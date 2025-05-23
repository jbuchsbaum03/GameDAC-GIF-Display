import os
import sys
import winshell
from win32com.client import Dispatch

import hid # For USB HID communication
from PIL import Image, ImageSequence, ImageDraw, ImageFont
import json
import time
from threading import Thread, Event

import tkinter as tk
from tkinter import filedialog
import pystray
from pystray import Menu, MenuItem

#############################################################################
#####                             USB/GIF CODE                          #####
#############################################################################

VENDOR_ID = 0x1038
PRODUCT_IDS = [0x12cb, 0x12cd, 0x12e0, 0x12e5]
SCREEN_WIDTH = 128
SCREEN_HEIGHT = 64
SCREEN_REPORT_SPLIT_SZ = 64
REPORT_SIZE = 1024

#Font Configuration
FONT_NAME = "PixelOperator.ttf"
FONT_SIZE = 16

class OLED_GIF:
    def __init__(self):
        self.device = None
        self.running = True
        self.frameDelaySeconds = 0.05
        self.invert = 0
        self.currentGIF = 0
        self.gif_frames_cache = {}

        self.connect_device()

    #########################################################################

    def connect_device(self):
        # Attempts to find and open the SteelSeries device via USB HID
        if self.device:
            try:
                self.device.close()
            except:
                pass
            self.device = None

        for device_dict in hid.enumerate(VENDOR_ID):
            if device_dict['product_id'] in PRODUCT_IDS:
                print(f"Found potential device: VID={hex(device_dict['vendor_id'])}, "
                      f"PID={hex(device_dict['product_id'])}, "
                      f"Path={device_dict['path'].decode()}, "
                      f"Interface={device_dict['interface_number']}")
                if device_dict['interface_number'] == 4 or device_dict['interface_number'] == -1:
                    try:
                        h = hid.device()
                        h.open_path(device_dict['path'])
                        prod_string = h.get_product_string()
                        mfg_string = h.get_manufacturer_string()
                        print(f"Successfully opened: {prod_string} / {mfg_string}")
                        self.device = h
                        return True
                    except Exception as e:
                        print(f"Could not open device {device_dict['path'].decode()}: {e}")
                        if h:
                            h.close()
        print("Compatible USB device not found.")
        return False

    #########################################################################

    def _create_draw_report(self, bitmap_segment, dst_x_on_screen, dst_y_on_screen):

        seg_w = bitmap_segment.width
        seg_h = bitmap_segment.height

        report = bytearray(REPORT_SIZE)
        report[0] = 0x06  # HID Report ID
        report[1] = 0x93  # Command ID for drawing
        report[2] = dst_x_on_screen & 0xFF
        report[3] = dst_y_on_screen & 0xFF
        report[4] = seg_w & 0xFF
        report[5] = seg_h & 0xFF

        stride_h = seg_h

        for x_in_seg in range(seg_w):
            for y_in_seg in range(seg_h):
                pixel_value_in_pil_image = bitmap_segment.getpixel((x_in_seg, y_in_seg))
                
                is_foreground_pixel = pixel_value_in_pil_image > 0 

                pixel_is_on_in_report = is_foreground_pixel if not self.invert else not is_foreground_pixel

                if pixel_is_on_in_report:
                    ri = x_in_seg * stride_h + y_in_seg
                    byte_index_in_report_data = ri // 8
                    bit_index_in_byte = ri % 8
                    report[6 + byte_index_in_report_data] |= (1 << bit_index_in_byte)
        return report

    #########################################################################

    def send_image_to_display(self, pil_image):

        if not self.device:
            if not self.connect_device():
                print("Failed to reconnect to device.")
                return False
        
        if pil_image.width != SCREEN_WIDTH or pil_image.height != SCREEN_HEIGHT:
            pil_image = pil_image.resize((SCREEN_WIDTH, SCREEN_HEIGHT)).convert('1')

        try:
            bitmap_segment1 = pil_image.crop((0, 0, SCREEN_REPORT_SPLIT_SZ, SCREEN_HEIGHT))
            report1 = self._create_draw_report(bitmap_segment1, 0, 0)
            self.device.send_feature_report(report1)

            bitmap_segment2 = pil_image.crop((SCREEN_REPORT_SPLIT_SZ, 0, SCREEN_WIDTH, SCREEN_HEIGHT))
            report2 = self._create_draw_report(bitmap_segment2, SCREEN_REPORT_SPLIT_SZ, 0)
            self.device.send_feature_report(report2)
            return True
        except Exception as e:
            print(f"Error sending frame to USB device: {e}")
            self.device = None 
            return False

    #########################################################################

    def process_gif(self, gif_path):
        if gif_path in self.gif_frames_cache:
            return self.gif_frames_cache[gif_path]

        frames = []
        try:
            with Image.open(gif_path) as img:
                for frame_pil in ImageSequence.Iterator(img):
                    frame_rgba = frame_pil.convert("RGBA")
                    
                    # Handle transparency data in GIFs
                    bg_aware_frame = Image.new("RGBA", frame_rgba.size, (0,0,0,255))
                    bg_aware_frame.paste(frame_rgba, (0,0), frame_rgba)

                    resized_frame = bg_aware_frame.resize((SCREEN_WIDTH, SCREEN_HEIGHT), Image.Resampling.LANCZOS)
                    
                    monochrome_frame = resized_frame.convert('1', dither=Image.Dither.FLOYDSTEINBERG)
                    frames.append(monochrome_frame)
            
            if frames:
                self.gif_frames_cache[gif_path] = frames
            return frames
        except FileNotFoundError:
            print(f"Error: GIF file not found at {gif_path}")
            return []
        except Exception as e:
            print(f"Error processing GIF {gif_path} with PIL: {e}")
            return []
        
    #########################################################################

    def playGIF(self, gif_path):
        print(f"Attempting to play GIF: {gif_path}")
        pil_gif_frames = self.process_gif(gif_path)
        if not pil_gif_frames:
            print(f"Could not process GIF: {gif_path}")
            self.display_error_message("GIF Error")
            return

        while self.running:
            if not pil_gif_frames: 
                break
            for current_pil_frame in pil_gif_frames:
                if not self.running:
                    break
                if not self.send_image_to_display(current_pil_frame):
                    print("Stopping GIF playback due to send error.")
                    self.stopGIF()
                    break
                time.sleep(self.frameDelaySeconds)
            if not self.running:
                break

    #########################################################################

    def playGIFCycle(self, gif_paths):
        print("Starting GIF cycle...")
        processed_pil_gifs = []
        for path in gif_paths:
            frames = self.process_gif(path) # Use the Pillow-based processing
            if frames:
                processed_pil_gifs.append(frames)
            else:
                print(f"Skipping {path} in cycle due to processing error.")

        if not processed_pil_gifs:
            print("No valid GIFs to cycle.")
            self.display_error_message("No GIFs")
            return

        ##                                        ##
        ## Store Report Array Here to Pre-Process ##
        ##                                        ##

        self.gif_start_event = Event()
        self.next_gif_event = Event()
        
        timer_thread = Thread(target=self.incrementTimer, args=(len(processed_pil_gifs),))
        timer_thread.daemon = True
        timer_thread.start()

        while self.running:
            self.gif_start_event.set() 
            self.next_gif_event.clear()
            
            while not self.next_gif_event.is_set() and self.running:
                for current_pil_frame in processed_pil_gifs[self.currentGIF]:
                    if (not self.running):
                        self.next_gif_event.set()
                        break
                    if not self.send_image_to_display(current_pil_frame):
                        print("Stopping GIF cycle due to send error.")
                        self.stopGIF()
                        break 
                    time.sleep(self.frameDelaySeconds)

    #########################################################################

    def incrementTimer(self, num_gifs_in_cycle):
        while self.running:
            if (not self.running):
                break

            self.gif_start_event.wait() 
            self.gif_start_event.clear()

            # Manual Timer Required 
            start_time = time.time()
            cycle_duration_per_gif = 20 
            while self.running and (time.time() - start_time < cycle_duration_per_gif):
                if not self.running:
                    break
                time.sleep(0.1)

            if self.running:
                self.currentGIF = (self.currentGIF + 1) % num_gifs_in_cycle
                self.next_gif_event.set() 
            else:
                break

    #########################################################################

    def stopGIF(self):
        self.running = False

    #########################################################################

    def display_error_message(self, message_text, timer):
        error_thread = Thread(target=self.display_error_helper, args=(message_text, timer), daemon=True)
        error_thread.start()

    def display_error_helper(self, message_text, timer):
        if not self.device:
            if not self.connect_device():
                return 

        try:
            font = ImageFont.truetype(FONT_NAME, FONT_SIZE)
        except IOError:
            font = ImageFont.load_default()

        image = Image.new('1', (SCREEN_WIDTH, SCREEN_HEIGHT), color=0) 
        draw = ImageDraw.Draw(image)
        
        try:
            bbox = draw.textbbox((0,0), message_text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        except AttributeError: 
            text_width, text_height = draw.textsize(message_text, font=font)

        x = (SCREEN_WIDTH - text_width) // 2
        y = (SCREEN_HEIGHT - text_height) // 2
        draw.text((x, y), message_text, font=font, fill=1) 

        self.send_image_to_display(image)

        if (timer):
            time.sleep(timer)
            original_invert_state = self.invert
            self.invert = False 
            black_screen_img = Image.new('1', (SCREEN_WIDTH, SCREEN_HEIGHT), color=0) 
            self.send_image_to_display(black_screen_img)
            self.invert = original_invert_state

    #########################################################################

    def quit_connection(self):
        if self.device:
            try:
                self.display_error_message("Spin Dial To Reset ->", 0)
            except Exception as e:
                print(f"Error closing device: {e}")
            time.sleep(0.1)
            self.device = None

#############################################################################
#####                             GUI CODE                              #####
#############################################################################

class GUI:
    def __init__(self, root, gif_player):
        self.root = root
        root.title("OLED GIF Display (USB)")
        root.geometry("325x270") 

        if hasattr(sys, '_MEIPASS'): 
            self.icon_path = os.path.join(sys._MEIPASS, "oled_gif.ico")
        else:
            self.icon_path = "oled_gif.ico" 

        try:
            root.iconbitmap(self.icon_path)
            self.tray_icon_image = Image.open(self.icon_path)
        except Exception as e:
            print(f"Error loading icon: {e}. Using default.")
            self.tray_icon_image = None 

        self.gif_player = gif_player 

        # Location for saving preferences
        documents_folder = os.path.expanduser("~\\Documents")
        self.game_dac_folder = os.path.join(documents_folder, "GameDAC GIF Display")
        self.pref_file_path = os.path.join(self.game_dac_folder, "preferences.json")

        # Establish Icon
        ico_menu = Menu(MenuItem("Show", self.show_window), MenuItem("Quit", self.quit))
        self.icon = pystray.Icon("oled_gif", self.tray_icon_image, "OLED GIF", menu=ico_menu)
        self.icon.run_detached()

        # Window Behavior #
        root.protocol("WM_DELETE_WINDOW", self.quit)
        root.bind("<Unmap>", self.to_tray)

        #####################################################################

        # Status Label #
        self.status_label = tk.Label(root, text="Please select a GIF", fg="black")
        self.status_label.pack(pady=5)

        # GIF Label #
        self.gif_label = tk.Label(root, text="No GIF Selected", fg="red")
        self.gif_label.pack(after=self.status_label)

        # Control Frame #
        control_frame = tk.Frame(root)
        control_frame.pack(pady=10)

        # Start Button #
        self.start_button = tk.Button(control_frame, text="Start", command=self.startGIF)
        self.start_button.config(background="lightgreen")
        self.start_button.pack(padx=5, side=tk.LEFT)

        # Invert Color
        self.invert_button = tk.Button(control_frame, text="Invert", command=self.invertColors)
        self.invert_button.config(background="#b0b0b0")
        self.invert_button.pack(padx=5, side=tk.LEFT)

        # Stop Button #
        self.stop_button = tk.Button(control_frame, text="Stop", command=self.stopGIF)
        self.stop_button.config(background="#ff6054", state=tk.DISABLED)
        self.stop_button.pack(padx=5, side=tk.LEFT)
        
        # Frame for File Buttons #
        file_frame = tk.Frame(root)
        file_frame.pack(after=control_frame, pady=10)

        # Browse Button #
        self.browse_button = tk.Button(file_frame, text="Choose GIF", command=self.browseGIF)
        self.browse_button.pack(side=tk.LEFT, padx=5)

        # Save Button #
        self.save_button = tk.Button(file_frame, text="Save GIF", command=self.saveGIF)
        self.save_button.pack(side=tk.LEFT, padx=5)

        # Clear Save Button #
        self.clear_button = tk.Button(file_frame, text="Clear Save", command=self.clearGIF)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        # Cycle Frame
        cycle_frame = tk.Frame(root)
        cycle_frame.pack(after=file_frame, pady=5)

        # Cycle GIFs Checkbox
        self.cycleVar = tk.BooleanVar()
        self.checkCycle = tk.Checkbutton(cycle_frame, text="Cycle multiple GIFs", variable=self.cycleVar, command=self.cycleToggle)
        self.checkCycle.pack(side=tk.LEFT, padx=5)

        # Select GIFs Button
        self.multiButton = tk.Button(cycle_frame, text="Open Folder", command=self.openFolder)
        self.multiButton.pack(side=tk.RIGHT, padx=5)

        # Options Frame
        options_frame = tk.Frame(root)
        options_frame.pack(after=cycle_frame, pady=10)

        # Start Minimized Checkbox
        self.minVar = tk.BooleanVar()
        self.chkmin = tk.Checkbutton(options_frame, text="Start in System Tray", variable=self.minVar, command=self.savePreferences)
        self.chkmin.pack(side=tk.LEFT, padx=5)

        # Run on Startup Checkbox
        self.startVar = tk.BooleanVar()
        self.checkStart = tk.Checkbutton(options_frame, text="Run on Startup", variable=self.startVar, command=self.savePreferences)
        self.checkStart.pack(side=tk.RIGHT, padx=5)


        # Check Loaded GIF / Settings #
        self.loadPreferences()

        if (self.minVar.get()):
            self.to_tray("<Unmap>")

        if not (self.cycleVar.get()):
            if (self.gif_path):
                self.save_button.config(state=tk.NORMAL)
                self.multiButton.config(state=tk.DISABLED)
                self.gif_label.config(text=f"Using {os.path.basename(self.gif_path)}", fg="black")
                self.startGIF()
            else:
                self.save_button.config(state=tk.DISABLED)
                self.start_button.config(state=tk.DISABLED)
                self.invert_button.config(state=tk.DISABLED)
                self.clear_button.config(state=tk.DISABLED)
        else:
            self.browse_button.config(state=tk.DISABLED)
            self.gif_label.config(text="Cycling GIFs Folder!", fg="black")
            self.startGIF()
        

    #############################################################################

    def startGIF(self):
        if (self.cycleVar.get()):
            self.startCycle()
        elif (self.gif_path):
            self.gif_player.running = True
            self.start_button.config(state=tk.DISABLED)
            
            self.stop_button.config(state=tk.NORMAL)
            self.invert_button.config(state=tk.NORMAL)

            self.status_label.config(text="Playing...", fg="green")
            gif_thread = Thread(target=self.gif_player.playGIF, args=(self.gif_path,))
            gif_thread.start()

    def startCycle(self):
        count = 0
        gif_paths = []
        multiGIF_Folder = os.path.join(self.game_dac_folder, "Cycle GIFs")
        os.makedirs(multiGIF_Folder, exist_ok=True)
        for file in os.listdir(multiGIF_Folder):
            if file.lower().endswith(".gif"):
                count+= 1
                gif_paths.append(os.path.join(multiGIF_Folder, file))
        if (count == 0):
            notif = "No GIFs in folder!"
            notif_thread = Thread(target=self.tempText, args=(self.gif_label, notif, "red"))
            notif_thread.start()
            self.gif_player.display_error_message("Add GIFs!")
            self.invert_button.config(state=tk.DISABLED)
        else:
            self.gif_player.running = True
            self.start_button.config(state=tk.DISABLED)

            self.invert_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.NORMAL)
            self.status_label.config(text="Playing...", fg="green")

            cycle_thread = Thread(target=self.gif_player.playGIFCycle, args=(gif_paths,))
            cycle_thread.start()


    def stopGIF(self):
        self.gif_player.stopGIF()
        self.stop_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.NORMAL)
        self.status_label.config(text=f"Stopped", fg="red")
        self.gif_player.display_error_message("GIF Stopped!", 2)

    def invertColors(self):
        self.gif_player.stopGIF()
        time.sleep(0.2)
        if (self.gif_player.invert):
            self.gif_player.invert = 0
        else:
            self.gif_player.invert = 1
        self.savePreferences()
        self.startGIF()

    def cycleToggle(self):
        self.stopGIF()
        
        if (self.cycleVar.get()):
            self.multiButton.config(state=tk.NORMAL)
            self.browse_button.config(state=tk.DISABLED)
            self.gif_label.config(text=f"Cycling GIFs folder!", fg="black")
        else:
            self.multiButton.config(state=tk.DISABLED)
            self.browse_button.config(state=tk.NORMAL)

            if (self.gif_path):
                self.gif_label.config(text=f"Using {os.path.basename(self.gif_path)}")
            else:
                self.gif_label.config(text="No GIF Selected", fg="red")
                self.start_button.config(state=tk.DISABLED)
                self.invert_button.config(state=tk.DISABLED)
                self.stop_button.config(state=tk.DISABLED)

        self.savePreferences()

    #############################################################################

    def browseGIF(self):
        wasRun = self.gif_player.running
        file_path = filedialog.askopenfilename(filetypes=[("GIF files", "*.gif")])
        if file_path:
            self.gif_player.stopGIF()
            if (not wasRun):
                self.gif_player.display_error_message("New GIF Selected!", 1)
            self.gif_path = file_path
            self.gif_label.config(text=f"Using {os.path.basename(file_path)}", fg="black")
            self.status_label.config(text="Waiting...", fg="black")
            self.start_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)
            time.sleep(0.1)
        if (wasRun):
            self.startGIF()

    def openFolder(self):
        os.makedirs(self.game_dac_folder, exist_ok=True)
        multiGIF_Folder = os.path.join(self.game_dac_folder, "Cycle GIFs")
        os.makedirs(multiGIF_Folder, exist_ok=True)
        os.startfile(multiGIF_Folder)

    def saveGIF(self):
        self.clear_button.config(state=tk.NORMAL)
        self.savePreferences()

        notif = "GIF Saved!"
        notif_thread = Thread(target=self.tempText, args=(self.gif_label, notif, "green"))
        notif_thread.start()

    def clearGIF(self):
        self.clear_button.config(state=tk.DISABLED)
        self.savePreferences()

        notif = "Save cleared!"
        notif_thread = Thread(target=self.tempText, args=(self.gif_label, notif, "red"))
        notif_thread.start()

    def tempText(self, label, message, color):
        prevText = label.cget("text")
        prevColor = label.cget("fg")
        label.config(text=message, fg=color)
        time.sleep(2)
        label.config(text=prevText, fg=prevColor)

    def savePreferences(self):
        os.makedirs(self.game_dac_folder, exist_ok=True)

        if (self.clear_button.cget("state") == tk.DISABLED):
            gifPath = None
        else:
            gifPath = self.gif_path

        data = {
            "start_min": self.minVar.get(),
            "startup": self.startVar.get(),
            "saved_gif": gifPath,
            "inverted": self.gif_player.invert,
            "cycle": self.cycleVar.get()
        }

        with open(self.pref_file_path, "w") as file:
            json.dump(data, file, indent=4)

        if (self.startVar.get()):
            self.add_to_startup()
        else:
            self.remove_from_startup()

    def loadPreferences(self):
        if os.path.exists(self.pref_file_path):
            with open(self.pref_file_path) as file:
                data = json.load(file)
                self.minVar.set(data.get("start_min"))
                self.startVar.set(data.get("startup"))
                self.gif_path = data.get("saved_gif")
                self.gif_player.invert = data.get("inverted")
                self.cycleVar.set(data.get("cycle"))
        else:
            self.minVar.set(False)
            self.startVar.set(False)
            self.gif_path = None
            self.cycleVar.set(False)

    #########################################################################

    def to_tray(self, event):
        self.root.withdraw()
        self.icon.visible = True

    def show_window(self):
        self.root.deiconify()
        self.root.lift()

    def quit(self):
        self.icon.stop()
        self.gif_player.stopGIF()
        self.gif_player.quit_connection()
        self.root.quit()

    def add_to_startup(self):
        startup_folder = winshell.startup()
        shortcut_path = os.path.join(startup_folder, "OLED_GIF.lnk")
        if not (os.path.exists(shortcut_path)):
            target = sys.executable

            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            shortcut.TargetPath = target
            shortcut.WorkingDirectory = os.path.dirname(target)
            shortcut.IconLocation = self.icon_path
            shortcut.Save()

    def remove_from_startup(self):
        startup_folder = winshell.startup()
        shortcut_path = os.path.join(startup_folder, "OLED_GIF.lnk")
        if (os.path.exists(shortcut_path)):
            os.remove(shortcut_path)

#############################################################################

if __name__ == "__main__":
    root = tk.Tk()
    gui = GUI(root,OLED_GIF())
    root.mainloop()

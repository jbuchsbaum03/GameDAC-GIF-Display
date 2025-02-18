import sys
import requests
from PIL import Image, ImageDraw
import pystray
from pystray import Menu, MenuItem
import json
import time
import os
import tkinter as tk
from tkinter import filedialog
from threading import Thread


#############################################################################
#####                             GIF CODE                              #####
#############################################################################

class OLED_GIF:
    def __init__(self):
        corePropsPath = r"C:\SteelSeries Replacement\Data\SteelSeries\GG\coreProps.json"
        self.sseAddress = f'http://{json.load(open(corePropsPath))["address"]}'
        self.game = "OLED_GIF"
        self.game_display_name = 'Display OLED GIF'
        self.event = "DISPLAY_GIF"
        self.running = True
        self.frameDelaySeconds = 0.025 #25 milliseconds
        
        self.registerGame()
        self.bindGameEvent()

    def registerGame(self):
        #Registers the game with SSE3
        data = {"game": self.game, "game_display_name": self.game_display_name, "developer": "TaleXVI"}
        requests.post(f'{self.sseAddress}/game_metadata', json=data)

    def bindGameEvent(self):
        #Binds an event for the OLED display
        data = {
            "game": self.game,
            "event": self.event,
            "value_optional": True,
            "handlers": [{
                "device-type": "screened-128x52",
                "mode": "screen",
                "zone": "one",
                "datas": [{
                    "length-millis": (self.frameDelaySeconds * 1000) + 50,
                    "has-text": False,
                    "image-data": [0]
                }]
                }]
        }
        requests.post(f'{self.sseAddress}/bind_game_event', json=data)

    #########################################################################

    def sendFrame(self, bitmap):
        #Sends the GIF frame to the OLED screen

        data = {
            "game": self.game,
            "event": self.event,
            "data": {
                "frame": {
                    "image-data-128x52": bitmap
                }
            }
        }
        requests.post(f'{self.sseAddress}/game_event', json=data)

    #########################################################################

    def playGIF(self, gif_path):
        gif_frames = processGIF(gif_path)
        while self.running:
            for frame in gif_frames:
                if (not self.running):
                    break
                self.sendFrame(frame)
                time.sleep(self.frameDelaySeconds)

    #########################################################################

    def stopGIF(self):
        self.running = False

    #########################################################################

    def removeGame(self):
        #Removes this application from Engine
        data = {
            "game": self.game
        }
        requests.post(f'{self.sseAddress}/remove_game', json=data)


    def removeGameEvent(self):
        #Removes this application from Engine
        data = {
            "game": self.game,
            "event": self.event
        }
        requests.post(f'{self.sseAddress}/remove_game', json=data)    


#############################################################################

def processGIF(gif_path):
    gif = Image.open(gif_path)
    gif.seek(0)

    gif_frames = []

    for frame in range(gif.n_frames):
        try:
            gif.seek(frame)
        except:
            break

        image = gif.copy().convert("1")
        image = image.resize((128,52))

        bitmap = [0] * 832

        index = 0
        for y in range(52):
            byte = 0
            for x in range(128):
                pixel = 0 if image.getpixel((x, y)) == 0 else 1
                byte |= (pixel << (7 - (x % 8)))
                if (x + 1) % 8 == 0:
                    bitmap[index] = byte
                    index += 1
                    byte = 0

        gif_frames.append(bitmap)

    return gif_frames

#############################################################################
#####                             GUI CODE                              #####
#############################################################################

class GUI:
    def __init__(self, root, gif_player):
        self.root = root
        root.title("OLED GIF Display")
        root.geometry("300x200")
        self.gif_player = gif_player
        
        documents_folder = os.path.expanduser("~\\Documents")
        self.game_dac_folder = os.path.join(documents_folder, "GameDAC GIF Display")
        self.save_file_path = os.path.join(self.game_dac_folder, "saved_gif")
        self.pref_file_path = os.path.join(self.game_dac_folder, "preferences")

        self.gif_path = self.loadGIF()

        #####################################################################

        # Icon Setup #
        self.icon = pystray.Icon("gif_icon", self.create_icon(), menu=Menu(MenuItem("Show", self.show_window), MenuItem("Quit", self.quit)))
        self.icon.run_detached()

        # Window Behavior #
        root.protocol("WM_DELETE_WINDOW", self.quit)
        root.bind("<Unmap>", self.to_tray)

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

        # Stop Button #
        self.stop_button = tk.Button(control_frame, text="Stop", command=self.stopGIF)
        self.stop_button.config(background="#ff6054", state=tk.DISABLED)
        self.stop_button.pack(padx=5, side=tk.LEFT)
        
        # Frame for File Buttons #
        file_frame = tk.Frame(root)
        file_frame.pack(after=control_frame, pady=10)

        # Browse Button #
        self.browse_button = tk.Button(file_frame, text="Browse GIF", command=self.browseGIF)
        self.browse_button.pack(side=tk.LEFT, padx=5)

        # Save Button #
        self.save_button = tk.Button(file_frame, text="Save GIF", command=self.saveGIF)
        self.save_button.pack(side=tk.LEFT, padx=5)

        # Clear Save Button #
        self.clear_button = tk.Button(file_frame, text="Clear Save", command=self.clearGIF)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        # Start Minimized Checkbox
        self.checked = tk.BooleanVar()
        self.checkbox = tk.Checkbutton(root, text="Start in system tray", variable=self.checked, command=self.savePreference)
        self.checkbox.pack(pady=10, after=file_frame)

        # Check Loaded GIF / Settings #
        if (self.gif_path):
            self.save_button.config(state=tk.NORMAL)
            self.gif_label.config(text=f"Using {os.path.basename(self.gif_path)}", fg="black")
            self.startGIF()

        else:
            self.save_button.config(state=tk.DISABLED)
            self.start_button.config(state=tk.DISABLED)
            self.clear_button.config(state=tk.DISABLED)

        if (self.loadPreference()):
            self.checked.set(True)
            self.to_tray("<Unmap>")


    #############################################################################

    def startGIF(self):
        if (self.gif_path):
            self.gif_player.running = True
            self.start_button.config(state=tk.DISABLED)
            
            self.stop_button.config(state=tk.NORMAL)
            self.status_label.config(text=f"Playing...", fg="green")

            gif_thread = Thread(target=self.gif_player.playGIF, args=(self.gif_path,))
            gif_thread.start()

    def stopGIF(self):
        self.gif_player.stopGIF()
        self.stop_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.NORMAL)
        self.status_label.config(text=f"Stopped", fg="red")

    #############################################################################

    def browseGIF(self):
        file_path = filedialog.askopenfilename(filetypes=[("GIF files", "*.gif")])
        if file_path:
            self.gif_path = file_path
            self.gif_label.config(text=f"Using {os.path.basename(file_path)}", fg="black")
            self.status_label.config(text="Waiting...", fg="black")
            self.start_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)

    def saveGIF(self):
        os.makedirs(self.game_dac_folder, exist_ok=True)

        with open(self.save_file_path, "w") as file:
            json.dump({"saved_gif": self.gif_path}, file)

        self.clear_button.config(state=tk.NORMAL)

        notif = "GIF Saved!"
        notif_thread = Thread(target=self.tempText, args=(self.gif_label, notif, "green"))
        notif_thread.start()

    def loadGIF(self):
        if os.path.exists(self.save_file_path):
            with open(self.save_file_path) as file:
                data = json.load(file)
                return data.get("saved_gif")
        else:
            return None

    def clearGIF(self):
        with open(self.save_file_path, "w") as file:
            json.dump({"saved_gif": None}, file)
    
        self.clear_button.config(state=tk.DISABLED)

        notif = "Save cleared!"
        notif_thread = Thread(target=self.tempText, args=(self.gif_label, notif, "red"))
        notif_thread.start()

    def tempText(self, label, message, color):
        prevText = label.cget("text")
        prevColor = label.cget("fg")
        label.config(text=message, fg=color)
        time.sleep(1)
        label.config(text=prevText, fg=prevColor)


    def loadPreference(self):
        if os.path.exists(self.pref_file_path):
            with open(self.pref_file_path) as file:
                data = json.load(file)
                return data.get("start_min")
        else:
            return None

    def savePreference(self):
        os.makedirs(self.game_dac_folder, exist_ok=True)
        with open(self.pref_file_path, "w") as file:
            json.dump({"start_min": self.checked.get()}, file)

    #########################################################################

    def create_icon(self):
        image = Image.new("RGB", (64, 64), color=(50,50,255))
        draw = ImageDraw.Draw(image)
        draw.rectangle([0,0,64,64], fill="blue")
        return image

    def to_tray(self, event):
        self.root.withdraw()
        self.icon.visible = True

    def show_window(self):
        self.root.deiconify()
        self.root.lift()

    def quit(self):
        self.icon.stop()
        self.stopGIF()
        self.root.quit()


#############################################################################

if __name__ == "__main__":
    root = tk.Tk()
    gui = GUI(root,OLED_GIF())
    root.mainloop()
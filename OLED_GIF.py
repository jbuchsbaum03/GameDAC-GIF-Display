import os
import sys
import winshell
from win32com.client import Dispatch

import requests
from PIL import Image as Image
import cv2
import json
import time
import numpy as np
from threading import Thread

import tkinter as tk
from tkinter import filedialog
import pystray
from pystray import Menu, MenuItem



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
        self.frameDelaySeconds = 0.001
        #0.001 = 1ms || 0.025 = 25ms
        self.invert = 0
        # 0 = No; 1 = Yes

        self.registerGame()
        self.bindGameEvent()

    #########################################################################

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
        gif_frames = processGIF(gif_path, self.invert)
        time.sleep(0.5)
        while self.running:
            for frame in gif_frames:
                if (not self.running):
                    break
                self.sendFrame(frame)
                time.sleep(self.frameDelaySeconds)


    def playGIFCycle(self, gif_paths):
        processedGIFs = []
        for path in gif_paths:
            gif_frames = processGIF(path, self.invert)

        while self.running:
            for gif_frames in processedGIFs:
                for frame in gif_frames:
                    if (not self.running):
                        break
                    self.sendFrame(frame)
                    time.sleep(self.frameDelaySeconds)
                time.sleep(0.5)


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

def processGIF(gif_path, invert):
    gif = cv2.VideoCapture(gif_path)
    gif_frames = []
    method = cv2.THRESH_BINARY_INV if invert else cv2.THRESH_BINARY

    while True:
        success,frame = gif.read()
        if not success:
            break

        
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        frame = cv2.resize(frame, (128, 52), interpolation=cv2.INTER_CUBIC)
        _, frame = cv2.threshold(frame, 128, 1, method)
        

        bytemap = [0] * 832
        index = 0

        for y in range(52):
            byte = 0
            for x in range(128):
                pixel = int(frame[y,x])
                byte |= (pixel << (7 - (x % 8)))
                if (x + 1) % 8 == 0:
                    bytemap[index] = byte
                    index += 1
                    byte = 0

        gif_frames.append(bytemap)

    gif.release()
    return gif_frames

#############################################################################
#####                             GUI CODE                              #####
#############################################################################

class GUI:
    def __init__(self, root, gif_player):
        self.root = root
        root.title("OLED GIF Display")
        root.geometry("325x240")
    
        if hasattr(sys, '_MEIPASS'):
            self.icon_path = os.path.join(sys._MEIPASS, "oled_gif.ico")
        else:
            self.icon_path = "oled_gif.ico"

        root.iconbitmap(self.icon_path)

        self.gif_player = gif_player
        
        documents_folder = os.path.expanduser("~\\Documents")
        self.game_dac_folder = os.path.join(documents_folder, "GameDAC GIF Display")
        self.pref_file_path = os.path.join(self.game_dac_folder, "preferences")

        #####################################################################

        # Icon Setup #
        self.icon = pystray.Icon("oled_gif", Image.open(self.icon_path), menu=Menu(MenuItem("Show", self.show_window), MenuItem("Quit", self.quit)))
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
                self.clear_button.config(state=tk.DISABLED)
        elif (self.cycleVar.get()):
            self.browse_button.config(state=tk.DISABLED)
            self.gif_label.config(text=f"Cycling GIFs Folder!", fg="black")
        

    #############################################################################

    def startGIF(self):
        if ((not self.cycleVar.get()) and self.gif_path):
            self.running = True
            self.start_button.config(state=tk.DISABLED)
            
            self.stop_button.config(state=tk.NORMAL)
            self.status_label.config(text=f"Playing...", fg="green")

            gif_thread = Thread(target=self.gif_player.playGIF, args=(self.gif_path,))
            gif_thread.start()

        elif (self.cycleVar.get()):
            self.startCycle()

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

    def stopGIF(self):
        self.gif_player.stopGIF()
        self.stop_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.NORMAL)
        self.status_label.config(text=f"Stopped", fg="red")

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
        notif = "Mode changed!"
        self.stopGIF()
        self.status_label.config(text='Waiting...', fg='black')
        notif_thread = Thread(target=self.tempText, args=(self.gif_label, notif, "green"))
        notif_thread.start()

        if (self.cycleVar.get()):
            self.multiButton.config(state=tk.NORMAL)
            self.browse_button.config(state=tk.DISABLED)
            self.gif_label.config(text=f"Cycling GIFs folder!", fg="black")
        else:
            self.multiButton.config(state=tk.DISABLED)
            self.browse_button.config(state=tk.NORMAL)
            self.gif_label.config(text=(f"Using {os.path.basename(self.gif_path)}") if (self.gif_path) else "Please select a GIF")
        self.savePreferences()



    #############################################################################

    def browseGIF(self):
        file_path = filedialog.askopenfilename(filetypes=[("GIF files", "*.gif")])
        if file_path:
            self.gif_player.stopGIF()
            self.gif_path = file_path
            self.gif_label.config(text=f"Using {os.path.basename(file_path)}", fg="black")
            self.status_label.config(text="Waiting...", fg="black")
            self.start_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)

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
        time.sleep(1)
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
            json.dump(data, file)

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
        self.stopGIF()
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
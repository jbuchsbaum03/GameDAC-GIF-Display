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
        corePropsPath = r"C:\ProgramData\SteelSeries\GG\coreProps.json"
        self.sseAddress = f'http://{json.load(open(corePropsPath))["address"]}'
        self.game = "OLED_GIF"
        self.game_display_name = 'Display OLED GIF'
        self.event = "DISPLAY_GIF"
        self.running = True
        self.frameDelaySeconds = 0 #0.025 = 25 milliseconds
        
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
                    "length-millis": (self.frameDelaySeconds * 1000) + 25,
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
    gif = cv2.VideoCapture(gif_path)
    gif_frames = []

    while True:
        success,frame = gif.read()
        if not success:
            break

        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        _, frame = cv2.threshold(frame, 128, 255, cv2.THRESH_BINARY)
        frame = cv2.resize(frame, (128, 52), interpolation=cv2.INTER_CUBIC)

        bytemap = [0] * 832
        index = 0

        for y in range(52):
            byte = 0
            for x in range(128):
                pixel = 0 if frame[y, x] == 0 else 1
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
        root.geometry("300x250")
    
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
        self.minVar = tk.BooleanVar()
        self.chkmin = tk.Checkbutton(root, text="Start in System Tray", variable=self.minVar, command=self.savePreferences)
        self.chkmin.pack(pady=10, after=file_frame)

        # Run on Startup Checkbox
        self.startVar = tk.BooleanVar()
        self.checkStart = tk.Checkbutton(root, text="Run on Startup", variable=self.startVar, command=self.savePreferences)
        self.checkStart.pack(after=self.chkmin)

        # Check Loaded GIF / Settings #
        self.loadPreferences()

        if (self.gif_path):
            self.save_button.config(state=tk.NORMAL)
            self.gif_label.config(text=f"Using {os.path.basename(self.gif_path)}", fg="black")
            self.startGIF()
        else:
            self.save_button.config(state=tk.DISABLED)
            self.start_button.config(state=tk.DISABLED)
            self.clear_button.config(state=tk.DISABLED)

        if (self.minVar.get()):
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
            self.gif_player.stopGIF()
            self.gif_path = file_path
            self.gif_label.config(text=f"Using {os.path.basename(file_path)}", fg="black")
            self.status_label.config(text="Waiting...", fg="black")
            self.start_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)

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
            "saved_gif": gifPath
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
        else:
            self.minVar.set(False)
            self.startVar.set(False)
            self.gif_path = None

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

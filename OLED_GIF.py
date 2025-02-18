import json
import requests
from PIL import Image

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
        corePropsPath = r"C:\ProgramData\SteelSeries\GG\coreProps.json"
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
#####                             GUI CODE                              #####
#############################################################################

class GUI:
    def __init__(self, root, gif_player):
        self.root = root
        root.title("OLED GIF Display")
        root.geometry("300x220")
        self.gif_player = gif_player
        self.gif_path = None

        
        self.status_label = tk.Label(root, text="Please select a GIF", fg="black")
        self.status_label.pack(pady=10)

        self.start_button = tk.Button(root, text="Start", command=self.startGIF)
        self.start_button.pack(pady=10)

        self.stop_button = tk.Button(root, text="Stop", command=self.stopGIF)
        self.stop_button.pack(pady=10)

        self.stop_button.config(state=tk.DISABLED)

        self.browse_button = tk.Button(root, text="Browse GIF", command=self.browseGIF)
        self.browse_button.pack(pady=10)

        self.gif_label = tk.Label(root, text="No GIF Selected", fg="red")
        self.gif_label.pack(pady=5, after=self.browse_button)

        

#############################################################################

    def startGIF(self):

        if (self.gif_path):
            self.gif_player.running = True
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.status_label.config(text=f"Playing...", fg="black")

            gif_thread = Thread(target=self.gif_player.playGIF, args=(self.gif_path,))
            gif_thread.start()

    def stopGIF(self):
        self.gif_player.stopGIF()
        self.stop_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.NORMAL)
        self.status_label.config(text=f"Stopped", fg="red")

    def browseGIF(self):
        file_path = filedialog.askopenfilename(filetypes=[("GIF files", "*.gif")])
        if file_path:
            self.gif_path = file_path
            self.gif_label.config(text=f"Using {os.path.basename(file_path)}", fg="green")


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
    
    print(f" GIF has {len(gif_frames)} frames")

    return gif_frames

#############################################################################

if __name__ == "__main__":
    root = tk.Tk()
    gui = GUI(root,OLED_GIF())
    root.mainloop()
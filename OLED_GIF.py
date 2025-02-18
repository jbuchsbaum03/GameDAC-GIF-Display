import json
import requests
import time
from PIL import Image

from os import getenv
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
        self.frame_delay = 100  # 100ms per frame (adjust as needed)
        self.running = True
        
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
                    "length-millis": self.frame_delay,
                    "has-text": False,
                    "image-data": [0]
                }]
                }]
        }
        requests.post(f'{self.sseAddress}/bind_game_event', json=data)

    #########################################################################

    def sendFrame(self, image):
        #Sends the GIF frame to the OLED screen

        image = image.convert("1")
        image = image.resize((128,52))

        bitmap = []

        # Process the image row by row (pack bits into bytes)
        for y in range(52):
            byte = 0
            for x in range(128):
                pixel = 0 if image.getpixel((x, y)) == 0 else 1
                byte |= (pixel << (7 - (x % 8)))
                if (x + 1) % 8 == 0:
                    bitmap.append(byte)
                    byte = 0

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

        gif = Image.open(gif_path)
        while self.running:
            for frame in range(gif.n_frames):
                if (not self.running):
                    break
                gif.seek(frame)
                self.sendFrame(gif)

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
        self.gif_player = gif_player

        self.start_button = tk.Button(root, text="Start", command=self.startGIF)
        self.start_button.pack(pady=10)

        self.stop_button = tk.Button(root, text="Stop", command=self.stopGIF)
        self.stop_button.pack(pady=10)

        self.stop_button.config(state=tk.DISABLED)

        self.browse_button = tk.Button(root, text="Browse GIF", command=self.browseGIF)
        self.browse_button.pack(pady=10)

        self.gif_path = None

#############################################################################

    def startGIF(self):
        self.gif_player.running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        gif_thread = Thread(target=self.gif_player.playGIF, args=(self.gif_path,))
        gif_thread.start()

    def stopGIF(self):
        self.gif_player.stopGIF()
        self.stop_button.config(state=tk.DISABLED)
        self.start_button.config(state=tk.NORMAL)

    def browseGIF(self):
        file_path = filedialog.askopenfilename(filetypes=[("GIF files", "*.gif")])
        if file_path:
            self.gif_path = file_path


#############################################################################

if __name__ == "__main__":
    oled = OLED_GIF()

    root = tk.Tk()
    root.title("OLED GIF Display")
    root.geometry("300x200")
    gui = GUI(root,oled)

    root.mainloop()
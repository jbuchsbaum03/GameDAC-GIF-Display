import json
import sys
import time
import requests
from os import getenv

class OLED_GIF:
    def __init__(self):
        corePropsPath = r"C:\ProgramData\SteelSeries\GG\coreProps.json"
        self.sseAddress = f'http://{json.load(open(corePropsPath))["address"]}'
        self.game = "OLED_TEXT"
        self.game_display_name = 'Display OLED Text'
        self.event = "DISPLAY"
        # self.frame_delay = 0.1  # 100ms per frame (adjust as needed)
        
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
                "device-type": "screened",
                "mode": "screen",
                "zone": "one",
                "datas": [{
                    "has-text": True,
                    "context-frame-key": "custom-text",
                    "length-millis": 2000
                }]
                }]
        }
        requests.post(f'{self.sseAddress}/bind_game_event', json=data)


    def displayText(self, text):
        #Sends the text to the OLED screen
        data = {
            "game": self.game,
            "event": self.event,
            "data": {
                "frame": {
                    "custom-text": text
                }
            }
        }
        response = requests.post(f'{self.sseAddress}/game_event', json=data)
        if response.status_code == 200:
            print("Okay")
        else:
            print(f"Failed to display text: {response.status_code}, {response.text}")



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

if __name__ == "__main__":
    GIFPlayer = OLED_GIF()
    GIFPlayer.bindGameEvent()

    while True:
        text = input("Display Message: ")
        if (text == "Exit"):
            sys.exit(0)
        elif (text == "REMOVE"):
            GIFPlayer.removeGameEvent()
            GIFPlayer.removeGame()
            sys.exit(0)
        else:
            GIFPlayer.displayText(text)


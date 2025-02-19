# GameDAC-GIF-Display
This repo contains 2 python scripts: OLED_GIF.py and OLED_TEXT.py.

OLED_GIF.py has all the code for the purposes of this project. OLED_TEXT.py was a smaller project I completed first, but left in the repo in case anyone finds it beneficial to reference.

This project enables owners of a SteelSeries GameDAC or Nova Pro & Base Station to select a gif and loop it on the OLED embedded in these devices. ANY gif can be used- it will automatically be resized and converted to black and white so that it can be displayed on the screen. No promises that the displayed gif will be of high quality though, especially if your input is very high motion or very colorful. I recommend sticking to MOSTLY black and white gifs to ensure the output is clean.

## Notice
This program requires SteelSeriesGG / SteelSeriesEngine to be running to communicate with OLED devices.
Additionally, the program assumes that your coreProps.json file is located at "C:\ProgramData\SteelSeries\GG", which it should be if you installed GG in its default directory.
This can be changed at the top of the OLED_GIF.py file, but you will have to manually compile the program.

## How to Use

### Using the EXE:
* Download OLED_GIF.exe from the assests section of this repository
* Run the .exe

### Manually Compile:
* Download / Clone the git repository
* In the console/cmd, navigate to the directory where you downloaded the files
* Ensure that you have pyinstaller (if not, run pip install pyinstaller -> if this confuses you use the .exe)
* Run the following command: <br> pyinstaller --onefile --noconsole --icon=oled_gif.ico --add-data "oled_gif.ico;." OLED_GIF.py
* Locate the compiled .exe
* Run the .exe

### Interacting with the Program
* 'Start' starts playback of the selected gif.
* 'Stop' stops playback of the selected gif.
* 'Browse GIF' allows you to select a local gif file to play.
* 'Save GIF' stores the currently selected gif to be automatically loaded and started on program start.
* 'Clear Save' removes the currently saved gif.
* 'Start in System Tray' enables the program to start directly into the system tray (i.e. minimized/hidden)
* 'Run on Startup' enables the program to run at startup/login

The idea behind 'start in system tray' and 'run on startup' is that both can be enabled to automatically play your saved gif when you start your PC. Nice right?

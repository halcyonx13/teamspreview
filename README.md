# TeamsPreview

The Microsoft Teams pop-out preview window is inflexible. It can't be resized and has a large GUI that can't be hidden.

This is a python application that uses pyglet and [windows-capture](https://github.com/NiiightmareXD/windows-capture/tree/main/windows-capture-python) to replicate the Teams video call window into a more flexible window. **It can be moved, resized, zoomed, and panned.**

Drawback is you can't minimise the Teams call window (keep it open in the background).

It's a bit resource heavy but it works for me, and maybe it will work for you.

![alt text](https://raw.githubusercontent.com/halcyonx13/teamspreview/refs/heads/main/teamspreview_screenshot.jpg)

## How to use:

### Installing

A .exe file generated with pyinstaller is included in releases. That is probably the easiest way to use this. Only tested on Windows 10 22H2. If it doesn't work on your system, you will have to fix it yourself.

[Download it here](https://github.com/halcyonx13/teamspreview/releases/)

### Using

When your Teams video call has started, open `teamspreview.exe` and choose the meeting window by moving with the arrow keys and pressing Enter. Do not minimise the video call window, simply switch to another application and the preview will appear.

#### Controls

* Left mouse button drag: Move window
* Right mouse button drag: Pan frame
* Mousewheel: Zoom in/out
* Space: Return to window selection screen
* F10: Reset zoom/pan
* F11: show FPS counter
* F12: swap colour mode (RGBA/BGRA)

#### Advanced

* JSON config file is saved in your home directory (user profile folder) and is named `.teams_preview_config.json`
* You can add a line for FPS to set a custom FPS limit, e.g. `"fps": 10`
* You can also set `"vsync": true` if you wanted vsync for whatever reason
* If you want to use the preview with applications other than Teams, set the window filter with `"titlefilter": "Skype"` <- (this will only show windows with the word Skype in the title in the window selection screen)
* If you want to rebuild the .exe file, I simply used `pyinstaller teamspreview.py --onefile --noconsole `



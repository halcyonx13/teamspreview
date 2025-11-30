# TeamsPreview

The Microsoft Teams pop-out preview window is terrible. It can't be resized and has an obnoxiously large amount of GUI.

This is a python application that uses pyglet and [windows-capture](https://github.com/NiiightmareXD/windows-capture/tree/main/windows-capture-python) to replicate the Teams video call window into a more flexible window.

Drawback is you can't minimise the Teams call window (keep it open in the background).

It's a bit resource heavy but it works for me, and maybe it will work for you.

A .exe file generated with pyinstaller is included in releases. That is probably the easiest way to use this. Only tested on Windows 10 22H2. If it doesn't work on your system, fix it yourself.
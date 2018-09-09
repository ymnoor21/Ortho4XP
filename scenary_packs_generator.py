#!/usr/bin/env python3
import os
from pathlib import Path
from win32com.client import Dispatch
 
def createWindowsShortcutDirectory(path, target='', workDir='', icon=''):    
    dispatch = Dispatch('WScript.Shell')
    shortcut = dispatch.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = workDir

    if icon != '':
        shortcut.IconLocation = icon

    shortcut.save()

if __name__ == "__main__":
	dir_path = os.path.dirname(os.path.realpath(__file__))
	tiles_dir = '\\Tiles'
	custom_scenary_path = "D:\\X-Plane 11\\Custom Scenery"

	tiles_dir_path = dir_path + tiles_dir

	for directory in os.listdir(tiles_dir_path):
		path = custom_scenary_path + "\\" + directory + ".lnk"
		target = tiles_dir_path + "\\" + directory 
		my_file = Path(path)

		if not my_file.exists():
			createWindowsShortcutDirectory(path, target)

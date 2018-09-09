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
	xplane_directory = "D:\\X-Plane 11"
	custom_scenary = xplane_directory + "\\Custom Scenery"

	tiles_dir_path = dir_path + tiles_dir

	for directory in os.listdir(tiles_dir_path):
		path = custom_scenary + "\\" + directory + ".lnk"
		target = tiles_dir_path + "\\" + directory 
		my_file = Path(path)

		if not my_file.exists():
			createWindowsShortcutDirectory(path, target)

	# delete and copy yOrtho4XP_Overlays
	yOverlay_path = custom_scenary + "\\yOrtho4XP_Overlays.lnk"
	yOverlay_target = dir_path + "\\yOrtho4XP_Overlays"

	yOverlay_lnk = Path(yOverlay_path)

	if not yOverlay_lnk.exists():
		createWindowsShortcutDirectory(yOverlay_path, yOverlay_target)

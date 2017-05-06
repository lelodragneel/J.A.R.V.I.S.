import win32api as w32
import winxptheme
import win32gui
import win32file
import win32process
import subprocess
import winshell
import os
import sys
from string import ascii_uppercase


def getComputerName():
    name = w32.GetComputerName()
    return name


# def GetWindowTheme():
#     if winxptheme.IsThemeActive():
#         print('k')
#         name = winxptheme.GetCurrentThemeName()
#         return str(name)


def getHandlesOfVisibleWindows():
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            name = win32gui.GetWindowText()
            hwnds.append(name)
        return True
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds


def openApplication(appName):
    subprocess.Popen(appName)


def delete(sourcePath):
    winshell.delete_file(sourcePath, no_confirm=True)


def undelete(sourcePath):
    winshell.undelete(sourcePath)


def recycledDate(sourcePath):
    recycleBin = winshell.recycle_bin()
    print(recycleBin.versions(sourcePath))


def rename(sourcePath, targetPath):
    winshell.rename_file(sourcePath, targetPath, no_confirm=True)


def copy(sourcePath, targetPath):
    winshell.copy_file(sourcePath, targetPath, no_confirm=True)


def move(sourcePath, targetPath):
    winshell.move_file(sourcePath, targetPath, no_confirm=True)


def createShortcut(filePath):
    winshell.CreateShortcut(
        Path=os.path.join(winshell.desktop(), "shortcut.lnk"),
        Target=filePath,
        Icon=(filePath, 0),
        Description="Shortcut"
    )


def getDiskFreeSpace(volumeLetter):
    freeBytes, totalBytes, totalFreeBytes = w32.GetDiskFreeSpaceEx(volumeLetter)
    return totalFreeBytes


def getVolumeLetters():
    activeVolumes = []
    for i in ascii_uppercase:
        if win32file.GetDriveType(i + ":\\") > 1:
            activeVolumes.append(([i],[win32file.GetDriveType(i + ":\\")]))
    return activeVolumes
import win32api
import winxptheme
import win32gui
import win32file
import win32process
import subprocess
import winshell
import os
import sys
import win32con
from string import ascii_uppercase


def getWindowWithTitle(title):
    hwndMain = win32gui.FindWindow(None, title)
    return hwndMain


def getWindowWithPID(pid):

    def callback(hwnd, handlers):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            ctid, cpid = win32process.GetWindowThreadProcessId(hwnd)
            if cpid == pid:
                handlers.append(win32gui.GetWindowText(hwnd))
        return True

    handlers = []
    win32gui.EnumWindows(callback, handlers)
    return handlers


def getAllItems(dir):
    files = []
    for entry in os.scandir(dir):
        files.append(entry.name)
    return files


def getComputerName():
    name = win32api.GetComputerName()
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
            name = win32gui.GetWindowText(hwnd)
            hwnds.append([hwnd, name])
        return True
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)
    return hwnds


# def minimizeAll():
#     for i in getHandlesOfVisibleWindows():
#         win32gui.PostMessage(i[0], win32con.SW_MINIMIZE)


def openApplication(appName):
    subprocess.Popen(appName)


def delete(sourcePath):
    winshell.delete_file(sourcePath, no_confirm=True)


def undelete(sourcePath):
    winshell.undelete(sourcePath)


def getRecycledDate(sourcePath):
    recycleBin = winshell.recycle_bin()
    return recycleBin.versions(sourcePath)


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
    freeBytes, totalBytes, totalFreeBytes = win32api.GetDiskFreeSpaceEx(volumeLetter)
    return totalFreeBytes


def getVolumeLetters():
    activeVolumes = []
    for i in ascii_uppercase:
        if win32file.GetDriveType(i + ":\\") > 1:
            activeVolumes.append(([i],[win32file.GetDriveType(i + ":\\")]))
    return activeVolumes
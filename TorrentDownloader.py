import os
import time
import subprocess
import win32api
from pywinauto.application import Application

folderPath = 'D:\\Downloads'
fileExt = '.torrent'
bittorrentPath = 'C:\\Users\\rospa\\AppData\\Roaming\\bittorrent\\BitTorrent.exe'

existingFiles = set()

while True:
    currentFiles = set(os.listdir(folderPath))
    newFiles = [file for file in (currentFiles - existingFiles) if file.endswith(fileExt)]

    if newFiles:
        print("\nNew {} files detected:".format(fileExt))

        for file in newFiles:
            print(file)
            filePath = os.path.join(folderPath, file)

            try:
                subprocess.run([bittorrentPath, filePath], check = True)

                app = Application().connect(path = bittorrentPath)
                app.window(title_re = '.*Add New Torrent.*').wait('ready', timeout = 10)
                app.window(title_re = '.*Add New Torrent.*').OK.click()
            except subprocess.CalledProcessError as e:
                print(f"Error executing {file}: {e}")

    existingFiles = currentFiles

    time.sleep(60)

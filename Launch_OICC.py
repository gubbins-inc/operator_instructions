import configparser
import os
import subprocess
from win32com.client import Dispatch
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import askyesno


def get_version_number(app_location):
    parser = Dispatch("Scripting.FileSystemObject")
    version = parser.GetFileVersion(app_location)
    return version


def get_versions():
    config = configparser.ConfigParser()
    config.read('settings.ini')
    try:
        localversion = (config['info']['sw_version'])
        localversion = float(localversion)
    except KeyError:
        localversion = False

    try:
        server_file = r'Codebase\server_settings.ini'
        paff = config['rootpath']['path']
        rootpath = paff + '\\'
        server_side = os.path.abspath(rootpath + server_file)
        config.read(server_side)
        serverversion = (config['info']['sw_version'])
        serverversion = float(serverversion)
    except KeyError:
        serverversion = False
    print("server: " + str(serverversion), "settings: " + str(localversion))
    if serverversion is False or localversion is False:
        print("unable to determine local and / or server versions of software")
        return True

    else:
        if serverversion > localversion:
            print("local version reported in settings file appears to be out of date")
            print("checking...")
            command = "OICC.exe --version"
            pipe = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

            while True:
                line = pipe.stdout.readline()
                if line:
                    found_version = float(line)
                if not line:
                    break
            print("server: " + str(serverversion), "settings: " + str(localversion), "local: " + str(found_version))

            if found_version > localversion:
                config2 = configparser.ConfigParser()
                config2.read('settings.ini')
                print("updating stored version number in settings file")
                config2.set('info', 'sw_version', str(found_version))
                with open('settings.ini', 'w') as configfile:
                    config2.write(configfile)
                if get_versions():
                    return True

            if serverversion > found_version:
                print("there is a newer version of the software available")
                # do update and close
                app_file = r"Codebase\update_OICC.exe"
                app_location = os.path.abspath(rootpath + app_file)
                version = get_version_number(app_location)
                vnum = float(version[0:3])
                print("installer reports it is at version " + str(vnum))

                if vnum > found_version:
                    root = tk.Tk()
                    root.overrideredirect(1)
                    root.title('')
                    root.geometry('0x0')
                    root.withdraw()

                    # click event handler
                    def confirm():
                        answer = askyesno(title='Update Available',
                                          message='There is an updated version of OICC, would you like to install?')
                        if answer:
                            root.destroy()
                            # do update
                            command = app_location
                            pipe = subprocess.Popen(command, shell=True)
                            return False
                        else:
                            root.destroy()
                            return True

                    confirm()
                    root.mainloop()
                else:
                    print("the installer file is not newer than the installed file...")
                    return True

        else:
            print("you are running the latest version of the software")
            return True


if get_versions():
    command = "OICC.exe"
    pipe = subprocess.Popen(command, shell=True)
else:
    os.sys.exit()

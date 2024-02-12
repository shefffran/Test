import ssd_checker

import PySimpleGUI as sg
import platform
import socket
import psutil
import os
import wmi
import pandas as pd
import shutil
import subprocess
import re

# --------------------
# os info
# -------------------


platform.node()
sys = platform.uname()

computer = wmi.WMI()
computer_info = computer.Win32_ComputerSystem()[0]
os_info = computer.Win32_OperatingSystem()[0]
proc_info = computer.Win32_Processor()[0]
gpu_info = computer.Win32_VideoController()[0]

os_name = os_info.Name.encode('utf-8').split(b"|")[0]
os_version = ' '.join([os_info.Version, os_info.BuildNumber])
system_ram = float(os_info.TotalVisibleMemorySize) / 1048576  #
new_system_ram = str(system_ram)[:-13]

ws = wmi.WMI(namespace='root/Microsoft/Windows/Storage')


totalSSD = int()
freeSSD  = int()
totalHDD = int()
freeHDD  = int()


print(ssd_checker.division)
for disk, d in zip(psutil.disk_partitions(), ws.MSFT_PhysicalDisk()):
    print(disk.mountpoint)
    if disk.fstype:
        if d.MediaType ==4:
            totalSSD += int(psutil.disk_usage(disk.mountpoint).total)
            freeSSD  += int(psutil.disk_usage(disk.mountpoint).free)
            print ("ssd")

        if d.MediaType ==3:
            totalHDD += int(psutil.disk_usage(disk.mountpoint).total)
            freeHDD  += int(psutil.disk_usage(disk.mountpoint).free)
            print ("Hdd")



ssdGB = round(totalSSD / (1024.0 ** 3), 4)
ssdFree = round(freeSSD / (1024.0 ** 3), 4)
hddGB = round(totalHDD / (1024.0 ** 3), 4)
hddFree = round(freeHDD / (1024.0 ** 3), 4)


# --------------------
#Testing
# --------------------



# ---------------------

# disk_freeSpace = str(float(d.FreeSpace) / 1000000000)
# new_disk_freeSpace = disk_freeSpace[:-6] + " GB"
#
# disk_Size = str(float(d.Size) / 1000000000) #steya error -------------------------------------------------------------------------------------
# new_disk_Size = disk_Size[:-6] + " GB"
# if(d.MediaType == 4):
#     ssdGB = new_disk_Size
#     ssdFree = new_disk_freeSpace
#
# print(f" Disk Size: {d.Caption, new_disk_Size, new_disk_freeSpace, d.DriveType}")




# print(f"Disk Types:")
# for d in ws.MSFT_PhysicalDisk():
#     print(" " + d.Model)
#
# c = wmi.WMI()
# for d in c.Win32_LogicalDisk():
#     disk_freeSpace = str(float(d.FreeSpace) / 1000000000)
#     new_disk_freeSpace = disk_freeSpace[:-6] + " GB"
#
#     disk_Size = str(float(d.Size) / 1000000000)
#     new_disk_Size = disk_Size[:-6] + " GB"
#     print(f" Disk Size: {d.Caption, new_disk_Size, new_disk_freeSpace, d.DriveType}")







# --------------------
#GUI
# --------------------


sg.theme('BrownBlue')


layout = [
    [sg.Text('FileBrowse',justification='center')],
    [sg.Input(expand_x=True, key="-FilePath-"),sg.FileBrowse(file_types=(("MIDI files", "*.xlsx"),))],

    [sg.Text('OS Type',size=(17,1)), sg.InputText(os_name,key='Os Name')],

    [sg.Text('PC Name',size=(17,1)), sg.InputText(sys.node[:-5],key='PC Name')],

    [sg.Text('CPU Model',size=(17,1)), sg.InputText(format(proc_info.Name),key='CPU Model')],
    [sg.Text('CPU Date',size=(17,1)), sg.InputText('',key='CPU Date')],

    [sg.Text('Ram (GB)',size=(17,1)), sg.InputText(format(new_system_ram),key='Ram (GB)')],
    [sg.Text('Ram Tech.',size=(17,1)), sg.InputText('',key='Ram Tech.')],

    [sg.Text('SSD (GB)',size=(17,1)), sg.InputText(ssdGB,key='SSD (GB)')],
    [sg.Text('Free Space(GB)',size=(17,1)), sg.InputText(ssdFree,key='Free Space(GB)')],

    [sg.Text('HDD (GB)',size=(17,1)), sg.InputText(hddGB,key='HDD (GB)')],
    [sg.Text('HDD Free Space (GB)',size=(17,1)), sg.InputText(hddFree,key='HDD Free Space (GB)')],

    [sg.Text('GPU',size=(17,1)), sg.InputText(format(gpu_info.Name),key='GPU')],

    [sg.Text('Description',size=(17,1)), sg.InputText('',key='Description')],

    [sg.Save(size=(15,1)),sg.Exit(size=(15,1))]
          ]


window = sg.Window('Hello Bitch',layout)

oneTimeIf = int(0)


while True:
    event,values = window.read()
    try:
        if oneTimeIf == 0:
            FilePath = values["-FilePath-"]
            EXCEL_FILE = FilePath
            df = pd.read_excel(EXCEL_FILE)
            oneTimeIf = 1
    except:
        sg.popup('Error')
        break

    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event =='Save':
        print(event,values)
        df = df._append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Saved')


window.close()

# -------------------
#
# --------------------

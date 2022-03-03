# It is recommended to disable UAC since this runs powershell commands
# as Administrator and things can fail if you don't click 'yes' fast enough
import subprocess
import json
import PySimpleGUI as sg
import time
import tempfile
from operator import itemgetter

# Create powershell scripts in temp dir, execute them, and load them into dicts
def powershell():
    absdir = tempfile.gettempdir()
    file = open(absdir + "/get_dr.ps1", "wt")
    file.write(r'Get-Disk | Get-StorageReliabilityCounter | Select-Object -Property "*" | convertto-json | out-file -Encoding ASCII $Env:TEMP/drives.json')
    file.close()

    file = open(absdir + "/get_sn.ps1", "wt")
    file.write(r'Get-PhysicalDisk | convertto-json | out-file -Encoding ASCII $Env:TEMP\sn.json')
    file.close()
    subprocess.run(['powershell.exe', 'Start-Process powershell.exe -argumentlist "-file", "$Env:TEMP/get_dr.ps1" -verb "runas"'], shell=True)
    subprocess.run(['powershell.exe', 'Start-Process powershell.exe -argumentlist "-file", "$Env:TEMP/get_sn.ps1" -verb "runas"'], shell=True)
    # Necessary for system with lots of drives or UAC enabled so user has time to click yes
    time.sleep(10)
    with open('{}/sn.json'.format(absdir), 'r') as snf:
        sn_dict = json.loads(snf.read())

    with open('{}/drives.json'.format(absdir), 'r') as drf:
        dr_dict = json.loads(drf.read())
    return sn_dict, dr_dict

# Dirty but it works... Takes the raw dicts from powershell() and turns it into a sg friendly format
def sort_data(sn_dict, dr_dict):
    device_ids = []
    serial_numbers = []
    size = []

    for i in sn_dict:
        if i['SerialNumber'] is None:
            i['SerialNumber'] = 0
        if i['Size'] is None:
            i['Size'] = 0
        device_ids.append(i['DeviceId'])
        serial_numbers.append(i['SerialNumber'])
        size.append(round(i['Size'] / 1099511627776, 2))

    for i, j in enumerate(device_ids):
        if len(j) == 1:
            device_ids[i] = "0{}".format(j)

    combined = zip(device_ids, serial_numbers, size)
    sorted_pairs = sorted(combined)
    tuples = zip(*sorted_pairs)
    device_ids, serial_numbers, size = [list(i) for i in tuples]
    device_ids2 = []
    temperatures = []
    power_on_hours = []
    load_unloads = []
    start_stops = []

    for i in dr_dict:
        if i['Temperature'] is None:
            i['Temperature'] = 0
        if i['PowerOnHours'] is None:
            i['PowerOnHours'] = 0
        if i['LoadUnloadCycleCount'] is None:
            i['LoadUnloadCycleCount'] = 0
        if i['StartStopCycleCount'] is None:
            i['StartStopCycleCount'] = 0
        temperatures.append(i['Temperature'])
        power_on_hours.append(round(i['PowerOnHours'] / 8760, 2))
        load_unloads.append(i['LoadUnloadCycleCount'])
        start_stops.append(i['StartStopCycleCount'])
        device_ids2.append(i['DeviceId'])

    for i, j in enumerate(device_ids2):
        if len(j) == 1:
            device_ids2[i] = "0{}".format(j)

    combined = zip(device_ids2, temperatures, power_on_hours, load_unloads, start_stops)
    sorted_pairs = sorted(combined)
    tuples = zip(*sorted_pairs)

    device_ids2, temperatures, power_on_hours, load_unloads, start_stops = [list(i) for i in tuples]
    master_dict = {"0": []}
    for i, device_id in enumerate(device_ids2):
        device_id = int(device_id)
        device_ids2[i] = str(device_id)
        master_dict["0"].append([
            str(device_id),
            serial_numbers[i],
            temperatures[i],
            power_on_hours[i],
            load_unloads[i],
            start_stops[i],
            size[i]
        ])

    return device_ids2, master_dict

# Execute the functions
s, d = powershell()
keys, data = sort_data(s, d)

# A Simple GUI
sg.theme('Dark Tan Blue')
headings = ['ID', 'SN', 'Temp', 'Years On', 'Load Cycles', 'Power Cycles', 'Size TB']
header = [[sg.Text('  ')] + [sg.Text(h, size=(12, 1)) for h in headings]]
window = sg.Window('Disk Info',
                   [
                       [
                           sg.Table(
                               values=data["0"],
                               key="0",
                               headings=headings,
                               num_rows=25,
                               hide_vertical_scroll=True,
                               def_col_width=10,
                               auto_size_columns=False,
                               enable_click_events=True
                           )
                       ]
                   ],
                   font='Courier 12'
                   )
event, values = window.read()
evs = [0, 0, 0, 0, 0, 0, 0]

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):  # If user closed the window
        break
    inputs = data["0"]
    if '+CLICKED+' == event[1]:
        # Magic from Mr. Redinger
        for x in range(7):
            if (-1, x) == event[2]:
                evs[x] = 1 if evs[x] == 0 else 0
                inputs = sorted(inputs, key=itemgetter(x), reverse=(evs[x] != 0))
        window.Element(event[0]).update(values=inputs)

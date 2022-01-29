"""
Objective: launch the applications listed in the table in Business Apps.xlsx
workbook, each in it's designated desktop and wait for the speicifed seconds
to account for loading times (estimated) and animations.

Note: We assume that the required number of virtual desktops is already created 
until we have a method to check and create the required desktops beforehand.
"""

from pandas import read_excel
import subprocess
from time import sleep
import keyboard as kb

# Read the applications data from the designated file
# for configured starts
apps = read_excel('Business Apps.xlsx')

desktops = apps['Destination Desktop'].unique() # List of indeces of desktops utilised
def start_apps(desktop):
    for app, starting_time in apps[['Application Path', 'Wait Period in Secs']][apps['Destination Desktop'] == desktop].values.tolist():
        subprocess.Popen(app, creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_BREAKAWAY_FROM_JOB); sleep(starting_time)

start_apps(desktops[0])
for desktop in desktops[1:]:
    kb.send('ctrl+win+right')
    start_apps(desktop)

for _ in range(len(desktops) - 1):
    kb.send('ctrl+win+left'); sleep(3)
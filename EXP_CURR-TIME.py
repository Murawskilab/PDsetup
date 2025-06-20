"""
By:  Giedrius Puidokas and Siddhartha Saggar \n
Aim: Measure dark current from DUT as a function of time.\n
==================\n
Suggestions:\n
1. This script depends on libraries: SMU.py.\n
2. Install KKeysight software and drivers for controlling the SMU.\n
3. Ensure dark condition: either manually trigger Light-blocker to cut light beam path, or turn-off LD.\n
4. Raw data is saved in a new folder named "Dark Current" within the folder location chosen by the user.\n
"""

from SMU import SMUDevice
from FlipMirror import FlipMirror
from LightBlock import LightBlock
import tkinter as tk
from tkinter import filedialog
import numpy as np
import matplotlib.pyplot as plt
import csv
import time
import os

### USER TO SET/DEFINE VALUES HERE ###
device_name = 'devicename'  # Filename of saved rawdata includes this name. Ensure keeping the name in ' '.
N_pts = 32  # Number of points to be recorded as a function of time. Preferably, be 2 to the power of an integer.
del_t = 0.1  # Time period (in s) between any two adjacent datapoints recorded.
voltage = 0.5 # applied voltage bias in V, where - or + is also dependent on the connections made in the setup.
NPLC = 5  # Number of Power Line Cycles. Duration of 1NPLC = 1/national-powergrid-AC-frequency-in-Hz.
N_meas = 15  # Number of current-time traces to be recorded. Records N_pts at del_t sampling, N_meas times.
save_plots = True  # "true" for plots to be saved as .png files.
show_plots = [True, 0.1]  # "true" for plots to be shown after each measurement. Second number shows duration in s.
###### END OF DATA ENTRY SECTION ######

start_time = time.time()  # Only to keep a check on how long time the script takes to be executed.

# Selecting a folder to save the results
root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()
print("Selected folder path to save results to:", folder_path)
if not folder_path:
    print('File selection cancelled.')
    quit()
# Create folder for results if it doesn't already exist
if not os.path.exists(os.path.join(folder_path, 'Dark Current')):
    os.makedirs(os.path.join(folder_path, 'Dark Current'))
# *******************************************************************

# Device initialization and abbreviating (giving shorthand alias to) instrument-names for ease of command-writing
SMU = SMUDevice()
SMU.connect()
time.sleep(0.3)
SMU.write_command(f":SOURce:VOLTage:LEVel:IMMediate:AMPLitude {voltage}")
time.sleep(0.3)
# *******************************************************************

# Some arrays to store results and some definitions
output_current = []
output_time = []
def show_currenttime_plot(otime, ocurrent, sname):
    plt.plot(otime, ocurrent)
    plt.xlabel('Time (s)')
    plt.ylabel('Current (A)')
    plt.title(f'{device_name} ')
    if save_plots:
        fig_path = os.path.join(folder_path, sname)
        plt.savefig(fig_path)
    if show_plots[0]:
        plt.show(block=False)
        plt.pause(show_plots[1])
        plt.close()
# Detect range function. It is vulnerable to float conversion errors, change to string handling for redundancy
def detect_range(current):
    allowed_ranges = [20e-12, 200e-12, 2e-9, 20e-9, 200e-9, 2e-6, 20e-6, 200e-6, 2e-3, 20e-3]
    detected_range = None
    for current_range in allowed_ranges:  # Loop through all ranges from lowest to highest
        if np.abs(current) <= current_range:
            detected_range = current_range
            break  # stop when correct range found
    if detected_range is None:
        print("Could not detect current range.")
    return detected_range
# *******************************************************************

# Record current-time traces, based on the user-set conditions 
SMU.set_current_range("AUTO")
SMU.measurement_speed("MED")
SMU.trigger_settings(mtype="AINT", count=10)
SMU.initiate("ACQuire")
IRange = detect_range(np.max(SMU.get_current()))
SMU.set_current_range(IRange)
SMU.measurement_speed(NPLC)
SMU.trigger_settings(mtype="TIMer", count=N_pts, period=del_t)
print("Measurement starts now")
for idx in range(N_meas):
    SMU.initiate("ACQuire", timeout=600)
    output_current = SMU.get_current()
    output_time = SMU.get_time()
    print("Measured", device_name, "; for N:", N_pts, " ; del-t:", del_t, " ; version:", idx+1, "/", N_meas,".")
    file_name = f"Dark Current/IT_{device_name}_{voltage}V_{del_t}s_{N_pts}pts_{idx+1}.csv" # for saving the raw data
    file_path = os.path.join(folder_path, file_name)
    with open(file_path, 'w', newline='') as file:
        writer = csv.writer(file, delimiter='\t')
        writer.writerow(["Time", "Current"])  # write header
        for t, c in zip(output_time, output_current):
            writer.writerow([t, c])  # write data
    plot_name = f"{device_name}_{voltage}V_{del_t}s_{N_pts}pts_{idx+1}.png"
    show_currenttime_plot(output_time, output_current, plot_name)
# **********************************************************************

# Disconnect with the instruments
SMU.write_command(":SOURce:VOLTage:LEVel:IMMediate:AMPLitude 0")
time.sleep(0.2)
SMU.disconnect()

duration = time.time() - start_time
print("The script took ", duration, " seconds to run.")
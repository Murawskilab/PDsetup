"""
By: Giedrius Puidokas\n
Aim: Measure current as function of applied voltage bias, from steady-state photodiode DUT.\n
Illumination condition is either dark or CW mode incident light of known wavelength and optical power. \n
==================\n
Suggestions:\n
1. This script depends on libraries: SMU.py.\n
2. Install KKeysight software and drivers for controlling the SMU.\n
3. Steady state illumination condition is not a variable in this experiment, it is to be recorded by the user. \n
4. For dark condition: either manually trigger Light-blocker to cut light beam path, or turn-off LD.\n
"""

from SMU import SMUDevice
import numpy as np
import matplotlib.pyplot as plt
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import csv
import os

### USER TO SET/DEFINE VALUES HERE ###
device_name = 'devicename'  # Filename of saved rawdata includes this name. Ensure keeping the name in ' '.
illum_cond = 'D' # Illumination Condition. Filename of saved rawdata includes this name. Ensure keeping the name in ' '.
V_stt = 1 # Starting value of the applied voltage bias range, in Volts.
V_end = -1 # End value of the applied voltage bias range, in Volts.
N_pts = 201 # Number of points to be recorded in single sweep across the voltage bias range.
save_plots = True  # "true" for plots to be saved as .png files.
show_plots = [True, 10]  # "true" for plots to be shown after each measurement. Second number shows duration in s.
###### END OF DATA ENTRY SECTION ######

start_time = time.time()  # Only to keep a check on how long time the script takes to be executed.

# Device initialization and abbreviating (giving shorthand alias to) instrument-names for ease of command-writing
SMU = SMUDevice()
SMU.connect()
print("Keysight electrometer (SMU) connected.")
SMU.trigger_settings(mtype="AINT")
measurement_speed = "MED"  # Possibilities: SHOR, MED, LONG, *number*
# *******************************************************************

# Selecting a folder to save the results
root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()
print("Selected folder path to save results to:", folder_path)
if not folder_path:
    print('File selection cancelled.')
    quit()
# *******************************************************************

# ******************Measurements***********************************
SMU.vs_function(ftype="SINGle", vstart=V_stt, vend=V_end  , points=N_pts, speed=measurement_speed)
time.sleep(0.2)
SMU.set_current_range("AUTO")
time.sleep(0.2)
SMU.initiate("ALL")     # ACQuire = measurement, TRANsient = source, ALL = both. For IV we need both
source = SMU.get_source()   # This just gets the measured data from Keysight
current = SMU.get_current()

# Create the plot
plt.figure(figsize=(10, 6))
plt.grid(True, which="both")    # "both" probably redundant, too lazy to check
plt.semilogy(source, np.abs(current), 'b-')  # abs(current) to make sure its plottable in log
# Add title and labels
plt.title(f'I-V_meas_{device_name}_{illum_cond}')
plt.xlabel('Source')
plt.ylabel('Current')
if save_plots:
    plot_path = os.path.join(folder_path, f"I-V_meas_{device_name}_{illum_cond}.png")
    plt.savefig(plot_path)
if show_plots[0]:
    plt.show(block=False)
    plt.pause(show_plots[1])
    plt.close()

# Saving rawdata
file_name = f"I-V_meas_{device_name}_{illum_cond}.csv"
file_path = os.path.join(folder_path, file_name)
# Write data to the CSV file
with open(file_path, 'w', newline='') as file:
    writer = csv.writer(file, delimiter='\t')
    writer.writerow(["Source", "Current"])  # write header
    for s, c in zip(source, current):
        writer.writerow([s, c])  # write data
# # ******************************************************************************

# Disconnect with the instruments
SMU.write_command(":SOURce:VOLTage:LEVel:IMMediate:AMPLitude 0")
SMU.disconnect()

duration = time.time() - start_time
print("The script took ", duration, " seconds to run.")

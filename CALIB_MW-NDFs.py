"""
By:  Siddhartha Saggar\n
Aim: Measure transmittance of NDF filter combinations in the Motorized Wheelset for set incident wavelength.\n
==================\n
Suggestions:\n
1. This script depends on libraries: LightBlock.py, TLPM.py, and Wheels.py.\n
2. Wheel_Calibration.txt provides required rotation values for WS corresponding to different filter pairs.\n
3. Install XiLab software package from Standa (WS) with its drivers to control motorized wheelset controller.\n
4. Install Thorlabs optical powermeter related software with its drivers to control its display console.\n
5. The last filter combination listed in Wheel_Calibration.txt is used as the reference (assumed 100% transmittance) \n 
   Ensure this is intentional — typically, this should be the 'no filter' configuration\n
6. OPM has its own filter to provide reliable results, for illumination in 10 micrwatt domain or higher. \n
"""

from Wheels import Filters
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from ctypes import byref,create_string_buffer,c_bool,c_double,c_voidp
from LightBlock import LightBlock
from TLPM import TLPM
import numpy as np
import time
import os
import csv


### USER TO SET/DEFINE VALUES HERE ###
measurement_name = 'devicename'  # Filename of saved rawdata includes this name. Ensure keeping the name in ' '.
number_of_points = 4 # number of points of optcal power measurement at each filter-pair combination
measurement_interval = 1  # time period (in s) between adjacent datapoints of optical power, per filter-conmbination
WL = 532 # enter peak wavelength in nm. Script assumes the incident light to be monochromatic.
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
# ********************************************************************************

# Device initialization and abbreviating (giving shorthand alias to) instrument-names for ease of command-writing
WH = Filters()
print("Motorized-filter-wheel (MW) connected.")
LB = LightBlock()
LB.connect() # initiating connection of the system to instruments/devices.
print("Light blocker (LB) connected.")
tlPM = TLPM()
resourceName = create_string_buffer(b'USB0::0x1313::0x8075::P5001149::INSTR')
tlPM.open(resourceName, c_bool(True), c_bool(True))
print("Optical powermeter (OPM) connected.")
wavelength = c_double(WL)
tlPM.setWavelength(wavelength)
print("Wavelength on OPM set to:", WL, "nm")
# ********************************************************************************

# Creating key Lists
filter_pos = []
move_pos = []
calibration = []
N_meas = number_of_points + 1
all_optical_power = []  # List to store Optical_power lists from each iteration
all_Ttime = []  # List to store Ttime lists from each iteration
averages = []  # List to store averages for each iteration
# ********************************************************************************

# Create folder for saving results if it doesn't already exist
if not os.path.exists(os.path.join(folder_path, 'Results dump')):
    os.makedirs(os.path.join(folder_path, 'Results dump'))
print("'Results dump' folder exists.")
# ********************************************************************************

# Measurement in dark condition
LB.move('block') # block the light beam path, preventing DUT exposure to maximum optical power.
print("No optical signal falling on DUT, now. User asked to remove filter of the optical powermeter (OPM).")
messagebox.showinfo('Script paused', 'For dark condition, remove filter from the OPM. Then, click OK to continue.')
Ttime_dark = []
Optical_power_dark = []
for j in range(N_meas): # Measures optical power in dark, for number_of_points one-by-one
    power = c_double()
    tlPM.measPower(byref(power))
    Optical_power_dark.append(power.value)
    print("Dark measurement", j + 1, "/", N_meas, ":", Optical_power_dark[j], "W")
    Ttime_dark.append(time.time()-start_time)
    time.sleep(measurement_interval)
average_power_dark = np.mean(Optical_power_dark[1:])
print("Average Optical Power in Dark condition:", average_power_dark)
print("Note that this is actually the offset within the OPM.")
D_file_name = f"Results dump/RawData_PM-t_{measurement_name}_{number_of_points}pts_{measurement_interval}s_0-0.csv"
D_file_path = os.path.join(folder_path, D_file_name)
with open(D_file_path, 'w', newline='') as file:
    writer = csv.writer(file, delimiter='\t')
    writer.writerow(["Time", "Optical power"])  # write header
    for t, pwr in zip(Ttime_dark, Optical_power_dark):
        writer.writerow([t, pwr])  # write data
print(f"Data for dark condition saved to {D_file_path}")  # status update
# *******************************************************************

# MEASUREMENT OF MOTORIZED WHEEL's FILTERS START HERE !!
WH.calibrate()  # ensures slot#1 of both wheels for light path. this is the starting-position. check Wheels.py.
WH.move(-87)
print("Now starting for different filter combinations.")  # status update
calib_file_path = "Wheel_Calibration.txt"
if not os.path.exists(calib_file_path):
    print(f"Error: {calib_file_path} not found.")
    quit()
with open(calib_file_path, 'r') as file:
    next(file)  # Skip the header line
    for line in file:
        columns = line.strip().split()  # Split the line into columns
        if len(columns) >= 2:  # Check if the line has enough columns
            filter_pos.append(columns[0])
            move_pos.append(int(columns[1]))
        else:
            print("Problem with number of columns in calibration file (check whitespace rows)")
if not os.path.exists(calib_file_path):
    print(f"Error: {calib_file_path} not found.")
    quit()
print("Filter positions accessed.")
messagebox.showinfo('Script paused', '(low light) Remove filter from the OPM. Then, click OK to continue.')
print('User asked to remove filter from Si-PD, the powermeter.')
LB.move('unblock')
for i in range(len(filter_pos)):
    WH.move(move_pos[i])
    # print("Moving to: ", filter_pos[i])
    if filter_pos[i] == "4":  # CRITICAL FOR NOT ALLOWING HIGHER INTENSITIES TO FALL ON OPM WITHOUT ITS OWN FILTER
        LB.move('block')
        print('User asked to bring-in filter of the OPM.')
        messagebox.showinfo('Script paused', 'Bring-in filter on the OPM. Then, click OK to continue.')
        LB.move('unblock')
    else:
        print("Moving to: ", filter_pos[i])
    Ttime = []
    Optical_power = []
    for j in range(N_meas): # Measures optical power for "number_of_points" points, one-by-one
        power = c_double()
        tlPM.measPower(byref(power))
        Optical_power.append(power.value)
        Ttime.append(time.time() - start_time)
        time.sleep(measurement_interval)
        # *******************************************************************
        # Save Optical power versus time for each filter combination (i)
        file_name = f"Results dump/RawData_PM-t_{measurement_name}_{number_of_points}pts_" \
                    f"{measurement_interval}s_{filter_pos[i]}.csv"
        file_path = os.path.join(folder_path, file_name)
        # Write data to the CSV file
        with open(file_path, 'w', newline='') as file:
            writer = csv.writer(file, delimiter='\t')
            writer.writerow(["Time", "Optical power"])  # write header
            for t, pwr in zip(Ttime, Optical_power):
                writer.writerow([t, pwr])  # write data
   
    # Store the lists
    all_optical_power.append(Optical_power)
    all_Ttime.append(Ttime)
    # Calculate the average for this iteration
    avg_power = ( np.mean(Optical_power[1:]) ) - average_power_dark
    averages.append((avg_power))  # Store average optical power sequentially
    print(f"Iteration {i + 1}: Average Optical Power at {filter_pos[i]} = {avg_power}")
# *******************************************************************

# Calculating trasmitivity of each Filter-wheel combination
if len(averages) > 1:  # Ensure there is more than one element
    last_value = averages[-1]  # Get the last element
    ratios = [val / last_value if val != 0 else None for val in averages]
else:
    print("the AVERAGES matrix is empty. Check files saved in the Results Dump.")
# *******************************************************************

# Saving data of Filter-wheel combination as "Average Optical Power | Filter Combination"
file_name = f"PM_T-ratios_{measurement_name}_{number_of_points}pts_{measurement_interval}s.csv"
file_path = os.path.join(folder_path, file_name)
with open(file_path, 'w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Average Optical Power", "Filter-Combination", "Transmitivity(0to1)"])  # Header
    writer.writerows([(avg_power, filter_pos[i], ratios[i]) for i, avg_power in enumerate(averages)])
print(f"Results saved to {file_path}")
# *******************************************************************

# Disconnect with the instruments
LB.move('block')
LB.disconnect()
WH.disconnect()

duration = time.time() - start_time
print("The script took ", duration, " seconds to run.")
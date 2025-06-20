"""
By: Siddhartha Saggar and Giedrius Puidokas\n
Aim: Record photocurrent as function of incident optical power (CW mode), from steady-state photodiode DUT.\n
Current from DUT is measured in dark and then CW mode illumination, to fetch corresponding photocurrent.\n
==================\n
Suggestions:\n
1. This script depends on libraries: SMU.py, LightBlock.py, Flipmirror.py, TLPM.py, and Wheels.py.\n
2. Also dependent on Wheel_Calibration.txt for NDF transmittance values corresponding to defined wavelength.\n
3. Wheel_Calibration.txt also provides required rotation values for WS corresponding to different filter pairs.\n 
4. Install XiLab software package from Standa (WS) with its drivers to control motorized wheelset controller.\n
5. Install Thorlabs optical powermeter related software with its drivers to control its display console.\n
6. Similarly Keysight's software with drivers for controlling the SMU.\n
"""

from FlipMirror import FlipMirror
from SMU import SMUDevice
from Wheels import Filters
from LightBlock import LightBlock
import tkinter as tk
from tkinter import filedialog
import numpy as np
import matplotlib.pyplot as plt
from ctypes import byref,create_string_buffer,c_bool,c_int16,c_double,c_voidp
from TLPM import TLPM
import math
import time
import os
import csv
import threading

### USER TO SET/DEFINE VALUES HERE ###
device_name = 'devicename'  # Filename of saved rawdata includes this name. Ensure keeping the name in ' '.
wl = 532  # wl=wavelength in nm, with 3 significant digits & no decimals. Script assumes monochromatic light source.
datapoints = 32  # Number of points to be recorded as a function of time. Should be 2 to the power of an integer.
sampling_time = 0.1  # Time period (in s) between any two adjacent datapoints recorded.
voltage = 0.5 # applied voltage bias in V, where - or + is also dependent on the connections made in the setup.
measurement_speed = 5  # denotes NPLC. Duration of 1NPLC = 1/national-powergrid-AC-frequency-in-Hz.
save_plots = True  # "true" for plots to be saved as .png files.
show_plots = [True, 10]  # "true" for plots to be shown after each measurement. Second number shows duration in s.
#######################################
N_d_prior = 6  # Number of measured points to be ignored prior to dark current signal recording.
N_d_after = 6  # Number of measured points to be ignored after the dark current signal recording.
N_i_prior = 6  # Number of measured points to be ignored prior to current signal under illumination recording.
N_i_after = 6  # Number of measured points to be ignored after the current signal under illumination recording.
number_of_measurements = 1  # in number of loops. Repeats the whole experiment again and saves all data uniquely.
###### END OF DATA ENTRY SECTION ######

start_time = time.time()  # Only to keep a check on how long time the script takes to be executed.

#Formulation to calculate total points to be measured for each case of filter-combination
N_dark = N_d_prior + datapoints + N_d_after
# cut_start_pts_dark = ((cut_start_time_dark) / (sampling_time)) + 1
# cut_end_pts_dark = ((cut_end_time_dark) / (sampling_time)) + 1
# cut_start_pts_illum = ((cut_start_time_illum) / (sampling_time)) + 1
# cut_end_pts_illum = ((cut_end_time_illum) / (sampling_time)) + 1
N_illum = N_i_prior + datapoints + N_i_after
total_points = N_dark + N_illum # This is set into SMU setting in line 178, along with sampling_time.
# *******************************************************************

# Importing calibration stuff (If trying to understand the code, check out the file)
filter_pos = []
move_pos = []
calibration = []
wavelength = str(wl) # defining "wavelength" as string, so searching in Wheel_Calibration.txt is possible.
file_path = "Wheel_Calibration.txt"
with open(file_path, 'r') as file:
    next(file)  # Skip the header line
    for line in file:
        columns = line.strip().split()  # Split the line into columns
        if len(columns) >= 3:  # Check if the line has enough columns
            filter_pos.append(columns[0])
            move_pos.append(int(columns[1]))
            if wavelength.lower() == '532':
                calibration.append(float(columns[2]))
            elif wavelength.lower() == '407':
                calibration.append(float(columns[3]))
            elif wavelength.lower() == '639':
                calibration.append(float(columns[4]))
            else:
                print("Problem with acquiring wheel calibration data. Check the file for data related to ", wl,"nm.")
        else:
            print("Problem with number of columns in calibration file (check whitespace rows)")
# *******************************************************************

# Device initialization and abbreviating (giving shorthand alias to) instrument-names for ease of command-writing
SMU = SMUDevice()
WH = Filters()
FM = FlipMirror()
LB = LightBlock()
LB.connect() # initiating connection of the system to instruments/devices.
print("Light blocker (LB) connected.")
FM.connect()
print("Motorized flipper (FM) beamsplitting filter connected.")
SMU.connect()
print("Keysight electrometer (SMU) connected.")
SMU.write_command(f":SOURce:VOLTage:LEVel:IMMediate:AMPLitude {voltage}")

# Initialize optical powermeter connection
tlPM = TLPM()
resourceName = create_string_buffer(b'USB0::0x1313::0x8075::P5001149::INSTR') # specific address of Thorlabs PM400
tlPM.open(resourceName, c_bool(True), c_bool(True))
print("Optical powermeter (OPM) connected.")
OPM_wl = c_double(wl)
tlPM.setWavelength(OPM_wl)
print("Wavelength on OPM set to:", wl, "nm")
tlPM.setPowerUnit(c_int16(0))
print("Unit of optical power set to: Watt (W).")
# *******************************************************************

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

# Selecting a folder to save the results
root = tk.Tk()
root.withdraw()
folder_path = filedialog.askdirectory()
print("Selected folder path to save results to:", folder_path)
if not folder_path:
    print('File selection cancelled.')
    quit()
# Create folder for results if it doesn't already exist
if not os.path.exists(os.path.join(folder_path, 'Results dump')):
    os.makedirs(os.path.join(folder_path, 'Results dump'))
# *******************************************************************


# EXPERIMENTATIONS ARE NOW ON !!!

for meas_num in range(number_of_measurements):
    # 1st step -- Re-orienting the motorized wheetset to 1-1 position
    print("Blocking the light beam path, preventing unrequired exposure of DUT to maximum optical power.")
    LB.move('block') # block the light beam path, preventing DUT exposure to maximum optical power.
    WH.calibrate() # ensures slot#1 on both wheels (No NDFs) in to the light beam path.
    # 2nd step -- Record optical power in dark (OPM_dark)
    print("Beam-splitter moved in to the light beam path.")
    FM.move('on')  # move beam-splitter into the light beam path.
    print('Initiating optical power measurement in dark condition.')
    Optical_power_dark = []
    OPM_dark = []  # optical power (in W) when light-beam blocked by blocker (LB).
    for j in range(11):  # Measures optical power in dark, for number_of_points one-by-one
        power = c_double()
        tlPM.measPower(byref(power))
        Optical_power_dark.append(power.value)
        print("Dark measurement", j + 1, "/11:", Optical_power_dark[j], "W")
        time.sleep(1) # makes the script wait 1 second before taking next record of the optical power.
        OPM_dark = np.mean(Optical_power_dark[1:]) # Helps take average of the last 10 datapoints measured.
    print("Average Optical Power in Dark condition:", OPM_dark, "W")
    print("Technically, this is the offset within the OPM (powermeter).")
    # 3rd step -- Measure initial current from DUT in dark -- only to assess lowest current range
    FM.move('off')  # move beam-splitter out of the light beam path.
    SMU.trigger_settings(mtype="AINT", count=30, period=None)  # initial current measurement -- won't be recorded.
    time.sleep(0.3)  # acts like hold time in s.
    SMU.measurement_speed("MED")
    time.sleep(0.3)  # acts like hold time in s.
    SMU.set_current_range("AUTO")
    time.sleep(0.3)  # acts like hold time in s.
    print("Dummy initiate to stabilize") 
    SMU.initiate('ACQuire', timeout=1000)  # helps in determining the current range.
    I_for_range = SMU.get_current()[-1]  # takes the final point from the measurement (can be changed).
    print(f"Range determination from: {I_for_range} A")  # for check during troubleshooting.
    IRange = detect_range(1.03 * I_for_range)  # the multiplier is used to give some room to avoid overflow.
    SMU.set_current_range(IRange)  # SETS CURRENT RANGE (INITIAL) ON THE SMU.
    print(f"SMU condition. Current range set to: {IRange} A.") 
    SMU.trigger_settings(mtype="TIMer", count=total_points, period=sampling_time)  # SETS N and del(t) ON SMU.
    print(f"SMU condition. Sampling time set to:  {sampling_time} s.")
    print(f"SMU condition. Total points per scan set to: {total_points} pts. (Datapoints reqd.:  {datapoints} pts)")
    SMU.measurement_speed(measurement_speed) # SETS NPLC VALUE ON SMU
    print(f"SMU condition. NPLC set to: {measurement_speed}.")
    need_range_change = False  # a bool, used to check whether range change is needed, so no to do it every time.
    # 4th step -- Record maximum optical power (laser_power)
    print("Unblocking the light beam path to record optical power.")
    LB.move('unblock')  # unblocks the light beam path.
    FM.move('on')  # move beam-splitter into the light beam path, for optical power measurement.
    power_meas_1 = [] # a list to store floating point numbers
    ginti = 0 # Counter for measuring opticalpower. Average of 10 measurements is considered.
    while ginti < 11: # to record optical power 11 times.
        power_1 = c_double()
        tlPM.measPower(byref(power_1))
        power_meas_1.append(power_1.value)
        print("Measurement", ginti + 1, "/11:", power_meas_1[ginti], "W")
        ginti += 1
        time.sleep(1) # hold time (in s) between two adjacent optical power measurements.
    laser_power = np.mean(power_meas_1[1:]) - OPM_dark # Takes average of the last 10 datapoints measured.
    print(f"Mean of max. optical power = {laser_power} W")
    calibration = np.array(calibration)  #
    Pinc = calibration[~np.isnan(calibration)]  # Remove 'nan' values. They are used to skip measurements
    Pinc = np.multiply(Pinc, laser_power)  # Multiplies calculated laser power by transmittance array
    print("Blocking the light beam path.")
    LB.move('block') # block the incident light path to keep DUT in dark.
    FM.move('off')  # move beam-splitter out of the light beam path.

    # Some arrays to store results
    Output_Current = []
    Current_Error = []
    Dark_Current = []
    Dark_Error = []
    Photocurrent = []
    Photocurrent_Error = []

    # Create a figure and axis for concurrent display
    fig, (ax1, ax2) = plt.subplots(2, 1)

    # 5th step -- Loop over each position (see file) of the Motorized Wheelset
    for i in range(len(filter_pos)):
        WH.move(move_pos[i])
        print("Moving to: ", filter_pos[i], "for measurement loop number", meas_num+1)
        # In some cases, for the wheel to reach required position, two moves are needed. To avoid measuring after the
        # first of such moves, NaN is used in transmittance column. When script finds this, it skips the measurement.
        if np.isnan(calibration[i]):
            print("NaN detected - skipping measurement (normal procedure)")
            continue
        if need_range_change:  # This is to prevent sending set range command every time
            SMU.set_current_range(IRange)
            need_range_change = False
        # A single measurement is split into two parts, half of it is in dark, half under illumination.
        # Threading is used to allow two commands run concurrently. It helps in controlling the conditions of
        # dark current measurement while also ensuring timely shutter-movement for measurement under illumination. 
        timer = threading.Timer(sampling_time * (N_dark), lambda: LB.move('unblock'))
        timer.start()
        # print("measurement started")
        SMU.initiate('ACQuire', timeout=1000)
        # print("measurement ended")
        meas_curr = SMU.get_current()
        ttime = SMU.get_time()  # double t to avoid confusing with other functions
        timer.cancel()  # Threading timer has to be defined and stopped every time it is used

        # Checks for overflow, if found, increases the range by 1 order and remeasures. This works best, when
        # the photocurrent measured by SMU under reverse bias is positive magnitude. So, connnect accordingly.
        while np.any(np.isnan(meas_curr)) or any(x > 1 for x in meas_curr):
            print("Overflow detected, repeating measurement with higher range")
            IRange = IRange * 10
            print(IRange)
            SMU.set_current_range(IRange)
            timer = threading.Timer(sampling_time * (N_dark), lambda: LB.move('unblock') )
            timer.start()
            SMU.initiate('ACQuire', timeout=1000)
            meas_curr = SMU.get_current()
            ttime = SMU.get_time()
            timer.cancel()
        LB.move('block')  # blocks the incident light path to keep DUT in dark.
        
        # Calculations
        # Mean dark current is "Dark_Current" here.
        Dark_Current.append(np.mean(meas_curr[int(math.ceil(N_d_prior)):int(math.ceil(N_d_prior + datapoints))])))
        Dark_Error.append((np.std(meas_curr[int(math.ceil(N_d_prior)):int(math.ceil(N_d_prior + datapoints))])))
        # Mean of measured current under illumination is "Output_Current" here.
        Output_Current.append(np.mean(meas_curr[int(math.ceil(N_dark + N_i_prior)):int(
            math.ceil(N_dark + N_i_prior + datapoints))])))
        Current_Error.append(np.std(meas_curr[int(math.ceil(N_dark + N_i_prior)):int(
            math.ceil(N_dark + N_i_prior + datapoints))])))
        # Mean photocurrent is "Photocurrent" here
        Photocurrent.append(Output_Current[-1] - Dark_Current[-1])
        Photocurrent_Error.append(np.sqrt((np.square(Current_Error[-1])) + (np.square(Dark_Error[-1]))))
        if IRange < detect_range(1.2*Output_Current[-1]):  # 1.2 is arbitrary, using previous average for comparison
            IRange = detect_range(1.2*Output_Current[-1])
            print("need range change")
            need_range_change = True
        # *******************************************************************************

        # Section to save file with raw data
        file_name = f"Results dump/Raw data {device_name} {Pinc[-1]}W low_intensity measurement{meas_num+1}" \
                    f"{filter_pos[i]} {voltage}V {measurement_speed} {total_points}pts {sampling_time}s.csv"
        file_path = os.path.join(folder_path, file_name)
        # Write data to the CSV file
        with open(file_path, 'w', newline='') as file:
            writer = csv.writer(file, delimiter='\t')
            writer.writerow(["Time", "Current"])  # write header
            for t, c, in zip(ttime, meas_curr):
                writer.writerow([t, c])  # write data
        # *******************************************************************************

        # Plot the data (updating plot)
        # Clear the axes
        ax1.cla()
        ax2.cla()
        # Plot the new line with error bars
        ax1.errorbar(Pinc[:len(Dark_Current)], Dark_Current, yerr=Dark_Error, label='Dark Current', fmt='o')
        ax1.errorbar(Pinc[:len(Dark_Current)], Output_Current, yerr=Current_Error, label='Light Current', fmt='o')
        ax2.errorbar(Pinc[:len(Dark_Current)], Photocurrent, yerr=Photocurrent_Error, fmt='o')
        # Set log-log or log-linear scale
        ax1.set_xscale('log')
        ax2.set_xscale('log')
        ax2.set_yscale('log')
        ax1.grid(True)
        ax2.grid(True)
        # Auto-scale
        ax1.relim()
        ax1.autoscale_view()
        ax1.legend()
        ax2.relim()
        ax2.autoscale_view()
        fig.canvas.draw() # Redraw the figure
        fig.canvas.flush_events()
        plt.pause(0.1) # Small pause to ensure plot updates
        # *******************************************************************************

    # Measured current vs optical power data
    file_name = f"Low intensity current output {device_name} {voltage}V measurement{meas_num+1}" \
                f" {measurement_speed} {total_points}pts {sampling_time}s.csv"
    file_path = os.path.join(folder_path, file_name)
    # Write data to the CSV file
    with open(file_path, 'w', newline='') as file:
        writer = csv.writer(file, delimiter='\t')
        writer.writerow(["Incident_Power", "Dark_Current", "Dark_Error", "Current", "Current_Error"])  # write header
        for p, d, der, i, er in zip(Pinc, Dark_Current, Dark_Error, Output_Current, Current_Error):
            writer.writerow([p, d, der, i, er])  # write data

    # Photocurrent vs optical power data
    file_name = f"Low intensity photocurrent {device_name} {voltage}V measurement{meas_num+1}" \
                f" {measurement_speed} {total_points}pts {sampling_time}s.csv"
    file_path = os.path.join(folder_path, file_name)
    # Write data to the CSV file
    with open(file_path, 'w', newline='') as file:
        writer = csv.writer(file, delimiter='\t')
        writer.writerow(["Incident_Power", "Photocurrent", "Photocurrent_Error"])  # write header
        for p, ph, er in zip(Pinc, Photocurrent, Photocurrent_Error):
            writer.writerow([p, ph, er])  # write data

    if save_plots:
        file_path = os.path.join(folder_path, f"Low intensity measurement {device_name} measurement{meas_num+1}"
                                              f"{voltage}V {measurement_speed} {total_points}pts.png")
        plt.savefig(file_path)
    if show_plots[0]:
        plt.show(block=False)
        plt.pause(show_plots[1])
        plt.close()

# Disconnect with the instruments
SMU.write_command(":SOURce:VOLTage:LEVel:IMMediate:AMPLitude 0")
FM.disconnect()
SMU.disconnect()
WH.disconnect()
LB.move('block')
LB.disconnect()

duration = time.time() - start_time
print("The script took ", duration, " seconds to run.")
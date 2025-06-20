########################################################
##   This is for Thorlabs Motorized Flipper MFF102/M  ##
##         serial number of product: 37009202         ##
##  Primary goal: automate block/unblock of lightpath ##
##          for use in OPD Measurement Setup.         ##
##    THIS FILE ACTS AS A LIBRARY FOR CONTROLLING     ##
##           THE MOTORIZED FLIPPER 37009202           ##
##             Done by Siddhartha Saggar.             ##
##               designed on 12.03.2025               ##
########################################################

import ctypes
import os
import time

class LightBlock:
    def __init__(self, serial_no='37009202'):  # replace with your device's serial number
        self.serial_no = serial_no.encode('utf-8')
        self.flipper_dll = None

    def connect(self):
        # Path to the Kinesis folder
        path = 'C:\\Program Files\\Thorlabs\\Kinesis'
        os.chdir(path)
        # Load the DLL
        self.flipper_dll = ctypes.cdll.LoadLibrary(
            'C:\\Program Files\\Thorlabs\\Kinesis\\Thorlabs.MotionControl.FilterFlipper.dll')
        # Initialize the device
        self.flipper_dll.TLI_BuildDeviceList()
        # Open the device
        self.flipper_dll.FF_Open(self.serial_no)
        # Start polling the device at 200ms intervals
        self.flipper_dll.FF_StartPolling(self.serial_no, 200)
        # Allow some time for the device to initialize
        time.sleep(1)
        print("Motorized Light-blocker connected")

    def disconnect(self):
        # Stop polling the device
        self.flipper_dll.FF_StopPolling(self.serial_no)
        # Close the device
        self.flipper_dll.FF_Close(self.serial_no)
        print("Motorized Light-blocker disconneted")

    def move(self, command):
        # Get the current position
        # position = self.flipper_dll.FF_GetPosition(self.serial_no)

        # Move to the specified position
        if command == "unblock":
            next_position = 2
        elif command == "block":
            next_position = 1
        else:
            print("Motorized Light-blocker: Invalid command!")
            return

        self.flipper_dll.FF_MoveToPosition(self.serial_no, next_position)

        # Allow some time for the device to move
        time.sleep(0.5)

        # Get the new position
        # position = self.flipper_dll.FF_GetPosition(self.serial_no)

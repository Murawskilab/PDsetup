import urllib
import tempfile
import re
from pyximc import *


class Filters:
    def __init__(self):
        global lib, device_id  # These are basically so that you don't have to input them everywhere

        # Get to folder where libraries live
        cur_dir = os.path.abspath(os.path.dirname(__file__))
        ximc_dir = os.path.join(cur_dir, "ximc")
        ximc_package_dir = os.path.join(ximc_dir, "crossplatform", "wrappers", "python")
        sys.path.append(ximc_package_dir)
        libdir = os.path.join(ximc_dir, "win64")
        if sys.version_info >= (3, 8):
            os.add_dll_directory(libdir)
        else:
            os.environ["Path"] = libdir + ";" + os.environ["Path"]
        result = lib.set_bindy_key(os.path.join(ximc_dir, "win32", "keyfile.sqlite").encode("utf-8"))
        if result != Result.Ok:
            lib.set_bindy_key("keyfile.sqlite".encode("utf-8"))  # Search for the key file in the current directory.
        probe_flags = EnumerateFlags.ENUMERATE_PROBE + EnumerateFlags.ENUMERATE_NETWORK
        enum_hints = b"addr="
        devenum = lib.enumerate_devices(probe_flags, enum_hints)
        dev_count = lib.get_device_count(devenum)
        controller_name = controller_name_t()
        for dev_ind in range(0, dev_count):
            enum_name = lib.get_device_name(devenum, dev_ind)
            result = lib.get_enumerate_device_controller_name(devenum, dev_ind, byref(controller_name))
        flag_virtual = 0
        open_name = None
        if len(sys.argv) > 1:
            open_name = sys.argv[1]
        elif dev_count > 0:
            open_name = lib.get_device_name(devenum, 0)
        elif sys.version_info >= (3, 0):
            self.virtual_controller()
        if not open_name:
            exit(1)
        if type(open_name) is str:
            open_name = open_name.encode()
        print("Wheels Device: " + repr(open_name))
        device_id = lib.open_device(open_name)
        eng = engine_settings_t()
        eng.MicrostepMode = MicrostepMode.MICROSTEP_MODE_FRAC_256

    # This is not used externally
    def virtual_controller(self):
        tempdir = tempfile.gettempdir() + "/testdevice.bin"
        if os.altsep:
            tempdir = tempdir.replace(os.sep, os.altsep)
        uri = urllib.parse.urlunparse(urllib.parse.ParseResult(scheme="file",
                                                               netloc=None, path=tempdir, params=None, query=None,
                                                               fragment=None))
        open_name = re.sub(r'^file', 'xi-emu', uri).encode()
        flag_virtual = 1
        print("The real controller is not found or busy with another app.")
        print("The virtual controller is opened to check the operation of the library.")
        print("If you want to open a real controller, connect it or close the application that uses it.")

    # Not used for now, just leftover. uSteps are not used also (no need for that much precision)
    def get_position(self):
        x_pos = get_position_t()
        result = lib.get_position(device_id, byref(x_pos))
        if result == Result.Ok:
            print("Position: {0} steps, {1} microsteps".format(x_pos.Position, x_pos.uPosition))
        return x_pos.Position, x_pos.uPosition

    def disconnect(self):
        lib.close_device(byref(cast(device_id, POINTER(c_int))))
        print("Wheel disconnected")

    def set_speed(self, speed):
        # Create move settings structure
        mvst = move_settings_t()
        # Get current move settings from controller
        result = lib.get_move_settings(device_id, byref(mvst))
        # Change current speed
        mvst.Speed = int(speed)
        # Write new move settings to controller
        result = lib.set_move_settings(device_id, byref(mvst))
        # Print command return status. It will be 0 if all is OK  # For troubleshooting
        # print("Write command result: " + repr(result))  # For troubleshooting

    # This is your bread and butter. The wheel makes a full 360 in 200 steps. 200 will go one way, -200 the other
    def move(self, position):
        # print("Going to {0}".format(distance))  # For troubleshooting
        lib.command_move(device_id, position)
        self.wait_for_stop(30)
        # print("Result: " + repr(result))  # For troubleshooting

    # Waits for the wheel to stop. Uses time after stop in ms as input
    def wait_for_stop(self, interval):
        # print("\nWaiting for stop")  # For troubleshooting
        result = lib.command_wait_for_stop(device_id, interval)
        # print("Result: " + repr(result))  # For troubleshooting

    # So the following calibrate function moves a lot to one direction, then the other, then to 0.
    def calibrate(self):
        self.move(290)
        self.move(-290)
        self.move(0)


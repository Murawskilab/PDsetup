import pyvisa
import time
import matplotlib.pyplot as plt


class SMUDevice:
    def __init__(self):
        self.rm = pyvisa.ResourceManager()
        self.smu = None

    def connect(self):
        smu_address = "USB0::0x2A8D::0x9B01::MY61390205::0::INSTR"
        self.smu = self.rm.open_resource(smu_address)
        time.sleep(0.1)
        # Send the *IDN? command to the SMU
        self.smu.write("*IDN?")
        time.sleep(1)  # Add a delay before reading the response.
        response = self.smu.read()
        print("SMU Identification:", response)

    def disconnect(self):
        if not self.wait_for_completion():
            print("Warning: Operation did not complete within the timeout.")
        if self.smu:
            self.smu.close()
            self.smu = None
            print("Disconnected from SMU")
        else:
            print("No connection to SMU")

    def trigger_settings(self, mtype=None, count=None, period=None):
        allowed_types = ["AINT", "BUS", "TIMer", "INT1", "INT2",
                         "LAN", "EXT1", "EXT2", "EXT3", "EXT4",
                         "EXT5", "EXT6", "EXT7", "TIN"]

        if mtype is not None and mtype in allowed_types:
            type_str = f":TRIGger:ACQuire:SOURce {mtype}"
            self.smu.write(type_str)

        if count is not None and 1 <= count <= 100000:
            count_str = f":TRIGger:ACQuire:COUNt {count}"
            self.smu.write(count_str)

        if period is not None and period >= 0.0001:
            period_str = f":TRIGger:ACQuire:TIMer {period}"
            self.smu.write(period_str)

    def vs_function(self, ftype=None, vstart=None, vend=None, points=None, speed=None):
        allowed_types = ["single", "double", "off"]

        if ftype.lower() != "off" or None:
            self.smu.write(":sour:volt:mode swe")
            self.smu.write(f":TRIGger:COUNt {1}")

        if ftype is not None and ftype.lower() in allowed_types:
            if ftype.lower() == "off":
                type_str = f":SOURce:VOLTage:MODE FIXED"
            else:
                type_str = f":SOURce:SWEep:STAir {ftype}"
            self.smu.write(type_str)

        if ftype.lower() == "single":
            self.smu.write(f":TRIGger:COUNt {points}")

        if ftype.lower() == "double":
            self.smu.write(f":TRIGger:COUNt {2 * points}")

        if ftype.lower() not in allowed_types:
            print("Warning: invalid measurement type syntax")

        if vstart is not None:
            vstart_str = f":SOUR:VOLT:STAR {vstart}"
            self.smu.write(vstart_str)

        if vend is not None:
            vend_str = f":SOUR:VOLT:STOP {vend}"
            self.smu.write(vend_str)

        if points is not None and 1 <= points <= 100000:
            points_str = f":SOURce:SWEep:POINts {points}"
            self.smu.write(points_str)

        if speed is not None:
            self.measurement_speed(speed)

    def initiate(self, command_type, timeout=1000):
        if command_type == 'ACQuire':
            command = ":INITiate:ACQuire"
        elif command_type == 'TRANsient':
            command = ":INITiate:TRANsient"
        elif command_type == 'ALL':
            command = ":INITiate:ALL"
        else:
            print("Invalid command type.")
            return

        self.write_command(command, timeout)
        if not self.wait_for_completion():
            print("Warning: Operation did not complete within the timeout.")

    def get_current(self):
        current_str = self.smu.query(":fetc:arr:curr?")
        if not self.wait_for_completion():
            print("Warning: Query did not complete within the timeout.")
        current = [float(value) for value in current_str.split(',')]
        return current

    def get_source(self):
        source_str = self.smu.query(":fetc:arr:sour?")
        if not self.wait_for_completion():
            print("Warning: Query did not complete within the timeout.")
        source = [float(value) for value in source_str.split(',')]
        return source

    def get_time(self):
        time_str = self.smu.query(":fetc:arr:time?")
        if not self.wait_for_completion():
            print("Warning: Query did not complete within the timeout.")
        ttime = [float(value) for value in time_str.split(',')]
        return ttime

    def wait_for_completion(self, timeout=400):
        self.smu.write("*OPC?")
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                response = self.smu.read().strip()
                if response == "1":
                    # print("Operation complete.")
                    return True
            except pyvisa.errors.VisaIOError as e:
                if e.error_code == pyvisa.constants.StatusCode.error_timeout:
                    continue  # Ignore timeout errors
                else:
                    raise  # Re-raise any other exceptions
        print("Operation did not complete within the timeout (you can change it in the SMU.py script.")
        return False

    def query_operation_completion(self):
        response = self.smu.query("*OPC?")
        return int(response)

    def check_operation_completion(self):
        response = self.smu.query("*OPC?")
        print(response)
        return int(response) == 1

    def write_command(self, command, timeout=400):
        self.smu.write(command)
        if not self.wait_for_completion(timeout):
            print("Warning: Command did not complete within the timeout.")

    def set_current_range(self, current_range):
        allowed_ranges = [2e-12, 20e-12, 200e-12, 2e-9, 20e-9, 200e-9, 2e-6, 20e-6, 200e-6, 2e-3, 20e-3, "AUTO"]
        if current_range not in allowed_ranges:
            print(f"Warning: The requested range {current_range} is not in the allowed ranges.")

        if current_range == "AUTO":
            command = ":SENS:CURR:DC:RANG:AUTO 1"
        else:
            command = f":SENS:CURR:DC:RANG {current_range}"
        self.smu.write(command)
        if not self.wait_for_completion():
            print("Warning: Operation did not complete within the timeout.")
        print(f"Current measurement range set to {current_range} A.")

    def monitor_current(self, current_range, interval, num_points):
        # Set current range
        self.set_current_range(current_range)

        # Initialize lists for storing time and current data
        time_data = []
        current_data = []

        # Set up plot
        plt.ion()
        fig, ax = plt.subplots()
        line, = ax.plot(time_data, current_data)
        ax.set_xlabel('Time (s)')
        ax.set_ylabel('Current (A)')
        ax.set_title('Time-Current Data')

        start_time = time.time()
        for i in range(num_points):
            # Measure current
            current = self.smu.query(":MEAS:CURR?")
            time.sleep(0.01)
            current = float(current)
            current_data.append(current)
            time_data.append(time.time() - start_time)

            # Update plot
            line.set_xdata(time_data)
            line.set_ydata(current_data)
            ax.relim()
            ax.autoscale_view(True, True, True)
            plt.draw()
            plt.pause(0.01)

            # Wait for the specified interval before next measurement
            time.sleep(interval)

        plt.ioff()
        plt.show()

        return time_data, current_data

    def measurement_speed(self, speed):
        if speed in ["SHOR", "MED", "LONG"]:
            time.sleep(0.1)
            self.smu.write(f":SENS:CURR:APER:AUTO ON")
            time.sleep(0.1)
            self.smu.write(f":SENS:CURR:APER:AUTO:MODE {speed}")
        else:
            try:
                nplc = float(speed)  # check if speed is a valid number
                self.smu.write(f":SENS:CURR:DC:NPLC {nplc}")
                if speed > 100:
                    print("Warning: too slow measurement speed, the instrument will use the max of 100")
                elif speed < 5e-4:
                    print("Warning: too high measurement speed, the instrument will use the min of 5e-4")
            except ValueError:
                print(f"Invalid input: {speed}. Expected 'SHOR', 'MED', 'LONG' or a number.")

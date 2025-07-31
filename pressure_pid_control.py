import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from collections import deque
from datetime import datetime, timedelta
import pyvisa
import time
import openpyxl
import threading
import tkinter as tk

# --- GUI Thread Function for live PID tuning ---
def start_gui():
    def apply_values():
        try:
            globals()['KP'] = float(kp_entry.get())
            globals()['KI'] = float(ki_entry.get())
            globals()['KD'] = float(kd_entry.get())
            print(f"✅ PID updated: KP={KP}, KI={KI}, KD={KD}")
        except ValueError:
            print("⚠️ Invalid PID entry")

    root = tk.Tk()
    root.title("PID Controller Tuning")

    tk.Label(root, text="Kp").grid(row=0, column=0)
    tk.Label(root, text="Ki").grid(row=1, column=0)
    tk.Label(root, text="Kd").grid(row=2, column=0)

    kp_entry = tk.Entry(root)
    ki_entry = tk.Entry(root)
    kd_entry = tk.Entry(root)

    kp_entry.grid(row=0, column=1)
    ki_entry.grid(row=1, column=1)
    kd_entry.grid(row=2, column=1)

    kp_entry.insert(0, str(KP))
    ki_entry.insert(0, str(KI))
    kd_entry.insert(0, str(KD))

    apply_button = tk.Button(root, text="Apply", command=apply_values)
    apply_button.grid(row=3, columnspan=2)

    root.mainloop()

# --- Initial PID Values ---
KP, KI, KD = 1, 0.09, 0.001

# Start GUI in background
gui_thread = threading.Thread(target=start_gui, daemon=True)
gui_thread.start()

# --- VISA Setup ---
rm = pyvisa.ResourceManager()
Ruska = rm.open_resource("GPIB0::1::INSTR")
Ruska.baud_rate = 9600
Ruska.read_termination = '\n'
Ruska.write_termination = '\n'

alicat = rm.open_resource('ASRL5::INSTR')
alicat.baud_rate = 19200
alicat.data_bits = 8
alicat.parity = pyvisa.constants.Parity.none
alicat.stop_bits = pyvisa.constants.StopBits.one
alicat.read_termination = '\r'
alicat.write_termination = '\r'
alicat.timeout = 1000

alicat.write('@@=a')
alicat.write('@@=b')
alicat.write('@@=e')
time.sleep(0.1)

# --- Logging Setup ---
timestamp_now = datetime.now().strftime('%Y%m%d_%H%M%S')
excel_filename = f"PressureControlLog_{timestamp_now}.xlsx"
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Pressure & Flow Log"
ws.append(["Timestamp", "Pressure (Pa)",
           "FlowSet_A", "Actual_A",
           "FlowSet_B", "Actual_B",
           "FlowSet_E", "Actual_E"])

# --- Plot Setup ---
plt.ion()
fig, ax = plt.subplots()
pressure_values = deque(maxlen=200)
timestamps = deque(maxlen=200)
line, = ax.plot([], [], 'b-', label="Pressure (Pa)")
ax.set_title("Live Pressure")
ax.set_xlabel("Time")
ax.set_ylabel("Pressure (Pa)")
ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
ax.legend()

# --- Control Setup ---
TARGET_PRESSURE = 550000
MAX_FLOW_A = 500.0
MIN_FLOW_A = 0.0
MAX_FLOW_B = 10.0
MIN_FLOW_B = 0.0
MAX_FLOW_E = 500.0
MIN_FLOW_E = 5.0  # <- Updated from 0.0
SAVE_INTERVAL = timedelta(minutes=2)
integral = 0.0
previous_error = 0.0
previous_time = time.time()
data_buffer = []
last_save_time = datetime.now()
Ruska.write(f'PRES {TARGET_PRESSURE}')

# --- Console Table Header ---
header_printed = False

try:
    while True:
        current_time = datetime.now()
        timestamp_str = current_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        now = time.time()
        dt = now - previous_time

        # --- Read pressure ---
        Ruska.write("MEAS?")
        try:
            pressure = float(Ruska.read().strip())
        except:
            continue

        # --- PID Calculation with Anti-Windup ---
        error = TARGET_PRESSURE - pressure

        # --- Predict integral growth ---
        predicted_integral = integral + error * dt
        derivative = (error - previous_error) / dt if dt > 0 else 0

        # --- Predict PID output using predicted integral ---
        predicted_output = KP * error + KI * predicted_integral + KD * derivative

        # --- Apply anti-windup logic based on actuator limits ---
        if predicted_output > MAX_FLOW_A:
            pid_output = MAX_FLOW_A
            # --- Only allow integral to decay slowly if output is saturated ---
            integral *= 0.98
        elif predicted_output < -MAX_FLOW_E:
            pid_output = -MAX_FLOW_E
            integral *= 0.98
        else:
            integral = predicted_integral
            pid_output = predicted_output

        previous_error = error
        previous_time = now

        if pid_output >= 0:
            flow_set_A = min(pid_output, MAX_FLOW_A)
            flow_set_B = min(flow_set_A, MAX_FLOW_B)
            flow_set_E = 0.0
        else:
            flow_set_A = 0.0
            flow_set_B = 0.0
            flow_set_E = min(-pid_output, MAX_FLOW_E)

        # Enforce rounding and limits
        flow_set_A = round(max(flow_set_A, MIN_FLOW_A), 1)
        flow_set_B = round(max(flow_set_B, MIN_FLOW_B), 3)  # <- updated for precision
        flow_set_E = round(max(flow_set_E, MIN_FLOW_E), 2)  # <- updated for new min limit

        previous_error = error
        previous_time = now

        # --- Send new setpoints ---
        alicat.write(f'aS{flow_set_A:.1f}')
        time.sleep(0.03)
        alicat.write(f'bS{flow_set_B:.3f}')  # <- updated for precision
        time.sleep(0.03)
        alicat.write(f'eS{flow_set_E:.2f}')
        time.sleep(0.03)

        # --- Read actual flows ---
        try:
            alicat.write('a')
            time.sleep(0.03)
            parts_a = alicat.read().strip().split()
            actual_A = float(parts_a[3]) if len(parts_a) > 3 else None
        except:
            actual_A = None

        try:
            alicat.write('b')
            time.sleep(0.03)
            parts_b = alicat.read().strip().split()
            actual_B = float(parts_b[4]) if len(parts_b) > 4 else None
        except:
            actual_B = None

        try:
            alicat.write('e')
            time.sleep(0.03)
            parts_e = alicat.read().strip().split()
            actual_E = float(parts_e[4]) if len(parts_e) > 4 else None
        except:
            actual_E = None

        # --- Print clean console log ---
        if not header_printed:
            print(f"{'Time':<23} | {'Pressure':>10} | "
                  f"{'A_set':>7}/{ 'A_act':<7} | "
                  f"{'B_set':>7}/{ 'B_act':<7} | "
                  f"{'E_set':>7}/{ 'E_act':<7} | {'Err':>8}")
            print("-" * 90)
            header_printed = True

        print(f"{timestamp_str:<23} | {pressure:10.2f} | "
              f"{flow_set_A:7.1f}/{actual_A if actual_A is not None else '---':<7} | "
              f"{flow_set_B:7.3f}/{actual_B if actual_B is not None else '---':<7} | "  # <- updated
              f"{flow_set_E:7.2f}/{actual_E if actual_E is not None else '---':<7} | "
              f"{error:8.1f}")

        # --- Store data for Excel ---
        data_buffer.append((timestamp_str, pressure,
                            f"{flow_set_A:.1f}", f"{actual_A:.1f}" if actual_A is not None else '',
                            f"{flow_set_B:.3f}", f"{actual_B:.3f}" if actual_B is not None else '',
                            f"{flow_set_E:.2f}", f"{actual_E:.2f}" if actual_E is not None else ''))

        # --- Update plot ---
        pressure_values.append(pressure)
        timestamps.append(current_time)
        line.set_xdata(timestamps)
        line.set_ydata(pressure_values)
        ax.relim()
        ax.autoscale_view()
        fig.autofmt_xdate()
        plt.pause(0.01)

        # --- Periodically save to Excel ---
        if datetime.now() - last_save_time > SAVE_INTERVAL:
            for row in data_buffer:
                ws.append(row)
            wb.save(excel_filename)
            data_buffer.clear()
            last_save_time = datetime.now()

        time.sleep(0.5)

except KeyboardInterrupt:
    print("🛑 Stopped by user.")
finally:
    for row in data_buffer:
        ws.append(row)
    wb.save(excel_filename)
    Ruska.close()
    alicat.close()
    plt.ioff()
    plt.show()

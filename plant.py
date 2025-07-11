# Plant Monitoring System (with Simulation Mode if Arduino not connected)
import ctypes
import sys
import os

def set_app_icon():
    if sys.platform == "win32":
        try:
            myappid = u'plant.monitor.app.001'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception as e:
            print(f"Error setting taskbar icon: {e}")

try:
    import serial
    import tkinter as tk
    from tkinter import ttk
    from ttkbootstrap import Style
    from ttkbootstrap.widgets import Meter
    from datetime import datetime
    import csv
    import matplotlib.pyplot as plt
    import random
    import requests
    import pandas as pd
except ImportError as e:
    print("Missing library:", e)
    print("Please install all required modules:")
    print("pip install pyserial ttkbootstrap matplotlib requests pandas")
    exit()

# --- Try Connecting to Arduino ---
arduino_connected = False
try:
    serial_port = "COM5"  # Change this to match your system
    baud_rate = 9600
    arduino = serial.Serial(serial_port, baud_rate, timeout=1)
    arduino_connected = True
    print(" Arduino connected successfully.")
except serial.SerialException:
    print(" Arduino not connected. Switching to simulation mode.")

# --- Global Variables ---
sensor_data_history = []
unique_dates = set()
thingspeak_api_key = "Q00C5KPTB59DTMJL"
thingspeak_url = "https://api.thingspeak.com/update"

# --- GUI Setup ---
style = Style("darkly")
root = style.master
root.title("Plant Monitoring System")

main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill=tk.BOTH, expand=True)

# Title
title_label = ttk.Label(main_frame, text="Plant Monitoring System", font=("Helvetica", 20, "bold"))
title_label.pack(pady=10)

# Meters
soil_meter = Meter(main_frame, bootstyle="success", subtext="Soil Moisture", textright="%", metertype="full")
soil_meter.pack(side=tk.LEFT, padx=10, pady=10)

temp_meter = Meter(main_frame, bootstyle="info", subtext="Temperature", textright="°C", metertype="full")
temp_meter.pack(side=tk.LEFT, padx=10, pady=10)

humidity_meter = Meter(main_frame, bootstyle="warning", subtext="Humidity", textright="%", metertype="full")
humidity_meter.pack(side=tk.LEFT, padx=10, pady=10)

# Labels for Light and Pump
light_label = ttk.Label(main_frame, text="Light Status: Unknown", font=("Helvetica", 14), bootstyle="secondary")
light_label.pack(pady=10)

pump_label = ttk.Label(main_frame, text="Water Pump: Unknown", font=("Helvetica", 14), bootstyle="secondary")
pump_label.pack(pady=10)

# Buttons
button_frame = ttk.Frame(main_frame, padding=10)
button_frame.pack(fill=tk.X, pady=10)

ttk.Button(button_frame, text="Soil Moisture Graph", command=lambda: plot_graph("soil_moisture")).pack(side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="Temperature Graph", command=lambda: plot_graph("temperature")).pack(side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="Humidity Graph", command=lambda: plot_graph("humidity")).pack(side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="Summary", command=lambda: summarize_data()).pack(side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="Export CSV", command=lambda: export_csv()).pack(side=tk.RIGHT, padx=10)

# --- UI Update Functions ---
def update_light_status(status):
    if status == "Sufficient":
        light_label.configure(text="Light Status: Sufficient", bootstyle="success")
    elif status == "Insufficient":
        light_label.configure(text="Light Status: Insufficient", bootstyle="danger")
    else:
        light_label.configure(text="Light Status: Unknown", bootstyle="secondary")

def update_pump_status(status):
    if status == "On":
        pump_label.configure(text="Water Pump: On", bootstyle="success")
    elif status == "Off":
        pump_label.configure(text="Water Pump: Off", bootstyle="danger")
    else:
        pump_label.configure(text="Water Pump: Unknown", bootstyle="secondary")

def send_to_thingspeak(soil, temp, humid, light_status):
    try:
        response = requests.get(
            thingspeak_url,
            params={
                "api_key": thingspeak_api_key,
                "field1": soil,
                "field2": temp,
                "field3": humid,
                "field4": 1 if light_status == "Sufficient" else 0,
            },
        )
        if response.status_code == 200:
            print(" Data sent to ThingSpeak.")
        else:
            print(f" ThingSpeak error: HTTP {response.status_code}")
    except Exception as e:
        print(f"Error sending to ThingSpeak: {e}")

# --- Duplicate Check ---
def is_duplicate_date(timestamp):
    # date_only = timestamp.split(" ")[0]
    # if date_only in unique_dates:
    #     return True
    # unique_dates.add(date_only)
    return False

# --- Process Data ---
def process_data(data):
    try:
        parts = data.split()

        if arduino_connected:
            soil = int(parts[1])
            temp = float(parts[3])
            humid = float(parts[5])
            light = parts[7]
        else:
            soil = random.randint(200, 900)
            temp = random.uniform(20, 40)
            humid = random.uniform(40, 90)
            light = random.choice(["Sufficient", "Insufficient"])

        soil_percent = round(((1023 - soil) / 850) * 100, 1)
        temp_percent = round(temp, 1)
        humid_percent = round(humid, 1)

        soil_meter.configure(amountused=soil_percent)
        temp_meter.configure(amountused=temp_percent)
        humidity_meter.configure(amountused=humid_percent)

        update_light_status(light)

        pump = "On" if soil_percent < 70 else "Off"
        update_pump_status(pump)

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if is_duplicate_date(timestamp):
            print(" Duplicate date — skipping entry.")
            return

        sensor_data_history.append({
            "timestamp": timestamp,
            "soil_moisture": soil,
            "temperature": temp,
            "humidity": humid,
            "light_status": light,
            "pump_status": pump
        })

        send_to_thingspeak(soil_percent, temp_percent, humid_percent, light)
    except Exception as e:
        print(f" Data processing error: {e}")

# --- Read from Arduino or Simulate ---
def read_from_arduino():
    if arduino_connected and arduino.in_waiting > 0:
        try:
            data = arduino.readline().decode().strip()
            process_data(data)
        except:
            pass
    else:
        simulated_data = "Soil 512 Temp 28 Humidity 60 Light Sufficient"
        process_data(simulated_data)
    root.after(1000, read_from_arduino)

# --- Plot Graph ---
def plot_graph(sensor_type):
    try:
        timestamps = [entry["timestamp"] for entry in sensor_data_history]
        values = [entry[sensor_type] for entry in sensor_data_history]

        plt.style.use("dark_background")
        plt.figure(figsize=(10, 6))
        plt.plot(timestamps, values, marker="o", label=sensor_type.capitalize())
        plt.xlabel("Timestamp")
        plt.ylabel(sensor_type.capitalize())
        plt.title(f"{sensor_type.capitalize()} Over Time")
        plt.xticks(rotation=45, ha="right")
        plt.legend()
        plt.tight_layout()
        plt.show()
    except Exception as e:
        print(f"Graph error: {e}")

# --- Export CSV ---
def export_csv():
    try:
        filename = f"sensor_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        with open(filename, mode="w", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=["timestamp", "soil_moisture", "temperature", "humidity", "light_status", "pump_status"])
            writer.writeheader()
            writer.writerows(sensor_data_history)
        print(f" Data exported to {filename}")
    except Exception as e:
        print(f"CSV export error: {e}")

# --- Summary with Pandas ---
def summarize_data():
    try:
        if not sensor_data_history:
            print("No data to summarize.")
            return
        df = pd.DataFrame(sensor_data_history)
        summary = df[["soil_moisture", "temperature", "humidity"]].describe()
        print("\n📊 Sensor Data Summary:")
        print(summary)
    except Exception as e:
        print(f"Summary error: {e}")

# --- Start App ---
read_from_arduino()
root.iconbitmap("tree.ico")
set_app_icon()
root.mainloop()

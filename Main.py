import tkinter as tk
from tkinter import messagebox
import json
import threading
import subprocess
import time

# Function to load data from total_load.json
def load_data():
    try:
        with open('total_load.json', 'r') as file:
            data = json.load(file)
            return data.get('total_wattage', 50.0), data.get('average_usage_hours', 6.0)
    except (FileNotFoundError, json.JSONDecodeError):
        return 50.0, 6.0

# Function to update the fields with data from total_load.json
def update_fields():
    total_wattage, average_usage_hours = load_data()
    req_power_var.set(total_wattage)
    ave_time_var.set(average_usage_hours)
    root.after(5000, update_fields)  # Auto-update every 5 seconds

# Function to run load.py using threading
def run_load_sched():
    threading.Thread(target=lambda: subprocess.run(['python', 'load.py']), daemon=True).start()

# GUI Setup
root = tk.Tk()
root.title("Load Scheduler")

# Variables
req_power_var = tk.DoubleVar()
ave_time_var = tk.DoubleVar()

# Load initial data
total_wattage, average_usage_hours = load_data()
req_power_var.set(total_wattage)
ave_time_var.set(average_usage_hours)

# Widgets
tk.Label(root, text="Req. Power (W):").grid(row=0, column=0, padx=10, pady=5)
tk.Entry(root, textvariable=req_power_var).grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Ave. Time (hrs):").grid(row=1, column=0, padx=10, pady=5)
tk.Entry(root, textvariable=ave_time_var).grid(row=1, column=1, padx=10, pady=5)

load_sched_btn = tk.Button(root, text="Load Sched", command=run_load_sched)
load_sched_btn.grid(row=2, column=0, columnspan=2, pady=10)

# Auto-update fields
update_fields()

root.mainloop()

import tkinter as tk
from tkinter import ttk
import pandas as pd
import tkinter.messagebox as messagebox
import csv
import math

# Import all predefined values from predefined_values.py
from predefined_values import *

# -------------------------
# Global Variables & Data
# -------------------------
appliance_data = pd.read_csv('Appliances.csv')
total_wattage = 0
total_usage_hours = 0
appliance_count = 0
total_consumption_kWh = 0  # Accumulated appliance kWh consumption

csv_filename = "load_Sched.csv"

# -------------------------
# Function Definitions
# -------------------------

def update_fields(*args):
    """
    Updates the Rated Power field when the Appliance field changes.
    Auto-fills rated power if the appliance exists; otherwise, leaves blank.
    Also triggers recalculation of the solar generation set.
    """
    appliance_name = appliance_var.get()
    if appliance_name in appliance_data['Appliance'].values:
        appliance_info = appliance_data[appliance_data['Appliance'] == appliance_name].iloc[0]
        rated_power_combobox.set(appliance_info['Rated Power (W)'])
    else:
        rated_power_combobox.set("")
    calculate_gen_set()

def save_to_csv():
    """
    Saves the current appliance schedule and solar generation set results to a CSV file.
    """
    try:
        with open(csv_filename, mode='w', newline='') as file:
            writer = csv.writer(file)
            # Write appliance details header
            writer.writerow(
                ["Appliance", "Power (W)", "PF", "Eff(%)", "Surge(W)", "Usage (Hrs)", "Count", "Consumption (kWh)"]
            )
            # Write each appliance row
            for row in tree.get_children():
                writer.writerow(tree.item(row)['values'])
            # Write totals for appliances
            writer.writerow([])
            writer.writerow(["Total Wattage", total_wattage])
            avg_usage_time = total_usage_hours / appliance_count if appliance_count else 0
            writer.writerow(["Average Usage Time", f"{avg_usage_time:.2f} hrs"])
            writer.writerow(["Total Consumption (kWh)", f"{total_consumption_kWh:.4f}"])

            # Write solar gen set results summary
            writer.writerow([])
            writer.writerow(["Solar Gen Set Summary", summary_label.cget("text")])
            writer.writerow([])
            writer.writerow(["Solar Component", "Requirement/Selection", "Details"])
            for item in solar_tree.get_children():
                writer.writerow(solar_tree.item(item)['values'])
    except PermissionError:
        messagebox.showerror("File Error", f"Permission denied when trying to write to '{csv_filename}'. "
                                             "Please ensure the file is not open in another application and that you have write permissions.")


def recalc_totals():
    """
    Recalculates the overall totals from the Treeview data,
    updates summary labels, recalculates the solar gen set,
    and saves the CSV file.
    """
    global total_wattage, total_usage_hours, appliance_count, total_consumption_kWh
    total_wattage = 0
    total_usage_hours = 0
    appliance_count = 0
    total_consumption_kWh = 0

    for row in tree.get_children():
        values = tree.item(row)['values']
        try:
            power = float(str(values[1]).replace(",", ""))
            usage = float(values[5])
            count = float(values[6])
            consumption = float(str(values[7]).replace(",", ""))
        except Exception:
            continue
        total_wattage += power * count
        total_usage_hours += usage * count
        appliance_count += count
        total_consumption_kWh += consumption

    total_wattage_label.config(text=f"Total Wattage: {total_wattage:,.0f} W")
    avg_usage_time = total_usage_hours / appliance_count if appliance_count else 0
    average_usage_time_label.config(text=f"Avg. Usage: {avg_usage_time:.2f} hrs")
    total_consumption_label.config(text=f"Total Consumption: {total_consumption_kWh:,.4f} kWh")

    calculate_gen_set()
    save_to_csv()

def add_appliance():
    """
    Adds an appliance entry to the Treeview and recalculates totals.
    """
    appliance = appliance_var.get()
    try:
        rated_power = float(rated_power_combobox.get())
    except ValueError:
        messagebox.showwarning("Input Error", "Please enter a valid number for Rated Power (W).")
        return

    # Check if the appliance exists in the CSV data.
    if appliance in appliance_data['Appliance'].values:
        appliance_info = appliance_data[appliance_data['Appliance'] == appliance].iloc[0]
        surge_power = float(appliance_info['Surge Power (W)'])
        efficiency = float(appliance_info['Efficiency (%)']) / 100
        power_factor = float(appliance_info['Power Factor (PF)'])
    else:
        surge_power = rated_power  # For custom appliances, assume surge equals rated power.
        efficiency = 1.0           # 100% efficiency.
        power_factor = 1.0         # Unity power factor.

    try:
        usage_hours = float(usage_hours_combobox.get())
    except ValueError:
        usage_hours = 6  # default usage hours

    try:
        appliance_count_input = round(float(counts_combobox.get()))
    except ValueError:
        appliance_count_input = 1

    # Calculate consumption in kWh (adjusted for power factor and efficiency)
    consumption = rated_power * usage_hours * appliance_count_input * power_factor * efficiency / 1000

    # Format numbers for display
    rated_power_formatted = f"{rated_power:,}"
    consumption_formatted = f"{consumption:,.4f}"

    tree.insert("", "end", values=(
        appliance,
        rated_power_formatted,
        f"{power_factor:.2f}",
        f"{efficiency * 100:.0f}",
        f"{surge_power:,}",
        usage_hours,
        appliance_count_input,
        consumption_formatted
    ))
    recalc_totals()

def delete_selected():
    """
    Deletes selected rows from the Treeview and recalculates totals.
    """
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showinfo("Delete", "No item selected for deletion.")
        return
    for item in selected_items:
        tree.delete(item)
    recalc_totals()

def on_combobox_keyrelease(event):
    """
    Filters appliance names in the Appliance combobox as the user types.
    """
    typed_value = appliance_var.get()
    if typed_value:
        filtered_appliances = [item for item in appliance_data['Appliance']
                               if isinstance(item, str) and typed_value.lower() in item.lower()]
    else:
        filtered_appliances = appliance_data['Appliance'].tolist()
    appliance_combobox['values'] = filtered_appliances

def on_tree_select(event):
    """
    When a Treeview row is selected, populates the input fields.
    """
    selected_items = tree.selection()
    if selected_items:
        item_values = tree.item(selected_items[0])['values']
        appliance_var.set(item_values[0])
        rated_power_combobox.set(item_values[1].replace(",", ""))
        usage_hours_combobox.set(item_values[5])
        counts_combobox.set(item_values[6])
    calculate_gen_set()

def on_tree_double_click(event):
    """
    Enables inline editing for Appliance and Rated Power cells.
    """
    region = tree.identify("region", event.x, event.y)
    if region != "cell":
        return

    col = tree.identify_column(event.x)
    row = tree.identify_row(event.y)
    if not row:
        return
    col_num = int(col.replace("#", "")) - 1
    if col_num not in (0, 1):
        return

    x, y, width, height = tree.bbox(row, col)
    current_value = tree.item(row, "values")[col_num]
    entry = tk.Entry(tree)
    entry.place(x=x, y=y, width=width, height=height)
    entry.insert(0, current_value)
    entry.focus()

    def on_focus_out(event):
        new_value = entry.get()
        current_values = list(tree.item(row, "values"))
        if col_num == 1:
            try:
                new_val_float = float(new_value)
                current_values[1] = f"{new_val_float:,}"
            except ValueError:
                messagebox.showwarning("Input Error", "Please enter a valid number for Rated Power (W).")
                entry.destroy()
                return
        else:
            current_values[col_num] = new_value
        tree.item(row, values=current_values)
        entry.destroy()
        recalc_totals()

    entry.bind("<FocusOut>", on_focus_out)
    entry.bind("<Return>", lambda event: on_focus_out(event))

def calculate_gen_set():
    """
    Calculates the Solar Generation Set requirements based on current appliance loads and solar parameters.
    Displays results in the solar_tree and summary_label.
    """
    # Clear previous results
    for item in solar_tree.get_children():
        solar_tree.delete(item)

    if total_consumption_kWh <= 0 or total_wattage <= 0:
        summary_label.config(text="")
        return

    try:
        system_voltage = float(system_voltage_var.get())
        dod = float(dod_var.get())
        solar_efficiency = float(solar_eff_var.get())
        panel_size = float(panel_size_var.get())
    except ValueError:
        messagebox.showerror("Input Error", "Please ensure all solar parameters are valid numbers.")
        return

    # Define constants and margins
    sun_hours = 5.5             # Average daily sun hours
    battery_margin = 1.20       # 20% extra battery capacity margin
    inverter_margin = 1.25      # 25% extra inverter capacity margin
    pv_margin = 1.20            # 20% extra PV capacity margin

    daily_consumption_Wh = total_consumption_kWh * 1000
    battery_Ah_required = (daily_consumption_Wh * battery_margin) / (system_voltage * (dod / 100))

    inverter_required = total_wattage * inverter_margin
    inverter_selected = next((inv for inv in sorted(INVERTER_SIZES) if inv >= inverter_required), None)
    if inverter_selected is None:
        inverter_selected = f"> {max(INVERTER_SIZES)}"

    pv_capacity_required = (daily_consumption_Wh * pv_margin) / (sun_hours * (solar_efficiency / 100))
    num_panels = math.ceil(pv_capacity_required / panel_size)
    total_pv_capacity = num_panels * panel_size

    current_per_panel = panel_size / system_voltage
    total_pv_current = num_panels * current_per_panel

    mppt_selected = next((mppt for mppt in sorted(MPPT_SIZES) if mppt >= total_pv_current), None)
    if mppt_selected is None:
        mppt_selected = f"> {max(MPPT_SIZES)}"

    scc_selected = next((scc for scc in sorted(SCC_SIZES) if scc >= total_pv_current), None)
    if scc_selected is None:
        scc_selected = f"> {max(SCC_SIZES)}"

    # --- Refined DC Circuit Breaker Calculation ---
    # Use a 1.20 safety margin for the total PV current.
    dc_breaker_required = total_pv_current * 1.20
    dc_breaker_selected = next((dc for dc in sorted(DC_BREAKER_SIZES) if dc >= dc_breaker_required), None)
    if dc_breaker_selected is None:
        dc_breaker_selected = f"> {max(DC_BREAKER_SIZES)}"

    # --- Refined AC Breaker Calculation ---
    # Calculate the inverter AC current (assuming 230V AC output) and apply 1.25 margin.
    inverter_ac_current = (inverter_selected / 230) if isinstance(inverter_selected, (int, float)) else (inverter_required / 230)
    ac_breaker_required = inverter_ac_current * 1.25
    ac_breaker_selected = next((ac for ac in sorted(BREAKER_SIZES) if ac >= ac_breaker_required), None)
    if ac_breaker_selected is None:
        ac_breaker_selected = f"> {max(BREAKER_SIZES)}"

    # --- Refined Cable (Wire) Sizing ---
    # Calculate the current from the inverter output, then apply a 1.25 margin.
    inverter_current = (inverter_selected / system_voltage) if isinstance(inverter_selected, (int, float)) else (inverter_required / system_voltage)
    cable_required = inverter_current * 1.25
    cable_selected = None
    for size in sorted(CABLE_SIZES):
        if AMPACITY_RATING.get(size, 0) >= cable_required:
            cable_selected = size
            break
    if cable_selected is None:
        cable_selected = f"> {max(CABLE_SIZES)}"

    # Populate solar_tree with calculation results
    solar_tree.insert("", "end", values=(
        "Daily Consumption (Wh)", f"{daily_consumption_Wh:,.0f}", "From appliance loads"
    ))
    solar_tree.insert("", "end", values=(
        "Battery Capacity (Ah)", f"{battery_Ah_required:,.0f}", f"{system_voltage}V, DOD: {dod}%"
    ))
    solar_tree.insert("", "end", values=(
        "Inverter Size (W)", f"{inverter_selected}", f"Required: {inverter_required:,.0f}W"
    ))
    solar_tree.insert("", "end", values=(
        "PV Array Capacity (W)", f"{total_pv_capacity:,.0f}", f"{num_panels} panels @ {panel_size}W each"
    ))
    solar_tree.insert("", "end", values=(
        "MPPT Controller (A)", f"{mppt_selected}", f"Total PV Current: {total_pv_current:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "Charge Controller (A)", f"{scc_selected}", f"Total PV Current: {total_pv_current:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "DC Circuit Breaker (A)", f"{dc_breaker_selected}", f"Required: {dc_breaker_required:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "AC Circuit Breaker (A)", f"{ac_breaker_selected}", f"Inverter AC Current: {inverter_ac_current:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "Cable Size (mm²)", f"{cable_selected}", f"Required Ampacity for {cable_required:.2f}A"
    ))

    # Update summary label with key details
    summary_text = (
        f"Battery: {system_voltage}V, {battery_Ah_required:,.0f}Ah | "
        f"Panels: {num_panels} ({total_pv_capacity:,.0f}W total) | "
        f"Inverter: {inverter_selected}W | "
        f"MPPT: {mppt_selected}A | "
        f"SCC: {scc_selected}A"
    )
    summary_label.config(text=summary_text)

# -------------------------
# Tkinter Root Window & Layout
# -------------------------
root = tk.Tk()
root.title("Appliance Power Consumption & Solar Gen Set Calculator")
root.geometry("1200x750")

root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=2)

# -------------------------
# Appliance Section UI
# -------------------------
top_frame = ttk.Frame(root, padding="5")
top_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

appliance_label = ttk.Label(top_frame, text="Appliance:")
appliance_label.grid(row=0, column=0, padx=5, pady=2, sticky="w")
appliance_var = tk.StringVar()
appliance_combobox = ttk.Combobox(top_frame, textvariable=appliance_var,
                                  values=appliance_data['Appliance'].tolist(), width=20)
appliance_combobox.grid(row=0, column=1, padx=5, pady=2)
appliance_var.set(appliance_data.iloc[0]['Appliance'])
appliance_var.trace_add("write", update_fields)
appliance_combobox.bind("<<ComboboxSelected>>", lambda event: calculate_gen_set())
appliance_combobox.bind("<KeyRelease>", on_combobox_keyrelease)

rated_power_label = ttk.Label(top_frame, text="Rated Power (W):")
rated_power_label.grid(row=0, column=2, padx=5, pady=2, sticky="w")
rated_power_combobox = ttk.Combobox(top_frame, width=8)
rated_power_combobox.grid(row=0, column=3, padx=5, pady=2)
rated_power_combobox.set(appliance_data.iloc[0]['Rated Power (W)'])

usage_hours_label = ttk.Label(top_frame, text="Usage Hours:")
usage_hours_label.grid(row=0, column=4, padx=5, pady=2, sticky="w")
usage_hours_combobox = ttk.Combobox(top_frame, values=[str(i) for i in range(1, 25)], width=5)
usage_hours_combobox.grid(row=0, column=5, padx=5, pady=2)
usage_hours_combobox.set("6")
usage_hours_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
usage_hours_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

counts_label = ttk.Label(top_frame, text="Count:")
counts_label.grid(row=0, column=6, padx=5, pady=2, sticky="w")
counts_combobox = ttk.Combobox(top_frame, values=[str(i) for i in range(1, 21)], width=5)
counts_combobox.grid(row=0, column=7, padx=5, pady=2)
counts_combobox.set("1")
counts_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
counts_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

add_button = ttk.Button(top_frame, text="Add", command=add_appliance, width=8)
add_button.grid(row=0, column=8, padx=5, pady=2)

delete_button = ttk.Button(top_frame, text="Delete", command=delete_selected, width=8)
delete_button.grid(row=0, column=9, padx=5, pady=2)

total_wattage_label = ttk.Label(top_frame, text="Total Wattage: 0 W")
total_wattage_label.grid(row=0, column=10, padx=5, pady=2, sticky="w")
average_usage_time_label = ttk.Label(top_frame, text="Avg. Usage: 0 hrs")
average_usage_time_label.grid(row=0, column=11, padx=5, pady=2, sticky="w")
total_consumption_label = ttk.Label(top_frame, text="Total Consumption: 0 kWh")
total_consumption_label.grid(row=0, column=12, padx=5, pady=2, sticky="w")

tree = ttk.Treeview(root,
                    columns=("Appliance", "Power (W)", "PF", "Eff(%)", "Surge(W)", "Usage (Hrs)", "Count", "Consumption (kWh)"),
                    show="headings", selectmode="extended", height=8)
tree.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
for col in tree["columns"]:
    tree.heading(col, text=col)
    tree.column(col, width=120, anchor="center")
tree.bind("<<TreeviewSelect>>", on_tree_select)
tree.bind("<Double-1>", on_tree_double_click)

# -------------------------
# Solar Gen Set Section UI
# -------------------------
solar_frame = ttk.LabelFrame(root, text="Solar Gen Set Requirements", padding="5")
solar_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)

system_voltage_label = ttk.Label(solar_frame, text="System Voltage (V):")
system_voltage_label.grid(row=0, column=0, padx=5, pady=2, sticky="w")
system_voltage_var = tk.StringVar()
system_voltage_combobox = ttk.Combobox(solar_frame, textvariable=system_voltage_var,
                                       values=[str(v) for v in sorted(VOLTAGES)],
                                       width=8)
system_voltage_combobox.grid(row=0, column=1, padx=5, pady=2)
system_voltage_combobox.set("24")
system_voltage_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
system_voltage_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

dod_label = ttk.Label(solar_frame, text="Depth of Discharge (%):")
dod_label.grid(row=0, column=2, padx=5, pady=2, sticky="w")
dod_var = tk.StringVar()
dod_combobox = ttk.Combobox(solar_frame, textvariable=dod_var,
                            values=[str(d) for d in DOD],
                            width=8)
dod_combobox.grid(row=0, column=3, padx=5, pady=2)
dod_combobox.set("50")
dod_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
dod_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

solar_eff_label = ttk.Label(solar_frame, text="Solar Efficiency (%):")
solar_eff_label.grid(row=0, column=4, padx=5, pady=2, sticky="w")
solar_eff_var = tk.StringVar()
solar_eff_combobox = ttk.Combobox(solar_frame, textvariable=solar_eff_var,
                                  values=[str(e) for e in SOLAR_EFFICIENCY],
                                  width=8)
solar_eff_combobox.grid(row=0, column=5, padx=5, pady=2)
solar_eff_combobox.set("22")
solar_eff_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
solar_eff_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

panel_size_label = ttk.Label(solar_frame, text="Solar Panel Size (W):")
panel_size_label.grid(row=0, column=6, padx=5, pady=2, sticky="w")
panel_size_var = tk.StringVar()
panel_size_combobox = ttk.Combobox(solar_frame, textvariable=panel_size_var,
                                   values=[str(s) for s in PANEL_SIZES],
                                   width=8)
panel_size_combobox.grid(row=0, column=7, padx=5, pady=2)
panel_size_combobox.set("100")
panel_size_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
panel_size_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

solar_tree = ttk.Treeview(solar_frame,
                          columns=("Component", "Requirement/Selection", "Details"),
                          show="headings", height=12)
solar_tree.grid(row=1, column=0, columnspan=8, padx=5, pady=5, sticky="nsew")
for col in solar_tree["columns"]:
    solar_tree.heading(col, text=col)
    solar_tree.column(col, width=220, anchor="center")

summary_frame = ttk.Frame(solar_frame, padding="5")
summary_frame.grid(row=2, column=0, columnspan=8, sticky="ew", padx=5, pady=5)
summary_label = ttk.Label(summary_frame, text="", font=("Arial", 11, "bold"))
summary_label.pack(fill="x")

solar_frame.grid_rowconfigure(1, weight=1)
solar_frame.grid_columnconfigure(0, weight=1)

# -------------------------
# Start the Application
# -------------------------
root.mainloop()

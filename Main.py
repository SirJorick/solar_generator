import tkinter as tk
from tkinter import ttk
import pandas as pd
import tkinter.messagebox as messagebox
import csv
import math
import os
import sys
import webbrowser

# Import all predefined values from predefined_values.py
from predefined_values import *

# -------------------------
# Global Variables & Data
# -------------------------
appliance_data = pd.read_csv('Appliances.csv')
total_wattage = 0
total_usage_hours = 0
appliance_count = 0
total_consumption_kWh = 0  # Accumulated energy consumption in kWh

csv_filename = "load_Sched.csv"

# Global variables for computed sizing and drawing
_battery_Ah_req = None
_num_panels = None
_total_pv_capacity = None
_inverter_sel = None
_mppt_sel = None
_scc_sel = None
_active_balancer_sel = None
_fuse_sel = None
_system_voltage = None
_dc_breaker_sel = None
_ac_breaker_sel = None
_cable_sel = None


# -------------------------
# Function Definitions
# -------------------------
def update_fields(*args):
    """
    Updates the Rated Power field when the Appliance field changes.
    Auto-fills rated power if the appliance exists; otherwise, leaves blank.
    Then triggers recalculation of the solar generation set.
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
            writer.writerow(
                ["Appliance", "Power (W)", "PF", "Eff(%)", "Surge(W)", "Usage (Hrs)", "Count", "Consumption (kWh)"])
            for row in tree.get_children():
                writer.writerow(tree.item(row)['values'])
            writer.writerow([])
            writer.writerow(["Total Consumption (kWh)", f"{total_consumption_kWh:,.4f}"])
            writer.writerow([])
            writer.writerow(["Solar Gen Set Summary", summary_label.cget("text")])
            writer.writerow([])
            writer.writerow(["Solar Component", "Requirement/Selection", "Details"])
            for item in solar_tree.get_children():
                writer.writerow(solar_tree.item(item)['values'])
    except PermissionError:
        messagebox.showerror("File Error",
                             f"Permission denied when trying to write to '{csv_filename}'. Please ensure the file is not open in another application and that you have write permissions.")


def recalc_totals():
    """
    Recalculates the overall totals from the Treeview data,
    updates solar generation set calculations, and saves the CSV file.
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

    calculate_gen_set()
    save_to_csv()


def add_appliance():
    """
    Adds an appliance entry to the Treeview and recalculates totals.
    Consumption is calculated as: (rated_power * usage_hours * count) / 1000.
    """
    appliance = appliance_var.get()
    try:
        rated_power = float(rated_power_combobox.get())
    except ValueError:
        messagebox.showwarning("Input Error", "Please enter a valid number for Rated Power (W).")
        return

    if appliance in appliance_data['Appliance'].values:
        appliance_info = appliance_data[appliance_data['Appliance'] == appliance].iloc[0]
        surge_power = float(appliance_info['Surge Power (W)'])
        power_factor = float(appliance_info['Power Factor (PF)'])
        efficiency = float(appliance_info['Efficiency (%)'])
    else:
        surge_power = rated_power
        power_factor = 1.0
        efficiency = 100

    try:
        usage_hours = float(usage_hours_combobox.get())
    except ValueError:
        usage_hours = 6

    try:
        appliance_count_input = round(float(counts_combobox.get()))
    except ValueError:
        appliance_count_input = 1

    consumption = rated_power * usage_hours * appliance_count_input / 1000
    rated_power_formatted = f"{rated_power:,}"
    consumption_formatted = f"{consumption:,.4f}"

    tree.insert("", "end", values=(
        appliance,
        rated_power_formatted,
        f"{power_factor:.2f}",
        f"{efficiency:.0f}",
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
    Enables inline editing for Appliance (col 0), Rated Power (col 1),
    and Usage Hours (col 5). Recalculates consumption (col 7) after editing.
    """
    region = tree.identify("region", event.x, event.y)
    if region != "cell":
        return

    col = tree.identify_column(event.x)
    row = tree.identify_row(event.y)
    if not row:
        return
    col_num = int(col.replace("#", "")) - 1
    if col_num not in (0, 1, 5):
        return

    x, y, width, height = tree.bbox(row, col)
    current_value = tree.item(row, "values")[col_num]
    entry = tk.Entry(tree)
    entry.place(x=x, y=y, width=width, height=height)
    entry.insert(0, current_value)
    entry.focus()

    def on_focus_out(event):
        new_value = entry.get().strip()
        current_values = list(tree.item(row, "values"))
        if col_num == 1:
            try:
                new_val_float = float(new_value)
                current_values[1] = f"{new_val_float:,}"
            except ValueError:
                messagebox.showwarning("Input Error", "Please enter a valid number for Rated Power (W).")
                entry.destroy()
                return
        elif col_num == 5:
            try:
                usage = float(new_value)
                current_values[5] = usage
            except ValueError:
                messagebox.showwarning("Input Error", "Please enter a valid number for Usage Hours.")
                entry.destroy()
                return
            try:
                rated_power = float(str(current_values[1]).replace(",", ""))
                count = float(current_values[6])
                consumption = rated_power * usage * count / 1000
                current_values[7] = f"{consumption:,.4f}"
            except Exception as e:
                messagebox.showwarning("Calculation Error", f"Error recalculating consumption: {e}")
                entry.destroy()
                return
        else:
            current_values[col_num] = new_value
        tree.item(row, values=current_values)
        entry.destroy()
        recalc_totals()

    entry.bind("<FocusOut>", on_focus_out)
    entry.bind("<Return>", lambda event: on_focus_out(event))


def draw_setup():
    """
    Draws a detailed schematic of the solar generation set.
    """
    if total_consumption_kWh <= 0 or total_wattage <= 0:
        messagebox.showerror("Draw Error", "No data available to draw the solar setup.")
        return

    calculate_gen_set()
    draw_setup_figure()


def draw_setup_figure():
    """
    Uses matplotlib to draw a schematic diagram of the solar generation set.
    The drawing is re-centered with larger boxes for the solar panel array and load,
    and an additional "Outlets" box is added between the AC breaker and the load.
    The cable (wire) size is explicitly labeled.
    After saving, the file is directly opened.
    """
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle, FancyArrowPatch

    fig, ax = plt.subplots(figsize=(16, 10))
    ax.set_xlim(0, 22)
    ax.set_ylim(0, 18)
    ax.axis('off')

    # --- Define Component Positions and Sizes ---
    # DC Side Components:
    solar_pos = (2, 15)  # Solar Panels (bigger box)
    solar_size = (3, 1.5)
    mppt_pos = (2, 13)
    mppt_size = (2, 1)
    cc_pos = (2, 11)
    cc_size = (2, 1)
    dc_breaker_pos = (5, 11.5)
    dc_breaker_size = (1, 0.8)
    battery_pos = (7, 10)
    battery_size = (3, 3)

    # AC Side Components:
    active_balancer_pos = (12, 15)
    active_balancer_size = (1.5, 1)
    fuse_pos = (12, 13)
    fuse_size = (1.5, 1)
    inverter_pos = (12, 8)
    inverter_size = (3, 2)
    ac_breaker_pos = (16, 8.5)
    ac_breaker_size = (1, 1)
    outlets_pos = (18, 7.5)  # Outlets box
    outlets_size = (2, 1.5)
    load_pos = (21, 5)  # Load (bigger box)
    load_size = (3, 3)

    # --- Draw Components ---
    solar_box = Rectangle(solar_pos, *solar_size, fc='yellow', alpha=0.7)
    ax.add_patch(solar_box)
    ax.text(solar_pos[0] + solar_size[0] / 2, solar_pos[1] + solar_size[1] / 2,
            f"Solar Panels\n{_num_panels} panels\nTotal: {_total_pv_capacity:,} W",
            ha='center', va='center', fontsize=10, fontweight='bold')

    mppt_box = Rectangle(mppt_pos, *mppt_size, fc='lightblue', alpha=0.7)
    ax.add_patch(mppt_box)
    ax.text(mppt_pos[0] + mppt_size[0] / 2, mppt_pos[1] + mppt_size[1] / 2,
            f"MPPT\n{_mppt_sel} A", ha='center', va='center', fontsize=10, fontweight='bold')

    cc_box = Rectangle(cc_pos, *cc_size, fc='lightgreen', alpha=0.7)
    ax.add_patch(cc_box)
    ax.text(cc_pos[0] + cc_size[0] / 2, cc_pos[1] + cc_size[1] / 2,
            f"Charge Ctrl\n{_scc_sel} A", ha='center', va='center', fontsize=10, fontweight='bold')

    dc_box = Rectangle(dc_breaker_pos, *dc_breaker_size, fc='gold', alpha=0.7)
    ax.add_patch(dc_box)
    ax.text(dc_breaker_pos[0] + dc_breaker_size[0] / 2, dc_breaker_pos[1] + dc_breaker_size[1] / 2,
            f"DC Breaker\n{_dc_breaker_sel} A", ha='center', va='center', fontsize=9, fontweight='bold')

    battery_box = Rectangle(battery_pos, *battery_size, fc='orange', alpha=0.7)
    ax.add_patch(battery_box)
    ax.text(battery_pos[0] + battery_size[0] / 2, battery_pos[1] + battery_size[1] / 2,
            f"Battery Bank\n{_battery_Ah_req:,.0f} Ah @ {float(_system_voltage):.0f} V",
            ha='center', va='center', fontsize=10, fontweight='bold')

    ab_box = Rectangle(active_balancer_pos, *active_balancer_size, fc='violet', alpha=0.7)
    ax.add_patch(ab_box)
    ax.text(active_balancer_pos[0] + active_balancer_size[0] / 2, active_balancer_pos[1] + active_balancer_size[1] / 2,
            f"Balancer\n{_active_balancer_sel} A", ha='center', va='center', fontsize=10, fontweight='bold')

    fuse_box = Rectangle(fuse_pos, *fuse_size, fc='pink', alpha=0.7)
    ax.add_patch(fuse_box)
    ax.text(fuse_pos[0] + fuse_size[0] / 2, fuse_pos[1] + fuse_size[1] / 2,
            f"Fuse\n{_fuse_sel} A", ha='center', va='center', fontsize=10, fontweight='bold')

    inverter_box = Rectangle(inverter_pos, *inverter_size, fc='red', alpha=0.7)
    ax.add_patch(inverter_box)
    ax.text(inverter_pos[0] + inverter_size[0] / 2, inverter_pos[1] + inverter_size[1] / 2,
            f"Inverter\n{_inverter_sel} W", ha='center', va='center', fontsize=10, fontweight='bold')

    ac_box = Rectangle(ac_breaker_pos, *ac_breaker_size, fc='salmon', alpha=0.7)
    ax.add_patch(ac_box)
    ax.text(ac_breaker_pos[0] + ac_breaker_size[0] / 2, ac_breaker_pos[1] + ac_breaker_size[1] / 2,
            f"AC Breaker\n{_ac_breaker_sel} A", ha='center', va='center', fontsize=10, fontweight='bold')

    outlets_box = Rectangle(outlets_pos, *outlets_size, fc='lightgrey', alpha=0.7)
    ax.add_patch(outlets_box)
    ax.text(outlets_pos[0] + outlets_size[0] / 2, outlets_pos[1] + outlets_size[1] / 2,
            "Outlets", ha='center', va='center', fontsize=10, fontweight='bold')

    load_box = Rectangle(load_pos, *load_size, fc='gray', alpha=0.7)
    ax.add_patch(load_box)
    ax.text(load_pos[0] + load_size[0] / 2, load_pos[1] + load_size[1] / 2,
            f"Load\n{total_wattage:,} W", ha='center', va='center', fontsize=10, fontweight='bold')

    # --- Draw Connection Arrows with Labels ---
    arrow1 = FancyArrowPatch((solar_pos[0] + solar_size[0], solar_pos[1] + solar_size[1] / 2),
                             (mppt_pos[0] + mppt_size[0], mppt_pos[1] + mppt_size[1] / 2),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow1)
    ax.text((solar_pos[0] + solar_size[0] + mppt_pos[0] + mppt_size[0]) / 2,
            solar_pos[1] + solar_size[1] / 2 + 0.5, "DC", fontsize=8, va='center')

    arrow2 = FancyArrowPatch((mppt_pos[0] + mppt_size[0] / 2, mppt_pos[1]),
                             (cc_pos[0] + cc_size[0] / 2, cc_pos[1] + cc_size[1]),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow2)
    ax.text(mppt_pos[0] + mppt_size[0] / 2,
            (mppt_pos[1] + cc_pos[1] + cc_size[1]) / 2, "DC", fontsize=8, va='center')

    arrow3 = FancyArrowPatch((cc_pos[0] + cc_size[0], cc_pos[1] + cc_size[1] / 2),
                             (dc_breaker_pos[0], dc_breaker_pos[1] + dc_breaker_size[1] / 2),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow3)
    ax.text((cc_pos[0] + cc_size[0] + dc_breaker_pos[0]) / 2,
            cc_pos[1] + cc_size[1] / 2, "DC", fontsize=8, va='center')

    arrow4 = FancyArrowPatch((dc_breaker_pos[0] + dc_breaker_size[0], dc_breaker_pos[1] + dc_breaker_size[1] / 2),
                             (battery_pos[0], battery_pos[1] + battery_size[1] / 2),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow4)
    ax.text((dc_breaker_pos[0] + dc_breaker_size[0] + battery_pos[0]) / 2,
            battery_pos[1] + battery_size[1] / 2, "DC", fontsize=8, va='center')

    arrow5 = FancyArrowPatch((battery_pos[0] + battery_size[0], battery_pos[1] + battery_size[1] * 0.8),
                             (active_balancer_pos[0], active_balancer_pos[1] + active_balancer_size[1] / 2),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow5)
    ax.text((battery_pos[0] + battery_size[0] + active_balancer_pos[0]) / 2,
            battery_pos[1] + battery_size[1] * 0.8, "DC", fontsize=8, va='center')

    arrow6 = FancyArrowPatch((battery_pos[0] + battery_size[0], battery_pos[1] + battery_size[1] * 0.5),
                             (fuse_pos[0], fuse_pos[1] + fuse_size[1] / 2),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow6)
    ax.text((battery_pos[0] + battery_size[0] + fuse_pos[0]) / 2,
            battery_pos[1] + battery_size[1] * 0.5, "DC", fontsize=8, va='center')

    arrow7 = FancyArrowPatch((fuse_pos[0] + fuse_size[0], fuse_pos[1] + fuse_size[1] / 2),
                             (inverter_pos[0] + inverter_size[0] / 2, inverter_pos[1] + inverter_size[1]),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow7)
    ax.text((fuse_pos[0] + fuse_size[0] + inverter_pos[0] + inverter_size[0] / 2) / 2,
            (fuse_pos[1] + inverter_pos[1] + inverter_size[1]) / 2, "AC", fontsize=8, va='center')

    arrow8 = FancyArrowPatch((inverter_pos[0] + inverter_size[0], inverter_pos[1] + inverter_size[1] / 2),
                             (ac_breaker_pos[0], ac_breaker_pos[1] + ac_breaker_size[1] / 2),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow8)
    ax.text((inverter_pos[0] + inverter_size[0] + ac_breaker_pos[0]) / 2,
            inverter_pos[1] + inverter_size[1] / 2, "AC", fontsize=8, va='center')

    arrow9 = FancyArrowPatch((ac_breaker_pos[0] + ac_breaker_size[0], ac_breaker_pos[1] + ac_breaker_size[1] / 2),
                             (outlets_pos[0], outlets_pos[1] + outlets_size[1] / 2),
                             arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow9)
    ax.text((ac_breaker_pos[0] + ac_breaker_size[0] + outlets_pos[0]) / 2,
            ac_breaker_pos[1] + ac_breaker_size[1] / 2, "AC", fontsize=8, va='center')

    arrow10 = FancyArrowPatch((outlets_pos[0] + outlets_size[0], outlets_pos[1] + outlets_size[1] / 2),
                              (load_pos[0], load_pos[1] + load_size[1] / 2),
                              arrowstyle='->', mutation_scale=15, color='black')
    ax.add_patch(arrow10)
    ax.text((outlets_pos[0] + outlets_size[0] + load_pos[0]) / 2,
            outlets_pos[1] + outlets_size[1] / 2, "AC", fontsize=8, va='center')

    arrow11 = FancyArrowPatch((battery_pos[0] + battery_size[0] / 2, battery_pos[1]),
                              (inverter_pos[0] + inverter_size[0] / 2, inverter_pos[1] + inverter_size[1]),
                              arrowstyle='->', mutation_scale=15, color='blue', linestyle='--')
    ax.add_patch(arrow11)
    ax.text((battery_pos[0] + battery_size[0] / 2 + inverter_pos[0] + inverter_size[0] / 2) / 2,
            (battery_pos[1] + inverter_pos[1] + inverter_size[1]) / 2, "DC Backup", fontsize=8, va='center',
            color='blue')

    ax.text((fuse_pos[0] + fuse_size[0] + inverter_pos[0] + inverter_size[0] / 2) / 2,
            inverter_pos[1] + inverter_size[1] + 0.5, f"Cable: {_cable_sel} mm²", fontsize=9, va='center', color='blue')

    try:
        filename = "Solar_Setup.png"
        fig.savefig(filename, dpi=150, bbox_inches='tight')
        plt.close(fig)
        open_image(filename)
    except Exception as e:
        messagebox.showerror("Save Error", f"Error saving the drawing: {e}")
        plt.close(fig)


def open_image(filepath):
    """
    Opens the image file using the default image viewer.
    """
    try:
        if sys.platform.startswith('win'):
            os.startfile(os.path.abspath(filepath))
        elif sys.platform.startswith('darwin'):
            os.system(f'open "{os.path.abspath(filepath)}"')
        else:
            os.system(f'xdg-open "{os.path.abspath(filepath)}"')
    except Exception as e:
        messagebox.showerror("Open Error", f"Error opening the image file: {e}")


def calculate_gen_set():
    """
    Calculates the Solar Generation Set requirements based on appliance loads and solar parameters.
    Stores computed values globally and populates the solar_tree and summary_label.

    Refinements:
      - Battery Capacity (Ah) = (Daily Consumption (Wh) * battery_margin) / (System Voltage * (DoD/100))
      - Inverter Required (W) = Total Appliance Wattage * inverter_margin
      - PV Array Sizing:
            Performance Ratio (PR) = 0.8 (typical)
            PV Capacity Required (W) = (Daily Consumption (Wh) * pv_margin) / (Sun Hours * PR)
            Number of Panels = ceil(PV Capacity Required / Panel Size)
      - The other components are selected based on the calculated currents with additional margins.
    """
    global total_consumption_kWh, total_wattage
    global _battery_Ah_req, _num_panels, _total_pv_capacity, _inverter_sel, _mppt_sel, _scc_sel
    global _active_balancer_sel, _fuse_sel, _system_voltage, _dc_breaker_sel, _ac_breaker_sel, _cable_sel

    for item in solar_tree.get_children():
        solar_tree.delete(item)

    if total_consumption_kWh <= 0 or total_wattage <= 0:
        summary_label.config(text="")
        return

    try:
        _system_voltage = float(system_voltage_var.get())
        dod = float(dod_var.get())
        panel_size = float(panel_size_var.get())
    except ValueError:
        messagebox.showerror("Input Error", "Please ensure all solar parameters are valid numbers.")
        return

    # Constants and margins
    sun_hours = 5.5  # Average daily sun hours
    battery_margin = 1.20  # 20% extra battery capacity margin
    inverter_margin = 1.25  # 25% extra inverter capacity margin
    pv_margin = 1.20  # 20% extra PV capacity margin
    performance_ratio = 0.8  # Typical system performance ratio

    daily_consumption_Wh = total_consumption_kWh * 1000
    _battery_Ah_req = (daily_consumption_Wh * battery_margin) / (_system_voltage * (dod / 100))

    inverter_required = total_wattage * inverter_margin
    _inverter_sel = next((inv for inv in sorted(INVERTER_SIZES) if inv >= inverter_required), None)
    if _inverter_sel is None:
        _inverter_sel = f"> {max(INVERTER_SIZES)}"

    # Revised PV capacity calculation using performance ratio
    pv_capacity_required = (daily_consumption_Wh * pv_margin) / (sun_hours * performance_ratio)
    _num_panels = math.ceil(pv_capacity_required / panel_size)
    _total_pv_capacity = _num_panels * panel_size

    current_per_panel = panel_size / _system_voltage
    total_pv_current = _num_panels * current_per_panel

    _mppt_sel = next((mppt for mppt in sorted(MPPT_SIZES) if mppt >= total_pv_current), None)
    if _mppt_sel is None:
        _mppt_sel = f"> {max(MPPT_SIZES)}"

    _scc_sel = next((scc for scc in sorted(SCC_SIZES) if scc >= total_pv_current), None)
    if _scc_sel is None:
        _scc_sel = f"> {max(SCC_SIZES)}"

    dc_breaker_required = total_pv_current * 1.20
    _dc_breaker_sel = next((dc for dc in sorted(DC_BREAKER_SIZES) if dc >= dc_breaker_required), None)
    if _dc_breaker_sel is None:
        _dc_breaker_sel = f"> {max(DC_BREAKER_SIZES)}"

    inverter_ac_current = (_inverter_sel / 230) if isinstance(_inverter_sel, (int, float)) else (
                inverter_required / 230)
    ac_breaker_required = inverter_ac_current * 1.25
    _ac_breaker_sel = next((ac for ac in sorted(BREAKER_SIZES) if ac >= ac_breaker_required), None)
    if _ac_breaker_sel is None:
        _ac_breaker_sel = f"> {max(BREAKER_SIZES)}"

    inverter_current = (_inverter_sel / _system_voltage) if isinstance(_inverter_sel, (int, float)) else (
                inverter_required / _system_voltage)
    cable_required = inverter_current * 1.25
    cable_selected = None
    for size in sorted(CABLE_SIZES):
        if AMPACITY_RATING.get(size, 0) >= cable_required:
            cable_selected = size
            break
    if cable_selected is None:
        cable_selected = f"> {max(CABLE_SIZES)}"
    _cable_sel = cable_selected

    active_balancer_required = _battery_Ah_req * 0.05
    active_balancer_required = max(active_balancer_required, 5)
    _active_balancer_sel = next((bal for bal in sorted(ACTIVE_BALANCER_SIZES) if bal >= active_balancer_required), None)
    if _active_balancer_sel is None:
        _active_balancer_sel = f"> {max(ACTIVE_BALANCER_SIZES)}"

    fuse_required = inverter_current * 1.25
    _fuse_sel = next((f for f in sorted(FUSE_SIZES) if f >= fuse_required), None)
    if _fuse_sel is None:
        _fuse_sel = f"> {max(FUSE_SIZES)}"

    solar_tree.insert("", "end", values=(
        "Daily Consumption (Wh)", f"{daily_consumption_Wh:,.0f}", "From appliance loads"
    ))
    solar_tree.insert("", "end", values=(
        "Battery Capacity (Ah)", f"{_battery_Ah_req:,.0f}", f"{_system_voltage}V, DOD: {dod}%"
    ))
    solar_tree.insert("", "end", values=(
        "Inverter Size (W)", f"{_inverter_sel}", f"Required: {inverter_required:,.0f}W"
    ))
    solar_tree.insert("", "end", values=(
        "PV Array Capacity (W)", f"{_total_pv_capacity:,.0f}", f"{_num_panels} panels @ {panel_size}W each"
    ))
    solar_tree.insert("", "end", values=(
        "MPPT Controller (A)", f"{_mppt_sel}", f"Total PV Current: {total_pv_current:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "Charge Controller (A)", f"{_scc_sel}", f"Total PV Current: {total_pv_current:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "DC Circuit Breaker (A)", f"{_dc_breaker_sel}", f"Required: {dc_breaker_required:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "AC Circuit Breaker (A)", f"{_ac_breaker_sel}", f"Inverter AC Current: {inverter_ac_current:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "Cable Size (mm²)", f"{_cable_sel}", f"Required Ampacity for {cable_required:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "Active Balancer (A)", f"{_active_balancer_sel}", f"Required: {active_balancer_required:.2f}A"
    ))
    solar_tree.insert("", "end", values=(
        "Fuse Size (A)", f"{_fuse_sel}", f"Required: {fuse_required:.2f}A"
    ))

    summary_text = (
        f"Battery: {_system_voltage}V, {_battery_Ah_req:,.0f}Ah | "
        f"Panels: {_num_panels} ({_total_pv_capacity:,}W total) | "
        f"Inverter: {_inverter_sel}W | "
        f"MPPT: {_mppt_sel}A | "
        f"SCC: {_scc_sel}A | "
        f"Balancer: {_active_balancer_sel}A | "
        f"Fuse: {_fuse_sel}A"
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
                                  values=appliance_data['Appliance'].tolist(), width=45)
appliance_combobox.grid(row=0, column=1, padx=5, pady=2)
appliance_var.set(appliance_data.iloc[0]['Appliance'])
appliance_var.trace_add("write", update_fields)
appliance_combobox.bind("<<ComboboxSelected>>", lambda event: calculate_gen_set())
appliance_combobox.bind("<KeyRelease>", on_combobox_keyrelease)

rated_power_label = ttk.Label(top_frame, text="Rated Power (W):")
rated_power_label.grid(row=0, column=2, padx=5, pady=2, sticky="w")
rated_power_combobox = ttk.Combobox(top_frame, width=10)
rated_power_combobox.grid(row=0, column=3, padx=5, pady=2)
rated_power_combobox.set(appliance_data.iloc[0]['Rated Power (W)'])

usage_hours_label = ttk.Label(top_frame, text="Usage Hours:")
usage_hours_label.grid(row=0, column=4, padx=5, pady=2, sticky="w")
usage_hours_combobox = ttk.Combobox(top_frame, values=[str(i) for i in range(1, 25)], width=8)
usage_hours_combobox.grid(row=0, column=5, padx=5, pady=2)
usage_hours_combobox.set("6")
usage_hours_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
usage_hours_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

counts_label = ttk.Label(top_frame, text="Count:")
counts_label.grid(row=0, column=6, padx=5, pady=2, sticky="w")
counts_combobox = ttk.Combobox(top_frame, values=[str(i) for i in range(1, 21)], width=8)
counts_combobox.grid(row=0, column=7, padx=5, pady=2)
counts_combobox.set("1")
counts_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
counts_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

add_button = ttk.Button(top_frame, text="Add", command=add_appliance, width=8)
add_button.grid(row=0, column=8, padx=5, pady=2)

delete_button = ttk.Button(top_frame, text="Delete", command=delete_selected, width=8)
delete_button.grid(row=0, column=9, padx=5, pady=2)

draw_button = ttk.Button(top_frame, text="Draw", command=draw_setup, width=8)
draw_button.grid(row=0, column=10, padx=5, pady=2)

tree = ttk.Treeview(root,
                    columns=(
                    "Appliance", "Power (W)", "PF", "Eff(%)", "Surge(W)", "Usage (Hrs)", "Count", "Consumption (kWh)"),
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

# Removed Solar Efficiency combobox entirely

panel_size_label = ttk.Label(solar_frame, text="Solar Panel Size (W):")
panel_size_label.grid(row=0, column=4, padx=5, pady=2, sticky="w")
panel_size_var = tk.StringVar()
panel_size_combobox = ttk.Combobox(solar_frame, textvariable=panel_size_var,
                                   values=[str(s) for s in PANEL_SIZES],
                                   width=8)
panel_size_combobox.grid(row=0, column=5, padx=5, pady=2)
panel_size_combobox.set("100")
panel_size_combobox.bind("<<ComboboxSelected>>", lambda event: recalc_totals())
panel_size_combobox.bind("<KeyRelease>", lambda event: recalc_totals())

solar_tree = ttk.Treeview(solar_frame,
                          columns=("Component", "Requirement/Selection", "Details"),
                          show="headings", height=12)
solar_tree.grid(row=1, column=0, columnspan=6, padx=5, pady=5, sticky="nsew")
for col in solar_tree["columns"]:
    solar_tree.heading(col, text=col)
    solar_tree.column(col, width=220, anchor="center")

summary_frame = ttk.Frame(solar_frame, padding="5")
summary_frame.grid(row=2, column=0, columnspan=6, sticky="ew", padx=5, pady=5)
summary_label = ttk.Label(summary_frame, text="", font=("Arial", 11, "bold"))
summary_label.pack(fill="x")

solar_frame.grid_rowconfigure(1, weight=1)
solar_frame.grid_columnconfigure(0, weight=1)

# -------------------------
# Start the Application
# -------------------------
root.mainloop()

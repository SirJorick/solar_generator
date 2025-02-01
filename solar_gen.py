import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import subprocess


from country_data import COUNTRY  # Import country data from country_data.py
from solar_components import (
    INVERTER_SIZES, BREAKER_SIZES, VOLTAGES, SOLAR_EFFICIENCY,
    PANEL_SIZES, USAGE_HOURS, DC_BREAKER_SIZES, MPPT_SIZES, SCC_SIZES,
    CABLE_SIZES, WIRE_RESISTANCES, AMPACITY_RATING
)

# Function to calculate wire size based on current (DC side and AC side)
def calculate_wire_size(current, is_dc=True):
    wire_sizes = CABLE_SIZES
    wire_resistances = WIRE_RESISTANCES if is_dc else WIRE_RESISTANCES

    for size in wire_sizes:
        ampacity = AMPACITY_RATING.get(size, 0)
        if current <= ampacity:
            return size, wire_resistances[size]
    raise ValueError("No suitable wire size found for the required current.")

# Function to calculate power harvested
def calculate_power_harvested(panel_size, solar_irradiance, efficiency, country_efficiency, harvest_length):
    return (
        panel_size * solar_irradiance * efficiency / 100 * country_efficiency / 100 * harvest_length
    )

# Correct Panel Harvest Calculation
def calculate_power_harvested(panel_size, solar_irradiance, panel_efficiency, country_efficiency, harvest_length):
    # Panel size in watts, solar irradiance in kWh/m²/day, efficiency percentages for the panel and location
    harvested_power = panel_size * (solar_irradiance * panel_efficiency / 100) * (country_efficiency / 100) * harvest_length
    return harvested_power


# Function to calculate battery amp-hours
def calculate_battery_ah(wattage, usage_hours, voltage, battery_efficiency, dod):
    return (wattage * usage_hours) / (voltage * (battery_efficiency / 100) * (dod / 100))

# Function to get the correct inverter size
def get_correct_inverter(wattage):
    for size in INVERTER_SIZES:
        if size >= wattage:
            return size
    raise ValueError("No suitable inverter size found. Consider higher capacity inverters.")

# Function to calculate AC-side breaker size dynamically based on the selected country
def calculate_ac_breaker(inverter_size, selected_country):
    country_data = COUNTRY.get(selected_country, {})
    ac_voltage = country_data.get('AC', 230)  # Default to 230V if no AC value is found

    ac_current = inverter_size / ac_voltage  # Calculate current in amps

    for breaker in BREAKER_SIZES:
        if breaker >= ac_current:
            return breaker
    raise ValueError("No suitable AC-side breaker size found. Consider higher capacity breakers.")

# Function to calculate DC-side breaker size
def calculate_dc_breaker(panel_wattage, panel_voltage):
    dc_current = panel_wattage / panel_voltage  # Calculate current in amps
    for breaker in DC_BREAKER_SIZES:
        if breaker >= dc_current:
            return breaker
    raise ValueError("No suitable DC-side breaker size found. Consider higher capacity breakers.")

# Function to calculate MPPT size (in Amps)
def calculate_mppt_size(total_power, panel_voltage):
    mppt_size = total_power / panel_voltage  # Calculate the current (in Amps)

    for size in MPPT_SIZES:
        if size >= mppt_size:
            return size
    raise ValueError("No suitable MPPT size found. Consider higher capacity MPPTs.")

# Function to calculate SCC size
def calculate_scc(total_power, voltage):
    scc_current = total_power / voltage  # Calculate current (in Amps)
    for size in SCC_SIZES:
        if size >= scc_current:
            return size
    raise ValueError("No suitable SCC size found. Consider higher capacity SCCs.")

# Function to reset input fields
def reset_inputs():
    wattage_input.delete(0, tk.END)
    wattage_input.insert(0, "800")
    country_combo.set("Philippines")
    voltage_combo.set(12)
    efficiency_combo.set(17)
    usage_combo.set(8)
    dod_combo.set(50)
    panel_size_combo.set(100)

# Helper function to validate inputs
def validate_input(value, field_name, min_value=None, max_value=None):
    try:
        value = float(value)
        if min_value is not None and value < min_value:
            raise ValueError(f"{field_name} must be at least {min_value}.")
        if max_value is not None and value > max_value:
            raise ValueError(f"{field_name} must not exceed {max_value}.")
        return value
    except ValueError:
        raise ValueError(f"Invalid value for {field_name}. Please enter a numeric value.")

# Function to display data
def display_data():
    try:
        # Get selected country and its data
        selected_country = country_combo.get()
        data = COUNTRY.get(selected_country, {})

        # Validate and get the input values
        wattage = validate_input(wattage_input.get(), "Input Wattage", min_value=1)
        selected_voltage = validate_input(voltage_combo.get(), "Voltage")
        selected_efficiency = validate_input(efficiency_combo.get(), "Solar Panel Efficiency", 10, 25)
        selected_usage_hours = validate_input(usage_combo.get(), "Usage Hours", min_value=1, max_value=24)
        selected_dod = validate_input(dod_combo.get(), "Depth of Discharge", 10, 100)
        selected_panel_size = validate_input(panel_size_combo.get(), "Panel Size", min_value=50, max_value=500)

        # Extract solar data from the selected country
        solar_irradiance = data.get('solar_irradiance')
        harvest_length = data.get('harvest_length')
        country_efficiency = data.get('efficiency')

        # Validate solar data availability
        if solar_irradiance is None or harvest_length is None:
            raise ValueError(f"Missing solar data for {selected_country}.")

        # Calculate power harvested per panel
        power_harvested = calculate_power_harvested(
            selected_panel_size, solar_irradiance, selected_efficiency, country_efficiency, harvest_length
        )

        # Calculate total number of panels required
        if power_harvested > 0:
            panels_required = -(-wattage // power_harvested)  # Ceiling division
            total_harvested_power = power_harvested * panels_required
        else:
            panels_required = "N/A"
            total_harvested_power = "N/A"

        # Calculate battery AH
        total_battery_ah = calculate_battery_ah(
            wattage, selected_usage_hours, selected_voltage, 85, selected_dod
        )
        total_battery_power = total_battery_ah * selected_voltage  # Battery power in Watt-hours

        # Get the correct inverter size
        correct_inverter = get_correct_inverter(wattage)

        # Calculate AC-side breaker size
        ac_breaker_size = calculate_ac_breaker(correct_inverter, selected_country)

        # Calculate DC-side breaker size
        dc_breaker_size = calculate_dc_breaker(wattage, selected_voltage)

        # Calculate MPPT size
        mppt_size = calculate_mppt_size(total_harvested_power, selected_voltage)

        # Calculate SCC size
        recommended_scc = calculate_scc(total_harvested_power, selected_voltage)

        # Calculate wire sizes
        dc_wire_size, dc_resistance = calculate_wire_size(wattage / selected_voltage, is_dc=True)
        ac_wire_size, ac_resistance = calculate_wire_size(correct_inverter / data.get('AC', 230), is_dc=False)

        # Prepare the result for display
        result_lines = [
            f"Country: {selected_country}",
            f"Optimal Orientation: {data.get('optimal_orientation', 'N/A')}",
            f"Efficiency: {data.get('efficiency', 'N/A')}%",
            f"Inclination: {data.get('inclination', 'N/A')}",
            f"Harvest Length: {harvest_length} months",
            f"Solar Irradiance: {solar_irradiance} kWh/m²",
            f"AC Voltage: {data.get('AC', 230)}V",
            f"Frequency: {data.get('HZ', 'N/A')}Hz",
            f"Seasonal Period: {data.get('seasonal_period', 'N/A')}",
            f"Weather Conditions: {data.get('weather_conditions', 'N/A')}",
            "",
            f"Voltage: {selected_voltage:,.1f}V",
            f"Efficiency: {selected_efficiency:,.1f}%",
            f"Usage Hours: {selected_usage_hours:,.1f} hrs",
            f"DOD: {selected_dod:,.1f}%",
            f"Panel Size: {selected_panel_size:,.1f}W",
            "",
            f"Input Wattage: {wattage:,.1f}W",
            f"Harvest/Panel: {power_harvested:,.1f}W",
            f"Total Panels Required: {panels_required}",
            f"Total Harvested Wattage: {total_harvested_power:,.1f}W",
            "",
            f"Total Battery AH: {total_battery_ah:,.1f} AH",
            f"Total Battery Power: {total_battery_power:,.1f} Wh",
            "",
            f"Recommended Inverter Size: {correct_inverter}",
            f"AC-side Breaker Size: {ac_breaker_size}A",
            f"DC-side Breaker Size: {dc_breaker_size}A",
            f"Recommended MPPT Size: {mppt_size}A",
            f"Recommended SCC Size: {recommended_scc}A",
            "",
            f"DC Wire Size: {dc_wire_size} mm² (Resistance: {dc_resistance} Ohms/m)",
            f"AC Wire Size: {ac_wire_size} mm² (Resistance: {ac_resistance} Ohms/m)"
        ]

        # Display the results
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, "\n".join(result_lines))

    except ValueError as e:
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, f"Error: {str(e)}")

# Enable right-click for copying/pasting
def show_context_menu(event):
    context_menu.post(event.x_root, event.y_root)

def copy_text():
    try:
        selected_text = result_text.get("sel.first", "sel.last")
        if selected_text:
            result_text.clipboard_clear()  # Clear current clipboard
            result_text.clipboard_append(selected_text)  # Copy selected text
        else:
            messagebox.showwarning("No Text Selected", "Please select some text to copy.")
    except tk.TclError:
        messagebox.showwarning("No Text Selected", "Please select some text to copy.")

def paste_text():
    try:
        result_text.insert(tk.INSERT, root.clipboard_get())  # Paste clipboard content
    except tk.TclError:
        messagebox.showerror("Error", "Nothing to paste!")

# Function for Load Schedule (runs load.py)
def load_schedule():
    """Run the load.py script."""
    try:
        subprocess.run(["python", "load.py"], check=True)  # Adjust path if necessary
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Failed to load schedule: {e}")

# Function for calculation display (as an example)
def display_data():
    """Placeholder for actual solar data calculation."""
    messagebox.showinfo("Display Data", "Displaying solar data...")

# Function to reset inputs
def reset_inputs():
    """Reset the input fields."""
    country_combo.set("Philippines")
    voltage_combo.set(12)
    efficiency_combo.set(17)
    usage_combo.set(8)
    dod_combo.set(50)
    panel_size_combo.set(100)
    wattage_input.delete(0, tk.END)
    wattage_input.insert(0, "1000")
    result_text.delete(1.0, tk.END)

# Function to show the context menu for right-click options
def show_context_menu(event):
    """Show right-click context menu."""
    context_menu.post(event.x_root, event.y_root)

# Function for copying text (for the context menu)
def copy_text():
    result_text.event_generate("<<Copy>>")

# Function for pasting text (for the context menu)
def paste_text():
    result_text.event_generate("<<Paste>>")

# Create the main window
root = tk.Tk()
root.title("Country Solar Data Display")

# Group input widgets in frames for better layout
input_frame = ttk.Frame(root, padding="10")
input_frame.pack(fill=tk.BOTH, expand=True)

result_frame = ttk.Frame(root, padding="10")
result_frame.pack(fill=tk.BOTH, expand=True)

# Country selection
country_label = tk.Label(input_frame, text="Country:")
country_label.grid(row=0, column=0, sticky=tk.W)

country_combo = ttk.Combobox(input_frame, values=["Philippines", "USA", "Canada"], state="readonly")
country_combo.set("Philippines")
country_combo.grid(row=0, column=1, sticky=tk.EW)

# Voltage selection
voltage_label = tk.Label(input_frame, text="Voltage:")
voltage_label.grid(row=1, column=0, sticky=tk.W)

voltage_combo = ttk.Combobox(input_frame, values=[12, 24, 48], state="readonly")
voltage_combo.set(12)
voltage_combo.grid(row=1, column=1, sticky=tk.EW)

# Solar panel efficiency selection
efficiency_label = tk.Label(input_frame, text="Solar Panel Efficiency:")
efficiency_label.grid(row=2, column=0, sticky=tk.W)

efficiency_combo = ttk.Combobox(input_frame, values=[15, 17, 20, 22], state="readonly")
efficiency_combo.set(17)
efficiency_combo.grid(row=2, column=1, sticky=tk.EW)

# Usage hours
usage_label = tk.Label(input_frame, text="Usage Hours:")
usage_label.grid(row=3, column=0, sticky=tk.W)

usage_combo = ttk.Combobox(input_frame, values=[4, 6, 8, 10], state="readonly")
usage_combo.set(8)
usage_combo.grid(row=3, column=1, sticky=tk.EW)

# Depth of discharge
dod_label = tk.Label(input_frame, text="Depth of Discharge (DOD):")
dod_label.grid(row=4, column=0, sticky=tk.W)

dod_combo = ttk.Combobox(input_frame, values=[50, 60, 70, 80, 90, 100], state="readonly")
dod_combo.set(50)
dod_combo.grid(row=4, column=1, sticky=tk.EW)

# Panel size
panel_size_label = tk.Label(input_frame, text="Panel Size (W):")
panel_size_label.grid(row=5, column=0, sticky=tk.W)

panel_size_combo = ttk.Combobox(input_frame, values=[50, 100, 150, 200], state="readonly")
panel_size_combo.set(100)
panel_size_combo.grid(row=5, column=1, sticky=tk.EW)

# Wattage input
wattage_label = tk.Label(input_frame, text="Wattage (W):")
wattage_label.grid(row=6, column=0, sticky=tk.W)

wattage_input = ttk.Entry(input_frame)
wattage_input.insert(0, "1000")
wattage_input.grid(row=6, column=1, sticky=tk.EW)

# Result text area
result_text = tk.Text(result_frame, wrap=tk.WORD, height=10, width=80)
result_text.pack(fill=tk.BOTH, expand=True)

# Create the context menu for right-click options
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Copy", command=copy_text)
context_menu.add_command(label="Paste", command=paste_text)

# Bind right-click event to show context menu
result_text.bind("<Button-3>", show_context_menu)

# Buttons
button_frame = ttk.Frame(root, padding="10")
button_frame.pack(fill=tk.X, expand=False)

display_button = ttk.Button(button_frame, text="Calculate", command=display_data)
display_button.pack(side=tk.LEFT, padx=5)

reset_button = ttk.Button(button_frame, text="Reset", command=reset_inputs)
reset_button.pack(side=tk.LEFT, padx=5)

# Load Schedule Button
load_sched_button = ttk.Button(button_frame, text="Load Sched", command=load_schedule)
load_sched_button.pack(side=tk.LEFT, padx=5)

# Adjust the window size
root.geometry("350x900")

# Run the main application loop
root.mainloop()
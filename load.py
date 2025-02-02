import tkinter as tk
from tkinter import ttk
import pandas as pd
import tkinter.messagebox as messagebox
import json

# Load appliance data from CSV
appliance_data = pd.read_csv('Appliances.csv')

# Set up the Tkinter root window
root = tk.Tk()
root.title("Appliance Power Consumption Calculator")

# Variables to track totals
total_wattage = 0
total_usage_hours = 0
appliance_count = 0

# Function to update the dependent fields (rated power, surge power, etc.)
def update_fields(*args):
    appliance_name = appliance_var.get()  # use the StringVar associated with the combobox

    # Check if the appliance is valid and exists in the list
    if appliance_name in appliance_data['Appliance'].values:
        # Find the row in appliance_data that matches the selected appliance
        appliance_info = appliance_data[appliance_data['Appliance'] == appliance_name].iloc[0]

        # Update the combobox and entry fields with the appliance's values
        rated_power_combobox.set(appliance_info['Rated Power (W)'])
        surge_power_entry.config(state="normal")
        surge_power_entry.delete(0, tk.END)
        surge_power_entry.insert(0, appliance_info['Surge Power (W)'])
        surge_power_entry.config(state="disabled")

        efficiency_entry.config(state="normal")
        efficiency_entry.delete(0, tk.END)
        efficiency_entry.insert(0, f"{appliance_info['Efficiency (%)']}%")
        efficiency_entry.config(state="disabled")

        power_factor_entry.config(state="normal")
        power_factor_entry.delete(0, tk.END)
        power_factor_entry.insert(0, appliance_info['Power Factor (PF)'])
        power_factor_entry.config(state="disabled")
    else:
        # Reset fields if appliance is invalid
        surge_power_entry.config(state="normal")
        surge_power_entry.delete(0, tk.END)
        surge_power_entry.config(state="disabled")

        efficiency_entry.config(state="normal")
        efficiency_entry.delete(0, tk.END)
        efficiency_entry.config(state="disabled")

        power_factor_entry.config(state="normal")
        power_factor_entry.delete(0, tk.END)
        power_factor_entry.config(state="disabled")

# Function to compute power consumption based on user inputs
def add_appliance():
    global total_wattage, total_usage_hours, appliance_count

    appliance = appliance_var.get()  # use the variable from the combobox
    try:
        rated_power = float(rated_power_combobox.get())
    except ValueError:
        messagebox.showwarning("Input Error", "Please select a valid appliance.")
        return

    # Check if surge power entry is empty before converting to float
    surge_power_str = surge_power_entry.get()
    if surge_power_str == '':
        messagebox.showwarning("Input Error", "Surge Power cannot be empty.")
        return

    try:
        surge_power = float(surge_power_str)
    except ValueError:
        messagebox.showwarning("Input Error", "Please enter a valid number for Surge Power.")
        return

    # Remove the '%' from efficiency and convert to float
    efficiency_str = efficiency_entry.get()
    try:
        efficiency = float(efficiency_str.replace('%', '').strip()) / 100
    except ValueError:
        efficiency = 0

    try:
        power_factor = float(power_factor_entry.get())
    except ValueError:
        power_factor = 0

    # Convert usage hours to float (for fractional hours)
    try:
        usage_hours = float(usage_hours_combobox.get())
    except ValueError:
        usage_hours = 6  # Default to 6 hours if invalid input is provided

    # Handle appliance count as a float and round to the nearest integer
    try:
        appliance_count_input = float(counts_combobox.get())
        appliance_count_input = round(appliance_count_input)  # Round to nearest integer
    except ValueError:
        appliance_count_input = 1  # Default to 1 if input is invalid

    # Calculating total power consumption (in kWh)
    consumption = rated_power * usage_hours * appliance_count_input * power_factor * efficiency / 1000

    # Update totals
    total_wattage += rated_power * appliance_count_input  # Sum up rated wattage
    total_usage_hours += usage_hours * appliance_count_input  # Sum up usage hours
    appliance_count += appliance_count_input  # Sum up appliance count

    # Format values with commas for easy reading
    rated_power_formatted = f"{rated_power:,}"
    surge_power_formatted = f"{surge_power:,}"
    consumption_formatted = f"{consumption:,.4f}"  # Four decimal places for consumption

    # Insert result into the Treeview with formatted values
    tree.insert("", "end", values=(
        appliance,
        rated_power_formatted,
        f"{power_factor:.2f}",
        f"{efficiency * 100:.0f}",
        surge_power_formatted,
        usage_hours,
        appliance_count_input,
        consumption_formatted
    ))

    # Update total wattage and average usage time labels
    total_wattage_label.config(text=f"Total Wattage: {total_wattage:,.0f} W")
    average_usage_time = total_usage_hours / appliance_count if appliance_count else 0
    average_usage_time_label.config(text=f"Average Usage Time: {average_usage_time:.2f} hours")

# Auto-complete functionality for combobox
def on_combobox_keyrelease(event):
    # Get the value typed by the user
    typed_value = appliance_var.get()

    # Filter the appliance list to match the typed value
    if typed_value:
        filtered_appliances = [item for item in appliance_data['Appliance']
                               if isinstance(item, str) and item.lower().startswith(typed_value.lower())]
    else:
        filtered_appliances = appliance_data['Appliance'].tolist()

    # Update the combobox list with the filtered items
    appliance_combobox['values'] = filtered_appliances

# Function to handle row selection from Treeview and fill the fields
def on_tree_select(event):
    selected_items = tree.selection()
    if selected_items:
        # If multiple items are selected, update with the first selected item.
        item_values = tree.item(selected_items[0])['values']
        appliance_name = item_values[0]
        rated_power = item_values[1]
        power_factor = item_values[2]
        efficiency = item_values[3]
        surge_power = item_values[4]
        usage_hours = item_values[5]
        appliance_count_val = item_values[6]

        appliance_var.set(appliance_name)
        rated_power_combobox.set(rated_power.replace(",", ""))  # Remove commas for numeric input

        surge_power_entry.config(state="normal")
        surge_power_entry.delete(0, tk.END)
        surge_power_entry.insert(0, surge_power)
        surge_power_entry.config(state="disabled")

        efficiency_entry.config(state="normal")
        efficiency_entry.delete(0, tk.END)
        efficiency_entry.insert(0, efficiency)
        efficiency_entry.config(state="disabled")

        power_factor_entry.config(state="normal")
        power_factor_entry.delete(0, tk.END)
        power_factor_entry.insert(0, power_factor)
        power_factor_entry.config(state="disabled")

        usage_hours_combobox.set(usage_hours)
        counts_combobox.set(appliance_count_val)

# Function to remove selected appliance(s) from Treeview
def remove_appliance():
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("No Selection", "Please select one or more items to remove.")
        return

    for item in selected_items:
        tree.delete(item)


def save_to_csv():
    # Disable the Save button to prevent multiple triggers
    save_button.config(state="disabled")

    # Check if there are any rows in the Treeview
    if not tree.get_children():
        messagebox.showwarning("No Data", "There are no appliances to save.")
        save_button.config(state="normal")
        return

    # Get all rows from the Treeview
    rows = [tree.item(row)['values'] for row in tree.get_children()]

    # Convert to DataFrame and save as CSV
    df = pd.DataFrame(
        rows,
        columns=[
            "Appliance", "Rated Power (W)", "Power Factor (PF)", "Efficiency (%)",
            "Surge Power (W)", "Usage Hours", "Appliance Count", "Consumption (kWh)"
        ]
    )
    try:
        df.to_csv("load_Sched.csv", index=False)

        # Calculate average usage hours
        average_usage_time = total_usage_hours / appliance_count if appliance_count else 0

        # Prepare the totals data dictionary for JSON
        data = {
            "total_wattage": total_wattage,
            "average_usage_hours": average_usage_time
        }
        # Save the totals to total_load.json
        with open("total_load.json", "w") as f:
            json.dump(data, f)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving the data: {e}")
        save_button.config(state="normal")
        return

    # Get all rows from the Treeview
    rows = []
    for row in tree.get_children():
        rows.append(tree.item(row)['values'])

    # Convert to DataFrame and save as CSV
    df = pd.DataFrame(rows, columns=["Appliance", "Rated Power (W)", "Power Factor (PF)", "Efficiency (%)",
                                     "Surge Power (W)", "Usage Hours", "Appliance Count", "Consumption (kWh)"])
    try:
        df.to_csv("load_Sched.csv", index=False)
        messagebox.showinfo("Success", "Data has been saved successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving the data: {e}")

# ---------------------
# Layout Setup
# ---------------------

# Create a frame for organizing the input widgets in rows
frame1 = ttk.Frame(root, padding="5")
frame1.grid(row=0, column=0, padx=5, pady=5)

# Create a StringVar for the appliance combobox
appliance_var = tk.StringVar()

# Group 1 - First row: Appliance selection
appliance_label = ttk.Label(frame1, text="Appliance:")
appliance_label.grid(row=0, column=0, padx=5, pady=5)

appliance_combobox = ttk.Combobox(frame1, textvariable=appliance_var,
                                  values=appliance_data['Appliance'].tolist(), width=43)
appliance_var.set(appliance_data.iloc[0]['Appliance'])  # Default to the first appliance
appliance_combobox.grid(row=0, column=1, padx=5, pady=5)
appliance_var.trace_add("write", update_fields)

rated_power_label = ttk.Label(frame1, text="Rated Power (W):")
rated_power_label.grid(row=0, column=2, padx=5, pady=5)

rated_power_combobox = ttk.Combobox(frame1,
                                    values=appliance_data['Rated Power (W)'].tolist(), width=8)
rated_power_combobox.set(appliance_data.iloc[0]['Rated Power (W)'])
rated_power_combobox.grid(row=0, column=3, padx=5, pady=5)

# Group 2 - Second row: Efficiency and Surge Power
efficiency_label = ttk.Label(frame1, text="Efficiency (%):")
efficiency_label.grid(row=1, column=0, padx=5, pady=5)

efficiency_entry = ttk.Entry(frame1, width=10)
efficiency_entry.grid(row=1, column=1, padx=5, pady=5)
efficiency_entry.config(state="disabled")

surge_power_label = ttk.Label(frame1, text="Surge Power (W):")
surge_power_label.grid(row=1, column=2, padx=5, pady=5)

surge_power_entry = ttk.Entry(frame1, width=10)
surge_power_entry.grid(row=1, column=3, padx=5, pady=5)
surge_power_entry.config(state="disabled")

# Group 3 - Third row: Usage Hours and Appliance Count
usage_hours_label = ttk.Label(frame1, text="Usage Hours:")
usage_hours_label.grid(row=2, column=0, padx=5, pady=5)

usage_hours_combobox = ttk.Combobox(frame1, values=[str(i) for i in range(1, 25)], width=10)
usage_hours_combobox.set("6")
usage_hours_combobox.grid(row=2, column=1, padx=5, pady=5)

counts_label = ttk.Label(frame1, text="Appliance Count:")
counts_label.grid(row=2, column=2, padx=5, pady=5)

counts_combobox = ttk.Combobox(frame1, values=[str(i) for i in range(1, 11)], width=10)
counts_combobox.set("1")
counts_combobox.grid(row=2, column=3, padx=5, pady=5)

# Group 4 - Fourth row: Power Factor
power_factor_label = ttk.Label(frame1, text="Power Factor (PF):")
power_factor_label.grid(row=3, column=0, padx=5, pady=5)

power_factor_entry = ttk.Entry(frame1, width=10)
power_factor_entry.grid(row=3, column=1, padx=5, pady=5)
power_factor_entry.config(state="disabled")

# Bind the key release event for auto-complete on the appliance combobox
appliance_combobox.bind("<KeyRelease>", on_combobox_keyrelease)

# Add Appliance Button (spanning full width in its row)
add_button = ttk.Button(frame1, text="Add", command=add_appliance)
add_button.grid(row=4, column=0, columnspan=4, pady=5)

# Treeview for displaying the appliance list with multiple selection enabled
tree = ttk.Treeview(root, columns=("Appliance", "Power (W)", "PF", "Eff(%)", "Surge(W)",
                                   "Usage (Hrs)", "Count", "Consumption (kWh)"),
                    show="headings", selectmode="extended")
tree.grid(row=1, column=0, padx=5, pady=5)

# Configure columns in the Treeview
for col in tree["columns"]:
    tree.heading(col, text=col)
    tree.column(col, width=80, anchor="center")

# Bind the selection event to update fields when a row is selected
tree.bind("<<TreeviewSelect>>", on_tree_select)

# Create a frame for Save, Remove, Total Wattage and Average Usage Time (all on one row)
action_frame = ttk.Frame(root, padding="5")
action_frame.grid(row=2, column=0, pady=5, sticky="ew")

# Save and Remove buttons
save_button = ttk.Button(action_frame, text="Save", command=save_to_csv)
save_button.grid(row=0, column=0, padx=5)

remove_button = ttk.Button(action_frame, text="Remove", command=remove_appliance)
remove_button.grid(row=0, column=1, padx=5)

# Total Wattage and Average Usage Time Labels
total_wattage_label = ttk.Label(action_frame, text="Total Wattage: 0 W")
total_wattage_label.grid(row=0, column=2, padx=20)

average_usage_time_label = ttk.Label(action_frame, text="Ave. Time: 0 hours")
average_usage_time_label.grid(row=0, column=3, padx=20)

# Start the Tkinter event loop
root.mainloop()

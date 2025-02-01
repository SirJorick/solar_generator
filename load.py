import tkinter as tk
from tkinter import ttk
import pandas as pd
from tkinter import simpledialog
import tkinter.messagebox as messagebox

# Load appliance data from CSV
appliance_data = pd.read_csv('Appliances.csv')

# Set up the Tkinter root window
root = tk.Tk()
root.title("Appliance Power Consumption Calculator")

# Function to update the dependent fields (rated power, surge power, etc.)
def update_fields(event=None):
    appliance_name = appliance_combobox.get()

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
    appliance = appliance_combobox.get()
    rated_power = float(rated_power_combobox.get())

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
    efficiency = float(efficiency_str.replace('%', '').strip()) / 100

    power_factor = float(power_factor_entry.get())

    # Convert usage hours to float (for fractional hours)
    try:
        usage_hours = float(usage_hours_combobox.get())
    except ValueError:
        usage_hours = 6  # Default to 6 hours if invalid input is provided

    # Handle appliance count as a float and round to the nearest integer
    try:
        appliance_count = float(counts_combobox.get())
        appliance_count = round(appliance_count)  # Round to nearest integer
    except ValueError:
        appliance_count = 1  # Default to 1 if input is invalid

    # Calculating total power consumption
    consumption = rated_power * usage_hours * appliance_count * power_factor * efficiency / 1000  # in kWh

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
        appliance_count,
        consumption_formatted
    ))

# Auto-complete functionality for combobox
def on_combobox_keyrelease(event):
    # Get the value typed by the user
    typed_value = appliance_combobox.get()

    # Filter the appliance list to match the typed value
    if typed_value:
        filtered_appliances = [item for item in appliance_data['Appliance'] if isinstance(item, str) and item.lower().startswith(typed_value.lower())]
    else:
        filtered_appliances = appliance_data['Appliance'].tolist()

    # Update the combobox list with the filtered items
    appliance_combobox['values'] = filtered_appliances
    if filtered_appliances:
        appliance_combobox.set(filtered_appliances[0])

# Function to handle row selection from Treeview and fill the fields
def on_tree_select(event):
    selected_item = tree.selection()
    if selected_item:
        item_values = tree.item(selected_item)['values']
        appliance_name = item_values[0]
        rated_power = item_values[1]
        power_factor = item_values[2]
        efficiency = item_values[3]
        surge_power = item_values[4]
        usage_hours = item_values[5]
        appliance_count = item_values[6]

        appliance_combobox.set(appliance_name)
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
        counts_combobox.set(appliance_count)

# Function to remove selected appliance from Treeview
def remove_appliance():
    selected_item = tree.selection()
    if selected_item:
        tree.delete(selected_item)

# Function to save the tree data to a CSV file with validation and popup confirmation
def save_to_csv():
    # Check if there are any rows in the Treeview
    if not tree.get_children():
        messagebox.showwarning("No Data", "There are no appliances to save.")
        return

    # Get all rows from the Treeview
    rows = []
    for row in tree.get_children():
        rows.append(tree.item(row)['values'])

    # Convert to DataFrame and save as CSV
    df = pd.DataFrame(rows, columns=["Appliance", "Rated Power (W)", "Power Factor (PF)", "Efficiency (%)", "Surge Power (W)", "Usage Hours", "Appliance Count", "Consumption (kWh)"])
    try:
        df.to_csv("load_Sched.csv", index=False)
        messagebox.showinfo("Success", "Data has been saved successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving the data: {e}")

# Create a frame for organizing the widgets
frame1 = ttk.Frame(root, padding="5")
frame1.grid(row=0, column=0, padx=5, pady=5)

# Appliance selection combobox
appliance_label = ttk.Label(frame1, text="Appliance:")
appliance_label.grid(row=0, column=0, padx=5, pady=5)

appliance_combobox = ttk.Combobox(frame1, values=appliance_data['Appliance'].tolist(), width=43)
appliance_combobox.set(appliance_data.iloc[0]['Appliance'])  # Default to the first appliance
appliance_combobox.grid(row=0, column=1, padx=5, pady=5)
appliance_combobox.bind("<<ComboboxSelected>>", update_fields)  # Trigger update_fields on combobox selection change
appliance_combobox.bind("<KeyRelease>", on_combobox_keyrelease)  # Trigger on key release for auto-complete

# Rated Power combobox
rated_power_label = ttk.Label(frame1, text="Rated Power (W):")
rated_power_label.grid(row=0, column=2, padx=5, pady=5)

rated_power_combobox = ttk.Combobox(frame1, values=appliance_data['Rated Power (W)'].tolist(), width=8)
rated_power_combobox.set(appliance_data.iloc[0]['Rated Power (W)'])  # Default to the first appliance's rated power
rated_power_combobox.grid(row=0, column=3, padx=5, pady=5)

# Efficiency text box (disabled)
efficiency_label = ttk.Label(frame1, text="Efficiency (%):")
efficiency_label.grid(row=0, column=4, padx=5, pady=5)

efficiency_entry = ttk.Entry(frame1, width=10)
efficiency_entry.grid(row=0, column=5, padx=5, pady=5)
efficiency_entry.config(state="disabled")  # Make it disabled

# Power Factor text box (disabled)
power_factor_label = ttk.Label(frame1, text="Power Factor (PF):")
power_factor_label.grid(row=0, column=6, padx=5, pady=5)

power_factor_entry = ttk.Entry(frame1, width=10)
power_factor_entry.grid(row=0, column=7, padx=5, pady=5)
power_factor_entry.config(state="disabled")  # Make it disabled

# Surge Power text box (disabled)
surge_power_label = ttk.Label(frame1, text="Surge Power (W):")
surge_power_label.grid(row=0, column=8, padx=5, pady=5)

surge_power_entry = ttk.Entry(frame1, width=10)
surge_power_entry.grid(row=0, column=9, padx=5, pady=5)
surge_power_entry.config(state="disabled")  # Make it disabled

# Usage Hours combobox
usage_hours_label = ttk.Label(frame1, text="Usage Hours:")
usage_hours_label.grid(row=1, column=0, padx=5, pady=5)

usage_hours_combobox = ttk.Combobox(frame1, values=[1, 2, 3, 4, 5, 6], width=5)
usage_hours_combobox.set(1)  # Default to 6 hours
usage_hours_combobox.grid(row=1, column=1, padx=5, pady=5)

# Appliance Count combobox
counts_label = ttk.Label(frame1, text="Appliance Count:")
counts_label.grid(row=1, column=2, padx=5, pady=5)

counts_combobox = ttk.Combobox(frame1, values=[1, 2, 3, 4, 5], width=5)
counts_combobox.set(1)  # Default to 1 appliance
counts_combobox.grid(row=1, column=3, padx=5, pady=5)

# Add Appliance button
add_button = ttk.Button(frame1, text="Add Appliance", command=add_appliance)
add_button.grid(row=1, column=4, padx=5, pady=5)

# Create the treeview to display appliance data
# Create the treeview to display appliance data
tree = ttk.Treeview(root, columns=("Appliance", "Rated Power (W)", "Power Factor (PF)", "Efficiency (%)", "Surge Power (W)", "Usage Hours", "Appliance Count", "Consumption (kWh)"), show="headings", height=20)  # Set height to 20 or any value you prefer
tree.grid(row=1, column=0, columnspan=8, padx=5, pady=5, sticky="nsew")  # Allow it to expand in all directions

# Configure the grid layout to allow the treeview to expand when the window is resized
root.grid_rowconfigure(1, weight=1)  # Make the second row (where the treeview is) expandable
root.grid_columnconfigure(0, weight=1)  # Make the first column expandable (if you want the entire treeview to resize horizontally)
root.grid_columnconfigure(1, weight=1)  # Make the second column expandable (if you want the entire treeview to resize horizontally)


# Define columns in Treeview
for col in tree["columns"]:
    tree.heading(col, text=col)

# Bind selection event to Treeview
tree.bind("<<TreeviewSelect>>", on_tree_select)

# Add Remove Appliance and Save buttons
remove_button = ttk.Button(root, text="Remove Appliance", command=remove_appliance)
remove_button.grid(row=2, column=0, padx=5, pady=5)

save_button = ttk.Button(root, text="Save to CSV", command=save_to_csv)
save_button.grid(row=2, column=1, padx=5, pady=5)

# Start the main loop for the application
root.mainloop()

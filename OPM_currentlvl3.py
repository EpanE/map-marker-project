import folium
import pandas as pd
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os

# Initialize the DataFrame to hold marker data
df = pd.DataFrame(columns=['Latitude', 'Longitude', 'Von', 'Voff', 'Description', 'Color', 'Icon'])

# Function to add or update a marker
def add_marker():
    description = desc_entry.get()
    von = von_entry.get()
    voff = voff_entry.get()
    latitude = lat_entry.get()
    longitude = lon_entry.get()

    # Validate input
    if not description or not von or not voff or not latitude or not longitude:
        messagebox.showerror("Input Error", "All fields must be filled out")
        return

    try:
        latitude = float(latitude)
        longitude = float(longitude)
        von = float(von)
        voff = float(voff)
    except ValueError:
        messagebox.showerror("Input Error", "Latitude, Longitude, Von, and Voff must be numbers")
        return

    # Determine color based on Voff value
    color = 'green' if -1.5 <= voff <= -0.8 else 'red'
    icon = 'info-sign'

    selected_index = coordinates_listbox.curselection()
    if selected_index:
        # Update existing marker
        index = selected_index[0]
        df.iloc[index] = [latitude, longitude, von, voff, description, color, icon]
        coordinates_listbox.delete(index)
        coordinates_listbox.insert(index, f"{description} ({latitude}, {longitude}) - Von: {von}, Voff: {voff}")
    else:
        # Add new marker
        df.loc[len(df)] = [latitude, longitude, von, voff, description, color, icon]
        coordinates_listbox.insert(tk.END, f"{description} ({latitude}, {longitude}) - Von: {von}, Voff: {voff}")

    # Clear the input fields
    clear_input_fields()

def clear_input_fields():
    desc_entry.delete(0, tk.END)
    von_entry.delete(0, tk.END)
    voff_entry.delete(0, tk.END)
    lat_entry.delete(0, tk.END)
    lon_entry.delete(0, tk.END)

# Function to create the map and open it in the browser
def create_map():
    if len(df) == 0:
        messagebox.showerror("Error", "No markers to add to the map")
        return

    initial_location = [df['Latitude'].iloc[0], df['Longitude'].iloc[0]]
    mymap = folium.Map(location=initial_location, zoom_start=12)

    for index, row in df.iterrows():
        folium.Marker(
            location=[row['Latitude'], row['Longitude']],
            popup=row['Description'],
            icon=folium.Icon(color=row['Color'], icon=row['Icon'])
        ).add_to(mymap)

    output_map_file = 'marked_map.html'
    try:
        mymap.save(output_map_file)
        webbrowser.open('file://' + os.path.realpath(output_map_file))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save or open the map: {str(e)}")

# Function to load data from an Excel file with specific headers
def load_from_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            # Read the Excel file
            excel_df = pd.read_excel(file_path)

            # Convert all column headers to lowercase for easier comparison
            excel_df.columns = excel_df.columns.str.lower()

            # Check if required columns exist in the file
            required_columns = ['latitude', 'longitude', 'von', 'voff', 'description']
            if not all(column in excel_df.columns for column in required_columns):
                raise ValueError("The Excel file must contain the following headers: latitude, longitude, von, voff, description")

            # Clear the current DataFrame and Listbox before loading new data
            global df
            df = pd.DataFrame(columns=['Latitude', 'Longitude', 'Von', 'Voff', 'Description', 'Color', 'Icon'])
            coordinates_listbox.delete(0, tk.END)

            # Process each row and insert it into the DataFrame and Listbox
            for index, row in excel_df.iterrows():
                latitude = row['latitude']
                longitude = row['longitude']
                von = row['von']
                voff = row['voff']
                description = row['description']

                # Determine color based on the value of voff
                color = 'green' if -1.5 <= voff <= -0.8 else 'red'
                icon = 'info-sign'

                # Add to DataFrame
                df.loc[len(df)] = [latitude, longitude, von, voff, description, color, icon]

                # Insert into Listbox for display
                coordinates_listbox.insert(tk.END, f"{description} ({latitude}, {longitude}) - Von: {von}, Voff: {voff}")

            messagebox.showinfo("Success", "Data loaded successfully from Excel file")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data from Excel: {str(e)}")

# Function to remove selected marker
def remove_marker():
    selected_index = coordinates_listbox.curselection()
    if selected_index:
        index = selected_index[0]
        df.drop(index, inplace=True)
        df.reset_index(drop=True, inplace=True)
        coordinates_listbox.delete(index)
    else:
        messagebox.showerror("Error", "No marker selected for removal")

# Function to edit selected marker
def edit_marker():
    selected_index = coordinates_listbox.curselection()
    if selected_index:
        index = selected_index[0]
        row = df.iloc[index]
        desc_entry.delete(0, tk.END)
        desc_entry.insert(0, row['Description'])
        von_entry.delete(0, tk.END)
        von_entry.insert(0, str(row['Von']))
        voff_entry.delete(0, tk.END)
        voff_entry.insert(0, str(row['Voff']))
        lat_entry.delete(0, tk.END)
        lat_entry.insert(0, str(row['Latitude']))
        lon_entry.delete(0, tk.END)
        lon_entry.insert(0, str(row['Longitude']))
    else:
        messagebox.showerror("Error", "No marker selected for editing")

# Initialize the main application window
root = tk.Tk()
root.title("Map Marker Tool")
root.geometry("700x500")

# Create a style for the UI
style = ttk.Style(root)
style.theme_use('clam')

# Main frame
main_frame = ttk.Frame(root, padding="10 10 10 10")
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Make the grid cells resizable
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
main_frame.grid_rowconfigure(7, weight=1)
main_frame.grid_columnconfigure(1, weight=1)

# Labels and Entries for marker details
ttk.Label(main_frame, text="Description:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
desc_entry = ttk.Entry(main_frame)
desc_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Von:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
von_entry = ttk.Entry(main_frame)
von_entry.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Voff:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
voff_entry = ttk.Entry(main_frame)
voff_entry.grid(row=2, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Latitude:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
lat_entry = ttk.Entry(main_frame)
lat_entry.grid(row=3, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Longitude:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
lon_entry = ttk.Entry(main_frame)
lon_entry.grid(row=4, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# Add marker button
add_marker_btn = ttk.Button(main_frame, text="Add/Update Marker", command=add_marker)
add_marker_btn.grid(row=5, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Listbox to display the coordinates and associated information
coordinates_listbox = tk.Listbox(main_frame)
coordinates_listbox.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

# Button to create the map
create_map_btn = ttk.Button(main_frame, text="Create Map", command=create_map)
create_map_btn.grid(row=7, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Button to load data from Excel
load_excel_btn = ttk.Button(main_frame, text="Load from Excel", command=load_from_excel)
load_excel_btn.grid(row=8, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Button to remove selected marker
remove_marker_btn = ttk.Button(main_frame, text="Remove Marker", command=remove_marker)
remove_marker_btn.grid(row=9, column=0, pady=10, sticky=(tk.W, tk.E))

# Button to edit selected marker
edit_marker_btn = ttk.Button(main_frame, text="Edit Marker", command=edit_marker)
edit_marker_btn.grid(row=9, column=1, pady=10, sticky=(tk.W, tk.E))

# Start the Tkinter event loop
root.mainloop()
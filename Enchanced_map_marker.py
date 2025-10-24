import folium
import pandas as pd
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Initialize the DataFrame to hold marker data
df = pd.DataFrame(columns=['Latitude', 'Longitude', 'Label', 'Color'])


# Function to add a marker to the map
def add_marker():
    # Get user input
    name = name_entry.get()
    color = color_entry.get()
    latitude = lat_entry.get()
    longitude = lon_entry.get()

    # Validate input
    if not name or not color or not latitude or not longitude:
        messagebox.showerror("Input Error", "All fields must be filled out")
        return

    try:
        latitude = float(latitude)
        longitude = float(longitude)
    except ValueError:
        messagebox.showerror("Input Error", "Latitude and Longitude must be numbers")
        return

    # Add the data to the dataframe
    df.loc[len(df)] = [latitude, longitude, name, color]

    # Update the listbox with the new entry
    coordinates_listbox.insert(tk.END, f"{name} ({latitude}, {longitude}) - {color}")

    # Clear the input fields
    name_entry.delete(0, tk.END)
    color_entry.delete(0, tk.END)
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
            popup=row['Label'],
            icon=folium.Icon(color=row['Color'], icon='info-sign')
        ).add_to(mymap)

    output_map_file = 'marked_map.html'
    mymap.save(output_map_file)
    webbrowser.open(output_map_file)


# Function to load data from an Excel file
def load_from_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        excel_df = pd.read_excel(file_path)
        for index, row in excel_df.iterrows():
            df.loc[len(df)] = [row['Latitude'], row['Longitude'], row['Label'], row['Color']]
            coordinates_listbox.insert(tk.END,
                                       f"{row['Label']} ({row['Latitude']}, {row['Longitude']}) - {row['Color']}")


# Initialize the main application window
root = tk.Tk()
root.title("Map Marker Tool")
root.geometry("600x400")

# Create a style for the UI
style = ttk.Style(root)
style.theme_use('clam')

# Main frame
main_frame = ttk.Frame(root, padding="10 10 10 10")
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Make the grid cells resizable
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
main_frame.grid_rowconfigure(5, weight=1)
main_frame.grid_columnconfigure(1, weight=1)

# Labels and Entries for marker details
ttk.Label(main_frame, text="Name:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
name_entry = ttk.Entry(main_frame)
name_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Color:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
color_entry = ttk.Entry(main_frame)
color_entry.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Latitude:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
lat_entry = ttk.Entry(main_frame)
lat_entry.grid(row=2, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Longitude:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
lon_entry = ttk.Entry(main_frame)
lon_entry.grid(row=3, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# Add marker button
add_marker_btn = ttk.Button(main_frame, text="Add Marker", command=add_marker)
add_marker_btn.grid(row=4, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Listbox to display the added coordinates
coordinates_listbox = tk.Listbox(main_frame, height=10)
coordinates_listbox.grid(row=5, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))

# Load from Excel button
load_excel_btn = ttk.Button(main_frame, text="Load from Excel", command=load_from_excel)
load_excel_btn.grid(row=6, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Create map button
create_map_btn = ttk.Button(main_frame, text="Create Map", command=create_map)
create_map_btn.grid(row=7, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Run the application
root.mainloop()

import folium
import pandas as pd
import webbrowser
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog


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

    # Clear the input fields
    name_entry.delete(0, tk.END)
    color_entry.delete(0, tk.END)
    lat_entry.delete(0, tk.END)
    lon_entry.delete(0, tk.END)


# Function to create the map and open it in the browser
def create_map():
    # Create a map centered around the first coordinate
    if len(df) == 0:
        messagebox.showerror("Error", "No markers to add to the map")
        return

    initial_location = [df['Latitude'].iloc[0], df['Longitude'].iloc[0]]
    mymap = folium.Map(location=initial_location, zoom_start=12)

    # Loop through the dataframe and add markers
    for index, row in df.iterrows():
        folium.Marker(
            location=[row['Latitude'], row['Longitude']],
            popup=row['Label'],
            icon=folium.Icon(color=row['Color'], icon='info-sign')
        ).add_to(mymap)

    # Save the map to an HTML file
    output_map_file = 'marked_map.html'
    mymap.save(output_map_file)
    print(f"Map saved to {output_map_file}")

    # Automatically open the map in the default web browser
    webbrowser.open(output_map_file)


# Initialize the main application window
root = tk.Tk()
root.title("Map Marker Tool")

# Initialize the dataframe to hold marker data
df = pd.DataFrame(columns=['Latitude', 'Longitude', 'Label', 'Color'])

# UI Labels and Entries for marker details
tk.Label(root, text="Name:").grid(row=0, column=0, padx=5, pady=5)
name_entry = tk.Entry(root)
name_entry.grid(row=0, column=1, padx=5, pady=5)

tk.Label(root, text="Color:").grid(row=1, column=0, padx=5, pady=5)
color_entry = tk.Entry(root)
color_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Latitude:").grid(row=2, column=0, padx=5, pady=5)
lat_entry = tk.Entry(root)
lat_entry.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Longitude:").grid(row=3, column=0, padx=5, pady=5)
lon_entry = tk.Entry(root)
lon_entry.grid(row=3, column=1, padx=5, pady=5)

# Buttons to add a marker and create the map
add_marker_btn = tk.Button(root, text="Add Marker", command=add_marker)
add_marker_btn.grid(row=4, column=0, columnspan=2, pady=10)

create_map_btn = tk.Button(root, text="Create Map", command=create_map)
create_map_btn.grid(row=5, column=0, columnspan=2, pady=10)

# Run the application
root.mainloop()

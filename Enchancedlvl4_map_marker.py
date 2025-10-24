import folium
import pandas as pd
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser
from PIL import ImageTk, Image
import io
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager


output_map_file = 'marked_map.html'

# Initialize the DataFrame to hold marker data
df = pd.DataFrame(columns=['Latitude', 'Longitude', 'Label', 'Color', 'Icon'])


# Function to add or update a marker
def add_marker():
    # Get user input
    name = name_entry.get()
    color = color_entry.get()
    latitude = lat_entry.get()
    longitude = lon_entry.get()
    icon = icon_var.get()

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

    selected_index = coordinates_listbox.curselection()
    if selected_index:
        # Update existing marker
        index = selected_index[0]
        df.iloc[index] = [latitude, longitude, name, color, icon]
        coordinates_listbox.delete(index)
        coordinates_listbox.insert(index, f"{name} ({latitude}, {longitude}) - {color} - {icon}")
    else:
        # Add new marker
        df.loc[len(df)] = [latitude, longitude, name, color, icon]
        coordinates_listbox.insert(tk.END, f"{name} ({latitude}, {longitude}) - {color} - {icon}")

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
    mymap = folium.Map(location=initial_location, zoom_start=zoom_scale.get(), tiles=map_style.get())

    for index, row in df.iterrows():
        folium.Marker(
            location=[row['Latitude'], row['Longitude']],
            popup=row['Label'],
            icon=folium.Icon(color=row['Color'], icon=row['Icon'])
        ).add_to(mymap)

    output_map_file = 'marked_map.html'
    mymap.save(output_map_file)
    webbrowser.open(output_map_file)


# Function to delete a marker
def delete_marker():
    selected_index = coordinates_listbox.curselection()
    if selected_index:
        index = selected_index[0]
        coordinates_listbox.delete(index)
        df.drop(df.index[index], inplace=True)
        df.reset_index(drop=True, inplace=True)
    else:
        messagebox.showerror("Delete Error", "No marker selected to delete")


# Function to load data from an Excel file
def load_from_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        excel_df = pd.read_excel(file_path)
        for index, row in excel_df.iterrows():
            df.loc[len(df)] = [row['Latitude'], row['Longitude'], row['Label'], row['Color'],
                               row.get('Icon', 'info-sign')]
            coordinates_listbox.insert(tk.END,
                                       f"{row['Label']} ({row['Latitude']}, {row['Longitude']}) - {row['Color']} - {row.get('Icon', 'info-sign')}")


# Function to save project to a CSV file
def save_project():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if file_path:
        df.to_csv(file_path, index=False)
        messagebox.showinfo("Save Successful", "Project saved successfully")


# Function to load project from a CSV file
def load_project():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        global df
        df = pd.read_csv(file_path)
        coordinates_listbox.delete(0, tk.END)
        for index, row in df.iterrows():
            coordinates_listbox.insert(tk.END,
                                       f"{row['Label']} ({row['Latitude']}, {row['Longitude']}) - {row['Color']} - {row['Icon']}")


# Function to pick a color
def pick_color():
    color_code = colorchooser.askcolor(title="Choose color")[1]
    if color_code:
        color_entry.delete(0, tk.END)
        color_entry.insert(tk.END, color_code)


# Function to edit an existing marker
def edit_marker(event):
    selected_index = coordinates_listbox.curselection()
    if selected_index:
        index = selected_index[0]
        selected_marker = df.iloc[index]
        name_entry.delete(0, tk.END)
        name_entry.insert(tk.END, selected_marker['Label'])
        color_entry.delete(0, tk.END)
        color_entry.insert(tk.END, selected_marker['Color'])
        lat_entry.delete(0, tk.END)
        lat_entry.insert(tk.END, selected_marker['Latitude'])
        lon_entry.delete(0, tk.END)
        lon_entry.insert(tk.END, selected_marker['Longitude'])
        icon_var.set(selected_marker['Icon'])


# Function to search for markers
def search_markers():
    query = search_entry.get().lower()
    coordinates_listbox.delete(0, tk.END)
    for index, row in df.iterrows():
        if query in row['Label'].lower() or query in row['Color'].lower() or query in str(
                row['Latitude']) or query in str(row['Longitude']):
            coordinates_listbox.insert(tk.END,
                                       f"{row['Label']} ({row['Latitude']}, {row['Longitude']}) - {row['Color']} - {row['Icon']}")


# Function to export the map as an image
def export_map_as_image():
    create_map()  # Ensure the map is created and saved as HTML

    # Setup Chrome WebDriver
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    try:
        driver.get(f"file://{output_map_file}")
        image = driver.get_screenshot_as_png()

        # Save the screenshot as an image file
        with open('marked_map.png', 'wb') as file:
            file.write(image)
        messagebox.showinfo("Export Successful", "Map exported as image successfully")
    finally:
        driver.quit()


# Initialize the main application window
root = tk.Tk()
root.title("Map Marker Tool")
root.geometry("700x600")

# Create a style for the UI
style = ttk.Style(root)
style.theme_use('clam')

# Main frame
main_frame = ttk.Frame(root, padding="10 10 10 10")
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Make the grid cells resizable
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)
main_frame.grid_rowconfigure(9, weight=1)
main_frame.grid_columnconfigure(1, weight=1)

# Labels and Entries for marker details
ttk.Label(main_frame, text="Name:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
name_entry = ttk.Entry(main_frame)
name_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Color:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
color_entry = ttk.Entry(main_frame)
color_entry.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
color_picker_btn = ttk.Button(main_frame, text="Pick Color", command=pick_color)
color_picker_btn.grid(row=1, column=2, padx=5, pady=5, sticky=tk.W)

ttk.Label(main_frame, text="Latitude:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
lat_entry = ttk.Entry(main_frame)
lat_entry.grid(row=2, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Longitude:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
lon_entry = ttk.Entry(main_frame)
lon_entry.grid(row=3, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

ttk.Label(main_frame, text="Icon:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
icon_var = tk.StringVar(value='info-sign')
icon_menu = ttk.Combobox(main_frame, textvariable=icon_var, values=['info-sign', 'cloud', 'home', 'flag', 'star'])
icon_menu.grid(row=4, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# Add Marker Button
add_marker_btn = ttk.Button(main_frame, text="Add/Update Marker", command=add_marker)
add_marker_btn.grid(row=5, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Listbox to display the coordinates and associated information
coordinates_listbox = tk.Listbox(main_frame)
coordinates_listbox.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))
coordinates_listbox.bind('<Double-1>', edit_marker)

# Search functionality
ttk.Label(main_frame, text="Search:").grid(row=7, column=0, padx=5, pady=5, sticky=tk.W)
search_entry = ttk.Entry(main_frame)
search_entry.grid(row=7, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
search_button = ttk.Button(main_frame, text="Search", command=search_markers)
search_button.grid(row=7, column=2, padx=5, pady=5, sticky=tk.W)

# Button to delete a marker
delete_marker_btn = ttk.Button(main_frame, text="Delete Marker", command=delete_marker)
delete_marker_btn.grid(row=8, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Buttons to create the map and save/load projects
create_map_btn = ttk.Button(main_frame, text="Create Map", command=create_map)
create_map_btn.grid(row=9, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

save_project_btn = ttk.Button(main_frame, text="Save Project", command=save_project)
save_project_btn.grid(row=10, column=0, pady=10, sticky=(tk.W, tk.E))

load_project_btn = ttk.Button(main_frame, text="Load Project", command=load_project)
load_project_btn.grid(row=10, column=1, pady=10, sticky=(tk.W, tk.E))

# Button to load data from Excel
load_excel_btn = ttk.Button(main_frame, text="Load from Excel", command=load_from_excel)
load_excel_btn.grid(row=11, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Zoom level control
ttk.Label(main_frame, text="Zoom Level:").grid(row=12, column=0, padx=5, pady=5, sticky=tk.W)
zoom_scale = tk.Scale(main_frame, from_=1, to=18, orient=tk.HORIZONTAL)
zoom_scale.set(10)
zoom_scale.grid(row=12, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# Map style options
ttk.Label(main_frame, text="Map Style:").grid(row=13, column=0, padx=5, pady=5, sticky=tk.W)
map_style = tk.StringVar(value='OpenStreetMap')
map_style_menu = ttk.Combobox(main_frame, textvariable=map_style,
                              values=['OpenStreetMap', 'Stamen Terrain', 'Stamen Toner', 'Stamen Watercolor',
                                      'CartoDB positron', 'CartoDB dark_matter'])
map_style_menu.grid(row=13, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

# Button to export the map as an image
export_image_btn = ttk.Button(main_frame, text="Export as Image", command=export_map_as_image)
export_image_btn.grid(row=14, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

# Start the Tkinter event loop
root.mainloop()

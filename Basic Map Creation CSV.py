import folium
import pandas as pd

# Function to create a map and add markers
def create_map(coordinates_file, output_map_file):
    # Read the coordinates from the CSV file
    df = pd.read_csv(coordinates_file)

    # Create a map centered around the first coordinate
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
    mymap.save(output_map_file)
    print(f"Map saved to {output_map_file}")

# Define the input CSV file and output HTML file
coordinates_file = 'coordinates.csv'  # Replace with your CSV file path
output_map_file = 'marked_map.html'   # The output HTML file for the map

# Call the function to create the map
create_map(coordinates_file, output_map_file)

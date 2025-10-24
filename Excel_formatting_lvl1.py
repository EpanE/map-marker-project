import pandas as pd
import random

# Load the original Excel file
input_file = 'cp station - rdx.xlsx'  # Replace with your input file name
output_file = 'cp station - reformated.xlsx'  # Output file name
df = pd.read_excel(input_file)

# Reorder the columns using the correct case-sensitive names
df_reformatted = df[['latitude', 'longitude', 'Code']].rename(columns={'Code': 'Label'})

# Define possible colors
colors = ['red', 'orange', 'blue', 'green']

# Assign a random color to each row
df_reformatted['color'] = [random.choice(colors) for _ in range(len(df))]

# Default icon assignment
signs = [ 'info-sign' , 'cloud' , 'flag']
df_reformatted['icon'] = [random.choice(signs) for _ in range(len(df))]# Default icon (can be customized)

# Add new columns
df_reformatted['name'] = df['Country']  # Assuming 'name' is the country name

# Save the reformatted data to a new Excel file
df_reformatted.to_excel(output_file, index=False)

print(f'Reformatted data with random colors saved to {output_file}')

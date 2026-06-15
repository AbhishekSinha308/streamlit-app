import pandas as pd
from datetime import datetime

# File paths
weather_file_path = r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\Filtered_WeatherDB_Data.xlsx'
ncfile_file_path = r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\Filtered_NcFile_data.xlsx'
output_file_path = r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\Temperature_Difference_Results.xlsx'

# Read the Excel files
weather_df = pd.read_excel(weather_file_path)
ncfile_df = pd.read_excel(ncfile_file_path)

# Get today's date
today_date = datetime.today().strftime('%Y-%m-%d')
print(f"Using today's date: {today_date}")

# Convert TIMESTAMP_DAY to datetime format
weather_df['TIMESTAMP_DAY'] = pd.to_datetime(weather_df['TIMESTAMP_DAY'], errors='coerce').dt.strftime('%Y-%m-%d')

# Debugging: Check unique dates in weather data
print("Unique dates in weather data:", weather_df['TIMESTAMP_DAY'].unique())

# Filter weather data for today's date
filtered_weather = weather_df[weather_df['TIMESTAMP_DAY'] == today_date].copy()

if filtered_weather.empty:
    print(f"No weather data found for {today_date}.")
else:
    print(f"Found {len(filtered_weather)} rows of weather data for {today_date}.")

# Convert Filtered_TMP column to lists
if 'Filtered_TMP' in filtered_weather.columns:
    filtered_weather['Filtered_TMP_List'] = filtered_weather['Filtered_TMP'].apply(
        lambda x: list(map(float, x.split(','))) if isinstance(x, str) else []
    )
else:
    print("Column 'Filtered_TMP' not found in weather data!")
    exit()

# Round latitude & longitude to avoid floating-point mismatches
filtered_weather[['LATITUDE', 'LONGITUDE']] = filtered_weather[['LATITUDE', 'LONGITUDE']].round(6)
ncfile_df[['LATITUDE', 'LONGITUDE']] = ncfile_df[['LATITUDE', 'LONGITUDE']].round(6)

# Debugging: Check unique latitudes & longitudes
print("Weather lat/lon samples:", filtered_weather[['LATITUDE', 'LONGITUDE']].drop_duplicates().head())
print("NcFile lat/lon samples:", ncfile_df[['LATITUDE', 'LONGITUDE']].drop_duplicates().head())

# Filter ncfile data based on latitude and longitude
filtered_ncfile = ncfile_df[ncfile_df['LATITUDE'].isin(filtered_weather['LATITUDE']) & 
                             ncfile_df['LONGITUDE'].isin(filtered_weather['LONGITUDE'])]

if filtered_ncfile.empty:
    print("No matching latitude and longitude found in NcFile data.")
else:
    print(f"Found {len(filtered_ncfile)} rows of NcFile data.")

# Merge the filtered weather and ncfile data
merged_df = pd.merge(filtered_weather, filtered_ncfile, on=['LATITUDE', 'LONGITUDE'], how='inner')

if merged_df.empty:
    print("No matching rows found after merging weather and ncfile data.")
    exit()
else:
    print(f"Found {len(merged_df)} rows after merging.")

# Calculate temperature differences
result_list = []

for _, row in merged_df.iterrows():
    lat = row['LATITUDE']
    lon = row['LONGITUDE']
    tmp_list = row['Filtered_TMP_List']
    
    timestep = int(row['Time Step']) - 13 

    if 0 <= timestep < len(tmp_list):
        diff = row['Temperature (°C)'] - tmp_list[timestep]
        result_list.append({
            'LATITUDE': lat,
            'LONGITUDE': lon,
            'Time Step': row['Time Step'],
            'Temp_Difference': round(diff, 1)
        })
    else:
        print(f"Warning: Time step {row['Time Step']} is out of range for weather data at ({lat}, {lon}).")
    
# Create a DataFrame for results
result_df = pd.DataFrame(result_list)

# Save results if not empty
if not result_df.empty:
    result_df.to_excel(output_file_path, index=False)
    print(f"Results saved to {output_file_path}.")
else:
    print("No valid results to save.")

# Print the results DataFrame
print(result_df)

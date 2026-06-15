import os
import re
import sys
import netCDF4
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

# Force UTF-8 encoding for output
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# Use Path objects for cross-platform compatibility
directory_path = Path(r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha')

output_excel_path = directory_path / 'NcFile_data.xlsx'
lat_lon_output_path = directory_path / 'LatLonCounts.csv'

file_pattern = r'(00Z|06Z|12Z|18Z)_(\d{8})_RE\.nc'

files_with_datetime = []

for file_name in os.listdir(str(directory_path)):
    match = re.match(file_pattern, file_name)
    if match:
        file_time = match.group(1)     
        file_date = match.group(2)       

        file_datetime = pd.to_datetime(
            file_date + file_time[:-1],
            format='%m%d%Y%H'
        )
        files_with_datetime.append((file_name, file_datetime))

files_with_datetime.sort(key=lambda x: x[1], reverse=True)

current_datetime = datetime.now()
cutoff_datetime = current_datetime - timedelta(days=2.1)

recent_files = [
    (file_name, file_datetime)
    for file_name, file_datetime in files_with_datetime
    if file_datetime >= cutoff_datetime
]

if not recent_files:
    print("No files found for the last 2 days.")
    exit()

print(f"Found {len(recent_files)} files from the last 2 days.")

latest_file, latest_datetime = recent_files[0]
latest_file_path = directory_path / latest_file

print(f"Processing latest file: {latest_file_path}")

dataset = netCDF4.Dataset(str(latest_file_path), 'r')

print("\nAvailable variables:")
for var in dataset.variables:
    print(var)

latitudes = dataset.variables['lat'][:]
longitudes = dataset.variables['lon'][:]

print(f"\nLatitude count: {len(latitudes)}")
print(f"Longitude count: {len(longitudes)}")

tmp_2m = dataset.variables['TMP_2maboveground'][:]
tp = dataset.variables['tp'][:]

print("\nVariable Shapes:")
print("TMP_2maboveground:", tmp_2m.shape)
print("tp:", tp.shape)

def get_nearest_index(array, value):
    return int(np.abs(array - value).argmin())


lat_lon_pairs = [
    {"LATITUDE": 20.1, "LONGITUDE": 74},
    {"LATITUDE": 8.9, "LONGITUDE": 77.8},
    {"LATITUDE": 15.8, "LONGITUDE": 78}
]
output_data = []

for pair in lat_lon_pairs:
    lat = pair["LATITUDE"]
    lon = pair["LONGITUDE"]

    lat_idx = get_nearest_index(latitudes, lat)
    lon_idx = get_nearest_index(longitudes, lon)

    temperature_kelvin = tmp_2m[:, 0, lat_idx, lon_idx]
    precipitation = tp[:, lat_idx, lon_idx]

    temperature_celsius = temperature_kelvin - 273.15

    for time_step, (temp, precip) in enumerate(zip(temperature_celsius, precipitation)):
        output_data.append({
            "LATITUDE": lat,
            "LONGITUDE": lon,
            "Time Step": time_step,
            "Temperature (°C)": round(float(temp), 2),
            "Total Precipitation (mm)": round(float(precip), 3)
        })
dataset.close()

df = pd.DataFrame(output_data)

if output_excel_path.exists():
    output_excel_path.unlink()
    print(f"Removed existing file: {output_excel_path}")

# Specify encoding and engine for Excel export
df.to_excel(str(output_excel_path), index=False, engine="openpyxl")
print(f"Extracted data saved to: {output_excel_path}")

lat_lon_counts = []

for file_name, file_datetime in recent_files:
    file_path = directory_path / file_name

    try:
        dataset = netCDF4.Dataset(str(file_path), 'r')

        latitudes = dataset.variables['lat'][:]
        longitudes = dataset.variables['lon'][:]

        lat_lon_counts.append({
            "FTP File Name": file_name,
            "FTP Datetime": file_datetime.strftime('%d-%m-%Y %H:%M'),
            "NC_FileLatitude Count": len(latitudes),
            "NC_FileLongitude Count": len(longitudes)
        })

        dataset.close()

        print(f"{file_name} → Lat: {len(latitudes)}, Lon: {len(longitudes)}")

    except Exception as e:
        # Handle Unicode characters in error messages safely
        error_msg = str(e).encode('utf-8', errors='replace').decode('utf-8')
        print(f"Error processing {file_name}: {error_msg}")

lat_lon_df = pd.DataFrame(lat_lon_counts)

# Specify UTF-8 encoding with BOM for Windows Excel compatibility
lat_lon_df.to_csv(str(lat_lon_output_path), index=False, encoding='utf-8-sig')

print(f"Latitude/Longitude counts saved to: {lat_lon_output_path}")
import pandas as pd
from datetime import datetime
import ast  

def extract_and_rearrange_wind_speed(file_path, weather_file_path):
   
    df = pd.read_excel(file_path, sheet_name='Query2_Results')
    df['time'] = pd.to_datetime(df['time'], format='%H:%M').dt.time

    df_filtered = df[df['time'] >= pd.to_datetime('06:00', format='%H:%M').time()]

    result = []
    for (lat, lon), group in df_filtered.groupby(['latitude_name', 'longitude_name']):
     
        group_data = group[['latitude_name', 'longitude_name', 'wind_speed']].reset_index(drop=True)
        result.append(group_data)


    final_df = pd.concat(result, axis=1)


    columns_to_remove = ['latitude_name', 'longitude_name', 'wind_speed']
    final_df = final_df.drop(columns=[col for col in columns_to_remove if col in final_df.columns], errors='ignore')


    weather_df = pd.read_excel(weather_file_path)

    specific_latitudes_longitudes = [
        (8.9, 78.0),
        (17.1, 74.0),
        (22.0, 70.6)
    ]

   
    for lat, lon in specific_latitudes_longitudes:
      
        matching_data = weather_df[(weather_df['LATITUDE'] == lat) & (weather_df['LONGITUDE'] == lon)]

        if not matching_data.empty:

            combined_wind_speeds = []

     
            for _, row in matching_data.iterrows():
            
                wind_speed_data_str = row['WIND_SPEED_925MB']

                if isinstance(wind_speed_data_str, str):
                    try:
                        wind_speed_data = ast.literal_eval(wind_speed_data_str)  
                    except (ValueError, SyntaxError):
                        print(f"Warning: Failed to parse wind speed data for LAT={lat}, LON={lon}, TIMESTAMP_DAY={row['TIMESTAMP_DAY']}")
                        continue
                else:
                    wind_speed_data = wind_speed_data_str  

                combined_wind_speeds.extend(list(wind_speed_data.values()))

            wind_speed_df = pd.DataFrame(
                combined_wind_speeds,
                columns=[f"Wind_Speed_925MB_{lat}_{lon}"]
            )

          
            final_df = pd.concat([final_df, wind_speed_df], axis=1)
        else:
            print(f"Warning: No data found for LAT={lat}, LON={lon} in WeatherDB_data.xlsx")

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name='WeatherDB_Wind_Speed', index=False)

    print(f"Column-wise data saved to {file_path} in the 'WeatherDB_Wind_Speed' sheet.")


    copy_wind_speed_by_date(file_path)

def copy_wind_speed_by_date(file_path):
 
    df = pd.read_excel(file_path, sheet_name='Query2_Results')

    df['weather_date'] = pd.to_datetime(df['weather_date'])


    df['time'] = pd.to_datetime(df['time'], format='%H:%M').dt.time

    result = []

    for (lat, lon), group in df.groupby(['latitude_name', 'longitude_name']):
     
        group = group.sort_values(by=['weather_date', 'time'])

        unique_dates = group['weather_date'].unique()

        for i, weather_date in enumerate(unique_dates):
            if i == 0:
            
                filtered_data = group[(group['weather_date'] == weather_date) & (group['time'] >= datetime.strptime('06:00', '%H:%M').time())]
            else:
        
                filtered_data = group[group['weather_date'] == weather_date]

            for _, row in filtered_data.iterrows():
                result.append({
                    'Latitude': lat,
                    'Longitude': lon,
                    'Weather_Date': weather_date,
                    'Time': row['time'],
                    'Wind_Speed': row['wind_speed']
                })
    result_df = pd.DataFrame(result)

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        result_df.to_excel(writer, sheet_name='RE_6thBlock_WindData', index=False)

    print(f"Data copied to 'RE WindData' sheet in {file_path}.")

def compare_wind_speed_data(file_path):

    weatherdb_df = pd.read_excel(file_path, sheet_name='WeatherDB_Wind_Speed')
    re_df = pd.read_excel(file_path, sheet_name='RE_6thBlock_WindData')

 
    comparison_data = []
    specific_latitudes_longitudes = [
        (8.9, 78.0),
        (17.1, 74.0),
        (22.0, 70.6)
    ]

    non_zero_differences = False

    for lat, lon in specific_latitudes_longitudes:
        wind_speed_column = f"Wind_Speed_925MB_{lat}_{lon}"

        if wind_speed_column in weatherdb_df.columns:
   
            re_filtered = re_df[(re_df['Latitude'] == lat) & (re_df['Longitude'] == lon)]

            for i, wind_speed in enumerate(weatherdb_df[wind_speed_column]):
                re_wind_speed = re_filtered['Wind_Speed'].iloc[i] if i < len(re_filtered) else None
                difference = round(wind_speed - re_wind_speed, 2) if re_wind_speed is not None else None

    
                if difference != 0:
                    non_zero_differences = True

                comparison_data.append({
                    'Latitude': lat,
                    'Longitude': lon,
                    'WeatherDB_Wind_Speed': round(wind_speed, 2),
                    'RE_Wind_Speed': round(re_wind_speed, 2) if re_wind_speed is not None else None,
                    'Difference': difference
                })

    comparison_df = pd.DataFrame(comparison_data)

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        comparison_df.to_excel(writer, sheet_name='Comparison_Result', index=False)

    if non_zero_differences:
        print("Data Mismatches found in the wind speed data.")
    else:
        print("Data Validtaed NO mismatch found.")

    print(f"Comparison data saved to 'Comparison_Result' sheet in {file_path}.")

file_path = r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\query2_results.xlsx'
weather_file_path = r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\WeatherDB_data.xlsx'

extract_and_rearrange_wind_speed(file_path, weather_file_path)
compare_wind_speed_data(file_path)

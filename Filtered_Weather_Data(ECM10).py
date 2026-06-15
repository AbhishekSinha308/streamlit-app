import pandas as pd

file_path = r"\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\WeatherDB_data.xlsx"
output_path = r"\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\Filtered_WeatherDB_Data.xlsx"


data = pd.read_excel(file_path, dtype={'TIMESTAMP_DAY': str})

def process_tmp_column(row):
    tmp_data = row['TMP_2maboveground']
    timestamp = row['TIMESTAMP_DAY']  
    
    try:
        tmp_dict = eval(tmp_data)  
        
       
        filtered_dict = {int(k): round(v, 2) for k, v in tmp_dict.items() if 750 <= int(k) <= 1050}
        
        if not filtered_dict:
            print(f"Alert: Filtered data for timestamp {timestamp} doesn't contain values between 750 and 1410.")
     
        return ','.join(str(v) for v in filtered_dict.values())
    except Exception as e:
        print(f"Error processing row with timestamp {timestamp}: {e}")
        return ""


data['Filtered_TMP'] = data.apply(process_tmp_column, axis=1)

data['TIMESTAMP_DAY'] = data['TIMESTAMP_DAY'].astype(str)


data.to_excel(output_path, index=False)

print(f"Filtered data has been saved to {output_path}")

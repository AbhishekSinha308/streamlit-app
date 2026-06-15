import pandas as pd

file_path =  r"\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\NcFile_data.xlsx"
data = pd.read_excel(file_path)  

print("Columns in the file:", data.columns)
print("Preview of the data:")
print(data.head())

if 'Time Step' in data.columns:
    print("Unique Time Step values:", data['Time Step'].unique())
    print("Time Step data type before conversion:", data['Time Step'].dtype)

    if not pd.api.types.is_numeric_dtype(data['Time Step']):
        data['Time Step'] = pd.to_numeric(data['Time Step'], errors='coerce')
        print("Time Step data type after conversion:", data['Time Step'].dtype)

    print("Minimum Time Step value:", data['Time Step'].min())
    print("Maximum Time Step value:", data['Time Step'].max())

    lower_limit, upper_limit = 13, 18 
    filtered_data = data[(data['Time Step'] >= lower_limit) & (data['Time Step'] <= upper_limit)].copy()
    print("Filtered DataFrame:")
    print(filtered_data)

    if filtered_data.empty:
        print(f"No data found for Time Step between {lower_limit} and {upper_limit}.")
    else:
      
        filtered_data['Temperature (°C)'] = filtered_data['Temperature (°C)'].round(3)
        output_path = r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\Filtered_NcFile_data.xlsx'
        filtered_data.to_excel(output_path, index=False) 
        print(f"Filtered data has been saved to {output_path}")
else:
    print("Time Step column is missing from the input file.")

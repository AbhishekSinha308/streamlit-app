import mysql.connector
from mysql.connector import Error
import pandas as pd
from datetime import datetime
import openpyxl

def fetch_mysql_data(host, user, password, database, query):
    try:
        print(f"Attempting to connect to the {database} database...")

        connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database
        )

        if connection.is_connected():
            print("Connected to the database")

            cursor = connection.cursor()
            print("Executing query...")
            cursor.execute(query)

            result = cursor.fetchall()
            print("Query executed. Fetching results...")

            columns = [desc[0] for desc in cursor.description]
            print(f"Columns: {columns}")

            return result, columns

    except Error as e:
        print(f"Error: {e.errno}, SQLSTATE: {e.sqlstate}, Message: {e.msg}")
        raise 
    finally:
        if 'connection' in locals() and connection.is_connected():
            cursor.close()
            connection.close()
            print("Connection closed")

def save_to_excel(data1, columns1, data2, columns2, file_name=r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\query_results.xlsx'):
    df1 = pd.DataFrame(data1, columns=columns1)
    df2 = pd.DataFrame(data2, columns=columns2)

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='Query1_Results', index=False)
        df2.to_excel(writer, sheet_name='Query2_Results', index=False)

    print(f"Data has been written to {file_name}")

def remove_process_file_name_blocks(file_name=r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\query_results.xlsx', sheet_name="Query1_Results"):
    # Load the workbook and the sheet
    wb = openpyxl.load_workbook(file_name)
    sheet = wb[sheet_name]

    # Get all rows of data starting from row 2 (skip header)
    rows = list(sheet.iter_rows(min_row=2, values_only=True))

    # Prepare a list to hold the indices of rows to keep    
    rows_to_keep = []
    current_lat_lon = None
    block_counter = 0
    row_counter = 0

    for row in rows:
        lat_name, lon_name = row[0], row[1]

        # If the latitude and longitude change, reset the counter for blocks
        if (lat_name, lon_name) != current_lat_lon:
            current_lat_lon = (lat_name, lon_name)
            block_counter = 0

        # For the first 6 rows of a block (after the first row), skip them
        if block_counter < 6:
            block_counter += 1  # Increment the block counter, skip first 6 rows
        else:
            rows_to_keep.append(row_counter)  # Keep the rest of the rows

        row_counter += 1

    # Now delete rows from the sheet that are not in `rows_to_keep`
    # Work in reverse order to prevent index shifting while deleting
    for idx in range(len(sheet['A']), 1, -1):  # Start from the bottom
        if idx - 2 not in rows_to_keep:  # -2 to adjust for the header offset
            sheet.delete_rows(idx)

    # Save the modified workbook
    wb.save(file_name)
    print(f"Process file name blocks have been removed and the data is saved to {file_name}")

if __name__ == "__main__":
    host = "3.6.48.236"
    user = "abhishek"
    password = "Abhishek@1212"

    today_date = datetime.today().strftime('%Y-%m-%d')
    print(f"Using today's date: {today_date}")

    database1 = "spfs"
    query1 = f"""
    SELECT l.latitude_name, e.longitude_name, b.weather_date, b.weather_time, b.process_file_name,
       f.param_name, c.param_value, b.data_source
    FROM 
        spfs.grid_point_config a
    JOIN 
        spfs.latitude l ON a.latitude_uid = l.latitude_uid
    JOIN 
        spfs.longitude e ON a.longitude_uid = e.longitude_uid
    JOIN 
        spfs.weather_forecast b ON a.grid_point_config_uid = b.grid_point_config_uid
    JOIN 
        spfs.weather_forecast_data c ON b.weather_forecast_uid = c.weather_forecast_uid
    JOIN 
        spfs.weather_param f ON f.weather_param_uid = c.weather_param_uid
    WHERE  
        b.weather_date >= '{today_date}'
        AND f.weather_param_uid = 51
        AND (
            (CAST(l.latitude_name AS CHAR) = '20.1' AND CAST(e.longitude_name AS CHAR) = '74') 
            OR (CAST(l.latitude_name AS CHAR) = '8.9' AND CAST(e.longitude_name AS CHAR) = '77.8')
            OR (CAST(l.latitude_name AS CHAR) = '15.8' AND CAST(e.longitude_name AS CHAR) = '78')
            OR (CAST(l.latitude_name AS CHAR) = '26.7' AND CAST(e.longitude_name AS CHAR) = '71.2')
        )
        AND b.data_source = 5
        AND a.data_source = b.data_source;
    """

    database2 = "hawa_new"
    query2 = f"""
    SELECT a.weather_date,a.master_file_name, a.time, b.latitude_name, c.longitude_name, a.wind_direction, a.wind_speed, a.hieght, a.level
    FROM hawa_new.wind_speed_direction a
    JOIN latitude b ON a.latitude = b.id
    JOIN longitude c ON a.longitude = c.id
    WHERE a.weather_date >= '{today_date}'
      AND a.data_source = 10
      AND (
        (b.latitude_name = '17.1' AND c.longitude_name = '74') OR
        (b.latitude_name = '8.9' AND c.longitude_name = '78') OR
        (b.latitude_name = '22' AND c.longitude_name = '70.6') OR
        (b.latitude_name = '17.3' AND c.longitude_name = '73.9')
      )
      AND a.level = '925 hpa'
    ORDER BY 
      CASE 
        WHEN b.latitude_name = '17.1' AND c.longitude_name = '74' THEN 1
        WHEN b.latitude_name = '8.9' AND c.longitude_name = '78' THEN 2
        WHEN b.latitude_name = '22' AND c.longitude_name = '70.6' THEN 3
        WHEN b.latitude_name = '17.3' AND c.longitude_name = '73.9' THEN 4
      END,
      a.weather_date, a.time;
    """

    data1, columns1 = fetch_mysql_data(host, user, password, database1, query1)
    print("Results for the first query:", data1)

    data2, columns2 = fetch_mysql_data(host, user, password, database2, query2)
    print("Results for the second query:", data2)

    save_to_excel(data1, columns1, data2, columns2, r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\query_results.xlsx')
    save_to_excel(data1, columns1, data2, columns2, r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\query2_results.xlsx')

    # Call the function to remove the process file name blocks for both files
    remove_process_file_name_blocks(r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\query_results.xlsx', sheet_name="Query1_Results")
    remove_process_file_name_blocks(r'\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\query2_results.xlsx', sheet_name="Query1_Results")

import sys
import io
import os

# ============================================================
# FIX WINDOWS ENCODING FIRST (before any imports or prints)
# ============================================================
if sys.platform == "win32":
    try:
        # Force UTF-8 encoding on Windows
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', newline='')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', newline='')
    except Exception as e:
        print(f"Warning: Could not set UTF-8 encoding: {e}")

# Now safe to import everything else
import pandas as pd
from pymongo import MongoClient
from ftplib import FTP
from datetime import datetime, timedelta
import re
import getpass
from urllib.parse import quote_plus
import time

print("[INFO] ECM10 Python:", sys.executable)

# ============================================================
# CONFIG PATHS
# ============================================================
SHARE_PATH = r"\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha"
FILE_HISTORY_PATH = os.path.join(SHARE_PATH, "FileHistory.csv")
WEATHER_DB_PATH = os.path.join(SHARE_PATH, "WeatherDB_data.xlsx")
LATLONG_OUTPUT_PATH = os.path.join(SHARE_PATH, "Weather_Updated_LatLong.xlsx")

# Global MongoDB connection (reuse it!)
_MONGO_CLIENT = None

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def validate_environment():
    print("\n========== ECM10 MASTER ENV ==========")
    print("[INFO] User        :", getpass.getuser())
    print("[INFO] Python      :", sys.executable)
    print("[INFO] Working Dir :", os.getcwd())
    print("[INFO] Share Exists:", os.path.exists(SHARE_PATH))
    print("======================================\n")

    if not os.path.exists(SHARE_PATH):
        raise RuntimeError("Network share not accessible")


def convert_utc_to_ist(utc_time):
    """Convert UTC time to IST (UTC+5:30)."""
    return utc_time + timedelta(hours=5, minutes=30)


def fetch_files_for_current_date(ftp, file_prefixes):
    """Fetch all files for the current or previous date, including late uploads."""
    try:
        files = []
        ftp.retrlines("LIST", files.append)
        print("[INFO] Available files on FTP server:")
        for file in files:
            print(f"       {file}")

        pattern = re.compile(rf"({'|'.join(file_prefixes)})_(\d{{2}})(\d{{2}})(\d{{4}})_RE\.nc")
        file_details = []

        for file in files:
            parts = file.split()
            filename = parts[-1]
            match = pattern.search(filename)
            if match:
                month = int(match.group(2))
                day = int(match.group(3))
                year = int(match.group(4))
                file_date = datetime(year, month, day)
                file_date_str = f"{parts[5]} {parts[6]} {parts[7]}"
                utc_time = datetime.strptime(f"{file_date_str} {year}", "%b %d %H:%M %Y")
                ist_time = convert_utc_to_ist(utc_time)

                print(f"[OK]  File: {filename}, UTC Time: {utc_time}, IST Time: {ist_time}")

                cutoff_time = datetime.now() - timedelta(hours=36)
                if utc_time >= cutoff_time:
                    file_details.append((file_date, filename, utc_time, ist_time))

        if not file_details:
            print("[WARNING] No files found matching the criteria.")
            return None, []

        file_details.sort(key=lambda x: x[2], reverse=True)
        latest_file = file_details[0]
        print("\n[INFO] Files for the current or previous date on FTP:")
        for file in file_details:
            print(f"       - {file[1]} (IST: {file[3]}, UTC: {file[2]})")
        return latest_file, file_details
    except Exception as e:
        print(f"[ERROR] Error fetching files for the current date: {e}")
        return None, []


def check_missing_files(file_details, file_prefixes, ftp):
    """Check if any files are missing from the sequence for past dates."""
    missing_files = []
    current_time = datetime.now()
    file_dates = [file[1] for file in file_details]
    previous_date = (datetime.now() - timedelta(days=1)).strftime("%m%d%Y")
    expected_files = [f"{prefix}_{previous_date}_IND.nc" for prefix in file_prefixes]
    expected_times = {
        "00Z": "11:46",
        "06Z": "17:45",
        "12Z": "23:45",
        "18Z": "05:46",
    }

    for expected_file in expected_files:
        if expected_file not in file_dates:
            prefix = expected_file.split("_")[0]
            expected_time_str = expected_times.get(prefix)
            if expected_time_str:
                expected_file_time = datetime.strptime(f"{previous_date} {expected_time_str}", "%m%d%Y %H:%M")
                if prefix == "18Z":
                    expected_file_time += timedelta(days=1)
                if current_time > expected_file_time:
                    ftp_files = []
                    ftp.retrlines("LIST", ftp_files.append)
                    if not any(expected_file in line for line in ftp_files):
                        missing_files.append(expected_file)
    if missing_files:
        print(f"[ALERT] Missing files for past date: {', '.join(missing_files)}")
    else:
        print("[OK]   No files are missing for the past date.")


# ============================================================
# FIXED MONGODB CONNECTION WITH POOLING
# ============================================================

def get_mongo_client():
    """Get or create global MongoDB client with connection pooling."""
    global _MONGO_CLIENT
    
    if _MONGO_CLIENT is not None:
        try:
            _MONGO_CLIENT.admin.command("ping")
            print("[OK]   Reusing existing MongoDB connection")
            return _MONGO_CLIENT
        except:
            _MONGO_CLIENT = None
    
    username = "abhishek"
    password = "Abhishek@7707"
    password_enc = quote_plus(password)

    auth_sources = ["admin", "weather"]
    auth_mechanisms = ["SCRAM-SHA-256", "SCRAM-SHA-1", "DEFAULT"]

    for auth_source in auth_sources:
        for mechanism in auth_mechanisms:
            try:
                conn_str = (
                    f"mongodb://{username}:{password_enc}@13.126.158.236:27017/"
                    f"weather?authSource={auth_source}&authMechanism={mechanism}"
                    f"&maxPoolSize=50&minPoolSize=10"  # Connection pooling
                )
                client = MongoClient(conn_str, serverSelectionTimeoutMS=5000)
                client.admin.command("ping")
                print(f"[OK]   MongoDB connected using authSource='{auth_source}', mechanism='{mechanism}'")
                _MONGO_CLIENT = client
                return client
            except Exception as e:
                print(f"[WARNING] Failed authSource='{auth_source}', mechanism='{mechanism}' -> {e}")

    raise Exception("ERROR: All MongoDB authentication attempts failed")


def fetch_file_history():
    """Fetch the latest records from the FileHistory collection and save CSV."""
    try:
        client = get_mongo_client()
        db = client["weather"]
        file_history_collection = db["FilesHistory"]
        print("\n[INFO] Fetching file history...")

        query = {"dataSource": "ECM"}
        results = list(file_history_collection.find(query).sort([("processedDateTime", -1)]).limit(6))

        if not results:
            print("[WARNING] No records found in FileHistory matching the query.")
            return []

        save_to_csv(results)
        return results
    except Exception as e:
        print(f"[ERROR] Error fetching file history: {e}")
        sys.exit(1)


def save_to_csv(records):
    """Save fetched file history records to a CSV file."""
    data = []
    for record in records:
        processed_date_time = record.get('processedDateTime')
        ist_time = convert_utc_to_ist(processed_date_time) if processed_date_time else None
        data.append({
            "DB_File Name": record.get('fileName'),
            "File Processed into DB at": ist_time.strftime("%Y-%m-%d %H:%M:%S") if ist_time else None,
            "File Status": record.get('fileStatus'),
            "Description": record.get('fileDescription'),
            "File Processed Start Hours": record.get('fromTimeHour'),
            "File Processed End Hours": record.get('toTimeHour')
        })
    df = pd.DataFrame(data)
    df.to_csv(FILE_HISTORY_PATH, index=False)
    print(f"[OK]   File history records saved to {FILE_HISTORY_PATH}")


# ============================================================
# OPTIMIZED DATABASE QUERIES
# ============================================================

def fetch_and_list_updates(latest_file_name):
    """
    OPTIMIZED: Fetch weather data using SINGLE BATCH QUERY instead of 8 separate queries
    OLD: 8 separate database calls (one per coordinate)
    NEW: 1 database call using $or operator
    """
    try:
        client = get_mongo_client()
        db = client["weather"]
        weather_collection = db["WEATHER_DATA"]
        print(f"\n[INFO] Fetching updates for file: {latest_file_name}")

        coordinates = [
            {"LATITUDE": 20.1, "LONGITUDE": 74},
            {"LATITUDE": 8.9, "LONGITUDE": 77.8},
            {"LATITUDE": 15.8, "LONGITUDE": 78},
            {"LATITUDE": 17.1, "LONGITUDE": 74},
            {"LATITUDE": 8.9, "LONGITUDE": 78},
            {"LATITUDE": 22, "LONGITUDE": 70.6},
            {"LATITUDE": 17.3, "LONGITUDE": 73.9},
            {"LATITUDE": 26.7, "LONGITUDE": 71.2}
        ]

        today = datetime.now().date()
        timestamp_filter = datetime.combine(today, datetime.min.time())

        # BUILD SINGLE QUERY WITH $or INSTEAD OF LOOP
        or_conditions = []
        for coords in coordinates:
            or_conditions.append({
                "LATITUDE": coords["LATITUDE"],
                "LONGITUDE": coords["LONGITUDE"]
            })

        # SINGLE DATABASE QUERY (not 8!)
        query = {
            "DATA_SOURCE": "ECM",
            "$or": or_conditions,
            "TIMESTAMP_DAY": {"$gte": timestamp_filter}
        }
        
        print(f"[TIME] Fetching data with single batch query (8 coordinates)...")
        start_time = time.time()
        results = list(weather_collection.find(query).sort("TIMESTAMP_DAY", 1))
        elapsed = time.time() - start_time
        print(f"[OK]   Fetched {len(results)} records in {elapsed:.2f} seconds")

        # Process results
        data_to_save = []
        for r in results:
            data_to_save.append({
                "LATITUDE": r.get("LATITUDE"),
                "LONGITUDE": r.get("LONGITUDE"),
                "TIMESTAMP_DAY": r.get("TIMESTAMP_DAY"),
                "GHI": r.get("GHI"),
                "WIND_SPEED_100M": r.get("WIND_SPEED_100M"),
                "WIND_SPEED_10M": r.get("WIND_SPEED_10M"),
                "WIND_SPEED_925MB": r.get("WIND_SPEED_925MB"),
                "WIND_SPEED_950MB": r.get("WIND_SPEED_950MB"),
                "TMP_2maboveground": r.get("TMP_2maboveground"),
                "TOTAL_PRECIPITATION": r.get("TOTAL_PRECIPITATION"),
                "FILE": latest_file_name
            })

        if data_to_save:
            df = pd.DataFrame(data_to_save)
            if os.path.exists(WEATHER_DB_PATH):
                os.remove(WEATHER_DB_PATH)
            df.to_excel(WEATHER_DB_PATH, index=False, engine='openpyxl')
            print(f"[OK]   Weather data saved to {WEATHER_DB_PATH}")
        else:
            print("[WARNING] No weather data found to save.")
    except Exception as e:
        print(f"[ERROR] Error fetching weather updates: {e}")
        sys.exit(1)


def export_latlong_updates():
    """
    OPTIMIZED: Fetch all data at once instead of looping through file names
    OLD: N separate queries (one per file)
    NEW: 1 or 2 queries using aggregation
    """
    try:
        client = get_mongo_client()
        db = client["weather"]
        weather_collection = db["WEATHER_DATA"]

        file_history_df = pd.read_csv(FILE_HISTORY_PATH)
        file_names = file_history_df['DB_File Name'].tolist()

        print(f"\n[TIME] Processing {len(file_names)} files...")
        start_time = time.time()

        # OPTIMIZED: Use aggregation pipeline instead of loop
        # This performs grouping on the server, not client side
        pipeline = [
            {
                "$match": {
                    "DATA_SOURCE": "ECM",
                    "fileName": {"$in": file_names}
                }
            },
            {
                "$group": {
                    "_id": {
                        "fileName": "$fileName",
                        "LATITUDE": "$LATITUDE",
                        "LONGITUDE": "$LONGITUDE"
                    }
                }
            },
            {
                "$sort": {"_id.fileName": 1}
            }
        ]

        print("[TIME] Running aggregation pipeline on MongoDB server...")
        results = list(weather_collection.aggregate(pipeline))
        elapsed = time.time() - start_time
        print(f"[OK]   Aggregation completed in {elapsed:.2f} seconds")

        # Process results
        updates = []
        counts_dict = {}

        for rec in results:
            file_name = rec["_id"]["fileName"]
            lat = rec["_id"]["LATITUDE"]
            lon = rec["_id"]["LONGITUDE"]
            
            updates.append({
                "File Name": file_name,
                "LATITUDE": lat,
                "LONGITUDE": lon
            })
            
            if file_name not in counts_dict:
                counts_dict[file_name] = 0
            counts_dict[file_name] += 1

        counts = [{"File Name": k, "Count": v} for k, v in counts_dict.items()]

        df_updates = pd.DataFrame(updates)
        df_counts = pd.DataFrame(counts)
        
        with pd.ExcelWriter(LATLONG_OUTPUT_PATH, engine='openpyxl') as writer:
            df_updates.to_excel(writer, sheet_name='Sheet1', index=False)
            df_counts.to_excel(writer, sheet_name='Sheet2', index=False)
        
        print(f"[OK]   Latitude/Longitude updates saved to {LATLONG_OUTPUT_PATH}")
    except Exception as e:
        print(f"[ERROR] Error exporting Lat/Long updates: {e}")
        sys.exit(1)


# ============================================================
# MAIN
# ============================================================
def main():
    server = "weather.50hertz.in"
    username = "analyst"
    password = "M@Lp*267D"

    try:
        ftp = FTP(server)
        ftp.login(user=username, passwd=password)
        print(f"\n[OK]   Logged in successfully to {server}")

        ftp.cwd("/ftp-data/ECMWF_26/RE")
        print("[OK]   Changed directory to /ftp-data/ECMWF_26/RE")

        file_prefixes = ["00Z", "06Z", "12Z", "18Z"]
        latest_file, file_details = fetch_files_for_current_date(ftp, file_prefixes)
        check_missing_files(file_details, file_prefixes, ftp)
        ftp.quit()
        print("[OK]   FTP connection closed.")

        file_history_records = fetch_file_history()
        if latest_file:
            fetch_and_list_updates(latest_file[1])
            export_latlong_updates()
        else:
            print("[WARNING] No files found for the current date.")
        
        print("\n" + "="*50)
        print("[OK]   ALL OPERATIONS COMPLETED SUCCESSFULLY!")
        print("="*50 + "\n")
    except Exception as e:
        print(f"\n[ERROR] Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main() 
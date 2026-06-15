import ftplib
from datetime import datetime, timedelta
import os
import shutil
import re
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from openpyxl import Workbook

SMTP_SERVER = "mx1.50hertz.in"
SMTP_PORT = 25
EMAIL_FROM = "abhishek.sinha@50hertz.in"
EMAIL_TO = ["abhishek.sinha@50hertz.in"]

MIN_SIZE_BYTES = 275 * 1024 * 1024 

def sanitize_filename(name):
    name = re.sub(r'[^\x00-\x7F]+', '', name)
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    return name.strip()
def is_real_file(ftp, name):
    try:
        ftp.size(name)
        return True
    except:
        return False
def safe_copy_to_network(src, dst, retries=10, wait=10):
    for i in range(retries):
        try:
            shutil.copy2(src, dst)
            return True
        except PermissionError:
            print(f"File locked, retrying copy ({i+1}/{retries})...")
            time.sleep(wait)
        except Exception as e:
            print(f"Copy failed ({i+1}): {e}")
            time.sleep(wait)
    return False
def check_file_size(path):
    size_bytes = os.path.getsize(path)
    size_mb = round(size_bytes / (1024 * 1024), 2)
    return size_bytes >= MIN_SIZE_BYTES, size_mb

def extract_date_from_filename(filename):
    try:
        part = filename.split("_")[1] 
        return datetime.strptime(part, "%m%d%Y").date()
    except:
        return None
def send_email(latest_date, previous_date, table_rows_html):
    subject = "ECMWF FTP – Latest & Previous Date Report"

    html_body = f"""
    <html>
    <body style="font-family:Calibri, Arial;">
        <p>Dear Team,</p>

        <p>This report includes <b>ONLY the latest and previous date files</b>.</p>

        <b>FTP DETAILS</b><br>
        Path: /ftp-data/ECMWF_26/RE<br>
        Latest Date: {latest_date}<br>
        Previous Date: {previous_date}<br><br>

        <table border="1" cellpadding="6" cellspacing="0">
            <tr>
                <th>Filename</th>
                <th>Date</th>
                <th>Status</th>
                <th>Size (MB)</th>
            </tr>
            {table_rows_html}
        </table>

        <br><i>This is a system-generated email.</i>
    </body>
    </html>
    """
    try:
        msg = MIMEMultipart("alternative")
        msg["From"] = EMAIL_FROM
        msg["To"] = ", ".join(EMAIL_TO)
        msg["Subject"] = subject
        msg.attach(MIMEText(html_body, "html"))

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
        server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
        server.quit()

        print("Email sent successfully")

    except Exception as e:
        print("Email failed:", e)

def download_latest_and_previous_files(ftp_host, ftp_directory, network_dir, output_excel):

    print("Script started")

    temp_dir = r"C:\Users\Abhishek Sinha\Desktop\PYTHON\_temp"
    os.makedirs(temp_dir, exist_ok=True)
    os.makedirs(network_dir, exist_ok=True)

    excel_data = [("Filename", "Date", "Status", "Size (MB)")]
    email_rows = []

    print("Connecting to FTP...")
    ftp = ftplib.FTP(ftp_host, timeout=120)
    ftp.login("analyst", "M@Lp*267D")
    ftp.cwd(ftp_directory)
    print("FTP connected")
    files = [f for f in ftp.nlst() if f not in (".", "..") and is_real_file(ftp, f)]

    print("\nFiles on FTP:")
    for f in files:
        print(" -", f)

    file_dates = {}
    for f in files:
        file_date = extract_date_from_filename(f)
        if file_date:
            file_dates[f] = file_date

    if not file_dates:
        print("No valid files found based on filename date")
        ftp.quit()
        return

    sorted_files = sorted(file_dates.items(), key=lambda x: x[1], reverse=True)

    latest_date = sorted_files[0][1]
    previous_date = latest_date - timedelta(days=1)

    print(f"Latest date   : {latest_date}")
    print(f"Previous date : {previous_date}")
    latest_files = [f for f, t in sorted_files if t == latest_date]
    previous_files = [f for f, t in sorted_files if t == previous_date]

    if not latest_files:
        print("No files found for latest date")

    if not previous_files:
        print("No files found for previous date")
    def process_files(file_list, file_date):
        for file in file_list:
            safe = sanitize_filename(file)

            final_path = os.path.join(network_dir, safe)
            temp_path = os.path.join(temp_dir, safe)
            temp_downloading = temp_path + ".downloading"

            print(f"\nProcessing: {safe}")

            if os.path.exists(final_path):
                ok, size_mb = check_file_size(final_path)
                status = "Skipped (Already Exists)" if ok else "Size Issue (<275MB)"
            else:
                try:
                    with open(temp_downloading, "wb") as f:
                        ftp.retrbinary(f"RETR {file}", f.write)
                        f.flush()
                        os.fsync(f.fileno())

                    time.sleep(2)
                    os.rename(temp_downloading, temp_path)

                except Exception as e:
                    print("FTP download failed:", e)
                    status = "Download Failed"
                    size_mb = 0
                    excel_data.append((safe, file_date, status, size_mb))
                    continue

                if safe_copy_to_network(temp_path, final_path):
                    os.remove(temp_path)
                    ok, size_mb = check_file_size(final_path)
                    status = "Downloaded" if ok else "Downloaded (Size Issue)"
                else:
                    status = "Copy Failed"
                    size_mb = 0

            excel_data.append((safe, file_date, status, size_mb))
            email_rows.append(
                f"<tr><td>{safe}</td><td>{file_date}</td><td>{status}</td><td>{size_mb}</td></tr>"
            )

    process_files(latest_files, latest_date)
    process_files(previous_files, previous_date)

    ftp.quit()
    print("FTP connection closed")
    wb = Workbook()
    ws = wb.active
    ws.title = "FTP Report"
    for row in excel_data:
        ws.append(row)
    wb.save(output_excel)

    print(f"Excel saved: {output_excel}")

    send_email(latest_date, previous_date, "".join(email_rows))

    print("Script finished")

download_latest_and_previous_files(
    ftp_host="35.154.41.156",
    ftp_directory="/ftp-data/ECMWF_26/RE",
    network_dir=r"\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha",
    output_excel=r"\\172.16.0.65\share\24_23_QA_Team\Abhishek_Sinha\FTP_File_Name.xlsx"
)
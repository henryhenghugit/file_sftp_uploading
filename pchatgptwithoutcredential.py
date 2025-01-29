import os
import openpyxl
import paramiko
import datetime
import logging
from typing import List

# Constants
SFTP_SERVER = ''
SFTP_USERNAME = ''
SFTP_PASSWORD = ''  # Securely retrieve this
SFTP_BASE_DIR = ''
SEARCH_DIRECTORIES = [
    "Z:\\ts_utp\\Assets\\University of Toronto Press\\Publishing", 
    "Z:\\ts_utp\\Assets\\U of Toronto Press\\Publishing"
]
EXCEL_PATH = "C:\\Users\\MIS\\Desktop\\HH\\ISBNs-for-file-transfer_15Jan2025.xlsx"
LOGS_PATH = "C:\\Users\\MIS\\Desktop\\HH\\logs\\"

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Functions
def setup_sftp_connection():
    """Setup the SFTP connection."""
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(hostname=SFTP_SERVER, username=SFTP_USERNAME, password=SFTP_PASSWORD, port=22)
    sftp = ssh.open_sftp()
    return sftp, ssh

def ensure_remote_directory_exists(sftp, remote_path: str):
    """Ensure the remote directory exists, create if not."""
    try:
        sftp.stat(remote_path)
    except FileNotFoundError:
        sftp.mkdir(remote_path)

def upload_file(sftp, local_path: str, remote_path: str):
    """Upload a file to the remote server."""
    local_path = local_path.replace("\\", "\\\\")
    sftp.put(local_path, remote_path)

def read_isbns_from_excel(path: str) -> List[str]:
    """Read ISBNs and TIDs from the Excel file."""
    wb = openpyxl.load_workbook(path)
    ws = wb['Sheet1']
    return [(isbn.value, tid.value) for isbn, tid in ws.iter_rows(min_row=2, min_col=1)]

def main():
    current_datetime = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    notfound_isbn = []
    uploaded_isbn = []

    file_name_notfound = f"notfound_{current_datetime}.log"
    file_name_upload = f"found_{current_datetime}.log"

    # Open logs
    notfound_log = open(os.path.join(LOGS_PATH, file_name_notfound), 'w')
    uploaded_log = open(os.path.join(LOGS_PATH, file_name_upload), 'w')

    # Setup SFTP
    sftp, ssh = setup_sftp_connection()

    try:
        logging.info(f"Current working directory: {sftp.getcwd()}")
        logging.info(f"Remote folder contents: {sftp.listdir()}")

        # Process ISBNs from Excel
        isbns_and_tids = read_isbns_from_excel(EXCEL_PATH)

        for isbn, tid in isbns_and_tids:
            found = False

            # Search for local files
            for directory in SEARCH_DIRECTORIES:
                for root, dirs, files in os.walk(directory):
                    for file in files:
                        local_path = os.path.join(root, file)

                        if str(isbn) in os.path.basename(local_path) and any(ext in file for ext in ["jpg", "pdf", "epub"]):
                            found = True
                            remote_path = os.path.join(SFTP_BASE_DIR, str(tid)).replace("\\", "/")
                            print ("remote_path is " + remote_path)
                            remote_file_path = os.path.join(remote_path, str(isbn), os.path.basename(local_path)).replace("\\", "/")
                            print ("remote_file_path is " + remote_file_path)

                            logging.info(f"Found a local file: {local_path}")
                            ensure_remote_directory_exists(sftp, remote_path)
                            ensure_remote_directory_exists(sftp, os.path.dirname(remote_file_path))

                            upload_file(sftp, local_path, remote_file_path)

                            uploaded_isbn.append(str(isbn))
                            uploaded_log.write(f"{isbn}\n")
                            break
                if found:
                    break

            if not found:
                notfound_isbn.append(str(isbn))

        # Write not found ISBNs to log
        unique_notfound_isbns = set(notfound_isbn)
        for isbn in unique_notfound_isbns:
            notfound_log.write(f"{isbn}\n")

    except Exception as e:
        logging.error(f"An error occurred: {e}")
    
    finally:
        # Clean up
        notfound_log.close()
        uploaded_log.close()
        sftp.close()
        ssh.close()

if __name__ == "__main__":
    main()
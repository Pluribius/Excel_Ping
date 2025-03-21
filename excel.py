import openpyxl
import subprocess
import time

def ping_ips_from_excel():
    """
    Prompts the user for an Excel file path and a column letter,
    then pings each IP address found in that column until an empty cell is encountered.
    """
    while True:
        excel_path = input("Enter the path to your Excel file: ")
        try:
            workbook = openpyxl.load_workbook(excel_path)
            break
        except FileNotFoundError:
            print("Error: File not found. Please enter a valid path.")
        except Exception as e:
            print(f"Error opening the Excel file: {e}")
            return

    while True:
        column_letter = input("Enter the column letter containing the IP addresses (e.g., A, B, C): ").upper()
        if not column_letter.isalpha() or len(column_letter) != 1:
            print("Invalid input. Please enter a single letter for the column.")
        else:
            break

    sheet = workbook.active  # Assuming you want to use the active sheet

    if sheet.max_row == 0:
        print("The Excel sheet is empty.")
        return

    print("\nStarting to ping IP addresses...")
    row_num = 1
    while True:
        cell = sheet[f"{column_letter}{row_num}"]
        ip_address = cell.value

        if ip_address is None or str(ip_address).strip() == "":
            print("\nReached an empty cell. Stopping the process.")
            break

        if not is_valid_ip(str(ip_address)):
            print(f"Skipping invalid IP address: {ip_address} in row {row_num}")
            row_num += 1
            continue

        print(f"\nPinging IP address: {ip_address} (Row {row_num})")
        try:
            # Construct the ping command based on the operating system
            if os.name == 'nt':  # Windows
                command = ['ping', '-n', '1', '-w', '1000', str(ip_address)]
            else:  # Linux and macOS
                command = ['ping', '-c', '1', '-W', '1', str(ip_address)]

            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            stdout, stderr = process.communicate()
            return_code = process.returncode

            if return_code == 0:
                print(f"  {ip_address} is reachable.")
            else:
                print(f"  {ip_address} is NOT reachable.")
                if stderr:
                    print(f"  Error: {stderr.decode('utf-8').strip()}")

        except FileNotFoundError:
            print("Error: 'ping' command not found on this system.")
            break
        except Exception as e:
            print(f"An error occurred while pinging {ip_address}: {e}")

        time.sleep(1)  # Wait for a short duration between pings
        row_num += 1

def is_valid_ip(ip_str):
    """
    Checks if a given string is a valid IPv4 address.
    """
    parts = ip_str.split('.')
    if len(parts) != 4:
        return False
    for part in parts:
        if not part.isdigit():
            return False
        if not 0 <= int(part) <= 255:
            return False
    return True

if __name__ == "__main__":
    import os
    ping_ips_from_excel()
import ctypes
import datetime
import time
import uuid
import pyodbc
from concurrent.futures import ThreadPoolExecutor
from threading import Thread
import pythoncom
import win32com.client
import os
import subprocess
import sys

# Windows API definitions
class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_ulong)]

kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32

# Database configuration for SQL Server
MAIN_DB_CONNECTION = {
    "server": "XCNTCN01",
    "database": "monitoring",
    "username": "",
    "password": "",
    "driver": "{ODBC Driver 17 for SQL Server}"
}

def get_laptop_serial_number():
    try:
        pythoncom.CoInitialize()
        obj = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        svc = obj.ConnectServer('.', 'root\\cimv2')
        items = svc.ExecQuery("SELECT SerialNumber FROM Win32_BIOS")
        return items[0].SerialNumber.strip()
    except Exception as e:
        print(f"Error getting serial number: {e}")
        return f"UNKNOWN_{uuid.uuid4()}"


def is_database_online():
    try:
        conn_str = f"DRIVER={MAIN_DB_CONNECTION['driver']};SERVER={MAIN_DB_CONNECTION['server']};DATABASE={MAIN_DB_CONNECTION['database']};Trusted_Connection=yes"
        conn = pyodbc.connect(conn_str, timeout=3)
        conn.close()
        return True
    except pyodbc.Error as e:
        print(f"Database connection error: {e}")
        return False

def get_idle_time_seconds():
    lastInputInfo = LASTINPUTINFO()
    lastInputInfo.cbSize = ctypes.sizeof(lastInputInfo)
    if not user32.GetLastInputInfo(ctypes.byref(lastInputInfo)):
        return 0
    current_ticks = kernel32.GetTickCount64()
    return (current_ticks - lastInputInfo.dwTime) // 1000

class UptimeAgent:
    def __init__(self):
        self.session_start = time.time()
        self.laptop_sn = get_laptop_serial_number()
        self.is_running = True
        self.executor = ThreadPoolExecutor(max_workers=2)
        self.offline_uptime = 0
        self.offline_idle_time = 0
        self.previous_idle_time = 0
        self.per_minute_idle_time = 0 # Initialize variable to track idle time within a minute

    def track_data(self):
        while self.is_running:
            start_of_minute = datetime.datetime.now().replace(second=0, microsecond=0)
            end_of_minute = start_of_minute + datetime.timedelta(minutes=1)
            self.per_minute_idle_time = 0 # Reset idle time at the start of each minute

            while datetime.datetime.now() < end_of_minute and self.is_running:
                time.sleep(1) # Check every second
                idle_time_seconds = get_idle_time_seconds()
                print(f"Current idle: {idle_time_seconds}")
                if idle_time_seconds > self.previous_idle_time: # Assuming idle time increases if the device is idle
                    self.per_minute_idle_time += 1
                self.previous_idle_time = idle_time_seconds

            current_time = datetime.datetime.now().replace(second=0, microsecond=0)
            date = current_time.date().isoformat()
            time_str = current_time.strftime('%H:%M:%S')
            uptime = 60 # A full minute has passed

            print(f"End of minute. Total idle time this minute: {self.per_minute_idle_time} seconds")

            if is_database_online():
                print("Online: Mengupload data ke database.")
                self.upload_to_main_db(date, time_str, uptime, self.per_minute_idle_time) # Upload per-minute idle time
                if self.offline_uptime > 0:
                    print(f"Mengupload data akumulasi offline: {self.offline_uptime} detik uptime, {self.offline_idle_time} detik idle")
                    self.upload_to_main_db(date, time_str, self.offline_uptime, self.offline_idle_time)
                    self.offline_uptime = 0
                    self.offline_idle_time = 0
            else:
                print("Offline: Menyimpan data sementara.")
                self.offline_uptime += uptime
                self.offline_idle_time += self.per_minute_idle_time # Accumulate per-minute idle time

    def get_system_info(self):
        try:
            pythoncom.CoInitialize()
            obj = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            svc = obj.ConnectServer('.', 'root\\cimv2')

            system_info = svc.ExecQuery("SELECT Manufacturer, Model FROM Win32_ComputerSystem")
            system_manufacture = system_info[0].Manufacturer.strip() if system_info else "UNKNOWN"
            system_model = system_info[0].Model.strip() if system_info else "UNKNOWN"

            bios_info = svc.ExecQuery("SELECT Version FROM Win32_BIOS")
            bios_version = bios_info[0].Version.strip() if bios_info else "UNKNOWN"

            user_name_microsoft=os.getlogin()

            return {
                'SYSTEM_MANUFACTURE': system_manufacture,
                'BIOS_VERSION': bios_version,
                'SYSTEM_MODEL': system_model,
                'USER_NAME_MICROSOFT': user_name_microsoft
            }

        except Exception as e:
            print(f"Error getting system info: {e}")
            return {
                'SYSTEM_MANUFACTURE': "ERROR",
                'BIOS_VERSION': "ERROR",
                'SYSTEM_MODEL': "ERROR",
                'USER_NAME_MICROSOFT': os.environ.get('USERNAME', 'UNKNOWN')
            }

    def upload_to_main_db(self, date, time_str, uptime, idle_time):
        conn = None
        try:
            conn_str = f"DRIVER={MAIN_DB_CONNECTION['driver']};SERVER={MAIN_DB_CONNECTION['server']};DATABASE={MAIN_DB_CONNECTION['database']};Trusted_Connection=yes"
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()

            cursor.execute("SELECT SN FROM laptops WHERE SN = ?", (self.laptop_sn,))
            if not cursor.fetchone():
                print(f"Laptop SN {self.laptop_sn} tidak ditemukan di database.")
                return

            cursor.execute("""
                SELECT date, uptime, idle_time FROM (
                    SELECT date, uptime, idle_time, ROW_NUMBER() OVER (ORDER BY date DESC, time DESC) as rn
                    FROM daily_uptimes
                    WHERE laptop_sn = ?
                ) as latest_record
                WHERE rn = 1
            """, (self.laptop_sn,))
            row = cursor.fetchone()

            today = datetime.datetime.strptime(date, "%Y-%m-%d").date()

            if row:
                last_date, existing_uptime, existing_idle_time = row

                if last_date == today:
                    new_uptime = existing_uptime + uptime
                    # Now 'idle_time' parameter contains the per-minute total idle time
                    new_idle_time = existing_idle_time + idle_time
                    print(f"Updating uptime and idle time: uptime={new_uptime}, idle_time={new_idle_time}")

                    cursor.execute("""
                        UPDATE daily_uptimes
                        SET uptime = ?, idle_time = ?, updated_at = GETDATE()
                        WHERE laptop_sn = ? AND date = ?
                    """, (new_uptime, new_idle_time, self.laptop_sn, today))
                else:
                    print(f"Inserting new uptime and idle time: uptime={uptime}, idle_time={idle_time}")

                    cursor.execute("""
                        INSERT INTO daily_uptimes (laptop_sn, date, time, uptime, idle_time, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, GETDATE(), GETDATE())
                    """, (self.laptop_sn, date, time_str, uptime, int(idle_time)))
            else:
                print(f"Inserting initial uptime and idle time: uptime={uptime}, idle_time={idle_time}")

                cursor.execute("""
                    INSERT INTO daily_uptimes (laptop_sn, date, time, uptime, idle_time, created_at, updated_at)
                    VALUES (?, ?, ?, ?, ?, GETDATE(), GETDATE())
                """, (self.laptop_sn, date, time_str, uptime, int(idle_time)))

            conn.commit()
        except Exception as e:
            print(f"Error: {e}")
        finally:
            if conn:
                conn.close()

if __name__ == "__main__":
    try:
        agent = UptimeAgent()
        agent.track_data()
    except KeyboardInterrupt:
        print("Agent terminated.")
    except Exception as e:
        print(f"An unexpected error occurred in main loop: {e}")
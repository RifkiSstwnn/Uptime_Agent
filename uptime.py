import ctypes
import datetime
import time
import uuid
import win32com.client
from concurrent.futures import ThreadPoolExecutor
from threading import Thread
import pythoncom
import mysql.connector

# Windows API definitions
class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_ulong)]

kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32

# Database configuration
MAIN_DB_CONNECTION = {
    "host": "localhost",
    "user": "root",
    "password": "",
    "database": "db_ujicoba"
}

def get_laptop_serial_number():
    try:
        pythoncom.CoInitialize()
        obj = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        svc = obj.ConnectServer('.', 'root\\cimv2')
        items = svc.ExecQuery("SELECT SerialNumber FROM Win32_BIOS")
        return items[0].SerialNumber.strip()
    except Exception as e:
        return f"UNKNOWN_{uuid.uuid4()}"

def get_idle_time_seconds():
    lastInputInfo = LASTINPUTINFO()
    lastInputInfo.cbSize = ctypes.sizeof(lastInputInfo)
    if not user32.GetLastInputInfo(ctypes.byref(lastInputInfo)):
        return 0
    current_ticks = kernel32.GetTickCount64()
    return (current_ticks - lastInputInfo.dwTime) // 1000

def is_database_online():
    try:
        conn = mysql.connector.connect(**MAIN_DB_CONNECTION)
        conn.close()
        return True
    except mysql.connector.Error:
        return False

class UptimeAgent:
    def __init__(self):
        self.session_start = time.time()
        self.laptop_sn = get_laptop_serial_number()
        self.is_running = True
        self.executor = ThreadPoolExecutor(max_workers=2)
        
        # Variabel akumulasi saat offline
        self.offline_uptime = 0
        self.offline_idle_time = 0

    def track_data(self):
        while self.is_running:
            next_minute = datetime.datetime.now().replace(second=0, microsecond=0) + datetime.timedelta(minutes=1)
            time.sleep((next_minute - datetime.datetime.now()).total_seconds())

            current_time = datetime.datetime.now().replace(second=0, microsecond=0)
            date = current_time.date().isoformat()
            time_str = current_time.strftime('%H:%M:%S')

            uptime = 60  # Tiap menit, uptime bertambah 60 detik
            idle_time = get_idle_time_seconds()

            if is_database_online():
                print("Online: Mengupload data ke database.")
                self.upload_to_main_db(date, time_str, uptime, idle_time)
                
                # Jika ada data offline, kirim juga
                if self.offline_uptime > 0:
                    print(f"Mengupload data akumulasi offline: {self.offline_uptime} detik uptime, {self.offline_idle_time} detik idle")
                    self.upload_to_main_db(date, time_str, self.offline_uptime, self.offline_idle_time)
                    
                    # Reset akumulasi setelah berhasil diunggah
                    self.offline_uptime = 0
                    self.offline_idle_time = 0
            else:
                print("Offline: Menyimpan data sementara.")
                self.offline_uptime += uptime
                self.offline_idle_time += idle_time

    def upload_to_main_db(self, date, time_str, uptime, idle_time):
        conn = None
        try:
            conn = mysql.connector.connect(**MAIN_DB_CONNECTION)
            cursor = conn.cursor()

            cursor.execute("SELECT SN FROM laptops WHERE SN = %s", (self.laptop_sn,))
            if not cursor.fetchone():
                print(f"Laptop SN {self.laptop_sn} tidak ditemukan di database.")
                return

            # Cek apakah ada data sebelumnya
            cursor.execute('''
                SELECT uptime, idle_time FROM (
                    SELECT uptime, idle_time FROM daily_uptimes 
                    WHERE laptop_sn = %s 
                    ORDER BY date DESC, time DESC 
                    LIMIT 1
                ) AS latest_record
            ''', (self.laptop_sn,))
            row = cursor.fetchone()

            if row:
           # Update existing record by accumulating uptime and idle_time
                existing_uptime, existing_idle_time = row
                new_uptime = existing_uptime + uptime  # Add 60 seconds for each minute
                new_idle_time = existing_idle_time + int(idle_time)
                cursor.execute('''
                    UPDATE daily_uptimes
                    SET uptime = %s, idle_time = %s, updated_at = NOW()
                    WHERE laptop_sn = %s AND date = (
                        SELECT MAX(date) FROM (
                            SELECT date FROM daily_uptimes WHERE laptop_sn = %s
                        ) AS latest_date
                    ) AND time = (
                        SELECT MAX(time) FROM (
                            SELECT time FROM daily_uptimes WHERE laptop_sn = %s AND date = (
                                SELECT MAX(date) FROM daily_uptimes WHERE laptop_sn = %s
                            )
                        ) AS latest_time
                    )
                ''', (new_uptime, new_idle_time, self.laptop_sn, self.laptop_sn, self.laptop_sn, self.laptop_sn))
            else:
                # Insert new record
                cursor.execute('''
                    INSERT INTO daily_uptimes (laptop_sn, date, time, uptime, idle_time, created_at, updated_at)
                    VALUES (%s, %s, %s, %s, %s, NOW(), NOW())
                ''', (self.laptop_sn, date, time_str, 60, int(idle_time)))  # Set initial uptime to 60 seconds

            conn.commit()
        except mysql.connector.Error as e:
            print(f"Main DB error: {e}")
        finally:
            if conn:
                conn.close()

if __name__ == "__main__":
    try:
        agent = UptimeAgent()
        agent.track_data()
    except KeyboardInterrupt:
        print("Agent terminated.")

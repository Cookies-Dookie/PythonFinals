from dotenv import load_dotenv
import pyodbc 
import os
from faculty_model import Faculty
from datetime import timedelta, datetime


class DatabaseManager:
    """Handles all communication with the MSSQL database using pyodbc."""

    def __init__(self):
        load_dotenv()  
        self.driver = '{ODBC Driver 17 for SQL Server}' 
        self.server = os.getenv('DB_SERVER')
        self.database = os.getenv('DB_NAME')
        self.user = os.getenv('DB_USER')
        self.password = os.getenv('DB_PASSWORD')

        self.conn_str = (
            f'DRIVER={self.driver};'
            f'SERVER={self.server};'
            f'DATABASE={self.database};'
            f'UID={self.user};'
            f'PWD={self.password}'
        )

    def execute_non_query(self, sql_query, params=None):
        """Executes INSERT, UPDATE, DELETE queries."""
        conn = None
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            
            if params:
                cursor.execute(sql_query, params)
            else:
                cursor.execute(sql_query)
                
            conn.commit() 
            return True, cursor.rowcount
        except pyodbc.Error as err:
            return False, str(err)
        finally:
            if conn:
                conn.close()

    def add_faculty(self, faculty_id, full_name, department):
        """Inserts a new faculty member into the dbo.Faculty table."""
        if self.load_faculty_info(faculty_id):
            return False, f"Faculty ID {faculty_id} already exists."
            
        sql = "INSERT INTO dbo.Faculty (FacultyID, FullName, Department) VALUES (?, ?, ?)"
        params = (faculty_id, full_name, department)
        return self.execute_non_query(sql, params)

    def load_faculty_info(self, faculty_id):
        """Fetches faculty details based on ID from dbo.Faculty."""
        conn = None
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()
            
            sql = "SELECT FacultyID, FullName, Department FROM dbo.Faculty WHERE FacultyID = ?"
            cursor.execute(sql, (faculty_id,))
            
            result = cursor.fetchone()

            if result:
                return Faculty(result[0], result[1], result[2] if result[2] else "N/A")
            return None
        except pyodbc.Error as err:
            print(f"Read Error: {err}")
            return None
        finally:
            if conn:
                conn.close()

    def get_attendance_report(self):
        """Fetches all faculty and applies the 8-hour shift logic."""
        conn = None
        report_data = []
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()

            sql = """
            SELECT 
                F.FacultyID, 
                F.FullName, 
                T1.Timestamp AS LastActionTime,
                T1.Action AS LastAction,
                (SELECT MAX(Timestamp) FROM dbo.Attendance WHERE FacultyID = F.FacultyID AND Action = 'Check-In' 
                 AND Timestamp < T1.Timestamp) AS PreviousCheckInTime 
            FROM dbo.Faculty AS F
            OUTER APPLY (
                SELECT TOP 1 *
                FROM dbo.Attendance AS A
                WHERE A.FacultyID = F.FacultyID
                ORDER BY A.Timestamp DESC
            ) AS T1
            ORDER BY F.FullName;
            """
            cursor.execute(sql)
            results = cursor.fetchall()
            
            for row in results:
                data = {
                    'FacultyID': row[0],
                    'FullName': row[1],
                    'LastActionTime': row[2],
                    'LastAction': row[3],
                    'PreviousCheckInTime': row[4],
                    'HoursRendered': None,
                    'Note': "No Records"
                }
                
                if data['LastAction'] == 'Check-In':
                    data['Note'] = "Time-In" 
                elif data['LastAction'] == 'Check-Out' and data['PreviousCheckInTime']:
                    time_diff = data['LastActionTime'] - data['PreviousCheckInTime']
                    data['HoursRendered'] = time_diff
                    
                    if time_diff >= timedelta(hours=8):
                        data['Note'] = "**properly done**"
                    else:
                        data['Note'] = "**not done**"
                
                report_data.append(data)

            return report_data
        
        except pyodbc.Error as err:
            print(f"Report Error: {err}")
            return []
        finally:
            if conn:
                conn.close()

    def get_raw_time_data(self):
        """Fetches all Check-In and Check-Out times for today from dbo.Attendance."""
        conn = None
        try:
            conn = pyodbc.connect(self.conn_str)
            cursor = conn.cursor()

            sql = """
            SELECT 
                FacultyID, 
                Timestamp, 
                Action 
            FROM dbo.Attendance 
            WHERE CAST(Timestamp AS DATE) = CAST(GETDATE() AS DATE) 
            ORDER BY Timestamp;
            """
            cursor.execute(sql)
            results = cursor.fetchall()
            
            organized_data = {}
            for row in results:
                fid, timestamp, action = row[0], row[1], row[2]
                if fid not in organized_data:
                    organized_data[fid] = {'check_ins': [], 'check_outs': []}
                
                if action == 'Check-In':
                    organized_data[fid]['check_ins'].append(timestamp)
                elif action == 'Check-Out':
                    organized_data[fid]['check_outs'].append(timestamp)

            return organized_data
        
        except pyodbc.Error as err:
            print(f"Analytics Data Error: {err}")
            return {}
        finally:
            if conn:
                conn.close()
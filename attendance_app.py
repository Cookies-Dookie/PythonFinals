# attendance_app.py

import tkinter as tk
from tkinter import messagebox, ttk
from db_manager import DatabaseManager
from datetime import timedelta, datetime, time
import matplotlib.pyplot as plt
import numpy as np 
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Define professional fonts
LARGE_FONT = ('Arial', 35, 'bold') 
MEDIUM_FONT = ('Arial', 25, 'bold')
INPUT_FONT = ('Arial', 16)
SMALL_FONT = ('Arial', 12)

class AttendanceApp(tk.Tk):
    """Main GUI Application using Tkinter and pyodbc for MSSQL."""

    def __init__(self):
        super().__init__()
        self.title("Faculty Attendance System - MSSQL Edition")
        self.geometry("800x600") 
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- DB Manager Initialization ---
        # Using a RAW STRING to correctly handle the backslash in the instance name.
        self.db_manager = DatabaseManager(server=r'sql.bsite.net\MSSQL2016', database='aspnetfp_') 

        container = tk.Frame(self)
        container.grid(row=0, column=0, sticky='nsew')
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        
        for F in (AttendanceFrame, AddFacultyFrame, ReportFrame, AnalyticsFrame): 
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew") 

        self.create_navigation_menu()
        self.show_frame("AttendanceFrame")

    def create_navigation_menu(self):
        """Creates the navigation menu bar at the top with enlarged fonts."""
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=0, font=MEDIUM_FONT) 
        menu_bar.add_cascade(label="Navigation", menu=file_menu, font=MEDIUM_FONT) 

        file_menu.add_command(label="1. Attendance (In/Out)", command=lambda: self.show_frame("AttendanceFrame"))
        file_menu.add_command(label="2. Add Faculty", command=lambda: self.show_frame("AddFacultyFrame"))
        file_menu.add_command(label="3. View Report", command=lambda: self.show_frame("ReportFrame"))
        file_menu.add_command(label="4. View Analytics", command=lambda: self.show_frame("AnalyticsFrame"))
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)

    def show_frame(self, page_name):
        """Brings the requested frame to the front and triggers refresh if needed."""
        frame = self.frames[page_name]
        if page_name == "ReportFrame":
            frame.refresh_report() 
        elif page_name == "AnalyticsFrame":
            frame.generate_analytics()
        frame.tkraise()

# --- FRAME 1: ATTENDANCE (Check In/Out - VERTICAL BUTTONS) ---
class AttendanceFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1) 
        self.grid_rowconfigure(3, weight=1) 
        self.grid_columnconfigure(0, weight=1) 

        tk.Label(self, text="Faculty Attendance (Check In/Out)", font=LARGE_FONT).grid(row=0, column=0, pady=30, padx=20, sticky='ew')
        
        input_frame = tk.Frame(self)
        input_frame.grid(row=1, column=0, pady=10, sticky='ew')
        input_frame.grid_columnconfigure(1, weight=1)
        
        tk.Label(input_frame, text="Faculty ID:", font=MEDIUM_FONT).grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.id_entry = tk.Entry(input_frame, font=INPUT_FONT)
        self.id_entry.grid(row=0, column=1, padx=10, pady=10, sticky='ew')

        # --- Button Placement (VERTICAL STACK - Smaller, Resizeable) ---
        
        check_in_btn = tk.Button(self, text="CHECK IN", font=MEDIUM_FONT, bg='green', fg='white', 
                                 command=lambda: self.check_in_out("Check-In"))
        check_in_btn.grid(row=2, column=0, padx=200, pady=10, sticky='nsew') 

        check_out_btn = tk.Button(self, text="CHECK OUT", font=MEDIUM_FONT, bg='red', fg='white', 
                                  command=lambda: self.check_in_out("Check-Out"))
        check_out_btn.grid(row=3, column=0, padx=200, pady=10, sticky='nsew')

    def check_in_out(self, action):
        faculty_id = self.id_entry.get().strip()

        if not faculty_id:
            messagebox.showwarning("Input Error", "Please enter Faculty ID.")
            return

        faculty_obj = self.controller.db_manager.load_faculty_info(faculty_id)

        if not faculty_obj:
            messagebox.showerror("Validation Error", f"Faculty ID '{faculty_id}' not found.")
            return

        sql = "INSERT INTO dbo.Attendance (FacultyID, Action) VALUES (?, ?)"
        params = (faculty_id, action)

        success, result = self.controller.db_manager.execute_non_query(sql, params)

        if success:
            messagebox.showinfo("Success", f"Attendance for {faculty_obj.full_name} has been recorded.\nAction: {action}.")
        else:
            messagebox.showerror("Error", f"Failed to record {action}. DB Error: {result}")

        self.id_entry.delete(0, tk.END) 

# --- FRAME 2: ADD FACULTY ---
class AddFacultyFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(3, weight=0)
        self.grid_rowconfigure(4, weight=1)

        tk.Label(self, text="Add New Faculty Member", font=LARGE_FONT).grid(row=0, column=0, pady=30, padx=20, sticky='ew')

        input_frame = tk.Frame(self)
        input_frame.grid(row=1, column=0, pady=10, sticky='ew')
        input_frame.grid_columnconfigure(1, weight=1)
        
        tk.Label(input_frame, text="Faculty ID:", font=MEDIUM_FONT).grid(row=0, column=0, padx=10, pady=10, sticky='w')
        self.id_entry = tk.Entry(input_frame, font=INPUT_FONT)
        self.id_entry.grid(row=0, column=1, padx=10, pady=10, sticky='ew')

        tk.Label(input_frame, text="Full Name:", font=MEDIUM_FONT).grid(row=1, column=0, padx=10, pady=10, sticky='w')
        self.name_entry = tk.Entry(input_frame, font=INPUT_FONT)
        self.name_entry.grid(row=1, column=1, padx=10, pady=10, sticky='ew')
        
        tk.Label(input_frame, text="Department:", font=MEDIUM_FONT).grid(row=2, column=0, padx=10, pady=10, sticky='w')
        self.dept_entry = tk.Entry(input_frame, font=INPUT_FONT)
        self.dept_entry.grid(row=2, column=1, padx=10, pady=10, sticky='ew')

        tk.Button(self, text="ADD FACULTY", font=MEDIUM_FONT, bg='blue', fg='white', 
                  command=self.add_worker_to_db).grid(row=3, column=0, pady=20, padx=200, sticky='ew')
    
    def add_worker_to_db(self):
        faculty_id = self.id_entry.get().strip()
        full_name = self.name_entry.get().strip()
        department = self.dept_entry.get().strip()
        
        if not all([faculty_id, full_name, department]):
            messagebox.showwarning("Input Error", "All fields are required.")
            return

        success, result = self.controller.db_manager.add_faculty(faculty_id, full_name, department)

        if success:
            messagebox.showinfo("Success", f"Faculty {full_name} (ID: {faculty_id}) added successfully.")
            self.id_entry.delete(0, tk.END)
            self.name_entry.delete(0, tk.END)
            self.dept_entry.delete(0, tk.END)
        else:
            messagebox.showerror("Error", f"Failed to add faculty. Error: {result}")

# --- FRAME 3: REPORT VIEW ---
class ReportFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        tk.Label(self, text="Attendance Status Report", font=LARGE_FONT).grid(row=0, column=0, pady=20, sticky='ew')
        
        style = ttk.Style()
        style.configure("Treeview.Heading", font=(SMALL_FONT[0], SMALL_FONT[1], 'bold'))
        style.configure("Treeview", font=SMALL_FONT, rowheight=25) 
        
        self.tree = ttk.Treeview(self, columns=('ID', 'Name', 'Last Action', 'Hours', 'Note'), show='headings')
        self.tree.heading('ID', text='ID')
        self.tree.heading('Name', text='Name')
        self.tree.heading('Last Action', text='Last Action')
        self.tree.heading('Hours', text='Hours Rendered')
        self.tree.heading('Note', text='8-Hour Status')
        
        self.tree.column('ID', width=80)
        self.tree.column('Name', width=200)
        self.tree.column('Last Action', width=150)
        self.tree.column('Hours', width=120)
        self.tree.column('Note', width=150)
        
        self.tree.grid(row=1, column=0, sticky='nsew', padx=10, pady=10) 
        
    def refresh_report(self):
        """Clears old data and loads the latest report from the database."""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        report_data = self.controller.db_manager.get_attendance_report()
        
        for record in report_data:
            hours = f"{record['HoursRendered'].total_seconds() / 3600:.2f} hrs" if record['HoursRendered'] else "N/A"
            action_time_str = record['LastActionTime'].strftime('%H:%M:%S %m/%d') if record['LastActionTime'] else "N/A"

            self.tree.insert('', tk.END, values=(
                record['FacultyID'],
                record['FullName'],
                f"{record['LastAction']} @ {action_time_str}" if record['LastAction'] else record['Note'],
                hours,
                record['Note']
            ))

# --- FRAME 4: ANALYTICS & GRAPHS ---
class AnalyticsFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller 
        
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        tk.Label(self, text="Attendance Analytics & Visualization", font=LARGE_FONT).grid(row=0, column=0, pady=15, sticky='ew')
        
        self.chart_frame = tk.Frame(self)
        self.chart_frame.grid(row=1, column=0, sticky='nsew', padx=10, pady=10)
        
        self.chart_frame.grid_rowconfigure(0, weight=1)
        self.chart_frame.grid_columnconfigure(0, weight=1)
        
        self.text_output = tk.Text(self.chart_frame, height=8, state=tk.DISABLED, font=SMALL_FONT)
        self.text_output.grid(row=1, column=0, sticky='ew', pady=(10, 0))

        self.fig, self.ax = plt.subplots(figsize=(6, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        
        self.canvas_widget.grid(row=0, column=0, sticky='nsew') 

    def calculate_stats(self, data):
        """Calculates average times, rates, and on-time percentage."""
        all_check_in_times = []
        all_check_out_times = []
        on_time_count = 0
        total_check_ins = 0
        
        ON_TIME_BOUNDARY = time(8, 0, 0)

        for fid, times in data.items():
            total_check_ins += len(times['check_ins'])
            
            for t_in in times['check_ins']:
                time_only = t_in.time()
                all_check_in_times.append(time_only)
                
                if time_only <= ON_TIME_BOUNDARY:
                    on_time_count += 1

            all_check_out_times.extend(times['check_outs'])
        
        avg_in_sec, avg_out_sec = None, None
        
        if all_check_in_times:
            in_seconds = [t.hour * 3600 + t.minute * 60 + t.second for t in all_check_in_times]
            avg_in_sec = np.mean(in_seconds)

        if all_check_out_times:
            out_seconds = [t.hour * 3600 + t.minute * 60 + t.second for t in all_check_out_times]
            avg_out_sec = np.mean(out_seconds)

        total_events = total_check_ins + len(all_check_out_times)
        in_rate = (total_check_ins / total_events) * 100 if total_events > 0 else 0
        out_rate = (len(all_check_out_times) / total_events) * 100 if total_events > 0 else 0
        
        on_time_percentage = (on_time_count / total_check_ins) * 100 if total_check_ins > 0 else 0

        return avg_in_sec, avg_out_sec, in_rate, out_rate, on_time_percentage
        
    def format_seconds_to_time(self, seconds):
        """Converts total seconds back to HH:MM:SS format."""
        if seconds is None: return "N/A"
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)
        return f"{h:02d}:{m:02d}:{s:02d}"

    def generate_analytics(self):
        """Fetches data, calculates stats, and plots/displays results."""
        raw_data = self.controller.db_manager.get_raw_time_data()
        
        if not raw_data:
            self.update_text_output("No attendance data found for today to generate analytics.")
            self.ax.clear()
            self.ax.text(0.5, 0.5, "No Data Today", ha='center', va='center')
            self.canvas.draw()
            return
            
        avg_in_sec, avg_out_sec, in_rate, out_rate, on_time_percentage = self.calculate_stats(raw_data)

        # 1. Plotting Average Time In and Time Out
        self.ax.clear()
        times = [t for t in [avg_in_sec, avg_out_sec] if t is not None]
        labels = [l for l, t in zip(['Avg. Check In', 'Avg. Check Out'], [avg_in_sec, avg_out_sec]) if t is not None]
        
        self.ax.bar(labels, times, color=['skyblue', 'salmon']) 
        self.ax.set_ylabel('Time (Seconds from Midnight)')
        self.ax.set_title('Average Time In/Out Today') 

        for i, val in enumerate(times):
             self.ax.text(i, val + 500, self.format_seconds_to_time(val), ha='center', va='bottom', fontsize=9)

        self.canvas.draw()
        
        # 2. Text Output (Rates and Percentage)
        in_time_formatted = self.format_seconds_to_time(avg_in_sec)
        out_time_formatted = self.format_seconds_to_time(avg_out_sec)
        
        report_text = f"""
        --- Daily Attendance Analytics ({datetime.now().strftime('%Y-%m-%d')}) ---
        Average Check-In Time:    {in_time_formatted}
        Average Check-Out Time:   {out_time_formatted}
        
        Rate of Time In Events:   {in_rate:.2f}% (Total check-ins vs. all events)
        Rate of Time Out Events:  {out_rate:.2f}% (Total check-outs vs. all events)
        
        Percentage On-Time/Early: {on_time_percentage:.2f}% (Checked in at or before 8:00 AM)
        """
        self.update_text_output(report_text)
        
    def update_text_output(self, text):
        """Helper function to update the read-only text box."""
        self.text_output.config(state=tk.NORMAL)
        self.text_output.delete('1.0', tk.END)
        self.text_output.insert(tk.END, text)
        self.text_output.config(state=tk.DISABLED)


if __name__ == "__main__":
    app = AttendanceApp()
    app.mainloop()

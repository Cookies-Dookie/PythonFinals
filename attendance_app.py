import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, time
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np
from db_manager import DatabaseManager
import subprocess, sys

def ensure_odbc_driver():
    try:
        import pyodbc
        drivers = pyodbc.drivers()
        if not any("SQL Server" in d for d in drivers):
            print("⚙️ Installing ODBC Driver 17 for SQL Server...")
            subprocess.run(["msiexec", "/i", "msodbcsql.msi", "/quiet", "/norestart"], check=True)
            print("ODBC driver installed successfully.")
    except Exception as e:
        print("ODBC install failed:", e)
        sys.exit(1)

class AttendanceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Faculty Attendance System")
        self.geometry("1200x750")
        self.configure(bg="#F8FAFC")
        self.db_manager = DatabaseManager()

        self.sidebar = tk.Frame(self, bg="#3A86FF", width=240)
        self.sidebar.pack(side="left", fill="y")

        self.container = tk.Frame(self, bg="#F8FAFC")
        self.container.pack(side="right", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (AttendanceFrame, AddFacultyFrame, ReportFrame, AnalyticsFrame):
            frame = F(parent=self.container, controller=self)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.build_sidebar()
        self.show_frame("AttendanceFrame")

    def build_sidebar(self):
        logo = tk.Label(self.sidebar, text="Faculty System", bg="#3A86FF", fg="white", font=("Segoe UI Semibold", 18))
        logo.pack(pady=(40, 30))

        buttons = [
            ("Attendance", "AttendanceFrame"),
            ("Add Faculty", "AddFacultyFrame"),
            ("Reports", "ReportFrame"),
            ("Analytics", "AnalyticsFrame")
        ]

        for text, frame_name in buttons:
            btn = tk.Label(self.sidebar, text=text, bg="#3A86FF", fg="white", font=("Segoe UI", 14),
                           padx=25, pady=15, anchor="w", cursor="hand2")
            btn.pack(fill="x", pady=5)
            btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#2563EB"))
            btn.bind("<Leave>", lambda e, b=btn: b.config(bg="#3A86FF"))
            btn.bind("<Button-1>", lambda e, f=frame_name: self.show_frame(f))

        tk.Label(self.sidebar, bg="#3A86FF").pack(expand=True, fill="both")

        exit_btn = tk.Button(self.sidebar, text="Exit", bg="#EF4444", fg="white", font=("Segoe UI", 12, "bold"),
                             relief="flat", cursor="hand2", command=self.quit)
        exit_btn.pack(pady=30, padx=20, fill="x")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        if page_name == "ReportFrame":
            frame.refresh_report()
        elif page_name == "AnalyticsFrame":
            frame.generate_analytics()
        frame.tkraise()

class BaseFrame(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#F8FAFC")
        self.controller = controller

    def build_title(self, title):
        tk.Label(self, text=title, font=("Segoe UI Semibold", 24), bg="#F8FAFC", fg="#1E293B").pack(pady=(20, 10))

    def make_card(self, parent, padding=30):
        frame = tk.Frame(parent, bg="white", bd=0, highlightbackground="#E2E8F0", highlightthickness=1)
        frame.pack(padx=padding, pady=padding, fill="x")
        return frame

class AttendanceFrame(BaseFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.build_title("Faculty Attendance")

        form = self.make_card(self)
        tk.Label(form, text="Faculty ID:", font=("Segoe UI", 16), bg="white", fg="#1E293B").grid(row=0, column=0, padx=20, pady=25, sticky="e")
        self.id_entry = ttk.Entry(form, font=("Segoe UI", 14), width=30)
        self.id_entry.grid(row=0, column=1, padx=20, pady=25)

        btns = tk.Frame(self, bg="#F8FAFC")
        btns.pack(pady=20)
        self.build_button(btns, "Check In", "#10B981", lambda: self.record("Check-In")).pack(side="left", padx=30)
        self.build_button(btns, "Check Out", "#F87171", lambda: self.record("Check-Out")).pack(side="left", padx=30)

    def build_button(self, parent, text, color, cmd):
        return tk.Button(parent, text=text, font=("Segoe UI Semibold", 14), bg=color, fg="white",
                         activebackground=color, activeforeground="white",
                         padx=30, pady=12, bd=0, relief="flat", cursor="hand2", command=cmd)

    def record(self, action):
        fid = self.id_entry.get().strip()
        if not fid:
            messagebox.showwarning("Input Error", "Please enter Faculty ID.")
            return
        f = self.controller.db_manager.load_faculty_info(fid)
        if not f:
            messagebox.showerror("Error", f"Faculty '{fid}' not found.")
            return
        sql = "INSERT INTO dbo.Attendance (FacultyID, Action, TimeStamp) VALUES (?, ?, ?)"
        ok, res = self.controller.db_manager.execute_non_query(sql, (fid, action, datetime.now()))
        if ok:
            messagebox.showinfo("Success", f"{f.full_name} successfully marked for {action}.")
        else:
            messagebox.showerror("Error", f"DB Error: {res}")
        self.id_entry.delete(0, tk.END)

class AddFacultyFrame(BaseFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.build_title("Add New Professor")

        form = self.make_card(self)
        fields = ["Faculty ID:", "Full Name:", "Department:"]
        self.entries = []

        for i, field in enumerate(fields):
            tk.Label(form, text=field, font=("Segoe UI", 14), bg="white", fg="#1E293B").grid(row=i, column=0, padx=20, pady=20, sticky="e")
            e = ttk.Entry(form, font=("Segoe UI", 13), width=40)
            e.grid(row=i, column=1, padx=20, pady=20)
            self.entries.append(e)

        self.build_button(self, "Add Faculty", "#3B82F6", self.add_faculty).pack(pady=25)

    def build_button(self, parent, text, color, cmd):
        return tk.Button(parent, text=text, font=("Segoe UI Semibold", 14), bg=color, fg="white",
                         activebackground=color, activeforeground="white",
                         padx=30, pady=12, bd=0, relief="flat", cursor="hand2", command=cmd)

    def add_faculty(self):
        fid, name, dept = [e.get().strip() for e in self.entries]
        if not all([fid, name, dept]):
            messagebox.showwarning("Input Error", "All fields are required.")
            return
        ok, res = self.controller.db_manager.add_faculty(fid, name, dept)
        if ok:
            messagebox.showinfo("Success", f"{name} has been added successfully.")
            for e in self.entries: e.delete(0, tk.END)
        else:
            messagebox.showerror("Error", f"Failed: {res}")

class ReportFrame(BaseFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.build_title("Attendance Reports")

        card = self.make_card(self)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 12), background="#3B82F6", foreground="white")
        style.configure("Treeview", font=("Segoe UI", 11), rowheight=30, background="white", fieldbackground="white", foreground="#1E293B")

        self.tree = ttk.Treeview(card, columns=("ID", "Name", "Action", "Hours", "Note"), show="headings")
        for c in ("ID", "Name", "Action", "Hours", "Note"):
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor="center", width=170)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

    def refresh_report(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        data = self.controller.db_manager.get_attendance_report()
        for r in data:
            hrs = f"{r['HoursRendered'].total_seconds()/3600:.2f} hrs" if r['HoursRendered'] else "N/A"
            t = r['LastActionTime'].strftime('%H:%M:%S %m/%d') if r['LastActionTime'] else "N/A"
            self.tree.insert('', 'end', values=(r['FacultyID'], r['FullName'], f"{r['LastAction']} at {t}", hrs, r['Note']))

class AnalyticsFrame(BaseFrame):
    def __init__(self, parent, controller):
        super().__init__(parent, controller)
        self.build_title("Attendance Analytics")

        card = self.make_card(self)
        self.fig, self.ax = plt.subplots(figsize=(7, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=card)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)

        self.output = tk.Text(card, height=8, font=("Segoe UI", 11), bg="white", fg="#1E293B", relief="flat")
        self.output.pack(fill="x", pady=10)
        self.output.config(state=tk.DISABLED)

    def generate_analytics(self):
        data = self.controller.db_manager.get_raw_time_data()
        self.ax.clear()
        if not data:
            self.ax.text(0.5, 0.5, "No Data", ha="center", va="center", color="#1E293B")
            self.canvas.draw()
            return
        avg_in, avg_out, in_rate, out_rate, ontime = self.calculate(data)
        self.ax.bar(["Check-In", "Check-Out"], [avg_in, avg_out], color=["#3B82F6", "#EF4444"])
        for i, v in enumerate([avg_in, avg_out]):
            self.ax.text(i, v + 300, self.format(v), ha="center", color="#1E293B")
        self.ax.set_title("Average Attendance Times", color="#1E293B")
        self.canvas.draw()
        text = f"""
Daily Analytics ({datetime.now().strftime('%Y-%m-%d')})
Avg Check-In:  {self.format(avg_in)}
Avg Check-Out: {self.format(avg_out)}
Check-In Rate: {in_rate:.2f}%
Check-Out Rate:{out_rate:.2f}%
On-Time Rate:  {ontime:.2f}%
"""
        self.update_text(text)

    def calculate(self, data):
        all_in, all_out, ontime, total_in = [], [], 0, 0
        ON_TIME = time(8, 0, 0)
        for fid, t in data.items():
            total_in += len(t["check_ins"])
            for tin in t["check_ins"]:
                all_in.append(tin.time())
                if tin.time() <= ON_TIME: ontime += 1
            all_out.extend([tout.time() for tout in t["check_outs"]])
        avg_in = np.mean([t.hour * 3600 + t.minute * 60 for t in all_in]) if all_in else 0
        avg_out = np.mean([t.hour * 3600 + t.minute * 60 for t in all_out]) if all_out else 0
        total = total_in + len(all_out)
        in_rate = (total_in / total) * 100 if total else 0
        out_rate = (len(all_out) / total) * 100 if total else 0
        on_time_pct = (ontime / total_in) * 100 if total_in else 0
        return avg_in, avg_out, in_rate, out_rate, on_time_pct

    def format(self, sec):
        h, m = int(sec // 3600), int((sec % 3600) // 60)
        return f"{h:02d}:{m:02d}"

    def update_text(self, text):
        self.output.config(state=tk.NORMAL)
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, text)
        self.output.config(state=tk.DISABLED)

if __name__ == "__main__":
    ensure_odbc_driver()
    app = AttendanceApp()
    app.mainloop()

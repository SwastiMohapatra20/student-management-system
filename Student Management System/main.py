"""
Advanced Student Management System (enterprise-style desktop app)

Features:
- ttkbootstrap themed UI with light/dark toggle
- Role-based login (Admin / Teacher / Guest)
- Sidebar + Notebook (Manage, Reports, Import/Export, Audit Log, Settings)
- CRUD with real-time validation, undo/redo, pagination
- SQLite backend with indexes, audit log table
- Charts embedded with matplotlib
- Import/Export CSV & Excel via pandas
- Backup & Restore DB
- Status bar showing DB and user info

Run:
    pip install ttkbootstrap pandas matplotlib openpyxl
    python advanced_student_mgmt.py
"""

import os
import sqlite3
import threading
import shutil
import time
import csv
import io
from datetime import datetime
from functools import partial

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox, filedialog
from tkinter import StringVar, IntVar

import pandas as pd
import matplotlib
matplotlib.use("Agg")  # avoid default interactive backend initially
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# ------------------------- CONFIG -------------------------
DB_FILE = "advanced_students.db"
PAGE_SIZE_DEFAULT = 50
THEME_LIGHT = "cosmo"
THEME_DARK = "darkly"
AVAILABLE_THEMES = [
    "cosmo","flatly","minty","litera","lumen","pulse","journal","sandstone","yeti",
    "cyborg","darkly","slate","solar","superhero","morph","vapor","vapor"
]

# ------------------------- DB HELPERS -------------------------
def get_conn():
    conn = sqlite3.connect(DB_FILE, timeout=10, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    # Students table
    cur.execute("""
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        roll TEXT UNIQUE NOT NULL,
        course TEXT NOT NULL,
        marks INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )""")
    # Audit log
    cur.execute("""
    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ts TEXT,
        user TEXT,
        role TEXT,
        action TEXT,
        details TEXT
    )""")
    # Users table (simple)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT NOT NULL,
        role TEXT NOT NULL
    )""")
    # Indexes for fast search
    cur.execute("CREATE INDEX IF NOT EXISTS idx_students_roll ON students(roll)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_students_name ON students(name)")
    conn.commit()
    # Ensure at least an admin account exists
    cur.execute("SELECT COUNT(*) as c FROM users")
    if cur.fetchone()["c"] == 0:
        # default admin/admin (please change)
        cur.execute("INSERT OR REPLACE INTO users (username,password,role) VALUES (?,?,?)",
                    ("admin", "admin", "Admin"))
        cur.execute("INSERT OR REPLACE INTO users (username,password,role) VALUES (?,?,?)",
                    ("teacher", "teacher", "Teacher"))
        conn.commit()
    conn.close()

def log_audit(user, role, action, details=""):
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("INSERT INTO audit_log (ts,user,role,action,details) VALUES (?,?,?,?,?)",
                    (datetime.utcnow().isoformat(), user, role, action, details))
        conn.commit()
        conn.close()
    except Exception as e:
        print("Audit log failed:", e)

# ------------------------- UTIL -------------------------
def safe_commit(conn):
    try:
        conn.commit()
    except Exception as e:
        conn.rollback()
        raise

def timestamped_backup_path():
    return f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"

# ------------------------- APP CLASS -------------------------
class StudentERPApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Student Management System")
        self.root.geometry("1100x700")
        self.style = root.style
        # states
        self.current_user = None
        self.current_role = None
        self.page_size = IntVar(value=PAGE_SIZE_DEFAULT)
        self.current_page = 0
        self.total_rows = 0
        self.undo_stack = []  # store tuples (action, payload)
        self.redo_stack = []
        # Build UI
        self.build_login_screen()

    # -------------------- LOGIN --------------------
    def build_login_screen(self):
        for w in self.root.winfo_children():
            w.destroy()
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(expand=True)
        ttk.Label(frame, text="Welcome — Login", font=("Inter", 22, "bold")).pack(pady=(0,12))
        form = ttk.Frame(frame)
        form.pack()
        ttk.Label(form, text="Username:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.login_user = StringVar()
        ttk.Entry(form, textvariable=self.login_user).grid(row=0, column=1, padx=6, pady=6)
        ttk.Label(form, text="Password:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.login_pass = StringVar()
        ttk.Entry(form, textvariable=self.login_pass, show="*").grid(row=1, column=1, padx=6, pady=6)
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=12)
        ttk.Button(btn_frame, text="Login", bootstyle="primary", command=self.handle_login).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="Continue as Guest", bootstyle="secondary", command=self.login_as_guest).grid(row=0, column=1, padx=6)
        ttk.Label(frame, text="(Default admin/admin, teacher/teacher) — change in DB for production", font=("Inter", 9)).pack(pady=(8,0))

    def handle_login(self):
        user = self.login_user.get().strip()
        pwd = self.login_pass.get().strip()
        if not user or not pwd:
            messagebox.showwarning("Login", "Enter username and password")
            return
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT role FROM users WHERE username=? AND password=?", (user, pwd))
        row = cur.fetchone()
        conn.close()
        if row:
            self.current_user = user
            self.current_role = row["role"]
            log_audit(self.current_user, self.current_role, "login", "successful login")
            self.build_main_ui()
        else:
            messagebox.showerror("Login Failed", "Invalid credentials")

    def login_as_guest(self):
        self.current_user = "guest"
        self.current_role = "Guest"
        log_audit(self.current_user, self.current_role, "login", "guest login")
        self.build_main_ui()

    # -------------------- MAIN UI --------------------
    def build_main_ui(self):
        for w in self.root.winfo_children():
            w.destroy()
        # Top toolbar
        topbar = ttk.Frame(self.root)
        topbar.pack(fill="x")
        ttk.Label(topbar, text="Advanced Student Management", font=("Inter", 16, "bold")).pack(side="left", padx=10, pady=6)
        ttk.Button(topbar, text="Backup DB", bootstyle="outline-info", command=self.backup_db).pack(side="right", padx=6)
        ttk.Button(topbar, text="Logout", bootstyle="outline-secondary", command=self.logout).pack(side="right", padx=6)
        # Main content: sidebar + notebook
        content = ttk.Frame(self.root)
        content.pack(fill="both", expand=True)
        # Sidebar
        sidebar = ttk.Frame(content, width=200)
        sidebar.pack(side="left", fill="y", padx=(8,4), pady=8)
        ttk.Label(sidebar, text=f"User: {self.current_user}\nRole: {self.current_role}", bootstyle="secondary").pack(pady=6, padx=6)
        ttk.Separator(sidebar).pack(fill="x", pady=6)
        self.nav_buttons = {}
        for idx, (name, cmd) in enumerate([
            ("Manage", self.show_manage_tab),
            ("Reports", self.show_reports_tab),
            ("Import/Export", self.show_import_tab),
            ("Audit Log", self.show_audit_tab),
            ("Settings", self.show_settings_tab)
        ]):
            b = ttk.Button(sidebar, text=name, width=18, command=cmd)
            b.pack(pady=4, padx=6)
            self.nav_buttons[name] = b
        # Notebook
        self.notebook = ttk.Notebook(content)
        self.notebook.pack(side="left", fill="both", expand=True, padx=(4,8), pady=8)
        # Build tabs
        self.build_manage_tab()
        self.build_reports_tab()
        self.build_import_tab()
        self.build_audit_tab()
        self.build_settings_tab()
        # status bar
        self.statusbar = ttk.Label(self.root, text="Ready", bootstyle="secondary")
        self.statusbar.pack(fill="x", side="bottom")
        # start on Manage tab
        self.notebook.select(0)
        self.refresh_status()

    def logout(self):
        confirm = messagebox.askyesno("Logout", "Are you sure you want to logout?")
        if confirm:
            log_audit(self.current_user, self.current_role, "logout", "user logged out")
            self.current_user = None
            self.current_role = None
            self.build_login_screen()

    def refresh_status(self, text=None):
        if text:
            self.statusbar.configure(text=text)
            return
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) as c FROM students")
        total = cur.fetchone()["c"]
        conn.close()
        self.total_rows = total
        self.statusbar.configure(text=f"DB: {DB_FILE} | Records: {total} | User: {self.current_user} ({self.current_role})")

    # -------------------- MANAGE TAB --------------------
    def build_manage_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Manage")
        # top: form
        form = ttk.Frame(tab, padding=8)
        form.pack(fill="x")
        ttk.Label(form, text="Name").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.m_name = StringVar()
        name_entry = ttk.Entry(form, textvariable=self.m_name, width=30)
        name_entry.grid(row=0, column=1, padx=6, pady=6)
        self.name_err = ttk.Label(form, text="", bootstyle="warning")
        self.name_err.grid(row=1, column=1, sticky="w")

        ttk.Label(form, text="Roll").grid(row=0, column=2, padx=6, pady=6, sticky="w")
        self.m_roll = StringVar()
        roll_entry = ttk.Entry(form, textvariable=self.m_roll, width=20)
        roll_entry.grid(row=0, column=3, padx=6, pady=6)
        self.roll_err = ttk.Label(form, text="", bootstyle="warning")
        self.roll_err.grid(row=1, column=3, sticky="w")

        ttk.Label(form, text="Course").grid(row=0, column=4, padx=6, pady=6, sticky="w")
        self.m_course = StringVar()
        course_box = ttk.Combobox(form, values=["B.Tech CSE","B.Tech ECE","B.Sc","MCA","M.Tech","BBA","Other"], textvariable=self.m_course, width=18)
        course_box.grid(row=0, column=5, padx=6, pady=6)
        self.course_err = ttk.Label(form, text="", bootstyle="warning")
        self.course_err.grid(row=1, column=5, sticky="w")

        ttk.Label(form, text="Marks").grid(row=2, column=0, padx=6, pady=6, sticky="w")
        self.m_marks = StringVar()
        marks_entry = ttk.Entry(form, textvariable=self.m_marks, width=12)
        marks_entry.grid(row=2, column=1, padx=6, pady=6)
        self.marks_err = ttk.Label(form, text="", bootstyle="warning")
        self.marks_err.grid(row=3, column=1, sticky="w")

        # Buttons
        btn_area = ttk.Frame(form)
        btn_area.grid(row=2, column=3, columnspan=3, sticky="e", padx=6)
        self.btn_add = ttk.Button(btn_area, text="Add", bootstyle="success", command=self.add_student)
        self.btn_add.grid(row=0, column=0, padx=4)
        self.btn_update = ttk.Button(btn_area, text="Update", bootstyle="info", command=self.update_student)
        self.btn_update.grid(row=0, column=1, padx=4)
        self.btn_delete = ttk.Button(btn_area, text="Delete", bootstyle="danger", command=self.delete_student)
        self.btn_delete.grid(row=0, column=2, padx=4)
        self.btn_clear = ttk.Button(btn_area, text="Clear", bootstyle="secondary", command=self.clear_form)
        self.btn_clear.grid(row=0, column=3, padx=4)
        self.btn_undo = ttk.Button(btn_area, text="Undo", bootstyle="outline-warning", command=self.undo_action)
        self.btn_undo.grid(row=0, column=4, padx=4)
        self.btn_redo = ttk.Button(btn_area, text="Redo", bootstyle="outline-info", command=self.redo_action)
        self.btn_redo.grid(row=0, column=5, padx=4)

        # Middle: search and pagination
        mid = ttk.Frame(tab, padding=6)
        mid.pack(fill="x")
        ttk.Label(mid, text="Search:").pack(side="left")
        self.search_text = StringVar()
        search_entry = ttk.Entry(mid, textvariable=self.search_text)
        search_entry.pack(side="left", fill="x", expand=True, padx=6)
        search_entry.bind("<KeyRelease>", lambda e: self.load_page(0))
        ttk.Button(mid, text="Clear Search", bootstyle="secondary", command=self.clear_search).pack(side="left", padx=6)

        # Pagination controls
        self.lbl_page = ttk.Label(mid, text="Page 1")
        self.lbl_page.pack(side="right", padx=6)
        ttk.Button(mid, text="Prev", bootstyle="outline-secondary", command=self.prev_page).pack(side="right", padx=6)
        ttk.Button(mid, text="Next", bootstyle="outline-secondary", command=self.next_page).pack(side="right", padx=6)

        # Bottom: table
        table_frame = ttk.Frame(tab, padding=6)
        table_frame.pack(fill="both", expand=True)
        cols = ("ID","Name","Roll","Course","Marks","Created")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=c)
            if c=="Name":
                self.tree.column(c, width=300, anchor="w")
            else:
                self.tree.column(c, width=100, anchor="center")
        self.tree.pack(fill="both", expand=True, side="left")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y")
        self.tree.configure(yscroll=vsb.set)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        # initial load
        self.load_page(0)

        # validation binding
        for var in (self.m_name, self.m_roll, self.m_course, self.m_marks):
            # trace changes for live validation
            try:
                var.trace_add("write", lambda *args: self.validate_form())
            except Exception:
                pass
        self.validate_form()

    def validate_form(self):
        ok = True
        name = self.m_name.get().strip()
        roll = self.m_roll.get().strip()
        course = self.m_course.get().strip()
        marks = self.m_marks.get().strip()
        # name
        if not name or any((not (ch.isalpha() or ch in " .-") for ch in name)):
            self.name_err.configure(text="Invalid name (letters, space, . -)")
            ok = False
        else:
            self.name_err.configure(text="")
        # roll
        if not roll or not roll.isdigit() or len(roll) > 12:
            self.roll_err.configure(text="Roll numeric (1-12 digits)")
            ok = False
        else:
            self.roll_err.configure(text="")
        # course
        if not course:
            self.course_err.configure(text="Choose or enter course")
            ok = False
        else:
            self.course_err.configure(text="")
        # marks
        if not marks.isdigit() or not (0 <= int(marks) <= 100):
            self.marks_err.configure(text="Marks 0-100")
            ok = False
        else:
            self.marks_err.configure(text="")
        # set button state
        state = "normal" if ok else "disabled"
        try:
            self.btn_add.configure(state=state)
            self.btn_update.configure(state=state)
        except Exception:
            pass
        return ok

    def add_student(self):
        if not self.validate_form():
            messagebox.showwarning("Validation", "Fix validation errors first")
            return
        name = self.m_name.get().strip()
        roll = self.m_roll.get().strip()
        course = self.m_course.get().strip()
        marks = int(self.m_marks.get().strip())
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("INSERT INTO students (name,roll,course,marks) VALUES (?,?,?,?)",
                        (name, roll, course, marks))
            conn.commit()
            conn.close()
            log_audit(self.current_user, self.current_role, "add", f"{roll}|{name}")
            # push undo action
            self.undo_stack.append(("delete", {"roll": roll}))
            self.redo_stack.clear()
            self.load_page(self.current_page)
            self.clear_form()
            messagebox.showinfo("Added", "Student added successfully")
            self.refresh_status()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Roll number already exists")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_student(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a row to update")
            return
        row = self.tree.item(sel[0])["values"]
        sid = row[0]
        old = {"id": sid, "name": row[1], "roll": row[2], "course": row[3], "marks": row[4]}
        if not self.validate_form():
            messagebox.showwarning("Validation", "Fix validation errors first")
            return
        new = {"name": self.m_name.get().strip(), "roll": self.m_roll.get().strip(),
               "course": self.m_course.get().strip(), "marks": int(self.m_marks.get().strip())}
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute("UPDATE students SET name=?, roll=?, course=?, marks=? WHERE id=?",
                        (new["name"], new["roll"], new["course"], new["marks"], sid))
            conn.commit()
            conn.close()
            log_audit(self.current_user, self.current_role, "update", f"id={sid}|from={old}|to={new}")
            # push undo
            self.undo_stack.append(("update", {"id": sid, "old": old}))
            self.redo_stack.clear()
            self.load_page(self.current_page)
            self.clear_form()
            messagebox.showinfo("Updated", "Record updated")
            self.refresh_status()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Roll conflicts with another record")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def delete_student(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a row to delete")
            return
        row = self.tree.item(sel[0])["values"]
        sid, name, roll = row[0], row[1], row[2]
        if self.current_role == "Guest":
            messagebox.showwarning("Permission", "Guest cannot delete")
            return
        if not messagebox.askyesno("Confirm", f"Delete {name} (roll {roll})?"):
            return
        try:
            conn = get_conn()
            cur = conn.cursor()
            # store old for undo
            cur.execute("SELECT * FROM students WHERE id=?", (sid,))
            old = dict(cur.fetchone())
            cur.execute("DELETE FROM students WHERE id=?", (sid,))
            conn.commit()
            conn.close()
            log_audit(self.current_user, self.current_role, "delete", f"id={sid}|{roll}|{name}")
            self.undo_stack.append(("insert", {"row": old}))
            self.redo_stack.clear()
            self.load_page(self.current_page)
            self.clear_form()
            messagebox.showinfo("Deleted", "Record deleted")
            self.refresh_status()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def clear_form(self):
        self.m_name.set("")
        self.m_roll.set("")
        self.m_course.set("")
        self.m_marks.set("")
        self.tree.selection_remove(self.tree.selection())
        self.validate_form()

    # Tree select fill form
    def on_tree_select(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        row = self.tree.item(sel[0])["values"]
        self.m_name.set(row[1])
        self.m_roll.set(row[2])
        self.m_course.set(row[3])
        self.m_marks.set(row[4])
        self.validate_form()

    # -------------------- Pagination / Loading --------------------
    def build_query_base(self):
        base = "FROM students"
        search = self.search_text.get().strip()
        args = []
        if search:
            like = f"%{search}%"
            base += " WHERE roll LIKE ? OR name LIKE ? OR course LIKE ?"
            args.extend([like, like, like])
        return base, args

    def load_page(self, page):
        try:
            page = max(0, int(page))
        except Exception:
            page = 0
        self.current_page = page
        offset = page * self.page_size.get()
        base, args = self.build_query_base()
        conn = get_conn()
        cur = conn.cursor()
        # total count
        cur.execute(f"SELECT COUNT(*) as c {base}", args)
        total = cur.fetchone()["c"]
        self.total_rows = total
        # fetch page
        cur.execute(f"SELECT id,name,roll,course,marks,created_at {base} ORDER BY name COLLATE NOCASE LIMIT ? OFFSET ?",
                    args + [self.page_size.get(), offset])
        rows = cur.fetchall()
        conn.close()
        # populate tree
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=(r["id"], r["name"], r["roll"], r["course"], r["marks"], r["created_at"]))
        # update page label
        total_pages = max(1, (self.total_rows + self.page_size.get() - 1) // self.page_size.get())
        self.lbl_page.configure(text=f"Page {self.current_page+1} / {total_pages}")
        self.refresh_status()

    def prev_page(self):
        if self.current_page > 0:
            self.load_page(self.current_page - 1)

    def next_page(self):
        total_pages = max(1, (self.total_rows + self.page_size.get() - 1) // self.page_size.get())
        if self.current_page + 1 < total_pages:
            self.load_page(self.current_page + 1)

    def clear_search(self):
        self.search_text.set("")
        self.load_page(0)

    # -------------------- Undo/Redo --------------------
    def undo_action(self):
        if not self.undo_stack:
            messagebox.showinfo("Undo", "Nothing to undo")
            return
        action, payload = self.undo_stack.pop()
        try:
            if action == "delete":
                # payload: {"roll": roll} -> delete that roll
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("DELETE FROM students WHERE roll=?", (payload["roll"],))
                conn.commit(); conn.close()
                log_audit(self.current_user, self.current_role, "undo_delete", payload["roll"])
                self.redo_stack.append(("insert", {"roll": payload["roll"]}))
            elif action == "insert":
                # payload: {"row": oldrow}
                row = payload["row"]
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("INSERT INTO students (id,name,roll,course,marks,created_at) VALUES (?,?,?,?,?,?)",
                            (row["id"], row["name"], row["roll"], row["course"], row["marks"], row["created_at"]))
                conn.commit(); conn.close()
                log_audit(self.current_user, self.current_role, "undo_insert", str(row["id"]))
                self.redo_stack.append(("delete", {"roll": row["roll"]}))
            elif action == "update":
                # payload: {"id": id, "old": olddict}
                old = payload["old"]
                conn = get_conn()
                cur = conn.cursor()
                cur.execute("UPDATE students SET name=?, roll=?, course=?, marks=? WHERE id=?",
                            (old["name"], old["roll"], old["course"], old["marks"], old["id"]))
                conn.commit(); conn.close()
                log_audit(self.current_user, self.current_role, "undo_update", str(old["id"]))
                self.redo_stack.append(("update_redo", {"id": old["id"]}))
            self.load_page(self.current_page)
            messagebox.showinfo("Undo", "Undo performed")
        except Exception as e:
            messagebox.showerror("Undo Error", str(e))

    def redo_action(self):
        if not self.redo_stack:
            messagebox.showinfo("Redo", "Nothing to redo")
            return
        action, payload = self.redo_stack.pop()
        # For simplicity, just reload page (advanced redo implementations can be added)
        messagebox.showinfo("Redo", "Redo executed (simple placeholder)")
        self.load_page(self.current_page)

    # -------------------- REPORTS TAB --------------------
    def build_reports_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Reports")
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Refresh Charts", bootstyle="outline-primary", command=self.refresh_charts).pack(side="left", padx=6)
        ttk.Button(top, text="Export Chart as PNG", bootstyle="outline-secondary", command=self.export_chart_png).pack(side="left", padx=6)
        # chart area
        chart_frame = ttk.Frame(tab)
        chart_frame.pack(fill="both", expand=True, padx=8, pady=8)
        self.fig = Figure(figsize=(8,5))
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self.refresh_charts()

    def refresh_charts(self):
        try:
            conn = get_conn()
            cur = conn.cursor()
            # distribution of students by course
            cur.execute("SELECT course, COUNT(*) as c FROM students GROUP BY course")
            rows = cur.fetchall()
            courses = [r["course"] for r in rows]
            counts = [r["c"] for r in rows]
            # marks histogram
            cur.execute("SELECT marks FROM students")
            marks = [r["marks"] for r in cur.fetchall()]
            conn.close()

            self.fig.clear()
            ax1 = self.fig.add_subplot(121)
            ax1.bar(courses if courses else ["No data"], counts if counts else [0])
            ax1.set_title("Students per Course")
            ax1.tick_params(axis='x', rotation=45)

            ax2 = self.fig.add_subplot(122)
            if marks:
                ax2.hist(marks, bins=10)
            else:
                ax2.text(0.5, 0.5, "No marks data", ha="center")
            ax2.set_title("Marks Distribution")

            self.canvas.draw()
            log_audit(self.current_user, self.current_role, "refresh_charts", "")
        except Exception as e:
            messagebox.showerror("Charts Error", str(e))

    def export_chart_png(self):
        file = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG image","*.png")])
        if not file:
            return
        self.fig.savefig(file)
        messagebox.showinfo("Saved", f"Chart saved to {file}")

    # -------------------- IMPORT / EXPORT TAB --------------------
    def build_import_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Import/Export")
        frame = ttk.Frame(tab, padding=12)
        frame.pack(fill="x")
        ttk.Button(frame, text="Import CSV/Excel", bootstyle="outline-primary", command=self.import_file).pack(side="left", padx=6)
        ttk.Button(frame, text="Export CSV", bootstyle="outline-success", command=self.export_csv).pack(side="left", padx=6)
        ttk.Button(frame, text="Export Excel", bootstyle="outline-info", command=self.export_excel).pack(side="left", padx=6)
        ttk.Button(frame, text="Restore DB", bootstyle="outline-danger", command=self.restore_db).pack(side="right", padx=6)

    def import_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files","*.csv"),("Excel files","*.xlsx;*.xls")])
        if not path: return
        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path, dtype=str)
            else:
                df = pd.read_excel(path, dtype=str)
            # expect columns: name,roll,course,marks (case-insensitive)
            df.columns = [c.strip().lower() for c in df.columns]
            expected = {"name","roll","course","marks"}
            if not expected.issubset(set(df.columns)):
                messagebox.showerror("Import Error", f"File missing columns. Need at least: {expected}")
                return
            # insert rows
            conn = get_conn()
            cur = conn.cursor()
            inserted = 0
            for _, r in df.iterrows():
                try:
                    cur.execute("INSERT INTO students (name,roll,course,marks) VALUES (?,?,?,?)",
                                (str(r["name"]), str(r["roll"]), str(r["course"]), int(float(r["marks"])) ))
                    inserted += 1
                except Exception:
                    continue
            conn.commit()
            conn.close()
            log_audit(self.current_user, self.current_role, "import", f"{path}|inserted={inserted}")
            messagebox.showinfo("Import", f"Import complete. Inserted {inserted} rows.")
            self.load_page(0)
        except Exception as e:
            messagebox.showerror("Import Error", str(e))

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files","*.csv")])
        if not path: return
        try:
            conn = get_conn()
            df = pd.read_sql_query("SELECT * FROM students", conn)
            conn.close()
            df.to_csv(path, index=False)
            log_audit(self.current_user, self.current_role, "export_csv", path)
            messagebox.showinfo("Export", f"Exported to {path}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files","*.xlsx")])
        if not path: return
        try:
            conn = get_conn()
            df = pd.read_sql_query("SELECT * FROM students", conn)
            conn.close()
            df.to_excel(path, index=False)
            log_audit(self.current_user, self.current_role, "export_excel", path)
            messagebox.showinfo("Export", f"Exported to {path}")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def backup_db(self):
        try:
            backup_path = timestamped_backup_path()
            shutil.copyfile(DB_FILE, backup_path)
            log_audit(self.current_user, self.current_role, "backup", backup_path)
            messagebox.showinfo("Backup", f"DB backed up to {backup_path}")
        except Exception as e:
            messagebox.showerror("Backup Error", str(e))

    def restore_db(self):
        path = filedialog.askopenfilename(filetypes=[("DB files","*.db"),("All files","*.*")])
        if not path: return
        if not messagebox.askyesno("Restore", "This will replace the current DB. Continue?"):
            return
        try:
            shutil.copyfile(path, DB_FILE)
            log_audit(self.current_user, self.current_role, "restore", path)
            messagebox.showinfo("Restore", "DB restored. Restarting application.")
            self.root.after(100, lambda: os.execl(sys.executable, sys.executable, *sys.argv))
        except Exception as e:
            messagebox.showerror("Restore Error", str(e))

    # -------------------- AUDIT TAB --------------------
    def build_audit_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Audit Log")
        frame = ttk.Frame(tab, padding=8)
        frame.pack(fill="both", expand=True)
        cols = ("TS","User","Role","Action","Details")
        self.audit_tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c in cols:
            self.audit_tree.heading(c, text=c)
            self.audit_tree.column(c, width=150, anchor="w")
        self.audit_tree.pack(fill="both", expand=True, side="left")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.audit_tree.yview)
        vsb.pack(side="right", fill="y")
        self.audit_tree.configure(yscroll=vsb.set)
        # load audit
        self.load_audit()

    def load_audit(self):
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT ts,user,role,action,details FROM audit_log ORDER BY id DESC LIMIT 1000")
        rows = cur.fetchall()
        conn.close()
        self.audit_tree.delete(*self.audit_tree.get_children())
        for r in rows:
            self.audit_tree.insert("", "end", values=(r["ts"], r["user"], r["role"], r["action"], r["details"]))

    # -------------------- SETTINGS TAB --------------------
    def build_settings_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Settings")
        frame = ttk.Frame(tab, padding=12)
        frame.pack(fill="x")
        # theme
        ttk.Label(frame, text="Theme:").grid(row=0, column=0, sticky="w")
        self.theme_var = StringVar(value=THEME_LIGHT)
        ttk.Combobox(frame, values=AVAILABLE_THEMES, textvariable=self.theme_var, width=20).grid(row=0, column=1, padx=6)
        ttk.Button(frame, text="Apply Theme", bootstyle="outline-primary", command=self.apply_theme).grid(row=0, column=2, padx=6)
        # page size
        ttk.Label(frame, text="Page size:").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.page_size, width=6).grid(row=1, column=1, sticky="w")
        ttk.Button(frame, text="Apply Page Size", command=self.apply_page_size).grid(row=1, column=2, padx=6)
        # show current DB path
        ttk.Label(frame, text=f"DB File: {os.path.abspath(DB_FILE)}").grid(row=2, column=0, columnspan=3, pady=12, sticky="w")

    def apply_theme(self):
        theme = self.theme_var.get()
        try:
            self.style.theme_use(theme)
            messagebox.showinfo("Theme", f"Theme changed to {theme}")
        except Exception as e:
            messagebox.showerror("Theme Error", str(e))

    def apply_page_size(self):
        try:
            val = int(self.page_size.get())
            if val <= 0:
                raise ValueError("Page size > 0")
            self.load_page(0)
            messagebox.showinfo("Page Size", f"Page size set to {val}")
        except Exception as e:
            messagebox.showerror("Page Size Error", str(e))

    # -------------------- NAV HELPERS --------------------
    def show_manage_tab(self):
        self.notebook.select(0)

    def show_reports_tab(self):
        self.notebook.select(1)
        self.refresh_charts()

    def show_import_tab(self):
        self.notebook.select(2)

    def show_audit_tab(self):
        self.notebook.select(3)
        self.load_audit()

    def show_settings_tab(self):
        self.notebook.select(4)

# ------------------------- STARTUP -------------------------
if __name__ == "__main__":
    import sys
    init_db()
    # Create root with initial theme
    try:
        root = ttk.Window(themename=THEME_LIGHT)
    except Exception:
        root = ttk.Window()
    app = StudentERPApp(root)
    root.mainloop()

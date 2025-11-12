# main.py
import os
import sys
import threading
import time
import configparser
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import pyodbc
import keyboard   # requires admin for some operations on Windows
import pandas as pd

# ---------- Config ----------
CFG = configparser.ConfigParser()
CFG.read('config.ini')
DB_PATH = CFG.get('Database', 'DB_PATH')
TIMEZONE = CFG.get('App', 'TIMEZONE', fallback='Asia/Kolkata')
EXPORT_FOLDER = CFG.get('App', 'EXPORT_FOLDER', fallback=os.path.join(os.getcwd(), 'reports'))
if not os.path.exists(EXPORT_FOLDER):
    os.makedirs(EXPORT_FOLDER, exist_ok=True)

# ---------- DB Helpers ----------
def get_conn():
    # Use Access ODBC driver (Windows). Must exist on client machines by default with Access installed.
    conn_str = r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;" % DB_PATH
    return pyodbc.connect(conn_str, autocommit=True, timeout=5)

def check_password(password):
    # returns full name or None
    try:
        cn = get_conn()
        cur = cn.cursor()
        cur.execute("SELECT ID, FirstName, LastName FROM Users WHERE [Password]=?", (password,))
        row = cur.fetchone()
        cur.close()
        cn.close()
        if row:
            return f"{row.FirstName} {row.LastName}", row.ID
    except Exception as e:
        print("DB error check_password:", e)
    return None, None

def ensure_today_record(fullname):
    # Find record for current date and user; if not exists create and return ID
    today_str = datetime.now().strftime("%d-%m-%Y")
    try:
        cn = get_conn()
        cur = cn.cursor()
        # assume Logs has Date column as text (dd-mm-yyyy) in column 'LogDate'
        # Adjust SQL to your actual column names as needed
        cur.execute("SELECT ID FROM Attendance WHERE FullName=? AND DateValue(LoginTime)=Date()", (fullname,))
        row = cur.fetchone()
        if row:
            rec_id = row.ID
        else:
            cur.execute("INSERT INTO Attendance (FullName, LoginTime) VALUES (?, ?)", (fullname, datetime.now()))
            # get last inserted ID: Access doesn't support SELECT SCOPE_IDENTITY; use SELECT MAX(ID) trick
            cur.execute("SELECT MAX(ID) as ID FROM Attendance")
            rec_id = cur.fetchone().ID
        cur.close()
        cn.close()
        return rec_id
    except Exception as e:
        print("DB error ensure_today_record:", e)
        return None

def update_field(rec_id, field_name, dt):
    try:
        cn = get_conn()
        cur = cn.cursor()
        cur.execute(f"UPDATE Attendance SET [{field_name}] = ? WHERE ID = ?", (dt, rec_id))
        cur.close()
        cn.close()
    except Exception as e:
        print("DB error update_field:", e)

def calc_totals(rec_id):
    try:
        cn = get_conn()
        cur = cn.cursor()
        cur.execute("SELECT DayStart, DayEnd, MorningTeaBreakStart, MorningTeaBreakEnd, LunchBreakStart, LunchBreakEnd, AfternoonTeaBreakStart, AfternoonTeaBreakEnd FROM Attendance WHERE ID=?", (rec_id,))
        row = cur.fetchone()
        if not row:
            cur.close()
            cn.close()
            return
        ds, de, mts, mte, ls, le, ats, ate = row
        def mins(a, b):
            if a and b:
                return int((b - a).total_seconds() / 60)
            return None
        t_morning = mins(mts, mte)
        t_lunch = mins(ls, le)
        t_afternoon = mins(ats, ate)
        total_work_mins = None
        if ds and de:
            total = int((de - ds).total_seconds() / 60)
            sub = 0
            for v in (t_morning, t_lunch, t_afternoon):
                if v:
                    sub += v
            total_work_mins = total - sub
        # Update totals (store minutes or hours as you prefer)
        cur.execute("UPDATE Attendance SET TotalMorningTeaTime=?, TotalLunchTime=?, TotalAfternoonTeaTime=?, TotalHoursWorked=? WHERE ID=?",
                    (t_morning, t_lunch, t_afternoon, (round(total_work_mins/60,2) if total_work_mins is not None else None), rec_id))
        cur.close()
        cn.close()
    except Exception as e:
        print("DB error calc_totals:", e)

# ---------- Keyboard blocking ----------
blocked = False
blocked_keys = ['alt', 'tab', 'windows', 'win', 'left windows', 'right windows', 'ctrl+alt+tab']
# Note: keyboard module key names vary. We'll block common keys. Ctrl+Alt+Del cannot be blocked.

def block_keyboard():
    global blocked
    try:
        # block many keys
        for k in blocked_keys:
            try:
                keyboard.block_key(k)
            except Exception:
                pass
        # set a global hook to suppress other keys
        keyboard.hook(suppress_events)
        blocked = True
    except Exception as e:
        print("block_keyboard error:", e)

def unblock_keyboard():
    global blocked
    try:
        keyboard.unhook_all()
        for k in blocked_keys:
            try:
                keyboard.unblock_key(k)
            except Exception:
                pass
        blocked = False
    except Exception as e:
        print("unblock_keyboard error:", e)

def suppress_events(e):
    # allow only mouse clicks and allow text keys when password field has focus
    # We'll suppress all key events unless our app indicates keyboard_enabled
    if AppState.keyboard_enabled:
        return
    # Suppress key by returning True to block event
    return True

# ---------- App state ----------
class AppState:
    logged_in = False
    user_fullname = None
    user_id = None
    current_rec_id = None
    day_started = False
    keyboard_enabled = False

# ---------- GUI ----------
class TMSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TMS - Time Management System")
        self.geometry("820x420")
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        # top
        self.lbl_date = tk.Label(self, text="", font=("Arial", 14))
        self.lbl_date.place(x=20, y=10)
        self.lbl_time = tk.Label(self, text="", font=("Arial", 14))
        self.lbl_time.place(x=250, y=10)
        tk.Label(self, text="Person Name:", font=("Arial", 10)).place(x=20,y=50)
        self.txt_person = tk.Entry(self, font=("Arial", 12), state='readonly', width=30)
        self.txt_person.place(x=120,y=50)
        tk.Label(self, text="Password:", font=("Arial", 10)).place(x=20,y=90)
        self.pwd = tk.Entry(self, show='*', font=("Arial", 12), width=25)
        self.pwd.place(x=120,y=90)
        self.btn_login = tk.Button(self, text="Login", command=self.do_login)
        self.btn_login.place(x=380,y=88)
        # Buttons
        btn_specs = [
            ("Day Start","btn_daystart", self.day_start, 20, 140),
            ("Morning Tea Start","btn_mtstart", self.mt_start, 160, 140),
            ("Morning Tea End","btn_mtend", self.mt_end, 340, 140),
            ("Lunch Start","btn_lstart", self.l_start, 20, 200),
            ("Lunch End","btn_lend", self.l_end, 160, 200),
            ("Afternoon Start","btn_atstart", self.at_start, 340, 200),
            ("Afternoon End","btn_atend", self.at_end, 20, 260),
            ("Day End","btn_dayend", self.day_end, 160, 260),
        ]
        self.buttons = {}
        for text, name, cmd, x, y in btn_specs:
            b = tk.Button(self, text=text, width=18, command=cmd, state='disabled')
            b.place(x=x, y=y)
            self.buttons[name] = b
        # reminder schedule (example lunch at 13:30)
        self.reminders = [
            ("Lunch", "13:30:00", self.lunch_reminder),
            # add other reminders as ("Name", "HH:MM:SS", callback)
        ]
        # Info label
        self.info = tk.Label(self, text="Status: Not logged in", font=("Arial", 10))
        self.info.place(x=20, y=320)
        # start clock thread
        threading.Thread(target=self.clock_loop, daemon=True).start()
        # on start block keyboard until login/daystart
        block_keyboard()
        AppState.keyboard_enabled = False

    def clock_loop(self):
        while True:
            now = datetime.now()
            self.lbl_date.config(text=now.strftime("%d-%m-%Y"))
            self.lbl_time.config(text=now.strftime("%I:%M:%S %p"))
            # check reminders
            for name, hhmmss, cb in self.reminders:
                if now.strftime("%H:%M:%S") == hhmmss:
                    self.after(0, cb)
            time.sleep(1)

    def do_login(self):
        pwd_val = self.pwd.get().strip()
        if not pwd_val:
            messagebox.showinfo("TMS", "Please enter password")
            return
        full, uid = check_password(pwd_val)
        if not full:
            messagebox.showerror("TMS", "Invalid password")
            return
        AppState.logged_in = True
        AppState.user_fullname = full
        AppState.user_id = uid
        self.txt_person.config(state='normal')
        self.txt_person.delete(0, tk.END)
        self.txt_person.insert(0, full)
        self.txt_person.config(state='readonly')
        # record LoginTime in DB (create row if needed)
        rec = ensure_today_record(full)
        AppState.current_rec_id = rec
        update_field(rec, 'LoginTime', datetime.now())
        self.info.config(text=f"Logged in as {full} (LoginTime saved)")
        # After login show popup telling to press Day Start
        messagebox.showinfo("TMS", "Please click on 'Day Start' button. Otherwise you are not performing any operation(s) in your system.")
        # enable only Day Start (others disabled)
        self.enable_buttons_for_state()
        # allow typing in password field? we keep keyboard disabled until Day Start; but for password entry we allowed earlier since user used mouse to click and typed — to support typing before Day Start, set keyboard_enabled True until Day Start?
        # We'll allow typing into pwd field BEFORE Day Start using temporary enable of keyboard for the password field only:
        # For simplicity, we will enable keyboard just long enough for 10 seconds for typing, then re-block.
        self.temp_enable_keyboard_for_password()

    def temp_enable_keyboard_for_password(self):
        # enable keyboard for 10 sec so user can type password (if needed)
        AppState.keyboard_enabled = True
        unblock_keyboard()
        # After 10 seconds, re-block if DayStart not pressed
        def reblock_later():
            time.sleep(10)
            if not AppState.day_started:
                AppState.keyboard_enabled = False
                block_keyboard()
        threading.Thread(target=reblock_later, daemon=True).start()

    def enable_buttons_for_state(self):
        if not AppState.logged_in:
            # only login is possible
            for k,b in self.buttons.items():
                b.config(state='disabled')
            self.buttons['btn_daystart'].config(state='normal')
        else:
            # after login but before DayStart: only DayStart enabled
            if not AppState.day_started:
                for k,b in self.buttons.items():
                    b.config(state='disabled')
                self.buttons['btn_daystart'].config(state='normal')
            else:
                # after DayStart: enable start-type buttons and DayEnd
                for k,b in self.buttons.items():
                    b.config(state='normal')
                self.buttons['btn_daystart'].config(state='disabled')  # can't DayStart again

    # ---------- Buttons callbacks ----------
    def day_start(self):
        if not AppState.logged_in:
            messagebox.showerror("TMS", "Please login first")
            return
        AppState.day_started = True
        # unblock keyboard fully (requires admin to block earlier)
        AppState.keyboard_enabled = True
        try:
            unblock_keyboard()
        except Exception:
            pass
        update_field(AppState.current_rec_id, 'DayStart', datetime.now())
        calc_totals(AppState.current_rec_id)
        self.info.config(text="Day Started")
        self.enable_buttons_for_state()
        messagebox.showinfo("TMS", "Day Start recorded. You can now use the keyboard.")

    def mt_start(self):
        update_field(AppState.current_rec_id, 'MorningTeaBreakStart', datetime.now())
        self.info.config(text="Morning Tea Start recorded")
        # disable others except mt_end
        for k,b in self.buttons.items():
            b.config(state='disabled')
        self.buttons['btn_mtend'].config(state='normal')

    def mt_end(self):
        update_field(AppState.current_rec_id, 'MorningTeaBreakEnd', datetime.now())
        calc_totals(AppState.current_rec_id)
        self.info.config(text="Morning Tea End recorded")
        self.enable_buttons_after_break()

    def l_start(self):
        update_field(AppState.current_rec_id, 'LunchBreakStart', datetime.now())
        self.info.config(text="Lunch Start recorded")
        for k,b in self.buttons.items():
            b.config(state='disabled')
        self.buttons['btn_lend'].config(state='normal')

    def l_end(self):
        update_field(AppState.current_rec_id, 'LunchBreakEnd', datetime.now())
        calc_totals(AppState.current_rec_id)
        self.info.config(text="Lunch End recorded")
        self.enable_buttons_after_break()

    def at_start(self):
        update_field(AppState.current_rec_id, 'AfternoonTeaBreakStart', datetime.now())
        self.info.config(text="Afternoon Tea Start recorded")
        for k,b in self.buttons.items():
            b.config(state='disabled')
        self.buttons['btn_atend'].config(state='normal')

    def at_end(self):
        update_field(AppState.current_rec_id, 'AfternoonTeaBreakEnd', datetime.now())
        calc_totals(AppState.current_rec_id)
        self.info.config(text="Afternoon Tea End recorded")
        self.enable_buttons_after_break()

    def day_end(self):
        update_field(AppState.current_rec_id, 'DayEnd', datetime.now())
        calc_totals(AppState.current_rec_id)
        self.info.config(text="Day End recorded")
        # disable everything for the day
        for k,b in self.buttons.items():
            b.config(state='disabled')

    def enable_buttons_after_break(self):
        # enable start buttons and dayend
        for name, btn in self.buttons.items():
            btn.config(state='normal')
        self.buttons['btn_daystart'].config(state='disabled')

    # ---------- Reminders ----------
    def lunch_reminder(self):
        # show a popup with options: Start Lunch, Skip 10 minutes, Not Consider
        resp = messagebox.askquestion("Lunch Time", "Lunch Time - Please Start the Lunch Break.\nChoose Yes to Start Lunch, No to Skip 10 minutes.\nCancel to Not Consider lunch.")
        if resp == 'yes':
            self.l_start()
            # bring app to front
            self.lift()
            self.focus_force()
        elif resp == 'no':
            # skip 10 minutes -> schedule a one-off reminder after 10 mins
            t = datetime.now() + timedelta(minutes=10)
            t_str = t.strftime("%H:%M:%S")
            self.reminders.append(("LunchRetry", t_str, self.lunch_reminder))
            # minimize app (per your request)
            self.iconify()
        else:
            # Not considered -> set Lunch start and end to 00:00:00
            zero = datetime(1970,1,1,0,0,0)
            update_field(AppState.current_rec_id, 'LunchBreakStart', zero)
            update_field(AppState.current_rec_id, 'LunchBreakEnd', zero)
            calc_totals(AppState.current_rec_id)
            messagebox.showinfo("TMS", "Lunch marked as Not Considered.")

    def on_close(self):
        # Do not allow closing easily; minimize to tray could be implemented.
        # For now, confirm close (but your policy may require app to reopen)
        if messagebox.askokcancel("Quit", "Are you sure you want to exit TMS?"):
            try:
                unblock_keyboard()
            except Exception:
                pass
            self.destroy()

# ---------- Run ----------
def main():
    app = TMSApp()
    app.mainloop()

if __name__ == '__main__':
    main()

"""
Job Application Tracker | ระบบติดตามสถานะการสมัครงาน

โครงสร้างของไฟล์นี้ (เรียงจากบนลงล่าง):
  1) imports และการตั้งค่า High-DPI / ฟอนต์
  2) ฐานข้อมูล SQLite และค่าคงที่ (สถานะ ช่องทาง สี)
  3) ฟังก์ชัน CRUD: เพิ่ม / อ่าน / แก้ไข / ลบ ใบสมัคร
  4) DatePicker: วิดเจ็ตเลือกวันที่แบบปฏิทิน
  5) JobTrackerApp: หน้าต่างหลัก 4 แท็บ
       - แดชบอร์ดสรุป  - ติดตามผล  - รายการทั้งหมด  - บันทึกการสมัครงานใหม่
  6) Entry Point: ส่วนที่เริ่มรันโปรแกรม
"""
import os  # ทำงานกับพาธไฟล์/โฟลเดอร์ (ใช้หาโฟลเดอร์ fonts)
import sys  # ตรวจระบบปฏิบัติการที่ใช้อยู่ (sys.platform)
import csv  # เขียนไฟล์ CSV ตอนส่งออกข้อมูล
import sqlite3  # เชื่อมต่อและสั่งงานฐานข้อมูล SQLite
import webbrowser  # เปิดลิงก์ประกาศงานในเบราว์เซอร์
from datetime import datetime, date  # จัดการวันที่และเวลา
import tkinter as tk  # ไลบรารีสร้างหน้าต่างโปรแกรม (GUI)
from tkinter import ttk, messagebox, filedialog  # ttk=วิดเจ็ตธีมใหม่ / messagebox=กล่องข้อความ / filedialog=หน้าต่างเลือกไฟล์

# ------------------------------------------------------------------------------
# รองรับ High-DPI บน Windows และโหลดฟอนต์ Prompt เข้าสู่ GDI Process
# ------------------------------------------------------------------------------
# ถ้าเป็น Windows ให้เปิดโหมดรองรับจอความละเอียดสูง เพื่อให้ตัวอักษรไม่เบลอ
if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # วิธีหลัก (Windows 8.1 ขึ้นไป)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()  # วิธีสำรอง สำหรับ Windows รุ่นเก่า
        except Exception:
            pass

def register_local_fonts():
    """โหลดฟอนต์ Prompt จากโฟลเดอร์ fonts เข้าสู่ Windows GDI เพื่อให้ Tkinter เรียกใช้ได้ทันที"""
    if sys.platform == "win32":
        try:
            import ctypes
            base_dir = os.path.dirname(os.path.abspath(__file__))
            fonts_dir = os.path.join(base_dir, "fonts")
            if os.path.isdir(fonts_dir):
                # FR_PRIVATE = โหลดฟอนต์ให้ใช้ได้เฉพาะโปรแกรมนี้ ไม่ติดตั้งลงเครื่อง
                FR_PRIVATE = 0x10
                # วนอ่านไฟล์ .ttf / .otf ทุกไฟล์ในโฟลเดอร์ fonts แล้วโหลดเข้าระบบ
                for fname in os.listdir(fonts_dir):
                    if fname.lower().endswith((".ttf", ".otf")):
                        fpath = os.path.join(fonts_dir, fname)
                        ctypes.windll.gdi32.AddFontResourceExW(fpath, FR_PRIVATE, 0)
        except Exception:
            pass

# เรียกโหลดฟอนต์ทันทีเมื่อเริ่มโปรแกรม
register_local_fonts()

# กำหนด Typography หลัก
# ชื่อฟอนต์หลัก (ถ้าไม่มีในเครื่อง Tkinter จะใช้ฟอนต์สำรองแทน)
# รูปแบบฟอนต์ = (ชื่อฟอนต์, ขนาด, ลักษณะ) ตั้งเป็นค่าคงที่เพื่อเรียกใช้ซ้ำทั่วโปรแกรม
FONT_FAMILY = "Prompt"
FONT_TITLE = (FONT_FAMILY, 15, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 9)
FONT_HEADING = (FONT_FAMILY, 11, "bold")
FONT_CARD_TITLE = (FONT_FAMILY, 9)
FONT_CARD_VAL = (FONT_FAMILY, 18, "bold")
FONT_BODY = (FONT_FAMILY, 9)
FONT_BODY_BOLD = (FONT_FAMILY, 9, "bold")
FONT_BUTTON = (FONT_FAMILY, 9, "bold")
FONT_SMALL = (FONT_FAMILY, 8)

# ------------------------------------------------------------------------------
# ฐานข้อมูล SQLite & ค่าคงที่
# ------------------------------------------------------------------------------
DB_NAME = "job_tracker.db"  # ชื่อไฟล์ฐานข้อมูล (สร้างในโฟลเดอร์ที่รันโปรแกรม)

def get_db_connection():
    """เปิดการเชื่อมต่อฐานข้อมูล (ทุกฟังก์ชัน CRUD เรียกใช้ฟังก์ชันนี้)"""
    conn = sqlite3.connect(DB_NAME)
    # ทำให้อ่านค่าด้วยชื่อคอลัมน์ได้ เช่น row["company"] แทนการใช้เลขลำดับ
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """สร้างตาราง applications ถ้ายังไม่มี (ทำงานทุกครั้งที่เปิดโปรแกรม)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company TEXT NOT NULL,
            position TEXT NOT NULL,
            applied_date TEXT NOT NULL,
            channel TEXT NOT NULL,
            job_url TEXT,
            status TEXT NOT NULL,
            notes TEXT,
            updated_at TEXT
        )
    """)
    conn.commit()
    conn.close()

# เตรียมฐานข้อมูลให้พร้อมก่อนเริ่มใช้งาน
init_db()

# ตัวเลือกสถานะการสมัคร (ใช้ในฟอร์ม ตัวกรอง และกราฟ)
STATUS_OPTIONS = [
    "กำลังรอผล",
    "นัดสัมภาษณ์",
    "ผ่านสัมภาษณ์รอบแรก",
    "ได้งานแล้ว",
    "ปฏิเสธ"
]

# ตัวเลือกช่องทางการสมัคร
CHANNEL_OPTIONS = [
    "LinkedIn",
    "เว็บบริษัท",
    "Email",
    "JobsDB / JobThai",
    "เพื่อนแนะนำ",
    "อื่นๆ"
]

# สีประจำแต่ละสถานะ (ใช้ในกราฟโดนัท)
STATUS_COLORS = {
    "กำลังรอผล": "#2563EB",         # น้ำเงิน
    "นัดสัมภาษณ์": "#D97706",        # ส้มอำพัน
    "ผ่านสัมภาษณ์รอบแรก": "#7C3AED",   # ม่วง
    "ได้งานแล้ว": "#059669",          # เขียว
    "ปฏิเสธ": "#DC2626"             # แดง
}

# ------------------------------------------------------------------------------
# Business Logic & CRUD Operations
# ------------------------------------------------------------------------------
def calculate_days_waiting(applied_date_str):
    """นับจำนวนวันตั้งแต่วันที่สมัครถึงวันนี้ (คืน 0 ถ้ารูปแบบวันที่ผิด)"""
    try:
        applied_date = datetime.strptime(applied_date_str, "%Y-%m-%d").date()
        today = date.today()
        return (today - applied_date).days
    except Exception:
        return 0

def db_create(company, position, applied_date, channel, job_url, status, notes):
    """CREATE: เพิ่มใบสมัครใหม่ 1 รายการ

    ใช้ ? แทนค่าในคำสั่ง SQL (parameterized query) เพื่อป้องกัน SQL Injection
    และตัดช่องว่างหน้า-หลังข้อความด้วย strip() พร้อมบันทึกเวลาอัปเดต
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO applications (company, position, applied_date, channel, job_url, status, notes, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (company.strip(), position.strip(), applied_date, channel, job_url.strip(), status, notes.strip(), now_str))
    conn.commit()
    conn.close()

def db_read_all():
    """READ: ดึงใบสมัครทั้งหมด เรียงวันที่สมัครล่าสุดก่อน (ถ้าวันที่เท่ากันเรียงตาม id ล่าสุดก่อน)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM applications ORDER BY applied_date DESC, id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def db_read_one(app_id):
    """READ: ดึงใบสมัคร 1 รายการตาม id (ไม่พบจะได้ None)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM applications WHERE id = ?", (app_id,))
    row = cursor.fetchone()
    conn.close()
    return row

def db_update(app_id, company, position, applied_date, channel, job_url, status, notes):
    """UPDATE: แก้ไขข้อมูลทุกช่องของใบสมัครตาม id พร้อมอัปเดตเวลาแก้ไข"""
    conn = get_db_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        UPDATE applications
        SET company = ?, position = ?, applied_date = ?, channel = ?, 
            job_url = ?, status = ?, notes = ?, updated_at = ?
        WHERE id = ?
    """, (company.strip(), position.strip(), applied_date, channel, job_url.strip(), status, notes.strip(), now_str, app_id))
    conn.commit()
    conn.close()

def db_update_status(app_id, new_status):
    """UPDATE: เปลี่ยนเฉพาะสถานะของใบสมัคร (ใช้กับปุ่มลัดในหน้าติดตามผล)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("UPDATE applications SET status = ?, updated_at = ? WHERE id = ?", (new_status, now_str, app_id))
    conn.commit()
    conn.close()

def db_reapply_followup(app_id, current_notes):
    """UPDATE: ติดตามผลซ้ำ โดยรีเซ็ตวันที่สมัครเป็นวันนี้ และต่อท้ายหมายเหตุว่าติดตามเมื่อไหร่"""
    conn = get_db_connection()
    cursor = conn.cursor()
    today_str = date.today().strftime("%Y-%m-%d")
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    followup_msg = f"[ติดตามผลซ้ำเมื่อ {today_str}]"
    # ถ้ามีหมายเหตุเดิมให้ต่อท้ายบรรทัดใหม่ ถ้าไม่มีก็ใช้ข้อความติดตามผลอย่างเดียว
    updated_notes = f"{current_notes}\n{followup_msg}".strip() if current_notes else followup_msg
    
    cursor.execute("""
        UPDATE applications
        SET applied_date = ?, notes = ?, updated_at = ?
        WHERE id = ?
    """, (today_str, updated_notes, now_str, app_id))
    conn.commit()
    conn.close()

def db_delete(app_id):
    """DELETE: ลบใบสมัครตาม id"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM applications WHERE id = ?", (app_id,))
    conn.commit()
    conn.close()

def db_seed_sample_data():
    """ใส่ข้อมูลตัวอย่าง 7 รายการ เฉพาะเมื่อตารางยังว่าง (กันข้อมูลซ้ำ)"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) as count FROM applications")
    count = cursor.fetchone()['count']
    if count == 0:
        # แต่ละรายการ = (บริษัท, ตำแหน่ง, วันที่สมัคร, ช่องทาง, ลิงก์, สถานะ, หมายเหตุ, เวลาอัปเดต)
        sample_records = [
            ("Agoda", "Software Engineer Intern", "2026-09-10", "LinkedIn", "https://careers.agoda.com", "กำลังรอผล", "ส่งใบสมัครพร้อมแนบ GitHub Portfolio", "2026-09-10 10:00:00"),
            ("SCB 10X", "Backend Developer (Go/Python)", "2026-09-28", "เว็บบริษัท", "https://scb10x.com/careers", "กำลังรอผล", "สมัครตำแหน่งฝึกงานช่วงปิดเทอม", "2026-09-28 14:30:00"),
            ("LINE MAN Wongnai", "Frontend Developer Intern", "2026-10-18", "LinkedIn", "https://careers.lmwn.com", "กำลังรอผล", "รอ HR ติดต่อกลับรอบ Resume Screening", "2026-10-18 09:15:00"),
            ("KBTG", "Data Scientist Intern", "2026-10-05", "Email", "https://kbtg.tech", "นัดสัมภาษณ์", "สัมภาษณ์รอบแรกผ่าน Zoom วันที่ 28 ต.ค.", "2026-10-10 11:20:00"),
            ("True Digital Group", "Full Stack Developer", "2026-09-15", "เว็บบริษัท", "https://careers.truedigital.com", "ผ่านสัมภาษณ์รอบแรก", "ผ่านด่านเทสโค้ด รอสัมภาษณ์กับ Tech Lead", "2026-10-01 16:45:00"),
            ("Bitkub", "Junior Python Developer", "2026-08-20", "LinkedIn", "https://careers.bitkub.com", "ได้งานแล้ว", "เซ็นสัญญาและเริ่มงานเดือนหน้า", "2026-09-25 15:00:00"),
            ("Shopee", "QA Automation Engineer", "2026-08-10", "อื่นๆ", "https://careers.shopee.co.th", "ปฏิเสธ", "บริษัทแจ้งว่ารับคนเต็มจำนวนแล้ว", "2026-09-01 13:00:00"),
        ]
        # executemany = เพิ่มหลายแถวในคำสั่งเดียว
        cursor.executemany("""
            INSERT INTO applications (company, position, applied_date, channel, job_url, status, notes, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, sample_records)
        conn.commit()
    conn.close()

# ------------------------------------------------------------------------------
# ตัวเลือกวันที่แบบปฏิทิน (DatePicker) - ไม่ต้องติดตั้งไลบรารีเพิ่ม
# ------------------------------------------------------------------------------
# ชื่อเดือนและตัวย่อวันภาษาไทย สำหรับแสดงในปฏิทิน
THAI_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
               "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
THAI_WEEKDAYS = ["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"]


class DatePicker(tk.Frame):
    """ช่องแสดงวันที่ (YYYY-MM-DD) พร้อมปุ่มกดเปิดปฏิทินให้เลือกวัน"""

    def __init__(self, parent, initial=None, bg="#FFFFFF", width=14):
        """สร้างช่องวันที่ + ปุ่มเปิดปฏิทิน

        parent = วิดเจ็ตที่จะวางลงไป, initial = วันที่เริ่มต้น (ไม่ใส่ = วันนี้), bg = สีพื้นหลัง
        """
        super().__init__(parent, bg=bg)
        self._value = tk.StringVar()
        # ช่องแสดงวันที่แบบอ่านอย่างเดียว (readonly) จะได้พิมพ์รูปแบบผิดไม่ได้
        self.entry = ttk.Entry(self, textvariable=self._value, width=width, state="readonly", justify="center")
        self.entry.pack(side="left")
        # ปุ่มรูปปฏิทิน กดเพื่อเปิดหน้าต่างเลือกวันที่
        self.btn = tk.Button(
            self, text="📅", font=FONT_BODY, relief="flat", bg="#E2E8F0", fg="#0F172A",
            activebackground="#CBD5E1", cursor="hand2", padx=6, command=self._open_popup
        )
        self.btn.pack(side="left", padx=(4, 0))
        # คลิกที่ช่องวันที่ก็เปิดปฏิทินได้เหมือนกัน
        self.entry.bind("<Button-1>", lambda e: self._open_popup())
        # ตั้งค่าเริ่มต้น: ใช้วันที่ที่ส่งมา ถ้าไม่ส่งมาใช้วันนี้
        self.set(initial or date.today())
        self._popup = None

    def get(self):
        """คืนวันที่ที่เลือกอยู่ เป็นข้อความรูปแบบ YYYY-MM-DD"""
        return self._value.get()

    def set(self, value):
        """ตั้งวันที่ รับได้ทั้ง date/datetime หรือข้อความ YYYY-MM-DD"""
        if isinstance(value, (date, datetime)):
            value = value.strftime("%Y-%m-%d")
        self._value.set(str(value))

    def _current_date(self):
        """อ่านวันที่ในช่องเป็นวัตถุ date ถ้าอ่านไม่ได้ให้ใช้วันนี้แทน"""
        try:
            return datetime.strptime(self._value.get(), "%Y-%m-%d").date()
        except ValueError:
            return date.today()

    def _open_popup(self):
        """เปิดหน้าต่างปฏิทิน (ถ้าเปิดอยู่แล้วให้ปิด)"""
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.destroy()
            return
        selected = self._current_date()
        self._view_year, self._view_month = selected.year, selected.month

        # สร้างหน้าต่างย่อยสำหรับปฏิทิน
        pop = tk.Toplevel(self)
        self._popup = pop
        pop.title("เลือกวันที่")
        pop.resizable(False, False)
        pop.transient(self.winfo_toplevel())
        pop.configure(bg="#FFFFFF", padx=10, pady=10)
        self.update_idletasks()
        # วางปฏิทินไว้ใต้ช่องวันที่
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 4
        pop.geometry(f"+{x}+{y}")

        # แถบบนสุด: ปุ่มเดือนก่อนหน้า / ชื่อเดือนและปี / ปุ่มเดือนถัดไป
        header = tk.Frame(pop, bg="#FFFFFF")
        header.pack(fill="x")
        tk.Button(header, text="◀", font=FONT_BODY, relief="flat", bg="#E2E8F0", width=3,
                  command=lambda: self._shift_month(-1)).pack(side="left")
        tk.Button(header, text="▶", font=FONT_BODY, relief="flat", bg="#E2E8F0", width=3,
                  command=lambda: self._shift_month(1)).pack(side="right")
        self._title = tk.Label(header, font=FONT_BODY_BOLD, bg="#FFFFFF", fg="#0F172A")
        self._title.pack(side="left", expand=True)

        # พื้นที่แสดงตารางวันที่ (ถูกวาดใหม่ทุกครั้งที่เปลี่ยนเดือน)
        self._grid = tk.Frame(pop, bg="#FFFFFF")
        self._grid.pack(pady=(8, 6))

        # ปุ่มลัดเลือกวันนี้
        tk.Button(pop, text="วันนี้", font=FONT_BUTTON, relief="flat", bg="#DBEAFE", fg="#1D4ED8",
                  cursor="hand2", command=lambda: self._choose(date.today())).pack(fill="x")

        self._render_month()
        pop.bind("<Escape>", lambda e: pop.destroy())
        pop.focus_set()

    def _shift_month(self, delta):
        """เลื่อนเดือน delta = -1 ย้อนกลับ, +1 ไปข้างหน้า (ข้ามปีให้อัตโนมัติ)"""
        m = self._view_month + delta
        y = self._view_year
        if m < 1:
            m, y = 12, y - 1
        elif m > 12:
            m, y = 1, y + 1
        self._view_year, self._view_month = y, m
        self._render_month()

    def _render_month(self):
        """วาดตารางวันของเดือนที่กำลังดูอยู่ เป็นปุ่มวันที่เรียงตามวันในสัปดาห์"""
        import calendar
        for w in self._grid.winfo_children():
            w.destroy()
        self._title.config(text=f"{THAI_MONTHS[self._view_month - 1]} {self._view_year}")
        for col, name in enumerate(THAI_WEEKDAYS):
            tk.Label(self._grid, text=name, font=FONT_SMALL, fg="#64748B", bg="#FFFFFF", width=4).grid(row=0, column=col)

        selected = self._current_date()
        today = date.today()
        # ให้สัปดาห์เริ่มที่วันอาทิตย์ (6 = Sunday)
        # monthdayscalendar คืนข้อมูลเป็นรายสัปดาห์ ช่องที่ไม่ใช่วันของเดือนนี้จะเป็น 0
        cal = calendar.Calendar(firstweekday=6)
        for r, week in enumerate(cal.monthdayscalendar(self._view_year, self._view_month), start=1):
            for c, day in enumerate(week):
                # ช่องว่างนอกเดือน ไม่ต้องสร้างปุ่ม
                if day == 0:
                    continue
                d = date(self._view_year, self._view_month, day)
                # สีวันปกติ / วันที่เลือกอยู่ = น้ำเงิน / วันนี้ = ฟ้าอ่อน
                bg, fg = "#FFFFFF", "#0F172A"
                if d == selected:
                    bg, fg = "#2563EB", "#FFFFFF"
                elif d == today:
                    bg, fg = "#DBEAFE", "#1D4ED8"
                tk.Button(
                    self._grid, text=str(day), font=FONT_BODY, width=3, relief="flat",
                    bg=bg, fg=fg, activebackground="#BFDBFE", cursor="hand2",
                    command=lambda dd=d: self._choose(dd)  # dd=d เพื่อจำวันที่ของปุ่มนั้นไว้ (ไม่ให้ค่าทับกันในลูป)
                ).grid(row=r, column=c, padx=1, pady=1)

    def _choose(self, d):
        """เมื่อกดเลือกวันที่: บันทึกค่าลงช่อง แล้วปิดปฏิทิน"""
        self.set(d)
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.destroy()


# ------------------------------------------------------------------------------
# คลาสหน้าต่างแอปพลิเคชันหลัก (Desktop GUI)
# ------------------------------------------------------------------------------
class JobTrackerApp(tk.Tk):
    """หน้าต่างหลักของโปรแกรม (สืบทอดจาก tk.Tk)

    ประกอบด้วย 4 แท็บ: แดชบอร์ดสรุป / ติดตามผล / รายการทั้งหมด / บันทึกการสมัครงานใหม่
    """
    def __init__(self):
        """สร้างหน้าต่าง ตั้งค่าธีม สร้างส่วนหัวและแท็บทั้ง 4 แล้วโหลดข้อมูลครั้งแรก"""
        # เรียกตัวสร้างของ tk.Tk เพื่อสร้างหน้าต่างหลักขึ้นมาก่อน
        super().__init__()
        
        self.title("Job Application Tracker | ระบบติดตามสถานะการสมัครงาน")
        self.minsize(1100, 700)  # ขนาดเล็กสุดที่ผู้ใช้ย่อหน้าต่างได้
        self.configure(bg="#F8FAFC")
        self.center_window(1220, 760)  # เปิดหน้าต่างขนาด 1220x760 ตรงกลางจอ
        
        self.setup_styles()  # ตั้งสีและฟอนต์ของวิดเจ็ต ttk
        self.create_header()  # สร้างส่วนหัวด้านบน
        
        # แท็บเมนูหลัก (Notebook)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=18, pady=(0, 14))
        
        # สร้างเฟรมว่างสำหรับแต่ละแท็บ (ยังไม่ใส่เนื้อหา)
        self.tab_dashboard = ttk.Frame(self.notebook)
        self.tab_alerts = ttk.Frame(self.notebook)
        self.tab_manage = ttk.Frame(self.notebook)
        self.tab_add = ttk.Frame(self.notebook)
        
        # เพิ่มแท็บเข้า Notebook พร้อมข้อความบนหัวแท็บ
        self.notebook.add(self.tab_dashboard, text=" 📊 แดชบอร์ดสรุป ")
        self.notebook.add(self.tab_alerts, text=" 🔔 ติดตามผล ")
        self.notebook.add(self.tab_manage, text=" 📋 รายการทั้งหมด ")
        self.notebook.add(self.tab_add, text=" ➕ บันทึกการสมัครงานใหม่ ")
        
        # สร้างเนื้อหาข้างในแต่ละแท็บ
        self.build_dashboard_tab()
        self.build_alerts_tab()
        self.build_manage_tab()
        self.build_add_tab()
        
        # โหลดข้อมูลจากฐานข้อมูลมาแสดงครั้งแรก
        self.refresh_all_data()
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self.refresh_all_data())  # รีเฟรชข้อมูลทุกครั้งที่สลับแท็บ

    def center_window(self, width, height):
        """ตั้งขนาดหน้าต่างและวางไว้กลางหน้าจอ"""
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def setup_styles(self):
        """กำหนดหน้าตา (สี ฟอนต์ ระยะห่าง) ของแท็บ ช่องกรอก และตาราง"""
        style = ttk.Style(self)
        # ธีม clam ปรับสีได้มากกว่าธีมเริ่มต้นของ Windows
        if 'clam' in style.theme_names():
            style.theme_use('clam')
            
        # หน้าตาของแท็บ (TNotebook = กรอบรวม, TNotebook.Tab = หัวแท็บแต่ละอัน)
        style.configure("TNotebook", background="#F8FAFC", borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            font=FONT_BUTTON,
            padding=[18, 9],
            background="#E2E8F0",
            foreground="#475569",
            borderwidth=0
        )
        # เปลี่ยนสีหัวแท็บที่ถูกเลือก (selected) เป็นน้ำเงิน ตัวอักษรขาว
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#2563EB")],
            foreground=[("selected", "#FFFFFF")]
        )
        
        # ช่องกรอกข้อความและ Combobox
        style.configure("TEntry", font=FONT_BODY, padding=6)
        style.configure("TCombobox", font=FONT_BODY, padding=5)
        
        style.configure(
            "Treeview",
            font=FONT_BODY,
            rowheight=32,  # ความสูงของแต่ละแถวในตาราง
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            borderwidth=0
        )
        style.configure(
            "Treeview.Heading",
            font=FONT_BODY_BOLD,
            background="#F1F5F9",
            foreground="#1E293B",
            relief="flat",
            padding=[8, 6]
        )
        style.map("Treeview.Heading", background=[("active", "#E2E8F0")])  # เปลี่ยนสีหัวตารางเมื่อเอาเมาส์ชี้

    def create_header(self):
        """สร้างส่วนหัว: ชื่อโปรแกรม (ซ้าย) และปุ่มข้อมูลตัวอย่าง (ขวา)"""
        header = tk.Frame(self, bg="#FFFFFF", padx=22, pady=14, bd=1, relief="solid")
        header.configure(highlightbackground="#E2E8F0", highlightthickness=1)
        header.pack(fill="x", padx=18, pady=14)
        
        left = tk.Frame(header, bg="#FFFFFF")
        left.pack(side="left")
        
        tk.Label(left, text="💼 Job Application Tracker", font=FONT_TITLE, fg="#0F172A", bg="#FFFFFF").pack(anchor="w")
        tk.Label(left, text="ระบบติดตามสถานะการสมัครงานและฝึกงาน", font=FONT_SUBTITLE, fg="#64748B", bg="#FFFFFF").pack(anchor="w", pady=(2, 0))
        
        right = tk.Frame(header, bg="#FFFFFF")
        right.pack(side="right")
        
        # ปุ่มสร้างข้อมูลตัวอย่าง (command = ฟังก์ชันที่ทำงานเมื่อกดปุ่ม)
        demo_btn = tk.Button(
            right,
            text="🎲 ข้อมูลตัวอย่าง",
            font=FONT_BUTTON,
            bg="#EEF2F6",
            fg="#2563EB",
            activebackground="#DBEAFE",
            activeforeground="#1D4ED8",
            relief="groove",
            bd=1,
            padx=14,
            pady=5,
            cursor="hand2",
            command=self.on_seed_demo
        )
        demo_btn.pack()

    # ==========================================================================
    # TAB 1: DASHBOARD
    # ==========================================================================
    def build_dashboard_tab(self):
        """สร้างกรอบเปล่าของแท็บแดชบอร์ด (เนื้อหาถูกวาดใน render_dashboard)"""
        self.dash_content = tk.Frame(self.tab_dashboard, bg="#F8FAFC", padx=16, pady=14)
        self.dash_content.pack(fill="both", expand=True)

    def render_dashboard(self, rows):
        """วาดแดชบอร์ดใหม่ทั้งหมดจากข้อมูล rows: การ์ดตัวเลข แถบแจ้งเตือน และกราฟ 2 ชนิด"""
        # ล้างของเก่าออกก่อนวาดใหม่ ไม่เช่นนั้นจะซ้อนกัน
        for widget in self.dash_content.winfo_children():
            widget.destroy()

        # คำนวณตัวเลขสรุปจากข้อมูลทั้งหมด
        total = len(rows)
        waiting = sum(1 for r in rows if r['status'] == "กำลังรอผล")
        interview = sum(1 for r in rows if r['status'] in ["นัดสัมภาษณ์", "ผ่านสัมภาษณ์รอบแรก"])
        accepted = sum(1 for r in rows if r['status'] == "ได้งานแล้ว")
        rejected = sum(1 for r in rows if r['status'] == "ปฏิเสธ")
        success_rate = (accepted / total * 100) if total > 0 else 0.0  # อัตราความสำเร็จ = ได้งาน ÷ สมัครทั้งหมด × 100
        
        # นับรายการที่รอผลนาน: วิกฤต = เกิน 30 วัน, เตือน = 15-30 วัน
        crit_count = sum(1 for r in rows if r['status'] == "กำลังรอผล" and calculate_days_waiting(r['applied_date']) > 30)
        warn_count = sum(1 for r in rows if r['status'] == "กำลังรอผล" and 14 < calculate_days_waiting(r['applied_date']) <= 30)
        total_alerts = crit_count + warn_count

        # แถวการ์ดสรุปตัวเลข (Metrics)
        metrics_frame = tk.Frame(self.dash_content, bg="#F8FAFC")
        metrics_frame.pack(fill="x", pady=(0, 14))

        # ข้อมูลการ์ดสรุป = (ชื่อ, ค่า, หน่วย, สีตัวเลข, สีแถบ)
        cards_data = [
            ("สมัครทั้งหมด", f"{total}", "ที่", "#2563EB", "#DBEAFE"),
            ("กำลังรอผล", f"{waiting}", "ที่", "#D97706", "#FEF3C7"),
            ("นัดสัมภาษณ์", f"{interview}", "ที่", "#7C3AED", "#EDE9FE"),
            ("ได้งานแล้ว", f"{accepted}", "ที่", "#059669", "#D1FAE5"),
            ("ปฏิเสธ", f"{rejected}", "ที่", "#DC2626", "#FEE2E2"),
            ("ความสำเร็จ", f"{success_rate:.1f}", "%", "#0284C7", "#E0F2FE")
        ]

        # วนสร้างการ์ดทีละใบ
        for title, val, unit, fg_color, bg_color in cards_data:
            card = tk.Frame(metrics_frame, bg="#FFFFFF", padx=14, pady=12, bd=1, relief="solid")
            card.configure(highlightbackground="#E2E8F0", highlightthickness=1)
            card.pack(side="left", fill="both", expand=True, padx=4)
            
            # Indicator bar ด้านบนการ์ด
            bar = tk.Frame(card, bg=fg_color, height=3)
            bar.pack(fill="x", pady=(0, 8))
            
            tk.Label(card, text=title, font=FONT_CARD_TITLE, fg="#64748B", bg="#FFFFFF").pack(anchor="w")
            
            val_frame = tk.Frame(card, bg="#FFFFFF")
            val_frame.pack(anchor="w", pady=(4, 0))
            tk.Label(val_frame, text=val, font=FONT_CARD_VAL, fg=fg_color, bg="#FFFFFF").pack(side="left")
            tk.Label(val_frame, text=f" {unit}", font=FONT_BODY, fg="#94A3B8", bg="#FFFFFF").pack(side="left", padx=(2, 0), pady=(6, 0))

        # แถบแจ้งเตือนค้างผล
        if total_alerts > 0:
            alert_bar = tk.Frame(self.dash_content, bg="#FFFBEB", padx=16, pady=10, bd=1, relief="solid")
            alert_bar.configure(highlightbackground="#F59E0B", highlightthickness=1)
            alert_bar.pack(fill="x", pady=(0, 14))
            
            msg = f"⚠️ มีรายการรอผลเกินกำหนด {total_alerts} รายการ (เกิน 30 วัน: {crit_count} | เกิน 14 วัน: {warn_count})"
            tk.Label(alert_bar, text=msg, font=FONT_BODY_BOLD, fg="#92400E", bg="#FFFBEB").pack(side="left")
            
            goto_btn = tk.Button(
                alert_bar,
                text="ดูรายการแจ้งเตือน ➔",
                font=FONT_BUTTON,
                bg="#F59E0B",
                fg="#FFFFFF",
                relief="flat",
                padx=10,
                pady=3,
                cursor="hand2",
                command=lambda: self.notebook.select(self.tab_alerts)
            )
            goto_btn.pack(side="right")

        # กราฟสรุปผล (Canvas Charts)
        charts_frame = tk.Frame(self.dash_content, bg="#F8FAFC")
        charts_frame.pack(fill="both", expand=True)
        
        # Donut Chart
        left_box = tk.Frame(charts_frame, bg="#FFFFFF", padx=16, pady=14, bd=1, relief="solid")
        left_box.configure(highlightbackground="#E2E8F0", highlightthickness=1)
        left_box.pack(side="left", fill="both", expand=True, padx=(0, 6))
        
        tk.Label(left_box, text="🥧 สัดส่วนสถานะการสมัครงาน", font=FONT_HEADING, fg="#0F172A", bg="#FFFFFF").pack(anchor="w")
        canvas_pie = tk.Canvas(left_box, bg="#FFFFFF", highlightthickness=0, height=270)
        canvas_pie.pack(fill="both", expand=True, pady=8)
        self.draw_donut_chart(canvas_pie, rows, total)

        # Bar Chart
        right_box = tk.Frame(charts_frame, bg="#FFFFFF", padx=16, pady=14, bd=1, relief="solid")
        right_box.configure(highlightbackground="#E2E8F0", highlightthickness=1)
        right_box.pack(side="right", fill="both", expand=True, padx=(6, 0))
        
        tk.Label(right_box, text="🌐 ช่องทางการสมัครงาน", font=FONT_HEADING, fg="#0F172A", bg="#FFFFFF").pack(anchor="w")
        canvas_bar = tk.Canvas(right_box, bg="#FFFFFF", highlightthickness=0, height=270)
        canvas_bar.pack(fill="both", expand=True, pady=8)
        self.draw_bar_chart(canvas_bar, rows)

    def draw_donut_chart(self, canvas, rows, total):
        """วาดกราฟโดนัทสัดส่วนสถานะ: แต่ละสถานะเป็นส่วนโค้ง ขนาดตามจำนวนใบสมัคร"""
        if total == 0:
            canvas.create_text(200, 130, text="ยังไม่มีข้อมูลสำหรับการแสดงผล", font=FONT_BODY, fill="#94A3B8")
            return

        counts = {s: sum(1 for r in rows if r['status'] == s) for s in STATUS_OPTIONS}  # นับจำนวนใบสมัครของแต่ละสถานะ
        cx, cy, r_outer, r_inner = 135, 130, 95, 52  # จุดศูนย์กลาง และรัศมีวงนอก/วงใน
        start_angle = 0.0

        for status, count in counts.items():
            if count == 0:
                continue
            extent = (count / total) * 360.0  # มุมของส่วนนี้ = สัดส่วน × 360 องศา
            color = STATUS_COLORS.get(status, "#94A3B8")
            # วาดส่วนโค้งของสถานะนี้ ต่อจากมุมที่วาดค้างไว้ (start_angle)
            canvas.create_arc(
                cx - r_outer, cy - r_outer, cx + r_outer, cy + r_outer,
                start=start_angle, extent=extent, fill=color, outline="#FFFFFF", width=2
            )
            start_angle += extent

        # เจาะรูตรงกลางสร้าง Donut Hole
        canvas.create_oval(cx - r_inner, cy - r_inner, cx + r_inner, cy + r_inner, fill="#FFFFFF", outline="#FFFFFF")
        canvas.create_text(cx, cy - 8, text=f"{total}", font=(FONT_FAMILY, 15, "bold"), fill="#0F172A")
        canvas.create_text(cx, cy + 12, text="สมัครทั้งหมด", font=FONT_SMALL, fill="#64748B")

        # Legend ทางขวา
        legend_x = 265
        legend_y = 40
        for status in STATUS_OPTIONS:
            cnt = counts[status]
            pct = (cnt / total * 100) if total > 0 else 0
            color = STATUS_COLORS.get(status, "#94A3B8")
            
            canvas.create_oval(legend_x, legend_y + 2, legend_x + 10, legend_y + 12, fill=color, outline="")
            canvas.create_text(
                legend_x + 18, legend_y + 7, anchor="w",
                text=f"{status}: {cnt} ({pct:.0f}%)", font=FONT_BODY, fill="#334155"
            )
            legend_y += 32

    def draw_bar_chart(self, canvas, rows):
        """วาดกราฟแท่งแนวนอน แสดงจำนวนใบสมัครของแต่ละช่องทาง"""
        if not rows:
            canvas.create_text(200, 130, text="ยังไม่มีข้อมูลสำหรับการแสดงผล", font=FONT_BODY, fill="#94A3B8")
            return

        channel_counts = {c: sum(1 for r in rows if r['channel'] == c) for c in CHANNEL_OPTIONS}
        active_channels = [(c, cnt) for c, cnt in channel_counts.items() if cnt > 0]
        if not active_channels:
            active_channels = list(channel_counts.items())[:4]

        max_val = max([cnt for _, cnt in active_channels] + [1])  # ค่าสูงสุด ใช้กำหนดความยาวแท่งให้เป็นสัดส่วนกัน
        y_pos = 28
        bar_height = 20
        max_bar_width = 240
        start_x = 135

        for ch_name, cnt in active_channels:
            canvas.create_text(
                start_x - 12, y_pos + (bar_height // 2), anchor="e",
                text=ch_name, font=FONT_BODY_BOLD, fill="#475569"
            )
            
            bar_w = int((cnt / max_val) * max_bar_width) if max_val > 0 else 0  # ความยาวแท่ง = สัดส่วนเทียบกับค่าสูงสุด
            if bar_w < 8 and cnt > 0:  # ถ้าค่าน้อยมาก ให้แท่งยาวอย่างน้อย 8 px จะได้มองเห็น
                bar_w = 8
                
            canvas.create_rectangle(
                start_x, y_pos, start_x + bar_w, y_pos + bar_height,
                fill="#3B82F6", outline="#2563EB"
            )
            
            canvas.create_text(
                start_x + bar_w + 8, y_pos + (bar_height // 2), anchor="w",
                text=f"{cnt} ที่", font=FONT_BODY_BOLD, fill="#0F172A"
            )
            y_pos += 38

    # ==========================================================================
    # TAB 2: SMART ALERTS
    # ==========================================================================
    def build_alerts_tab(self):
        """สร้างโครงแท็บติดตามผล: หัวข้อ + พื้นที่เลื่อนดูรายการได้ (Canvas + Scrollbar)"""
        top_frame = tk.Frame(self.tab_alerts, bg="#FFFFFF", padx=18, pady=12, bd=1, relief="solid")
        top_frame.configure(highlightbackground="#E2E8F0", highlightthickness=1)
        top_frame.pack(fill="x", padx=14, pady=12)
        
        tk.Label(top_frame, text="🔔 รายการติดตามผล", font=FONT_HEADING, fg="#0F172A", bg="#FFFFFF").pack(anchor="w")
        tk.Label(top_frame, text="รายการสถานะ 'กำลังรอผล' ที่เกิน 14 วัน หรือ 30 วัน", font=FONT_SUBTITLE, fg="#64748B", bg="#FFFFFF").pack(anchor="w", pady=(2, 0))

        container = tk.Frame(self.tab_alerts, bg="#F8FAFC")
        container.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        
        # เทคนิคทำให้เลื่อนได้: วาง Frame ไว้ใน Canvas แล้วผูกกับ Scrollbar
        self.alert_canvas = tk.Canvas(container, bg="#F8FAFC", highlightthickness=0)
        self.alert_scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.alert_canvas.yview)
        self.alert_scroll_frame = tk.Frame(self.alert_canvas, bg="#F8FAFC")
        
        self.alert_scroll_frame.bind(
            "<Configure>",
            lambda e: self.alert_canvas.configure(scrollregion=self.alert_canvas.bbox("all"))
        )
        self.alert_win = self.alert_canvas.create_window((0, 0), window=self.alert_scroll_frame, anchor="nw")
        self.alert_canvas.configure(yscrollcommand=self.alert_scrollbar.set)
        
        self.alert_canvas.bind("<Configure>", lambda e: self.alert_canvas.itemconfig(self.alert_win, width=e.width))
        
        # เลื่อนด้วยลูกกลิ้งเมาส์ (ทำงานเฉพาะตอนเมาส์อยู่เหนือพื้นที่นี้)
        def _on_alert_wheel(event):
            self.alert_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.alert_canvas.bind("<Enter>", lambda e: self.alert_canvas.bind_all("<MouseWheel>", _on_alert_wheel))
        self.alert_canvas.bind("<Leave>", lambda e: self.alert_canvas.unbind_all("<MouseWheel>"))

        self.alert_canvas.pack(side="left", fill="both", expand=True)
        self.alert_scrollbar.pack(side="right", fill="y")

    def render_alerts(self, rows):
        """แสดงการ์ดแจ้งเตือนของรายการที่ 'กำลังรอผล' เกิน 14 วัน"""
        for w in self.alert_scroll_frame.winfo_children():
            w.destroy()

        # คัดเฉพาะรายการสถานะ "กำลังรอผล" ที่รอเกิน 14 วัน
        delayed_items = []
        for r in rows:
            if r['status'] == "กำลังรอผล":
                days = calculate_days_waiting(r['applied_date'])
                if days > 14:
                    delayed_items.append((r, days))

        # แสดงจำนวนรายการค้างบนหัวแท็บ เช่น ติดตามผล (3)
        tab_text = f" 🔔 ติดตามผล ({len(delayed_items)}) " if delayed_items else " 🔔 ติดตามผล "
        self.notebook.tab(self.tab_alerts, text=tab_text)

        # ไม่มีรายการค้าง -> แสดงข้อความ "ไม่มีรายการค้างเกินกำหนด" แล้วจบฟังก์ชัน
        if not delayed_items:
            empty_box = tk.Frame(self.alert_scroll_frame, bg="#FFFFFF", padx=36, pady=32, bd=1, relief="solid")
            empty_box.configure(highlightbackground="#E2E8F0", highlightthickness=1)
            empty_box.pack(fill="x", padx=10, pady=24)
            
            tk.Label(empty_box, text="🎉 ไม่มีรายการค้างเกินกำหนด", font=FONT_TITLE, fg="#059669", bg="#FFFFFF").pack()
            tk.Label(empty_box, text="ทุกรายการสมัครงานยังอยู่ในระยะเวลาที่เหมาะสม", font=FONT_BODY, fg="#64748B", bg="#FFFFFF").pack(pady=4)
            return

        for item, days in delayed_items:
            app_id = item['id']
            # เกิน 30 วัน = สีแดง (วิกฤต) / 15-30 วัน = สีเหลือง (เตือน)
            is_critical = (days > 30)
            
            card_bg = "#FEF2F2" if is_critical else "#FFFBEB"
            border_color = "#EF4444" if is_critical else "#F59E0B"
            badge_bg = "#DC2626" if is_critical else "#D97706"
            level_text = f"🚨 เกิน {days} วัน" if is_critical else f"⚠️ รอผล {days} วัน"
            
            card = tk.Frame(self.alert_scroll_frame, bg=card_bg, padx=16, pady=12, bd=1, relief="solid")
            card.configure(highlightbackground=border_color, highlightthickness=1)
            card.pack(fill="x", padx=10, pady=6)
            
            top_row = tk.Frame(card, bg=card_bg)
            top_row.pack(fill="x")
            
            company_lbl = tk.Label(
                top_row,
                text=f"🏢 {item['company']}",
                font=FONT_HEADING,
                fg="#991B1B" if is_critical else "#92400E",
                bg=card_bg
            )
            company_lbl.pack(side="left")
            
            pos_lbl = tk.Label(top_row, text=f" — {item['position']}", font=FONT_BODY, fg="#334155", bg=card_bg)
            pos_lbl.pack(side="left")
            
            badge_lbl = tk.Label(top_row, text=level_text, font=FONT_BUTTON, fg="#FFFFFF", bg=badge_bg, padx=8, pady=2)
            badge_lbl.pack(side="right")
            
            info_text = f"📅 วันที่สมัคร: {item['applied_date']}  •  🌐 ช่องทาง: {item['channel']}"
            if item['notes']:
                info_text += f"\n📝 หมายเหตุ: {item['notes']}"
            
            info_lbl = tk.Label(card, text=info_text, font=FONT_BODY, fg="#475569", bg=card_bg, justify="left")
            info_lbl.pack(anchor="w", pady=(6, 8))
            
            btn_row = tk.Frame(card, bg=card_bg)
            btn_row.pack(fill="x")
            
            tk.Label(btn_row, text="การจัดการ: ", font=FONT_BODY_BOLD, fg="#1E293B", bg=card_bg).pack(side="left")
            
            # ปุ่มลัดในการ์ด: ติดตามผลซ้ำ / นัดสัมภาษณ์ / ได้งานแล้ว / ลบ
            # (lambda ใช้ส่งค่าของรายการนั้นเข้าฟังก์ชัน)
            btn_reapply = tk.Button(
                btn_row, text="🔄 ติดตามผลซ้ำ", font=FONT_BUTTON,
                bg="#2563EB", fg="#FFFFFF", padx=10, pady=3, relief="groove", cursor="hand2",
                command=lambda aid=app_id, n=item['notes'], c=item['company']: self.on_quick_reapply(aid, n, c)
            )
            btn_reapply.pack(side="left", padx=4)
            
            btn_int = tk.Button(
                btn_row, text="🎯 นัดสัมภาษณ์", font=FONT_BUTTON,
                bg="#7C3AED", fg="#FFFFFF", padx=10, pady=3, relief="groove", cursor="hand2",
                command=lambda aid=app_id, c=item['company']: self.on_quick_status(aid, "นัดสัมภาษณ์", c)
            )
            btn_int.pack(side="left", padx=4)
            
            btn_acc = tk.Button(
                btn_row, text="🎉 ได้งานแล้ว", font=FONT_BUTTON,
                bg="#059669", fg="#FFFFFF", padx=10, pady=3, relief="groove", cursor="hand2",
                command=lambda aid=app_id, c=item['company']: self.on_quick_status(aid, "ได้งานแล้ว", c)
            )
            btn_acc.pack(side="left", padx=4)
            
            btn_del = tk.Button(
                btn_row, text="🗑️ ลบ", font=FONT_BUTTON,
                bg="#DC2626", fg="#FFFFFF", padx=10, pady=3, relief="groove", cursor="hand2",
                command=lambda aid=app_id, c=item['company']: self.on_delete(aid, c)
            )
            btn_del.pack(side="right", padx=4)

    # ==========================================================================
    # TAB 3: MANAGE ALL RECORDS
    # ==========================================================================
    def build_manage_tab(self):
        """สร้างแท็บรายการทั้งหมด: แถบเครื่องมือ (ค้นหา/กรอง/ปุ่ม) และตารางข้อมูล"""
        # แถบเครื่องมือด้านบน: ค้นหา ตัวกรอง และปุ่มต่างๆ
        tool_frame = tk.Frame(self.tab_manage, bg="#FFFFFF", padx=14, pady=10, bd=1, relief="solid")
        tool_frame.configure(highlightbackground="#E2E8F0", highlightthickness=1)
        tool_frame.pack(fill="x", padx=14, pady=(12, 10))
        
        tk.Label(tool_frame, text="🔍 ค้นหา:", font=FONT_BODY_BOLD, bg="#FFFFFF").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(tool_frame, textvariable=self.search_var, width=18)
        self.search_entry.pack(side="left", padx=(4, 12))
        self.search_entry.bind("<KeyRelease>", lambda e: self.filter_records())  # กรองข้อมูลทันทีที่พิมพ์
        
        tk.Label(tool_frame, text="สถานะ:", font=FONT_BODY_BOLD, bg="#FFFFFF").pack(side="left")
        self.filter_status_var = tk.StringVar(value="ทั้งหมด")
        self.cb_status = ttk.Combobox(tool_frame, textvariable=self.filter_status_var, 
                                      values=["ทั้งหมด"] + STATUS_OPTIONS, width=14, state="readonly")
        self.cb_status.pack(side="left", padx=(4, 12))
        self.cb_status.bind("<<ComboboxSelected>>", lambda e: self.filter_records())  # กรองใหม่เมื่อเลือกค่าใน dropdown
        
        tk.Label(tool_frame, text="ช่องทาง:", font=FONT_BODY_BOLD, bg="#FFFFFF").pack(side="left")
        self.filter_channel_var = tk.StringVar(value="ทั้งหมด")
        self.cb_channel = ttk.Combobox(tool_frame, textvariable=self.filter_channel_var, 
                                       values=["ทั้งหมด"] + CHANNEL_OPTIONS, width=12, state="readonly")
        self.cb_channel.pack(side="left", padx=(4, 12))
        self.cb_channel.bind("<<ComboboxSelected>>", lambda e: self.filter_records())  # กรองใหม่เมื่อเลือกค่าใน dropdown
        
        export_btn = tk.Button(
            tool_frame, text="📥 ส่งออก CSV", font=FONT_BUTTON,
            bg="#059669", fg="#FFFFFF", padx=10, pady=3, relief="groove", cursor="hand2",
            command=self.on_export_csv
        )
        export_btn.pack(side="right", padx=4)

        url_btn = tk.Button(
            tool_frame, text="🔗 เปิดลิงก์งาน", font=FONT_BODY,
            bg="#EEF2F6", fg="#2563EB", padx=8, pady=3, relief="groove", cursor="hand2",
            command=self.on_open_url
        )
        url_btn.pack(side="right", padx=4)
        
        edit_btn = tk.Button(
            tool_frame, text="✏️ แก้ไข", font=FONT_BUTTON,
            bg="#2563EB", fg="#FFFFFF", padx=10, pady=3, relief="groove", cursor="hand2",
            command=self.on_edit_selected
        )
        edit_btn.pack(side="right", padx=4)
        
        del_btn = tk.Button(
            tool_frame, text="🗑️ ลบ", font=FONT_BUTTON,
            bg="#EF4444", fg="#FFFFFF", padx=10, pady=3, relief="groove", cursor="hand2",
            command=self.on_delete_selected
        )
        del_btn.pack(side="right", padx=4)

        # ตารางข้อมูล Treeview
        table_frame = tk.Frame(self.tab_manage, bg="#FFFFFF", bd=1, relief="solid")
        table_frame.configure(highlightbackground="#E2E8F0", highlightthickness=1)
        table_frame.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        
        # ชื่อคอลัมน์ของตาราง (ชื่อภายใน)
        cols = ("id", "company", "position", "applied_date", "days_waiting", "channel", "status", "job_url", "notes")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        
        # ตั้งหัวคอลัมน์ที่ผู้ใช้เห็น
        self.tree.heading("id", text="ID")
        self.tree.heading("company", text="ชื่อบริษัท")
        self.tree.heading("position", text="ตำแหน่งงาน")
        self.tree.heading("applied_date", text="วันที่สมัคร")
        self.tree.heading("days_waiting", text="รอผล")
        self.tree.heading("channel", text="ช่องทาง")
        self.tree.heading("status", text="สถานะ")
        self.tree.heading("job_url", text="ลิงก์งาน")
        self.tree.heading("notes", text="หมายเหตุ")
        
        # ตั้งความกว้างและการจัดตำแหน่งของแต่ละคอลัมน์ (anchor: w=ชิดซ้าย, center=กึ่งกลาง)
        self.tree.column("id", width=45, anchor="center")
        self.tree.column("company", width=150, anchor="w")
        self.tree.column("position", width=170, anchor="w")
        self.tree.column("applied_date", width=100, anchor="center")
        self.tree.column("days_waiting", width=85, anchor="center")
        self.tree.column("channel", width=105, anchor="center")
        self.tree.column("status", width=125, anchor="center")
        self.tree.column("job_url", width=160, anchor="w")
        self.tree.column("notes", width=220, anchor="w")
        
        # แถบเลื่อนแนวตั้งและแนวนอน ผูกกับตาราง
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # จัดวางด้วย grid ให้ตารางขยายเต็มพื้นที่ โดยแถบเลื่อนอยู่ที่ขอบ
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        
        self.tree.bind("<Double-1>", lambda e: self.on_edit_selected())  # ดับเบิลคลิกแถวเพื่อเปิดหน้าต่างแก้ไข
        self.setup_treeview_sorting()

        # สีพื้นแถวตามสถานะ: แดง = รอเกิน 30 วัน, เหลือง = เกิน 14 วัน, เขียว = ได้งานแล้ว
        self.tree.tag_configure("critical", background="#FEE2E2")
        self.tree.tag_configure("warning", background="#FEF3C7")
        self.tree.tag_configure("accepted", background="#D1FAE5")

    def setup_treeview_sorting(self):
        """ผูกการคลิกหัวคอลัมน์เข้ากับการเรียงข้อมูล"""
        for col in self.tree["columns"]:
            self.tree.heading(col, command=lambda _c=col: self.sort_treeview_column(_c, False))

    def sort_treeview_column(self, col, reverse):
        """เรียงตารางตามคอลัมน์ที่คลิก คลิกซ้ำเพื่อสลับ น้อย→มาก / มาก→น้อย"""
        items = []
        for child in self.tree.get_children(""):
            val = self.tree.set(child, col)
            # แปลงเป็นตัวเลขก่อนเรียง ไม่เช่นนั้นจะเรียงแบบข้อความ (เช่น 10 มาก่อน 2)
            if col == "id":
                try:
                    val = int(val)
                except ValueError:
                    pass
            elif col == "days_waiting":
                try:
                    val = int(str(val).replace(" วัน", "").strip())
                except ValueError:
                    pass
            items.append((val, child))

        items.sort(reverse=reverse)
        for index, (_, child) in enumerate(items):
            self.tree.move(child, "", index)

        self.tree.heading(col, command=lambda _c=col: self.sort_treeview_column(_c, not reverse))  # คลิกครั้งถัดไปให้เรียงกลับทิศ

    def filter_records(self):
        """กรองข้อมูลตามคำค้น + สถานะ + ช่องทาง แล้วแสดงในตาราง

        ใช้ข้อมูลที่แคชไว้ใน all_cached_rows จึงไม่ต้องอ่านฐานข้อมูลใหม่ทุกครั้งที่กรอง
        """
        q = self.search_var.get().strip().lower()
        stat_filter = self.filter_status_var.get()
        chan_filter = self.filter_channel_var.get()
        
        # ล้างตารางก่อนใส่ข้อมูลที่กรองแล้ว
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for r in self.all_cached_rows:
            # ข้ามแถวที่ไม่ตรงกับเงื่อนไขกรอง (ค้นหาจากชื่อบริษัทหรือตำแหน่ง ไม่สนตัวพิมพ์เล็ก/ใหญ่)
            if q and (q not in r['company'].lower() and q not in r['position'].lower()):
                continue
            if stat_filter != "ทั้งหมด" and r['status'] != stat_filter:
                continue
            if chan_filter != "ทั้งหมด" and r['channel'] != chan_filter:
                continue
                
            days = calculate_days_waiting(r['applied_date'])
            # กำหนดสีแถวตามสถานะและจำนวนวันที่รอ
            tag = ""
            if r['status'] == "กำลังรอผล":
                if days > 30:
                    tag = "critical"
                elif days > 14:
                    tag = "warning"
            elif r['status'] == "ได้งานแล้ว":
                tag = "accepted"
                
            # เพิ่มแถวลงตาราง (tags ใช้เลือกสีที่ตั้งไว้ใน tag_configure)
            self.tree.insert("", "end", values=(
                r['id'], r['company'], r['position'], r['applied_date'], 
                f"{days} วัน", r['channel'], r['status'], r['job_url'], r['notes']
            ), tags=(tag,))

    # ==========================================================================
    # TAB 4: ADD APPLICATION
    # ==========================================================================
    def build_add_tab(self):
        """สร้างฟอร์มบันทึกใบสมัครใหม่ (ช่องที่มี * ต้องกรอก)"""
        outer_frame = tk.Frame(self.tab_add, bg="#F8FAFC", padx=20, pady=16)
        outer_frame.pack(fill="both", expand=True)

        form_frame = tk.Frame(outer_frame, bg="#FFFFFF", padx=28, pady=22, bd=1, relief="solid")
        form_frame.configure(highlightbackground="#E2E8F0", highlightthickness=1)
        form_frame.pack(fill="x", padx=10, pady=6)
        
        tk.Label(form_frame, text="➕ บันทึกข้อมูลการสมัครงาน", font=FONT_HEADING, fg="#0F172A", bg="#FFFFFF").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 16)
        )
        
        # ชื่อบริษัท
        tk.Label(form_frame, text="ชื่อบริษัท *", font=FONT_BODY_BOLD, bg="#FFFFFF").grid(row=1, column=0, sticky="w", pady=6)
        self.entry_add_company = ttk.Entry(form_frame, width=45)
        self.entry_add_company.grid(row=1, column=1, sticky="w", pady=6)
        
        # ตำแหน่งงาน
        tk.Label(form_frame, text="ตำแหน่งงาน *", font=FONT_BODY_BOLD, bg="#FFFFFF").grid(row=2, column=0, sticky="w", pady=6)
        self.entry_add_pos = ttk.Entry(form_frame, width=45)
        self.entry_add_pos.grid(row=2, column=1, sticky="w", pady=6)
        
        # วันที่สมัคร
        tk.Label(form_frame, text="วันที่สมัคร *", font=FONT_BODY_BOLD, bg="#FFFFFF").grid(row=3, column=0, sticky="w", pady=6)
        date_box = tk.Frame(form_frame, bg="#FFFFFF")
        date_box.grid(row=3, column=1, sticky="w", pady=6)
        
        self.entry_add_date = DatePicker(date_box, bg="#FFFFFF")
        self.entry_add_date.pack(side="left")
        tk.Label(date_box, text="(กดเพื่อเลือกวันที่)", font=FONT_SMALL, fg="#64748B", bg="#FFFFFF").pack(side="left", padx=8)
        
        # ช่องทาง
        tk.Label(form_frame, text="ช่องทางรับสมัคร", font=FONT_BODY_BOLD, bg="#FFFFFF").grid(row=4, column=0, sticky="w", pady=6)
        self.cb_add_channel = ttk.Combobox(form_frame, values=CHANNEL_OPTIONS, width=28, state="readonly")
        self.cb_add_channel.set("LinkedIn")
        self.cb_add_channel.grid(row=4, column=1, sticky="w", pady=6)
        
        # สถานะเริ่มต้น
        tk.Label(form_frame, text="สถานะเริ่มต้น", font=FONT_BODY_BOLD, bg="#FFFFFF").grid(row=5, column=0, sticky="w", pady=6)
        self.cb_add_status = ttk.Combobox(form_frame, values=STATUS_OPTIONS, width=28, state="readonly")
        self.cb_add_status.set("กำลังรอผล")
        self.cb_add_status.grid(row=5, column=1, sticky="w", pady=6)
        
        # ลิงก์ประกาศงาน
        tk.Label(form_frame, text="ลิงก์ประกาศงาน", font=FONT_BODY_BOLD, bg="#FFFFFF").grid(row=6, column=0, sticky="w", pady=6)
        self.entry_add_url = ttk.Entry(form_frame, width=45)
        self.entry_add_url.grid(row=6, column=1, sticky="w", pady=6)
        
        # หมายเหตุ
        tk.Label(form_frame, text="บันทึกเพิ่มเติม", font=FONT_BODY_BOLD, bg="#FFFFFF").grid(row=7, column=0, sticky="nw", pady=6)
        self.txt_add_notes = tk.Text(form_frame, width=45, height=4, font=FONT_BODY)
        self.txt_add_notes.grid(row=7, column=1, sticky="w", pady=6)
        
        # ปุ่มดำเนินการ
        # ปุ่มบันทึกและปุ่มล้างฟอร์ม
        btn_box = tk.Frame(form_frame, bg="#FFFFFF")
        btn_box.grid(row=8, column=1, sticky="w", pady=(16, 0))
        
        save_btn = tk.Button(
            btn_box, text="💾 บันทึกข้อมูล", font=FONT_BUTTON,
            bg="#2563EB", fg="#FFFFFF", padx=16, pady=6, relief="groove", cursor="hand2",
            command=self.on_save_add
        )
        save_btn.pack(side="left", padx=(0, 10))
        
        clear_btn = tk.Button(
            btn_box, text="ล้างข้อมูล", font=FONT_BODY,
            bg="#F1F5F9", fg="#475569", padx=12, pady=6, relief="groove", cursor="hand2",
            command=self.clear_add_form
        )
        clear_btn.pack(side="left")

    def on_save_add(self):
        """กดบันทึกในฟอร์ม: ตรวจข้อมูล -> บันทึกลงฐานข้อมูล -> ล้างฟอร์ม -> ไปหน้ารายการทั้งหมด"""
        company = self.entry_add_company.get().strip()
        position = self.entry_add_pos.get().strip()
        applied_date = self.entry_add_date.get().strip()
        channel = self.cb_add_channel.get().strip()
        status = self.cb_add_status.get().strip()
        job_url = self.entry_add_url.get().strip()
        notes = self.txt_add_notes.get("1.0", "end").strip()
        
        # ตรวจว่ากรอกช่องบังคับ (*) ครบหรือไม่
        if not company or not position:
            messagebox.showwarning("แจ้งเตือน", "กรุณาระบุชื่อบริษัทและตำแหน่งงาน")
            return
            
        try:
            datetime.strptime(applied_date, "%Y-%m-%d")  # ตรวจรูปแบบวันที่ (ถ้าผิดจะเกิด ValueError)
        except ValueError:
            messagebox.showerror("แจ้งเตือน", "กรุณากรอกวันที่ในรูปแบบ YYYY-MM-DD")
            return
            
        # บันทึกลงฐานข้อมูล แล้วแจ้งผล ล้างฟอร์ม และรีเฟรชหน้าจอ
        db_create(company, position, applied_date, channel, job_url, status, notes)
        messagebox.showinfo("สำเร็จ", f"บันทึกข้อมูลการสมัครงานบริษัท {company} เรียบร้อยแล้ว")
        self.clear_add_form()
        self.refresh_all_data()
        self.notebook.select(self.tab_manage)

    def clear_add_form(self):
        """ล้างฟอร์มกลับเป็นค่าเริ่มต้น (วันที่ = วันนี้, ช่องทาง = LinkedIn, สถานะ = กำลังรอผล)"""
        self.entry_add_company.delete(0, "end")
        self.entry_add_pos.delete(0, "end")
        self.entry_add_date.set(date.today())
        self.cb_add_channel.set("LinkedIn")
        self.cb_add_status.set("กำลังรอผล")
        self.entry_add_url.delete(0, "end")
        self.txt_add_notes.delete("1.0", "end")

    # ==========================================================================
    # ACTION HANDLERS
    # ==========================================================================
    def refresh_all_data(self):
        """ดึงข้อมูลล่าสุดจากฐานข้อมูล แล้วอัปเดตทั้ง 3 หน้า (แดชบอร์ด / ติดตามผล / ตาราง)"""
        self.all_cached_rows = db_read_all()
        self.render_dashboard(self.all_cached_rows)
        self.render_alerts(self.all_cached_rows)
        self.filter_records()

    def on_seed_demo(self):
        """กดปุ่มข้อมูลตัวอย่าง: ใส่ข้อมูลตัวอย่างลงฐานข้อมูล (เฉพาะเมื่อยังว่าง)"""
        db_seed_sample_data()
        messagebox.showinfo("ข้อมูลตัวอย่าง", "สร้างข้อมูลตัวอย่างเรียบร้อยแล้ว")
        self.refresh_all_data()

    def on_quick_reapply(self, app_id, current_notes, company):
        """ปุ่มลัด "ติดตามผลซ้ำ" ในหน้าติดตามผล"""
        db_reapply_followup(app_id, current_notes)
        messagebox.showinfo("สำเร็จ", f"รีเซ็ตวันที่สมัครบริษัท {company} เป็นวันปัจจุบันแล้ว")
        self.refresh_all_data()

    def on_quick_status(self, app_id, new_status, company):
        """ปุ่มลัดเปลี่ยนสถานะในหน้าติดตามผล (นัดสัมภาษณ์ / ได้งานแล้ว)"""
        db_update_status(app_id, new_status)
        messagebox.showinfo("สำเร็จ", f"เปลี่ยนสถานะบริษัท {company} เป็น '{new_status}' แล้ว")
        self.refresh_all_data()

    def on_delete(self, app_id, company):
        """ลบรายการ โดยถามยืนยันก่อนเสมอ"""
        if messagebox.askyesno("ยืนยันการลบ", f"ต้องการลบรายการของบริษัท '{company}' หรือไม่?"):
            db_delete(app_id)
            messagebox.showinfo("สำเร็จ", "ลบข้อมูลเรียบร้อยแล้ว")
            self.refresh_all_data()

    def on_delete_selected(self):
        """ลบรายการที่เลือกอยู่ในตาราง"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกรายการที่ต้องการลบในตาราง")
            return
        # อ่านค่าของแถวที่เลือก: vals[0] = id, vals[1] = ชื่อบริษัท
        vals = self.tree.item(selected[0], "values")
        app_id = vals[0]
        company = vals[1]
        self.on_delete(app_id, company)

    def on_open_url(self):
        """เปิดลิงก์ประกาศงานของแถวที่เลือกในเบราว์เซอร์"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกรายการที่ต้องการเปิดลิงก์")
            return
        vals = self.tree.item(selected[0], "values")
        url = vals[7]
        if not url:
            messagebox.showinfo("แจ้งเตือน", "รายการนี้ไม่ได้ระบุลิงก์ประกาศงาน")
            return
        webbrowser.open(url)

    def on_edit_selected(self):
        """เปิดหน้าต่างแก้ไขข้อมูลของแถวที่เลือก (ดับเบิลคลิกแถวก็เรียกฟังก์ชันนี้)"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("แจ้งเตือน", "กรุณาเลือกรายการที่ต้องการแก้ไข")
            return
            
        app_id = self.tree.item(selected[0], "values")[0]
        row = db_read_one(app_id)
        if not row:
            return

        # สร้างหน้าต่างแก้ไข (Toplevel) แบบ modal: grab_set() ล็อกไม่ให้กดหน้าต่างหลักระหว่างแก้
        dialog = tk.Toplevel(self)
        dialog.title(f"แก้ไขข้อมูล: {row['company']}")
        dialog.transient(self)
        dialog.grab_set()
        dialog.configure(bg="#F8FAFC", padx=22, pady=18)

        # คำนวณตำแหน่งให้หน้าต่างแก้ไขอยู่กลางหน้าต่างหลัก
        d_w, d_h = 520, 480
        self.update_idletasks()
        d_x = self.winfo_x() + max(0, (self.winfo_width() - d_w) // 2)
        d_y = self.winfo_y() + max(0, (self.winfo_height() - d_h) // 2)
        dialog.geometry(f"{d_w}x{d_h}+{d_x}+{d_y}")

        tk.Label(dialog, text=f"✏️ แก้ไขข้อมูล (ID #{app_id})", font=FONT_HEADING, bg="#F8FAFC").pack(anchor="w", pady=(0, 12))

        form = tk.Frame(dialog, bg="#F8FAFC")
        form.pack(fill="both", expand=True)

        tk.Label(form, text="ชื่อบริษัท:", font=FONT_BODY_BOLD, bg="#F8FAFC").grid(row=0, column=0, sticky="w", pady=4)
        # ฟอร์มแก้ไข: ใส่ค่าเดิมของรายการไว้ในแต่ละช่อง
        e_company = ttk.Entry(form, width=32)
        e_company.insert(0, row['company'])
        e_company.grid(row=0, column=1, sticky="w", pady=4)

        tk.Label(form, text="ตำแหน่งงาน:", font=FONT_BODY_BOLD, bg="#F8FAFC").grid(row=1, column=0, sticky="w", pady=4)
        e_pos = ttk.Entry(form, width=32)
        e_pos.insert(0, row['position'])
        e_pos.grid(row=1, column=1, sticky="w", pady=4)

        tk.Label(form, text="วันที่สมัคร:", font=FONT_BODY_BOLD, bg="#F8FAFC").grid(row=2, column=0, sticky="w", pady=4)
        e_date = DatePicker(form, initial=row['applied_date'], bg="#F8FAFC")
        e_date.grid(row=2, column=1, sticky="w", pady=4)

        tk.Label(form, text="ช่องทาง:", font=FONT_BODY_BOLD, bg="#F8FAFC").grid(row=3, column=0, sticky="w", pady=4)
        cb_channel = ttk.Combobox(form, values=CHANNEL_OPTIONS, width=25, state="readonly")
        cb_channel.set(row['channel'])
        cb_channel.grid(row=3, column=1, sticky="w", pady=4)

        tk.Label(form, text="สถานะ:", font=FONT_BODY_BOLD, bg="#F8FAFC").grid(row=4, column=0, sticky="w", pady=4)
        cb_status = ttk.Combobox(form, values=STATUS_OPTIONS, width=25, state="readonly")
        cb_status.set(row['status'])
        cb_status.grid(row=4, column=1, sticky="w", pady=4)

        tk.Label(form, text="ลิงก์งาน:", font=FONT_BODY_BOLD, bg="#F8FAFC").grid(row=5, column=0, sticky="w", pady=4)
        e_url = ttk.Entry(form, width=32)
        e_url.insert(0, row['job_url'] or "")
        e_url.grid(row=5, column=1, sticky="w", pady=4)

        tk.Label(form, text="หมายเหตุ:", font=FONT_BODY_BOLD, bg="#F8FAFC").grid(row=6, column=0, sticky="nw", pady=4)
        txt_notes = tk.Text(form, width=32, height=3, font=FONT_BODY)
        txt_notes.insert("1.0", row['notes'] or "")
        txt_notes.grid(row=6, column=1, sticky="w", pady=4)

        def save_edit():
            """ตรวจข้อมูลที่แก้ แล้วบันทึกลงฐานข้อมูล ปิดหน้าต่าง และรีเฟรชหน้าจอ"""
            new_comp = e_company.get().strip()
            new_p = e_pos.get().strip()
            new_d = e_date.get().strip()
            new_ch = cb_channel.get().strip()
            new_st = cb_status.get().strip()
            new_u = e_url.get().strip()
            new_n = txt_notes.get("1.0", "end").strip()

            if not new_comp or not new_p:
                messagebox.showwarning("แจ้งเตือน", "กรุณากรอกชื่อบริษัทและตำแหน่งงาน")
                return
            try:
                datetime.strptime(new_d, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("แจ้งเตือน", "กรุณากรอกวันที่ในรูปแบบ YYYY-MM-DD")
                return

            db_update(app_id, new_comp, new_p, new_d, new_ch, new_u, new_st, new_n)
            dialog.destroy()
            messagebox.showinfo("สำเร็จ", f"อัปเดตข้อมูล {new_comp} เรียบร้อยแล้ว")
            self.refresh_all_data()

        # ปุ่มบันทึกและปุ่มยกเลิก
        btn_box = tk.Frame(dialog, bg="#F8FAFC")
        btn_box.pack(fill="x", pady=(14, 0))
        
        tk.Button(
            btn_box, text="💾 บันทึก", font=FONT_BUTTON,
            bg="#2563EB", fg="#FFFFFF", padx=14, pady=4, relief="groove", cursor="hand2",
            command=save_edit
        ).pack(side="right", padx=4)
        
        tk.Button(
            btn_box, text="ยกเลิก", font=FONT_BODY,
            bg="#E2E8F0", fg="#334155", padx=10, pady=4, relief="groove", cursor="hand2",
            command=dialog.destroy
        ).pack(side="right", padx=4)

    def on_export_csv(self):
        """ส่งออกข้อมูลทั้งหมดเป็นไฟล์ CSV (เปิดใน Excel ได้)"""
        if not self.all_cached_rows:
            messagebox.showwarning("แจ้งเตือน", "ไม่มีข้อมูลสำหรับส่งออก")
            return
            
        # เปิดหน้าต่างให้ผู้ใช้เลือกตำแหน่งและชื่อไฟล์ที่จะบันทึก
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files (Excel)", "*.csv"), ("All Files", "*.*")],
            initialfile=f"job_applications_{date.today().strftime('%Y%m%d')}.csv",
            title="บันทึกไฟล์ CSV"
        )
        # ผู้ใช้กดยกเลิก -> ไม่ต้องทำอะไรต่อ
        if not file_path:
            return

        try:
            # encoding utf-8-sig ใส่ BOM นำหน้าไฟล์ เพื่อให้ Excel อ่านภาษาไทยไม่เพี้ยน
            with open(file_path, mode="w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["ID", "ชื่อบริษัท", "ตำแหน่งงาน", "วันที่สมัคร", "ระยะเวลารอ (วัน)", 
                                 "ช่องทาง", "สถานะ", "ลิงก์ประกาศงาน", "หมายเหตุ", "อัปเดตล่าสุด"])
                for r in self.all_cached_rows:
                    days = calculate_days_waiting(r['applied_date'])
                    writer.writerow([
                        r['id'], r['company'], r['position'], r['applied_date'], 
                        days, r['channel'], r['status'], r['job_url'], r['notes'], r['updated_at']
                    ])
            messagebox.showinfo("สำเร็จ", f"ส่งออกข้อมูลเป็นไฟล์ CSV เรียบร้อยแล้ว:\n{file_path}")
        except Exception as e:
            messagebox.showerror("เกิดข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ได้: {e}")

# ------------------------------------------------------------------------------
# Entry Point
# ------------------------------------------------------------------------------
# ทำงานเฉพาะเมื่อรันไฟล์นี้โดยตรง (ถ้าถูก import จากไฟล์อื่นจะไม่เปิดโปรแกรม)
if __name__ == "__main__":
    app = JobTrackerApp()
    app.mainloop()  # เริ่มวนลูปรอรับการกระทำของผู้ใช้ (โปรแกรมค้างอยู่ตรงนี้จนปิดหน้าต่าง)
import tkinter as tk
from tkinter import messagebox
from database import get_user_by_username, add_user, verify_password, get_all_users

class LoginWindow(tk.Toplevel):
    def __init__(self, parent):
        print("DEBUG: داخل LoginWindow __init__ (قبل از super().__init__).") # این خط باید باشد
        super().__init__(parent)
        self.parent = parent
        self.user_id = None
        self.user_role = None
        
        print("DEBUG: LoginWindow __init__ آغاز شد (بعد از super().__init__).") # این خط را اضافه کنید
        
        self.title("ورود به سیستم")
        self.geometry("300x200")
        self.resizable(False, False)
        
        # --- Make this window modal and transient ---
        self.grab_set()        # تمام ورودی‌ها را به این پنجره محدود می‌کند
        self.transient(parent) # این پنجره را موقت برای پنجره والد قرار می‌دهد
        
        self.create_widgets()
        self.center_window()
        self.protocol("WM_DELETE_WINDOW", self.on_closing) # مدیریت دکمه بستن پنجره

        # --- اطمینان از نمایش و فوکوس پنجره قبل از بلاک کردن والد ---
        self.update_idletasks() # پردازش تمام رویدادهای در حال انتظار برای رندر پنجره
        self.deiconify()        # اطمینان از اینکه پنجره Toplevel قابل مشاهده است
        self.lift()             # آوردن پنجره به بالای همه پنجره‌ها
        self.focus_set()        # فوکوس دادن به پنجره لاگین
        print("DEBUG: مؤلفه‌های گرافیکی LoginWindow ایجاد و قابل مشاهده شدند.") # این خط را اضافه کنید
        # --- پایان بخش اطمینان ---

        # بررسی وجود کاربران، در صورت عدم وجود، ایجاد ادمین اولیه را پیشنهاد می‌دهد
        # از after استفاده می‌شود تا اطمینان حاصل شود پنجره قبل از نمایش messagebox آماده است
        self.after(100, self.check_and_create_initial_admin) 

        # --- این خط مهم باعث می‌شود پنجره اصلی منتظر بسته شدن این پنجره بماند ---
        self.parent.wait_window(self) 
        print("DEBUG: LoginWindow.wait_window به پایان رسید.") # این خط را اضافه کنید

    def center_window(self):
        """پنجره Toplevel را نسبت به والد خود در مرکز قرار می‌دهد."""
        self.update_idletasks() # اطمینان از محاسبه ابعاد پنجره
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def create_widgets(self):
        """ویجت‌های فرم ورود را ایجاد می‌کند."""
        # نام کاربری
        username_label = tk.Label(self, text="نام کاربری:")
        username_label.pack(pady=5)
        self.username_entry = tk.Entry(self, width=30)
        self.username_entry.pack(pady=2)

        # رمز عبور
        password_label = tk.Label(self, text="رمز عبور:")
        password_label.pack(pady=5)
        self.password_entry = tk.Entry(self, width=30, show="*")
        self.password_entry.pack(pady=2)

        # دکمه ورود
        login_button = tk.Button(self, text="ورود", command=self.attempt_login)
        login_button.pack(pady=10)

        # اتصال کلید Enter به ورود
        self.username_entry.bind("<Return>", lambda event: self.password_entry.focus_set())
        self.password_entry.bind("<Return>", lambda event: self.attempt_login())
        self.username_entry.focus_set() # تنظیم فوکوس اولیه

    def attempt_login(self):
        """تلاش برای ورود کاربر با اطلاعات وارد شده."""
        print("DEBUG: تابع attempt_login در LoginWindow فراخوانی شد.") # این خط را اضافه کنید
        username = self.username_entry.get()
        password = self.password_entry.get()

        user_data = get_user_by_username(username)

        if user_data and verify_password(user_data['password_hash'], password):
            print("DEBUG: ورود موفقیت‌آمیز. در حال بستن پنجره ورود.") # این خط را اضافه کنید
            messagebox.showinfo("ورود موفق", "با موفقیت وارد شدید!", parent=self)
            self.user_id = user_data['id']
            self.user_role = user_data['role']
            self.destroy() # پنجره ورود را می‌بندد
        else:
            print("DEBUG: ورود ناموفق.") # این خط را اضافه کنید
            messagebox.showerror("خطای ورود", "نام کاربری یا رمز عبور اشتباه است.", parent=self)

    def check_and_create_initial_admin(self):
        """بررسی وجود کاربران و پیشنهاد ایجاد ادمین اولیه در صورت عدم وجود."""
        users = get_all_users()
        if not users:
            response = messagebox.askyesno(
                "ایجاد کاربر ادمین",
                "هیچ کاربری در سیستم وجود ندارد. آیا مایلید یک کاربر ادمین ایجاد کنید؟\n"
                "این کاربر به صورت پیش فرض نام کاربری 'admin' و رمز عبور 'admin123' خواهد داشت.",
                parent=self
            )
            if response:
                if add_user("admin", "admin123", "admin"):
                    messagebox.showinfo("کاربر ادمین ایجاد شد", "کاربر ادمین با نام کاربری 'admin' و رمز عبور 'admin123' ایجاد شد. لطفاً با آن وارد شوید.", parent=self)
                else:
                    messagebox.showerror("خطا", "خطا در ایجاد کاربر ادمین.", parent=self)
            else:
                messagebox.showwarning("هشدار", "بدون کاربر ادمین، ممکن است نتوانید به تمام قابلیت‌ها دسترسی داشته باشید.", parent=self)
        
        self.username_entry.focus_set()

    def on_closing(self):
        print("DEBUG: تابع on_closing در LoginWindow فراخوانی شد (دکمه X).") # این خط را اضافه کنید
        if messagebox.askokcancel("خروج", "آیا مطمئن هستید که می‌خواهید خارج شوید؟", parent=self):
            print("DEBUG: بستن پنجره ورود تایید شد.") # این خط را اضافه کنید
            self.destroy()
        else:
            print("DEBUG: بستن پنجره ورود لغو شد.") # این خط را اضافه کنید
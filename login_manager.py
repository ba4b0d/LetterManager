import tkinter as tk
from tkinter import messagebox, simpledialog, Toplevel, Label, Entry, Button, ttk
import os 
import hashlib 

# وارد کردن توابع مدیریت کاربر از database.py
from database import get_user_by_username, add_user, verify_password, get_all_users, update_user, update_user_password, delete_user, get_user_by_id

class LoginWindow(tk.Toplevel):
    def __init__(self, parent):
        print("DEBUG: داخل LoginWindow __init__ (قبل از super().__init__).")
        super().__init__(parent)
        self.parent = parent 
        self.user_id = None
        self.user_role = None
        
        print("DEBUG: LoginWindow __init__ آغاز شد (بعد از super().__init__).")
        
        self.title("ورود به سیستم")
        self.geometry("300x200")
        self.resizable(False, False)
        
        self.grab_set()        
        self.transient(parent) 
        
        self._center_window()
        self._create_login_widgets()
        self.protocol("WM_DELETE_WINDOW", self._on_closing) 

        self.update_idletasks() 
        self.deiconify()        
        self.lift()             
        self.focus_set()        
        self.update()           
        print("DEBUG: مؤلفه‌های گرافیکی LoginWindow ایجاد و قابل مشاهده شدند.") 
        self.username_entry.focus_set() 

        self.parent.wait_window(self) 

    def _center_window(self):
        self.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def _hash_password(self, password):
        """Hashes a password using SHA256. This is a private method for LoginWindow."""
        return hashlib.sha256(password.encode()).hexdigest()

    def _create_login_widgets(self):
        print("DEBUG: _create_login_widgets آغاز شد.") 
        tk.Label(self, text="نام کاربری:").pack(pady=5)
        self.username_entry = tk.Entry(self)
        self.username_entry.pack(pady=5)
        self.username_entry.focus_set()

        tk.Label(self, text="رمز عبور:").pack(pady=5)
        self.password_entry = tk.Entry(self, show="*")
        self.password_entry.pack(pady=5)

        tk.Button(self, text="ورود", command=self._authenticate_user).pack(pady=10)

        self.bind("<Return>", lambda event: self._authenticate_user())
        print("DEBUG: _create_login_widgets پایان یافت.") 

    def _authenticate_user(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        
        if verify_password(username, password):
            user = get_user_by_username(username)
            self.user_id = user['id']
            self.user_role = user['role']
            self.destroy() 
        else:
            messagebox.showerror("خطای ورود", "نام کاربری یا رمز عبور اشتباه است.", parent=self)

    def _on_closing(self):
        print("DEBUG: تابع _on_closing در LoginWindow فراخوانی شد (دکمه X).")
        if messagebox.askokcancel("خروج", "آیا می‌خواهید از برنامه خارج شوید؟", parent=self):
            self.parent.destroy() 
        else:
            pass

    # =========================================================
    # NEW FUNCTIONALITY: User Management for Admin 
    # =========================================================
    def open_user_management_window(self):
        if self.user_role != 'admin': 
            messagebox.showwarning("عدم دسترسی", "شما اجازه دسترسی به این بخش را ندارید.", parent=self.parent)
            return

        user_mgmt_window = Toplevel(self.parent) 
        user_mgmt_window.title("مدیریت کاربران")
        user_mgmt_window.geometry("500x400")
        self.parent.update_idletasks() 
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (user_mgmt_window.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (user_mgmt_window.winfo_height() // 2)
        user_mgmt_window.geometry(f"+{x}+{y}")

        user_mgmt_window.grab_set() 

        button_frame = ttk.Frame(user_mgmt_window)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="افزودن کاربر جدید", command=self._add_user_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="ویرایش کاربر انتخاب شده", command=self._edit_user_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="حذف کاربر انتخاب شده", command=self._delete_user_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="تغییر رمز عبور کاربر انتخاب شده", command=self._change_user_password_admin_dialog).pack(side=tk.LEFT, padx=5)

        tree_frame = ttk.Frame(user_mgmt_window)
        tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.user_treeview = ttk.Treeview(tree_frame, columns=("id", "username", "role"), show="headings", yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.user_treeview.yview)

        self.user_treeview.heading("id", text="شناسه")
        self.user_treeview.heading("username", text="نام کاربری")
        self.user_treeview.heading("role", text="نقش")

        self.user_treeview.column("id", width=50, stretch=tk.NO)
        self.user_treeview.column("username", width=150)
        self.user_treeview.column("role", width=100)
        
        self.user_treeview.pack(expand=True, fill="both")
        
        self._populate_user_treeview()

        user_mgmt_window.wait_window(user_mgmt_window) 

    def _populate_user_treeview(self):
        for i in self.user_treeview.get_children():
            self.user_treeview.delete(i)

        users = get_all_users()
        for user in users:
            self.user_treeview.insert("", tk.END, values=(user['id'], user['username'], user['role']))

    def _add_user_dialog(self):
        new_username = simpledialog.askstring("افزودن کاربر", "نام کاربری جدید را وارد کنید:", parent=self.parent) 
        if new_username:
            new_password = simpledialog.askstring("افزودن کاربر", f"رمز عبور برای کاربر '{new_username}' را وارد کنید:", show='*', parent=self.parent) 
            if new_password:
                new_role = simpledialog.askstring("افزودن کاربر", f"نقش کاربر '{new_username}' را وارد کنید (user یا admin):", initialvalue="user", parent=self.parent) 
                if new_role and new_role.lower() in ['user', 'admin']:
                    if add_user(new_username, new_password, new_role.lower()): 
                        messagebox.showinfo("موفقیت", f"کاربر '{new_username}' با موفقیت اضافه شد.", parent=self.parent) 
                        self._populate_user_treeview()
                else:
                    messagebox.showwarning("ورودی نامعتبر", "نقش وارد شده نامعتبر است. (فقط 'user' یا 'admin')", parent=self.parent) 

    def _edit_user_dialog(self):
        selected_item = self.user_treeview.focus()
        if not selected_item:
            messagebox.showwarning("انتخاب کاربر", "لطفاً کاربری را برای ویرایش انتخاب کنید.", parent=self.parent) 
            return
        
        values = self.user_treeview.item(selected_item, 'values')
        user_id = values[0]
        current_username = values[1]
        current_role = values[2]

        new_username = simpledialog.askstring("ویرایش کاربر", "نام کاربری جدید را وارد کنید:", initialvalue=current_username, parent=self.parent) 
        if new_username:
            new_role = simpledialog.askstring("ویرایش کاربر", "نقش جدید را وارد کنید (user یا admin):", initialvalue=current_role, parent=self.parent) 
            if new_role and new_role.lower() in ['user', 'admin']:
                if update_user(user_id, new_username, new_role.lower()):
                    messagebox.showinfo("موفقیت", f"کاربر '{new_username}' با موفقیت ویرایش شد.", parent=self.parent) 
                    self._populate_user_treeview()
            else:
                messagebox.showwarning("ورودی نامعتبر", "نقش وارد شده نامعتبر است. (فقط 'user' یا 'admin')", parent=self.parent) 

    def _delete_user_dialog(self):
        selected_item = self.user_treeview.focus()
        if not selected_item:
            messagebox.showwarning("انتخاب کاربر", "لطفاً کاربری را برای حذف انتخاب کنید.", parent=self.parent) 
            return

        values = self.user_treeview.item(selected_item, 'values')
        user_id = values[0]
        username_to_delete = values[1]

        if username_to_delete == 'admin':
            messagebox.showerror("خطا", "کاربر 'admin' را نمی‌توان حذف کرد.", parent=self.parent) 
            return
        
        if messagebox.askyesno("حذف کاربر", f"آیا مطمئن هستید که می‌خواهید کاربر '{username_to_delete}' را حذف کنید؟", parent=self.parent): 
            if delete_user(user_id):
                messagebox.showinfo("موفقیت", f"کاربر '{username_to_delete}' با موفقیت حذف شد.", parent=self.parent) 
                self._populate_user_treeview()

    def _change_user_password_admin_dialog(self):
        selected_item = self.user_treeview.focus()
        if not selected_item:
            messagebox.showwarning("انتخاب کاربر", "لطفاً کاربری را برای تغییر رمز عبور انتخاب کنید.", parent=self.parent) 
            return
        
        values = self.user_treeview.item(selected_item, 'values')
        user_id = values[0]
        username = values[1]

        new_password = simpledialog.askstring("تغییر رمز عبور", f"رمز عبور جدید برای کاربر '{username}' را وارد کنید:", show='*', parent=self.parent) 
        if new_password:
            confirm_password = simpledialog.askstring("تغییر رمز عبور", "رمز عبور جدید را مجدداً وارد کنید:", show='*', parent=self.parent)
            if new_password == confirm_password:
                if update_user_password(user_id, new_password): 
                    messagebox.showinfo("موفقیت", f"رمز عبور کاربر '{username}' با موفقیت تغییر یافت.", parent=self.parent) 
            else:
                messagebox.showerror("خطا", "رمز عبور جدید و تکرار آن یکسان نیستند.", parent=self.parent)
        else:
            messagebox.showwarning("ورودی نامعتبر", "رمز عبور جدید نمی‌تواند خالی باشد.", parent=self.parent) 

    def change_my_password(self):
        if self.user_id is None:
            messagebox.showwarning("ورود الزامی", "لطفاً ابتدا وارد سیستم شوید.", parent=self.parent)
            return

        current_password = simpledialog.askstring("تغییر رمز عبور", "رمز عبور فعلی خود را وارد کنید:", show='*', parent=self.parent)
        if not current_password:
            return

        if verify_password(get_user_by_id(self.user_id)['username'], current_password):
            new_password = simpledialog.askstring("تغییر رمز عبور", "رمز عبور جدید را وارد کنید:", show='*', parent=self.parent)
            if new_password:
                confirm_password = simpledialog.askstring("تغییر رمز عبور", "رمز عبور جدید را مجدداً وارد کنید:", show='*', parent=self.parent)
                if new_password == confirm_password:
                    if update_user_password(self.user_id, new_password): 
                        messagebox.showinfo("موفقیت", "رمز عبور شما با موفقیت تغییر یافت.", parent=self.parent)
                else:
                    messagebox.showerror("خطا", "رمز عبور جدید و تکرار آن یکسان نیستند.", parent=self.parent)
            else:
                messagebox.showwarning("ورودی نامعتبر", "رمز عبور جدید نمی‌تواند خالی باشد.", parent=self.parent)
        else:
            messagebox.showerror("خطا", "رمز عبور فعلی اشتباه است.", parent=self.parent)

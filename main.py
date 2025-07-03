import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import os
import shutil
from docx import Document
import jdatetime
import sqlite3
from ttkthemes import ThemedTk
from datetime import datetime
import sys
import traceback # اضافه کردن این خط برای دریافت traceback کامل خطا

# Import logical modules
from database import create_tables, get_db_connection, insert_letter, get_letters_from_db, get_letter_by_code
from settings_manager import load_settings, save_settings, company_name, full_company_name, default_save_path, letterhead_template_path, set_default_settings
from crm_logic import populate_organizations_treeview, populate_contacts_treeview, on_add_organization_button, on_edit_organization_button, on_delete_organization_button, on_add_contact_button, on_edit_contact_button, on_delete_contact_button, on_organization_select
from letter_generation_logic import on_generate_letter, generate_letter_number
from archive_logic import update_history_treeview, on_search_archive_button, on_open_letter_button
from helpers import convert_numbers_to_persian, replace_text_in_docx, show_progress_window, hide_progress_window, sort_column
from login_manager import LoginWindow

# --- Global Configurations (can be loaded from settings_manager) ---
BASE_FONT = ("Arial", 10)

class App:
    def __init__(self, root, user_id, user_role):
        self.root = root
        self.root.title("مدیریت ارتباط با مشتری و نامه‌نگاری")
        self.root.geometry("1000x700")
        self.user_id = user_id     # NEW: Store logged-in user's ID
        self.user_role = user_role # NEW: Store logged-in user's role

        # Apply a theme - Reverted to 'clam' or you can set to your preferred theme like 'arc'
        self.root.set_theme("clam") 
                                    

        # Initialize settings
        load_settings()

        # NEW: Disable settings tab for 'user' role
        if self.user_role == 'user':
            settings_tab_index = self.notebook.index(self.settings_frame)
            self.notebook.tab(settings_tab_index, state='disabled')
            # Optional: Hide it completely if desired
            # self.notebook.hide(settings_tab_index)
            messagebox.showinfo("دسترسی محدود", "شما به عنوان کاربر عادی وارد شدید. دسترسی به تنظیمات محدود است.")

        # Ensure correct initial paths are set up or default to current directory
        global default_save_path, letterhead_template_path
        if not default_save_path:
            default_save_path = os.path.join(os.path.expanduser("~"), "Documents", "GeneratedLetters")
            if not os.path.exists(default_save_path):
                os.makedirs(default_save_path)
            save_settings()
        
        # --- Status bar (INITIALIZED EARLY) ---
        self.status_bar = tk.Label(root, text="آماده به کار", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=BASE_FONT)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Initialize database tables
        create_tables()

        # --- UI Setup ---
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # Tabs
        self.tab_crm = ttk.Frame(self.notebook)
        self.tab_letter = ttk.Frame(self.notebook)
        self.tab_archive = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_letter, text="تولید نامه")
        self.notebook.add(self.tab_archive, text="آرشیو نامه‌ها")
        self.notebook.add(self.tab_crm, text="مدیریت مشتریان و مخاطبین")
        self.notebook.add(self.tab_settings, text="تنظیمات")

        # --- CRM Tab (tab_crm) ---
        self._setup_crm_tab()

        # --- Letter Generation Tab (tab_letter) ---
        self._setup_letter_tab()

        # --- Archive Tab (tab_archive) ---
        self._setup_archive_tab()

        # --- Settings Tab (tab_settings) ---
        self._setup_settings_tab()

        # Initial data population
        # populate_org_contact_combos() no longer directly relevant to letter tab, but still used for CRM comboboxes
        # This will be called from _on_tab_change for CRM tab, if needed
        self.update_history_treeview()

        # Bind tab change event to refresh data
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)

        # Variables to store selected organization/contact from dialogs
        self.selected_org_id = None
        self.selected_org_name = None
        self.selected_contact_id = None
        self.selected_contact_name = None


    # --- Helper methods for progress bar ---
    def show_progress(self, message="در حال پردازش..."):
        show_progress_window(message, self.root)

    def hide_progress(self):
        hide_progress_window()

    # --- CRM Tab Setup ---
    def _setup_crm_tab(self):
        # Organizations section
        org_frame = ttk.LabelFrame(self.tab_crm, text="سازمان‌ها", padding="10 10 10 10")
        org_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # Search and Add/Edit/Delete for Organizations
        org_search_frame = ttk.Frame(org_frame)
        org_search_frame.pack(fill=tk.X, pady=5)

        ttk.Label(org_search_frame, text="جستجوی سازمان:").pack(side=tk.RIGHT, padx=5)
        self.org_search_entry = ttk.Entry(org_search_frame)
        self.org_search_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        self.org_search_entry.bind("<Return>", lambda event: populate_organizations_treeview(self.org_search_entry.get(), self.org_treeview, self.status_bar))

        ttk.Button(org_search_frame, text="جستجو", command=lambda: populate_organizations_treeview(self.org_search_entry.get(), self.org_treeview, self.status_bar)).pack(side=tk.RIGHT, padx=5)

        org_buttons_frame = ttk.Frame(org_frame)
        org_buttons_frame.pack(fill=tk.X, pady=5)
        ttk.Button(org_buttons_frame, text="افزودن سازمان", command=lambda: on_add_organization_button(self.root, self.org_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.RIGHT, padx=5)
        ttk.Button(org_buttons_frame, text="ویرایش سازمان", command=lambda: on_edit_organization_button(self.root, self.org_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.RIGHT, padx=5)
        ttk.Button(org_buttons_frame, text="حذف سازمان", command=lambda: on_delete_organization_button(self.root, self.org_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.RIGHT, padx=5)

        # Organizations Treeview
        org_tree_frame = ttk.Frame(org_frame)
        org_tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        org_scrollbar = ttk.Scrollbar(org_tree_frame, orient="vertical")
        org_scrollbar.pack(side=tk.RIGHT, fill=tk.Y) 

        self.org_treeview = ttk.Treeview(org_tree_frame, columns=("id", "name", "industry", "phone", "email", "address", "description"), show="headings", yscrollcommand=org_scrollbar.set)
        org_scrollbar.config(command=self.org_treeview.yview)

        # Define columns and headings
        self.org_treeview.heading("id", text="شناسه", command=lambda: sort_column(self.org_treeview, "id", False))
        self.org_treeview.heading("name", text="نام سازمان", command=lambda: sort_column(self.org_treeview, "name", False))
        self.org_treeview.heading("industry", text="صنعت", command=lambda: sort_column(self.org_treeview, "industry", False))
        self.org_treeview.heading("phone", text="تلفن", command=lambda: sort_column(self.org_treeview, "phone", False))
        self.org_treeview.heading("email", text="ایمیل", command=lambda: sort_column(self.org_treeview, "email", False))
        self.org_treeview.heading("address", text="آدرس", command=lambda: sort_column(self.org_treeview, "address", False))
        self.org_treeview.heading("description", text="توضیحات", command=lambda: sort_column(self.org_treeview, "description", False))

        # Set column widths (adjust as needed)
        self.org_treeview.column("id", width=30, stretch=tk.NO)
        self.org_treeview.column("name", width=120)
        self.org_treeview.column("industry", width=80)
        self.org_treeview.column("phone", width=80)
        self.org_treeview.column("email", width=100)
        self.org_treeview.column("address", width=150)
        self.org_treeview.column("description", width=150)

        self.org_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True) 
        self.org_treeview.bind("<<TreeviewSelect>>", lambda event: on_organization_select(event, self.org_treeview, self.contact_treeview, self.status_bar))

        populate_organizations_treeview(org_treeview_ref=self.org_treeview, status_bar_ref=self.status_bar)


        # Contacts section
        contact_frame = ttk.LabelFrame(self.tab_crm, text="مخاطبین", padding="10 10 10 10")
        contact_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        # Configure grid weights for self.tab_crm to ensure equal column sizing
        self.tab_crm.grid_columnconfigure(0, weight=1)
        self.tab_crm.grid_columnconfigure(1, weight=1)
        self.tab_crm.grid_rowconfigure(0, weight=1)

        # Search and Add/Edit/Delete for Contacts
        contact_search_frame = ttk.Frame(contact_frame)
        contact_search_frame.pack(fill=tk.X, pady=5)

        ttk.Label(contact_search_frame, text="جستجوی مخاطب:").pack(side=tk.RIGHT, padx=5)
        self.contact_search_entry = ttk.Entry(contact_search_frame)
        self.contact_search_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        self.contact_search_entry.bind("<Return>", lambda event: populate_contacts_treeview(None, self.contact_search_entry.get(), self.contact_treeview, self.status_bar))

        ttk.Button(contact_search_frame, text="جستجو", command=lambda: populate_contacts_treeview(None, self.contact_search_entry.get(), self.contact_treeview, self.status_bar)).pack(side=tk.RIGHT, padx=5)

        contact_buttons_frame = ttk.Frame(contact_frame)
        contact_buttons_frame.pack(fill=tk.X, pady=5)
        ttk.Button(contact_buttons_frame, text="افزودن مخاطب", command=lambda: on_add_contact_button(self.root, self.contact_treeview, self.org_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.RIGHT, padx=5)
        ttk.Button(contact_buttons_frame, text="ویرایش مخاطب", command=lambda: on_edit_contact_button(self.root, self.contact_treeview, self.org_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.RIGHT, padx=5)
        ttk.Button(contact_buttons_frame, text="حذف مخاطب", command=lambda: on_delete_contact_button(self.root, self.contact_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.RIGHT, padx=5)

        # Contacts Treeview
        contact_tree_frame = ttk.Frame(contact_frame)
        contact_tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        contact_scrollbar = ttk.Scrollbar(contact_tree_frame, orient="vertical")
        contact_scrollbar.pack(side=tk.RIGHT, fill=tk.Y) 

        self.contact_treeview = ttk.Treeview(contact_tree_frame, columns=("id", "organization_id", "first_name", "last_name", "title", "phone", "email", "notes"), show="headings", yscrollcommand=contact_scrollbar.set)
        contact_scrollbar.config(command=self.contact_treeview.yview)

        # Define columns and headings
        self.contact_treeview.heading("id", text="شناسه", command=lambda: sort_column(self.contact_treeview, "id", False))
        self.contact_treeview.heading("organization_id", text="شناسه سازمان", command=lambda: sort_column(self.contact_treeview, "organization_id", False))
        self.contact_treeview.heading("first_name", text="نام", command=lambda: sort_column(self.contact_treeview, "first_name", False))
        self.contact_treeview.heading("last_name", text="نام خانوادگی", command=lambda: sort_column(self.contact_treeview, "last_name", False))
        self.contact_treeview.heading("title", text="عنوان", command=lambda: sort_column(self.contact_treeview, "title", False))
        self.contact_treeview.heading("phone", text="تلفن", command=lambda: sort_column(self.contact_treeview, "phone", False))
        self.contact_treeview.heading("email", text="ایمیل", command=lambda: sort_column(self.contact_treeview, "email", False))
        self.contact_treeview.heading("notes", text="یادداشت‌ها", command=lambda: sort_column(self.contact_treeview, "notes", False))

        # Set column widths (adjust as needed)
        self.contact_treeview.column("id", width=30, stretch=tk.NO)
        self.contact_treeview.column("organization_id", width=60, stretch=tk.NO)
        self.contact_treeview.column("first_name", width=80)
        self.contact_treeview.column("last_name", width=100)
        self.contact_treeview.column("title", width=80)
        self.contact_treeview.column("phone", width=80)
        self.contact_treeview.column("email", width=100)
        self.contact_treeview.column("notes", width=120)

        self.contact_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True) 

        populate_contacts_treeview(contact_treeview_ref=self.contact_treeview, status_bar_ref=self.status_bar)

    # --- Letter Generation Tab Setup ---
    def _setup_letter_tab(self):
        letter_frame = ttk.LabelFrame(self.tab_letter, text="تولید نامه جدید", padding="10 10 10 10")
        letter_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Row 1: Organization/Contact Selection (Now with Entry and Button)
        org_contact_frame = ttk.Frame(letter_frame)
        org_contact_frame.pack(fill=tk.X, pady=5)

        # Organization Selection
        ttk.Label(org_contact_frame, text=":سازمان").pack(side=tk.RIGHT, padx=5)
        self.entry_org_letter = ttk.Entry(org_contact_frame, width=30, state="readonly") # To display selected org name
        self.entry_org_letter.pack(side=tk.RIGHT, padx=5, expand=True, fill=tk.X)
        self.btn_select_org = ttk.Button(org_contact_frame, text="انتخاب سازمان", command=self._open_org_selection_dialog)
        self.btn_select_org.pack(side=tk.RIGHT, padx=5)

        # Contact Selection
        ttk.Label(org_contact_frame, text=":مخاطب").pack(side=tk.RIGHT, padx=5)
        self.entry_contact_letter = ttk.Entry(org_contact_frame, width=30, state="readonly") # To display selected contact name
        self.entry_contact_letter.pack(side=tk.RIGHT, padx=5, expand=True, fill=tk.X)
        self.btn_select_contact = ttk.Button(org_contact_frame, text="انتخاب مخاطب", command=self._open_contact_selection_dialog) # This will be implemented in a future step
        self.btn_select_contact.pack(side=tk.RIGHT, padx=5)

        # Row 2: Letter Type
        letter_type_frame = ttk.Frame(letter_frame)
        letter_type_frame.pack(fill=tk.X, pady=5)

        ttk.Label(letter_type_frame, text=":نوع نامه").pack(side=tk.RIGHT, padx=5)
        self.letter_types = {
            "FIN": "مالی",
            "HR": "منابع انسانی",
            "GEN": "عمومی"
        }
        self.combo_letter_type = ttk.Combobox(letter_type_frame, values=list(self.letter_types.values()), state="readonly", width=20)
        self.combo_letter_type.set(self.letter_types["FIN"])
        self.combo_letter_type.pack(side=tk.RIGHT, padx=5)


        # Row 3: Subject
        subject_frame = ttk.Frame(letter_frame)
        subject_frame.pack(fill=tk.X, pady=5)
        ttk.Label(subject_frame, text=":موضوع نامه").pack(side=tk.RIGHT, padx=5)
        self.entry_subject = ttk.Entry(subject_frame)
        self.entry_subject.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        self.entry_subject.bind("<Button-3>", self._show_text_context_menu)


        # Row 4: Letter Body
        body_frame = ttk.Frame(letter_frame)
        body_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        ttk.Label(body_frame, text=":متن نامه").pack(side=tk.TOP, anchor=tk.E, padx=5)
        self.text_letter_body = tk.Text(body_frame, wrap="word", height=15, font=BASE_FONT)
        self.text_letter_body.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.text_letter_body.bind("<Button-3>", self._show_text_context_menu)

        # Paste button for letter body
        self.btn_paste_body = ttk.Button(body_frame, text="چسباندن متن (Paste)", command=self.paste_text_to_body)
        self.btn_paste_body.pack(pady=5, padx=5, anchor=tk.W) 


        # Generate Button
        ttk.Button(letter_frame, text="تولید نامه", command=lambda: on_generate_letter(self)).pack(pady=10)

    # Method to paste text into the letter body
    def paste_text_to_body(self):
        """Pastes text from the clipboard into the letter body text area."""
        try:
            clipboard_content = self.root.clipboard_get() 
            self.text_letter_body.insert(tk.END, clipboard_content) 
        except tk.TclError:
            messagebox.showwarning("خطا در چسباندن", "کلیپ‌بورد خالی است یا محتوای متنی ندارد.", parent=self.root)
        except Exception as e:
            messagebox.showerror("خطا", f"خطایی در چسباندن متن رخ داد: {e}", parent=self.root)

    # --- Modal Dialog for Organization Selection ---
    def _open_org_selection_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("انتخاب سازمان")
        dialog.transient(self.root) # Make it modal relative to root
        dialog.grab_set() # Grab all events until this window is destroyed
        dialog.geometry("600x400")

        # Center the dialog on the screen
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        # Frame for search and treeview within the dialog
        dialog_frame = ttk.Frame(dialog, padding="10")
        dialog_frame.pack(fill=tk.BOTH, expand=True)

        # Search bar
        search_frame = ttk.Frame(dialog_frame)
        search_frame.pack(fill=tk.X, pady=5)
        ttk.Label(search_frame, text="جستجو:").pack(side=tk.RIGHT, padx=5)
        self.org_dialog_search_entry = ttk.Entry(search_frame)
        self.org_dialog_search_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        self.org_dialog_search_entry.bind("<Return>", lambda event: self._populate_org_dialog_treeview(self.org_dialog_search_entry.get()))
        ttk.Button(search_frame, text="جستجو", command=lambda: self._populate_org_dialog_treeview(self.org_dialog_search_entry.get())).pack(side=tk.RIGHT, padx=5)

        # Treeview for organizations
        tree_frame = ttk.Frame(dialog_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # MODIFIED: Scrollbar to the RIGHT
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # MODIFIED: Treeview to the LEFT of scrollbar
        self.org_dialog_treeview = ttk.Treeview(tree_frame, columns=("id", "name", "industry"), show="headings", yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.config(command=self.org_dialog_treeview.yview)

        self.org_dialog_treeview.heading("id", text="شناسه")
        self.org_dialog_treeview.heading("name", text="نام سازمان")
        self.org_dialog_treeview.heading("industry", text="صنعت")
        self.org_dialog_treeview.column("id", width=50, stretch=tk.NO)
        self.org_dialog_treeview.column("name", width=200)
        self.org_dialog_treeview.column("industry", width=150)
        self.org_dialog_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True) 

        # Bind double-click to select
        self.org_dialog_treeview.bind("<Double-1>", lambda event: self._select_org_from_dialog(dialog))

        # Buttons at the bottom
        button_frame = ttk.Frame(dialog_frame)
        button_frame.pack(fill=tk.X, pady=5)
        # MODIFIED: Buttons packed from LEFT, order reversed for visual appearance
        ttk.Button(button_frame, text="انتخاب", command=lambda: self._select_org_from_dialog(dialog)).pack(side=tk.LEFT, padx=5) 
        ttk.Button(button_frame, text="لغو", command=dialog.destroy).pack(side=tk.LEFT, padx=5) 

        # Populate initial data in the dialog's treeview
        self._populate_org_dialog_treeview()

        # Handle dialog closing (e.g., via X button)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy) 

        self.root.wait_window(dialog) 

    def _populate_org_dialog_treeview(self, search_term=""):
        """Populates the organization selection dialog's treeview."""
        for i in self.org_dialog_treeview.get_children():
            self.org_dialog_treeview.delete(i)

        conn = get_db_connection()
        cursor = conn.cursor()
        query = "SELECT id, name, industry FROM Organizations"
        params = []
        if search_term:
            query += " WHERE name LIKE ?"
            params.append(f"%{search_term}%")
        query += " ORDER BY name"
        
        cursor.execute(query, params)
        organizations = cursor.fetchall()
        conn.close()

        for org in organizations:
            self.org_dialog_treeview.insert("", tk.END, values=(org['id'], org['name'], org['industry']))

    def _select_org_from_dialog(self, dialog):
        """Called when an organization is selected from the dialog."""
        selected_item = self.org_dialog_treeview.focus()
        if selected_item:
            values = self.org_dialog_treeview.item(selected_item, 'values')
            self.selected_org_id = values[0]
            self.selected_org_name = values[1]
            
            # Update the main entry field
            self.entry_org_letter.config(state="normal")
            self.entry_org_letter.delete(0, tk.END)
            self.entry_org_letter.insert(0, self.selected_org_name)
            self.entry_org_letter.config(state="readonly")
            
            # Clear previous contact selection as organization changed
            self.selected_contact_id = None
            self.selected_contact_name = None
            self.entry_contact_letter.config(state="normal")
            self.entry_contact_letter.delete(0, tk.END)
            self.entry_contact_letter.config(state="readonly")

            dialog.destroy() 
        else:
            messagebox.showwarning("انتخاب سازمان", "لطفاً یک سازمان را انتخاب کنید.", parent=dialog)

    # --- New: Modal Dialog for Contact Selection ---
    def _open_contact_selection_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("انتخاب مخاطب")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("700x450")

        # Center the dialog
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        dialog_frame = ttk.Frame(dialog, padding="10")
        dialog_frame.pack(fill=tk.BOTH, expand=True)

        # Search bar
        search_frame = ttk.Frame(dialog_frame)
        search_frame.pack(fill=tk.X, pady=5)
        ttk.Label(search_frame, text="جستجو:").pack(side=tk.RIGHT, padx=5)
        self.contact_dialog_search_entry = ttk.Entry(search_frame)
        self.contact_dialog_search_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        self.contact_dialog_search_entry.bind("<Return>", lambda event: self._populate_contact_dialog_treeview(self.contact_dialog_search_entry.get(), self.selected_org_id))
        ttk.Button(search_frame, text="جستجو", command=lambda: self._populate_contact_dialog_treeview(self.contact_dialog_search_entry.get(), self.selected_org_id)).pack(side=tk.RIGHT, padx=5)

        # Treeview for contacts
        tree_frame = ttk.Frame(dialog_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # MODIFIED: Scrollbar to the RIGHT
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # MODIFIED: Treeview to the LEFT of scrollbar
        self.contact_dialog_treeview = ttk.Treeview(tree_frame, columns=("id", "first_name", "last_name", "title", "org_name"), show="headings", yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.config(command=self.contact_dialog_treeview.yview)

        self.contact_dialog_treeview.heading("id", text="شناسه")
        self.contact_dialog_treeview.heading("first_name", text="نام")
        self.contact_dialog_treeview.heading("last_name", text="نام خانوادگی")
        self.contact_dialog_treeview.heading("title", text="عنوان")
        self.contact_dialog_treeview.heading("org_name", text="سازمان")
        self.contact_dialog_treeview.column("id", width=50, stretch=tk.NO)
        self.contact_dialog_treeview.column("first_name", width=100)
        self.contact_dialog_treeview.column("last_name", width=120)
        self.contact_dialog_treeview.column("title", width=100)
        self.contact_dialog_treeview.column("org_name", width=150)
        self.contact_dialog_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True) 

        # Bind double-click to select
        self.contact_dialog_treeview.bind("<Double-1>", lambda event: self._select_contact_from_dialog(dialog))

        # Buttons at the bottom
        button_frame = ttk.Frame(dialog_frame)
        button_frame.pack(fill=tk.X, pady=5)
        # MODIFIED: Buttons packed from LEFT, order reversed for visual appearance
        ttk.Button(button_frame, text="انتخاب", command=lambda: self._select_contact_from_dialog(dialog)).pack(side=tk.LEFT, padx=5) 
        ttk.Button(button_frame, text="لغو", command=dialog.destroy).pack(side=tk.LEFT, padx=5) 


        # Populate initial data, potentially filtered by selected organization
        self._populate_contact_dialog_treeview(organization_id=self.selected_org_id)

        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        self.root.wait_window(dialog)

    def _populate_contact_dialog_treeview(self, search_term="", organization_id=None):
        """Populates the contact selection dialog's treeview, optionally filtered by search term and organization_id."""
        for i in self.contact_dialog_treeview.get_children():
            self.contact_dialog_treeview.delete(i)

        conn = get_db_connection()
        cursor = conn.cursor()
        query = """
            SELECT c.id, c.first_name, c.last_name, c.title, o.name AS org_name, c.organization_id
            FROM Contacts c
            LEFT JOIN Organizations o ON c.organization_id = o.id
        """
        params = []
        conditions = []

        if search_term:
            conditions.append("(c.first_name LIKE ? OR c.last_name LIKE ? OR c.title LIKE ? OR o.name LIKE ?)")
            params.extend([f"%{search_term}%", f"%{search_term}%", f"%{search_term}%", f"%{search_term}%"])
        
        if organization_id is not None:
            conditions.append("c.organization_id = ?")
            params.append(organization_id)
        
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        
        query += " ORDER BY c.last_name, c.first_name"
        
        cursor.execute(query, params)
        contacts = cursor.fetchall()
        conn.close()

        for contact in contacts:
            self.contact_dialog_treeview.insert("", tk.END, values=(
                contact['id'], 
                contact['first_name'], 
                contact['last_name'], 
                contact['title'], 
                contact['org_name'],
                contact['organization_id'] # Keep org_id for internal use if needed
            ))

    def _select_contact_from_dialog(self, dialog):
        """Called when a contact is selected from the dialog."""
        selected_item = self.contact_dialog_treeview.focus()
        if selected_item:
            values = self.contact_dialog_treeview.item(selected_item, 'values')
            self.selected_contact_id = values[0]
            self.selected_contact_name = f"{values[1]} {values[2]}" # First Name + Last Name
            
            # Ensure the correct organization is also set if not already
            # (This handles cases where a contact is selected without an org first)
            selected_contact_org_id = values[5] # The organization_id is the 6th value in the 'values' tuple
            if self.selected_org_id is None or self.selected_org_id != selected_contact_org_id:
                # If org wasn't set or is different, try to set it
                conn = get_db_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM Organizations WHERE id = ?", (selected_contact_org_id,))
                org_name = cursor.fetchone()
                conn.close()
                if org_name:
                    self.selected_org_id = selected_contact_org_id
                    self.selected_org_name = org_name['name']
                    self.entry_org_letter.config(state="normal")
                    self.entry_org_letter.delete(0, tk.END)
                    self.entry_org_letter.insert(0, self.selected_org_name)
                    self.entry_org_letter.config(state="readonly")


            # Update the main contact entry field
            self.entry_contact_letter.config(state="normal")
            self.entry_contact_letter.delete(0, tk.END)
            self.entry_contact_letter.insert(0, self.selected_contact_name)
            self.entry_contact_letter.config(state="readonly")
            
            dialog.destroy()
        else:
            messagebox.showwarning("انتخاب مخاطب", "لطفاً یک مخاطب را انتخاب کنید.", parent=dialog)

    # This method is now primarily used for CRM tab comboboxes, not letter tab
    def populate_org_contact_combos(self):
        """Populates organization and contact comboboxes for CRM tab."""
        conn = get_db_connection()
        cursor = conn.cursor()

        # Populate Organization Combobox (for CRM tab)
        cursor.execute("SELECT id, name FROM Organizations ORDER BY name")
        organizations = cursor.fetchall()
        org_names = ["---"]
        self.org_data_map = {"---": None}
        for org in organizations:
            org_names.append(org['name'])
            self.org_data_map[org['name']] = org['id']
        # self.combo_org_letter['values'] = org_names # Removed from here, handled by dialog

        # Populate Contact Combobox (all contacts initially, for CRM tab)
        cursor.execute("SELECT id, first_name, last_name, organization_id, title FROM Contacts ORDER BY first_name, last_name")
        contacts = cursor.fetchall()
        contact_full_names = ["---"]
        self.all_contacts_data = {"---": None}
        for contact in contacts:
            full_name = f"{contact['first_name']} {contact['last_name']}"
            contact_full_names.append(full_name)
            self.all_contacts_data[full_name] = {
                'id': contact['id'],
                'first_name': contact['first_name'],
                'last_name': contact['last_name'],
                'organization_id': contact['organization_id'],
                'title': contact['title']
            }
        # self.combo_contact_letter['values'] = contact_full_names # Removed from here, handled by dialog

        conn.close()

    # _on_org_letter_select is no longer needed as combobox is removed
    # def _on_org_letter_select(self, event):
    #    pass # Logic moved to _select_org_from_dialog for selected_org_id and contact handling


    # --- Archive Tab Setup ---
    def _setup_archive_tab(self):
        archive_frame = ttk.LabelFrame(self.tab_archive, text="آرشیو نامه‌ها", padding="10 10 10 10")
        archive_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Search bar for archive
        archive_search_frame = ttk.Frame(archive_frame)
        archive_search_frame.pack(fill=tk.X, pady=5)

        ttk.Label(archive_search_frame, text="جستجو در آرشیو:").pack(side=tk.RIGHT, padx=5)
        self.archive_search_entry = ttk.Entry(archive_search_frame)
        self.archive_search_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        self.archive_search_entry.bind("<Return>", lambda event: on_search_archive_button(self.archive_search_entry, self.history_treeview, self.status_bar))
        ttk.Button(archive_search_frame, text="جستجو", command=lambda: on_search_archive_button(self.archive_search_entry, self.history_treeview, self.status_bar)).pack(side=tk.RIGHT, padx=5)

        # Treeview for letter history
        history_tree_frame = ttk.Frame(archive_frame)
        history_tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        history_scrollbar = ttk.Scrollbar(history_tree_frame, orient="vertical")
        history_scrollbar.pack(side=tk.RIGHT, fill=tk.Y) 

        self.history_treeview = ttk.Treeview(history_tree_frame, columns=("code", "type", "date", "subject", "organization", "contact", "CreatedBy"), show="headings", yscrollcommand=history_scrollbar.set)
        history_scrollbar.config(command=self.history_treeview.yview)

        # Define columns and headings for history
        self.history_treeview.heading("code", text="کد نامه", command=lambda: sort_column(self.history_treeview, "code", False))
        self.history_treeview.heading("type", text="نوع نامه", command=lambda: sort_column(self.history_treeview, "type", False))
        self.history_treeview.heading("date", text="تاریخ", command=lambda: sort_column(self.history_treeview, "date", False))
        self.history_treeview.heading("subject", text="موضوع", command=lambda: sort_column(self.history_treeview, "subject", False))
        self.history_treeview.heading("organization", text="سازمان", command=lambda: sort_column(self.history_treeview, "organization", False))
        self.history_treeview.heading("contact", text="مخاطب", command=lambda: sort_column(self.history_treeview, "contact", False))

        self.history_treeview.column("code", width=100, stretch=tk.NO)
        self.history_treeview.column("type", width=80, stretch=tk.NO)
        self.history_treeview.column("date", width=80, stretch=tk.NO)
        self.history_treeview.column("subject", width=200)
        self.history_treeview.column("organization", width=150)
        self.history_treeview.column("contact", width=150)

        self.history_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True) 

        # Open Letter Button
        open_letter_button = ttk.Button(archive_frame, text="باز کردن نامه", command=lambda: on_open_letter_button(self.history_treeview, self.root, self.status_bar))
        open_letter_button.pack(pady=10)

    # --- Settings Tab Setup ---
    def _setup_settings_tab(self):
        settings_frame = ttk.LabelFrame(self.tab_settings, text="تنظیمات برنامه", padding="10 10 10 10")
        settings_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Company Name (Abbreviation)
        ttk.Label(settings_frame, text="نام اختصاری شرکت (کد نامه‌):").grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        self.entry_company_name_abbr = ttk.Entry(settings_frame, width=50)
        self.entry_company_name_abbr.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        self.entry_company_name_abbr.insert(0, company_name)

        # Full Company Name
        ttk.Label(settings_frame, text="نام کامل شرکت (در پابرگ نامه):").grid(row=1, column=0, sticky=tk.E, padx=5, pady=5)
        self.entry_full_company_name = ttk.Entry(settings_frame, width=50)
        self.entry_full_company_name.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        self.entry_full_company_name.insert(0, full_company_name)

        # Default Save Path
        ttk.Label(settings_frame, text="مسیر پیش‌فرض ذخیره نامه‌ها:").grid(row=2, column=0, sticky=tk.E, padx=5, pady=5)
        self.entry_save_path = ttk.Entry(settings_frame, width=50)
        self.entry_save_path.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        self.entry_save_path.insert(0, default_save_path)
        ttk.Button(settings_frame, text="انتخاب مسیر", command=self._select_save_path).grid(row=2, column=2, padx=5, pady=5)

        # Letterhead Template Path
        ttk.Label(settings_frame, text="مسیر فایل الگوی سربرگ (Word):").grid(row=3, column=0, sticky=tk.E, padx=5, pady=5)
        self.entry_template_path = ttk.Entry(settings_frame, width=50)
        self.entry_template_path.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)
        self.entry_template_path.insert(0, letterhead_template_path)
        ttk.Button(settings_frame, text="انتخاب فایل", command=self._select_template_file).grid(row=3, column=2, padx=5, pady=5)

        # Save Settings Button
        ttk.Button(settings_frame, text="ذخیره تنظیمات", command=self._save_settings_from_ui).grid(row=4, column=1, columnspan=2, pady=20)


        # Configure column weights for resizing
        settings_frame.grid_columnconfigure(1, weight=1)

    def on_generate_letter_wrapper(self):
        """Wrapper method to collect data and call on_generate_letter."""
        self.show_progress("در حال تولید نامه...")
        try:
            # Collect all necessary data from the UI fields
            letter_type = self.letter_type_var.get()
            subject = self.subject_entry.get()
            body = self.letter_body_text.get("1.0", tk.END).strip()
            selected_org_id = self.selected_org_id # Already stored when selecting org
            selected_contact_id = self.selected_contact_id # Already stored when selecting contact

            if not letter_type:
                messagebox.showwarning("ورودی ناقص", "لطفاً نوع نامه را انتخاب کنید.")
                return
            if not subject:
                messagebox.showwarning("ورودی ناقص", "لطفاً موضوع نامه را وارد کنید.")
                return
            if not body:
                messagebox.showwarning("ورودی ناقص", "لطفاً متن نامه را وارد کنید.")
                return
            # Organization and Contact are optional depending on letter type
            # For now, we assume they are selected from the CRM tab and stored in self.selected_org_id/contact_id

            # Call the main letter generation logic, passing the user_id
            on_generate_letter(
                root_window_ref=self.root,
                status_bar_ref=self.status_bar,
                letter_type=letter_type,
                subject=subject,
                body_content=body,
                organization_id=selected_org_id,
                contact_id=selected_contact_id,
                save_path=default_save_path, # Using global default_save_path
                letterhead_template=letterhead_template_path, # Using global letterhead_template_path
                user_id=self.user_id # Pass the current user's ID
            )

            # Clear fields after successful generation
            self.letter_type_var.set("")
            self.subject_entry.delete(0, tk.END)
            self.letter_body_text.delete("1.0", tk.END)
            self.selected_org_id = None
            self.selected_org_name = None
            self.selected_contact_id = None
            self.selected_contact_name = None
            self.org_display_entry.delete(0, tk.END)
            self.contact_display_entry.delete(0, tk.END)

            self.update_history_treeview() # Refresh archive after new letter is generated

        except Exception as e:
            messagebox.showerror("خطا در تولید نامه", f"خطایی رخ داد: {e}")
        finally:
            self.hide_progress()


    def _select_save_path(self):
        """Opens a directory dialog to select the default save path."""
        folder_selected = filedialog.askdirectory(parent=self.root)
        if folder_selected:
            self.entry_save_path.delete(0, tk.END)
            self.entry_save_path.insert(0, folder_selected)

    def _select_template_file(self):
        """Opens a file dialog to select the letterhead template file."""
        file_selected = filedialog.askopenfilename(
            parent=self.root,
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if file_selected:
            self.entry_template_path.delete(0, tk.END)
            self.entry_template_path.insert(0, file_selected)

    def _save_settings_from_ui(self):
        """Saves settings from the UI fields to settings.txt."""
        global company_name, full_company_name, default_save_path, letterhead_template_path
        company_name = self.entry_company_name_abbr.get().strip()
        full_company_name = self.entry_full_company_name.get().strip()
        default_save_path = self.entry_save_path.get().strip()
        letterhead_template_path = self.entry_template_path.get().strip()
        save_settings()
        messagebox.showinfo("تنظیمات", "تنظیمات با موفقیت ذخیره شد.", parent=self.root)
        if self.status_bar: self.status_bar.config(text="تنظیمات ذخیره شد.")

    # Helper method for context menu - Re-added Paste command
    def _show_text_context_menu(self, event):
        """Displays a right-click context menu for Text widgets."""
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="بریدن", command=lambda: event.widget.event_generate("<<Cut>>"))
        context_menu.add_command(label="کپی", command=lambda: event.widget.event_generate("<<Copy>>"))
        context_menu.add_command(label="چسباندن", command=lambda: event.widget.event_generate("<<Paste>>")) 
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    # Public methods for external modules to call (if needed)
    def update_history_treeview(self, search_term="", treeview_widget=None, status_bar_ref=None):
        update_history_treeview(search_term, treeview_widget or self.history_treeview, status_bar_ref or self.status_bar)

    def _on_tab_change(self, event):
        """Handles actions when a notebook tab is changed."""
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        if selected_tab == "مدیریت مشتریان و مخاطبین":
            populate_organizations_treeview(org_treeview_ref=self.org_treeview, status_bar_ref=self.status_bar)
            populate_contacts_treeview(contact_treeview_ref=self.contact_treeview, status_bar_ref=self.status_bar)
        elif selected_tab == "تولید نامه":
            # No longer need to populate comboboxes here, as they are replaced by entry/button
            # self.populate_org_contact_combos() 
            pass
        elif selected_tab == "آرشیو نامه‌ها":
            self.update_history_treeview()


    # --- Global Error Handling ---
# (اطمینان حاصل کنید که 'import sys' و 'import traceback' در ابتدای فایل شما وجود دارند)
# ... (کدهای قبلی) ...

if __name__ == "__main__":
    try:
        print("DEBUG: تنظیمات اولیه برنامه آغاز شد.")
        create_tables() 
        print("DEBUG: جداول دیتابیس بررسی/ایجاد شدند.")

        root = ThemedTk(theme="cl")
        # root.withdraw() # <--- این خط را کامنت کنید (یک # در ابتدای آن بگذارید) یا حذف کنید
        print("DEBUG: پنجره اصلی Tkinter ایجاد شد.") # پیام را تغییر دادم

        print("DEBUG: در حال ایجاد پنجره ورود (LoginWindow)...")
        login_window = LoginWindow(root) 
        print("DEBUG: پنجره ورود بسته شد. در حال بازیابی وضعیت ورود.")

        logged_in_user_id = login_window.user_id
        logged_in_user_role = login_window.user_role
        
        print(f"DEBUG: نتیجه ورود - شناسه کاربر: {logged_in_user_id}, نقش: {logged_in_user_role}")

        if logged_in_user_id is not None:
            print("DEBUG: کاربر وارد شده است. نمایش پنجره اصلی و ایجاد نمونه برنامه.")
            root.deiconify() # این خط ممکن است دیگر لازم نباشد اما فعلا نگه دارید
            app = App(root, logged_in_user_id, logged_in_user_role)
            print("DEBUG: نمونه برنامه ایجاد شد. شروع حلقه اصلی (mainloop).")
            root.mainloop()
            print("DEBUG: حلقه اصلی به پایان رسید.")
        else:
            print("DEBUG: ورود لغو یا ناموفق بود. در حال خروج از برنامه.")
            root.destroy()
            sys.exit(0)

    except Exception as e:
        error_message = f"یک خطای غیرمنتظره رخ داد:\nنوع خطا: {type(e).__name__}\nپیام خطا: {e}\n\n"
        error_message += "جزئیات کامل خطا (Traceback):\n"
        error_message += traceback.format_exc()

        print("\n" + "*"*50)
        print("CRITICAL ERROR CAUGHT:")
        print(error_message)
        print("*"*50 + "\n")
        messagebox.showerror("خطای بحرانی برنامه", error_message + "\n\nبرنامه بسته خواهد شد.", parent=None)
        sys.exit(1)
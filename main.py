import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import os
import shutil
from docx import Document
import jdatetime
import sqlite3
from ttkthemes import ThemedTk
from datetime import datetime

# Import logical modules
from database import create_tables, get_db_connection, insert_letter, get_letters_from_db, get_letter_by_code
from settings_manager import load_settings, save_settings, company_name, full_company_name, default_save_path, letterhead_template_path, set_default_settings
from crm_logic import populate_organizations_treeview, populate_contacts_treeview, on_add_organization_button, on_edit_organization_button, on_delete_organization_button, on_add_contact_button, on_edit_contact_button, on_delete_contact_button, on_organization_select
from letter_generation_logic import on_generate_letter, generate_letter_number
from archive_logic import update_history_treeview, on_search_archive_button, on_open_letter_button
from helpers import convert_numbers_to_persian, replace_text_in_docx, show_progress_window, hide_progress_window, sort_column

# --- Global Configurations (can be loaded from settings_manager) ---
BASE_FONT = ("Arial", 10)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("مدیریت ارتباط با مشتری و نامه‌نگاری")
        self.root.geometry("1000x700")

        # Apply a theme - Reverted to 'clam' or you can set to your preferred theme like 'arc'
        self.root.set_theme("clam") # Changed back to 'clam' or your original preferred theme if it was 'arc'
                                    # You can try 'plastik', 'alt', 'breeze', 'arc', 'elegance' etc.

        # Initialize settings
        load_settings()

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

        self.notebook.add(self.tab_crm, text="مدیریت مشتریان و مخاطبین")
        self.notebook.add(self.tab_letter, text="تولید نامه")
        self.notebook.add(self.tab_archive, text="آرشیو نامه‌ها")
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
        self.populate_org_contact_combos()
        self.update_history_treeview()

        # Bind tab change event to refresh data
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)

    # --- Helper methods for progress bar ---
    def show_progress(self, message="در حال پردازش..."):
        show_progress_window(message, self.root)

    def hide_progress(self):
        hide_progress_window()

    # --- CRM Tab Setup ---
    def _setup_crm_tab(self):
        # Organizations section
        org_frame = ttk.LabelFrame(self.tab_crm, text="سازمان‌ها", padding="10 10 10 10")
        org_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)

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
        org_scrollbar.pack(side=tk.LEFT, fill=tk.Y)

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

        self.org_treeview.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.org_treeview.bind("<<TreeviewSelect>>", lambda event: on_organization_select(event, self.org_treeview, self.contact_treeview, self.status_bar))

        populate_organizations_treeview(org_treeview_ref=self.org_treeview, status_bar_ref=self.status_bar)


        # Contacts section
        contact_frame = ttk.LabelFrame(self.tab_crm, text="مخاطبین", padding="10 10 10 10")
        contact_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True, padx=10, pady=5)

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
        contact_scrollbar.pack(side=tk.LEFT, fill=tk.Y)

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

        self.contact_treeview.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        populate_contacts_treeview(contact_treeview_ref=self.contact_treeview, status_bar_ref=self.status_bar)

    # --- Letter Generation Tab Setup ---
    def _setup_letter_tab(self):
        letter_frame = ttk.LabelFrame(self.tab_letter, text="تولید نامه جدید", padding="10 10 10 10")
        letter_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Row 1: Organization/Contact Selection
        org_contact_frame = ttk.Frame(letter_frame)
        org_contact_frame.pack(fill=tk.X, pady=5)

        ttk.Label(org_contact_frame, text="سازمان:").pack(side=tk.RIGHT, padx=5)
        self.org_data_map = {}
        self.combo_org_letter = ttk.Combobox(org_contact_frame, state="readonly", width=30)
        self.combo_org_letter.pack(side=tk.RIGHT, padx=5, expand=True, fill=tk.X)
        self.combo_org_letter.bind("<<ComboboxSelected>>", self._on_org_letter_select)

        ttk.Label(org_contact_frame, text="مخاطب:").pack(side=tk.RIGHT, padx=5)
        self.all_contacts_data = {}
        self.combo_contact_letter = ttk.Combobox(org_contact_frame, state="readonly", width=30)
        self.combo_contact_letter.pack(side=tk.RIGHT, padx=5, expand=True, fill=tk.X)

        # Row 2: Letter Type
        letter_type_frame = ttk.Frame(letter_frame)
        letter_type_frame.pack(fill=tk.X, pady=5)

        ttk.Label(letter_type_frame, text="نوع نامه:").pack(side=tk.RIGHT, padx=5)
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
        ttk.Label(subject_frame, text="موضوع نامه:").pack(side=tk.RIGHT, padx=5)
        self.entry_subject = ttk.Entry(subject_frame)
        self.entry_subject.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        self.entry_subject.bind("<Button-3>", self._show_text_context_menu)


        # Row 4: Letter Body
        body_frame = ttk.Frame(letter_frame)
        body_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        ttk.Label(body_frame, text="متن نامه:").pack(side=tk.TOP, anchor=tk.E, padx=5)
        self.text_letter_body = tk.Text(body_frame, wrap="word", height=15, font=BASE_FONT)
        self.text_letter_body.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.text_letter_body.bind("<Button-3>", self._show_text_context_menu)

        # Generate Button
        ttk.Button(letter_frame, text="تولید نامه", command=lambda: on_generate_letter(self)).pack(pady=10)

    def populate_org_contact_combos(self):
        """Populates organization and contact comboboxes."""
        conn = get_db_connection()
        cursor = conn.cursor()

        # Populate Organization Combobox
        cursor.execute("SELECT id, name FROM Organizations ORDER BY name")
        organizations = cursor.fetchall()
        org_names = ["---"]
        self.org_data_map = {"---": None}
        for org in organizations:
            org_names.append(org['name'])
            self.org_data_map[org['name']] = org['id']
        self.combo_org_letter['values'] = org_names
        self.combo_org_letter.set("---")

        # Populate Contact Combobox (all contacts initially)
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
        self.combo_contact_letter['values'] = contact_full_names
        self.combo_contact_letter.set("---")

        conn.close()

    def _on_org_letter_select(self, event):
        """Filters contact combobox based on selected organization."""
        selected_org_name = self.combo_org_letter.get()
        selected_org_id = self.org_data_map.get(selected_org_name)

        filtered_contact_names = ["---"]
        self.combo_contact_letter.set("---")

        if selected_org_id is not None:
            for full_name, contact_data in self.all_contacts_data.items():
                if contact_data and contact_data['organization_id'] == selected_org_id:
                    filtered_contact_names.append(full_name)
        else:
            for full_name in self.all_contacts_data.keys():
                if full_name != "---":
                    filtered_contact_names.append(full_name)

        self.combo_contact_letter['values'] = filtered_contact_names


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
        history_scrollbar.pack(side=tk.LEFT, fill=tk.Y)

        self.history_treeview = ttk.Treeview(history_tree_frame, columns=("code", "type", "date", "subject", "organization", "contact"), show="headings", yscrollcommand=history_scrollbar.set)
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

        self.history_treeview.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

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
        context_menu.add_command(label="چسباندن", command=lambda: event.widget.event_generate("<<Paste>>")) # Added back paste
        
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
            self.populate_org_contact_combos()
        elif selected_tab == "آرشیو نامه‌ها":
            self.update_history_treeview()


if __name__ == "__main__":
    # Create database tables if they don't exist
    create_tables()
    root = ThemedTk()
    app = App(root)
    root.mainloop()
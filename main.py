import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import os
import shutil
# from docx import Document # This might not be strictly needed here if replace_text_in_docx handles it
# import jdatetime # This might not be strictly needed here if other modules handle it
# import sqlite3 # This might not be strictly needed here if database module handles it
from ttkthemes import ThemedTk # Ensure this is imported for themes
# from datetime import datetime # This might not be strictly needed here if other modules handle it

# Import logical modules
from database import create_tables, get_db_connection, insert_letter, get_letters_from_db, get_letter_by_code
# Updated import to include full_company_name and set_default_settings
from settings_manager import load_settings, save_settings, company_name, full_company_name, default_save_path, letterhead_template_path, set_default_settings 
from crm_logic import populate_organizations_treeview, populate_contacts_treeview, on_add_organization_button, on_edit_organization_button, on_delete_organization_button, on_add_contact_button, on_edit_contact_button, on_delete_contact_button, on_organization_select
# Import the updated on_generate_letter function
from letter_generation_logic import on_generate_letter 
from archive_logic import update_history_treeview, on_search_archive_button, on_open_letter_button
from helpers import convert_numbers_to_persian, replace_text_in_docx, show_progress_window, hide_progress_window, sort_column

# --- Global Configurations (can be loaded from settings_manager) ---
BASE_FONT = ("Arial", 10) # Define BASE_FONT as it's used by helpers

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("مدیریت ارتباط با مشتری و نامه‌نگاری")
        self.root.geometry("1000x700") # Adjusted initial size
        
        # Load settings immediately when the app starts
        load_settings() # This populates global variables from settings_manager

        # Initialize global references for helper functions (better to pass as arguments)
        # As per previous discussions, we are passing app_instance (self) where needed.
        # But ensure initial setup of progress window is handled.
        self.progress_window = None # Initialize here
        
        # --- Configure ttk styles ---
        s = ttk.Style()
        # Configure a generic button style using BASE_FONT
        s.configure('TButton', font=BASE_FONT)
        # Configure other ttk widgets if they also need BASE_FONT, e.g.:
        s.configure('TLabel', font=BASE_FONT)
        s.configure('TEntry', font=BASE_FONT)
        s.configure('TCombobox', font=BASE_FONT)
        s.configure('TNotebook.Tab', font=BASE_FONT) # For tab titles
        s.configure('TLabelframe.Label', font=BASE_FONT) # For LabelFrame titles
        # You might need to add more styles based on your UI elements if they don't inherit
        
        # --- Data holders for CRM tab ---
        self.org_data_map = {} # To map organization names to IDs
        self.all_contacts_data = {} # To map contact full names to their data (including ID)
        
        self.contact_search_term = "" # To hold the current search term for contacts

        # Create notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, expand=True, fill="both")

        # Create individual tabs
        self.tab_dashboard = ttk.Frame(self.notebook)
        self.tab_letter_generation = ttk.Frame(self.notebook)
        self.tab_crm = ttk.Frame(self.notebook)
        self.tab_archive = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_dashboard, text="داشبورد")
        self.notebook.add(self.tab_letter_generation, text="تولید نامه")
        self.notebook.add(self.tab_crm, text="CRM")
        self.notebook.add(self.tab_archive, text="آرشیو")
        self.notebook.add(self.tab_settings, text="تنظیمات")

        # Status Bar
        # tk.Label supports font directly
        self.status_bar = tk.Label(root, text="برنامه آماده است.", bd=1, relief=tk.SUNKEN, anchor=tk.W, font=BASE_FONT)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Initializing tabs calls their _create_ methods
        self._create_dashboard_tab()
        self._create_letter_generation_tab()
        self._create_crm_tab()
        self._create_archive_tab()
        self._create_settings_tab()
        
        # Initial population of CRM data and combos
        self.populate_org_contact_combos()
        populate_organizations_treeview(org_treeview_ref=self.org_treeview, status_bar_ref=self.status_bar)
        populate_contacts_treeview(contact_treeview_ref=self.contact_treeview, status_bar_ref=self.status_bar)
        self.update_history_treeview() # Call instance method which uses self.history_treeview
        
        # Bind CRM tab selection to refresh data
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)

        # Context menu for text widgets
        self.root.bind_class("Text", "<Button-3>", self._show_text_context_menu)
        self.root.bind_class("Entry", "<Button-3>", self._show_text_context_menu)


    def _on_tab_change(self, event):
        selected_tab = self.notebook.tab(self.notebook.select(), "text")
        if selected_tab == "CRM":
            self.populate_org_contact_combos()
            populate_organizations_treeview(org_treeview_ref=self.org_treeview, status_bar_ref=self.status_bar)
            populate_contacts_treeview(contact_treeview_ref=self.contact_treeview, status_bar_ref=self.status_bar)
        elif selected_tab == "آرشیو":
            self.update_history_treeview() # Refresh archive when tab is selected
        elif selected_tab == "تولید نامه":
            self.populate_org_contact_combos() # Refresh combos when letter generation tab is selected
        elif selected_tab == "تنظیمات":
            # Refresh settings fields when the settings tab is selected
            # This is important after a save or if settings.txt was changed externally
            load_settings() # Reload global settings
            self.entry_company_name.delete(0, tk.END)
            self.entry_company_name.insert(0, company_name)
            self.entry_full_company_name.delete(0, tk.END)
            self.entry_full_company_name.insert(0, full_company_name)
            self.entry_default_save_path.delete(0, tk.END)
            self.entry_default_save_path.insert(0, default_save_path)
            self.entry_letterhead_template.delete(0, tk.END)
            self.entry_letterhead_template.insert(0, letterhead_template_path)


    def show_progress(self, message):
        # FIX: Removed BASE_FONT, as show_progress_window in helpers.py doesn't expect it directly
        show_progress_window(self.root, message) 

    def hide_progress(self):
        hide_progress_window()


    def _create_dashboard_tab(self):
        # Dashboard content (can be expanded later)
        label = ttk.Label(self.tab_dashboard, text="به داشبورد خوش آمدید!", font=("Arial", 16))
        label.pack(pady=50)

    def _create_letter_generation_tab(self):
        # Frame for inputs
        input_frame = ttk.LabelFrame(self.tab_letter_generation, text="اطلاعات نامه", padding=(20, 10))
        input_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Subject
        ttk.Label(input_frame, text="موضوع نامه:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5) # Font handled by style
        self.entry_subject = ttk.Entry(input_frame, width=70) # Font handled by style
        self.entry_subject.grid(row=0, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=5, padx=5)

        # Organization Combobox
        ttk.Label(input_frame, text="انتخاب سازمان:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5) # Font handled by style
        self.combo_org_letter = ttk.Combobox(input_frame, width=30, state="readonly") # Font handled by style
        self.combo_org_letter.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.combo_org_letter.set("---") # Default value
        self.combo_org_letter.bind("<<ComboboxSelected>>", self._on_letter_org_select) # Bind selection event

        # Contact Combobox
        ttk.Label(input_frame, text="انتخاب مخاطب:").grid(row=1, column=2, sticky=tk.W, pady=5, padx=5) # Font handled by style
        self.combo_contact_letter = ttk.Combobox(input_frame, width=30, state="readonly") # Font handled by style
        self.combo_contact_letter.grid(row=1, column=3, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.combo_contact_letter.set("---") # Default value

        # Letter Type Combobox
        self.letter_types = {
            "FIN": "مالی",
            "ADM": "اداری",
            "HR": "پرسنلی",
            "GEN": "عمومی"
        }
        ttk.Label(input_frame, text="نوع نامه:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5) # Font handled by style
        self.combo_letter_type = ttk.Combobox(input_frame, width=30, state="readonly", 
                                              values=list(self.letter_types.values())) # Font handled by style
        self.combo_letter_type.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.combo_letter_type.set(self.letter_types["FIN"]) # Default letter type


        # Letter Body
        ttk.Label(input_frame, text="متن نامه:").grid(row=3, column=0, sticky=tk.NW, pady=5, padx=5) # Font handled by style
        self.text_letter_body = tk.Text(input_frame, width=70, height=15, font=BASE_FONT, wrap=tk.WORD) # tk.Text supports font directly
        self.text_letter_body.grid(row=3, column=1, columnspan=3, sticky="nsew", pady=5, padx=5)
        
        # Configure row and column weights for resizing
        input_frame.grid_rowconfigure(3, weight=1)
        input_frame.grid_columnconfigure(1, weight=1)
        input_frame.grid_columnconfigure(3, weight=1)

        # Generate Button - font is now handled by ttk.Style
        generate_button = ttk.Button(self.tab_letter_generation, text="تولید نامه", 
                                     command=lambda: on_generate_letter(self))
        generate_button.pack(pady=10)

    def _on_letter_org_select(self, event):
        """Updates contact combobox based on selected organization in letter generation tab."""
        selected_org_name = self.combo_org_letter.get()
        self.combo_contact_letter.set("---") # Reset contact combo

        if selected_org_name == "---":
            # If no organization selected, populate with all contacts
            contact_names = sorted(self.all_contacts_data.keys())
        else:
            # Filter contacts by selected organization
            selected_org_id = self.org_data_map.get(selected_org_name)
            if selected_org_id:
                contact_names = []
                for contact_name, contact_data in self.all_contacts_data.items():
                    if contact_data.get('organization_id') == selected_org_id:
                        contact_names.append(contact_name)
                contact_names.sort()
            else:
                contact_names = [] # Should not happen if data maps are correct
        
        self.combo_contact_letter['values'] = ["---"] + contact_names

    def populate_org_contact_combos(self):
        """Populates organization and contact comboboxes for letter generation tab."""
        conn = get_db_connection()
        cursor = conn.cursor()

        # Populate Organization Combobox
        cursor.execute("SELECT id, name FROM Organizations ORDER BY name")
        organizations = cursor.fetchall()
        self.org_data_map = {"---": None} # Add default "---" option
        org_names = ["---"]
        for org in organizations:
            self.org_data_map[org['name']] = org['id']
            org_names.append(org['name'])
        self.combo_org_letter['values'] = org_names

        # Populate Contact Combobox (all contacts initially)
        cursor.execute("SELECT id, first_name, last_name, organization_id FROM Contacts ORDER BY first_name, last_name")
        contacts = cursor.fetchall()
        self.all_contacts_data = {} # Clear existing data
        contact_full_names = ["---"]
        for contact in contacts:
            full_name = f"{contact['first_name']} {contact['last_name']}"
            self.all_contacts_data[full_name] = {
                'id': contact['id'],
                'first_name': contact['first_name'],
                'last_name': contact['last_name'],
                'organization_id': contact['organization_id']
            }
            contact_full_names.append(full_name)
        self.combo_contact_letter['values'] = contact_full_names
        
        conn.close()


    def _create_crm_tab(self):
        # CRM tab content
        crm_frame = ttk.LabelFrame(self.tab_crm, text="مدیریت سازمان‌ها و مخاطبین", padding=(20, 10))
        crm_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=20, pady=20) # Adjusted padx/pady

        # Organization Management Section
        org_frame = ttk.LabelFrame(crm_frame, text="سازمان‌ها", padding=(10, 5))
        org_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=5, pady=5)

        # Organization Search
        org_search_frame = ttk.Frame(org_frame)
        org_search_frame.pack(fill=tk.X, pady=5)
        self.entry_org_search = ttk.Entry(org_search_frame, width=30) # Font handled by style
        self.entry_org_search.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.entry_org_search.bind("<Return>", lambda event: populate_organizations_treeview(self.entry_org_search.get(), self.org_treeview, self.status_bar))
        ttk.Button(org_search_frame, text="جستجو", command=lambda: populate_organizations_treeview(self.entry_org_search.get(), self.org_treeview, self.status_bar)).pack(side=tk.LEFT, padx=5)

        # Organization Treeview
        self.org_treeview = ttk.Treeview(org_frame, columns=("ID", "Name", "Industry", "Phone", "Email", "Address", "Description"), show="headings", selectmode="browse")
        self.org_treeview.pack(fill="both", expand=True)

        # Define headings and column widths
        self.org_treeview.heading("ID", text="شناسه", command=lambda: sort_column(self.org_treeview, "ID", False))
        self.org_treeview.heading("Name", text="نام سازمان", command=lambda: sort_column(self.org_treeview, "Name", False))
        self.org_treeview.heading("Industry", text="صنعت", command=lambda: sort_column(self.org_treeview, "Industry", False))
        self.org_treeview.heading("Phone", text="تلفن", command=lambda: sort_column(self.org_treeview, "Phone", False))
        self.org_treeview.heading("Email", text="ایمیل", command=lambda: sort_column(self.org_treeview, "Email", False))
        self.org_treeview.heading("Address", text="آدرس", command=lambda: sort_column(self.org_treeview, "Address", False))
        self.org_treeview.heading("Description", text="توضیحات", command=lambda: sort_column(self.org_treeview, "Description", False))

        # Set column widths (example values, adjust as needed)
        self.org_treeview.column("ID", width=40, stretch=tk.NO)
        self.org_treeview.column("Name", width=120, stretch=tk.YES)
        self.org_treeview.column("Industry", width=80, stretch=tk.YES)
        self.org_treeview.column("Phone", width=90, stretch=tk.YES)
        self.org_treeview.column("Email", width=120, stretch=tk.YES)
        self.org_treeview.column("Address", width=150, stretch=tk.YES)
        self.org_treeview.column("Description", width=100, stretch=tk.YES)

        # Bind selection event to populate contacts
        self.org_treeview.bind("<<TreeviewSelect>>", lambda event: on_organization_select(event, self.org_treeview, self.contact_treeview, self.status_bar))
        
        # Organization Buttons
        org_button_frame = ttk.Frame(org_frame)
        org_button_frame.pack(pady=10)
        ttk.Button(org_button_frame, text="افزودن سازمان", command=lambda: on_add_organization_button(self.root, self.org_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.LEFT, padx=5)
        ttk.Button(org_button_frame, text="ویرایش سازمان", command=lambda: on_edit_organization_button(self.root, self.org_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.LEFT, padx=5)
        ttk.Button(org_button_frame, text="حذف سازمان", command=lambda: on_delete_organization_button(self.root, self.org_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.LEFT, padx=5)

        # Contact Management Section
        contact_frame = ttk.LabelFrame(crm_frame, text="مخاطبین", padding=(10, 5))
        contact_frame.pack(side=tk.RIGHT, fill="both", expand=True, padx=5, pady=5)

        # Contact Search
        contact_search_frame = ttk.Frame(contact_frame)
        contact_search_frame.pack(fill=tk.X, pady=5)
        self.entry_contact_search = ttk.Entry(contact_search_frame, width=30) # Font handled by style
        self.entry_contact_search.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.entry_contact_search.bind("<Return>", lambda event: populate_contacts_treeview(None, self.entry_contact_search.get(), self.contact_treeview, self.status_bar))
        ttk.Button(contact_search_frame, text="جستجو", command=lambda: populate_contacts_treeview(None, self.entry_contact_search.get(), self.contact_treeview, self.status_bar)).pack(side=tk.LEFT, padx=5)

        # Contact Treeview
        self.contact_treeview = ttk.Treeview(contact_frame, columns=("ID", "FirstName", "LastName", "Title", "Organization", "Phone", "Email", "Notes"), show="headings", selectmode="browse")
        self.contact_treeview.pack(fill="both", expand=True)

        # Define headings and column widths
        self.contact_treeview.heading("ID", text="شناسه", command=lambda: sort_column(self.contact_treeview, "ID", False))
        self.contact_treeview.heading("FirstName", text="نام", command=lambda: sort_column(self.contact_treeview, "FirstName", False))
        self.contact_treeview.heading("LastName", text="نام خانوادگی", command=lambda: sort_column(self.contact_treeview, "LastName", False))
        self.contact_treeview.heading("Title", text="عنوان", command=lambda: sort_column(self.contact_treeview, "Title", False))
        self.contact_treeview.heading("Organization", text="سازمان", command=lambda: sort_column(self.contact_treeview, "Organization", False))
        self.contact_treeview.heading("Phone", text="تلفن", command=lambda: sort_column(self.contact_treeview, "Phone", False))
        self.contact_treeview.heading("Email", text="ایمیل", command=lambda: sort_column(self.contact_treeview, "Email", False))
        self.contact_treeview.heading("Notes", text="یادداشت‌ها", command=lambda: sort_column(self.contact_treeview, "Notes", False))

        # Set column widths
        self.contact_treeview.column("ID", width=40, stretch=tk.NO)
        self.contact_treeview.column("FirstName", width=80, stretch=tk.YES)
        self.contact_treeview.column("LastName", width=100, stretch=tk.YES)
        self.contact_treeview.column("Title", width=80, stretch=tk.YES)
        self.contact_treeview.column("Organization", width=100, stretch=tk.YES)
        self.contact_treeview.column("Phone", width=90, stretch=tk.YES)
        self.contact_treeview.column("Email", width=120, stretch=tk.YES)
        self.contact_treeview.column("Notes", width=100, stretch=tk.YES)

        # Contact Buttons
        contact_button_frame = ttk.Frame(contact_frame)
        contact_button_frame.pack(pady=10)
        ttk.Button(contact_button_frame, text="افزودن مخاطب", command=lambda: on_add_contact_button(self.root, self.contact_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.LEFT, padx=5)
        ttk.Button(contact_button_frame, text="ویرایش مخاطب", command=lambda: on_edit_contact_button(self.root, self.contact_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.LEFT, padx=5)
        ttk.Button(contact_button_frame, text="حذف مخاطب", command=lambda: on_delete_contact_button(self.root, self.contact_treeview, self.status_bar, self.populate_org_contact_combos)).pack(side=tk.LEFT, padx=5)

    def _create_archive_tab(self):
        # Archive tab content
        archive_frame = ttk.LabelFrame(self.tab_archive, text="آرشیو نامه‌ها", padding=(20, 10))
        archive_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Search bar for archive
        search_frame = ttk.Frame(archive_frame)
        search_frame.pack(fill=tk.X, pady=5)
        self.entry_archive_search = ttk.Entry(search_frame, width=50) # Font handled by style
        self.entry_archive_search.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        self.entry_archive_search.bind("<Return>", lambda event: on_search_archive_button(self.entry_archive_search.get(), self.history_treeview, self.status_bar))
        ttk.Button(search_frame, text="جستجو در آرشیو", command=lambda: on_search_archive_button(self.entry_archive_search.get(), self.history_treeview, self.status_bar)).pack(side=tk.LEFT, padx=5)

        # History Treeview
        self.history_treeview = ttk.Treeview(archive_frame, columns=("Code", "Type", "Date", "Subject", "Organization", "Contact"), show="headings", selectmode="browse")
        self.history_treeview.pack(fill="both", expand=True)

        # Define headings
        self.history_treeview.heading("Code", text="شماره نامه", command=lambda: sort_column(self.history_treeview, "Code", False))
        self.history_treeview.heading("Type", text="نوع", command=lambda: sort_column(self.history_treeview, "Type", False))
        self.history_treeview.heading("Date", text="تاریخ", command=lambda: sort_column(self.history_treeview, "Date", False))
        self.history_treeview.heading("Subject", text="موضوع", command=lambda: sort_column(self.history_treeview, "Subject", False))
        self.history_treeview.heading("Organization", text="سازمان", command=lambda: sort_column(self.history_treeview, "Organization", False))
        self.history_treeview.heading("Contact", text="مخاطب", command=lambda: sort_column(self.history_treeview, "Contact", False))

        # Set column widths
        self.history_treeview.column("Code", width=120, stretch=tk.NO)
        self.history_treeview.column("Type", width=80, stretch=tk.NO)
        self.history_treeview.column("Date", width=100, stretch=tk.NO)
        self.history_treeview.column("Subject", width=250, stretch=tk.YES)
        self.history_treeview.column("Organization", width=150, stretch=tk.YES)
        self.history_treeview.column("Contact", width=150, stretch=tk.YES)
        
        # Open Letter Button
        open_button = ttk.Button(archive_frame, text="باز کردن نامه انتخاب شده", 
                                 command=lambda: on_open_letter_button(self.history_treeview, self.status_bar, self.root))
        open_button.pack(pady=10)


    def _create_settings_tab(self):
        settings_frame = ttk.LabelFrame(self.tab_settings, text="تنظیمات برنامه", padding=(20, 10))
        settings_frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Company Abbreviation
        ttk.Label(settings_frame, text="نام اختصاری شرکت (برای کد نامه):").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5) # Font handled by style
        self.entry_company_name = ttk.Entry(settings_frame, width=50) # Font handled by style
        self.entry_company_name.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.entry_company_name.insert(0, company_name) # Load current setting

        # Full Company Name (New)
        ttk.Label(settings_frame, text="نام کامل شرکت (برای پایین نامه):").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5) # Font handled by style
        self.entry_full_company_name = ttk.Entry(settings_frame, width=50) # Font handled by style
        self.entry_full_company_name.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.entry_full_company_name.insert(0, full_company_name) # Load current setting

        # Default Save Path
        ttk.Label(settings_frame, text="مسیر پیش‌فرض ذخیره نامه‌ها:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5) # Font handled by style
        self.entry_default_save_path = ttk.Entry(settings_frame, width=50) # Font handled by style
        self.entry_default_save_path.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.entry_default_save_path.insert(0, default_save_path) # Load current setting
        ttk.Button(settings_frame, text="انتخاب پوشه", command=self._select_default_save_path).grid(row=2, column=2, sticky=tk.W, padx=5) # Font handled by style

        # Letterhead Template Path
        ttk.Label(settings_frame, text="مسیر فایل الگو سربرگ (docx):").grid(row=3, column=0, sticky=tk.W, pady=5, padx=5) # Font handled by style
        self.entry_letterhead_template = ttk.Entry(settings_frame, width=50) # Font handled by style
        self.entry_letterhead_template.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        self.entry_letterhead_template.insert(0, letterhead_template_path) # Load current setting
        ttk.Button(settings_frame, text="انتخاب فایل الگو", command=self._select_letterhead_template).grid(row=3, column=2, sticky=tk.W, padx=5) # Font handled by style

        # Save Settings Button
        save_button = ttk.Button(settings_frame, text="ذخیره تنظیمات", command=self._save_current_settings) # Font handled by style
        save_button.grid(row=4, column=0, columnspan=3, pady=15)
        
        # Configure column weights for resizing
        settings_frame.grid_columnconfigure(1, weight=1)

    def _select_default_save_path(self):
        folder_selected = filedialog.askdirectory(parent=self.root, 
                                                  initialdir=self.entry_default_save_path.get() or os.getcwd(),
                                                  title="انتخاب پوشه ذخیره نامه‌ها")
        if folder_selected:
            self.entry_default_save_path.delete(0, tk.END)
            self.entry_default_save_path.insert(0, folder_selected)

    def _select_letterhead_template(self):
        file_selected = filedialog.askopenfilename(parent=self.root, 
                                                   initialdir=os.path.dirname(self.entry_letterhead_template.get()) or os.getcwd(),
                                                   title="انتخاب فایل الگوی سربرگ",
                                                   filetypes=[("Word Documents", "*.docx")])
        if file_selected:
            self.entry_letterhead_template.delete(0, tk.END)
            self.entry_letterhead_template.insert(0, file_selected)

    def _save_current_settings(self):
        # Update global variables in settings_manager module
        import settings_manager # Re-import to access global variables directly
        settings_manager.company_name = self.entry_company_name.get().strip()
        settings_manager.full_company_name = self.entry_full_company_name.get().strip() 
        settings_manager.default_save_path = self.entry_default_save_path.get().strip()
        settings_manager.letterhead_template_path = self.entry_letterhead_template.get().strip()

        # Call save function from settings_manager
        settings_manager.save_settings()
        messagebox.showinfo("ذخیره تنظیمات", "تنظیمات با موفقیت ذخیره شد.", parent=self.root)
        
        # After saving, refresh the settings tab to show the newly saved values (optional but good for consistency)
        self._on_tab_change(None) # Pass None as event as it's an internal call


    # Helper method for context menu
    def _show_text_context_menu(self, event):
        """Displays a right-click context menu for Text and Entry widgets."""
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="بریدن", command=lambda: event.widget.event_generate("<<Cut>>"))
        context_menu.add_command(label="کپی", command=lambda: event.widget.event_generate("<<Copy>>"))
        context_menu.add_command(label="چسباندن", command=lambda: event.widget.event_generate("<<Paste>>"))
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    # Public methods for external modules to call (if needed)
    def update_history_treeview(self, search_term=""):
        update_history_treeview(search_term, self.history_treeview, self.status_bar)


if __name__ == "__main__":
    # Ensure database tables exist when the application starts
    create_tables()

    root = ThemedTk(theme="arc") # Use a theme for better aesthetics
    app = App(root)
    root.mainloop()
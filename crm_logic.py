import tkinter as tk
from tkinter import messagebox, ttk, simpledialog

from database import (
    get_db_connection,
    get_organizations_from_db,
    insert_organization,
    update_organization,
    delete_organization,
    get_contacts_from_db,
    insert_contact,
    update_contact,
    delete_contact,
    get_organization_by_id,
    get_contact_by_id
)

# Assuming BASE_FONT is defined globally or passed. For now, define locally if not available.
try:
    BASE_FONT
except NameError:
    BASE_FONT = ("Arial", 10)


def populate_organizations_treeview(search_term="", org_treeview_ref=None, status_bar_ref=None):
    """Populates the organizations Treeview with data from the database."""
    if org_treeview_ref:
        for item in org_treeview_ref.get_children():
            org_treeview_ref.delete(item)
        
        organizations = get_organizations_from_db(search_term)
        for org in organizations:
            org_treeview_ref.insert("", tk.END, values=(
                org['id'],
                org['name'],
                org['industry'],
                org['phone'],
                org['email'],
                org['address'],
                org['description']
            ), iid=org['id'])
        if status_bar_ref: status_bar_ref.config(text=f"نمایش {len(organizations)} سازمان.")

def populate_contacts_treeview(organization_id=None, search_term="", contact_treeview_ref=None, status_bar_ref=None):
    """Populates the contacts Treeview with data from the database."""
    if contact_treeview_ref:
        for item in contact_treeview_ref.get_children():
            contact_treeview_ref.delete(item)
        
        contacts = get_contacts_from_db(organization_id=organization_id, search_term=search_term)
        for contact in contacts:
            org_name = contact['organization_name'] if contact['organization_name'] else "---"
            contact_treeview_ref.insert("", tk.END, values=(
                contact['id'],
                contact['first_name'],
                contact['last_name'],
                org_name, # Display organization name
                contact['title'],
                contact['phone'],
                contact['email'],
                contact['notes']
            ), iid=contact['id'])
        if status_bar_ref: status_bar_ref.config(text=f"نمایش {len(contacts)} مخاطب.")


# --- Organization Management Functions (now called by dialogs) ---
# These functions are simplified as dialogs will handle data collection
def on_add_organization(name, industry, phone, email, address, description, org_treeview_ref, status_bar_ref, populate_org_contact_combos_callback, root_window_ref):
    """Handles adding a new organization from dialog data."""
    if not name:
        messagebox.showwarning("ورودی ناقص", "لطفاً نام سازمان را وارد کنید.", parent=root_window_ref)
        return False

    if insert_organization(name, industry, phone, email, address, description):
        messagebox.showinfo("موفقیت", "سازمان با موفقیت اضافه شد.", parent=root_window_ref)
        populate_organizations_treeview(org_treeview_ref=org_treeview_ref, status_bar_ref=status_bar_ref)
        populate_org_contact_combos_callback()
        if status_bar_ref: status_bar_ref.config(text=f"سازمان '{name}' اضافه شد.")
        return True
    else:
        messagebox.showerror("خطا", "خطا در افزودن سازمان.", parent=root_window_ref)
        return False

def on_edit_organization_button(root_window_ref, entry_name, entry_industry, entry_phone, entry_email, entry_address, text_description, org_treeview_ref, contact_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    """Handles editing an existing organization."""
    selected_items = org_treeview_ref.selection()
    if selected_items:
        org_id = org_treeview_ref.item(selected_items[0], 'iid')
        name = entry_name.get().strip()
        industry = entry_industry.get().strip()
        phone = entry_phone.get().strip()
        email = entry_email.get().strip()
        address = entry_address.get().strip()
        description = text_description.get("1.0", tk.END).strip()

        if not name:
            messagebox.showwarning("ورودی ناقص", "لطفاً نام سازمان را وارد کنید.", parent=root_window_ref)
            return

        if update_organization(org_id, name, industry, phone, email, address, description):
            messagebox.showinfo("موفقیت", "سازمان با موفقیت ویرایش شد.", parent=root_window_ref)
            populate_organizations_treeview(org_treeview_ref=org_treeview_ref, status_bar_ref=status_bar_ref)
            populate_contacts_treeview(org_treeview_ref.item(org_treeview_ref.selection()[0], 'iid') if org_treeview_ref.selection() else None, "", contact_treeview_ref, status_bar_ref)
            populate_org_contact_combos_callback() # Refresh combos in main app
            if status_bar_ref: status_bar_ref.config(text=f"سازمان '{name}' ویرایش شد.")
        else:
            messagebox.showerror("خطا", "خطا در ویرایش سازمان.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای ویرایش یک سازمان، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)

def on_delete_organization_button(root_window_ref, org_treeview_ref, contact_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    """Handles deleting an organization."""
    selected_items = org_treeview_ref.selection()
    if selected_items:
        org_id = org_treeview_ref.item(selected_items[0], 'iid')
        org_name = org_treeview_ref.item(selected_items[0], 'values')[1]
        
        response = messagebox.askyesno("تأیید حذف", f"آیا مطمئن هستید که می‌خواهید سازمان '{org_name}' را حذف کنید؟\n\nتمام مخاطبین و نامه‌های مرتبط با این سازمان (در صورت وجود) نیز ممکن است تحت تأثیر قرار گیرند.", parent=root_window_ref)
        if response:
            if delete_organization(org_id):
                messagebox.showinfo("موفقیت", "سازمان با موفقیت حذف شد.", parent=root_window_ref)
                populate_organizations_treeview(org_treeview_ref=org_treeview_ref, status_bar_ref=status_bar_ref)
                populate_contacts_treeview(None, "", contact_treeview_ref, status_bar_ref)
                populate_org_contact_combos_callback() # Refresh combos in main app
                if status_bar_ref: status_bar_ref.config(text=f"سازمان '{org_name}' حذف شد.")
            else:
                messagebox.showerror("خطا", "خطا در حذف سازمان.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای حذف یک سازمان، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)

# --- Contact Management Functions (now called by dialogs) ---
def on_add_contact(organization_id, first_name, last_name, title, phone, email, notes, contact_treeview_ref, status_bar_ref, populate_org_contact_combos_callback, root_window_ref):
    """Handles adding a new contact from dialog data."""
    if not first_name or not last_name:
        messagebox.showwarning("ورودی ناقص", "لطفاً نام و نام خانوادگی مخاطب را وارد کنید.", parent=root_window_ref)
        return False
    
    if insert_contact(organization_id, first_name, last_name, title, phone, email, notes):
        messagebox.showinfo("موفقیت", "مخاطب با موفقیت اضافه شد.", parent=root_window_ref)
        populate_contacts_treeview(organization_id, "", contact_treeview_ref, status_bar_ref)
        populate_org_contact_combos_callback()
        if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{first_name} {last_name}' اضافه شد.")
        return True
    else:
        messagebox.showerror("خطا", "خطا در افزودن مخاطب.", parent=root_window_ref)
        return False

def on_edit_contact_button(root_window_ref, entry_org_id, entry_first_name, entry_last_name, entry_title, entry_phone, entry_email, text_notes, contact_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    """Handles editing an existing contact."""
    selected_items = contact_treeview_ref.selection()
    if selected_items:
        contact_id = contact_treeview_ref.item(selected_items[0], 'iid')
        organization_id_str = entry_org_id.get().strip()
        first_name = entry_first_name.get().strip()
        last_name = entry_last_name.get().strip()
        title = entry_title.get().strip()
        phone = entry_phone.get().strip()
        email = entry_email.get().strip()
        notes = text_notes.get("1.0", tk.END).strip()

        if not first_name or not last_name:
            messagebox.showwarning("ورودی ناقص", "لطفاً نام و نام خانوادگی مخاطب را وارد کنید.", parent=root_window_ref)
            return
        
        organization_id = int(organization_id_str) if organization_id_str else None

        if update_contact(contact_id, organization_id, first_name, last_name, title, phone, email, notes):
            messagebox.showinfo("موفقیت", "مخاطب با موفقیت ویرایش شد.", parent=root_window_ref)
            populate_contacts_treeview(organization_id, "", contact_treeview_ref, status_bar_ref)
            populate_org_contact_combos_callback() # Refresh combos in main app
            if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{first_name} {last_name}' ویرایش شد.")
        else:
            messagebox.showerror("خطا", "خطا در ویرایش مخاطب.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای ویرایش یک مخاطب، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)

def on_delete_contact_button(root_window_ref, contact_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    """Handles deleting a contact."""
    selected_items = contact_treeview_ref.selection()
    if selected_items:
        contact_id = contact_treeview_ref.item(selected_items[0], 'iid')
        contact_name = f"{contact_treeview_ref.item(selected_items[0], 'values')[1]} {contact_treeview_ref.item(selected_items[0], 'values')[2]}"
        
        response = messagebox.askyesno("تأیید حذف", f"آیا مطمئن هستید که می‌خواهید مخاطب '{contact_name}' را حذف کنید؟", parent=root_window_ref)
        if response:
            if delete_contact(contact_id):
                messagebox.showinfo("موفقیت", "مخاطب با موفقیت حذف شد.", parent=root_window_ref)
                populate_contacts_treeview(None, "", contact_treeview_ref, status_bar_ref)
                populate_org_contact_combos_callback() # Refresh combos in main app
                if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{contact_name}' حذف شد.")
            else:
                messagebox.showerror("خطا", "خطا در حذف مخاطب.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای حذف یک مخاطب، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)

def on_organization_select(event, org_treeview_ref, contact_treeview_ref, status_bar_ref, entry_org_name, entry_org_industry, entry_org_phone, entry_org_email, entry_org_address, text_org_description):
    """Filters contacts based on the selected organization and populates org entry fields."""
    selected_items = org_treeview_ref.selection()
    if selected_items:
        selected_org_id = org_treeview_ref.item(selected_items[0], 'iid')
        org_data = get_organization_by_id(selected_org_id)
        if org_data:
            entry_org_name.delete(0, tk.END)
            entry_org_name.insert(0, org_data['name'])
            entry_org_industry.delete(0, tk.END)
            entry_org_industry.insert(0, org_data['industry'] if org_data['industry'] else "")
            entry_org_phone.delete(0, tk.END)
            entry_org_phone.insert(0, org_data['phone'] if org_data['phone'] else "")
            entry_org_email.delete(0, tk.END)
            entry_org_email.insert(0, org_data['email'] if org_data['email'] else "")
            entry_org_address.delete(0, tk.END)
            entry_org_address.insert(0, org_data['address'] if org_data['address'] else "")
            text_org_description.delete("1.0", tk.END)
            text_org_description.insert("1.0", org_data['description'] if org_data['description'] else "")

        populate_contacts_treeview(organization_id=selected_org_id, search_term="",
                                   contact_treeview_ref=contact_treeview_ref, status_bar_ref=status_bar_ref)
        if status_bar_ref: status_bar_ref.config(text=f"مخاطبین سازمان انتخاب شده فیلتر شدند.")
    else:
        # Clear organization entry fields if no organization is selected
        entry_org_name.delete(0, tk.END)
        entry_org_industry.delete(0, tk.END)
        entry_org_phone.delete(0, tk.END)
        entry_org_email.delete(0, tk.END)
        entry_org_address.delete(0, tk.END)
        text_org_description.delete("1.0", tk.END)
        
        # If no organization is selected, show all contacts
        populate_contacts_treeview(organization_id=None, search_term="",
                                   contact_treeview_ref=contact_treeview_ref, status_bar_ref=status_bar_ref)
        if status_bar_ref: status_bar_ref.config(text="فیلتر مخاطبین برداشته شد. نمایش همه مخاطبین.")

def on_contact_select(event, contact_treeview_ref, entry_org_id, entry_first_name, entry_last_name, entry_title, entry_phone, entry_email, text_notes):
    """Populates contact entry fields based on the selected contact."""
    selected_items = contact_treeview_ref.selection()
    if selected_items:
        selected_contact_id = contact_treeview_ref.item(selected_items[0], 'iid')
        contact_data = get_contact_by_id(selected_contact_id)
        if contact_data:
            entry_org_id.delete(0, tk.END)
            entry_org_id.insert(0, contact_data['organization_id'] if contact_data['organization_id'] else "")
            entry_first_name.delete(0, tk.END)
            entry_first_name.insert(0, contact_data['first_name'] if contact_data['first_name'] else "")
            entry_last_name.delete(0, tk.END)
            entry_last_name.insert(0, contact_data['last_name'] if contact_data['last_name'] else "")
            entry_title.delete(0, tk.END)
            entry_title.insert(0, contact_data['title'] if contact_data['title'] else "")
            entry_phone.delete(0, tk.END)
            entry_phone.insert(0, contact_data['phone'] if contact_data['phone'] else "")
            entry_email.delete(0, tk.END)
            entry_email.insert(0, contact_data['email'] if contact_data['email'] else "")
            text_notes.delete("1.0", tk.END)
            text_notes.insert("1.0", contact_data['notes'] if contact_data['notes'] else "")
    else:
        # Clear contact entry fields if no contact is selected
        entry_org_id.delete(0, tk.END)
        entry_first_name.delete(0, tk.END)
        entry_last_name.delete(0, tk.END)
        entry_title.delete(0, tk.END)
        entry_phone.delete(0, tk.END)
        entry_email.delete(0, tk.END)
        text_notes.delete("1.0", tk.END)


class AddOrganizationDialog(tk.Toplevel):
    def __init__(self, parent, populate_organizations_callback, populate_all_combos_callback, org_treeview_ref, status_bar_ref):
        super().__init__(parent)
        self.parent = parent
        self.populate_organizations_callback = populate_organizations_callback
        self.populate_all_combos_callback = populate_all_combos_callback
        self.org_treeview_ref = org_treeview_ref
        self.status_bar_ref = status_bar_ref

        self.title("افزودن سازمان جدید")
        self.geometry("450x400")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self._center_window()

        self._create_widgets()

    def _center_window(self):
        self.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        form_frame = ttk.Frame(self, padding="10")
        form_frame.pack(fill=tk.BOTH, expand=True)

        # Input fields for Organization
        ttk.Label(form_frame, text="نام سازمان:").grid(row=0, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_name = ttk.Entry(form_frame)
        self.entry_name.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="صنعت:").grid(row=1, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_industry = ttk.Entry(form_frame)
        self.entry_industry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="تلفن:").grid(row=2, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_phone = ttk.Entry(form_frame)
        self.entry_phone.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="ایمیل:").grid(row=3, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_email = ttk.Entry(form_frame)
        self.entry_email.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="آدرس:").grid(row=4, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_address = ttk.Entry(form_frame)
        self.entry_address.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="توضیحات:").grid(row=5, column=0, sticky=tk.E, padx=5, pady=2)
        self.text_description = tk.Text(form_frame, wrap="word", height=5, font=BASE_FONT)
        self.text_description.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=2)

        form_frame.grid_columnconfigure(1, weight=1)

        button_frame = ttk.Frame(self, padding="10")
        button_frame.pack(fill=tk.X, pady=10)
        ttk.Button(button_frame, text="ذخیره", command=self._on_save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="لغو", command=self.destroy).pack(side=tk.RIGHT, padx=5)

    def _on_save(self):
        name = self.entry_name.get().strip()
        industry = self.entry_industry.get().strip()
        phone = self.entry_phone.get().strip()
        email = self.entry_email.get().strip()
        address = self.entry_address.get().strip()
        description = self.text_description.get("1.0", tk.END).strip()

        if on_add_organization(name, industry, phone, email, address, description, self.org_treeview_ref, self.status_bar_ref, self.populate_all_combos_callback, self.parent):
            self.destroy()


class AddContactDialog(tk.Toplevel):
    def __init__(self, parent, populate_contacts_callback, populate_all_combos_callback, contact_treeview_ref, status_bar_ref):
        super().__init__(parent)
        self.parent = parent
        self.populate_contacts_callback = populate_contacts_callback
        self.populate_all_combos_callback = populate_all_combos_callback
        self.contact_treeview_ref = contact_treeview_ref
        self.status_bar_ref = status_bar_ref

        self.selected_org_id = None
        self.selected_org_name = ""

        self.title("افزودن مخاطب جدید")
        self.geometry("450x450")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self._center_window()

        self._create_widgets()

    def _center_window(self):
        self.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() // 2) - (self.winfo_width() // 2)
        y = self.parent.winfo_y() + (self.parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        form_frame = ttk.Frame(self, padding="10")
        form_frame.pack(fill=tk.BOTH, expand=True)

        # Organization Selection
        ttk.Label(form_frame, text="سازمان:").grid(row=0, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_org_name = ttk.Entry(form_frame, state="readonly")
        self.entry_org_name.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Button(form_frame, text="انتخاب سازمان", command=self._open_org_selection_dialog).grid(row=0, column=2, padx=5, pady=2)

        ttk.Label(form_frame, text="نام:").grid(row=1, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_first_name = ttk.Entry(form_frame)
        self.entry_first_name.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="نام خانوادگی:").grid(row=2, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_last_name = ttk.Entry(form_frame)
        self.entry_last_name.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="عنوان:").grid(row=3, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_title = ttk.Entry(form_frame)
        self.entry_title.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="تلفن:").grid(row=4, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_phone = ttk.Entry(form_frame)
        self.entry_phone.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="ایمیل:").grid(row=5, column=0, sticky=tk.E, padx=5, pady=2)
        self.entry_email = ttk.Entry(form_frame)
        self.entry_email.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(form_frame, text="یادداشت‌ها:").grid(row=6, column=0, sticky=tk.E, padx=5, pady=2)
        self.text_notes = tk.Text(form_frame, wrap="word", height=5, font=BASE_FONT)
        self.text_notes.grid(row=6, column=1, sticky=tk.EW, padx=5, pady=2)

        form_frame.grid_columnconfigure(1, weight=1)

        button_frame = ttk.Frame(self, padding="10")
        button_frame.pack(fill=tk.X, pady=10)
        ttk.Button(button_frame, text="ذخیره", command=self._on_save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="لغو", command=self.destroy).pack(side=tk.RIGHT, padx=5)

    def _open_org_selection_dialog(self):
        dialog = tk.Toplevel(self) # Parent is this dialog
        dialog.title("انتخاب سازمان")
        dialog.transient(self)
        dialog.grab_set()
        dialog.geometry("600x400")

        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        dialog_frame = ttk.Frame(dialog, padding="10")
        dialog_frame.pack(fill=tk.BOTH, expand=True)

        search_frame = ttk.Frame(dialog_frame)
        search_frame.pack(fill=tk.X, pady=5)
        ttk.Label(search_frame, text="جستجو:").pack(side=tk.RIGHT, padx=5)
        org_dialog_search_entry = ttk.Entry(search_frame)
        org_dialog_search_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)
        
        # Treeview for organizations
        tree_frame = ttk.Frame(dialog_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        org_dialog_treeview = ttk.Treeview(tree_frame, columns=("id", "name", "industry"), show="headings", yscrollcommand=tree_scrollbar.set)
        tree_scrollbar.config(command=org_dialog_treeview.yview)

        org_dialog_treeview.heading("id", text="شناسه")
        org_dialog_treeview.heading("name", text="نام سازمان")
        org_dialog_treeview.heading("industry", text="صنعت")
        org_dialog_treeview.column("id", width=50, stretch=tk.NO)
        org_dialog_treeview.column("name", width=200)
        org_dialog_treeview.column("industry", width=150)
        org_dialog_treeview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def _populate_org_dialog_treeview_internal(search_term=""):
            for i in org_dialog_treeview.get_children():
                org_dialog_treeview.delete(i)

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
                org_dialog_treeview.insert("", tk.END, values=(org['id'], org['name'], org['industry']))

        org_dialog_search_entry.bind("<Return>", lambda event: _populate_org_dialog_treeview_internal(org_dialog_search_entry.get()))
        ttk.Button(search_frame, text="جستجو", command=lambda: _populate_org_dialog_treeview_internal(org_dialog_search_entry.get())).pack(side=tk.RIGHT, padx=5)

        def _select_org():
            selected_item = org_dialog_treeview.focus()
            if selected_item:
                values = org_dialog_treeview.item(selected_item, 'values')
                self.selected_org_id = values[0]
                self.selected_org_name = values[1]
                self.entry_org_name.config(state="normal")
                self.entry_org_name.delete(0, tk.END)
                self.entry_org_name.insert(0, self.selected_org_name)
                self.entry_org_name.config(state="readonly")
                dialog.destroy()
            else:
                messagebox.showwarning("انتخاب سازمان", "لطفاً یک سازمان را انتخاب کنید.", parent=dialog)

        org_dialog_treeview.bind("<Double-1>", lambda event: _select_org())

        button_frame_dialog = ttk.Frame(dialog_frame)
        button_frame_dialog.pack(fill=tk.X, pady=5)
        ttk.Button(button_frame_dialog, text="انتخاب", command=_select_org).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame_dialog, text="لغو", command=dialog.destroy).pack(side=tk.LEFT, padx=5)

        _populate_org_dialog_treeview_internal()
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        self.wait_window(dialog) # Wait for this dialog to close

    def _on_save(self):
        first_name = self.entry_first_name.get().strip()
        last_name = self.entry_last_name.get().strip()
        title = self.entry_title.get().strip()
        phone = self.entry_phone.get().strip()
        email = self.entry_email.get().strip()
        notes = self.text_notes.get("1.0", tk.END).strip()

        if on_add_contact(self.selected_org_id, first_name, last_name, title, phone, email, notes, self.contact_treeview_ref, self.status_bar_ref, self.populate_all_combos_callback, self.parent):
            self.destroy()

# --- Organization and Contact functions (existing, simplified for dialog interaction) ---
# The original on_add_organization_button and on_add_contact_button are now split
# into on_add_organization and on_add_contact, which are called by the dialogs.
# The on_edit and on_delete functions remain largely the same, but will not
# interact with the main CRM tab's input fields anymore.

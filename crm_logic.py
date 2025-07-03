import tkinter as tk
from tkinter import messagebox, ttk
from database import (
    get_organizations, insert_organization, update_organization, delete_organization,
    get_contacts, insert_contact, update_contact, delete_contact,
    get_db_connection # <--- Add this import
)

def populate_organizations_treeview(search_term="", org_treeview_ref=None, status_bar_ref=None):
    """Populates the organizations Treeview with data from the database."""
    if org_treeview_ref: # Ensure org_treeview_ref is initialized
        for item in org_treeview_ref.get_children():
            org_treeview_ref.delete(item)
        
        organizations = get_organizations(search_term)
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
    if contact_treeview_ref: # Ensure contact_treeview_ref is initialized
        for item in contact_treeview_ref.get_children():
            contact_treeview_ref.delete(item)
        
        contacts = get_contacts(organization_id, search_term)
        for contact in contacts:
            contact_treeview_ref.insert("", tk.END, values=(
                contact['id'],
                contact['organization_id'],
                contact['first_name'],
                contact['last_name'],
                contact['title'],
                contact['phone'],
                contact['email'],
                contact['notes']
            ), iid=contact['id'])
        if status_bar_ref: status_bar_ref.config(text=f"نمایش {len(contacts)} مخاطب.")


def on_add_organization_button(root_window_ref, org_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    add_org_window = tk.Toplevel(root_window_ref)
    add_org_window.title("افزودن سازمان جدید")
    add_org_window.grab_set() # Make the window modal

    input_frame = ttk.Frame(add_org_window, padding="10")
    input_frame.pack(fill=tk.BOTH, expand=True)

    # Labels and Entries
    ttk.Label(input_frame, text="نام سازمان:").grid(row=0, column=0, sticky=tk.E, padx=5, pady=2)
    entry_name = ttk.Entry(input_frame, width=40)
    entry_name.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)

    ttk.Label(input_frame, text="صنعت:").grid(row=1, column=0, sticky=tk.E, padx=5, pady=2)
    entry_industry = ttk.Entry(input_frame, width=40)
    entry_industry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)

    ttk.Label(input_frame, text="تلفن:").grid(row=2, column=0, sticky=tk.E, padx=5, pady=2)
    entry_phone = ttk.Entry(input_frame, width=40)
    entry_phone.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)

    ttk.Label(input_frame, text="ایمیل:").grid(row=3, column=0, sticky=tk.E, padx=5, pady=2)
    entry_email = ttk.Entry(input_frame, width=40)
    entry_email.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)

    ttk.Label(input_frame, text="آدرس:").grid(row=4, column=0, sticky=tk.E, padx=5, pady=2)
    entry_address = ttk.Entry(input_frame, width=40)
    entry_address.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)

    ttk.Label(input_frame, text="توضیحات:").grid(row=5, column=0, sticky=tk.E, padx=5, pady=2)
    text_description = tk.Text(input_frame, width=40, height=5, font=("Arial", 10))
    text_description.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=2)

    input_frame.grid_columnconfigure(1, weight=1)

    def save_organization():
        name = entry_name.get().strip()
        industry = entry_industry.get().strip()
        phone = entry_phone.get().strip()
        email = entry_email.get().strip()
        address = entry_address.get().strip()
        description = text_description.get("1.0", tk.END).strip()

        if not name:
            messagebox.showerror("خطا", "نام سازمان نمی‌تواند خالی باشد.", parent=add_org_window)
            return

        if insert_organization(name, industry, phone, email, address, description):
            messagebox.showinfo("موفقیت", "سازمان با موفقیت اضافه شد.", parent=add_org_window)
            populate_organizations_treeview("", org_treeview_ref, status_bar_ref)
            populate_org_contact_combos_callback() # Callback to update main app combos
            add_org_window.destroy()
            if status_bar_ref: status_bar_ref.config(text=f"سازمان '{name}' اضافه شد.")
        else:
            messagebox.showerror("خطا", "خطا در افزودن سازمان.", parent=add_org_window)

    save_button = ttk.Button(add_org_window, text="ذخیره", command=save_organization)
    save_button.pack(pady=10)


def on_edit_organization_button(root_window_ref, org_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    selected_item = org_treeview_ref.selection()
    if selected_item:
        org_id = org_treeview_ref.item(selected_item[0], 'iid')
        current_values = org_treeview_ref.item(selected_item[0], 'values')

        edit_org_window = tk.Toplevel(root_window_ref)
        edit_org_window.title(f"ویرایش سازمان: {current_values[1]}")
        edit_org_window.grab_set()

        input_frame = ttk.Frame(edit_org_window, padding="10")
        input_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(input_frame, text="نام سازمان:").grid(row=0, column=0, sticky=tk.E, padx=5, pady=2)
        entry_name = ttk.Entry(input_frame, width=40)
        entry_name.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_name.insert(0, current_values[1])

        ttk.Label(input_frame, text="صنعت:").grid(row=1, column=0, sticky=tk.E, padx=5, pady=2)
        entry_industry = ttk.Entry(input_frame, width=40)
        entry_industry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_industry.insert(0, current_values[2])

        ttk.Label(input_frame, text="تلفن:").grid(row=2, column=0, sticky=tk.E, padx=5, pady=2)
        entry_phone = ttk.Entry(input_frame, width=40)
        entry_phone.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_phone.insert(0, current_values[3])

        ttk.Label(input_frame, text="ایمیل:").grid(row=3, column=0, sticky=tk.E, padx=5, pady=2)
        entry_email = ttk.Entry(input_frame, width=40)
        entry_email.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_email.insert(0, current_values[4])

        ttk.Label(input_frame, text="آدرس:").grid(row=4, column=0, sticky=tk.E, padx=5, pady=2)
        entry_address = ttk.Entry(input_frame, width=40)
        entry_address.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_address.insert(0, current_values[5])

        ttk.Label(input_frame, text="توضیحات:").grid(row=5, column=0, sticky=tk.E, padx=5, pady=2)
        text_description = tk.Text(input_frame, width=40, height=5, font=("Arial", 10))
        text_description.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=2)
        text_description.insert("1.0", current_values[6])

        input_frame.grid_columnconfigure(1, weight=1)

        def update_organization_data():
            name = entry_name.get().strip()
            industry = entry_industry.get().strip()
            phone = entry_phone.get().strip()
            email = entry_email.get().strip()
            address = entry_address.get().strip()
            description = text_description.get("1.0", tk.END).strip()

            if not name:
                messagebox.showerror("خطا", "نام سازمان نمی‌تواند خالی باشد.", parent=edit_org_window)
                return

            if update_organization(org_id, name, industry, phone, email, address, description):
                messagebox.showinfo("موفقیت", "سازمان با موفقیت ویرایش شد.", parent=edit_org_window)
                populate_organizations_treeview("", org_treeview_ref, status_bar_ref)
                populate_org_contact_combos_callback()
                edit_org_window.destroy()
                if status_bar_ref: status_bar_ref.config(text=f"سازمان '{name}' ویرایش شد.")
            else:
                messagebox.showerror("خطا", "خطا در ویرایش سازمان.", parent=edit_org_window)

        save_button = ttk.Button(edit_org_window, text="ذخیره تغییرات", command=update_organization_data)
        save_button.pack(pady=10)
    else:
        messagebox.showwarning("انتخاب نشده", "برای ویرایش یک سازمان، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)


def on_delete_organization_button(root_window_ref, org_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    selected_item = org_treeview_ref.selection()
    if selected_item:
        org_id = org_treeview_ref.item(selected_item[0], 'iid')
        org_name = org_treeview_ref.item(selected_item[0], 'values')[1]
        
        confirm = messagebox.askyesno("حذف سازمان", f"آیا مطمئن هستید که می‌خواهید سازمان '{org_name}' را حذف کنید؟ این کار مخاطبین مرتبط را نیز تحت تاثیر قرار می‌دهد.", parent=root_window_ref)
        if confirm:
            if delete_organization(org_id):
                messagebox.showinfo("موفقیت", "سازمان با موفقیت حذف شد.", parent=root_window_ref)
                populate_organizations_treeview("", org_treeview_ref, status_bar_ref)
                populate_contacts_treeview(contact_treeview_ref=root_window_ref.nametowidget(org_treeview_ref.winfo_parent()).nametowidget("!frame2.!treeview"), status_bar_ref=status_bar_ref) # Refresh contacts, assuming a known path
                populate_org_contact_combos_callback()
                if status_bar_ref: status_bar_ref.config(text=f"سازمان '{org_name}' حذف شد.")
            else:
                messagebox.showerror("خطا", "خطا در حذف سازمان.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای حذف یک سازمان، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)


def on_add_contact_button(root_window_ref, contact_treeview_ref, org_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    add_contact_window = None
    try:
        add_contact_window = tk.Toplevel(root_window_ref)
        add_contact_window.title("افزودن مخاطب جدید")
        add_contact_window.grab_set() # Make the window modal

        input_frame = ttk.Frame(add_contact_window, padding="10")
        input_frame.pack(fill=tk.BOTH, expand=True)

        # Get organizations for combobox
        conn = get_db_connection() # <--- This function needs to be imported
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM Organizations ORDER BY name")
        organizations = cursor.fetchall()
        org_names = ["---"] + [org['name'] for org in organizations]
        org_map = {org['name']: org['id'] for org in organizations}
        conn.close()

        # Labels and Entries (using grid for better alignment)
        ttk.Label(input_frame, text="سازمان:").grid(row=0, column=0, sticky=tk.E, padx=5, pady=2)
        combo_org = ttk.Combobox(input_frame, values=org_names, state="readonly", width=37)
        combo_org.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        combo_org.set("---") # Default value

        # Pre-select organization if one is selected in the main CRM tab
        selected_org_items = org_treeview_ref.selection()
        if selected_org_items:
            selected_org_name = org_treeview_ref.item(selected_org_items[0], 'values')[1]
            if selected_org_name in org_names:
                combo_org.set(selected_org_name)

        ttk.Label(input_frame, text="نام:").grid(row=1, column=0, sticky=tk.E, padx=5, pady=2)
        entry_first_name = ttk.Entry(input_frame, width=40)
        entry_first_name.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(input_frame, text="نام خانوادگی:").grid(row=2, column=0, sticky=tk.E, padx=5, pady=2)
        entry_last_name = ttk.Entry(input_frame, width=40)
        entry_last_name.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(input_frame, text="عنوان:").grid(row=3, column=0, sticky=tk.E, padx=5, pady=2)
        entry_title = ttk.Entry(input_frame, width=40)
        entry_title.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(input_frame, text="تلفن:").grid(row=4, column=0, sticky=tk.E, padx=5, pady=2)
        entry_phone = ttk.Entry(input_frame, width=40)
        entry_phone.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(input_frame, text="ایمیل:").grid(row=5, column=0, sticky=tk.E, padx=5, pady=2)
        entry_email = ttk.Entry(input_frame, width=40)
        entry_email.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=2)

        ttk.Label(input_frame, text="یادداشت‌ها:").grid(row=6, column=0, sticky=tk.E, padx=5, pady=2)
        text_notes = tk.Text(input_frame, width=40, height=5, font=("Arial", 10))
        text_notes.grid(row=6, column=1, sticky=tk.EW, padx=5, pady=2)

        input_frame.grid_columnconfigure(1, weight=1)

        def save_contact():
            org_name = combo_org.get()
            first_name = entry_first_name.get().strip()
            last_name = entry_last_name.get().strip()
            title = entry_title.get().strip()
            phone = entry_phone.get().strip()
            email = entry_email.get().strip()
            notes = text_notes.get("1.0", tk.END).strip()

            if not first_name or not last_name:
                messagebox.showerror("خطا", "نام و نام خانوادگی مخاطب نمی‌تواند خالی باشد.", parent=add_contact_window)
                return

            organization_id = None
            if org_name and org_name != "---":
                organization_id = org_map.get(org_name)

            if insert_contact(organization_id, first_name, last_name, title, phone, email, notes):
                messagebox.showinfo("موفقیت", "مخاطب با موفقیت اضافه شد.", parent=add_contact_window)
                populate_contacts_treeview(organization_id, "", contact_treeview_ref, status_bar_ref)
                populate_org_contact_combos_callback() # Callback to update main app combos
                add_contact_window.destroy()
                if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{first_name} {last_name}' اضافه شد.")
            else:
                messagebox.showerror("خطا", "خطا در افزودن مخاطب.", parent=add_contact_window)

        save_button = ttk.Button(add_contact_window, text="ذخیره", command=save_contact)
        save_button.pack(pady=10)

    except Exception as e:
        if add_contact_window and add_contact_window.winfo_exists():
            add_contact_window.destroy()
        messagebox.showerror("خطا", f"خطا در باز کردن پنجره افزودن مخاطب: {e}", parent=root_window_ref)
        print(f"DEBUG: Error opening add contact window: {e}") # Print to console for debugging
        return


def on_edit_contact_button(root_window_ref, contact_treeview_ref, org_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    selected_item = contact_treeview_ref.selection()
    if selected_item:
        contact_id = contact_treeview_ref.item(selected_item[0], 'iid')
        current_values = contact_treeview_ref.item(selected_item[0], 'values')

        edit_contact_window = tk.Toplevel(root_window_ref)
        edit_contact_window.title(f"ویرایش مخاطب: {current_values[2]} {current_values[3]}")
        edit_contact_window.grab_set()

        input_frame = ttk.Frame(edit_contact_window, padding="10")
        input_frame.pack(fill=tk.BOTH, expand=True)

        # Get organizations for combobox
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM Organizations ORDER BY name")
        organizations = cursor.fetchall()
        org_names = ["---"] + [org['name'] for org in organizations]
        org_map = {org['name']: org['id'] for org in organizations}
        conn.close()

        ttk.Label(input_frame, text="سازمان:").grid(row=0, column=0, sticky=tk.E, padx=5, pady=2)
        combo_org = ttk.Combobox(input_frame, values=org_names, state="readonly", width=37)
        combo_org.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        
        # Set current organization
        current_org_id = current_values[1]
        current_org_name = "---"
        for name, oid in org_map.items():
            if oid == current_org_id:
                current_org_name = name
                break
        combo_org.set(current_org_name)


        ttk.Label(input_frame, text="نام:").grid(row=1, column=0, sticky=tk.E, padx=5, pady=2)
        entry_first_name = ttk.Entry(input_frame, width=40)
        entry_first_name.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_first_name.insert(0, current_values[2])

        ttk.Label(input_frame, text="نام خانوادگی:").grid(row=2, column=0, sticky=tk.E, padx=5, pady=2)
        entry_last_name = ttk.Entry(input_frame, width=40)
        entry_last_name.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_last_name.insert(0, current_values[3])

        ttk.Label(input_frame, text="عنوان:").grid(row=3, column=0, sticky=tk.E, padx=5, pady=2)
        entry_title = ttk.Entry(input_frame, width=40)
        entry_title.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_title.insert(0, current_values[4])

        ttk.Label(input_frame, text="تلفن:").grid(row=4, column=0, sticky=tk.E, padx=5, pady=2)
        entry_phone = ttk.Entry(input_frame, width=40)
        entry_phone.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_phone.insert(0, current_values[5])

        ttk.Label(input_frame, text="ایمیل:").grid(row=5, column=0, sticky=tk.E, padx=5, pady=2)
        entry_email = ttk.Entry(input_frame, width=40)
        entry_email.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=2)
        entry_email.insert(0, current_values[6])

        ttk.Label(input_frame, text="یادداشت‌ها:").grid(row=6, column=0, sticky=tk.E, padx=5, pady=2)
        text_notes = tk.Text(input_frame, width=40, height=5, font=("Arial", 10))
        text_notes.grid(row=6, column=1, sticky=tk.EW, padx=5, pady=2)
        text_notes.insert("1.0", current_values[7])

        input_frame.grid_columnconfigure(1, weight=1)

        def update_contact_data():
            org_name = combo_org.get()
            first_name = entry_first_name.get().strip()
            last_name = entry_last_name.get().strip()
            title = entry_title.get().strip()
            phone = entry_phone.get().strip()
            email = entry_email.get().strip()
            notes = text_notes.get("1.0", tk.END).strip()

            if not first_name or not last_name:
                messagebox.showerror("خطا", "نام و نام خانوادگی مخاطب نمی‌تواند خالی باشد.", parent=edit_contact_window)
                return
            
            organization_id = None
            if org_name and org_name != "---":
                organization_id = org_map.get(org_name)


            if update_contact(contact_id, organization_id, first_name, last_name, title, phone, email, notes):
                messagebox.showinfo("موفقیت", "مخاطب با موفقیت ویرایش شد.", parent=edit_contact_window)
                populate_contacts_treeview(organization_id, "", contact_treeview_ref, status_bar_ref)
                populate_org_contact_combos_callback()
                edit_contact_window.destroy()
                if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{first_name} {last_name}' ویرایش شد.")
            else:
                messagebox.showerror("خطا", "خطا در ویرایش مخاطب.", parent=edit_contact_window)

        save_button = ttk.Button(edit_contact_window, text="ذخیره تغییرات", command=update_contact_data)
        save_button.pack(pady=10)
    else:
        messagebox.showwarning("انتخاب نشده", "برای ویرایش یک مخاطب، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)


def on_delete_contact_button(root_window_ref, contact_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    selected_item = contact_treeview_ref.selection()
    if selected_item:
        contact_id = contact_treeview_ref.item(selected_item[0], 'iid')
        contact_name = f"{contact_treeview_ref.item(selected_item[0], 'values')[2]} {contact_treeview_ref.item(selected_item[0], 'values')[3]}"
        
        confirm = messagebox.askyesno("حذف مخاطب", f"آیا مطمئن هستید که می‌خواهید مخاطب '{contact_name}' را حذف کنید؟", parent=root_window_ref)
        if confirm:
            if delete_contact(contact_id):
                messagebox.showinfo("موفقیت", "مخاطب با موفقیت حذف شد.", parent=root_window_ref)
                # Refresh contacts treeview, ensuring the organization_id filter is maintained if one was active
                # This requires getting the currently selected organization in the org_treeview_ref
                selected_org_items = org_treeview_ref.selection()
                if selected_org_items:
                    selected_org_id = org_treeview_ref.item(selected_org_items[0], 'iid')
                    populate_contacts_treeview(selected_org_id, "", contact_treeview_ref, status_bar_ref)
                else:
                    # Default to None for organization_id
                    # If a more precise refresh is needed here, org_treeview_ref would need to be passed.
                    populate_contacts_treeview(None, "", contact_treeview_ref, status_bar_ref) 
                populate_org_contact_combos_callback()
                if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{contact_name}' حذف شد.")
            else:
                messagebox.showerror("خطا", "خطا در حذف مخاطب.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای حذف یک مخاطب، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)

def on_organization_select(event, org_treeview_ref, contact_treeview_ref, status_bar_ref):
    """Filters contacts based on the selected organization."""
    selected_items = org_treeview_ref.selection()
    if selected_items:
        selected_org_id = org_treeview_ref.item(selected_items[0], 'iid')
        populate_contacts_treeview(organization_id=selected_org_id, search_term="", 
                                   contact_treeview_ref=contact_treeview_ref, status_bar_ref=status_bar_ref)
        if status_bar_ref: status_bar_ref.config(text=f"مخاطبین سازمان انتخاب شده فیلتر شدند.")
    else:
        populate_contacts_treeview(organization_id=None, search_term="", 
                                   contact_treeview_ref=contact_treeview_ref, status_bar_ref=status_bar_ref)
        if status_bar_ref: status_bar_ref.config(text="فیلتر مخاطبین برداشته شد.")
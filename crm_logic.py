import tkinter as tk
from tkinter import messagebox, ttk
from database import (
    get_organizations, insert_organization, update_organization, delete_organization,
    get_contacts, insert_contact, update_contact, delete_contact
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
                contact['first_name'],
                contact['last_name'],
                contact['title'],
                contact['phone'],
                contact['email'],
                contact['notes']
            ), iid=contact['id'])
        if status_bar_ref: status_bar_ref.config(text=f"نمایش {len(contacts)} مخاطب.")


# Functions for Add/Edit/Delete Organization/Contact - modified to accept UI references
def on_add_organization_button(root_window_ref, org_treeview_ref, status_bar_ref):
    """Handles adding a new organization."""
    add_org_window = tk.Toplevel(root_window_ref)
    add_org_window.title("افزودن سازمان جدید")
    add_org_window.grab_set() # Make window modal
    add_org_window.transient(root_window_ref)

    labels = ["نام سازمان:", "صنعت:", "تلفن:", "ایمیل:", "آدرس:", "توضیحات:"]
    entries = {}

    for i, text in enumerate(labels):
        ttk.Label(add_org_window, text=text).grid(row=i, column=0, padx=5, pady=5, sticky="w")
        entry = ttk.Entry(add_org_window, justify='right', width=40)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        entries[text] = entry

    def save_org():
        name = entries["نام سازمان:"].get().strip()
        industry = entries["صنعت:"].get().strip()
        phone = entries["تلفن:"].get().strip()
        email = entries["ایمیل:"].get().strip()
        address = entries["آدرس:"].get().strip()
        description = entries["توضیحات:"].get().strip()

        if not name:
            messagebox.showwarning("ورودی نامعتبر", "نام سازمان نمی‌تواند خالی باشد.", parent=add_org_window)
            return

        if insert_organization(name, industry, phone, email, address, description):
            messagebox.showinfo("موفقیت", "سازمان با موفقیت اضافه شد.", parent=add_org_window)
            populate_organizations_treeview("", org_treeview_ref, status_bar_ref) # Refresh the treeview
            if status_bar_ref: status_bar_ref.config(text=f"سازمان '{name}' اضافه شد.")
            add_org_window.destroy()
        else:
            messagebox.showerror("خطا", "خطا در افزودن سازمان. ممکن است نام سازمان تکراری باشد.", parent=add_org_window)

    ttk.Button(add_org_window, text="ذخیره", command=save_org).grid(row=len(labels), column=1, padx=5, pady=10, sticky="e")
    add_org_window.mainloop() # Use mainloop for modal behavior

def on_edit_organization_button(root_window_ref, org_treeview_ref, status_bar_ref):
    """Handles editing an existing organization."""
    selected_item = org_treeview_ref.focus()
    if selected_item:
        org_data = org_treeview_ref.item(selected_item, 'values')
        org_id = org_data[0] # ID is the first value

        edit_org_window = tk.Toplevel(root_window_ref)
        edit_org_window.title("ویرایش سازمان")
        edit_org_window.grab_set()
        edit_org_window.transient(root_window_ref)

        labels = ["نام سازمان:", "صنعت:", "تلفن:", "ایمیل:", "آدرس:", "توضیحات:"]
        entries = {}
        initial_values = {
            "نام سازمان:": org_data[1],
            "صنعت:": org_data[2],
            "تلفن:": org_data[3],
            "ایمیل:": org_data[4],
            "آدرس:": org_data[5],
            "توضیحات:": org_data[6]
        }

        for i, text in enumerate(labels):
            ttk.Label(edit_org_window, text=text).grid(row=i, column=0, padx=5, pady=5, sticky="w")
            entry = ttk.Entry(edit_org_window, justify='right', width=40)
            entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
            entry.insert(0, initial_values.get(text, ""))
            entries[text] = entry

        def update_org():
            name = entries["نام سازمان:"].get().strip()
            industry = entries["صنعت:"].get().strip()
            phone = entries["تلفن:"].get().strip()
            email = entries["ایمیل:"].get().strip()
            address = entries["آدرس:"].get().strip()
            description = entries["توضیحات:"].get().strip()

            if not name:
                messagebox.showwarning("ورودی نامعتبر", "نام سازمان نمی‌تواند خالی باشد.", parent=edit_org_window)
                return

            if update_organization(org_id, name, industry, phone, email, address, description):
                messagebox.showinfo("موفقیت", "سازمان با موفقیت ویرایش شد.", parent=edit_org_window)
                populate_organizations_treeview("", org_treeview_ref, status_bar_ref) # Refresh the treeview
                if status_bar_ref: status_bar_ref.config(text=f"سازمان '{name}' ویرایش شد.")
                edit_org_window.destroy()
            else:
                messagebox.showerror("خطا", "خطا در ویرایش سازمان. ممکن است نام سازمان تکراری باشد.", parent=edit_org_window)

        ttk.Button(edit_org_window, text="ذخیره تغییرات", command=update_org).grid(row=len(labels), column=1, padx=5, pady=10, sticky="e")
        edit_org_window.mainloop()
    else:
        messagebox.showwarning("انتخاب نشده", "برای ویرایش یک سازمان، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)

def on_delete_organization_button(root_window_ref, org_treeview_ref, contact_treeview_ref, status_bar_ref):
    """Handles deleting an organization."""
    selected_item = org_treeview_ref.focus()
    if selected_item:
        org_id = org_treeview_ref.item(selected_item, 'iid')
        org_name = org_treeview_ref.item(selected_item, 'values')[1]

        if messagebox.askyesno("تایید حذف", f"آیا از حذف سازمان '{org_name}' مطمئن هستید؟\n(این عمل مخاطبان مرتبط را نیز حذف می‌کند)", parent=root_window_ref):
            if delete_organization(org_id):
                messagebox.showinfo("موفقیت", f"سازمان '{org_name}' با موفقیت حذف شد.", parent=root_window_ref)
                populate_organizations_treeview("", org_treeview_ref, status_bar_ref) # Refresh organizations
                populate_contacts_treeview(None, "", contact_treeview_ref, status_bar_ref) # Refresh contacts (all, as selected org is gone)
                if status_bar_ref: status_bar_ref.config(text=f"سازمان '{org_name}' حذف شد.")
            else:
                messagebox.showerror("خطا", "خطا در حذف سازمان.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای حذف یک سازمان، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)


def on_add_contact_button(root_window_ref, contact_treeview_ref, org_treeview_ref, status_bar_ref):
    """Handles adding a new contact."""
    add_contact_window = tk.Toplevel(root_window_ref)
    add_contact_window.title("افزودن مخاطب جدید")
    add_contact_window.grab_set()
    add_contact_window.transient(root_window_ref)

    labels = ["نام:", "نام خانوادگی:", "عنوان/سمت:", "تلفن:", "ایمیل:", "یادداشت‌ها:"]
    entries = {}

    # Organization selection for new contact
    ttk.Label(add_contact_window, text="سازمان مرتبط:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    org_combo = ttk.Combobox(add_contact_window, state="readonly", justify='right', width=37)
    org_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    
    # Use direct import if needed, otherwise ensure database is in module scope
    import database # Temporarily import here if not globally imported in crm_logic
    conn = database.get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM Organizations ORDER BY name")
    orgs = cursor.fetchall()
    conn.close()

    org_names = [org['name'] for org in orgs]
    org_id_map = {org['name']: org['id'] for org in orgs}
    org_combo['values'] = ["--- انتخاب سازمان ---"] + org_names
    org_combo.set("--- انتخاب سازمان ---")

    # Pre-select organization if one is active in the CRM tab
    selected_org_item = org_treeview_ref.focus()
    if selected_org_item:
        pre_selected_org_name = org_treeview_ref.item(selected_org_item, 'values')[1]
        if pre_selected_org_name in org_names:
            org_combo.set(pre_selected_org_name)


    for i, text in enumerate(labels):
        ttk.Label(add_contact_window, text=text).grid(row=i+1, column=0, padx=5, pady=5, sticky="w")
        if text == "یادداشت‌ها:":
            entry = tk.Text(add_contact_window, wrap=tk.WORD, height=5, width=30)
            entry.grid(row=i+1, column=1, padx=5, pady=5, sticky="ew")
        else:
            entry = ttk.Entry(add_contact_window, justify='right', width=40)
            entry.grid(row=i+1, column=1, padx=5, pady=5, sticky="ew")
        entries[text] = entry

    def save_contact():
        selected_org_name = org_combo.get()
        organization_id = org_id_map.get(selected_org_name) if selected_org_name != "--- انتخاب سازمان ---" else None
        
        first_name = entries["نام:"].get().strip()
        last_name = entries["نام خانوادگی:"].get().strip()
        title = entries["عنوان/سمت:"].get().strip()
        phone = entries["تلفن:"].get().strip()
        email = entries["ایمیل:"].get().strip()
        notes = entries["یادداشت‌ها:"].get("1.0", tk.END).strip()

        if not first_name or not last_name:
            messagebox.showwarning("ورودی نامعتبر", "نام و نام خانوادگی نمی‌تواند خالی باشد.", parent=add_contact_window)
            return

        if insert_contact(organization_id, first_name, last_name, title, phone, email, notes):
            messagebox.showinfo("موفقیت", "مخاطب با موفقیت اضافه شد.", parent=add_contact_window)
            # Refresh contacts based on the currently selected organization in the CRM tab
            current_selected_org_item = org_treeview_ref.focus()
            if current_selected_org_item:
                current_org_id = org_treeview_ref.item(current_selected_org_item, 'iid')
                populate_contacts_treeview(current_org_id, "", contact_treeview_ref, status_bar_ref)
            else:
                populate_contacts_treeview(None, "", contact_treeview_ref, status_bar_ref)
            
            if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{first_name} {last_name}' اضافه شد.")
            add_contact_window.destroy()
        else:
            messagebox.showerror("خطا", "خطا در افزودن مخاطب.", parent=add_contact_window)

    ttk.Button(add_contact_window, text="ذخیره", command=save_contact).grid(row=len(labels)+1, column=1, padx=5, pady=10, sticky="e")
    add_contact_window.mainloop()

def on_edit_contact_button(root_window_ref, contact_treeview_ref, org_treeview_ref, status_bar_ref):
    """Handles editing an existing contact."""
    selected_item = contact_treeview_ref.focus()
    if selected_item:
        contact_data = contact_treeview_ref.item(selected_item, 'values')
        contact_id = contact_data[0]

        edit_contact_window = tk.Toplevel(root_window_ref)
        edit_contact_window.title("ویرایش مخاطب")
        edit_contact_window.grab_set()
        edit_contact_window.transient(root_window_ref)

        import database # Temporarily import here if not globally imported in crm_logic
        conn = database.get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT organization_id, first_name, last_name, title, phone, email, notes FROM Contacts WHERE id = ?", (contact_id,))
        current_contact_data = cursor.fetchone()
        
        cursor.execute("SELECT id, name FROM Organizations ORDER BY name")
        orgs = cursor.fetchall()
        conn.close()

        org_names = [org['name'] for org in orgs]
        org_id_map = {org['name']: org['id'] for org in orgs}
        id_org_map = {org['id']: org['name'] for org in orgs} # Reverse map for initial setting

        ttk.Label(edit_contact_window, text="سازمان مرتبط:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        org_combo = ttk.Combobox(edit_contact_window, state="readonly", justify='right', width=37)
        org_combo.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        org_combo['values'] = ["--- انتخاب سازمان ---"] + org_names
        
        initial_org_id = current_contact_data['organization_id']
        if initial_org_id and initial_org_id in id_org_map:
            org_combo.set(id_org_map[initial_org_id])
        else:
            org_combo.set("--- انتخاب سازمان ---")


        labels = ["نام:", "نام خانوادگی:", "عنوان/سمت:", "تلفن:", "ایمیل:", "یادداشت‌ها:"]
        entries = {}
        initial_values = {
            "نام:": current_contact_data['first_name'],
            "نام خانوادگی:": current_contact_data['last_name'],
            "عنوان/سمت:": current_contact_data['title'],
            "تلفن:": current_contact_data['phone'],
            "ایمیل:": current_contact_data['email'],
            "یادداشت‌ها:": current_contact_data['notes']
        }

        for i, text in enumerate(labels):
            ttk.Label(edit_contact_window, text=text).grid(row=i+1, column=0, padx=5, pady=5, sticky="w")
            if text == "یادداشت‌ها:":
                entry = tk.Text(edit_contact_window, wrap=tk.WORD, height=5, width=30)
                entry.grid(row=i+1, column=1, padx=5, pady=5, sticky="ew")
                entry.insert("1.0", initial_values.get(text, ""))
            else:
                entry = ttk.Entry(edit_contact_window, justify='right', width=40)
                entry.grid(row=i+1, column=1, padx=5, pady=5, sticky="ew")
                entry.insert(0, initial_values.get(text, ""))
            entries[text] = entry

        def update_contact_data():
            selected_org_name = org_combo.get()
            organization_id = org_id_map.get(selected_org_name) if selected_org_name != "--- انتخاب سازمان ---" else None
            
            first_name = entries["نام:"].get().strip()
            last_name = entries["نام خانوادگی:"].get().strip()
            title = entries["عنوان/سمت:"].get().strip()
            phone = entries["تلفن:"].get().strip()
            email = entries["ایمیل:"].get().strip()
            notes = entries["یادداشت‌ها:"].get("1.0", tk.END).strip()

            if not first_name or not last_name:
                messagebox.showwarning("ورودی نامعتبر", "نام و نام خانوادگی نمی‌تواند خالی باشد.", parent=edit_contact_window)
                return

            if update_contact(contact_id, organization_id, first_name, last_name, title, phone, email, notes):
                messagebox.showinfo("موفقیت", "مخاطب با موفقیت ویرایش شد.", parent=edit_contact_window)
                # Determine which org's contacts to refresh based on the current selection in CRM tab
                selected_crm_org_item = org_treeview_ref.focus()
                if selected_crm_org_item:
                    current_crm_org_id = org_treeview_ref.item(selected_crm_org_item, 'iid')
                    populate_contacts_treeview(current_crm_org_id, "", contact_treeview_ref, status_bar_ref)
                else: # No organization selected, refresh all contacts
                    populate_contacts_treeview(None, "", contact_treeview_ref, status_bar_ref)
                
                if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{first_name} {last_name}' ویرایش شد.")
                edit_contact_window.destroy()
            else:
                messagebox.showerror("خطا", "خطا در ویرایش مخاطب.", parent=edit_contact_window)

        ttk.Button(edit_contact_window, text="ذخیره تغییرات", command=update_contact_data).grid(row=len(labels)+1, column=1, padx=5, pady=10, sticky="e")
        edit_contact_window.mainloop()
    else:
        messagebox.showwarning("انتخاب نشده", "برای ویرایش یک مخاطب، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)


def on_delete_contact_button(root_window_ref, contact_treeview_ref, status_bar_ref):
    """Handles deleting a contact."""
    selected_item = contact_treeview_ref.focus()
    if selected_item:
        contact_id = contact_treeview_ref.item(selected_item, 'iid')
        contact_name = f"{contact_treeview_ref.item(selected_item, 'values')[1]} {contact_treeview_ref.item(selected_item, 'values')[2]}"

        if messagebox.askyesno("تایید حذف", f"آیا از حذف مخاطب '{contact_name}' مطمئن هستید؟", parent=root_window_ref):
            if delete_contact(contact_id):
                messagebox.showinfo("موفقیت", f"مخاطب '{contact_name}' با موفقیت حذف شد.", parent=root_window_ref)
                # Refresh contacts based on current selection (if any) or all
                # This function does not have org_treeview_ref, so we default to None for organization_id
                # If a more precise refresh is needed here, org_treeview_ref would need to be passed.
                populate_contacts_treeview(None, "", contact_treeview_ref, status_bar_ref) 
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
                                   contact_treeview_ref=contact_treeview_ref, status_bar_ref=status_bar_ref) # Show all contacts if no organization is selected
        if status_bar_ref: status_bar_ref.config(text=f"نمایش همه مخاطبین.")
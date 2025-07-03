import tkinter as tk
from tkinter import messagebox, ttk
from database import (
    get_db_connection,
    get_organizations_from_db,  # Corrected import name
    insert_organization,
    update_organization,
    delete_organization,
    get_contacts_from_db,      # Corrected import name
    insert_contact,
    update_contact,
    delete_contact,
    get_organization_by_id,
    get_contact_by_id
)

def populate_organizations_treeview(search_term="", org_treeview_ref=None, status_bar_ref=None):
    """Populates the organizations Treeview with data from the database."""
    if org_treeview_ref: # Ensure org_treeview_ref is initialized
        for item in org_treeview_ref.get_children():
            org_treeview_ref.delete(item)
        
        organizations = get_organizations_from_db(search_term) # Corrected function call
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
        
        contacts = get_contacts_from_db(organization_id=organization_id, search_term=search_term) # Corrected function call
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


def on_add_organization_button(root_window_ref, entry_name, entry_industry, entry_phone, entry_email, entry_address, text_description, org_treeview_ref, status_bar_ref):
    """Handles adding a new organization."""
    name = entry_name.get().strip()
    industry = entry_industry.get().strip()
    phone = entry_phone.get().strip()
    email = entry_email.get().strip()
    address = entry_address.get().strip()
    description = text_description.get("1.0", tk.END).strip()

    if not name:
        messagebox.showwarning("ورودی ناقص", "لطفاً نام سازمان را وارد کنید.", parent=root_window_ref)
        return

    if insert_organization(name, industry, phone, email, address, description):
        messagebox.showinfo("موفقیت", "سازمان با موفقیت اضافه شد.", parent=root_window_ref)
        # Clear fields
        entry_name.delete(0, tk.END)
        entry_industry.delete(0, tk.END)
        entry_phone.delete(0, tk.END)
        entry_email.delete(0, tk.END)
        entry_address.delete(0, tk.END)
        text_description.delete("1.0", tk.END)
        populate_organizations_treeview(org_treeview_ref=org_treeview_ref, status_bar_ref=status_bar_ref)
        if status_bar_ref: status_bar_ref.config(text=f"سازمان '{name}' اضافه شد.")
    else:
        messagebox.showerror("خطا", "خطا در افزودن سازمان.", parent=root_window_ref)

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
            populate_org_contact_combos_callback()
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
        org_name = org_treeview_ref.item(selected_items[0], 'values')[1] # Assuming name is at index 1
        
        response = messagebox.askyesno("تأیید حذف", f"آیا مطمئن هستید که می‌خواهید سازمان '{org_name}' را حذف کنید؟\n\nتمام مخاطبین و نامه‌های مرتبط با این سازمان (در صورت وجود) نیز ممکن است تحت تأثیر قرار گیرند.", parent=root_window_ref)
        if response:
            if delete_organization(org_id):
                messagebox.showinfo("موفقیت", "سازمان با موفقیت حذف شد.", parent=root_window_ref)
                populate_organizations_treeview(org_treeview_ref=org_treeview_ref, status_bar_ref=status_bar_ref)
                # After deleting an org, refresh contacts to show all or based on new selection
                populate_contacts_treeview(None, "", contact_treeview_ref, status_bar_ref) 
                populate_org_contact_combos_callback()
                if status_bar_ref: status_bar_ref.config(text=f"سازمان '{org_name}' حذف شد.")
            else:
                messagebox.showerror("خطا", "خطا در حذف سازمان.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای حذف یک سازمان، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)

def on_add_contact_button(root_window_ref, org_id_entry, entry_first_name, entry_last_name, entry_title, entry_phone, entry_email, text_notes, contact_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    """Handles adding a new contact."""
    organization_id_str = org_id_entry.get().strip()
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

    if insert_contact(organization_id, first_name, last_name, title, phone, email, notes):
        messagebox.showinfo("موفقیت", "مخاطب با موفقیت اضافه شد.", parent=root_window_ref)
        # Clear fields
        # org_id_entry.delete(0, tk.END) # Don't clear org_id_entry if user wants to add multiple contacts for same org
        entry_first_name.delete(0, tk.END)
        entry_last_name.delete(0, tk.END)
        entry_title.delete(0, tk.END)
        entry_phone.delete(0, tk.END)
        entry_email.delete(0, tk.END)
        text_notes.delete("1.0", tk.END)
        populate_contacts_treeview(organization_id, "", contact_treeview_ref, status_bar_ref)
        populate_org_contact_combos_callback()
        if status_bar_ref: status_bar_ref.config(text=f"مخاطب '{first_name} {last_name}' اضافه شد.")
    else:
        messagebox.showerror("خطا", "خطا در افزودن مخاطب.", parent=root_window_ref)

def on_edit_contact_button(root_window_ref, org_id_entry, entry_first_name, entry_last_name, entry_title, entry_phone, entry_email, text_notes, contact_treeview_ref, status_bar_ref, populate_org_contact_combos_callback):
    """Handles editing an existing contact."""
    selected_items = contact_treeview_ref.selection()
    if selected_items:
        contact_id = contact_treeview_ref.item(selected_items[0], 'iid')
        organization_id_str = org_id_entry.get().strip()
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
            populate_org_contact_combos_callback()
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
        contact_name = f"{contact_treeview_ref.item(selected_items[0], 'values')[1]} {contact_treeview_ref.item(selected_items[0], 'values')[2]}" # First and Last Name
        
        response = messagebox.askyesno("تأیید حذف", f"آیا مطمئن هستید که می‌خواهید مخاطب '{contact_name}' را حذف کنید؟", parent=root_window_ref)
        if response:
            if delete_contact(contact_id):
                messagebox.showinfo("موفقیت", "مخاطب با موفقیت حذف شد.", parent=root_window_ref)
                # After deleting a contact, refresh contacts view for the currently selected organization_id
                # This depends on how populate_contacts_treeview is designed to filter.
                # If it should show all contacts after delete, pass None for organization_id
                # If it should maintain the current filter, you need to retrieve the current organization_id
                # For simplicity, refreshing all contacts for now.
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
        # If no organization is selected, show all contacts
        populate_contacts_treeview(organization_id=None, search_term="", 
                                   contact_treeview_ref=contact_treeview_ref, status_bar_ref=status_bar_ref)
        if status_bar_ref: status_bar_ref.config(text="فیلتر مخاطبین برداشته شد. نمایش همه مخاطبین.")
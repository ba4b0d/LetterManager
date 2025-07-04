import os
from tkinter import messagebox, END
from database import get_letters_from_db, get_letter_by_code
from helpers import sort_column
import tkinter as tk # Needed for tk.END

def update_history_treeview(search_term="", history_treeview_ref=None, status_bar_ref=None, letter_types_map=None):
    """
    Populates the letter history Treeview with data from the database.
    Now accepts letter_types_map to convert raw letter types to Persian display names.
    """
    if history_treeview_ref: # Ensure history_treeview_ref is initialized
        for item in history_treeview_ref.get_children():
            history_treeview_ref.delete(item)

        letters = get_letters_from_db(search_term=search_term)

        for letter_data in letters:
            org_name = letter_data['organization_name'] if letter_data['organization_name'] else "---"
            contact_name = f"{letter_data['first_name']} {letter_data['last_name']}" if (letter_data['first_name'] and letter_data['last_name']) else "---"

            # Get the raw letter type from the database (e.g., "FIN", "HR")
            # Ensure 'letter_type_raw' is correctly fetched by database.py
            raw_letter_type = letter_data['letter_type_raw'] 
            
            # Convert to Persian display name using the provided map
            # Use .get() with a fallback to raw_letter_type in case the key is not found
            # This makes it robust even if letter_types_map is None or incomplete
            display_letter_type = letter_types_map.get(raw_letter_type, raw_letter_type) if letter_types_map else raw_letter_type

            display_values = (
                letter_data['letter_code_persian'],
                display_letter_type, # Use the converted Persian type
                letter_data['date_shamsi_persian'],
                letter_data['subject'],
                org_name,
                contact_name
            )
            history_treeview_ref.insert("", tk.END, values=display_values, iid=letter_data['id'])
        if status_bar_ref: status_bar_ref.config(text=f"نمایش {len(letters)} نامه.")


def on_search_archive_button(search_entry_text, history_treeview_ref, status_bar_ref, letter_types_map=None):
    """Handles the search button click for archive tab."""
    # Pass letter_types_map to update_history_treeview
    update_history_treeview(search_term=search_entry_text, history_treeview_ref=history_treeview_ref, status_bar_ref=status_bar_ref, letter_types_map=letter_types_map)

def on_open_letter_button(history_treeview_ref, root_window_ref, status_bar_ref):
    """Opens the selected letter file."""
    selected_item = history_treeview_ref.focus()
    if selected_item:
        letter_code_from_treeview = history_treeview_ref.item(selected_item, 'values')[0] # Assuming code is the first column
        letter_data_from_db = get_letter_by_code(letter_code_from_treeview)

        if letter_data_from_db and os.path.exists(letter_data_from_db['file_path']):
            try:
                os.startfile(letter_data_from_db['file_path'])
                messagebox.showinfo("باز کردن نامه", f"نامه '{letter_code_from_treeview}' با موفقیت باز شد.", parent=root_window_ref)
                if status_bar_ref: status_bar_ref.config(text=f"نامه '{letter_code_from_treeview}' باز شد.")
            except Exception as e:
                messagebox.showerror("خطا در باز کردن فایل",
                                    f"خطا در باز کردن فایل نامه '{letter_code_from_treeview}':\n\n"
                                    f"پیام خطا: {e}\n\n"
                                    "اطمینان حاصل کنید که فایل Word بسته است و دسترسی‌های لازم برای باز کردن آن را دارید."
                                    "\nهمچنین ممکن است برنامه مربوطه (مانند Microsoft Word) روی سیستم شما نصب نباشد.", parent=root_window_ref)
        else:
            messagebox.showerror("خطا: فایل یافت نشد",
                                f"فایل مربوط به شماره نامه '{letter_code_from_treeview}' در مسیر ذخیره شده یافت نشد یا حذف شده است.\n\n"
                                f"مسیر ذخیره شده: '{letter_data_from_db['file_path'] if letter_data_from_db else 'نامشخص'}'\n\n"
                                "لطفاً از وجود فایل در مسیر ذکر شده مطمئن شوید.", parent=root_window_ref)
    else:
        messagebox.showwarning("انتخاب نشده", "برای باز کردن یک نامه، ابتدا آن را از لیست انتخاب کنید.", parent=root_window_ref)


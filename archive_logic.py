import os
from tkinter import messagebox, END
from database import get_letters_from_db, get_letter_by_code
from helpers import sort_column
import tkinter as tk # Needed for tk.END

def update_history_treeview(search_term="", history_treeview_ref=None, status_bar_ref=None):
    """Populates the letter history Treeview with data from the database."""
    if history_treeview_ref: # Ensure history_treeview_ref is initialized
        for item in history_treeview_ref.get_children():
            history_treeview_ref.delete(item)

        letters = get_letters_from_db(search_term=search_term)

        for letter_data in letters:
            org_name = letter_data['organization_name'] if letter_data['organization_name'] else "---"
            contact_name = f"{letter_data['first_name']} {letter_data['last_name']}" if (letter_data['first_name'] and letter_data['last_name']) else "---"

            display_values = (
                letter_data['letter_code_persian'],
                letter_data['letter_type_persian'], # CHANGED: Corrected to use 'letter_type_persian'
                letter_data['date_shamsi_persian'],
                letter_data['subject'],
                org_name,
                contact_name
            )
            history_treeview_ref.insert("", tk.END, values=display_values, iid=letter_data['letter_code'])
        if status_bar_ref: status_bar_ref.config(text=f"نمایش {len(letters)} نامه.")


def on_search_archive_button(search_entry_ref, history_treeview_ref, status_bar_ref):
    """Handles the search button click for the letter archive."""
    search_term = search_entry_ref.get().strip()
    update_history_treeview(search_term, history_treeview_ref, status_bar_ref)

def on_open_letter_button(history_treeview_ref, root_window_ref, status_bar_ref):
    """Opens the selected letter file."""
    selected_items = history_treeview_ref.selection()
    if selected_items:
        letter_code_from_treeview = history_treeview_ref.item(selected_items[0], 'iid')
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
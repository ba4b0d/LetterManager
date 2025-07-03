import os
import shutil
import jdatetime
from datetime import datetime
from tkinter import messagebox, filedialog, END, W

from database import get_db_connection, insert_letter, get_letters_from_db 
from helpers import convert_numbers_to_persian, replace_text_in_docx, show_progress_window, hide_progress_window
from settings_manager import company_name, default_save_path, letterhead_template_path, full_company_name


def generate_letter_number(app_instance):
    """Generates a new letter code based on current date and database sequence."""
    year_shamsi = jdatetime.date.today().year
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    next_sequence_number = 1 # Default to 1 if no letters for the year or parsing fails

    # Try to find the highest existing sequence number for the current Shamsi year
    # This query fetches letter codes for the current year and orders them by ID to find the max sequence
    current_year_prefix = str(year_shamsi) + '/'
    cursor.execute("""
        SELECT letter_code FROM Letters
        WHERE date_shamsi_persian LIKE ?
        ORDER BY id DESC LIMIT 1
    """, (current_year_prefix + '%',))
    last_letter = cursor.fetchone()

    if last_letter:
        last_code = last_letter['letter_code'] # e.g., NGRR-FIN-1403-001
        try:
            # Extract the numeric sequence part (e.g., "001" from the example)
            # This assumes the format COMPANY-TYPE-YYYY-NNN
            parts = last_code.split('-')
            if len(parts) == 4:
                last_seq_str = parts[3] # Get the "NNN" part
                last_seq_num = int(last_seq_str)
                next_sequence_number = last_seq_num + 1
        except (ValueError, IndexError):
            # If parsing fails or format is unexpected, next_sequence_number remains 1 (default)
            pass 
            
    conn.close()
    
    # Get letter type abbreviation from selected value
    selected_letter_type_value = app_instance.combo_letter_type.get()
    letter_type_abbr = None
    for abbr, full_name in app_instance.letter_types.items():
        if full_name == selected_letter_type_value:
            letter_type_abbr = abbr
            break

    if not letter_type_abbr:
        letter_type_abbr = "GEN" # Fallback if no type selected or found

    # Format the letter number: COMPANY-TYPE-YYYY-NNN (e.g., NGRR-FIN-1403-001)
    letter_code = f"{company_name}-{letter_type_abbr}-{year_shamsi}-{next_sequence_number:03d}"
    letter_code_persian = convert_numbers_to_persian(letter_code)

    return letter_code, letter_code_persian


def on_generate_letter(app_instance):
    """Handles the letter generation process."""
    org_name = app_instance.combo_org_letter.get()
    contact_full_name = app_instance.combo_contact_letter.get()
    subject = app_instance.entry_subject.get().strip()
    body = app_instance.text_letter_body.get("1.0", END).strip()
    letter_type_display = app_instance.combo_letter_type.get()
    
    # Validation
    if org_name == "---":
        messagebox.showwarning("ورودی ناقص", "لطفاً سازمان مقصد نامه را انتخاب کنید.", parent=app_instance.root)
        return
    if contact_full_name == "---":
        messagebox.showwarning("ورودی ناقص", "لطفاً مخاطب مقصد نامه را انتخاب کنید.", parent=app_instance.root)
        return
    if not subject:
        messagebox.showwarning("ورودی ناقص", "لطفاً موضوع نامه را وارد کنید.", parent=app_instance.root)
        return
    if not body:
        messagebox.showwarning("ورودی ناقص", "لطفاً متن اصلی نامه را وارد کنید.", parent=app_instance.root)
        return
    if not letterhead_template_path or not os.path.exists(letterhead_template_path):
        messagebox.showerror("خطا", "مسیر فایل الگوی سربرگ (Word) در تنظیمات مشخص نشده یا فایل وجود ندارد. لطفاً در تب 'تنظیمات' آن را تنظیم کنید.", parent=app_instance.root)
        return

    app_instance.show_progress("در حال تولید نامه...")

    try:
        # Get Organization ID
        organization_id = app_instance.org_data_map.get(org_name)
        
        # Get Contact ID and details
        contact_data = app_instance.all_contacts_data.get(contact_full_name)
        contact_id = contact_data['id'] if contact_data else None
        
        # Get current date (Shamsi and Gregorian)
        today_j = jdatetime.date.today()
        date_shamsi_persian = convert_numbers_to_persian(f"{today_j.year}/{today_j.month:02d}/{today_j.day:02d}")
        date_gregorian = datetime.now().strftime("%Y-%m-%d")

        # Generate letter number
        letter_code, letter_code_persian = generate_letter_number(app_instance) # Pass app_instance

        # Determine letter_type abbreviation
        letter_type_abbr = None
        for abbr, full_name in app_instance.letter_types.items():
            if full_name == letter_type_display:
                letter_type_abbr = abbr
                break
        if not letter_type_abbr:
            letter_type_abbr = "GEN" # Fallback

        # Define replacements for the Word template
        replacements = {
            "[[DATE]]": date_shamsi_persian,
            "[[CODE]]": letter_code_persian,
            "[[ORGANIZATION_NAME]]": org_name,
            "[[CONTACT_NAME]]": contact_full_name,
            "[[SUBJECT]]": subject,
            "[[BODY]]": body,
            "[[COMPANY_NAME]]": full_company_name # Use full_company_name from settings
        }

        # Create target directory if it doesn't exist
        if not os.path.exists(default_save_path):
            os.makedirs(default_save_path)

        # Define new file path
        file_name = f"{letter_code} - {subject}.docx"
        new_file_path = os.path.join(default_save_path, file_name)

        # Copy template and replace text
        shutil.copyfile(letterhead_template_path, new_file_path)
        replace_text_in_docx(new_file_path, replacements)

        # Insert letter record into database
        if not insert_letter(letter_code, letter_code_persian, letter_type_abbr, 
                             date_gregorian, date_shamsi_persian, subject, body, 
                             organization_id, contact_id, new_file_path):
            messagebox.showwarning("هشدار", "نامه با موفقیت تولید شد، اما در ذخیره آن در پایگاه داده مشکلی پیش آمد.", parent=app_instance.root)
            # Still proceed to open file, but alert user
        
        messagebox.showinfo("عملیات موفق", f"فایل با نام {file_name} در مسیر '{default_save_path}' ذخیره و محتوای آن بروزرسانی شد.", parent=app_instance.root)
        
        try:
            os.startfile(new_file_path)
        except Exception as open_error:
            messagebox.showwarning("هشدار", f"فایل '{file_name}' با موفقیت کپی و ویرایش شد، اما در باز کردن آن خطایی رخ داد: {open_error}", parent=app_instance.root)

        # Clear form fields and update archive
        app_instance.entry_subject.delete(0, END)
        app_instance.text_letter_body.delete("1.0", END)
        app_instance.combo_org_letter.set("---")
        app_instance.combo_contact_letter.set("---")
        app_instance.combo_letter_type.set(app_instance.letter_types["FIN"]) # Reset to default type
        
        if app_instance.status_bar: app_instance.status_bar.config(text="شماره نامه جدید با موفقیت ایجاد شد.")
        app_instance.update_history_treeview("") # Update the archive tab
        app_instance.populate_org_contact_combos() # Re-populate combos to reflect any changes if new org/contact was added elsewhere

    except Exception as e:
        messagebox.showerror("خطا در تولید نامه", f"خطایی در فرآیند تولید نامه رخ داد: {e}", parent=app_instance.root)
        print(f"DEBUG: Error in on_generate_letter: {e}") # For console debugging
    finally:
        app_instance.hide_progress()
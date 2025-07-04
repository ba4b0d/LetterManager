import os
import shutil
import jdatetime
from datetime import datetime
from tkinter import messagebox, filedialog, END, W

from database import get_db_connection, insert_letter, get_letters_from_db 
from helpers import convert_numbers_to_persian, replace_text_in_docx, show_progress_window, hide_progress_window
from settings_manager import company_name, default_save_path, letterhead_template_path, full_company_name 


def generate_letter_number(letter_type_display, letter_types_map):
    """
    Generates a new letter code based on current date and database sequence.
    Now accepts letter_type_display and letter_types_map directly.
    """
    year_shamsi = jdatetime.date.today().year
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    next_sequence_number = 1 # Default to 1 if no letters for the year or parsing fails

    # Find the maximum existing sequence number for the current Shamsi year
    current_year_prefix = str(year_shamsi) + '/'
    cursor.execute(f"""
        SELECT MAX(CAST(SUBSTR(letter_code_persian, INSTR(letter_code_persian, '-') + INSTR(SUBSTR(letter_code_persian, INSTR(letter_code_persian, '-') + 1), '-') + INSTR(SUBSTR(letter_code_persian, INSTR(letter_code_persian, '-') + INSTR(SUBSTR(letter_code_persian, INSTR(letter_code_persian, '-') + 1), '-') + 1), '-') + 1) AS INTEGER))
        FROM Letters
        WHERE letter_code_persian LIKE '{company_name}-%-{year_shamsi}-%'
    """)

    max_sequence = cursor.fetchone()[0]
    
    if max_sequence is not None:
        next_sequence_number = max_sequence + 1
            
    conn.close()
    
    letter_type_abbr = None
    for abbr, full_name in letter_types_map.items(): # Use passed map
        if full_name == letter_type_display:
            letter_type_abbr = abbr
            break

    if not letter_type_abbr:
        letter_type_abbr = "GEN" 

    letter_code = f"{company_name}-{letter_type_abbr}-{year_shamsi}-{next_sequence_number:03d}"
    letter_code_persian = convert_numbers_to_persian(letter_code)

    return letter_code, letter_code_persian


def on_generate_letter(root_window_ref, status_bar_ref, letter_type_display, subject, body_content, organization_id, contact_id, save_path, letterhead_template, user_id, letter_types_map):
    """
    Handles the letter generation process.
    Now accepts all necessary parameters directly, including letter_types_map.
    """
    
    # Validation (using passed arguments)
    if not subject:
        messagebox.showwarning("ورودی ناقص", "لطفاً موضوع نامه را وارد کنید.", parent=root_window_ref)
        return
    if not body_content:
        messagebox.showwarning("ورودی ناقص", "لطفاً متن اصلی نامه را وارد کنید.", parent=root_window_ref)
        return
    
    # Check if both organization_id and contact_id are None
    if organization_id is None and contact_id is None:
        messagebox.showwarning("ورودی ناقص", "لطفاً حداقل یک سازمان یا مخاطب مقصد نامه را انتخاب کنید.", parent=root_window_ref)
        return
    
    if not letterhead_template or not os.path.exists(letterhead_template):
        messagebox.showerror("خطا", "مسیر فایل الگوی سربرگ (Word) در تنظیمات مشخص نشده یا فایل وجود ندارد. لطفاً در تب 'تنظیمات' آن را تنظیم کنید.", parent=root_window_ref)
        return

    # Use the passed status_bar_ref and root_window_ref for progress and messageboxes
    show_progress_window("در حال تولید نامه...", root_window_ref)

    try:
        # Retrieve organization and contact names for replacements
        org_name = ""
        contact_full_name = ""
        conn = get_db_connection()
        cursor = conn.cursor()

        if organization_id:
            cursor.execute("SELECT name FROM Organizations WHERE id = ?", (organization_id,))
            org_data = cursor.fetchone()
            if org_data:
                org_name = org_data['name']
        
        if contact_id:
            cursor.execute("SELECT first_name, last_name FROM Contacts WHERE id = ?", (contact_id,))
            contact_data = cursor.fetchone()
            if contact_data:
                contact_full_name = f"{contact_data['first_name']} {contact_data['last_name']}"
        
        conn.close()

        today_j = jdatetime.date.today()
        date_shamsi_persian = convert_numbers_to_persian(f"{today_j.year}/{today_j.month:02d}/{today_j.day:02d}")
        date_gregorian = datetime.now().strftime("%Y-%m-%d")

        # Pass letter_type_display and letter_types_map to generate_letter_number
        letter_code, letter_code_persian = generate_letter_number(letter_type_display, letter_types_map) 

        # Determine abbreviation from display name using the passed map
        letter_type_abbr = "GEN" # Default
        for abbr, full_name in letter_types_map.items():
            if full_name == letter_type_display:
                letter_type_abbr = abbr
                break

        replacements = {
            "[[DATE]]": date_shamsi_persian,
            "[[CODE]]": letter_code_persian,
            "[[ORGANIZATION_NAME]]": org_name, 
            "[[CONTACT_NAME]]": contact_full_name, 
            "[[SUBJECT]]": subject,
            "[[BODY]]": body_content, # Use body_content
            "[[COMPANY_NAME]]": full_company_name 
        }

        if not os.path.exists(save_path): # Use passed save_path
            os.makedirs(save_path)

        file_name_subject = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_')).strip()
        file_name = f"{letter_code} - {file_name_subject}.docx"
        new_file_path = os.path.join(save_path, file_name)

        shutil.copyfile(letterhead_template, new_file_path) # Use passed letterhead_template
        replace_text_in_docx(new_file_path, replacements)

        if not insert_letter(
            letter_code_prefix=company_name, # Assuming company_name is global or passed
            letter_code_number=int(letter_code.split('-')[-1]), # Extract number from code
            letter_code_persian=letter_code_persian,
            type=letter_type_abbr, # Use abbreviation
            date_gregorian=date_gregorian,
            date_shamsi_persian=date_shamsi_persian,
            subject=subject, 
            organization_id=organization_id,
            contact_id=contact_id,
            body=body_content,
            file_path=new_file_path,
            user_id=user_id # Use passed user_id
        ):
            messagebox.showwarning("هشدار", "نامه با موفقیت تولید شد، اما در ذخیره آن در پایگاه داده مشکلی پیش آمد.", parent=root_window_ref)
        
        messagebox.showinfo("عملیات موفق", f"فایل با نام {file_name} در مسیر '{save_path}' ذخیره و محتوای آن بروزرسانی شد.", parent=root_window_ref)
        
        try:
            os.startfile(new_file_path)
        except Exception as open_error:
            messagebox.showwarning("هشدار", f"فایل '{file_name}' با موفقیت کپی و ویرایش شد، اما در باز کردن آن خطایی رخ داد: {open_error}", parent=root_window_ref)

    except Exception as e:
        messagebox.showerror("خطا در تولید نامه", f"خطایی در فرآیند تولید نامه رخ داد: {e}", parent=root_window_ref)
        print(f"DEBUG: Error in on_generate_letter: {e}") 
    finally:
        hide_progress_window() # Call directly, no app_instance

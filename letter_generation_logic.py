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

    # Find the maximum existing sequence number for the current Shamsi year
    current_year_prefix = str(year_shamsi) + '/'
    # Changed query to directly select the maximum numerical part to be more robust
    cursor.execute(f"""
        SELECT MAX(CAST(SUBSTR(letter_code, INSTR(letter_code, '-') + INSTR(SUBSTR(letter_code, INSTR(letter_code, '-') + 1), '-') + INSTR(SUBSTR(letter_code, INSTR(letter_code, '-') + INSTR(SUBSTR(letter_code, INSTR(letter_code, '-') + 1), '-') + 1), '-') + 1) AS INTEGER))
        FROM Letters
        WHERE letter_code LIKE '{company_name}-%-{year_shamsi}-%'
    """)
    # Explanation of SUBSTR and INSTR for robust parsing:
    # 1. INSTR(letter_code, '-') finds the first hyphen.
    # 2. SUBSTR(letter_code, INSTR(letter_code, '-') + 1) gets the part after the first hyphen.
    # 3. INSTR(SUBSTR(letter_code, ...), '-') finds the second hyphen in the remaining part.
    # 4. INSTR(SUBSTR(letter_code, INSTR(letter_code, '-') + INSTR(SUBSTR(letter_code, INSTR(letter_code, '-') + 1), '-') + 1), '-') finds the third hyphen in the remaining part.
    # 5. Adding 1 to the final INSTR result gets the position right after the last hyphen before the number.
    # This complex SUBSTR/INSTR chain is to reliably get the "NNN" part from "COMPANY-TYPE-YYYY-NNN".
    # CAST(... AS INTEGER) converts it to a number so MAX works correctly.

    max_sequence = cursor.fetchone()[0]
    
    if max_sequence is not None:
        next_sequence_number = max_sequence + 1
            
    conn.close()
    
    selected_letter_type_value = app_instance.combo_letter_type.get()
    letter_type_abbr = None
    for abbr, full_name in app_instance.letter_types.items():
        if full_name == selected_letter_type_value:
            letter_type_abbr = abbr
            break

    if not letter_type_abbr:
        letter_type_abbr = "GEN" 

    letter_code = f"{company_name}-{letter_type_abbr}-{year_shamsi}-{next_sequence_number:03d}"
    letter_code_persian = convert_numbers_to_persian(letter_code)

    return letter_code, letter_code_persian


def on_generate_letter(app_instance):
    """Handles the letter generation process."""
    org_name = app_instance.combo_org_letter.get()
    contact_full_name = app_instance.combo_contact_letter.get()
    subject = app_instance.entry_subject.get().strip()
    body = app_instance.text_letter_body.get("1.0", END).strip()
    letter_type_display = app_instance.combo_letter_type.get() # This is letter_type_persian
    
    # Validation
    if not subject:
        messagebox.showwarning("ورودی ناقص", "لطفاً موضوع نامه را وارد کنید.", parent=app_instance.root)
        return
    if not body:
        messagebox.showwarning("ورودی ناقص", "لطفاً متن اصلی نامه را وارد کنید.", parent=app_instance.root)
        return
    if org_name == "---" and contact_full_name == "---":
        messagebox.showwarning("ورودی ناقص", "لطفاً حداقل یک سازمان یا مخاطب مقصد نامه را انتخاب کنید.", parent=app_instance.root)
        return
    if not letterhead_template_path or not os.path.exists(letterhead_template_path):
        messagebox.showerror("خطا", "مسیر فایل الگوی سربرگ (Word) در تنظیمات مشخص نشده یا فایل وجود ندارد. لطفاً در تب 'تنظیمات' آن را تنظیم کنید.", parent=app_instance.root)
        return

    app_instance.show_progress("در حال تولید نامه...")

    try:
        organization_id = None
        if org_name != "---": 
            organization_id = app_instance.org_data_map.get(org_name)
        
        contact_id = None
        if contact_full_name != "---": 
            contact_data = app_instance.all_contacts_data.get(contact_full_name)
            contact_id = contact_data['id'] if contact_data else None
        
        today_j = jdatetime.date.today()
        date_shamsi_persian = convert_numbers_to_persian(f"{today_j.year}/{today_j.month:02d}/{today_j.day:02d}")
        date_gregorian = datetime.now().strftime("%Y-%m-%d")

        letter_code, letter_code_persian = generate_letter_number(app_instance) 

        letter_type_abbr = None
        for abbr, full_name in app_instance.letter_types.items():
            if full_name == letter_type_display:
                letter_type_abbr = abbr
                break
        if not letter_type_abbr:
            letter_type_abbr = "GEN" 

        replacements = {
            "[[DATE]]": date_shamsi_persian,
            "[[CODE]]": letter_code_persian,
            "[[ORGANIZATION_NAME]]": org_name if org_name != "---" else "", 
            "[[CONTACT_NAME]]": contact_full_name if contact_full_name != "---" else "", 
            "[[SUBJECT]]": subject,
            "[[BODY]]": body,
            "[[COMPANY_NAME]]": full_company_name 
        }

        if not os.path.exists(default_save_path):
            os.makedirs(default_save_path)

        file_name_subject = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_')).strip()
        file_name = f"{letter_code} - {file_name_subject}.docx"
        new_file_path = os.path.join(default_save_path, file_name)

        shutil.copyfile(letterhead_template_path, new_file_path)
        replace_text_in_docx(new_file_path, replacements)

        if not insert_letter(letter_code, letter_code_persian, letter_type_abbr, letter_type_display,
                             date_gregorian, date_shamsi_persian, subject, 
                             organization_id, contact_id, body, new_file_path):
            messagebox.showwarning("هشدار", "نامه با موفقیت تولید شد، اما در ذخیره آن در پایگاه داده مشکلی پیش آمد.", parent=app_instance.root)
        
        messagebox.showinfo("عملیات موفق", f"فایل با نام {file_name} در مسیر '{default_save_path}' ذخیره و محتوای آن بروزرسانی شد.", parent=app_instance.root)
        
        try:
            os.startfile(new_file_path)
        except Exception as open_error:
            messagebox.showwarning("هشدار", f"فایل '{file_name}' با موفقیت کپی و ویرایش شد، اما در باز کردن آن خطایی رخ داد: {open_error}", parent=app_instance.root)

        app_instance.entry_subject.delete(0, END)
        app_instance.text_letter_body.delete("1.0", END)
        app_instance.combo_org_letter.set("---")
        app_instance.combo_contact_letter.set("---")
        app_instance.combo_letter_type.set(app_instance.letter_types["FIN"]) 
        
        if app_instance.status_bar: app_instance.status_bar.config(text="شماره نامه جدید با موفقیت ایجاد شد.")
        app_instance.update_history_treeview("") 
        app_instance.populate_org_contact_combos() 

    except Exception as e:
        messagebox.showerror("خطا در تولید نامه", f"خطایی در فرآیند تولید نامه رخ داد: {e}", parent=app_instance.root)
        print(f"DEBUG: Error in on_generate_letter: {e}") 
    finally:
        app_instance.hide_progress()
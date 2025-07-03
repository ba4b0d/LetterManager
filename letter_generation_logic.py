import os
import shutil
import jdatetime
from datetime import datetime
from tkinter import messagebox, filedialog, END, W

from database import get_db_connection, insert_letter, get_letters_from_db
from helpers import convert_numbers_to_persian, replace_text_in_docx, show_progress_window, hide_progress_window
from settings_manager import company_name, full_company_name, default_save_path, letterhead_template_path

def generate_letter_number(app_instance):
    """Generates a new letter code based on current date and database sequence."""
    year_shamsi = jdatetime.date.today().year

    conn = get_db_connection()
    cursor = conn.cursor()

    # Get the count of letters for the current Shamsi year and letter type
    # Assuming 'date_shamsi_persian' is stored in 'YYYY/MM/DD' format
    selected_letter_type_persian = app_instance.combo_letter_type.get()

    # Find the abbreviation for the selected Persian type
    letter_type_abbr = next((k for k, v in app_instance.letter_types.items() if v == selected_letter_type_persian), "GEN")

    # CHANGED: Query now correctly uses 'letter_type_persian' column
    cursor.execute("SELECT COUNT(*) FROM Letters WHERE date_shamsi_persian LIKE ? AND letter_type_persian = ?",
                    (str(year_shamsi) + '/%', selected_letter_type_persian))

    count_for_year_and_type = cursor.fetchone()[0]

    conn.close()

    # خط 34: تصحیح اشتباه تایپی از count_for_year_and_1 به count_for_year_and_type
    next_sequence_number = count_for_year_and_type + 1

    # Format the letter number: COMPANY-TYPE-YYYY-NNN (e.g., NGRR-FIN-1403-001)
    letter_code_english = f"{company_name}-{letter_type_abbr}-{year_shamsi}-{next_sequence_number:03d}"

    # Convert only the numbers in the code to Persian for display/storage
    letter_code_persian = convert_numbers_to_persian(letter_code_english)

    # CHANGED: Return letter_type_abbr as well
    return letter_code_english, letter_code_persian, selected_letter_type_persian, letter_type_abbr


def on_generate_letter(app_instance):
    """Handles the letter generation process."""
    subject = app_instance.entry_subject.get().strip()
    body = app_instance.text_letter_body.get("1.0", END).strip()

    if not subject or not body:
        messagebox.showwarning("ورودی ناقص", "لطفاً موضوع و متن نامه را وارد کنید.", parent=app_instance.root)
        return

    selected_org_name = app_instance.combo_org_letter.get()
    selected_contact_name = app_instance.combo_contact_letter.get()

    # Get organization ID and name for DB/replacements
    organization_id = None
    organization_name_for_db = None
    if selected_org_name and selected_org_name != "---":
        organization_id = app_instance.org_data_map.get(selected_org_name)
        organization_name_for_db = selected_org_name

    # Get contact ID and full name for DB/replacements
    contact_id = None
    contact_full_name_for_db = None
    contact_first_name_for_db = None
    contact_last_name_for_db = None
    # CHANGED: Added contact_title to fetch
    contact_title_for_db = None
    if selected_contact_name and selected_contact_name != "---":
        contact_data = app_instance.all_contacts_data.get(selected_contact_name)
        if contact_data:
            contact_id = contact_data['id']
            contact_first_name_for_db = contact_data['first_name']
            contact_last_name_for_db = contact_data['last_name']
            contact_full_name_for_db = f"{contact_first_name_for_db} {contact_last_name_for_db}"
            contact_title_for_db = contact_data.get('title', '') # Safely get title
        else: # Fallback if contact data not found (shouldn't happen if combos are populated correctly)
            contact_full_name_for_db = selected_contact_name
            contact_title_for_db = '' # Default empty title if not found

    # --- DEBUGGING LINE ---
    print(f"Letterhead template path: {letterhead_template_path}")
    # --- END DEBUGGING LINE ---

    # Check if a template file is selected and exists
    if not letterhead_template_path or not os.path.exists(letterhead_template_path):
        messagebox.showerror("خطا در الگو",
                             f"مسیر فایل الگوی سربرگ نامعتبر است یا فایل وجود ندارد: \n\n{letterhead_template_path}\n\nلطفاً مسیر صحیح را در بخش تنظیمات مشخص کنید.",
                             parent=app_instance.root)
        return

    # CHANGED: Receive letter_type_abbr from generate_letter_number
    letter_code_english, letter_code_persian, letter_type_persian, letter_type_abbr = generate_letter_number(app_instance)

    # Get current Shamsi date in Persian for display and database
    current_date_shamsi = jdatetime.date.today()
    date_shamsi_persian = convert_numbers_to_persian(current_date_shamsi.strftime("%Y/%m/%d"))
    date_gregorian = datetime.now().strftime("%Y-%m-%d") # Get current Gregorian date for DB storage

    app_instance.show_progress("در حال تولید نامه...")

    try:
        # Define output file path
        output_folder = default_save_path if default_save_path else os.getcwd()
        if not os.path.exists(output_folder):
            os.makedirs(output_folder) # Ensure the directory exists

        # Sanitize filename (remove characters that are invalid in file paths)
        sanitized_subject = "".join(c for c in subject if c.isalnum() or c in (' ', '-', '_', '.', '(', ')')).strip() # Added more allowed chars
        new_file_name = f"{letter_code_english} - {sanitized_subject}.docx"
        new_file_path = os.path.join(output_folder, new_file_name)

        # Copy the template to the new file path
        shutil.copyfile(letterhead_template_path, new_file_path)

        # Prepare replacements (Ensure your Word template uses these exact placeholders)
        replacements = {
            "[[DATE]]": date_shamsi_persian,
            "[[CODE]]": letter_code_persian,
            "[[ORGANIZATION_NAME]]": organization_name_for_db if organization_name_for_db else "---",
            "[[CONTACT_NAME]]": contact_full_name_for_db if contact_full_name_for_db else "---",
            "[[SUBJECT]]": subject,
            "[[BODY]]": body,
            "[[COMPANY_NAME]]": full_company_name # Use the new full_company_name for the footer
        }

        # Perform replacements in the copied document
        replace_text_in_docx(new_file_path, replacements)

        # Insert letter record into database
        try:
            # CHANGED: Corrected argument order and count for insert_letter
            insert_letter(
                letter_code=letter_code_english,
                letter_code_persian=letter_code_persian,
                letter_type_abbr=letter_type_abbr,         # Correctly pass abbreviation
                letter_type_persian=letter_type_persian,   # Correctly pass Persian type
                date_gregorian=date_gregorian,             # Correctly pass Gregorian date
                date_shamsi_persian=date_shamsi_persian,   # Correctly pass Shamsi date
                subject=subject,
                organization_id=organization_id,
                contact_id=contact_id,
                body_text=body,                            # Correctly pass body text
                file_path=new_file_path
            )
        except Exception as db_error:
            messagebox.showwarning("خطا در ذخیره در پایگاه داده", f"نامه با موفقیت تولید شد اما در ذخیره آن در پایگاه داده مشکلی پیش آمد.\n\nخطا: {db_error}", parent=app_instance.root)
            # Still proceed to open file, but alert user

        messagebox.showinfo("عملیات موفق", f"فایل با نام {new_file_name} در مسیر '{default_save_path}' ذخیره و محتوای آن بروزرسانی شد.", parent=app_instance.root)

        try:
            os.startfile(new_file_path)
        except Exception as open_error:
            messagebox.showwarning("هشدار", f"فایل '{new_file_name}' با موفقیت کپی و ویرایش شد، اما در باز کردن آن خطایی رخ داد: {open_error}", parent=app_instance.root)

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
        messagebox.showerror("خطا در تولید نامه", f"خطایی هنگام تولید نامه رخ داد: {e}", parent=app_instance.root)
    finally:
        app_instance.hide_progress()
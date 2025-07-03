import os
from tkinter import messagebox

# Global variables for settings. These will be loaded from/saved to settings.txt
# and will be accessed by other modules.
company_name = "NGRR" # Default company abbreviation for letter numbering
full_company_name = "" # New: Default full company name for letter footer
default_save_path = "" # Default path to save generated letters
letterhead_template_path = "" # Path to the Word document template for letterhead

def set_default_settings():
    """Sets default application settings and saves them."""
    global company_name, full_company_name, default_save_path, letterhead_template_path
    
    # Set sensible defaults
    company_name = "NGRR"
    full_company_name = "نام کامل شرکت شما" # Default for the new field
    
    # Attempt to set default_save_path to a 'Documents' folder or current working directory
    try:
        # For Windows
        default_save_path = os.path.join(os.path.expanduser("~"), "Documents", "GeneratedLetters")
        if not os.path.exists(default_save_path):
            os.makedirs(default_save_path)
    except Exception:
        # Fallback for other OS or if Documents folder isn't accessible
        default_save_path = os.getcwd()
    
    letterhead_template_path = "" # User will need to set this via UI

    save_settings() # Immediately save defaults to file
    
def load_settings():
    """Loads application settings from 'settings.txt'.
    If the file doesn't exist or is incomplete, it sets default settings.
    """
    global company_name, full_company_name, default_save_path, letterhead_template_path

    # Ensure default paths are handled initially if the file doesn't exist
    # (This block is primarily for initial setup if default_save_path is not yet in settings.txt)
    # This ensures that a default save path is available even if settings.txt is new or empty.
    initial_default_save_path = os.path.join(os.path.expanduser("~"), "Documents", "GeneratedLetters")
    if not os.path.exists(initial_default_save_path):
        os.makedirs(initial_default_save_path, exist_ok=True)


    if os.path.exists("settings.txt"):
        try:
            with open("settings.txt", "r", encoding="utf-8") as file:
                settings = file.readlines()
                # Ensure we have at least 4 lines for all settings
                # Index 0: company_name
                # Index 1: full_company_name
                # Index 2: default_save_path
                # Index 3: letterhead_template_path
                if len(settings) >= 4:
                    company_name = settings[0].strip()
                    full_company_name = settings[1].strip() 
                    default_save_path = settings[2].strip()
                    letterhead_template_path = settings[3].strip()
                else:
                    # If file exists but is incomplete (e.g., old version), load defaults and save
                    set_default_settings() 
        except Exception as e:
            messagebox.showerror("خطا در بارگذاری", f"خطا در بارگذاری تنظیمات: {e}\nتنظیمات پیش‌فرض اعمال شد.")
            set_default_settings() # Fallback to defaults on error
    else:
        # If settings.txt does not exist, create it with default settings
        set_default_settings()

def save_settings():
    """Saves current application settings (from global variables) to 'settings.txt'."""
    try:
        with open("settings.txt", "w", encoding="utf-8") as file:
            file.write(f"{company_name}\n")
            file.write(f"{full_company_name}\n") 
            file.write(f"{default_save_path}\n")
            file.write(f"{letterhead_template_path}\n")
    except Exception as e:
        messagebox.showerror("خطا در ذخیره", f"خطا در ذخیره تنظیمات: {e}")

# Initial load of settings when the module is imported
# This ensures global variables are populated as soon as settings_manager is used.
load_settings()
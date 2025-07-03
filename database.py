import sqlite3
import os
from tkinter import messagebox # Importing for error messages in DB operations

DATABASE_NAME = "crm.db" # مطمئن شوید این مسیر به فایل دیتابیس صحیح اشاره دارد (ممکن است crm_letters.db باشد)

def get_db_connection():
    """Establishes a connection to the SQLite database."""
    conn = sqlite3.connect(DATABASE_NAME)
    conn.row_factory = sqlite3.Row  # Allows accessing columns by name
    return conn

def create_tables():
    """Creates necessary tables in the database if they don't exist,
    and adds missing columns to existing tables."""
    conn = get_db_connection()
    cursor = conn.cursor()

    # Create Organizations table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Organizations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            industry TEXT,
            phone TEXT,
            email TEXT,
            address TEXT,
            description TEXT
        )
    """)

    # Create Contacts table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            organization_id INTEGER,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            title TEXT,
            phone TEXT,
            email TEXT,
            notes TEXT,
            FOREIGN KEY (organization_id) REFERENCES Organizations(id) ON DELETE SET NULL
        )
    """)
    # Add an index to contacts for faster lookup by organization_id
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_contacts_organization_id ON Contacts(organization_id)")


    # Create Letters table
    # This will create the table if it doesn't exist with all specified columns
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Letters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            letter_code TEXT NOT NULL UNIQUE,
            letter_code_persian TEXT NOT NULL UNIQUE,
            letter_type_abbr TEXT,
            letter_type_persian TEXT,
            date_gregorian TEXT,
            date_shamsi_persian TEXT,
            subject TEXT,
            organization_id INTEGER,
            contact_id INTEGER,
            body TEXT,
            file_path TEXT,
            FOREIGN KEY (organization_id) REFERENCES Organizations(id) ON DELETE SET NULL,
            FOREIGN KEY (contact_id) REFERENCES Contacts(id) ON DELETE SET NULL
        )
    """)

    # --- Schema Migration: Add missing columns to existing tables ---
    # Check if 'body' column exists in 'Letters' table, add if not
    cursor.execute("PRAGMA table_info(Letters)")
    columns = [col[1] for col in cursor.fetchall()] # Get column names
    
    if 'body' not in columns:
        try:
            cursor.execute("ALTER TABLE Letters ADD COLUMN body TEXT")
            conn.commit()
            print("DEBUG: Added 'body' column to Letters table.")
        except sqlite3.Error as e:
            print(f"DEBUG: Error adding 'body' column: {e}")

    # Check if 'file_path' column exists in 'Letters' table, add if not
    if 'file_path' not in columns: # Re-fetch or ensure list is fresh if multiple alters are possible
        try:
            cursor.execute("ALTER TABLE Letters ADD COLUMN file_path TEXT")
            conn.commit()
            print("DEBUG: Added 'file_path' column to Letters table.")
        except sqlite3.Error as e:
            print(f"DEBUG: Error adding 'file_path' column: {e}")

    conn.commit() # Commit any pending changes
    conn.close()

def insert_organization(name, industry, phone, email, address, description):
    """Inserts a new organization into the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO Organizations (name, industry, phone, email, address, description)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (name, industry, phone, email, address, description))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        messagebox.showerror("خطا", "سازمانی با این نام از قبل وجود دارد.", title="خطای ورودی")
        return False
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در افزودن سازمان: {e}", title="خطا")
        return False
    finally:
        conn.close()

def update_organization(org_id, name, industry, phone, email, address, description):
    """Updates an existing organization in the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            UPDATE Organizations SET name=?, industry=?, phone=?, email=?, address=?, description=?
            WHERE id=?
        """, (name, industry, phone, email, address, description, org_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        messagebox.showerror("خطا", "سازمانی با این نام از قبل وجود دارد.", title="خطای ورودی")
        return False
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در ویرایش سازمان: {e}", title="خطا")
        return False
    finally:
        conn.close()

def delete_organization(org_id):
    """Deletes an organization from the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Organizations WHERE id=?", (org_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در حذف سازمان: {e}", title="خطا")
        return False
    finally:
        conn.close()

def get_organizations_from_db(search_term=""):
    """Retrieves organizations from the database, optionally filtered by search term."""
    conn = get_db_connection()
    cursor = conn.cursor()
    query = "SELECT * FROM Organizations"
    params = []
    if search_term:
        query += " WHERE name LIKE ? OR industry LIKE ? OR phone LIKE ? OR email LIKE ?"
        search_pattern = '%' + search_term + '%'
        params = [search_pattern, search_pattern, search_pattern, search_pattern]
    cursor.execute(query, tuple(params))
    organizations = cursor.fetchall()
    conn.close()
    return [dict(org) for org in organizations] # Convert rows to dictionaries

def get_organization_by_id(org_id):
    """Retrieves a single organization record by its ID."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Organizations WHERE id=?", (org_id,))
    org = cursor.fetchone()
    conn.close()
    return dict(org) if org else None

def insert_contact(organization_id, first_name, last_name, title, phone, email, notes):
    """Inserts a new contact into the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO Contacts (organization_id, first_name, last_name, title, phone, email, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (organization_id, first_name, last_name, title, phone, email, notes))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در افزودن مخاطب: {e}", title="خطا")
        return False
    finally:
        conn.close()

def update_contact(contact_id, organization_id, first_name, last_name, title, phone, email, notes):
    """Updates an existing contact in the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            UPDATE Contacts SET organization_id=?, first_name=?, last_name=?, title=?, phone=?, email=?, notes=?
            WHERE id=?
        """, (organization_id, first_name, last_name, title, phone, email, notes, contact_id))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در ویرایش مخاطب: {e}", title="خطا")
        return False
    finally:
        conn.close()

def delete_contact(contact_id):
    """Deletes a contact from the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Contacts WHERE id=?", (contact_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در حذف مخاطب: {e}", title="خطا")
        return False
    finally:
        conn.close()

def get_contacts_from_db(organization_id=None, search_term=""):
    """Retrieves contacts from the database, optionally filtered by organization_id or search term."""
    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT C.*, O.name AS organization_name
        FROM Contacts C
        LEFT JOIN Organizations O ON C.organization_id = O.id
    """
    conditions = []
    params = []

    if organization_id:
        conditions.append("C.organization_id = ?")
        params.append(organization_id)
    
    if search_term:
        conditions.append("(C.first_name LIKE ? OR C.last_name LIKE ? OR C.title LIKE ? OR O.name LIKE ?)")
        search_pattern = '%' + search_term + '%'
        params.extend([search_pattern, search_pattern, search_pattern, search_pattern])

    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    cursor.execute(query, tuple(params))
    contacts = cursor.fetchall()
    conn.close()
    return [dict(contact) for contact in contacts] # Convert rows to dictionaries

def get_contact_by_id(contact_id):
    """Retrieves a single contact record by its ID."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT C.*, O.name AS organization_name
        FROM Contacts C
        LEFT JOIN Organizations O ON C.organization_id = O.id
        WHERE C.id=?
    """, (contact_id,))
    contact = cursor.fetchone()
    conn.close()
    return dict(contact) if contact else None

def insert_letter(letter_code, letter_code_persian, letter_type_abbr, letter_type_persian,
                  date_gregorian, date_shamsi_persian, subject, 
                  organization_id, contact_id, body, file_path):
    """Inserts a new letter into the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO Letters (letter_code, letter_code_persian, letter_type_abbr, letter_type_persian,
                                 date_gregorian, date_shamsi_persian, subject, 
                                 organization_id, contact_id, body, file_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (letter_code, letter_code_persian, letter_type_abbr, letter_type_persian,
              date_gregorian, date_shamsi_persian, subject, 
              organization_id, contact_id, body, file_path))
        conn.commit()
        return True
    except sqlite3.IntegrityError as e:
        if "UNIQUE constraint failed: Letters.letter_code" in str(e):
            messagebox.showerror(title="خطای ورودی", message="شماره نامه تکراری است. لطفاً دوباره تلاش کنید.")
        else:
            messagebox.showerror(title="خطا", message=f"خطا در افزودن نامه: {e}")
        return False
    except Exception as e:
        messagebox.showerror(title="خطا", message=f"خطا در افزودن نامه: {e}")
        return False
    finally:
        conn.close()

def get_letters_from_db(search_term=""):
    """Retrieves letters from the database, optionally filtered by search term."""
    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT 
            L.*, 
            O.name AS organization_name, 
            C.first_name, 
            C.last_name
        FROM Letters L
        LEFT JOIN Organizations O ON L.organization_id = O.id
        LEFT JOIN Contacts C ON L.contact_id = C.id
    """
    conditions = []
    params = []

    if search_term:
        conditions.append("(L.letter_code_persian LIKE ? OR L.subject LIKE ? OR O.name LIKE ? OR C.first_name LIKE ? OR C.last_name LIKE ? OR L.letter_type_persian LIKE ?)")
        search_pattern = '%' + search_term + '%'
        params.extend([search_pattern, search_pattern, search_pattern, search_pattern, search_pattern, search_pattern])

    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    query += " ORDER BY L.date_shamsi_persian DESC, L.id DESC"

    cursor.execute(query, tuple(params))
    letters = cursor.fetchall()
    conn.close()
    return [dict(letter) for letter in letters]

def get_letter_by_code(letter_code_display):
    """Retrieves a single letter record by its letter code (Persian)."""
    conn = get_db_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    query = """
        SELECT 
            L.*, 
            O.name AS organization_name, 
            C.first_name, 
            C.last_name,
            C.title AS contact_title
        FROM Letters L
        LEFT JOIN Organizations O ON L.organization_id = O.id
        LEFT JOIN Contacts C ON L.contact_id = C.id
        WHERE L.letter_code_persian = ?
    """
    try:
        cursor.execute(query, (letter_code_display,))
        row = cursor.fetchone()
        if row:
            return dict(row) 
        else:
            return None
    except Exception as e:
        print(f"Database error in get_letter_by_code: {e}")
        return None
    finally:
        conn.close()
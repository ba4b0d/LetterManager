import sqlite3
import os
import hashlib 
from tkinter import messagebox 
from datetime import datetime 

DATABASE_NAME = "crm.db" 

def get_db_connection():
    """Establishes a connection to the SQLite database."""
    conn = sqlite3.connect(DATABASE_NAME)
    conn.row_factory = sqlite3.Row  
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
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_contacts_organization_id ON Contacts (organization_id);")

    # Create Letters table - Ensure this matches the columns being inserted
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Letters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            letter_code_prefix TEXT NOT NULL,
            letter_code_number INTEGER NOT NULL,
            letter_code_persian TEXT NOT NULL UNIQUE, 
            type TEXT NOT NULL,
            date_shamsi_persian TEXT NOT NULL,
            subject TEXT NOT NULL,
            body TEXT NOT NULL,
            organization_id INTEGER,
            contact_id INTEGER,
            file_path TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            user_id INTEGER, 
            FOREIGN KEY (organization_id) REFERENCES Organizations(id) ON DELETE SET NULL,
            FOREIGN KEY (contact_id) REFERENCES Contacts(id) ON DELETE SET NULL,
            FOREIGN KEY (user_id) REFERENCES Users(id) ON DELETE SET NULL 
        )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_letters_user_id ON Letters (user_id);")


    # Create Users table 
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL, 
            role TEXT NOT NULL DEFAULT 'user' 
        )
    """)
    conn.commit()
    conn.close()

# --- User Management Functions ---
def _hash_password(password):
    """Hashes a password using SHA256. This is an internal helper function."""
    return hashlib.sha256(password.encode()).hexdigest()

def get_user_by_username(username):
    """Retrieves a user by username."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, password_hash, role FROM Users WHERE username = ?", (username,))
    user = cursor.fetchone()
    conn.close()
    return dict(user) if user else None

def get_user_by_id(user_id):
    """Retrieves a user by ID."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role, password_hash FROM Users WHERE id = ?", (user_id,)) 
    user = cursor.fetchone()
    conn.close()
    return dict(user) if user else None

def add_user(username, password, role):
    """Adds a new user to the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    hashed_password = _hash_password(password)
    try:
        cursor.execute("INSERT INTO Users (username, password_hash, role) VALUES (?, ?, ?)",
                       (username, hashed_password, role))
        conn.commit()
        return True
    except sqlite3.IntegrityError: 
        messagebox.showerror("خطا", "نام کاربری از قبل موجود است.")
        return False
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در افزودن کاربر: {e}")
        return False
    finally:
        conn.close()

def verify_password(username, password):
    """Verifies a user's password."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT password_hash FROM Users WHERE username = ?", (username,))
    user_data = cursor.fetchone()
    conn.close()
    if user_data:
        return user_data['password_hash'] == _hash_password(password)
    return False

def get_all_users():
    """Retrieves all users from the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role FROM Users ORDER BY username")
    users = cursor.fetchall()
    conn.close()
    return [dict(user) for user in users]

def update_user(user_id, username=None, role=None):
    """Updates an existing user's username or role."""
    conn = get_db_connection()
    cursor = conn.cursor()
    updates = []
    params = []
    if username:
        updates.append("username = ?")
        params.append(username)
    if role:
        updates.append("role = ?")
        params.append(role)
    
    if not updates:
        conn.close()
        return False 

    query = f"UPDATE Users SET {', '.join(updates)} WHERE id = ?"
    params.append(user_id)
    try:
        cursor.execute(query, tuple(params))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        messagebox.showerror("خطا", "نام کاربری از قبل موجود است.")
        return False
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در ویرایش کاربر: {e}")
        return False
    finally:
        conn.close()

def update_user_password(user_id, new_password):
    """Updates a user's password."""
    conn = get_db_connection()
    cursor = conn.cursor()
    hashed_password = _hash_password(new_password)
    try:
        cursor.execute("UPDATE Users SET password_hash = ? WHERE id = ?", (hashed_password, user_id))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در تغییر رمز عبور: {e}")
        return False
    finally:
        conn.close()

def delete_user(user_id):
    """Deletes a user from the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Users WHERE id = ?", (user_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در حذف کاربر: {e}")
        return False
    finally:
        conn.close()

# --- Letter Management Functions ---
# Modified insert_letter to accept individual parameters
def insert_letter(letter_code_prefix, letter_code_number, letter_code_persian, type, date_gregorian, date_shamsi_persian, subject, organization_id, contact_id, body, file_path, user_id):
    """Inserts a new letter record into the database."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO Letters (letter_code_prefix, letter_code_number, letter_code_persian, type, date_shamsi_persian, subject, body, organization_id, contact_id, file_path, user_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            letter_code_prefix,
            letter_code_number,
            letter_code_persian,
            type,
            date_shamsi_persian,
            subject,
            body,
            organization_id,
            contact_id,
            file_path,
            user_id,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S") # Add created_at
        ))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا در ذخیره نامه", f"خطا در ذخیره نامه در دیتابیس: {e}")
        return False
    finally:
        conn.close()

def get_letters_from_db(search_term=""):
    """Retrieves letter records from the database, optionally filtered by search term."""
    conn = get_db_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    query = """
        SELECT 
            L.*, 
            L.type AS letter_type_raw, -- Explicitly select and alias the 'type' column
            O.name AS organization_name, 
            C.first_name, 
            C.last_name,
            C.title AS contact_title,
            U.username AS created_by_username
        FROM Letters L
        LEFT JOIN Organizations O ON L.organization_id = O.id
        LEFT JOIN Contacts C ON L.contact_id = C.id
        LEFT JOIN Users U ON L.user_id = U.id
    """
    
    params = []
    conditions = []

    if search_term:
        search_pattern = f"%{search_term}%"
        conditions.append("(L.letter_code_persian LIKE ? OR L.type LIKE ? OR L.subject LIKE ? OR O.name LIKE ? OR C.first_name LIKE ? OR C.last_name LIKE ? OR U.username LIKE ?)")
        params.extend([search_pattern, search_pattern, search_pattern, search_pattern, search_pattern, search_pattern, search_pattern])

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
            L.type AS letter_type_raw, -- Explicitly select and alias the 'type' column
            O.name AS organization_name, 
            C.first_name, 
            C.last_name,
            C.title AS contact_title,
            U.username AS created_by_username 
        FROM Letters L
        LEFT JOIN Organizations O ON L.organization_id = O.id
        LEFT JOIN Contacts C ON L.contact_id = C.id
        LEFT JOIN Users U ON L.user_id = U.id 
        WHERE L.letter_code_persian = ?
    """
    try:
        cursor.execute(query, (letter_code_display,))
        letter = cursor.fetchone()
        return dict(letter) if letter else None
    except Exception as e:
        messagebox.showerror("خطا در بازیابی نامه", f"خطا در بازیابی نامه با کد {letter_code_display}: {e}")
        return None
    finally:
        conn.close()

# --- Organization and Contact functions ---
def get_organizations_from_db(search_term=""): 
    conn = get_db_connection()
    cursor = conn.cursor()
    query = "SELECT id, name, industry, phone, email, address, description FROM Organizations"
    params = []
    if search_term:
        query += " WHERE name LIKE ?"
        params.append(f"%{search_term}%")
    query += " ORDER BY name"
    cursor.execute(query, params)
    orgs = cursor.fetchall()
    conn.close()
    return [dict(org) for org in orgs]

def get_organization_by_id(org_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, industry, phone, email, address, description FROM Organizations WHERE id = ?", (org_id,))
    org = cursor.fetchone()
    conn.close()
    return dict(org) if org else None

def insert_organization(name, industry, phone, email, address, description):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO Organizations (name, industry, phone, email, address, description) VALUES (?, ?, ?, ?, ?, ?)",
                       (name, industry, phone, email, address, description))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        messagebox.showerror("خطا", "سازمانی با این نام از قبل موجود است.")
        return False
    finally:
        conn.close()

def update_organization(org_id, name, industry, phone, email, address, description):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE Organizations SET name=?, industry=?, phone=?, email=?, address=?, description=? WHERE id=?",
                       (name, industry, phone, email, address, description, org_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        messagebox.showerror("خطا", "سازمانی با این نام از قبل موجود است.")
        return False
    finally:
        conn.close()

def delete_organization(org_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Organizations WHERE id=?", (org_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در حذف سازمان: {e}")
        return False
    finally:
        conn.close()

def get_contacts_from_db(organization_id=None, search_term=""): 
    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT C.id, C.organization_id, C.first_name, C.last_name, C.title, C.phone, C.email, C.notes, O.name AS organization_name
        FROM Contacts C
        LEFT JOIN Organizations O ON C.organization_id = O.id
    """
    params = []
    conditions = []

    if organization_id is not None:
        conditions.append("C.organization_id = ?")
        params.append(organization_id)
    
    if search_term:
        search_pattern = f"%{search_term}%"
        conditions.append("(C.first_name LIKE ? OR C.last_name LIKE ? OR C.title LIKE ? OR O.name LIKE ?)")
        params.extend([search_pattern, search_pattern, search_pattern, search_pattern])

    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    query += " ORDER BY C.last_name, C.first_name"
    
    cursor.execute(query, params)
    contacts = cursor.fetchall()
    conn.close()
    return [dict(contact) for contact in contacts]

def get_contact_by_id(contact_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, organization_id, first_name, last_name, title, phone, email, notes FROM Contacts WHERE id = ?", (contact_id,))
    contact = cursor.fetchone()
    conn.close()
    return dict(contact) if contact else None

def insert_contact(organization_id, first_name, last_name, title, phone, email, notes):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO Contacts (organization_id, first_name, last_name, title, phone, email, notes) VALUES (?, ?, ?, ?, ?, ?, ?)",
                       (organization_id, first_name, last_name, title, phone, email, notes))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در افزودن مخاطب: {e}")
        return False
    finally:
        conn.close()

def update_contact(contact_id, organization_id, first_name, last_name, title, phone, email, notes):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE Contacts SET organization_id=?, first_name=?, last_name=?, title=?, phone=?, email=?, notes=? WHERE id=?",
                       (organization_id, first_name, last_name, title, phone, email, notes, contact_id))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در ویرایش مخاطب: {e}")
        return False
    finally:
        conn.close()

def delete_contact(contact_id):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Contacts WHERE id=?", (contact_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا", f"خطا در حذف مخاطب: {e}")
        return False
    finally:
        conn.close()

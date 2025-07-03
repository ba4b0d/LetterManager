import sqlite3
import os
from tkinter import messagebox # Importing for error messages in DB operations

DATABASE_NAME = "crm.db"

def get_db_connection():
    """Establishes a connection to the SQLite database."""
    conn = sqlite3.connect(DATABASE_NAME)
    conn.row_factory = sqlite3.Row  # Allows accessing columns by name
    return conn

def create_tables():
    """Creates necessary tables in the database if they don't exist."""
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
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_contacts_organization_id ON Contacts (organization_id)")


    # Create Letters table
    # CHANGED: 'letter_type' to 'letter_type_persian' and added 'letter_type_abbr', 'body_text'
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS Letters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            letter_code TEXT NOT NULL UNIQUE,
            letter_code_persian TEXT NOT NULL UNIQUE,
            letter_type_abbr TEXT NOT NULL,
            letter_type_persian TEXT NOT NULL,
            date_gregorian TEXT NOT NULL,
            date_shamsi_persian TEXT NOT NULL,
            subject TEXT NOT NULL,
            organization_id INTEGER,
            contact_id INTEGER,
            body_text TEXT, 
            file_path TEXT NOT NULL,
            FOREIGN KEY (organization_id) REFERENCES Organizations(id) ON DELETE SET NULL,
            FOREIGN KEY (contact_id) REFERENCES Contacts(id) ON DELETE SET NULL
        )
    """)
    conn.commit()
    conn.close()

def insert_organization(name, industry, phone, email, address, description):
    """Inserts a new organization record into the Organizations table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO Organizations (name, industry, phone, email, address, description)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (name, industry, phone, email, address, description))
        conn.commit()
        return True
    except sqlite3.IntegrityError as e:
        messagebox.showerror("خطا در پایگاه داده", f"سازمانی با این نام از قبل موجود است: {name}\n\nپیام خطا: {e}")
        return False
    finally:
        conn.close()

def get_organizations(search_term=""):
    """Retrieves organization records from the Organizations table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    query = "SELECT * FROM Organizations"
    params = []
    if search_term:
        query += " WHERE name LIKE ? OR industry LIKE ?"
        params = [f"%{search_term}%", f"%{search_term}%"]
    cursor.execute(query, tuple(params))
    organizations = cursor.fetchall()
    conn.close()
    return [dict(org) for org in organizations]

def update_organization(org_id, name, industry, phone, email, address, description):
    """Updates an existing organization record in the Organizations table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            UPDATE Organizations
            SET name = ?, industry = ?, phone = ?, email = ?, address = ?, description = ?
            WHERE id = ?
        """, (name, industry, phone, email, address, description, org_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError as e:
        messagebox.showerror("خطا در پایگاه داده", f"سازمانی با این نام از قبل موجود است: {name}\n\nپیام خطا: {e}")
        return False
    finally:
        conn.close()

def delete_organization(org_id):
    """Deletes an organization record from the Organizations table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Organizations WHERE id = ?", (org_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا در پایگاه داده", f"خطا در حذف سازمان: {e}")
        return False
    finally:
        conn.close()

def insert_contact(organization_id, first_name, last_name, title, phone, email, notes):
    """Inserts a new contact record into the Contacts table."""
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
        messagebox.showerror("خطا در پایگاه داده", f"خطا در ذخیره مخاطب: {e}")
        return False
    finally:
        conn.close()

def get_contacts(organization_id=None, search_term=""):
    """Retrieves contact records from the Contacts table, optionally filtered by organization."""
    conn = get_db_connection()
    cursor = conn.cursor()
    query = "SELECT * FROM Contacts"
    params = []
    conditions = []

    if organization_id is not None:
        conditions.append("organization_id = ?")
        params.append(organization_id)
    
    if search_term:
        conditions.append("(first_name LIKE ? OR last_name LIKE ? OR title LIKE ? OR email LIKE ?)")
        search_pattern = f"%{search_term}%"
        params.extend([search_pattern, search_pattern, search_pattern, search_pattern])

    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    cursor.execute(query, tuple(params))
    contacts = cursor.fetchall()
    conn.close()
    return [dict(contact) for contact in contacts]

def get_contact_by_id(contact_id):
    """Retrieves a single contact record by its ID."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Contacts WHERE id = ?", (contact_id,))
    contact = cursor.fetchone()
    conn.close()
    return dict(contact) if contact else None

def update_contact(contact_id, organization_id, first_name, last_name, title, phone, email, notes):
    """Updates an existing contact record in the Contacts table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            UPDATE Contacts
            SET organization_id = ?, first_name = ?, last_name = ?, title = ?, phone = ?, email = ?, notes = ?
            WHERE id = ?
        """, (organization_id, first_name, last_name, title, phone, email, notes, contact_id))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا در پایگاه داده", f"خطا در بروزرسانی مخاطب: {e}")
        return False
    finally:
        conn.close()

def delete_contact(contact_id):
    """Deletes a contact record from the Contacts table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Contacts WHERE id = ?", (contact_id,))
        conn.commit()
        return True
    except Exception as e:
        messagebox.showerror("خطا در پایگاه داده", f"خطا در حذف مخاطب: {e}")
        return False
    finally:
        conn.close()

# CHANGED: Added 'letter_type_abbr' parameter and corrected parameter order in INSERT statement
def insert_letter(letter_code, letter_code_persian, letter_type_abbr, letter_type_persian,
                  date_gregorian, date_shamsi_persian, subject, organization_id, contact_id, body_text, file_path):
    """Inserts a new letter record into the Letters table."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT INTO Letters (
                letter_code, letter_code_persian, letter_type_abbr, letter_type_persian,
                date_gregorian, date_shamsi_persian, subject, organization_id, contact_id, body_text, file_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            letter_code, letter_code_persian, letter_type_abbr, letter_type_persian,
            date_gregorian, date_shamsi_persian, subject, organization_id, contact_id, body_text, file_path
        ))
        conn.commit()
        return True
    except sqlite3.IntegrityError as e:
        messagebox.showerror("خطا در پایگاه داده", f"خطا در ذخیره نامه: {e}")
        return False
    finally:
        conn.close()

def get_letters_from_db(organization_id=None, contact_id=None, search_term=""):
    """Retrieves letter records from the Letters table, optionally filtered and searched."""
    conn = get_db_connection()
    conn.row_factory = sqlite3.Row # Ensure row_factory is set for this connection too
    cursor = conn.cursor()
    
    # CHANGED: Select all columns and join with Organizations and Contacts for display
    query = """
        SELECT 
            L.*, 
            O.name AS organization_name, 
            C.first_name, 
            C.last_name,
            C.title AS contact_title -- Added contact title
        FROM Letters L
        LEFT JOIN Organizations O ON L.organization_id = O.id
        LEFT JOIN Contacts C ON L.contact_id = C.id
    """
    params = []
    conditions = []

    if organization_id is not None:
        conditions.append("L.organization_id = ?")
        params.append(organization_id)
    
    if contact_id is not None:
        conditions.append("L.contact_id = ?")
        params.append(contact_id)

    if search_term:
        conditions.append("(L.letter_code_persian LIKE ? OR L.subject LIKE ? OR O.name LIKE ? OR C.first_name LIKE ? OR C.last_name LIKE ? OR L.letter_type_persian LIKE ?)") # Added letter_type_persian to search
        search_pattern = '%' + search_term + '%'
        params.extend([search_pattern, search_pattern, search_pattern, search_pattern, search_pattern, search_pattern]) # Extend for the new search term

    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    
    query += " ORDER BY L.date_shamsi_persian DESC, L.id DESC" # Order by date and then ID (for unique ordering)

    cursor.execute(query, tuple(params))
    letters = cursor.fetchall()
    conn.close()
    return [dict(letter) for letter in letters]

def get_letter_by_code(letter_code):
    """Retrieves a single letter record by its letter code (English)."""
    conn = get_db_connection()
    conn.row_factory = sqlite3.Row # Ensure row_factory is set
    cursor = conn.cursor()
    # CHANGED: Select all columns and join with Organizations and Contacts for display
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
        WHERE L.letter_code = ?
    """
    cursor.execute(query, (letter_code,))
    letter = cursor.fetchone()
    conn.close()
    return dict(letter) if letter else None
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import pandas as pd
import os
import json
import sqlite3
from collections import defaultdict
import re  # For parsing Extra column and checking markers
import traceback  # For detailed error logging
import time  # For slight delay on busy cursor
import shutil  # For file operations
import openpyxl  # For Excel file operations
# Import report templates
from report_templates import ReportManager

# Helper functions for cursor state
def set_busy_cursor(widget):
    """Set the cursor to a busy/wait cursor for the given widget."""
    widget.config(cursor="watch")
    widget.update_idletasks()
    
def set_default_cursor(widget):
    """Set the cursor back to the default for the given widget."""
    widget.config(cursor="")
    widget.update_idletasks()
try:
    from tkcalendar import DateEntry  # Import DateEntry
except ImportError:
    messagebox.showerror("Missing Library", "The 'tkcalendar' library is required.\nPlease install it using:\npip install tkcalendar")
    exit()

# --- Configuration ---
SCOPES = ['https://spreadsheets.google.com/feeds',
          'https://www.googleapis.com/auth/drive']
CREDENTIALS_FILE = "google_credentials.json"
SPREADSHEET_ID = "1Qgsb73esqN3iYroYszW5m1_I8tqZ-wXy5ZOsr-saBsU"  # Default ID
CONFIG_FILE = "pos_config.json"  # File to store configuration
BACKUP_FOLDER = 'sales_backup'  # Folder to store local Excel backups
DAILY_SUMMARY_SEPARATOR = "--- Daily Summary ---"  # Marker for summary block start
PURCHASE_MARKER = "PURCHASE - "  # Prefix for supplier name in Customer column for purchase rows
ICON_PATH = 'icons'  # Folder for icon images
DB_FILE = "pos_data.db"  # SQLite database file

# --- Default Configuration ---
DEFAULT_CONFIG = {
    "app_mode": "offline",  # Start in offline mode by default
    "spreadsheet_id": SPREADSHEET_ID,
    "backup_folder": BACKUP_FOLDER,
    "auto_backup": True,
    "backup_interval_days": 1,
    "last_backup": "",
    "theme": "light",
    "language": "en",
    "company_name": "Your Company",
    "company_address": "",
    "company_phone": "",
    "company_email": ""
}

# --- Utility Functions ---
def safe_float(value, default=0.0):
    """Safely converts a value to float, returning default on error."""
    try:
        if value is None or str(value).strip() == "":
            return default
        # Remove currency symbols, commas etc., keeping only digits, dot, and minus
        cleaned_value = re.sub(r'[^\d.-]+', '', str(value))
        if cleaned_value and cleaned_value != '-':  # Ensure not just a minus sign
            return float(cleaned_value)
        else:
            return default
    except (ValueError, TypeError):
        return default

def safe_int(value, default=0):
    """Safely converts a value to int (via float), returning default on error."""
    try:
        if value is None or str(value).strip() == "":
            return default
        # Allow conversion from "5.0" to 5 by first converting to float
        float_val = safe_float(value, default=0.0)  # Specify default explicitly as a float
        if float_val is not None:
            return int(float_val)
        else:
            return default  # If safe_float failed, return default int
    except (ValueError, TypeError):
        return default

def set_busy_cursor(widget):
    """Sets the cursor to 'watch' for the given widget."""
    if widget and widget.winfo_exists():
        try:
            widget.config(cursor="watch")
            widget.update_idletasks()
        except tk.TclError:
            print("Warning: Could not set busy cursor (widget might be destroyed).")

def set_default_cursor(widget):
    """Sets the cursor back to default for the given widget."""
    if widget and widget.winfo_exists():
        try:
            widget.config(cursor="")
            widget.update_idletasks()
        except tk.TclError:
            print("Warning: Could not set default cursor (widget might be destroyed).")

def ensure_directory_exists(dir_path):
    """Creates a directory if it doesn't exist."""
    if not os.path.exists(dir_path):
        try:
            os.makedirs(dir_path)
            print(f"Created directory: {dir_path}")
        except OSError as e:
            print(f"Error creating directory '{dir_path}': {e}")
            messagebox.showwarning("Startup Warning", f"Could not create directory:\n{dir_path}\n\nPlease create it manually.\nError: {e}")
            return False
    return True

class ConfigManager:
    """Manages application configuration."""
    
    def __init__(self, config_file=CONFIG_FILE):
        self.config_file = config_file
        self.config = self.load_config()
        
    def load_config(self):
        """Load configuration from file or create default."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            else:
                # Create a new config file with defaults
                self.save_config(DEFAULT_CONFIG)
                return DEFAULT_CONFIG.copy()
        except Exception as e:
            print(f"Error loading configuration: {e}")
            return DEFAULT_CONFIG.copy()
            
    def save_config(self, config=None):
        """Save configuration to file."""
        if config is not None:
            self.config = config
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=4)
            return True
        except Exception as e:
            print(f"Error saving configuration: {e}")
            return False
            
    def get(self, key, default=None):
        """Get a configuration value."""
        return self.config.get(key, default)
        
    def set(self, key, value):
        """Set a configuration value and save."""
        self.config[key] = value
        return self.save_config()

class DatabaseManager:
    """Manages the local SQLite database."""
    
    def __init__(self, db_file=DB_FILE):
        self.db_file = db_file
        self.connection = None
        self.create_tables()
        
    def connect(self):
        """Connect to the database."""
        if self.connection is None:
            self.connection = sqlite3.connect(self.db_file)
            self.connection.row_factory = sqlite3.Row
        return self.connection
        
    def close(self):
        """Close the database connection."""
        if self.connection:
            self.connection.close()
            self.connection = None
            
    def create_tables(self):
        """Create database tables if they don't exist."""
        conn = self.connect()
        cursor = conn.cursor()
        
        # Create Products table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            price REAL NOT NULL,
            current_stock REAL DEFAULT 0,
            active INTEGER DEFAULT 1
        )
        ''')
        
        # Create Customers table with Supplier flag
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            contact TEXT,
            credit_balance REAL DEFAULT 0,
            active INTEGER DEFAULT 1,
            is_supplier INTEGER DEFAULT 0
        )
        ''')
        
        # Create Transactions table to handle both sales and purchases
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            customer_id INTEGER NOT NULL,
            transaction_type TEXT NOT NULL,  -- 'sale' or 'purchase'
            total_amount REAL NOT NULL,
            cash_received REAL NOT NULL,
            previous_credit REAL NOT NULL,
            updated_credit REAL NOT NULL,
            notes TEXT,
            synced INTEGER DEFAULT 0,
            FOREIGN KEY (customer_id) REFERENCES customers (id)
        )
        ''')
        
        # Create Transaction Items table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS transaction_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            transaction_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            quantity REAL NOT NULL,
            unit_price REAL NOT NULL,
            subtotal REAL NOT NULL,
            FOREIGN KEY (transaction_id) REFERENCES transactions (id),
            FOREIGN KEY (product_id) REFERENCES products (id)
        )
        ''')
        
        # Maintain compatibility with existing code that uses 'sales' and 'sale_items' tables
        cursor.execute('''
        CREATE VIEW IF NOT EXISTS sales AS
        SELECT * FROM transactions
        ''')
        
        cursor.execute('''
        CREATE VIEW IF NOT EXISTS sale_items AS
        SELECT 
            id,
            transaction_id as sale_id,
            product_id,
            quantity,
            unit_price as price
        FROM transaction_items
        ''')
        
        # Create CustomerPrices table (for custom pricing)
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS customer_prices (
            customer_id INTEGER NOT NULL,
            product_id INTEGER NOT NULL,
            price REAL NOT NULL,
            PRIMARY KEY (customer_id, product_id),
            FOREIGN KEY (customer_id) REFERENCES customers(id),
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
        ''')
        
        # Create DailySummary table with inventory data field
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_summary (
            date TEXT PRIMARY KEY,
            total_sales REAL NOT NULL,
            total_purchases REAL NOT NULL,
            cash_received REAL NOT NULL,
            inventory_data TEXT,
            synced INTEGER DEFAULT 0
        )
        ''')
        
        # Create StockMovements table to track daily inventory changes
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS stock_movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            product_id INTEGER NOT NULL,
            quantity_change REAL NOT NULL,  -- positive for purchases, negative for sales
            transaction_id INTEGER,
            movement_type TEXT NOT NULL,  -- 'sale', 'purchase', 'adjustment', etc.
            notes TEXT,
            FOREIGN KEY (product_id) REFERENCES products(id),
            FOREIGN KEY (transaction_id) REFERENCES transactions(id)
        )
        ''')
        
        # Create daily stock tracking table
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS daily_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            product_id INTEGER NOT NULL,
            opening_stock REAL NOT NULL,
            purchases REAL DEFAULT 0,
            sales REAL DEFAULT 0,
            adjustments REAL DEFAULT 0,
            closing_stock REAL NOT NULL,
            FOREIGN KEY (product_id) REFERENCES products(id),
            UNIQUE (date, product_id)
        )
        ''')
        
        conn.commit()
        
    def add_or_update_product(self, name, price):
        """Add a new product or update if exists."""
        conn = self.connect()
        cursor = conn.cursor()
        try:
            cursor.execute('''
            INSERT INTO products (name, price) VALUES (?, ?)
            ON CONFLICT(name) DO UPDATE SET price = ?
            ''', (name, price, price))
            conn.commit()
            return True
        except Exception as e:
            print(f"Error adding/updating product: {e}")
            return False
            
    def add_or_update_customer(self, name, contact="", credit_balance=0, is_supplier=0):
        """Add a new customer or update if exists."""
        conn = self.connect()
        cursor = conn.cursor()
        try:
            cursor.execute('''
            INSERT INTO customers (name, contact, credit_balance, is_supplier) VALUES (?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET contact = ?, credit_balance = ?, is_supplier = ?
            ''', (name, contact, credit_balance, is_supplier, contact, credit_balance, is_supplier))
            conn.commit()
            return True
        except Exception as e:
            print(f"Error adding/updating customer: {e}")
            return False
            
    def get_all_products(self):
        """Get all active products."""
        conn = self.connect()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM products WHERE active = 1")
        return cursor.fetchall()
        
    def get_all_customers(self):
        """Get all active customers (not suppliers)."""
        conn = self.connect()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM customers WHERE active = 1 AND is_supplier = 0")
        return cursor.fetchall()
        
    def get_all_suppliers(self):
        """Get all active suppliers."""
        conn = self.connect()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM customers WHERE active = 1 AND is_supplier = 1")
        return cursor.fetchall()
        
    def get_customer_credit(self, customer_id):
        """Get customer's current credit balance."""
        conn = self.connect()
        cursor = conn.cursor()
        cursor.execute("SELECT credit_balance FROM customers WHERE id = ?", (customer_id,))
        result = cursor.fetchone()
        return result['credit_balance'] if result else 0
        
    def add_sale(self, date, customer_id, items, total_amount, cash_received, previous_credit, updated_credit, notes="", transaction_type="sale"):
        """Add a new transaction (sale or purchase)."""
        conn = self.connect()
        cursor = conn.cursor()
        try:
            # Start a transaction
            conn.execute("BEGIN TRANSACTION")
            
            # Insert transaction record
            cursor.execute('''
            INSERT INTO transactions (date, customer_id, transaction_type, total_amount, cash_received, previous_credit, updated_credit, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (date, customer_id, transaction_type, total_amount, cash_received, previous_credit, updated_credit, notes))
            
            transaction_id = cursor.lastrowid
            
            # Insert transaction items and update stock
            for item in items:
                product_id = item['product_id']
                quantity = item['quantity']
                price = item['price']
                
                # Insert item to the transaction
                cursor.execute('''
                INSERT INTO transaction_items (transaction_id, product_id, quantity, unit_price, subtotal)
                VALUES (?, ?, ?, ?, ?)
                ''', (transaction_id, product_id, quantity, price, quantity * price))
                
                # Get current stock
                cursor.execute("SELECT current_stock FROM products WHERE id = ?", (product_id,))
                result = cursor.fetchone()
                current_stock = 0
                if result and 'current_stock' in result.keys():
                    current_stock = result['current_stock']
                
                # Update stock based on transaction type
                if transaction_type == 'sale':
                    new_stock = current_stock - quantity
                    quantity_change = -quantity  # Negative for sales
                else:  # purchase
                    new_stock = current_stock + quantity
                    quantity_change = quantity  # Positive for purchases
                
                # Check if stock_movements table exists
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='stock_movements'")
                if cursor.fetchone():
                    # Record stock movement
                    cursor.execute('''
                    INSERT INTO stock_movements (date, product_id, quantity_change, transaction_id, movement_type, notes)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ''', (date, product_id, quantity_change, transaction_id, transaction_type, f"{transaction_type} of {quantity} units"))
                
                # Check if daily_stock table exists
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='daily_stock'")
                if cursor.fetchone():
                    # Update daily stock record or create if not exists for this product and date
                    cursor.execute('''
                    SELECT * FROM daily_stock WHERE date = ? AND product_id = ?
                    ''', (date, product_id))
                    daily_stock = cursor.fetchone()
                    
                    if daily_stock:
                        # Update existing record
                        if transaction_type == 'sale':
                            new_sales = daily_stock['sales'] + quantity
                            cursor.execute('''
                            UPDATE daily_stock 
                            SET sales = ?, closing_stock = ? 
                            WHERE date = ? AND product_id = ?
                            ''', (new_sales, new_stock, date, product_id))
                        else:  # purchase
                            new_purchases = daily_stock['purchases'] + quantity
                            cursor.execute('''
                            UPDATE daily_stock 
                            SET purchases = ?, closing_stock = ? 
                            WHERE date = ? AND product_id = ?
                            ''', (new_purchases, new_stock, date, product_id))
                    else:
                        # Create new daily stock record
                        if transaction_type == 'sale':
                            cursor.execute('''
                            INSERT INTO daily_stock 
                            (date, product_id, opening_stock, sales, purchases, adjustments, closing_stock)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                            ''', (date, product_id, current_stock, quantity, 0, 0, new_stock))
                        else:  # purchase
                            cursor.execute('''
                            INSERT INTO daily_stock 
                            (date, product_id, opening_stock, sales, purchases, adjustments, closing_stock)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                            ''', (date, product_id, current_stock, 0, quantity, 0, new_stock))
                
                # Update the product stock if current_stock column exists
                cursor.execute("PRAGMA table_info(products)")
                columns = [info[1] for info in cursor.fetchall()]
                if 'current_stock' in columns:
                    cursor.execute("UPDATE products SET current_stock = ? WHERE id = ?", (new_stock, product_id))
            
            # Update customer's credit balance based on client type
            if transaction_type == 'sale':
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='clients'")
                if cursor.fetchone():
                    cursor.execute('''
                    UPDATE clients SET credit_balance = ? WHERE id = ?
                    ''', (updated_credit, customer_id))
                else:
                    cursor.execute('''
                    UPDATE customers SET credit_balance = ? WHERE id = ?
                    ''', (updated_credit, customer_id))
            else:  # purchase
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='suppliers'")
                if cursor.fetchone():
                    cursor.execute('''
                    UPDATE suppliers SET credit_balance = ? WHERE id = ?
                    ''', (updated_credit, customer_id))
                else:
                    cursor.execute('''
                    UPDATE customers SET credit_balance = ? WHERE id = ?
                    ''', (updated_credit, customer_id))
            
            # Commit the transaction
            conn.commit()
            return transaction_id
        except Exception as e:
            conn.rollback()
            print(f"Error adding transaction: {e}")
            traceback.print_exc()
            return None
            
    def add_purchase(self, date, supplier_id, items, total_amount, cash_paid, previous_credit, updated_credit, notes=""):
        """Add a new purchase transaction from a supplier."""
        return self.add_sale(date, supplier_id, items, total_amount, cash_paid, previous_credit, updated_credit, notes, "purchase")
            
    def get_daily_sales(self, date):
        """Get all sales for a specific date."""
        conn = self.connect()
        cursor = conn.cursor()
        
        # Check table names to determine which query to use
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='transactions'")
        if cursor.fetchone():
            cursor.execute('''
            SELECT t.*, c.name as customer_name 
            FROM transactions t
            LEFT JOIN clients c ON t.customer_id = c.id
            WHERE t.date = ? AND t.transaction_type = 'sale'
            ''', (date,))
        else:
            cursor.execute('''
            SELECT s.*, c.name as customer_name 
            FROM sales s
            LEFT JOIN customers c ON s.customer_id = c.id
            WHERE s.date = ?
            ''', (date,))
            
        return cursor.fetchall()
        
    def get_daily_sales_detail(self, date):
        """Get detailed sales with items for a specific date."""
        conn = self.connect()
        cursor = conn.cursor()
        
        # Check table names to determine which query to use
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='transactions'")
        if cursor.fetchone():
            cursor.execute('''
            SELECT t.id, t.date, c.name as customer_name, t.total_amount, t.cash_received,
                   t.previous_credit, t.updated_credit, t.notes,
                   ti.product_id, p.name as product_name, ti.quantity, ti.unit_price as price
            FROM transactions t
            LEFT JOIN clients c ON t.customer_id = c.id
            LEFT JOIN transaction_items ti ON t.id = ti.transaction_id
            LEFT JOIN products p ON ti.product_id = p.id
            WHERE t.date = ? AND t.transaction_type = 'sale'
            ORDER BY t.id, ti.id
            ''', (date,))
        else:
            cursor.execute('''
            SELECT s.id, s.date, c.name as customer_name, s.total_amount, s.cash_received,
                   s.previous_credit, s.updated_credit, s.notes,
                   si.product_id, p.name as product_name, si.quantity, si.price
            FROM sales s
            LEFT JOIN customers c ON s.customer_id = c.id
            LEFT JOIN sale_items si ON s.id = si.sale_id
            LEFT JOIN products p ON si.product_id = p.id
            WHERE s.date = ? AND s.transaction_type = 'sale'
            ORDER BY s.id, si.id
            ''', (date,))
        
        sales = {}
        for row in cursor.fetchall():
            sale_id = row['id']
            if sale_id not in sales:
                sales[sale_id] = {
                    'id': sale_id,
                    'date': row['date'],
                    'customer_name': row['customer_name'],
                    'total_amount': row['total_amount'],
                    'cash_received': row['cash_received'],
                    'previous_credit': row['previous_credit'],
                    'updated_credit': row['updated_credit'],
                    'notes': row['notes'],
                    'items': []
                }
            
            if row['product_id']:
                sales[sale_id]['items'].append({
                    'product_id': row['product_id'],
                    'product_name': row['product_name'],
                    'quantity': row['quantity'],
                    'price': row['price']
                })
        
        return list(sales.values())
        
    def get_daily_purchases_detail(self, date):
        """Get detailed purchases with items for a specific date."""
        conn = self.connect()
        cursor = conn.cursor()
        
        # Check table names to determine which query to use
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='transactions'")
        if cursor.fetchone():
            cursor.execute('''
            SELECT t.id, t.date, s.name as supplier_name, t.total_amount, t.cash_received as cash_paid,
                   t.previous_credit, t.updated_credit, t.notes,
                   ti.product_id, p.name as product_name, ti.quantity, ti.unit_price as price
            FROM transactions t
            LEFT JOIN suppliers s ON t.customer_id = s.id
            LEFT JOIN transaction_items ti ON t.id = ti.transaction_id
            LEFT JOIN products p ON ti.product_id = p.id
            WHERE t.date = ? AND t.transaction_type = 'purchase'
            ORDER BY t.id, ti.id
            ''', (date,))
        else:
            cursor.execute('''
            SELECT s.id, s.date, c.name as supplier_name, s.total_amount, s.cash_received as cash_paid,
                   s.previous_credit, s.updated_credit, s.notes,
                   si.product_id, p.name as product_name, si.quantity, si.price
            FROM sales s
            LEFT JOIN customers c ON s.customer_id = c.id
            LEFT JOIN sale_items si ON s.id = si.sale_id
            LEFT JOIN products p ON si.product_id = p.id
            WHERE s.date = ? AND s.transaction_type = 'purchase'
            ORDER BY s.id, si.id
            ''', (date,))
        
        purchases = {}
        for row in cursor.fetchall():
            purchase_id = row['id']
            if purchase_id not in purchases:
                purchases[purchase_id] = {
                    'id': purchase_id,
                    'date': row['date'],
                    'supplier_name': row['supplier_name'],
                    'total_amount': row['total_amount'],
                    'cash_paid': row['cash_paid'],
                    'previous_credit': row['previous_credit'],
                    'updated_credit': row['updated_credit'],
                    'notes': row['notes'],
                    'items': []
                }
            
            if row['product_id']:
                purchases[purchase_id]['items'].append({
                    'product_id': row['product_id'],
                    'product_name': row['product_name'],
                    'quantity': row['quantity'],
                    'price': row['price']
                })
        
        return list(purchases.values())
        
    def save_daily_summary(self, date, total_sales, total_purchases, cash_received):
        """Save or update daily summary."""
        conn = self.connect()
        cursor = conn.cursor()
        try:
            cursor.execute('''
            INSERT INTO daily_summary (date, total_sales, total_purchases, cash_received)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(date) DO UPDATE SET
            total_sales = ?, total_purchases = ?, cash_received = ?
            ''', (date, total_sales, total_purchases, cash_received,
                  total_sales, total_purchases, cash_received))
            conn.commit()
            return True
        except Exception as e:
            print(f"Error saving daily summary: {e}")
            return False
            
    def get_daily_summary(self, date):
        """Get daily summary for a specific date."""
        conn = self.connect()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM daily_summary WHERE date = ?", (date,))
        return cursor.fetchone()
        
    def calculate_daily_inventory(self, date):
        """Calculate inventory changes for a specific date."""
        conn = self.connect()
        
        # Get daily stock data for the date
        cursor = conn.cursor()
        cursor.execute('''
        SELECT ds.*, p.name as product_name, p.price
        FROM daily_stock ds
        JOIN products p ON ds.product_id = p.id
        WHERE ds.date = ?
        ORDER BY p.name
        ''', (date,))
        
        inventory_summary = {}
        for row in cursor.fetchall():
            product_id = row['product_id']
            inventory_summary[product_id] = {
                'product_name': row['product_name'],
                'price': row['price'],
                'opening_stock': row['opening_stock'],
                'purchases': row['purchases'],
                'sales': row['sales'],
                'adjustments': row['adjustments'],
                'closing_stock': row['closing_stock'],
                'net_change': row['closing_stock'] - row['opening_stock']
            }
        
        # If no daily stock data exists, get it from stock movements
        if not inventory_summary:
            cursor.execute('''
            SELECT p.id as product_id, p.name as product_name, p.price, p.current_stock,
                   SUM(CASE WHEN sm.movement_type = 'purchase' THEN sm.quantity_change ELSE 0 END) as purchases,
                   SUM(CASE WHEN sm.movement_type = 'sale' THEN -sm.quantity_change ELSE 0 END) as sales,
                   SUM(CASE WHEN sm.movement_type NOT IN ('sale', 'purchase') THEN sm.quantity_change ELSE 0 END) as adjustments
            FROM products p
            LEFT JOIN stock_movements sm ON p.id = sm.product_id AND sm.date = ?
            WHERE p.active = 1
            GROUP BY p.id
            ORDER BY p.name
            ''', (date,))
            
            for row in cursor.fetchall():
                product_id = row['product_id']
                # Get the opening stock by subtracting changes from current stock
                opening_stock = row['current_stock'] - (row['purchases'] - row['sales'] + row['adjustments'])
                closing_stock = row['current_stock']
                
                inventory_summary[product_id] = {
                    'product_name': row['product_name'],
                    'price': row['price'],
                    'opening_stock': opening_stock,
                    'purchases': row['purchases'] or 0,
                    'sales': row['sales'] or 0,
                    'adjustments': row['adjustments'] or 0,
                    'closing_stock': closing_stock,
                    'net_change': closing_stock - opening_stock
                }
        
        # Calculate totals
        total_opening_value = 0
        total_closing_value = 0
        total_purchased_value = 0
        total_sold_value = 0
        
        for product_id, data in inventory_summary.items():
            price = data['price']
            total_opening_value += price * data['opening_stock']
            total_closing_value += price * data['closing_stock']
            total_purchased_value += price * data['purchases']
            total_sold_value += price * data['sales']
        
        # Return both product details and summary
        return {
            'products': inventory_summary,
            'total_opening_value': total_opening_value,
            'total_closing_value': total_closing_value,
            'total_purchased_value': total_purchased_value,
            'total_sold_value': total_sold_value,
            'net_value_change': total_closing_value - total_opening_value
        }
            
    def export_to_excel(self, start_date, end_date, filename):
        """Export data to Excel for a date range.
        
        This function saves all daily reports in a single Excel file with different sheets (tabs)
        for different dates. Each sheet follows the specified format with daily summary at the bottom.
        """
        conn = self.connect()
        
        try:
            # Check table names to determine which query to use
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='transactions'")
            use_new_schema = cursor.fetchone() is not None
            
            # Default Excel file path
            default_excel_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "milk_data.xlsx")
            
            # If no specific filename was provided, use the default
            if not filename:
                filename = default_excel_path
            
            # Get unique dates between start_date and end_date
            dates_query = '''
            SELECT DISTINCT date FROM transactions 
            WHERE date BETWEEN ? AND ? ORDER BY date
            '''
            cursor.execute(dates_query, (start_date, end_date))
            dates = [row['date'] for row in cursor.fetchall()]
            
            # If there are no dates within range, return
            if not dates:
                print(f"No transactions found between {start_date} and {end_date}")
                return False
            
            # Check if the file already exists
            file_exists = os.path.exists(filename)
            
            # Use ExcelWriter with openpyxl engine to allow editing existing files
            with pd.ExcelWriter(filename, engine='openpyxl', mode='a' if file_exists else 'w') as writer:
                # If the file exists, load it
                if file_exists:
                    try:
                        wb = openpyxl.load_workbook(filename)
                        # Access the book property through the underlying openpyxl writer
                        writer.sheets = {ws.title: ws for ws in wb.worksheets}
                        writer.book = wb
                    except Exception as e:
                        print(f"Error loading existing workbook: {e}")
                        # If can't load, create new workbook (handled automatically)
                
                # Process each date
                for date in dates:
                    # Create a sheet name from the date (YYYY-MM-DD becomes Sheet_YYYY-MM-DD)
                    sheet_name = f"{date}"
                    
                    # Skip if this sheet already exists
                    if file_exists and sheet_name in writer.book.sheetnames:
                        # You could choose to delete and recreate, but here we'll preserve it
                        print(f"Sheet {sheet_name} already exists, skipping.")
                        continue
                    
                    # Get transactions for this date
                    if use_new_schema:
                        transactions_df = pd.read_sql_query(
                            '''
                            SELECT t.id, t.date, c.name as entity_name, t.transaction_type, t.total_amount, 
                                   t.cash_received, t.previous_credit, t.updated_credit, t.notes
                            FROM transactions t
                            LEFT JOIN customers c ON t.customer_id = c.id
                            WHERE t.date = ?
                            ORDER BY t.transaction_type, t.id
                            ''', 
                            conn, 
                            params=[date]
                        )
                        
                        # Query transaction items
                        items_df = pd.read_sql_query(
                            '''
                            SELECT ti.transaction_id as sale_id, t.transaction_type, p.name as product_name, 
                                   ti.quantity, ti.unit_price as price,
                                   (ti.quantity * ti.unit_price) as subtotal
                            FROM transaction_items ti
                            JOIN products p ON ti.product_id = p.id
                            JOIN transactions t ON ti.transaction_id = t.id
                            WHERE t.date = ?
                            ORDER BY t.transaction_type, ti.transaction_id, ti.id
                            ''', 
                            conn, 
                            params=[date]
                        )
                    else:
                        transactions_df = pd.read_sql_query(
                            '''
                            SELECT s.id, s.date, c.name as entity_name, s.transaction_type, s.total_amount, 
                                   s.cash_received, s.previous_credit, s.updated_credit, s.notes
                            FROM sales s
                            LEFT JOIN customers c ON s.customer_id = c.id
                            WHERE s.date = ?
                            ORDER BY s.transaction_type, s.id
                            ''', 
                            conn, 
                            params=[date]
                        )
                        
                        # Query transaction items
                        items_df = pd.read_sql_query(
                            '''
                            SELECT si.sale_id, s.transaction_type, p.name as product_name, si.quantity, si.price,
                                   (si.quantity * si.price) as subtotal
                            FROM sale_items si
                            JOIN products p ON si.product_id = p.id
                            JOIN sales s ON si.sale_id = s.id
                            WHERE s.date = ?
                            ORDER BY s.transaction_type, si.sale_id, si.id
                            ''', 
                            conn, 
                            params=[date]
                        )
                    
                    # Query stock movements for this date
                    stock_movements_df = pd.read_sql_query(
                        '''
                        SELECT sm.date, p.name as product_name, sm.quantity_change, 
                               sm.movement_type, sm.notes
                        FROM stock_movements sm
                        JOIN products p ON sm.product_id = p.id
                        WHERE sm.date = ?
                        ORDER BY p.name, sm.movement_type
                        ''', 
                        conn, 
                        params=[date]
                    )
                    
                    # Query inventory data
                    inventory_df = pd.read_sql_query(
                        '''
                        SELECT p.id, p.name, p.price, p.current_stock, (p.price * p.current_stock) as inventory_value
                        FROM products p
                        WHERE p.active = 1
                        ORDER BY p.name
                        ''', 
                        conn
                    )
                    
                    # Create detailed daily report for this date
                    self._create_daily_report(writer, date, transactions_df, items_df, inventory_df, stock_movements_df)
            
            return True
            
        except Exception as e:
            import traceback
            print(f"Error exporting to Excel: {e}")
            traceback.print_exc()
            return False
        
        finally:
            # Always close the connection
            if conn:
                conn.close()
        
    def _create_daily_report(self, excel_writer, date, transactions_df, items_df, inventory_df, stock_movements_df):
        """Create a detailed daily report in Excel format similar to the example."""
        # Filter data for the specific date
        day_transactions = transactions_df[transactions_df['date'] == date]
        day_sales = day_transactions[day_transactions['transaction_type'] == 'sale']
        day_purchases = day_transactions[day_transactions['transaction_type'] == 'purchase']
        day_items = items_df[items_df['sale_id'].isin(day_transactions['id'])]
        
        # Group sales by customer and product
        sales_by_customer = {}
        purchases_by_supplier = {}
        
        # Process sales
        for _, sale in day_sales.iterrows():
            sale_id = sale['id']
            customer_name = sale['entity_name']
            
            if customer_name not in sales_by_customer:
                sales_by_customer[customer_name] = {
                    'total': sale['total_amount'],
                    'cash_received': sale['cash_received'],
                    'previous_credit': sale['previous_credit'],
                    'updated_credit': sale['updated_credit'],
                    'items': {}
                }
            
            # Add items
            sale_items = day_items[day_items['sale_id'] == sale_id]
            for _, item in sale_items.iterrows():
                product_name = item['product_name']
                if product_name not in sales_by_customer[customer_name]['items']:
                    sales_by_customer[customer_name]['items'][product_name] = 0
                sales_by_customer[customer_name]['items'][product_name] += item['quantity']
        
        # Process purchases
        for _, purchase in day_purchases.iterrows():
            purchase_id = purchase['id']
            supplier_name = purchase['entity_name']
            
            if supplier_name not in purchases_by_supplier:
                purchases_by_supplier[supplier_name] = {
                    'total': purchase['total_amount'],
                    'cash_paid': purchase['cash_received'],
                    'previous_credit': purchase['previous_credit'],
                    'updated_credit': purchase['updated_credit'],
                    'items': {}
                }
            
            # Add items
            purchase_items = day_items[day_items['sale_id'] == purchase_id]
            for _, item in purchase_items.iterrows():
                product_name = item['product_name']
                if product_name not in purchases_by_supplier[supplier_name]['items']:
                    purchases_by_supplier[supplier_name]['items'][product_name] = 0
                purchases_by_supplier[supplier_name]['items'][product_name] += item['quantity']
        
        # Calculate stock changes for each product
        stock_changes = {}
        
        # Group stock movements by product
        day_movements = stock_movements_df[stock_movements_df['date'] == date]
        
        # Connect to database to get daily stock data
        conn = self.connect()
        cursor = conn.cursor()
        
        # Get daily stock data for this date which will include opening/closing balances
        cursor.execute('''
        SELECT ds.*, p.name as product_name, p.price
        FROM daily_stock ds
        JOIN products p ON ds.product_id = p.id
        WHERE ds.date = ?
        ORDER BY p.name
        ''', (date,))
        
        daily_stock_records = cursor.fetchall()
        
        # If we have daily stock records, use those for the stock changes
        if daily_stock_records:
            for record in daily_stock_records:
                product_name = record['product_name']
                stock_changes[product_name] = {
                    'product_id': record['product_id'],
                    'price': record['price'],
                    'opening_stock': record['opening_stock'],
                    'total_sold': record['sales'],
                    'total_purchased': record['purchases'],
                    'net_change': record['closing_stock'] - record['opening_stock'],
                    'closing_stock': record['closing_stock']
                }
        else:
            # If no daily stock records, calculate from stock movements
            for product_name in day_movements['product_name'].unique():
                product_movements = day_movements[day_movements['product_name'] == product_name]
                
                # Calculate totals by movement type
                total_sold = abs(sum(product_movements[product_movements['movement_type'] == 'sale']['quantity_change']))
                total_purchased = sum(product_movements[product_movements['movement_type'] == 'purchase']['quantity_change'])
                net_change = sum(product_movements['quantity_change'])
                
                # Get current stock for this product
                cursor.execute('''
                SELECT id, price, current_stock
                FROM products
                WHERE name = ?
                ''', (product_name,))
                
                product_data = cursor.fetchone()
                if product_data:
                    closing_stock = product_data['current_stock']
                    opening_stock = closing_stock - net_change
                    
                    stock_changes[product_name] = {
                        'product_id': product_data['id'],
                        'price': product_data['price'],
                        'opening_stock': opening_stock,
                        'total_sold': total_sold,
                        'total_purchased': total_purchased,
                        'net_change': net_change,
                        'closing_stock': closing_stock
                    }
        
        # Close database connection
        conn.close()
        
        # Calculate monetary values
        total_opening_value = sum(data['price'] * data['opening_stock'] for data in stock_changes.values())
        total_closing_value = sum(data['price'] * data['closing_stock'] for data in stock_changes.values())
        total_sold_value = sum(data['price'] * data['total_sold'] for data in stock_changes.values())
        total_purchased_value = sum(data['price'] * data['total_purchased'] for data in stock_changes.values())
        
        # Now create the Excel sheet
        df = pd.DataFrame()
        
        # Start with header info and title
        row_index = 0
        df.at[row_index, 'A'] = f"Daily Report for {date}"
        row_index += 2
        
        # Add sales transactions section
        df.at[row_index, 'A'] = "SALES TRANSACTIONS"
        row_index += 1
        
        if sales_by_customer:
            df.at[row_index, 'A'] = "Customer Name"
            df.at[row_index, 'B'] = "FC 1L"
            df.at[row_index, 'C'] = "$ 1/2L"
            df.at[row_index, 'D'] = "$ 140ml"
            df.at[row_index, 'E'] = "$ 1L"
            df.at[row_index, 'F'] = "Curd 130"
            df.at[row_index, 'G'] = "Curd 1K"
            df.at[row_index, 'H'] = "Extra"
            df.at[row_index, 'I'] = "Total Amount"
            df.at[row_index, 'J'] = "Previous Credit"
            df.at[row_index, 'K'] = "Cash Received"
            df.at[row_index, 'L'] = "Updated Credit"
            row_index += 1
            
            # Add individual customer rows
            for customer, data in sales_by_customer.items():
                df.at[row_index, 'A'] = customer
                
                # Fill in product quantities
                for product, qty in data['items'].items():
                    if "FC 1L" in product.upper():
                        df.at[row_index, 'B'] = qty
                    elif "1/2L" in product.upper() or "500ML" in product.upper():
                        df.at[row_index, 'C'] = qty
                    elif "140ML" in product.upper():
                        df.at[row_index, 'D'] = qty
                    elif "1L" in product.upper() and "FC" not in product.upper():
                        df.at[row_index, 'E'] = qty
                    elif "CURD 130" in product.upper():
                        df.at[row_index, 'F'] = qty
                    elif "CURD 1K" in product.upper():
                        df.at[row_index, 'G'] = qty
                    else:
                        # Other products in Extra column
                        current_extra = df.at[row_index, 'H'] if pd.notna(df.at[row_index, 'H']) else ""
                        df.at[row_index, 'H'] = f"{current_extra}{product}: {qty}, "
                
                df.at[row_index, 'I'] = data['total']
                df.at[row_index, 'J'] = data['previous_credit']
                df.at[row_index, 'K'] = data['cash_received']
                df.at[row_index, 'L'] = data['updated_credit']
                row_index += 1
            
            # Calculate and add sales total row
            df.at[row_index, 'A'] = "TOTAL SALES"
            df.at[row_index, 'I'] = sum(data['total'] for data in sales_by_customer.values())
            df.at[row_index, 'K'] = sum(data['cash_received'] for data in sales_by_customer.values())
            row_index += 2
        else:
            df.at[row_index, 'A'] = "No sales transactions for this date"
            row_index += 2
        
        # Add purchase transactions section
        df.at[row_index, 'A'] = "PURCHASE TRANSACTIONS"
        row_index += 1
        
        if purchases_by_supplier:
            df.at[row_index, 'A'] = "Supplier"
            df.at[row_index, 'B'] = "Items Purchased"
            df.at[row_index, 'I'] = "Total Amount"
            df.at[row_index, 'J'] = "Previous Credit"
            df.at[row_index, 'K'] = "Cash Paid"
            df.at[row_index, 'L'] = "Updated Credit"
            row_index += 1
            
            # Add individual supplier rows
            for supplier, data in purchases_by_supplier.items():
                df.at[row_index, 'A'] = supplier
                
                # Combine items into a single string
                items_str = ", ".join([f"{product}: {qty}" for product, qty in data['items'].items()])
                df.at[row_index, 'B'] = items_str
                
                df.at[row_index, 'I'] = data['total']
                df.at[row_index, 'J'] = data['previous_credit']
                df.at[row_index, 'K'] = data['cash_paid']
                df.at[row_index, 'L'] = data['updated_credit']
                row_index += 1
            
            # Calculate and add purchases total row
            df.at[row_index, 'A'] = "TOTAL PURCHASES"
            df.at[row_index, 'I'] = sum(data['total'] for data in purchases_by_supplier.values())
            df.at[row_index, 'K'] = sum(data['cash_paid'] for data in purchases_by_supplier.values())
            row_index += 2
        else:
            df.at[row_index, 'A'] = "No purchase transactions for this date"
            row_index += 2
        
        # Add inventory daily summary section
        df.at[row_index, 'A'] = "DAILY INVENTORY SUMMARY"
        row_index += 1
        
        df.at[row_index, 'A'] = "Product Name"
        df.at[row_index, 'B'] = "Opening Stock"
        df.at[row_index, 'C'] = "Purchased"
        df.at[row_index, 'D'] = "Sold"
        df.at[row_index, 'E'] = "Closing Stock"
        df.at[row_index, 'F'] = "Unit Price"
        df.at[row_index, 'G'] = "Opening Value"
        df.at[row_index, 'H'] = "Closing Value"
        row_index += 1
        
        # Add product stock movements
        for product_name, change in stock_changes.items():
            df.at[row_index, 'A'] = product_name
            df.at[row_index, 'B'] = change['opening_stock']
            df.at[row_index, 'C'] = change['total_purchased']
            df.at[row_index, 'D'] = change['total_sold']
            df.at[row_index, 'E'] = change['closing_stock']
            df.at[row_index, 'F'] = change['price']
            df.at[row_index, 'G'] = change['price'] * change['opening_stock']
            df.at[row_index, 'H'] = change['price'] * change['closing_stock']
            row_index += 1
        
        # Add totals row
        df.at[row_index, 'A'] = "TOTALS"
        df.at[row_index, 'B'] = sum(change['opening_stock'] for change in stock_changes.values())
        df.at[row_index, 'C'] = sum(change['total_purchased'] for change in stock_changes.values())
        df.at[row_index, 'D'] = sum(change['total_sold'] for change in stock_changes.values())
        df.at[row_index, 'E'] = sum(change['closing_stock'] for change in stock_changes.values())
        df.at[row_index, 'G'] = total_opening_value
        df.at[row_index, 'H'] = total_closing_value
        row_index += 2
        
        # Skip several rows before adding the daily summary
        row_index += 2

        # Add final summary section at the bottom of the report with a slightly different format
        df.at[row_index, 'A'] = "--- Daily Summary ---"
        row_index += 1
        
        # Calculate product totals for the summary
        total_sold = {
            'FC 1L': 0,
            '1/2L': 0,
            '140ml': 0,
            '1L': 0,
            'Curd 130': 0,
            'Curd 1K': 0
        }
        
        # Count sold quantities by product type
        for customer, data in sales_by_customer.items():
            for product, qty in data['items'].items():
                if "FC 1L" in product.upper():
                    total_sold['FC 1L'] += qty
                elif "1/2L" in product.upper() or "500ML" in product.upper():
                    total_sold['1/2L'] += qty
                elif "140ML" in product.upper():
                    total_sold['140ml'] += qty
                elif "1L" in product.upper() and "FC" not in product.upper():
                    total_sold['1L'] += qty
                elif "CURD 130" in product.upper():
                    total_sold['Curd 130'] += qty
                elif "CURD 1K" in product.upper():
                    total_sold['Curd 1K'] += qty
        
        # Calculate purchases by product
        total_purchased = {
            'FC 1L': 0,
            '1/2L': 0,
            '140ml': 0,
            '1L': 0,
            'Curd 130': 0,
            'Curd 1K': 0
        }
        
        # Count purchased quantities by product type
        for supplier, data in purchases_by_supplier.items():
            for product, qty in data['items'].items():
                if "FC 1L" in product.upper():
                    total_purchased['FC 1L'] += qty
                elif "1/2L" in product.upper() or "500ML" in product.upper():
                    total_purchased['1/2L'] += qty
                elif "140ML" in product.upper():
                    total_purchased['140ml'] += qty
                elif "1L" in product.upper() and "FC" not in product.upper():
                    total_purchased['1L'] += qty
                elif "CURD 130" in product.upper():
                    total_purchased['Curd 130'] += qty
                elif "CURD 1K" in product.upper():
                    total_purchased['Curd 1K'] += qty
        
        # Add the Product Names row
        df.at[row_index, 'A'] = "Product Name"
        df.at[row_index, 'B'] = "Total Sold"
        df.at[row_index, 'C'] = "Total Bought"
        df.at[row_index, 'D'] = "Net Change (Bought-Sold)"
        row_index += 1
        
        # Add the summary rows for each product type
        for product in total_sold.keys():
            df.at[row_index, 'A'] = product
            df.at[row_index, 'B'] = total_sold[product]
            df.at[row_index, 'C'] = total_purchased[product]
            df.at[row_index, 'D'] = total_purchased[product] - total_sold[product]
            row_index += 1
            
        # Skip a row
        row_index += 1
        
        # Add the other summary information
        df.at[row_index, 'A'] = "Total Sales"
        df.at[row_index, 'B'] = sum(data['total'] for data in sales_by_customer.values()) if sales_by_customer else 0
        row_index += 1
        
        df.at[row_index, 'A'] = "Total Purchases"
        df.at[row_index, 'B'] = sum(data['total'] for data in purchases_by_supplier.values()) if purchases_by_supplier else 0
        row_index += 1
        
        df.at[row_index, 'A'] = "Cash Received"
        df.at[row_index, 'B'] = sum(data['cash_received'] for data in sales_by_customer.values()) if sales_by_customer else 0
        row_index += 1
        
        # Calculate values for the summary
        cash_in = sum(data['cash_received'] for data in sales_by_customer.values()) if sales_by_customer else 0
        cash_out = sum(data['cash_paid'] for data in purchases_by_supplier.values()) if purchases_by_supplier else 0
        
        # Add inventory value change
        df.at[row_index, 'A'] = "Inventory Value Change"
        df.at[row_index, 'B'] = total_closing_value - total_opening_value
        
        # Write to Excel with formatting
        sheet_name = f"Daily Report {date}"
        df.to_excel(excel_writer, sheet_name=sheet_name, index=False)
        
        # Apply some basic formatting
        worksheet = excel_writer.sheets[sheet_name]
        
        # The rest of the formatting will be handled by pandas

class GoogleSheetsManager:
    """Manages synchronization with Google Sheets."""
    
    def __init__(self, credentials_file=CREDENTIALS_FILE, spreadsheet_id=None):
        self.credentials_file = credentials_file
        self.spreadsheet_id = spreadsheet_id
        self.client = None
        self.spreadsheet = None
        
    def authenticate(self):
        """Authenticate with Google Sheets API."""
        try:
            # Check if credentials file exists
            if not os.path.exists(self.credentials_file):
                print(f"Credentials file not found: {self.credentials_file}")
                return False
                
            # Set up credentials
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                self.credentials_file, SCOPES)
                
            # Connect to Google Drive API
            self.client = gspread.authorize(creds)
            
            # Check if spreadsheet ID is available
            if not self.spreadsheet_id:
                print("No spreadsheet ID provided")
                return False
                
            # Open the spreadsheet
            self.spreadsheet = self.client.open_by_key(self.spreadsheet_id)
            print("Successfully connected to Google Sheets")
            return True
        except Exception as e:
            print(f"Google Sheets authentication failed: {e}")
            traceback.print_exc()
            return False
            
    def get_worksheet(self, date_str):
        """Get a worksheet for a specific date, creating if needed."""
        if not self.client or not self.spreadsheet:
            print("Not connected to Google Sheets")
            return None
            
        try:
            # Try to get the worksheet
            try:
                worksheet = self.spreadsheet.worksheet(date_str)
                return worksheet
            except gspread.exceptions.WorksheetNotFound:
                # Create new worksheet
                worksheet = self.spreadsheet.add_worksheet(
                    title=date_str, rows=100, cols=20)
                return worksheet
        except Exception as e:
            print(f"Error getting worksheet: {e}")
            return None
            
    def sync_sales_to_sheet(self, date_str, sales_data, header_row):
        """Sync sales data to Google Sheet for a specific date."""
        if not self.client or not self.spreadsheet:
            print("Not connected to Google Sheets")
            return False
            
        worksheet = self.get_worksheet(date_str)
        if not worksheet:
            return False
            
        try:
            # Add header row if needed
            try:
                current_values = worksheet.get_all_values()
                if not current_values or len(current_values) == 0:
                    worksheet.append_row(header_row)
            except:
                worksheet.append_row(header_row)
                
            # Add sales data
            for sale in sales_data:
                row_data = self._convert_sale_to_row(sale, header_row)
                worksheet.append_row(row_data)
                
            return True
        except Exception as e:
            print(f"Error syncing sales to sheet: {e}")
            return False
            
    def sync_purchases_to_sheet(self, date_str, purchases_data, header_row):
        """Sync purchases data to Google Sheet for a specific date."""
        if not self.client or not self.spreadsheet:
            print("Not connected to Google Sheets")
            return False
            
        worksheet = self.get_worksheet(date_str)
        if not worksheet:
            return False
            
        try:
            # Add purchases section header
            try:
                # Find the last row
                values = worksheet.get_all_values()
                last_row = len(values) + 1 if values else 1
                
                # Add a blank row for separation if there's data
                if last_row > 1:
                    worksheet.append_row([""] * len(header_row))
                    last_row += 1
                
                # Add a purchases section header
                purchases_header = ["PURCHASES"] + [""] * (len(header_row) - 1)
                worksheet.append_row(purchases_header)
                last_row += 1
                
                # Add the header row
                worksheet.append_row(header_row)
                last_row += 1
            except Exception as e:
                print(f"Error adding purchases header: {e}")
                # If there's an error, still try to append the header
                worksheet.append_row(header_row)
                
            # Add purchases data
            for purchase in purchases_data:
                row_data = self._convert_purchase_to_row(purchase, header_row)
                worksheet.append_row(row_data)
                
            return True
        except Exception as e:
            print(f"Error syncing purchases to sheet: {e}")
            return False
            
    def _convert_purchase_to_row(self, purchase, header_row):
        """Convert a purchase object to a row format for sheets."""
        row_data = []
        
        # Process each header column
        for header in header_row:
            if header == "Supplier Name":
                row_data.append(purchase.get('supplier_name', ''))
            elif header == "Total Amount for":
                row_data.append(purchase.get('total_amount', 0))
            elif header == "Previous Credit":
                row_data.append(purchase.get('previous_credit', 0))
            elif header == "Cash Paid":
                row_data.append(purchase.get('cash_paid', 0))
            elif header == "Updated Credit Balance":
                row_data.append(purchase.get('updated_credit', 0))
            elif header == "Extra":
                # Format items for the Extra column
                items_str = ", ".join([
                    f"{item['product_name']}: {item['quantity']}" 
                    for item in purchase.get('items', [])
                ])
                row_data.append(items_str)
            else:
                # Check if this is a product column
                found = False
                for item in purchase.get('items', []):
                    if header == item['product_name']:
                        row_data.append(item['quantity'])
                        found = True
                        break
                        
                if not found:
                    row_data.append("")
                    
        return row_data
            
    def _convert_sale_to_row(self, sale, header_row):
        """Convert a sale object to a row format for sheets."""
        row_data = []
        
        # Process each header column
        for header in header_row:
            if header == "Customer Name":
                row_data.append(sale.get('customer_name', ''))
            elif header == "Total Amount for":
                row_data.append(sale.get('total_amount', 0))
            elif header == "Previous Credit":
                row_data.append(sale.get('previous_credit', 0))
            elif header == "Cash Received":
                row_data.append(sale.get('cash_received', 0))
            elif header == "Updated Credit Balance":
                row_data.append(sale.get('updated_credit', 0))
            elif header == "Extra":
                # Format items for the Extra column
                items_str = ", ".join([
                    f"{item['product_name']}: {item['quantity']}" 
                    for item in sale.get('items', [])
                ])
                row_data.append(items_str)
            else:
                # Check if this is a product column
                found = False
                for item in sale.get('items', []):
                    if header == item['product_name']:
                        row_data.append(item['quantity'])
                        found = True
                        break
                        
                if not found:
                    row_data.append("")
                    
        return row_data

class BackupManager:
    """Manages backup of data to Excel files."""
    
    def __init__(self, backup_folder=BACKUP_FOLDER, db_manager=None):
        self.backup_folder = backup_folder
        self.db_manager = db_manager
        ensure_directory_exists(backup_folder)
        
    def backup_current_date(self, date=None):
        """Backup data for a specific date."""
        if not date:
            date = datetime.datetime.now().strftime("%Y-%m-%d")
            
        if not self.db_manager:
            print("Database manager not available")
            return False
            
        # Create filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.backup_folder, f"daily_backup_{date}_{timestamp}.xlsx")
        
        # Export to Excel
        return self.db_manager.export_to_excel(date, date, filename)
        
    def backup_date_range(self, start_date, end_date):
        """Backup data for a date range."""
        if not self.db_manager:
            print("Database manager not available")
            return False
            
        # Create filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.join(self.backup_folder, 
                               f"backup_{start_date}_to_{end_date}_{timestamp}.xlsx")
        
        # Export to Excel
        return self.db_manager.export_to_excel(start_date, end_date, filename)
        
    def backup_full_database(self):
        """Create a full backup of the database file."""
        if not self.db_manager:
            print("Database manager not available")
            return False
            
        # Close database connection
        self.db_manager.close()
        
        # Create filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(self.backup_folder, f"full_backup_{timestamp}.db")
        
        try:
            # Copy the database file
            shutil.copy2(self.db_manager.db_file, backup_file)
            return True
        except Exception as e:
            print(f"Error creating full backup: {e}")
            return False
        finally:
            # Reopen database connection
            self.db_manager.connect()

class ConfigScreen:
    """Configuration screen for the POS system."""
    
    def __init__(self, parent, config_manager, on_save_callback=None):
        self.parent = parent
        self.config_manager = config_manager
        self.on_save_callback = on_save_callback
        
        # Create the window
        self.window = tk.Toplevel(parent)
        self.window.title("POS System Configuration")
        self.window.geometry("600x500")
        self.window.grab_set()  # Make the window modal
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        self.general_tab = ttk.Frame(self.notebook, padding=10)
        self.online_tab = ttk.Frame(self.notebook, padding=10)
        self.backup_tab = ttk.Frame(self.notebook, padding=10)
        self.company_tab = ttk.Frame(self.notebook, padding=10)
        
        self.notebook.add(self.general_tab, text="General Settings")
        self.notebook.add(self.online_tab, text="Online Mode")
        self.notebook.add(self.backup_tab, text="Backup Settings")
        self.notebook.add(self.company_tab, text="Company Details")
        
        # Load current settings
        self.settings = self.config_manager.config.copy()
        
        # Create form controls
        self.create_general_settings()
        self.create_online_settings()
        self.create_backup_settings()
        self.create_company_settings()
        
        # Create bottom buttons
        self.create_buttons()
        
    def create_general_settings(self):
        """Create general settings controls."""
        # Mode selection
        ttk.Label(self.general_tab, text="Application Mode:").grid(row=0, column=0, sticky="w", pady=5)
        self.mode_var = tk.StringVar(value=self.settings.get("app_mode", "offline"))
        mode_frame = ttk.Frame(self.general_tab)
        mode_frame.grid(row=0, column=1, sticky="w", pady=5)
        ttk.Radiobutton(mode_frame, text="Online", value="online", variable=self.mode_var).pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="Offline", value="offline", variable=self.mode_var).pack(side="left", padx=5)
        
        # Theme selection
        ttk.Label(self.general_tab, text="Theme:").grid(row=1, column=0, sticky="w", pady=5)
        self.theme_var = tk.StringVar(value=self.settings.get("theme", "light"))
        theme_frame = ttk.Frame(self.general_tab)
        theme_frame.grid(row=1, column=1, sticky="w", pady=5)
        ttk.Radiobutton(theme_frame, text="Light", value="light", variable=self.theme_var).pack(side="left", padx=5)
        ttk.Radiobutton(theme_frame, text="Dark", value="dark", variable=self.theme_var).pack(side="left", padx=5)
        
        # Language selection
        ttk.Label(self.general_tab, text="Language:").grid(row=2, column=0, sticky="w", pady=5)
        self.language_var = tk.StringVar(value=self.settings.get("language", "en"))
        language_combo = ttk.Combobox(self.general_tab, textvariable=self.language_var,
                                   values=["en", "ta"], width=10, state="readonly")
        language_combo.grid(row=2, column=1, sticky="w", pady=5)
        ttk.Label(self.general_tab, text="(en=English, ta=Tamil)").grid(row=2, column=2, sticky="w", pady=5)
        
    def create_online_settings(self):
        """Create online mode settings controls."""
        # Credentials file
        ttk.Label(self.online_tab, text="Google Credentials File:").grid(row=0, column=0, sticky="w", pady=5)
        creds_frame = ttk.Frame(self.online_tab)
        creds_frame.grid(row=0, column=1, columnspan=2, sticky="w", pady=5)
        
        self.creds_path_var = tk.StringVar(value=CREDENTIALS_FILE)
        ttk.Entry(creds_frame, textvariable=self.creds_path_var, width=30).pack(side="left", padx=5)
        ttk.Button(creds_frame, text="Browse...", command=self.browse_credentials).pack(side="left", padx=5)
        
        # Spreadsheet ID
        ttk.Label(self.online_tab, text="Google Spreadsheet ID:").grid(row=1, column=0, sticky="w", pady=5)
        self.spreadsheet_id_var = tk.StringVar(value=self.settings.get("spreadsheet_id", ""))
        ttk.Entry(self.online_tab, textvariable=self.spreadsheet_id_var, width=40).grid(row=1, column=1, sticky="w", pady=5)
        
        # Test connection button
        ttk.Button(self.online_tab, text="Test Connection", command=self.test_connection).grid(row=2, column=0, pady=15)
        
        # Status indicator
        self.connection_status_var = tk.StringVar(value="Not connected")
        ttk.Label(self.online_tab, textvariable=self.connection_status_var).grid(row=2, column=1, sticky="w", pady=15)
        
        # Help text
        help_frame = ttk.LabelFrame(self.online_tab, text="Help")
        help_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)
        
        help_text = (
            "To use online mode, you need a Google Service Account credentials file.\n"
            "1. Go to the Google Cloud Console\n"
            "2. Create a project and enable the Google Sheets API\n"
            "3. Create service account credentials and download as JSON\n"
            "4. Share your Google Spreadsheet with the service account email\n"
            "5. Enter the Spreadsheet ID (found in the URL of your sheet)"
        )
        
        ttk.Label(help_frame, text=help_text, wraplength=500, justify="left").pack(padx=10, pady=10)
        
    def create_backup_settings(self):
        """Create backup settings controls."""
        # Backup folder
        ttk.Label(self.backup_tab, text="Backup Folder:").grid(row=0, column=0, sticky="w", pady=5)
        backup_frame = ttk.Frame(self.backup_tab)
        backup_frame.grid(row=0, column=1, columnspan=2, sticky="w", pady=5)
        
        self.backup_folder_var = tk.StringVar(value=self.settings.get("backup_folder", BACKUP_FOLDER))
        ttk.Entry(backup_frame, textvariable=self.backup_folder_var, width=30).pack(side="left", padx=5)
        ttk.Button(backup_frame, text="Browse...", command=self.browse_backup_folder).pack(side="left", padx=5)
        
        # Auto backup
        ttk.Label(self.backup_tab, text="Automatic Backups:").grid(row=1, column=0, sticky="w", pady=5)
        self.auto_backup_var = tk.BooleanVar(value=self.settings.get("auto_backup", True))
        ttk.Checkbutton(self.backup_tab, variable=self.auto_backup_var).grid(row=1, column=1, sticky="w", pady=5)
        
        # Backup interval
        ttk.Label(self.backup_tab, text="Backup Interval (days):").grid(row=2, column=0, sticky="w", pady=5)
        self.backup_interval_var = tk.IntVar(value=self.settings.get("backup_interval_days", 1))
        backup_interval_spin = ttk.Spinbox(self.backup_tab, from_=1, to=30, textvariable=self.backup_interval_var, width=5)
        backup_interval_spin.grid(row=2, column=1, sticky="w", pady=5)
        
        # Manual backup buttons
        ttk.Label(self.backup_tab, text="Manual Backup:").grid(row=3, column=0, sticky="w", pady=15)
        backup_buttons_frame = ttk.Frame(self.backup_tab)
        backup_buttons_frame.grid(row=3, column=1, sticky="w", pady=15)
        
        ttk.Button(backup_buttons_frame, text="Backup Today's Data", 
                  command=lambda: self.manual_backup("today")).pack(side="left", padx=5)
        ttk.Button(backup_buttons_frame, text="Backup Full Database", 
                  command=lambda: self.manual_backup("full")).pack(side="left", padx=5)
        
    def create_company_settings(self):
        """Create company details settings."""
        # Company name
        ttk.Label(self.company_tab, text="Company Name:").grid(row=0, column=0, sticky="w", pady=5)
        self.company_name_var = tk.StringVar(value=self.settings.get("company_name", ""))
        ttk.Entry(self.company_tab, textvariable=self.company_name_var, width=40).grid(row=0, column=1, sticky="w", pady=5)
        
        # Company address
        ttk.Label(self.company_tab, text="Address:").grid(row=1, column=0, sticky="w", pady=5)
        self.company_address_var = tk.StringVar(value=self.settings.get("company_address", ""))
        address_entry = ttk.Entry(self.company_tab, textvariable=self.company_address_var, width=40)
        address_entry.grid(row=1, column=1, sticky="w", pady=5)
        
        # Company phone
        ttk.Label(self.company_tab, text="Phone:").grid(row=2, column=0, sticky="w", pady=5)
        self.company_phone_var = tk.StringVar(value=self.settings.get("company_phone", ""))
        ttk.Entry(self.company_tab, textvariable=self.company_phone_var, width=40).grid(row=2, column=1, sticky="w", pady=5)
        
        # Company email
        ttk.Label(self.company_tab, text="Email:").grid(row=3, column=0, sticky="w", pady=5)
        self.company_email_var = tk.StringVar(value=self.settings.get("company_email", ""))
        ttk.Entry(self.company_tab, textvariable=self.company_email_var, width=40).grid(row=3, column=1, sticky="w", pady=5)
        
    def create_buttons(self):
        """Create bottom buttons."""
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(button_frame, text="Save & Close", command=self.save_settings).pack(side="right", padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.window.destroy).pack(side="right", padx=5)
        
    def browse_credentials(self):
        """Browse for credentials file."""
        filename = filedialog.askopenfilename(
            title="Select Google Credentials File",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if filename:
            self.creds_path_var.set(filename)
            
    def browse_backup_folder(self):
        """Browse for backup folder."""
        folder = filedialog.askdirectory(title="Select Backup Folder")
        if folder:
            self.backup_folder_var.set(folder)
            
    def test_connection(self):
        """Test Google Sheets connection."""
        creds_file = self.creds_path_var.get()
        sheet_id = self.spreadsheet_id_var.get()
        
        if not os.path.exists(creds_file):
            self.connection_status_var.set("Error: Credentials file not found")
            return
            
        if not sheet_id:
            self.connection_status_var.set("Error: Spreadsheet ID is required")
            return
            
        self.connection_status_var.set("Testing connection...")
        self.window.update_idletasks()
        
        # Try to connect
        sheets_manager = GoogleSheetsManager(creds_file, sheet_id)
        success = sheets_manager.authenticate()
        
        if success:
            self.connection_status_var.set("Connected successfully!")
        else:
            self.connection_status_var.set("Connection failed. Check credentials and ID.")
            
    def manual_backup(self, backup_type):
        """Perform manual backup."""
        # Get backup folder
        backup_folder = self.backup_folder_var.get()
        
        # Create backup manager
        db_manager = DatabaseManager()
        backup_manager = BackupManager(backup_folder, db_manager)
        
        if backup_type == "today":
            success = backup_manager.backup_current_date()
            if success:
                messagebox.showinfo("Backup Complete", 
                                   "Today's data has been backed up successfully.")
            else:
                messagebox.showerror("Backup Failed", 
                                    "Failed to create backup. Check logs for details.")
        elif backup_type == "full":
            success = backup_manager.backup_full_database()
            if success:
                messagebox.showinfo("Backup Complete", 
                                   "Full database backup created successfully.")
            else:
                messagebox.showerror("Backup Failed", 
                                    "Failed to create backup. Check logs for details.")
                
    def save_settings(self):
        """Save settings and close window."""
        # Collect all settings into a dictionary
        new_settings = {
            "app_mode": self.mode_var.get(),
            "theme": self.theme_var.get(),
            "language": self.language_var.get(),
            "spreadsheet_id": self.spreadsheet_id_var.get(),
            "backup_folder": self.backup_folder_var.get(),
            "auto_backup": self.auto_backup_var.get(),
            "backup_interval_days": self.backup_interval_var.get(),
            "company_name": self.company_name_var.get(),
            "company_address": self.company_address_var.get(),
            "company_phone": self.company_phone_var.get(),
            "company_email": self.company_email_var.get()
        }
        
        # Save to configuration
        self.config_manager.save_config(new_settings)
        
        # Call callback if provided
        if self.on_save_callback:
            self.on_save_callback(new_settings)
            
        # Close window
        self.window.destroy()

class POSApp:
    """Main POS application class."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Enhanced POS System")
        self.root.geometry("1200x800")
        
        # Initialize managers
        self.config_manager = ConfigManager()
        self.db_manager = DatabaseManager()
        self.backup_manager = BackupManager(
            self.config_manager.get("backup_folder", BACKUP_FOLDER), 
            self.db_manager)
        self.report_manager = ReportManager(self.db_manager, BACKUP_FOLDER)
        
        # Initialize ReportManager with company info
        company_info = {
            'name': self.config_manager.get('company_name', 'Enhanced POS System'),
            'address': self.config_manager.get('company_address', ''),
            'phone': self.config_manager.get('company_phone', ''),
            'email': self.config_manager.get('company_email', ''),
        }
        self.report_manager.set_company_info(company_info)
            
        # Initialize Google Sheets manager if in online mode
        self.sheets_manager = None
        if self.config_manager.get("app_mode") == "online":
            self.initialize_sheets_manager()
            
        # Configure styles
        self.style = ttk.Style()
        self.apply_theme()
        
        # Create main frame
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill="both", expand=True)
        
        # Create UI components
        self.create_menu_bar()
        self.create_status_bar()
        self.create_main_content()
        
    def initialize_sheets_manager(self):
        """Initialize the Google Sheets manager."""
        try:
            self.sheets_manager = GoogleSheetsManager(
                CREDENTIALS_FILE,
                self.config_manager.get("spreadsheet_id")
            )
            
            # Try to authenticate
            success = self.sheets_manager.authenticate()
            if success:
                self.set_status("Connected to Google Sheets", "online")
            else:
                self.set_status("Offline - Google Sheets authentication failed", "offline")
        except Exception as e:
            print(f"Error initializing Google Sheets: {e}")
            self.set_status("Offline - Google Sheets error", "offline")
            
    def apply_theme(self):
        """Apply the selected theme to the application."""
        theme = self.config_manager.get("theme", "light")
        
        if theme == "dark":
            # Dark theme colors
            bg_color = "#2d2d2d"
            fg_color = "#ffffff"
            button_bg = "#3d3d3d"
            button_fg = "#ffffff"
            highlight_bg = "#5d82c1"
            highlight_fg = "#ffffff"
        else:
            # Light theme colors
            bg_color = "#f5f5f5"
            fg_color = "#212121"
            button_bg = "#e0e0e0"
            button_fg = "#212121"
            highlight_bg = "#4a6da7"
            highlight_fg = "#ffffff"
            
        # Configure ttk styles
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, foreground=fg_color)
        self.style.configure("TButton", background=button_bg, foreground=button_fg)
        self.style.map("TButton",
                      background=[("active", highlight_bg)],
                      foreground=[("active", highlight_fg)])
                      
        self.style.configure("TNotebook", background=bg_color)
        self.style.configure("TNotebook.Tab", background=button_bg, foreground=button_fg)
        self.style.map("TNotebook.Tab",
                      background=[("selected", highlight_bg)],
                      foreground=[("selected", highlight_fg)])
                      
        # Configure root window
        self.root.configure(bg=bg_color)
        
    def create_menu_bar(self):
        """Create the application menu bar."""
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
        # File menu
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Configuration", command=self.open_config)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        self.menu_bar.add_cascade(label="File", menu=file_menu)
        
        # Data menu
        data_menu = tk.Menu(self.menu_bar, tearoff=0)
        data_menu.add_command(label="Backup Current Data", command=self.backup_current_data)
        data_menu.add_command(label="Export to Excel...", command=self.export_to_excel)
        
        # Quick Export submenu
        quick_export_menu = tk.Menu(data_menu, tearoff=0)
        quick_export_menu.add_command(label="Daily Sales Report", command=self.quick_export_daily)
        quick_export_menu.add_command(label="Monthly Summary Report", command=self.quick_export_monthly)
        quick_export_menu.add_command(label="Inventory Report", command=self.quick_export_inventory)
        quick_export_menu.add_command(label="Customer Credit Report", command=self.quick_export_credit)
        data_menu.add_cascade(label="Quick Export", menu=quick_export_menu)
        
        data_menu.add_separator()
        data_menu.add_command(label="Sync with Google Sheets", command=self.sync_with_sheets)
        self.menu_bar.add_cascade(label="Data", menu=data_menu)
        
        # Tools menu
        tools_menu = tk.Menu(self.menu_bar, tearoff=0)
        tools_menu.add_command(label="Products Manager", command=self.open_products_manager)
        tools_menu.add_command(label="Customers Manager", command=self.open_customers_manager)
        self.menu_bar.add_cascade(label="Tools", menu=tools_menu)
        
        # Help menu
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        self.menu_bar.add_cascade(label="Help", menu=help_menu)
        
    def create_status_bar(self):
        """Create status bar at the bottom of the window."""
        self.status_bar = ttk.Frame(self.root)
        self.status_bar.pack(side="bottom", fill="x")
        
        # Status text
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(self.status_bar, textvariable=self.status_var)
        self.status_label.pack(side="left", padx=10, pady=5)
        
        # Connection status
        self.connection_frame = ttk.Frame(self.status_bar)
        self.connection_frame.pack(side="right", padx=10, pady=5)
        
        self.connection_indicator = tk.Canvas(self.connection_frame, width=10, height=10, 
                                           bg=self.style.lookup("TFrame", "background"))
        self.connection_indicator.pack(side="left", padx=5)
        
        self.connection_var = tk.StringVar(value="Offline Mode")
        self.connection_label = ttk.Label(self.connection_frame, textvariable=self.connection_var)
        self.connection_label.pack(side="left")
        
        # Default to offline
        self.set_status("Ready", "offline")
        
    def set_status(self, message, connection_status=None):
        """Set status bar message and optionally connection status."""
        self.status_var.set(message)
        
        if connection_status:
            if connection_status == "online":
                self.connection_var.set("Online Mode")
                self.connection_indicator.create_oval(2, 2, 8, 8, fill="#4caf50", outline="")
            else:
                self.connection_var.set("Offline Mode")
                self.connection_indicator.create_oval(2, 2, 8, 8, fill="#ff9800", outline="")
                
    def create_main_content(self):
        """Create the main content area with tabs."""
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # Create tabs
        self.home_tab = ttk.Frame(self.notebook, padding=10)
        self.sales_tab = ttk.Frame(self.notebook, padding=10)
        self.purchases_tab = ttk.Frame(self.notebook, padding=10)
        self.reports_tab = ttk.Frame(self.notebook, padding=10)
        
        self.notebook.add(self.home_tab, text="Home")
        self.notebook.add(self.sales_tab, text="Sales Entry")
        self.notebook.add(self.purchases_tab, text="Purchase Entry")
        self.notebook.add(self.reports_tab, text="Reports")
        
        # Populate tabs
        self.create_home_tab()
        self.create_sales_tab()
        self.create_purchases_tab()
        self.create_reports_tab()
        
    def create_home_tab(self):
        """Create content for the home tab."""
        # Main container with card-like styling
        main_container = ttk.Frame(self.home_tab)
        main_container.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Header with company logo and name
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=(0, 20))
        
        company_name = self.config_manager.get("company_name", "Your Company")
        
        # Create a title with larger font
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side="top", fill="x")
        ttk.Label(title_frame, text=f"Welcome to {company_name}", 
                font=("Segoe UI", 22, "bold")).pack(pady=(0, 5))
        ttk.Label(title_frame, text="Point of Sale System", 
                font=("Segoe UI", 14)).pack(pady=(0, 10))
        
        # Date and time display
        current_date = datetime.datetime.now().strftime("%A, %d %B %Y")
        ttk.Label(title_frame, text=current_date, 
                font=("Segoe UI", 10)).pack(pady=(0, 5))
        
        # Create a two-column layout for content
        content_frame = ttk.Frame(main_container)
        content_frame.pack(fill="both", expand=True, pady=10)
        
        # Left column - System status and information
        left_column = ttk.Frame(content_frame)
        left_column.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # System Status Section
        status_frame = ttk.LabelFrame(left_column, text="System Status")
        status_frame.pack(fill="x", pady=(0, 15), padx=5)
        
        status_grid = ttk.Frame(status_frame, padding=10)
        status_grid.pack(fill="x")
        
        # Mode status with indicator
        ttk.Label(status_grid, text="Operation Mode:", 
                font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        mode_txt = "Online" if self.config_manager.get("app_mode") == "online" else "Offline"
        mode_color = "#4CAF50" if mode_txt == "Online" else "#FFA500"  # Green for online, orange for offline
        
        mode_frame = ttk.Frame(status_grid)
        mode_frame.grid(row=0, column=1, sticky="w", pady=5)
        
        # Create a colored indicator using Canvas
        mode_indicator = tk.Canvas(mode_frame, width=12, height=12, highlightthickness=0)
        mode_indicator.create_oval(2, 2, 10, 10, fill=mode_color, outline=mode_color)
        mode_indicator.pack(side="left", padx=(0, 5))
        
        ttk.Label(mode_frame, text=mode_txt, 
                font=("Segoe UI", 10)).pack(side="left")
        
        # Google Sheets status
        ttk.Label(status_grid, text="Google Sheets:", 
                font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", padx=10, pady=5)
                
        if hasattr(self, 'sheets_manager') and self.sheets_manager and self.sheets_manager.client:
            sheets_status = "Connected"
            sheets_color = "#4CAF50"  # Green for connected
        else:
            sheets_status = "Disconnected"
            sheets_color = "#F44336"  # Red for disconnected
            
        sheets_frame = ttk.Frame(status_grid)
        sheets_frame.grid(row=1, column=1, sticky="w", pady=5)
        
        sheets_indicator = tk.Canvas(sheets_frame, width=12, height=12, highlightthickness=0)
        sheets_indicator.create_oval(2, 2, 10, 10, fill=sheets_color, outline=sheets_color)
        sheets_indicator.pack(side="left", padx=(0, 5))
        
        ttk.Label(sheets_frame, text=sheets_status, 
                font=("Segoe UI", 10)).pack(side="left")
        
        # Backup status
        ttk.Label(status_grid, text="Last Backup:", 
                font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", padx=10, pady=5)
                
        last_backup = self.config_manager.get("last_backup", "Never")
        ttk.Label(status_grid, text=last_backup, 
                font=("Segoe UI", 10)).grid(row=2, column=1, sticky="w", pady=5)
        
        # Database info
        ttk.Label(status_grid, text="Database:", 
                font=("Segoe UI", 9, "bold")).grid(row=3, column=0, sticky="w", padx=10, pady=5)
                
        db_file = self.config_manager.get("db_file", "pos_data.db")
        ttk.Label(status_grid, text=f"{db_file}", 
                font=("Segoe UI", 10)).grid(row=3, column=1, sticky="w", pady=5)
        
        # Recent Activity Section (placeholder)
        activity_frame = ttk.LabelFrame(left_column, text="Recent Activity")
        activity_frame.pack(fill="both", expand=True, pady=(0, 10), padx=5)
        
        activity_list = ttk.Frame(activity_frame, padding=10)
        activity_list.pack(fill="both", expand=True)
        
        # Create a small treeview to show recent transactions
        columns = ("time", "type", "customer", "amount")
        self.activity_tree = ttk.Treeview(activity_list, columns=columns, 
                                     show="headings", selectmode="browse", height=5)
        
        self.activity_tree.heading("time", text="Time")
        self.activity_tree.heading("type", text="Type")
        self.activity_tree.heading("customer", text="Customer")
        self.activity_tree.heading("amount", text="Amount")
        
        self.activity_tree.column("time", width=80)
        self.activity_tree.column("type", width=80)
        self.activity_tree.column("customer", width=150)
        self.activity_tree.column("amount", width=80)
        
        self.activity_tree.pack(side="left", fill="both", expand=True)
        
        activity_scrollbar = ttk.Scrollbar(activity_list, orient="vertical", command=self.activity_tree.yview)
        activity_scrollbar.pack(side="right", fill="y")
        self.activity_tree.configure(yscrollcommand=activity_scrollbar.set)
        
        # Populate with dummy data (future enhancement: show real data)
        current_time = datetime.datetime.now()
        time_str = current_time.strftime("%H:%M")
        self.activity_tree.insert("", "end", values=(time_str, "Login", "System", "N/A"))
        
        # Right column - Quick actions and navigation
        right_column = ttk.Frame(content_frame)
        right_column.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # Quick Actions Section
        actions_frame = ttk.LabelFrame(right_column, text="Quick Actions")
        actions_frame.pack(fill="x", pady=(0, 15), padx=5)
        
        # Create a grid layout for actions
        action_grid = ttk.Frame(actions_frame, padding=10)
        action_grid.pack(fill="x")
        
        # Style for action buttons
        style = ttk.Style()
        style.configure("Action.TButton", font=("Segoe UI", 11))
        
        # Transaction buttons
        sales_btn = ttk.Button(action_grid, text="New Sale", 
                             command=lambda: self.notebook.select(self.sales_tab),
                             style="Action.TButton", width=15)
        sales_btn.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        purchase_btn = ttk.Button(action_grid, text="New Purchase", 
                                command=lambda: self.notebook.select(self.purchases_tab),
                                style="Action.TButton", width=15)
        purchase_btn.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # Report buttons
        report_btn = ttk.Button(action_grid, text="Today's Report", 
                              command=self.view_today_report,
                              style="Action.TButton", width=15)
        report_btn.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        
        export_btn = ttk.Button(action_grid, text="Export to Excel", 
                              command=lambda: self.export_to_excel(),
                              style="Action.TButton", width=15)
        export_btn.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        # Management buttons
        backup_btn = ttk.Button(action_grid, text="Backup Now", 
                              command=self.backup_current_data,
                              style="Action.TButton", width=15)
        backup_btn.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        
        config_btn = ttk.Button(action_grid, text="Configuration", 
                              command=self.open_config,
                              style="Action.TButton", width=15)
        config_btn.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        
        # Tips & Help Section
        help_frame = ttk.LabelFrame(right_column, text="Tips & Help")
        help_frame.pack(fill="both", expand=True, pady=(0, 10), padx=5)
        
        help_text = ttk.Frame(help_frame, padding=10)
        help_text.pack(fill="both", expand=True)
        
        tips = [
            "Press F1 at any time to open the help window",
            "Daily backups are created automatically",
            "Use the 'Export to Excel' feature to create detailed reports",
            "You can switch between Online and Offline modes in Configuration",
            "Always check customer credit before completing a sale"
        ]
        
        for i, tip in enumerate(tips):
            ttk.Label(help_text, text=f"• {tip}", 
                    wraplength=350, justify="left").pack(anchor="w", pady=2)
                 
    def create_sales_tab(self):
        """Create content for the sales entry tab."""
        # Main container with card-like styling
        main_container = ttk.Frame(self.sales_tab)
        main_container.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Title and header
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(header_frame, text="Sales Entry", font=("Segoe UI", 16, "bold")).pack(side="left")
        
        # Current date indicator on the right side
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        ttk.Label(header_frame, text=f"Today: {today}", font=("Segoe UI", 10)).pack(side="right")
        
        # Top section - Customer and Date
        top_section = ttk.LabelFrame(main_container, text="Transaction Details")
        top_section.pack(fill="x", pady=(0, 15), padx=5)
        
        details_frame = ttk.Frame(top_section)
        details_frame.pack(fill="x", padx=10, pady=10)
        
        # Two-column layout for details
        ttk.Label(details_frame, text="Date:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        today = datetime.datetime.now()
        self.date_entry = DateEntry(details_frame, width=15, 
                                  year=today.year, month=today.month, day=today.day,
                                  background='darkblue', foreground='white', borderwidth=2)
        self.date_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(details_frame, text="Customer:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.customer_var = tk.StringVar()
        self.customer_combo = ttk.Combobox(details_frame, textvariable=self.customer_var, width=25)
        self.customer_combo.grid(row=0, column=3, sticky="w", padx=5, pady=5)
        
        # Previous credit display
        ttk.Label(details_frame, text="Previous Credit:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.prev_credit_var = tk.StringVar(value="0.00")
        prev_credit_label = ttk.Label(details_frame, textvariable=self.prev_credit_var, width=15)
        prev_credit_label.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        # Notes field
        ttk.Label(details_frame, text="Notes:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.notes_var = tk.StringVar()
        notes_entry = ttk.Entry(details_frame, textvariable=self.notes_var, width=25)
        notes_entry.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        
        # Populate customer list
        self.update_customer_list()
        
        # Bind customer selection to credit update
        self.customer_combo.bind("<<ComboboxSelected>>", self.update_customer_credit)
        
        # Items section with inline add
        items_section = ttk.LabelFrame(main_container, text="Sale Items")
        items_section.pack(fill="both", expand=True, pady=(0, 15), padx=5)
        
        # Quick add item row
        add_item_frame = ttk.Frame(items_section)
        add_item_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        # Get products from database
        products = self.db_manager.get_all_products()
        product_names = [row['name'] for row in products]
        product_dict = {row['name']: row for row in products}
        
        ttk.Label(add_item_frame, text="Product:").grid(row=0, column=0, sticky="w", padx=5)
        self.new_product_var = tk.StringVar()
        product_combo = ttk.Combobox(add_item_frame, textvariable=self.new_product_var, width=30)
        product_combo['values'] = sorted(product_names)
        product_combo.grid(row=0, column=1, sticky="w", padx=5)
        
        ttk.Label(add_item_frame, text="Price:").grid(row=0, column=2, sticky="w", padx=5)
        self.new_price_var = tk.StringVar(value="0.00")
        price_entry = ttk.Entry(add_item_frame, textvariable=self.new_price_var, width=10)
        price_entry.grid(row=0, column=3, sticky="w", padx=5)
        
        ttk.Label(add_item_frame, text="Quantity:").grid(row=0, column=4, sticky="w", padx=5)
        self.new_quantity_var = tk.StringVar(value="1")
        quantity_entry = ttk.Entry(add_item_frame, textvariable=self.new_quantity_var, width=8)
        quantity_entry.grid(row=0, column=5, sticky="w", padx=5)
        
        add_button = ttk.Button(add_item_frame, text="Add Item", 
                             command=self.add_item_inline)
        add_button.grid(row=0, column=6, sticky="w", padx=10)
        
        # Update price when product selected
        def on_product_selected(event):
            product_name = self.new_product_var.get()
            if product_name in product_dict:
                self.new_price_var.set(f"{product_dict[product_name]['price']:.2f}")
                
        product_combo.bind("<<ComboboxSelected>>", on_product_selected)
        
        # Items table
        items_frame = ttk.Frame(items_section)
        items_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Items treeview with styled headers
        self.items_tree = ttk.Treeview(items_frame, columns=("product", "quantity", "price", "total"), 
                                    show="headings", selectmode="browse", height=8)
        self.items_tree.heading("product", text="Product")
        self.items_tree.heading("quantity", text="Quantity")
        self.items_tree.heading("price", text="Price")
        self.items_tree.heading("total", text="Total")
        
        self.items_tree.column("product", width=250)
        self.items_tree.column("quantity", width=100)
        self.items_tree.column("price", width=100)
        self.items_tree.column("total", width=100)
        
        self.items_tree.pack(side="left", fill="both", expand=True)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(items_frame, orient="vertical", command=self.items_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.items_tree.configure(yscrollcommand=scrollbar.set)
        
        # Action buttons row
        action_frame = ttk.Frame(items_section)
        action_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(action_frame, text="Remove Selected", command=self.remove_item).pack(side="left", padx=(0,5))
        ttk.Button(action_frame, text="Clear All", command=self.clear_items).pack(side="left")
        
        # Totals and payment section
        payment_section = ttk.LabelFrame(main_container, text="Payment Details")
        payment_section.pack(fill="x", pady=(0, 10), padx=5)
        
        payment_frame = ttk.Frame(payment_section)
        payment_frame.pack(fill="x", padx=10, pady=10)
        
        # Left column - totals
        totals_frame = ttk.Frame(payment_frame)
        totals_frame.pack(side="left", fill="x", expand=True)
        
        ttk.Label(totals_frame, text="Total Amount:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.total_amount_var = tk.StringVar(value="0.00")
        ttk.Label(totals_frame, textvariable=self.total_amount_var, font=("Segoe UI", 12, "bold")).grid(row=0, column=1, sticky="w", pady=5)
        
        ttk.Label(totals_frame, text="Cash Received:").grid(row=1, column=0, sticky="w", pady=5)
        self.cash_received_var = tk.StringVar(value="0.00")
        ttk.Entry(totals_frame, textvariable=self.cash_received_var, width=15).grid(row=1, column=1, sticky="w", pady=5)
        
        ttk.Label(totals_frame, text="Updated Credit:", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=5)
        self.updated_credit_var = tk.StringVar(value="0.00")
        ttk.Label(totals_frame, textvariable=self.updated_credit_var, font=("Segoe UI", 12)).grid(row=2, column=1, sticky="w", pady=5)
        
        # Right column - save buttons
        save_frame = ttk.Frame(payment_frame)
        save_frame.pack(side="right", padx=20)
        
        # Create button styles
        style = ttk.Style()
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Secondary.TButton", font=("Segoe UI", 9))
        
        # Add buttons with appropriate styles
        ttk.Button(save_frame, text="Calculate Totals", command=self.calculate_totals, 
                style="Secondary.TButton").pack(pady=5)
        
        # Add a clear/prominent save button
        save_btn_frame = ttk.Frame(save_frame)
        save_btn_frame.pack(pady=5)
        
        save_btn = ttk.Button(save_btn_frame, text="Save Transaction", command=self.save_transaction, 
                           style="Primary.TButton", width=20)
        save_btn.pack(fill="both", expand=True)
        
        # Bind events
        self.cash_received_var.trace_add("write", self.on_cash_received_change)
        
    def add_item_inline(self):
        """Add item directly from the inline form"""
        product_name = self.new_product_var.get()
        if not product_name:
            messagebox.showwarning("Input Error", "Please select a product")
            return
            
        try:
            quantity = float(self.new_quantity_var.get())
            price = float(self.new_price_var.get())
            
            if quantity <= 0:
                messagebox.showwarning("Input Error", "Quantity must be greater than zero")
                return
                
            total = quantity * price
            
            # Add to treeview
            self.items_tree.insert("", "end", values=(product_name, quantity, f"{price:.2f}", f"{total:.2f}"))
            
            # Clear the product selection for the next item
            self.new_product_var.set("")
            self.new_price_var.set("0.00")
            self.new_quantity_var.set("1")
            
            # Update total
            self.update_sales_total()
            
        except ValueError:
            messagebox.showwarning("Input Error", "Please enter valid numbers for quantity and price")
        
    def create_purchases_tab(self):
        """Create content for the purchases entry tab."""
        # Main container with card-like styling
        main_container = ttk.Frame(self.purchases_tab)
        main_container.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Title and header
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(header_frame, text="Purchase Entry", font=("Segoe UI", 16, "bold")).pack(side="left")
        
        # Current date indicator on the right side
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        ttk.Label(header_frame, text=f"Today: {today}", font=("Segoe UI", 10)).pack(side="right")
        
        # Top section - Supplier and Date
        top_section = ttk.LabelFrame(main_container, text="Transaction Details")
        top_section.pack(fill="x", pady=(0, 15), padx=5)
        
        details_frame = ttk.Frame(top_section)
        details_frame.pack(fill="x", padx=10, pady=10)
        
        # Two-column layout for details
        ttk.Label(details_frame, text="Date:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        today = datetime.datetime.now()
        self.purchase_date_entry = DateEntry(details_frame, width=15, 
                                          year=today.year, month=today.month, day=today.day,
                                          background='darkblue', foreground='white', borderwidth=2)
        self.purchase_date_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(details_frame, text="Supplier:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.supplier_var = tk.StringVar()
        self.supplier_combo = ttk.Combobox(details_frame, textvariable=self.supplier_var, width=25)
        self.supplier_combo.grid(row=0, column=3, sticky="w", padx=5, pady=5)
        
        # Previous credit display
        ttk.Label(details_frame, text="Previous Credit:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.supplier_prev_credit_var = tk.StringVar(value="0.00")
        prev_credit_label = ttk.Label(details_frame, textvariable=self.supplier_prev_credit_var, width=15)
        prev_credit_label.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        
        # Notes field
        ttk.Label(details_frame, text="Notes:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
        self.purchase_notes_var = tk.StringVar()
        notes_entry = ttk.Entry(details_frame, textvariable=self.purchase_notes_var, width=25)
        notes_entry.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        
        # Populate supplier list
        self.update_supplier_list()
        
        # Bind supplier selection to credit update
        self.supplier_combo.bind("<<ComboboxSelected>>", self.update_supplier_credit)
        
        # Items section with inline add
        items_section = ttk.LabelFrame(main_container, text="Purchase Items")
        items_section.pack(fill="both", expand=True, pady=(0, 15), padx=5)
        
        # Quick add item row
        add_item_frame = ttk.Frame(items_section)
        add_item_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        # Get products from database
        products = self.db_manager.get_all_products()
        product_names = [row['name'] for row in products]
        product_dict = {row['name']: row for row in products}
        
        ttk.Label(add_item_frame, text="Product:").grid(row=0, column=0, sticky="w", padx=5)
        self.new_purchase_product_var = tk.StringVar()
        product_combo = ttk.Combobox(add_item_frame, textvariable=self.new_purchase_product_var, width=30)
        product_combo['values'] = sorted(product_names)
        product_combo.grid(row=0, column=1, sticky="w", padx=5)
        
        ttk.Label(add_item_frame, text="Price:").grid(row=0, column=2, sticky="w", padx=5)
        self.new_purchase_price_var = tk.StringVar(value="0.00")
        price_entry = ttk.Entry(add_item_frame, textvariable=self.new_purchase_price_var, width=10)
        price_entry.grid(row=0, column=3, sticky="w", padx=5)
        
        ttk.Label(add_item_frame, text="Quantity:").grid(row=0, column=4, sticky="w", padx=5)
        self.new_purchase_quantity_var = tk.StringVar(value="1")
        quantity_entry = ttk.Entry(add_item_frame, textvariable=self.new_purchase_quantity_var, width=8)
        quantity_entry.grid(row=0, column=5, sticky="w", padx=5)
        
        add_button = ttk.Button(add_item_frame, text="Add Item", 
                             command=self.add_purchase_item_inline)
        add_button.grid(row=0, column=6, sticky="w", padx=10)
        
        # Update price when product selected
        def on_product_selected(event):
            product_name = self.new_purchase_product_var.get()
            if product_name in product_dict:
                self.new_purchase_price_var.set(f"{product_dict[product_name]['price']:.2f}")
                
        product_combo.bind("<<ComboboxSelected>>", on_product_selected)
        
        # Items table
        items_frame = ttk.Frame(items_section)
        items_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Items treeview with styled headers
        self.purchase_items_tree = ttk.Treeview(items_frame, columns=("product", "quantity", "price", "total"), 
                                             show="headings", selectmode="browse", height=8)
        self.purchase_items_tree.heading("product", text="Product")
        self.purchase_items_tree.heading("quantity", text="Quantity")
        self.purchase_items_tree.heading("price", text="Price")
        self.purchase_items_tree.heading("total", text="Total")
        
        self.purchase_items_tree.column("product", width=250)
        self.purchase_items_tree.column("quantity", width=100)
        self.purchase_items_tree.column("price", width=100)
        self.purchase_items_tree.column("total", width=100)
        
        self.purchase_items_tree.pack(side="left", fill="both", expand=True)
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(items_frame, orient="vertical", command=self.purchase_items_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.purchase_items_tree.configure(yscrollcommand=scrollbar.set)
        
        # Action buttons row
        action_frame = ttk.Frame(items_section)
        action_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(action_frame, text="Remove Selected", command=self.remove_purchase_item).pack(side="left", padx=(0,5))
        ttk.Button(action_frame, text="Clear All", command=self.clear_purchase_items).pack(side="left")
        
        # Totals and payment section
        payment_section = ttk.LabelFrame(main_container, text="Payment Details")
        payment_section.pack(fill="x", pady=(0, 10), padx=5)
        
        payment_frame = ttk.Frame(payment_section)
        payment_frame.pack(fill="x", padx=10, pady=10)
        
        # Left column - totals
        totals_frame = ttk.Frame(payment_frame)
        totals_frame.pack(side="left", fill="x", expand=True)
        
        ttk.Label(totals_frame, text="Total Amount:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.purchase_total_amount_var = tk.StringVar(value="0.00")
        ttk.Label(totals_frame, textvariable=self.purchase_total_amount_var, font=("Segoe UI", 12, "bold")).grid(row=0, column=1, sticky="w", pady=5)
        
        ttk.Label(totals_frame, text="Cash Paid:").grid(row=1, column=0, sticky="w", pady=5)
        self.purchase_cash_paid_var = tk.StringVar(value="0.00")
        ttk.Entry(totals_frame, textvariable=self.purchase_cash_paid_var, width=15).grid(row=1, column=1, sticky="w", pady=5)
        
        ttk.Label(totals_frame, text="Updated Credit:", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=5)
        self.purchase_updated_credit_var = tk.StringVar(value="0.00")
        ttk.Label(totals_frame, textvariable=self.purchase_updated_credit_var, font=("Segoe UI", 12)).grid(row=2, column=1, sticky="w", pady=5)
        
        # Right column - save buttons
        save_frame = ttk.Frame(payment_frame)
        save_frame.pack(side="right", padx=20)
        
        # Create button styles (reusing the ones from sales tab)
        style = ttk.Style()
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Secondary.TButton", font=("Segoe UI", 9))
        
        ttk.Button(save_frame, text="Calculate Totals", command=self.calculate_purchase_totals, 
                style="Secondary.TButton").pack(pady=5)
        
        # Add a clear/prominent save button
        save_btn_frame = ttk.Frame(save_frame)
        save_btn_frame.pack(pady=5)
        
        save_btn = ttk.Button(save_btn_frame, text="Save Purchase", command=self.save_purchase_transaction, 
                           style="Primary.TButton", width=20)
        save_btn.pack(fill="both", expand=True)
        
        # Bind events
        self.purchase_cash_paid_var.trace_add("write", self.on_purchase_cash_paid_change)
        
    def add_purchase_item_inline(self):
        """Add purchase item directly from the inline form"""
        product_name = self.new_purchase_product_var.get()
        if not product_name:
            messagebox.showwarning("Input Error", "Please select a product")
            return
            
        try:
            quantity = float(self.new_purchase_quantity_var.get())
            price = float(self.new_purchase_price_var.get())
            
            if quantity <= 0:
                messagebox.showwarning("Input Error", "Quantity must be greater than zero")
                return
                
            total = quantity * price
            
            # Add to treeview
            self.purchase_items_tree.insert("", "end", values=(product_name, quantity, f"{price:.2f}", f"{total:.2f}"))
            
            # Clear the product selection for the next item
            self.new_purchase_product_var.set("")
            self.new_purchase_price_var.set("0.00")
            self.new_purchase_quantity_var.set("1")
            
            # Update total
            self.update_purchase_total()
            
        except ValueError:
            messagebox.showwarning("Input Error", "Please enter valid numbers for quantity and price")
        
    def create_reports_tab(self):
        """Create content for the reports tab."""
        # Main container
        main_container = ttk.Frame(self.reports_tab)
        main_container.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Title and header
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(header_frame, text="Reports & Analysis", font=("Segoe UI", 16, "bold")).pack(side="left")
        
        # Date range selection in a card
        date_section = ttk.LabelFrame(main_container, text="Date Range")
        date_section.pack(fill="x", pady=(0, 15), padx=5)
        
        date_frame = ttk.Frame(date_section)
        date_frame.pack(fill="x", padx=10, pady=10)
        
        # Create a grid layout for date selection
        ttk.Label(date_frame, text="From:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        today = datetime.datetime.now()
        first_of_month = today.replace(day=1)
        self.from_date_entry = DateEntry(date_frame, width=15, 
                                       year=first_of_month.year, 
                                       month=first_of_month.month, 
                                       day=first_of_month.day,
                                       background='darkblue', foreground='white', borderwidth=2)
        self.from_date_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        
        ttk.Label(date_frame, text="To:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
        self.to_date_entry = DateEntry(date_frame, width=15, 
                                     year=today.year, month=today.month, day=today.day,
                                     background='darkblue', foreground='white', borderwidth=2)
        self.to_date_entry.grid(row=0, column=3, sticky="w", padx=5, pady=5)
        
        # Button frame with styled buttons
        button_frame = ttk.Frame(date_frame)
        button_frame.grid(row=0, column=4, padx=20, pady=5, sticky="e")
        
        # Create button styles (reuse existing ones)
        style = ttk.Style()
        style.configure("Action.TButton", font=("Segoe UI", 9))
        
        ttk.Button(button_frame, text="Generate Report", command=self.generate_report, 
                style="Action.TButton").pack(side="left", padx=5)
        ttk.Button(button_frame, text="Export to Excel", command=self.export_report_to_excel, 
                style="Action.TButton").pack(side="left", padx=5)
        
        # Quick report options
        quick_frame = ttk.Frame(date_section)
        quick_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        ttk.Label(quick_frame, text="Quick Reports:", font=("Segoe UI", 9, "bold")).pack(side="left", padx=5)
        ttk.Button(quick_frame, text="Today", command=self.view_today_report, 
                style="Action.TButton").pack(side="left", padx=5)
        ttk.Button(quick_frame, text="This Week", command=lambda: self.quick_export_weekly(), 
                style="Action.TButton").pack(side="left", padx=5)
        ttk.Button(quick_frame, text="This Month", command=lambda: self.quick_export_monthly(), 
                style="Action.TButton").pack(side="left", padx=5)
        
        # Report content section
        report_section = ttk.LabelFrame(main_container, text="Report Results")
        report_section.pack(fill="both", expand=True, pady=(0, 15), padx=5)
        
        # Notebook for different report views with styling
        report_notebook = ttk.Notebook(report_section)
        report_notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Sales tab
        self.sales_report_tab = ttk.Frame(report_notebook, padding=10)
        report_notebook.add(self.sales_report_tab, text="Sales")
        
        # Products tab
        self.products_report_tab = ttk.Frame(report_notebook, padding=10)
        report_notebook.add(self.products_report_tab, text="Products")
        
        # Customers tab
        self.customers_report_tab = ttk.Frame(report_notebook, padding=10)
        report_notebook.add(self.customers_report_tab, text="Customers")
        
        # Create styled sales report treeview
        sales_frame = ttk.Frame(self.sales_report_tab)
        sales_frame.pack(fill="both", expand=True)
        
        self.sales_tree = ttk.Treeview(sales_frame, 
                                     columns=("date", "customer", "total", "cash", "credit"),
                                     show="headings", selectmode="browse", height=10)
        self.sales_tree.heading("date", text="Date")
        self.sales_tree.heading("customer", text="Customer")
        self.sales_tree.heading("total", text="Total")
        self.sales_tree.heading("cash", text="Cash Received")
        self.sales_tree.heading("credit", text="Credit Balance")
        
        self.sales_tree.column("date", width=100)
        self.sales_tree.column("customer", width=200)
        self.sales_tree.column("total", width=100)
        self.sales_tree.column("cash", width=100)
        self.sales_tree.column("credit", width=100)
        
        self.sales_tree.pack(side="left", fill="both", expand=True)
        
        sales_scrollbar = ttk.Scrollbar(sales_frame, orient="vertical", command=self.sales_tree.yview)
        sales_scrollbar.pack(side="right", fill="y")
        self.sales_tree.configure(yscrollcommand=sales_scrollbar.set)
        
        # Create summary frame with a card-like appearance
        summary_frame = ttk.LabelFrame(self.sales_report_tab, text="Summary")
        summary_frame.pack(fill="x", pady=10)
        
        # Use a grid layout for better organization
        summary_grid = ttk.Frame(summary_frame, padding=10)
        summary_grid.pack(fill="x")
        
        ttk.Label(summary_grid, text="Total Sales:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.total_sales_var = tk.StringVar(value="0.00")
        ttk.Label(summary_grid, textvariable=self.total_sales_var, font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w", pady=5)
        
        ttk.Label(summary_grid, text="Total Cash Received:", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, sticky="w", padx=10, pady=5)
        self.total_cash_var = tk.StringVar(value="0.00")
        ttk.Label(summary_grid, textvariable=self.total_cash_var, font=("Segoe UI", 10, "bold")).grid(row=0, column=3, sticky="w", pady=5)
        
        ttk.Label(summary_grid, text="Transaction Count:", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.transaction_count_var = tk.StringVar(value="0")
        ttk.Label(summary_grid, textvariable=self.transaction_count_var, font=("Segoe UI", 10)).grid(row=1, column=1, sticky="w", pady=5)
        
        # Create products report treeview
        products_frame = ttk.Frame(self.products_report_tab)
        products_frame.pack(fill="both", expand=True)
        
        self.products_tree = ttk.Treeview(products_frame, 
                                        columns=("product", "quantity", "revenue"),
                                        show="headings", selectmode="browse")
        self.products_tree.heading("product", text="Product")
        self.products_tree.heading("quantity", text="Quantity Sold")
        self.products_tree.heading("revenue", text="Revenue")
        
        self.products_tree.column("product", width=250)
        self.products_tree.column("quantity", width=150)
        self.products_tree.column("revenue", width=150)
        
        self.products_tree.pack(side="left", fill="both", expand=True)
        
        products_scrollbar = ttk.Scrollbar(products_frame, orient="vertical", command=self.products_tree.yview)
        products_scrollbar.pack(side="right", fill="y")
        self.products_tree.configure(yscrollcommand=products_scrollbar.set)
        
        # Create customers report treeview
        customers_frame = ttk.Frame(self.customers_report_tab)
        customers_frame.pack(fill="both", expand=True)
        
        self.customers_tree = ttk.Treeview(customers_frame, 
                                         columns=("customer", "transactions", "total", "credit"),
                                         show="headings", selectmode="browse")
        self.customers_tree.heading("customer", text="Customer")
        self.customers_tree.heading("transactions", text="Transactions")
        self.customers_tree.heading("total", text="Total Purchased")
        self.customers_tree.heading("credit", text="Current Credit")
        
        self.customers_tree.column("customer", width=250)
        self.customers_tree.column("transactions", width=100)
        self.customers_tree.column("total", width=150)
        self.customers_tree.column("credit", width=150)
        
        self.customers_tree.pack(side="left", fill="both", expand=True)
        
        customers_scrollbar = ttk.Scrollbar(customers_frame, orient="vertical", command=self.customers_tree.yview)
        customers_scrollbar.pack(side="right", fill="y")
        self.customers_tree.configure(yscrollcommand=customers_scrollbar.set)
        
    def update_customer_list(self):
        """Update the customer dropdown list."""
        customers = self.db_manager.get_all_customers()
        customer_names = [row['name'] for row in customers]
        self.customer_combo['values'] = sorted(customer_names)
        
    def update_supplier_list(self):
        """Update the supplier dropdown list."""
        suppliers = self.db_manager.get_all_suppliers()
        supplier_names = [row['name'] for row in suppliers]
        self.supplier_combo['values'] = sorted(supplier_names)
        
    def update_customer_credit(self, event=None):
        """Update the displayed customer credit when customer is selected."""
        customer_name = self.customer_var.get()
        if not customer_name:
            self.prev_credit_var.set("0.00")
            return
            
        # Find customer in database
        conn = self.db_manager.connect()
        cursor = conn.cursor()
        cursor.execute("SELECT id, credit_balance FROM customers WHERE name = ?", (customer_name,))
        result = cursor.fetchone()
        
        if result:
            credit_balance = result['credit_balance']
            self.prev_credit_var.set(f"{credit_balance:.2f}")
        else:
            self.prev_credit_var.set("0.00")
            
    def update_supplier_credit(self, event=None):
        """Update the displayed supplier credit when supplier is selected."""
        supplier_name = self.supplier_var.get()
        if not supplier_name:
            self.supplier_prev_credit_var.set("0.00")
            return
            
        # Find supplier in database
        conn = self.db_manager.connect()
        cursor = conn.cursor()
        cursor.execute("SELECT id, credit_balance FROM customers WHERE name = ? AND is_supplier = 1", (supplier_name,))
        result = cursor.fetchone()
        
        if result:
            credit_balance = result['credit_balance']
            self.supplier_prev_credit_var.set(f"{credit_balance:.2f}")
        else:
            self.supplier_prev_credit_var.set("0.00")
            
    def add_item_dialog(self):
        """Show dialog to add an item to the sale."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Item")
        dialog.geometry("400x300")
        dialog.grab_set()  # Make the dialog modal
        
        ttk.Label(dialog, text="Select Product:").pack(pady=(20, 5))
        
        # Get products from database
        products = self.db_manager.get_all_products()
        product_names = [row['name'] for row in products]
        product_dict = {row['name']: row for row in products}
        
        product_var = tk.StringVar()
        product_combo = ttk.Combobox(dialog, textvariable=product_var, width=30)
        product_combo['values'] = sorted(product_names)
        product_combo.pack(pady=5)
        
        # Price field
        price_frame = ttk.Frame(dialog)
        price_frame.pack(pady=5)
        ttk.Label(price_frame, text="Price:").pack(side="left", padx=5)
        price_var = tk.StringVar(value="0.00")
        price_entry = ttk.Entry(price_frame, textvariable=price_var, width=10)
        price_entry.pack(side="left")
        
        # Quantity field
        quantity_frame = ttk.Frame(dialog)
        quantity_frame.pack(pady=5)
        ttk.Label(quantity_frame, text="Quantity:").pack(side="left", padx=5)
        quantity_var = tk.StringVar(value="1")
        quantity_entry = ttk.Entry(quantity_frame, textvariable=quantity_var, width=10)
        quantity_entry.pack(side="left")
        
        # Update price when product selected
        def on_product_selected(event):
            product_name = product_var.get()
            if product_name in product_dict:
                price_var.set(f"{product_dict[product_name]['price']:.2f}")
                
        product_combo.bind("<<ComboboxSelected>>", on_product_selected)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def add_item():
            product_name = product_var.get()
            if not product_name:
                messagebox.showwarning("Input Error", "Please select a product")
                return
                
            try:
                quantity = float(quantity_var.get())
                price = float(price_var.get())
                
                if quantity <= 0:
                    messagebox.showwarning("Input Error", "Quantity must be greater than zero")
                    return
                    
                total = quantity * price
                
                # Add to treeview
                self.items_tree.insert("", "end", values=(product_name, quantity, f"{price:.2f}", f"{total:.2f}"))
                
                # Update total
                self.update_sales_total()
                
                # Close dialog
                dialog.destroy()
            except ValueError:
                messagebox.showwarning("Input Error", "Please enter valid numbers for quantity and price")
                
        ttk.Button(button_frame, text="Add", command=add_item).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=10)
        
    def remove_item(self):
        """Remove selected item from the sale."""
        selected_item = self.items_tree.selection()
        if selected_item:
            self.items_tree.delete(selected_item)
            self.update_sales_total()
            
    def clear_items(self):
        """Clear all items from the sale."""
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)
        self.update_sales_total()
        
    def update_sales_total(self):
        """Update the sales total based on items in the treeview."""
        total = 0.0
        for item in self.items_tree.get_children():
            values = self.items_tree.item(item, 'values')
            total += float(values[3])  # Total column
            
        self.total_amount_var.set(f"{total:.2f}")
        self.calculate_totals()
        
    def on_cash_received_change(self, *args):
        """Recalculate totals when cash received changes."""
        try:
            # Get current values
            total_amount = safe_float(self.total_amount_var.get(), 0.0)
            cash_received = safe_float(self.cash_received_var.get(), 0.0)
            previous_credit = safe_float(self.prev_credit_var.get(), 0.0)
            
            # Calculate updated credit
            updated_credit = previous_credit + (total_amount - cash_received)
            
            # Update the credit display
            self.updated_credit_var.set(f"{updated_credit:.2f}")
        except Exception as e:
            print(f"Error in cash received handler: {e}")
            traceback.print_exc()
        
    def add_purchase_item_dialog(self):
        """Show dialog to add an item to the purchase."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Purchase Item")
        dialog.geometry("400x300")
        dialog.grab_set()  # Make the dialog modal
        
        ttk.Label(dialog, text="Select Product:").pack(pady=(20, 5))
        
        # Get products from database
        products = self.db_manager.get_all_products()
        product_names = [row['name'] for row in products]
        product_dict = {row['name']: row for row in products}
        
        product_var = tk.StringVar()
        product_combo = ttk.Combobox(dialog, textvariable=product_var, width=30)
        product_combo['values'] = sorted(product_names)
        product_combo.pack(pady=5)
        
        # Price field
        price_frame = ttk.Frame(dialog)
        price_frame.pack(pady=5)
        ttk.Label(price_frame, text="Price:").pack(side="left", padx=5)
        price_var = tk.StringVar(value="0.00")
        price_entry = ttk.Entry(price_frame, textvariable=price_var, width=10)
        price_entry.pack(side="left")
        
        # Quantity field
        quantity_frame = ttk.Frame(dialog)
        quantity_frame.pack(pady=5)
        ttk.Label(quantity_frame, text="Quantity:").pack(side="left", padx=5)
        quantity_var = tk.StringVar(value="1")
        quantity_entry = ttk.Entry(quantity_frame, textvariable=quantity_var, width=10)
        quantity_entry.pack(side="left")
        
        # Update price when product selected
        def on_product_selected(event):
            product_name = product_var.get()
            if product_name in product_dict:
                price_var.set(f"{product_dict[product_name]['price']:.2f}")
                
        product_combo.bind("<<ComboboxSelected>>", on_product_selected)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def add_item():
            product_name = product_var.get()
            if not product_name:
                messagebox.showwarning("Input Error", "Please select a product")
                return
                
            try:
                quantity = float(quantity_var.get())
                price = float(price_var.get())
                
                if quantity <= 0:
                    messagebox.showwarning("Input Error", "Quantity must be greater than zero")
                    return
                    
                total = quantity * price
                
                # Add to treeview
                self.purchase_items_tree.insert("", "end", values=(product_name, quantity, f"{price:.2f}", f"{total:.2f}"))
                
                # Update total
                self.update_purchase_total()
                
                # Close dialog
                dialog.destroy()
            except ValueError:
                messagebox.showwarning("Input Error", "Please enter valid numbers for quantity and price")
                
        ttk.Button(button_frame, text="Add", command=add_item).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=10)
    
    def remove_purchase_item(self):
        """Remove selected item from the purchase."""
        selected_item = self.purchase_items_tree.selection()
        if selected_item:
            self.purchase_items_tree.delete(selected_item)
            self.update_purchase_total()
            
    def clear_purchase_items(self):
        """Clear all items from the purchase."""
        for item in self.purchase_items_tree.get_children():
            self.purchase_items_tree.delete(item)
        self.update_purchase_total()
        
    def update_purchase_total(self):
        """Update the purchase total based on items in the treeview."""
        total = 0.0
        for item in self.purchase_items_tree.get_children():
            values = self.purchase_items_tree.item(item, 'values')
            total += float(values[3])  # Total column
            
        self.purchase_total_amount_var.set(f"{total:.2f}")
        self.calculate_purchase_totals()
        
    def on_purchase_cash_paid_change(self, *args):
        """Recalculate totals when cash paid changes."""
        try:
            # Get current values
            total_amount = float(self.purchase_total_amount_var.get())
            cash_paid = safe_float(self.purchase_cash_paid_var.get(), 0.0)
            previous_credit = safe_float(self.supplier_prev_credit_var.get(), 0.0)
            
            # For suppliers: updated credit = previous credit + total amount - cash paid
            # Note: this assumes that higher credit means the supplier owes more to us
            updated_credit = previous_credit + total_amount - cash_paid
            self.purchase_updated_credit_var.set(f"{updated_credit:.2f}")
        except Exception as e:
            print(f"Error in cash paid handler: {e}")
            traceback.print_exc()
        
    def calculate_purchase_totals(self):
        """Calculate updated credit based on purchase total and cash paid."""
        try:
            total_amount = float(self.purchase_total_amount_var.get())
            cash_paid = safe_float(self.purchase_cash_paid_var.get(), 0.0)
            previous_credit = safe_float(self.supplier_prev_credit_var.get(), 0.0)
            
            # For suppliers: updated credit = previous credit + total amount - cash paid
            # Note: this assumes that higher credit means the supplier owes more to us
            updated_credit = previous_credit + total_amount - cash_paid
            self.purchase_updated_credit_var.set(f"{updated_credit:.2f}")
        except ValueError:
            pass
            
    def save_purchase_transaction(self):
        """Save the current purchase transaction to the database."""
        # Validate inputs
        supplier_name = self.supplier_var.get()
        if not supplier_name:
            messagebox.showwarning("Input Error", "Please select a supplier")
            return
            
        items = self.purchase_items_tree.get_children()
        if not items:
            messagebox.showwarning("Input Error", "Please add at least one item")
            return
            
        try:
            # Get values
            date_str = self.purchase_date_entry.get_date().strftime("%Y-%m-%d")
            total_amount = float(self.purchase_total_amount_var.get())
            cash_paid = safe_float(self.purchase_cash_paid_var.get(), 0.0)
            previous_credit = safe_float(self.supplier_prev_credit_var.get(), 0.0)
            updated_credit = float(self.purchase_updated_credit_var.get())
            
            # Get supplier ID
            conn = self.db_manager.connect()
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM customers WHERE name = ? AND is_supplier = 1", (supplier_name,))
            result = cursor.fetchone()
            
            if not result:
                # Add new supplier
                self.db_manager.add_or_update_customer(supplier_name, "", previous_credit, 1)
                cursor.execute("SELECT id FROM customers WHERE name = ? AND is_supplier = 1", (supplier_name,))
                result = cursor.fetchone()
                
            supplier_id = result['id']
            
            # Prepare items list
            items_list = []
            for item in items:
                values = self.purchase_items_tree.item(item, 'values')
                product_name = values[0]
                quantity = float(values[1])
                price = float(values[2])
                
                # Get product ID
                cursor.execute("SELECT id FROM products WHERE name = ?", (product_name,))
                product_result = cursor.fetchone()
                
                if not product_result:
                    # Add new product
                    self.db_manager.add_or_update_product(product_name, price)
                    cursor.execute("SELECT id FROM products WHERE name = ?", (product_name,))
                    product_result = cursor.fetchone()
                    
                product_id = product_result['id']
                
                items_list.append({
                    'product_id': product_id,
                    'quantity': quantity,
                    'price': price
                })
                
            # Save transaction
            purchase_id = self.db_manager.add_purchase(
                date_str, supplier_id, items_list, total_amount, 
                cash_paid, previous_credit, updated_credit
            )
            
            if purchase_id:
                # Sync with Google Sheets if in online mode
                if (self.config_manager.get("app_mode") == "online" and 
                    hasattr(self, 'sheets_manager') and 
                    self.sheets_manager and 
                    self.sheets_manager.client):
                    
                    self.set_status("Syncing with Google Sheets...", "online")
                    purchases_data = self.db_manager.get_daily_purchases_detail(date_str)
                    header_row = ["Supplier Name", "Total Amount for", "Previous Credit", 
                                 "Cash Paid", "Updated Credit Balance", "Extra"]
                    self.sheets_manager.sync_purchases_to_sheet(date_str, purchases_data, header_row)
                    self.set_status("Purchase transaction saved and synced", "online")
                else:
                    self.set_status("Purchase transaction saved (offline mode)", "offline")
                    
                # Clear the form
                self.clear_purchase_items()
                self.supplier_var.set("")
                self.supplier_prev_credit_var.set("0.00")
                self.purchase_cash_paid_var.set("0.00")
                self.purchase_total_amount_var.set("0.00")
                self.purchase_updated_credit_var.set("0.00")
                
                messagebox.showinfo("Success", "Purchase transaction saved successfully")
            else:
                messagebox.showerror("Error", "Failed to save purchase transaction. Please check the logs.")
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
        
    def calculate_totals(self):
        """Calculate updated credit based on total and cash received."""
        try:
            total_amount = float(self.total_amount_var.get())
            cash_received = safe_float(self.cash_received_var.get(), 0.0)
            previous_credit = safe_float(self.prev_credit_var.get(), 0.0)
            
            # Updated credit = previous credit + total amount - cash received
            updated_credit = previous_credit + total_amount - cash_received
            self.updated_credit_var.set(f"{updated_credit:.2f}")
        except ValueError:
            pass
            
    def save_transaction(self):
        """Save the current transaction to the database."""
        # Validate inputs
        customer_name = self.customer_var.get()
        if not customer_name:
            messagebox.showwarning("Input Error", "Please select a customer")
            return
            
        items = self.items_tree.get_children()
        if not items:
            messagebox.showwarning("Input Error", "Please add at least one item")
            return
            
        try:
            # Get values
            date_str = self.date_entry.get_date().strftime("%Y-%m-%d")
            total_amount = float(self.total_amount_var.get())
            cash_received = safe_float(self.cash_received_var.get(), 0.0)
            previous_credit = safe_float(self.prev_credit_var.get(), 0.0)
            updated_credit = float(self.updated_credit_var.get())
            
            # Get customer ID
            conn = self.db_manager.connect()
            cursor = conn.cursor()
            cursor.execute("SELECT id FROM customers WHERE name = ?", (customer_name,))
            result = cursor.fetchone()
            
            if not result:
                # Add new customer
                self.db_manager.add_or_update_customer(customer_name, "", previous_credit)
                cursor.execute("SELECT id FROM customers WHERE name = ?", (customer_name,))
                result = cursor.fetchone()
                
            customer_id = result['id']
            
            # Prepare items list
            items_list = []
            for item in items:
                values = self.items_tree.item(item, 'values')
                product_name = values[0]
                quantity = float(values[1])
                price = float(values[2])
                
                # Get product ID
                cursor.execute("SELECT id FROM products WHERE name = ?", (product_name,))
                product_result = cursor.fetchone()
                
                if not product_result:
                    # Add new product
                    self.db_manager.add_or_update_product(product_name, price)
                    cursor.execute("SELECT id FROM products WHERE name = ?", (product_name,))
                    product_result = cursor.fetchone()
                    
                product_id = product_result['id']
                
                items_list.append({
                    'product_id': product_id,
                    'quantity': quantity,
                    'price': price
                })
                
            # Save transaction
            sale_id = self.db_manager.add_sale(
                date_str, customer_id, items_list, total_amount, 
                cash_received, previous_credit, updated_credit
            )
            
            if sale_id:
                # Sync with Google Sheets if in online mode
                if (self.config_manager.get("app_mode") == "online" and 
                    hasattr(self, 'sheets_manager') and 
                    self.sheets_manager and 
                    self.sheets_manager.client):
                    
                    self.set_status("Syncing with Google Sheets...", "online")
                    sales_data = self.db_manager.get_daily_sales_detail(date_str)
                    header_row = ["Customer Name", "Total Amount for", "Previous Credit", 
                                 "Cash Received", "Updated Credit Balance", "Extra"]
                    self.sheets_manager.sync_sales_to_sheet(date_str, sales_data, header_row)
                    self.set_status("Transaction saved and synced", "online")
                else:
                    self.set_status("Transaction saved (offline mode)", "offline")
                    
                # Clear the form
                self.clear_items()
                self.customer_var.set("")
                self.prev_credit_var.set("0.00")
                self.cash_received_var.set("0.00")
                self.total_amount_var.set("0.00")
                self.updated_credit_var.set("0.00")
                
                messagebox.showinfo("Success", "Transaction saved successfully")
            else:
                messagebox.showerror("Error", "Failed to save transaction. Please check the logs.")
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
            
    def generate_report(self):
        """Generate sales report for the selected date range."""
        try:
            from_date = self.from_date_entry.get_date().strftime("%Y-%m-%d")
            to_date = self.to_date_entry.get_date().strftime("%Y-%m-%d")
            
            # Clear existing data
            for tree in [self.sales_tree, self.products_tree, self.customers_tree]:
                for item in tree.get_children():
                    tree.delete(item)
                    
            # Connect to database
            conn = self.db_manager.connect()
            
            # Get sales data
            sales_df = pd.read_sql_query('''
            SELECT s.id, s.date, c.name as customer_name, s.total_amount, s.cash_received,
                   s.updated_credit
            FROM sales s
            LEFT JOIN customers c ON s.customer_id = c.id
            WHERE s.date BETWEEN ? AND ?
            ORDER BY s.date, s.id
            ''', conn, params=(from_date, to_date))
            
            # Populate sales tree
            total_sales = 0
            total_cash = 0
            
            for _, row in sales_df.iterrows():
                self.sales_tree.insert("", "end", values=(
                    row['date'],
                    row['customer_name'],
                    f"{row['total_amount']:.2f}",
                    f"{row['cash_received']:.2f}",
                    f"{row['updated_credit']:.2f}"
                ))
                
                total_sales += row['total_amount']
                total_cash += row['cash_received']
                
            # Update summary
            self.total_sales_var.set(f"{total_sales:.2f}")
            self.total_cash_var.set(f"{total_cash:.2f}")
            self.transaction_count_var.set(str(len(sales_df)))
            
            # Get product data
            products_df = pd.read_sql_query('''
            SELECT p.name as product_name, SUM(si.quantity) as quantity, 
                   SUM(si.quantity * si.price) as revenue
            FROM sale_items si
            JOIN products p ON si.product_id = p.id
            JOIN sales s ON si.sale_id = s.id
            WHERE s.date BETWEEN ? AND ?
            GROUP BY p.name
            ORDER BY revenue DESC
            ''', conn, params=(from_date, to_date))
            
            # Populate products tree
            for _, row in products_df.iterrows():
                self.products_tree.insert("", "end", values=(
                    row['product_name'],
                    f"{row['quantity']:.2f}",
                    f"{row['revenue']:.2f}"
                ))
                
            # Get customer data
            customers_df = pd.read_sql_query('''
            SELECT c.name as customer_name, COUNT(s.id) as transaction_count, 
                   SUM(s.total_amount) as total_purchased, c.credit_balance
            FROM customers c
            LEFT JOIN sales s ON c.id = s.customer_id AND s.date BETWEEN ? AND ?
            GROUP BY c.id
            HAVING transaction_count > 0
            ORDER BY total_purchased DESC
            ''', conn, params=(from_date, to_date))
            
            # Populate customers tree
            for _, row in customers_df.iterrows():
                self.customers_tree.insert("", "end", values=(
                    row['customer_name'],
                    row['transaction_count'],
                    f"{row['total_purchased']:.2f}",
                    f"{row['credit_balance']:.2f}"
                ))
                
            self.set_status(f"Report generated for {from_date} to {to_date}")
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
            
    def export_report_to_excel(self):
        """Export the current report to Excel."""
        try:
            from_date = self.from_date_entry.get_date().strftime("%Y-%m-%d")
            to_date = self.to_date_entry.get_date().strftime("%Y-%m-%d")
            
            # Ask for filename
            filename = filedialog.asksaveasfilename(
                title="Save Report",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"report_{from_date}_to_{to_date}.xlsx"
            )
            
            if not filename:
                return
                
            # Export to Excel
            success = self.db_manager.export_to_excel(from_date, to_date, filename)
            
            if success:
                messagebox.showinfo("Export Complete", f"Report has been exported to {filename}")
                self.set_status(f"Report exported to {filename}")
            else:
                messagebox.showerror("Export Failed", "Failed to export report.")
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
            
    def view_today_report(self):
        """View report for today's date."""
        try:
            today = datetime.datetime.now().date()
            
            # Ask if user wants to view the report in-app or export to Excel
            choice = messagebox.askyesno(
                "Report Options",
                "Do you want to export today's report to Excel with the stylish template?\n\n"
                "Select 'Yes' to use the Quick Export feature.\n"
                "Select 'No' to view the standard report in the app.",
                icon=messagebox.QUESTION
            )
            
            if choice:
                # Use the quick export feature
                set_busy_cursor(self.root)
                try:
                    report_path = self.report_manager.quick_export_daily(today.strftime("%Y-%m-%d"))
                    if report_path:
                        self.set_status(f"Daily report generated: {report_path}")
                        messagebox.showinfo("Report Generated", 
                                           f"Daily sales report for today has been generated.")
                        
                        # Ask if user wants to open the report
                        if messagebox.askyesno("Open Report", "Would you like to open the report now?"):
                            os.startfile(report_path) if os.name == 'nt' else os.system(f"xdg-open {report_path}")
                    else:
                        messagebox.showwarning("No Data", 
                                            f"No sales data found for today.")
                finally:
                    set_default_cursor(self.root)
            else:
                # Use the standard in-app report view
                self.from_date_entry.set_date(today)
                self.to_date_entry.set_date(today)
                self.notebook.select(self.reports_tab)
                self.generate_report()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
        
    def backup_current_data(self):
        """Backup current data to Excel file."""
        try:
            success = self.backup_manager.backup_current_date()
            
            if success:
                # Update last backup timestamp
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.config_manager.set("last_backup", timestamp)
                
                messagebox.showinfo("Backup Complete", "Current data has been backed up successfully")
                self.set_status("Backup completed successfully")
            else:
                messagebox.showerror("Backup Failed", "Failed to create backup. Check logs for details.")
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
            
    def export_to_excel(self):
        """Export data to Excel file for a custom date range."""
        try:
            # Ask for date range
            dialog = tk.Toplevel(self.root)
            dialog.title("Export to Excel")
            dialog.geometry("400x200")
            dialog.grab_set()
            
            # Create date fields
            date_frame = ttk.Frame(dialog, padding=10)
            date_frame.pack(fill="x")
            
            ttk.Label(date_frame, text="From Date:").grid(row=0, column=0, sticky="w", pady=5)
            today = datetime.datetime.now()
            first_of_month = today.replace(day=1)
            from_date = DateEntry(date_frame, width=12, 
                                year=first_of_month.year, 
                                month=first_of_month.month, 
                                day=first_of_month.day)
            from_date.grid(row=0, column=1, sticky="w", pady=5)
            
            ttk.Label(date_frame, text="To Date:").grid(row=1, column=0, sticky="w", pady=5)
            to_date = DateEntry(date_frame, width=12, 
                              year=today.year, month=today.month, day=today.day)
            to_date.grid(row=1, column=1, sticky="w", pady=5)
            
            # Buttons
            button_frame = ttk.Frame(dialog, padding=10)
            button_frame.pack(fill="x")
            
            def do_export():
                try:
                    from_date_str = from_date.get_date().strftime("%Y-%m-%d")
                    to_date_str = to_date.get_date().strftime("%Y-%m-%d")
                    
                    # Ask for filename
                    filename = filedialog.asksaveasfilename(
                        title="Save Export",
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                        initialfile=f"export_{from_date_str}_to_{to_date_str}.xlsx"
                    )
                    
                    if not filename:
                        return
                        
                    # Export to Excel
                    success = self.db_manager.export_to_excel(from_date_str, to_date_str, filename)
                    
                    if success:
                        messagebox.showinfo("Export Complete", 
                                           f"Data has been exported to {filename}")
                        self.set_status(f"Data exported to {filename}")
                    else:
                        messagebox.showerror("Export Failed", 
                                            "Failed to export data. Check logs for details.")
                        
                    dialog.destroy()
                        
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {str(e)}")
                    traceback.print_exc()
                    
            ttk.Button(button_frame, text="Export", command=do_export).pack(side="right", padx=5)
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
            
    def quick_export_daily(self):
        """Generate a quick daily sales report for the current date."""
        try:
            # Use current date by default
            date_str = datetime.datetime.now().strftime("%Y-%m-%d")
            
            # Show date selection dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Daily Sales Report")
            dialog.geometry("300x150")
            dialog.grab_set()
            
            ttk.Label(dialog, text="Select Date:").pack(pady=(20, 5))
            date_entry = DateEntry(dialog, width=12)
            date_entry.pack(pady=5)
            
            def do_export():
                try:
                    selected_date = date_entry.get_date().strftime("%Y-%m-%d")
                    set_busy_cursor(self.root)
                    dialog.destroy()
                    
                    # Generate filename
                    company_name = self.config_manager.get("company_name", "POS").replace(" ", "_")
                    filename = f"{company_name}_Daily_Report_{selected_date}.xlsx"
                    
                    # Export to Excel
                    success = self.db_manager.export_to_excel(selected_date, selected_date, filename)
                    
                    if success:
                        self.set_status(f"Daily report exported to {filename}")
                        messagebox.showinfo("Export Complete", 
                                          f"Daily sales report for {selected_date} has been exported to {filename}")
                        
                        # Ask if user wants to open the report
                        if messagebox.askyesno("Open Report", "Would you like to open the report now?"):
                            os.startfile(filename) if os.name == 'nt' else os.system(f"xdg-open {filename}")
                    else:
                        messagebox.showwarning("No Data", 
                                            f"No sales data found for {selected_date}.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to generate report: {str(e)}")
                    traceback.print_exc()
                finally:
                    set_default_cursor(self.root)
            
            button_frame = ttk.Frame(dialog)
            button_frame.pack(fill="x", pady=20)
            
            ttk.Button(button_frame, text="Generate", command=do_export).pack(side="right", padx=10)
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=10)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
            
    def quick_export_weekly(self):
        """Generate a quick weekly report."""
        try:
            # Calculate dates for current week (Monday to Sunday)
            today = datetime.datetime.now()
            # Get the current weekday (0 is Monday, 6 is Sunday)
            weekday = today.weekday()
            # Calculate Monday (start of week)
            start_date = today - datetime.timedelta(days=weekday)
            # Calculate Sunday (end of week)
            end_date = start_date + datetime.timedelta(days=6)
            
            # Show date range selection dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Weekly Report")
            dialog.geometry("400x200")
            dialog.grab_set()
            
            ttk.Label(dialog, text="Select Week Range:", font=("Segoe UI", 10, "bold")).pack(pady=(15, 5))
            
            # Date selection frame
            date_frame = ttk.Frame(dialog)
            date_frame.pack(pady=10, padx=20, fill="x")
            
            ttk.Label(date_frame, text="From:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            from_date = DateEntry(date_frame, width=12, 
                                year=start_date.year, month=start_date.month, day=start_date.day)
            from_date.grid(row=0, column=1, padx=5, pady=5, sticky="w")
            
            ttk.Label(date_frame, text="To:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
            to_date = DateEntry(date_frame, width=12, 
                              year=end_date.year, month=end_date.month, day=end_date.day)
            to_date.grid(row=0, column=3, padx=5, pady=5, sticky="w")
            
            # Quick selection buttons
            quick_frame = ttk.Frame(dialog)
            quick_frame.pack(pady=5)
            
            def set_this_week():
                # Reset to this week
                from_date.set_date(start_date)
                to_date.set_date(end_date)
                
            def set_last_week():
                # Set to last week
                last_week_start = start_date - datetime.timedelta(days=7)
                last_week_end = end_date - datetime.timedelta(days=7)
                from_date.set_date(last_week_start)
                to_date.set_date(last_week_end)
                
            ttk.Button(quick_frame, text="This Week", command=set_this_week).pack(side="left", padx=5)
            ttk.Button(quick_frame, text="Last Week", command=set_last_week).pack(side="left", padx=5)
            
            def do_export():
                try:
                    # Get selected dates
                    from_date_str = from_date.get_date().strftime("%Y-%m-%d")
                    to_date_str = to_date.get_date().strftime("%Y-%m-%d")
                    
                    # Check date order
                    if from_date.get_date() > to_date.get_date():
                        messagebox.showwarning("Invalid Date Range", "Start date must be before end date.")
                        return
                        
                    set_busy_cursor(self.root)
                    self.set_status("Generating weekly report...")
                    dialog.destroy()
                    
                    # Generate filename
                    company_name = self.config_manager.get("company_name", "POS").replace(" ", "_")
                    date_range = f"{from_date_str}_to_{to_date_str}"
                    filename = f"{company_name}_Weekly_Report_{date_range}.xlsx"
                    
                    # Export to Excel
                    success = self.db_manager.export_to_excel(from_date_str, to_date_str, filename)
                    
                    if success:
                        self.set_status(f"Weekly report exported to {filename}")
                        messagebox.showinfo("Export Complete", 
                                          f"Weekly report for {from_date_str} to {to_date_str} has been exported to {filename}")
                        
                        # Ask if user wants to open the report
                        if messagebox.askyesno("Open Report", "Would you like to open the report now?"):
                            os.startfile(filename) if os.name == 'nt' else os.system(f"xdg-open {filename}")
                    else:
                        messagebox.showwarning("No Data", 
                                             f"No sales data found for the selected date range.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to generate report: {str(e)}")
                    traceback.print_exc()
                finally:
                    set_default_cursor(self.root)
                    
            # Button frame
            button_frame = ttk.Frame(dialog)
            button_frame.pack(fill="x", pady=15, padx=20)
            
            ttk.Button(button_frame, text="Generate Report", command=do_export, 
                     style="Primary.TButton").pack(side="right", padx=5)
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
    
    def quick_export_monthly(self):
        """Generate a quick monthly summary report."""
        try:
            # Use current month by default
            today = datetime.datetime.now()
            
            # Show month selection dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Monthly Summary Report")
            dialog.geometry("300x180")
            dialog.grab_set()
            
            ttk.Label(dialog, text="Select Month and Year:").pack(pady=(20, 5))
            
            # Month and year selection
            selection_frame = ttk.Frame(dialog)
            selection_frame.pack(pady=5)
            
            month_var = tk.StringVar(value=today.strftime("%B"))
            year_var = tk.IntVar(value=today.year)
            
            months = ["January", "February", "March", "April", "May", "June", 
                    "July", "August", "September", "October", "November", "December"]
            
            ttk.Label(selection_frame, text="Month:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
            month_combo = ttk.Combobox(selection_frame, textvariable=month_var, values=months, width=10)
            month_combo.grid(row=0, column=1, padx=5, pady=5, sticky="w")
            
            ttk.Label(selection_frame, text="Year:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
            year_spin = ttk.Spinbox(selection_frame, from_=2000, to=2100, textvariable=year_var, width=8)
            year_spin.grid(row=1, column=1, padx=5, pady=5, sticky="w")
            
            def do_export():
                try:
                    month_idx = months.index(month_var.get()) + 1
                    year = year_var.get()
                    
                    set_busy_cursor(self.root)
                    dialog.destroy()
                    
                    # Calculate first and last day of month
                    first_day = datetime.datetime(year, month_idx, 1)
                    # Get last day by getting first day of next month and subtracting 1 day
                    if month_idx == 12:
                        last_day = datetime.datetime(year+1, 1, 1) - datetime.timedelta(days=1)
                    else:
                        last_day = datetime.datetime(year, month_idx+1, 1) - datetime.timedelta(days=1)
                    
                    # Format dates for filename and display
                    date_str = first_day.strftime("%Y-%m")
                    start_date_str = first_day.strftime("%Y-%m-%d")
                    end_date_str = last_day.strftime("%Y-%m-%d")
                    
                    # Generate filename
                    company_name = self.config_manager.get("company_name", "POS").replace(" ", "_")
                    filename = f"{company_name}_Monthly_Report_{date_str}.xlsx"
                    
                    # Export to Excel
                    success = self.db_manager.export_to_excel(start_date_str, end_date_str, filename)
                    
                    if success:
                        self.set_status(f"Monthly report exported to {filename}")
                        messagebox.showinfo("Export Complete", 
                                          f"Monthly report for {month_var.get()} {year} has been exported to {filename}")
                        
                        # Ask if user wants to open the report
                        if messagebox.askyesno("Open Report", "Would you like to open the report now?"):
                            os.startfile(filename) if os.name == 'nt' else os.system(f"xdg-open {filename}")
                    else:
                        messagebox.showwarning("No Data", 
                                            f"No data found for {month_var.get()} {year}.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to generate report: {str(e)}")
                    traceback.print_exc()
                finally:
                    set_default_cursor(self.root)
            
            button_frame = ttk.Frame(dialog)
            button_frame.pack(fill="x", pady=20)
            
            ttk.Button(button_frame, text="Generate", command=do_export).pack(side="right", padx=10)
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=10)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
    
    def quick_export_inventory(self):
        """Generate a quick inventory report."""
        try:
            set_busy_cursor(self.root)
            
            # Dialog to select date for inventory
            dialog = tk.Toplevel(self.root)
            dialog.title("Inventory Report")
            dialog.geometry("300x150")
            dialog.grab_set()
            
            ttk.Label(dialog, text="Select Date for Inventory:").pack(pady=(20, 5))
            date_entry = DateEntry(dialog, width=12)
            date_entry.pack(pady=5)
            
            def do_export():
                try:
                    selected_date = date_entry.get_date().strftime("%Y-%m-%d")
                    dialog.destroy()
                    
                    # Calculate inventory
                    inventory_data = self.db_manager.calculate_daily_inventory(selected_date)
                    
                    if inventory_data is not None and not inventory_data.empty:
                        # Generate filename
                        company_name = self.config_manager.get("company_name", "POS").replace(" ", "_")
                        filename = f"{company_name}_Inventory_Report_{selected_date}.xlsx"
                        
                        # Save to Excel
                        with pd.ExcelWriter(filename) as writer:
                            inventory_data.to_excel(writer, sheet_name="Inventory", index=False)
                            
                        self.set_status(f"Inventory report exported to {filename}")
                        messagebox.showinfo("Export Complete", 
                                         f"Inventory report for {selected_date} has been exported to {filename}")
                        
                        # Ask if user wants to open the report
                        if messagebox.askyesno("Open Report", "Would you like to open the report now?"):
                            os.startfile(filename) if os.name == 'nt' else os.system(f"xdg-open {filename}")
                    else:
                        messagebox.showwarning("No Data", 
                                            f"No inventory data found for {selected_date}.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to generate report: {str(e)}")
                    traceback.print_exc()
                finally:
                    set_default_cursor(self.root)
            
            button_frame = ttk.Frame(dialog)
            button_frame.pack(fill="x", pady=20)
            
            ttk.Button(button_frame, text="Generate", command=do_export).pack(side="right", padx=10)
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=10)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")
            traceback.print_exc()
            set_default_cursor(self.root)
    
    def quick_export_credit(self):
        """Generate a quick customer credit report."""
        try:
            set_busy_cursor(self.root)
            
            # Get all customers with their credit balances
            conn = self.db_manager.connect()
            cursor = conn.cursor()
            cursor.execute("SELECT id, name, contact, credit_balance FROM customers WHERE is_supplier = 0")
            customers = cursor.fetchall()
            
            if customers:
                # Create DataFrame for Excel export
                customers_df = pd.DataFrame(customers)
                
                # Generate filename
                company_name = self.config_manager.get("company_name", "POS").replace(" ", "_")
                today = datetime.datetime.now().strftime("%Y-%m-%d")
                filename = f"{company_name}_Credit_Report_{today}.xlsx"
                
                # Export to Excel
                with pd.ExcelWriter(filename) as writer:
                    customers_df.to_excel(writer, sheet_name="Customer Credit", index=False)
                    
                self.set_status(f"Credit report exported to {filename}")
                messagebox.showinfo("Export Complete", 
                                 f"Customer credit report has been exported to {filename}")
                
                # Ask if user wants to open the report
                if messagebox.askyesno("Open Report", "Would you like to open the report now?"):
                    os.startfile(filename) if os.name == 'nt' else os.system(f"xdg-open {filename}")
            else:
                messagebox.showwarning("No Data", "No customer data found.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")
            traceback.print_exc()
        finally:
            set_default_cursor(self.root)
            
    def sync_with_sheets(self):
        """Synchronize data with Google Sheets."""
        if self.config_manager.get("app_mode") != "online":
            if messagebox.askyesno("Mode Change", 
                                  "You're currently in offline mode. Switch to online mode?"):
                self.config_manager.set("app_mode", "online")
                self.initialize_sheets_manager()
            else:
                return
                
        if not hasattr(self, 'sheets_manager') or not self.sheets_manager or not self.sheets_manager.client:
            messagebox.showwarning("Not Connected", 
                                  "Not connected to Google Sheets. Please check your configuration.")
            self.open_config()
            return
            
        try:
            # Get sync date range
            dialog = tk.Toplevel(self.root)
            dialog.title("Sync with Google Sheets")
            dialog.geometry("400x200")
            dialog.grab_set()
            
            # Create date fields
            date_frame = ttk.Frame(dialog, padding=10)
            date_frame.pack(fill="x")
            
            ttk.Label(date_frame, text="From Date:").grid(row=0, column=0, sticky="w", pady=5)
            today = datetime.datetime.now()
            from_date = DateEntry(date_frame, width=12, 
                                year=today.year, month=today.month, day=today.day)
            from_date.grid(row=0, column=1, sticky="w", pady=5)
            
            ttk.Label(date_frame, text="To Date:").grid(row=1, column=0, sticky="w", pady=5)
            to_date = DateEntry(date_frame, width=12, 
                              year=today.year, month=today.month, day=today.day)
            to_date.grid(row=1, column=1, sticky="w", pady=5)
            
            # Buttons
            button_frame = ttk.Frame(dialog, padding=10)
            button_frame.pack(fill="x")
            
            def do_sync():
                try:
                    from_date_str = from_date.get_date().strftime("%Y-%m-%d")
                    to_date_str = to_date.get_date().strftime("%Y-%m-%d")
                    
                    # Show progress indicator
                    self.set_status("Syncing with Google Sheets...", "online")
                    dialog.destroy()
                    
                    # Process each date in the range
                    current_date = from_date.get_date()
                    end_date = to_date.get_date()
                    success_count = 0
                    
                    while current_date <= end_date:
                        date_str = current_date.strftime("%Y-%m-%d")
                        sales_data = self.db_manager.get_daily_sales_detail(date_str)
                        
                        if sales_data:
                            header_row = ["Customer Name", "Total Amount for", "Previous Credit", 
                                         "Cash Received", "Updated Credit Balance", "Extra"]
                            if self.sheets_manager.sync_sales_to_sheet(date_str, sales_data, header_row):
                                success_count += 1
                                
                        # Move to next day
                        current_date += datetime.timedelta(days=1)
                        
                    if success_count > 0:
                        messagebox.showinfo("Sync Complete", 
                                           f"Successfully synced {success_count} day(s) to Google Sheets")
                        self.set_status("Sync completed successfully", "online")
                    else:
                        messagebox.showinfo("Sync Complete", "No data found to sync for the selected dates")
                        self.set_status("Sync completed (no data found)", "online")
                        
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred during sync: {str(e)}")
                    traceback.print_exc()
                    self.set_status("Sync failed. See error message.", "online")
                    
            ttk.Button(button_frame, text="Sync", command=do_sync).pack(side="right", padx=5)
            ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="right", padx=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
            
    def open_products_manager(self):
        """Open the products manager dialog."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Products Manager")
        dialog.geometry("600x500")
        dialog.grab_set()
        
        # Create products treeview
        tree_frame = ttk.Frame(dialog, padding=10)
        tree_frame.pack(fill="both", expand=True)
        
        products_tree = ttk.Treeview(tree_frame, columns=("id", "name", "price"), 
                                   show="headings", selectmode="browse")
        products_tree.heading("id", text="ID")
        products_tree.heading("name", text="Product Name")
        products_tree.heading("price", text="Price")
        
        products_tree.column("id", width=50)
        products_tree.column("name", width=300)
        products_tree.column("price", width=100)
        
        products_tree.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=products_tree.yview)
        scrollbar.pack(side="right", fill="y")
        products_tree.configure(yscrollcommand=scrollbar.set)
        
        # Load products
        def load_products():
            for item in products_tree.get_children():
                products_tree.delete(item)
                
            products = self.db_manager.get_all_products()
            for product in products:
                products_tree.insert("", "end", values=(
                    product['id'],
                    product['name'],
                    f"{product['price']:.2f}"
                ))
                
        load_products()
        
        # Buttons for actions
        button_frame = ttk.Frame(dialog, padding=10)
        button_frame.pack(fill="x")
        
        def add_product():
            # Open add dialog
            add_dialog = tk.Toplevel(dialog)
            add_dialog.title("Add Product")
            add_dialog.geometry("300x150")
            add_dialog.grab_set()
            
            ttk.Label(add_dialog, text="Product Name:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
            name_var = tk.StringVar()
            ttk.Entry(add_dialog, textvariable=name_var, width=20).grid(row=0, column=1, padx=10, pady=5)
            
            ttk.Label(add_dialog, text="Price:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
            price_var = tk.StringVar(value="0.00")
            ttk.Entry(add_dialog, textvariable=price_var, width=20).grid(row=1, column=1, padx=10, pady=5)
            
            def do_add():
                try:
                    name = name_var.get().strip()
                    price = safe_float(price_var.get(), 0.0)
                    
                    if not name:
                        messagebox.showwarning("Input Error", "Please enter a product name")
                        return
                        
                    success = self.db_manager.add_or_update_product(name, price)
                    if success:
                        load_products()
                        add_dialog.destroy()
                    else:
                        messagebox.showerror("Error", "Failed to add product")
                        
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {str(e)}")
                    traceback.print_exc()
                    
            button_frame = ttk.Frame(add_dialog)
            button_frame.grid(row=2, column=0, columnspan=2, pady=15)
            ttk.Button(button_frame, text="Add", command=do_add).grid(row=0, column=0, padx=10)
            ttk.Button(button_frame, text="Cancel", command=add_dialog.destroy).grid(row=0, column=1, padx=10)
            
        def edit_product():
            selected = products_tree.selection()
            if not selected:
                messagebox.showwarning("Selection Required", "Please select a product to edit")
                return
                
            values = products_tree.item(selected, 'values')
            product_id = values[0]
            product_name = values[1]
            product_price = float(values[2])
            
            # Open edit dialog
            edit_dialog = tk.Toplevel(dialog)
            edit_dialog.title("Edit Product")
            edit_dialog.geometry("300x150")
            edit_dialog.grab_set()
            
            ttk.Label(edit_dialog, text="Product Name:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
            name_var = tk.StringVar(value=product_name)
            ttk.Entry(edit_dialog, textvariable=name_var, width=20).grid(row=0, column=1, padx=10, pady=5)
            
            ttk.Label(edit_dialog, text="Price:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
            price_var = tk.StringVar(value=f"{product_price:.2f}")
            ttk.Entry(edit_dialog, textvariable=price_var, width=20).grid(row=1, column=1, padx=10, pady=5)
            
            def do_edit():
                try:
                    name = name_var.get().strip()
                    price = safe_float(price_var.get(), 0.0)
                    
                    if not name:
                        messagebox.showwarning("Input Error", "Please enter a product name")
                        return
                        
                    # Update in database
                    conn = self.db_manager.connect()
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE products SET name = ?, price = ? WHERE id = ?",
                        (name, price, product_id)
                    )
                    conn.commit()
                    
                    load_products()
                    edit_dialog.destroy()
                        
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {str(e)}")
                    traceback.print_exc()
                    
            button_frame = ttk.Frame(edit_dialog)
            button_frame.grid(row=2, column=0, columnspan=2, pady=15)
            ttk.Button(button_frame, text="Save", command=do_edit).grid(row=0, column=0, padx=10)
            ttk.Button(button_frame, text="Cancel", command=edit_dialog.destroy).grid(row=0, column=1, padx=10)
            
        def delete_product():
            selected = products_tree.selection()
            if not selected:
                messagebox.showwarning("Selection Required", "Please select a product to delete")
                return
                
            values = products_tree.item(selected, 'values')
            product_id = values[0]
            product_name = values[1]
            
            if messagebox.askyesno("Confirm Delete", 
                                  f"Are you sure you want to delete '{product_name}'?\n\n"
                                  f"This will only mark it as inactive, not remove it from the database."):
                try:
                    # Soft delete in database
                    conn = self.db_manager.connect()
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE products SET active = 0 WHERE id = ?",
                        (product_id,)
                    )
                    conn.commit()
                    
                    load_products()
                        
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {str(e)}")
                    traceback.print_exc()
                    
        ttk.Button(button_frame, text="Add Product", command=add_product).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Edit Selected", command=edit_product).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Delete Selected", command=delete_product).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side="right", padx=5)
        
    def open_customers_manager(self):
        """Open the customers manager dialog."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Customers Manager")
        dialog.geometry("700x500")
        dialog.grab_set()
        
        # Create customers treeview
        tree_frame = ttk.Frame(dialog, padding=10)
        tree_frame.pack(fill="both", expand=True)
        
        customers_tree = ttk.Treeview(tree_frame, 
                                    columns=("id", "name", "contact", "credit"),
                                    show="headings", selectmode="browse")
        customers_tree.heading("id", text="ID")
        customers_tree.heading("name", text="Customer Name")
        customers_tree.heading("contact", text="Contact")
        customers_tree.heading("credit", text="Credit Balance")
        
        customers_tree.column("id", width=50)
        customers_tree.column("name", width=250)
        customers_tree.column("contact", width=150)
        customers_tree.column("credit", width=100)
        
        customers_tree.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=customers_tree.yview)
        scrollbar.pack(side="right", fill="y")
        customers_tree.configure(yscrollcommand=scrollbar.set)
        
        # Load customers
        def load_customers():
            for item in customers_tree.get_children():
                customers_tree.delete(item)
                
            customers = self.db_manager.get_all_customers()
            for customer in customers:
                customers_tree.insert("", "end", values=(
                    customer['id'],
                    customer['name'],
                    customer.get('contact', ''),
                    f"{customer.get('credit_balance', 0.0):.2f}"
                ))
                
        load_customers()
        
        # Buttons for actions
        button_frame = ttk.Frame(dialog, padding=10)
        button_frame.pack(fill="x")
        
        def add_customer():
            # Open add dialog
            add_dialog = tk.Toplevel(dialog)
            add_dialog.title("Add Customer")
            add_dialog.geometry("300x200")
            add_dialog.grab_set()
            
            ttk.Label(add_dialog, text="Customer Name:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
            name_var = tk.StringVar()
            ttk.Entry(add_dialog, textvariable=name_var, width=20).grid(row=0, column=1, padx=10, pady=5)
            
            ttk.Label(add_dialog, text="Contact:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
            contact_var = tk.StringVar()
            ttk.Entry(add_dialog, textvariable=contact_var, width=20).grid(row=1, column=1, padx=10, pady=5)
            
            ttk.Label(add_dialog, text="Initial Credit:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
            credit_var = tk.StringVar(value="0.00")
            ttk.Entry(add_dialog, textvariable=credit_var, width=20).grid(row=2, column=1, padx=10, pady=5)
            
            def do_add():
                try:
                    name = name_var.get().strip()
                    contact = contact_var.get().strip()
                    credit = safe_float(credit_var.get(), 0.0)
                    
                    if not name:
                        messagebox.showwarning("Input Error", "Please enter a customer name")
                        return
                        
                    success = self.db_manager.add_or_update_customer(name, contact, credit)
                    if success:
                        load_customers()
                        self.update_customer_list()  # Update main form dropdown
                        add_dialog.destroy()
                    else:
                        messagebox.showerror("Error", "Failed to add customer")
                        
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {str(e)}")
                    traceback.print_exc()
                    
            button_frame = ttk.Frame(add_dialog)
            button_frame.grid(row=3, column=0, columnspan=2, pady=15)
            ttk.Button(button_frame, text="Add", command=do_add).grid(row=0, column=0, padx=10)
            ttk.Button(button_frame, text="Cancel", command=add_dialog.destroy).grid(row=0, column=1, padx=10)
            
        def edit_customer():
            selected = customers_tree.selection()
            if not selected:
                messagebox.showwarning("Selection Required", "Please select a customer to edit")
                return
                
            values = customers_tree.item(selected, 'values')
            customer_id = values[0]
            customer_name = values[1]
            customer_contact = values[2]
            customer_credit = float(values[3])
            
            # Open edit dialog
            edit_dialog = tk.Toplevel(dialog)
            edit_dialog.title("Edit Customer")
            edit_dialog.geometry("300x200")
            edit_dialog.grab_set()
            
            ttk.Label(edit_dialog, text="Customer Name:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
            name_var = tk.StringVar(value=customer_name)
            ttk.Entry(edit_dialog, textvariable=name_var, width=20).grid(row=0, column=1, padx=10, pady=5)
            
            ttk.Label(edit_dialog, text="Contact:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
            contact_var = tk.StringVar(value=customer_contact)
            ttk.Entry(edit_dialog, textvariable=contact_var, width=20).grid(row=1, column=1, padx=10, pady=5)
            
            ttk.Label(edit_dialog, text="Credit Balance:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
            credit_var = tk.StringVar(value=f"{customer_credit:.2f}")
            ttk.Entry(edit_dialog, textvariable=credit_var, width=20).grid(row=2, column=1, padx=10, pady=5)
            
            def do_edit():
                try:
                    name = name_var.get().strip()
                    contact = contact_var.get().strip()
                    credit = safe_float(credit_var.get(), 0.0)
                    
                    if not name:
                        messagebox.showwarning("Input Error", "Please enter a customer name")
                        return
                        
                    # Update in database
                    conn = self.db_manager.connect()
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE customers SET name = ?, contact = ?, credit_balance = ? WHERE id = ?",
                        (name, contact, credit, customer_id)
                    )
                    conn.commit()
                    
                    load_customers()
                    self.update_customer_list()  # Update main form dropdown
                    edit_dialog.destroy()
                        
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {str(e)}")
                    traceback.print_exc()
                    
            button_frame = ttk.Frame(edit_dialog)
            button_frame.grid(row=3, column=0, columnspan=2, pady=15)
            ttk.Button(button_frame, text="Save", command=do_edit).grid(row=0, column=0, padx=10)
            ttk.Button(button_frame, text="Cancel", command=edit_dialog.destroy).grid(row=0, column=1, padx=10)
            
        def delete_customer():
            selected = customers_tree.selection()
            if not selected:
                messagebox.showwarning("Selection Required", "Please select a customer to delete")
                return
                
            values = customers_tree.item(selected, 'values')
            customer_id = values[0]
            customer_name = values[1]
            
            if messagebox.askyesno("Confirm Delete", 
                                  f"Are you sure you want to delete '{customer_name}'?\n\n"
                                  f"This will only mark it as inactive, not remove it from the database."):
                try:
                    # Soft delete in database
                    conn = self.db_manager.connect()
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE customers SET active = 0 WHERE id = ?",
                        (customer_id,)
                    )
                    conn.commit()
                    
                    load_customers()
                    self.update_customer_list()  # Update main form dropdown
                        
                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred: {str(e)}")
                    traceback.print_exc()
                    
        ttk.Button(button_frame, text="Add Customer", command=add_customer).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Edit Selected", command=edit_customer).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Delete Selected", command=delete_customer).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side="right", padx=5)
        
    def open_config(self):
        """Open the configuration dialog."""
        config_screen = ConfigScreen(self.root, self.config_manager, self.on_config_saved)
        
    def on_config_saved(self, new_settings):
        """Handle configuration changes."""
        # Apply theme
        self.apply_theme()
        
        # Update mode
        if new_settings.get("app_mode") == "online":
            self.initialize_sheets_manager()
        else:
            self.set_status("Running in offline mode", "offline")
            
        # Update company info in report manager
        company_info = {
            'name': new_settings.get('company_name', 'Enhanced POS System'),
            'address': new_settings.get('company_address', ''),
            'phone': new_settings.get('company_phone', ''),
            'email': new_settings.get('company_email', ''),
        }
        self.report_manager.set_company_info(company_info)
            
        # Update home tab to reflect changes
        self.create_home_tab()
        
    def show_about(self):
        """Show about dialog."""
        messagebox.showinfo("About Enhanced POS System", 
                           "Enhanced POS System\n\n"
                           "A user-friendly point of sale system with online and offline capabilities.\n\n"
                           "Features:\n"
                           "- Works seamlessly in both online and offline modes\n"
                           "- Local Excel backup functionality\n"
                           "- Google Sheets synchronization\n"
                           "- Comprehensive reporting\n"
                           "- Easy configuration")

# --- Main Application ---
def initialize_sample_data():
    """Initialize sample data from the provided lists."""
    try:
        # Check if init_data.py exists
        if os.path.exists('init_data.py'):
            print("Initializing sample data...")
            import init_data
            init_data.initialize_database()
            print("Sample data initialization complete.")
        else:
            print("Sample data initialization script not found. Skipping.")
    except Exception as e:
        print(f"Error initializing sample data: {e}")
        traceback.print_exc()

def main():
    """Main application entry point."""
    # Ensure directories exist
    ensure_directory_exists(BACKUP_FOLDER)
    ensure_directory_exists(ICON_PATH)
    
    # Initialize sample data
    initialize_sample_data()
    
    # Create and run the application
    root = tk.Tk()
    root.withdraw()  # Hide the main window until setup is complete
    
    app = POSApp(root)
    
    # Show the main window
    root.deiconify()
    
    # Run the application
    root.mainloop()

if __name__ == "__main__":
    main()
"""
FUNLEX CAKES AND CREAM – Sales & Stock Management App
======================================================
Secure internal tool for recording sales, tracking stock,
and exporting data to Excel.
"""

import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
from calendar import monthrange
from pathlib import Path
import uuid
import hashlib
from io import BytesIO

# ─────────────────────────────────────────────
# PATHS & CONFIG
# ─────────────────────────────────────────────
APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "funlex_data.db"
LOGO_PATH = APP_DIR / "funlex_logo.png"
DEFAULT_ADMIN_CODE = "FUNLEX2026"
DEFAULT_SALES_CODE = "SALES2026"
ADMIN_CODE_HASH = hashlib.sha256(DEFAULT_ADMIN_CODE.encode()).hexdigest()
SALES_CODE_HASH = hashlib.sha256(DEFAULT_SALES_CODE.encode()).hexdigest()

DEFAULT_PRODUCTS = [
    ("1000 Cake", "6250073", 1000, 0),
    ("2500 Slice Cake", "6250069", 2500, 0),
    ("1500 Yoghurt", "6225367", 1500, 0),
    ("4000 Flaky Chinchin", "6162121", 4000, 0),
    ("Cake Slice", "6149285", 3000, 0),
    ("Dudu Drink", "6015590", 500, 0),
    ("Puffpuff 110pcs", "PP110", 1100, 0),
    ("Small Sosa", "SS01", 500, 0),
    ("Vanilla Foil Cake", "VFC01", 2500, 0),
    ("Yoghurt Parfait 500ml", "YP500", 1500, 0),
    ("1200 Chinchin (22pcs)", "CH1200", 1200, 0),
    ("2000 Chinchin", "CH2000", 2000, 0),
    ("Big 5Alive", "B5A", 800, 0),
    ("Predator", "PRED", 600, 0),
    ("Small 5Alive", "S5A", 400, 0),
    ("Cake Slice + Parfait", "CSP01", 3500, 0),
    ("500g Cube Chinchin", "CCH500", 1500, 0),
    ("Can Malt", "MALT", 500, 0),
    ("Bottle Water", "BW01", 200, 0),
    ("Pop Corn", "PC01", 500, 0),
    ("200 Chinchin", "CH200", 200, 0),
    ("300 Cookies", "CK300", 300, 0),
    ("Zobo Drink", "ZOBO", 300, 0),
    ("Family Bread", "FB01", 800, 0),
    ("Yoghurt 35cl", "Y35", 400, 0),
    ("2K Cookies", "CK2K", 2000, 0),
    ("Spring Roll (50pcs)", "SR50", 3000, 0),
    ("Can Fanta", "FANTA", 400, 0),
    ("Smallchop 1pc", "SC01", 200, 0),
    ("Samosa (50pcs)", "SAM50", 3000, 0),
    ("Jam Doughnut 1kg (20pcs)", "JD20", 2500, 0),
    ("1000 Cookies", "CK1K", 1000, 0),
    ("Meatpie 1kg (20pcs)", "MP20", 3000, 0),
    ("Eggroll 22pcs", "ER22", 1500, 0),
    ("Red Velvet Foil Cake", "RVFC", 3500, 0),
    ("Red Velvet Cup Cake", "RVCC", 500, 0),
    ("Viju Yoghurt", "VIJU", 600, 0),
    ("2000 Flaky Chinchin", "FCH2K", 2000, 0),
    ("Hotdog Sausage (22pcs)", "HS22", 2000, 0),
    ("Berry Blast", "BB01", 800, 0),
    ("Can Coke", "COKE", 400, 0),
    ("Buns (24pcs)", "BUNS24", 1500, 0),
    ("Economy Loaf", "ELOAF", 600, 0),
    ("300 Cup Cake", "CC300", 300, 0),
    ("Doughnut 1kg (20pcs)", "DN20", 2000, 0),
    ("Can Sprite", "SPRITE", 400, 0),
    ("Meat Pie Single", "MPS", 500, 0),
    ("Sausage Roll", "SR01", 400, 0),
    ("Chin Chin Mixed", "CHMIX", 1500, 0),
    ("Fruit Juice 50cl", "FJ50", 700, 0),
]

# ─────────────────────────────────────────────
# DATABASE
# ─────────────────────────────────────────────
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            sku TEXT,
            unit_price REAL NOT NULL DEFAULT 0,
            cost_price REAL NOT NULL DEFAULT 0,
            current_stock REAL NOT NULL DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # migrate cost_price on older DBs
    try:
        cur.execute("PRAGMA table_info(products)")
        cols = [r[1] for r in cur.fetchall()]
        if "cost_price" not in cols:
            cur.execute("ALTER TABLE products ADD COLUMN cost_price REAL NOT NULL DEFAULT 0")
    except Exception:
        pass
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unique_id TEXT UNIQUE NOT NULL,
            sale_date TEXT NOT NULL,
            product_id INTEGER NOT NULL,
            product_name TEXT NOT NULL,
            opening_stock REAL NOT NULL,
            units_sold REAL NOT NULL,
            unit_price REAL NOT NULL,
            amount REAL NOT NULL,
            closing_stock REAL NOT NULL,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS stock_movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unique_id TEXT UNIQUE NOT NULL,
            move_date TEXT NOT NULL,
            product_id INTEGER NOT NULL,
            product_name TEXT NOT NULL,
            movement_type TEXT NOT NULL,
            quantity REAL NOT NULL,
            stock_before REAL NOT NULL,
            stock_after REAL NOT NULL,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS special_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unique_id TEXT UNIQUE NOT NULL,
            order_date TEXT NOT NULL,
            customer_name TEXT,
            product_description TEXT NOT NULL,
            quantity REAL NOT NULL,
            unit_price REAL NOT NULL DEFAULT 0,
            amount REAL NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'Pending',
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS branch_daily (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unique_id TEXT UNIQUE NOT NULL,
            record_date TEXT NOT NULL,
            product_id INTEGER,
            product_name TEXT NOT NULL,
            qty_taken REAL NOT NULL DEFAULT 0,
            qty_sold REAL NOT NULL DEFAULT 0,
            qty_remaining REAL NOT NULL DEFAULT 0,
            amount_sold REAL NOT NULL DEFAULT 0,
            unit_price REAL NOT NULL DEFAULT 0,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS raw_materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            unit TEXT NOT NULL DEFAULT 'kg',
            current_qty REAL NOT NULL DEFAULT 0,
            reorder_level REAL NOT NULL DEFAULT 0,
            unit_cost REAL NOT NULL DEFAULT 0,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS raw_material_movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unique_id TEXT UNIQUE NOT NULL,
            move_date TEXT NOT NULL,
            material_id INTEGER NOT NULL,
            material_name TEXT NOT NULL,
            movement_type TEXT NOT NULL,
            quantity REAL NOT NULL,
            qty_before REAL NOT NULL,
            qty_after REAL NOT NULL,
            unit_cost REAL NOT NULL DEFAULT 0,
            total_cost REAL NOT NULL DEFAULT 0,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (material_id) REFERENCES raw_materials(id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS expenses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unique_id TEXT UNIQUE NOT NULL,
            expense_date TEXT NOT NULL,
            category TEXT NOT NULL,
            description TEXT,
            amount REAL NOT NULL,
            branch TEXT NOT NULL DEFAULT 'Main',
            payment_method TEXT DEFAULT 'Cash',
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS staff (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT NOT NULL,
            role_title TEXT,
            branch TEXT NOT NULL DEFAULT 'Main',
            phone TEXT,
            monthly_salary REAL NOT NULL DEFAULT 0,
            start_date TEXT,
            is_active INTEGER NOT NULL DEFAULT 1,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS salary_payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unique_id TEXT UNIQUE NOT NULL,
            staff_id INTEGER NOT NULL,
            staff_name TEXT NOT NULL,
            pay_date TEXT NOT NULL,
            period_label TEXT,
            amount REAL NOT NULL,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (staff_id) REFERENCES staff(id)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            account_type TEXT NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS journal_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unique_id TEXT UNIQUE NOT NULL,
            entry_date TEXT NOT NULL,
            description TEXT,
            source_type TEXT,
            source_id TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS journal_lines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            journal_id INTEGER NOT NULL,
            account_code TEXT NOT NULL,
            account_name TEXT NOT NULL,
            debit REAL NOT NULL DEFAULT 0,
            credit REAL NOT NULL DEFAULT 0,
            FOREIGN KEY (journal_id) REFERENCES journal_entries(id)
        )
    """)
    cur.execute("SELECT COUNT(*) FROM accounts")
    if cur.fetchone()[0] == 0:
        default_accounts = [
            ("1000", "Cash", "Asset"),
            ("1100", "Bank / Transfer", "Asset"),
            ("1200", "Raw materials inventory", "Asset"),
            ("2000", "Accounts payable (suppliers)", "Liability"),
            ("2100", "Loans payable", "Liability"),
            ("4000", "Sales - Main", "Income"),
            ("4100", "Sales - Punch", "Income"),
            ("4200", "Sales - Special orders", "Income"),
            ("5000", "Cost of materials used", "Expense"),
            ("5100", "Spoilage / waste", "Expense"),
            ("5200", "Operating expenses", "Expense"),
            ("5300", "Salaries", "Expense"),
            ("5400", "Interest expense", "Expense"),
            ("5050", "Cost of goods sold", "Expense"),
            ("1250", "Finished goods inventory", "Asset"),
            ("3000", "Opening balance equity", "Equity"),
        ]
        for code, name, atype in default_accounts:
            cur.execute(
                "INSERT INTO accounts (code, name, account_type, is_active) VALUES (?, ?, ?, 1)",
                (code, name, atype),
            )
    # Ensure liability accounts exist on older databases
    for code, name, atype in [
        ("2000", "Accounts payable (suppliers)", "Liability"),
        ("2100", "Loans payable", "Liability"),
        ("5400", "Interest expense", "Expense"),
        ("5050", "Cost of goods sold", "Expense"),
        ("1250", "Finished goods inventory", "Asset"),
        ("3000", "Opening balance equity", "Equity"),
    ]:
        cur.execute("SELECT id FROM accounts WHERE code = ?", (code,))
        if not cur.fetchone():
            cur.execute(
                "INSERT INTO accounts (code, name, account_type, is_active) VALUES (?, ?, ?, 1)",
                (code, name, atype),
            )

    cur.execute("SELECT COUNT(*) FROM products")
    if cur.fetchone()[0] == 0:
        for name, sku, price, stock in DEFAULT_PRODUCTS:
            cur.execute(
                "INSERT INTO products (name, sku, unit_price, cost_price, current_stock) VALUES (?, ?, ?, ?, ?)",
                (name, sku, price, 0, stock),
            )

    cur.execute("SELECT COUNT(*) FROM raw_materials")
    if cur.fetchone()[0] == 0:
        default_materials = [
            ("Flour", "kg", 0, 10, 0),
            ("Butter", "kg", 0, 5, 0),
            ("Sugar", "kg", 0, 5, 0),
            ("Eggs", "crates", 0, 2, 0),
            ("Milk", "litres", 0, 5, 0),
            ("Yeast", "kg", 0, 1, 0),
            ("Cocoa powder", "kg", 0, 2, 0),
            ("Vanilla essence", "litres", 0, 1, 0),
            ("Baking powder", "kg", 0, 1, 0),
            ("Vegetable oil", "litres", 0, 5, 0),
        ]
        for name, unit, qty, reorder, cost in default_materials:
            cur.execute(
                """INSERT INTO raw_materials (name, unit, current_qty, reorder_level, unit_cost)
                   VALUES (?, ?, ?, ?, ?)""",
                (name, unit, qty, reorder, cost),
            )

    cur.execute("SELECT value FROM settings WHERE key = 'admin_code_hash'")
    if not cur.fetchone():
        cur.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?)",
            ("admin_code_hash", ADMIN_CODE_HASH),
        )
    cur.execute("SELECT value FROM settings WHERE key = 'sales_code_hash'")
    if not cur.fetchone():
        cur.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?)",
            ("sales_code_hash", SALES_CODE_HASH),
        )

    conn.commit()
    conn.close()


def get_setting(key: str, default: str = "") -> str:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT value FROM settings WHERE key = ?", (key,))
    row = cur.fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key: str, value: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
        (key, value),
    )
    conn.commit()
    conn.close()


def verify_login(code: str):
    """Return role 'admin' or 'sales' if code matches, else None."""
    code = code.strip()
    h = hashlib.sha256(code.encode()).hexdigest()
    admin_h = get_setting("admin_code_hash", ADMIN_CODE_HASH)
    sales_h = get_setting("sales_code_hash", SALES_CODE_HASH)
    if h == admin_h:
        return "admin"
    if h == sales_h:
        return "sales"
    return None


def change_role_code(role: str, new_code: str):
    """role: 'admin' or 'sales'"""
    h = hashlib.sha256(new_code.encode()).hexdigest()
    key = "admin_code_hash" if role == "admin" else "sales_code_hash"
    set_setting(key, h)


def get_all_products():
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT id, name, sku, unit_price, COALESCE(cost_price, 0) AS cost_price, current_stock FROM products ORDER BY name",
        conn,
    )
    conn.close()
    return df


def get_product_by_id(pid: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM products WHERE id = ?", (pid,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def add_sale(sale_date, product_id, units_sold, notes="", unit_price=None, payment_method="Cash"):
    """Record a sale. unit_price can override catalogue price for this sale only."""
    prod = get_product_by_id(product_id)
    if not prod:
        return None, "Product not found"
    opening = float(prod["current_stock"])
    units = float(units_sold)
    if units <= 0:
        return None, "Units sold must be greater than 0"
    if units > opening:
        return None, f"Insufficient stock. Available: {opening:.0f}"
    if unit_price is not None and float(unit_price) >= 0:
        price = float(unit_price)
    else:
        price = float(prod["unit_price"])
    amount = units * price
    closing = opening - units
    uid = f"SALE-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO sales
               (unique_id, sale_date, product_id, product_name, opening_stock,
                units_sold, unit_price, amount, closing_stock, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (uid, sale_date, product_id, prod["name"], opening, units, price, amount, closing, notes),
        )
        cur.execute("UPDATE products SET current_stock = ? WHERE id = ?", (closing, product_id))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()

    # Auto journal: revenue + COGS
    try:
        auto_journal_sale(sale_date, amount, prod["name"], uid, payment=payment_method or "Cash")
        cost_p = float(prod.get("cost_price") or 0)
        cogs_amt = units * cost_p
        if cogs_amt > 0:
            auto_journal_cogs(sale_date, cogs_amt, prod["name"], uid)
    except Exception:
        pass
    return uid, None


def add_stock_movement(move_date, product_id, movement_type, quantity, notes=""):
    prod = get_product_by_id(product_id)
    if not prod:
        return None, "Product not found"
    qty = float(quantity)
    if qty <= 0:
        return None, "Quantity must be greater than 0"
    before = float(prod["current_stock"])
    if movement_type in ("spoilage", "withdrawn", "adjustment_out"):
        after = before - qty
        if after < 0:
            return None, f"Cannot reduce stock below 0. Available: {before:.0f}"
    else:
        after = before + qty
    uid = f"STK-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO stock_movements
               (unique_id, move_date, product_id, product_name, movement_type,
                quantity, stock_before, stock_after, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (uid, move_date, product_id, prod["name"], movement_type, qty, before, after, notes),
        )
        cur.execute("UPDATE products SET current_stock = ? WHERE id = ?", (after, product_id))
        conn.commit()
        return uid, None
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()


def get_sales_df(start_date=None, end_date=None):
    conn = get_conn()
    query = "SELECT * FROM sales"
    params = []
    if start_date and end_date:
        query += " WHERE sale_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    query += " ORDER BY sale_date DESC, created_at DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def get_stock_movements_df(start_date=None, end_date=None):
    conn = get_conn()
    query = "SELECT * FROM stock_movements"
    params = []
    if start_date and end_date:
        query += " WHERE move_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    query += " ORDER BY move_date DESC, created_at DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def get_daily_totals():
    conn = get_conn()
    df = pd.read_sql_query(
        """SELECT sale_date AS date,
                  COUNT(*) AS transactions,
                  SUM(units_sold) AS total_units,
                  SUM(amount) AS total_amount
           FROM sales GROUP BY sale_date ORDER BY sale_date DESC""",
        conn,
    )
    conn.close()
    return df


def update_product_price(product_id, new_price, cost_price=None):
    conn = get_conn()
    cur = conn.cursor()
    if cost_price is None:
        cur.execute("UPDATE products SET unit_price = ? WHERE id = ?", (float(new_price), product_id))
    else:
        cur.execute(
            "UPDATE products SET unit_price = ?, cost_price = ? WHERE id = ?",
            (float(new_price), float(cost_price), product_id),
        )
    conn.commit()
    conn.close()
    return True, None


def update_product_cost(product_id, cost_price):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE products SET cost_price = ? WHERE id = ?", (float(cost_price), product_id))
    conn.commit()
    conn.close()
    return True, None



def add_new_product(name, sku, price, stock, cost_price=0):
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO products (name, sku, unit_price, cost_price, current_stock) VALUES (?, ?, ?, ?, ?)",
            (name, sku, price, float(cost_price or 0), stock),
        )
        conn.commit()
        return True, None
    except sqlite3.IntegrityError:
        return False, "Product name already exists"
    except Exception as e:
        return False, str(e)
    finally:
        conn.close()



def delete_sale(sale_id: int):
    """Delete a sale, restore stock, and reverse related journals (sale + COGS)."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT product_id, units_sold, unique_id FROM sales WHERE id = ?", (sale_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Sale not found"
    product_id, units = row["product_id"], float(row["units_sold"])
    sale_uid = row["unique_id"]
    try:
        cur.execute(
            "UPDATE products SET current_stock = current_stock + ? WHERE id = ?",
            (units, product_id),
        )
        cur.execute("DELETE FROM sales WHERE id = ?", (sale_id,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()
    try:
        reverse_journals_for_source("sale", sale_uid, reason="Sale deleted")
        reverse_journals_for_source("cogs", sale_uid, reason="Sale deleted")
    except Exception:
        pass
    return True, None


def update_sale(sale_id: int, new_units: float, notes: str = ""):
    """Update units sold on an existing sale (adjusts stock difference)."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT product_id, units_sold, unit_price, opening_stock FROM sales WHERE id = ?",
        (sale_id,),
    )
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Sale not found"
    product_id = row["product_id"]
    old_units = float(row["units_sold"])
    price = float(row["unit_price"])
    opening = float(row["opening_stock"])
    new_units = float(new_units)
    if new_units <= 0:
        conn.close()
        return False, "Units must be greater than 0"
    # Difference to apply to product stock
    diff = old_units - new_units  # positive means restore stock
    cur.execute("SELECT current_stock FROM products WHERE id = ?", (product_id,))
    current = float(cur.fetchone()["current_stock"])
    new_product_stock = current + diff
    if new_product_stock < 0:
        conn.close()
        return False, f"Not enough stock for this change. Available after restore: {current + old_units:.0f}"
    new_closing = opening - new_units
    new_amount = new_units * price
    try:
        cur.execute(
            """UPDATE sales SET units_sold = ?, amount = ?, closing_stock = ?, notes = ?
               WHERE id = ?""",
            (new_units, new_amount, new_closing, notes, sale_id),
        )
        cur.execute(
            "UPDATE products SET current_stock = ? WHERE id = ?",
            (new_product_stock, product_id),
        )
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def delete_stock_movement(move_id: int, force: bool = False):
    """Delete a stock movement and reverse its effect on product stock.
    If force=True and reverse would go negative, stock is set to 0 instead.
    """
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT product_id, product_name, movement_type, quantity FROM stock_movements WHERE id = ?",
        (move_id,),
    )
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Movement not found"
    product_id = row["product_id"]
    product_name = row["product_name"]
    mtype = row["movement_type"]
    qty = float(row["quantity"])
    # Reverse the movement
    if mtype in ("spoilage", "withdrawn", "adjustment_out"):
        delta = qty   # was removed → add back
        action = "add back"
    else:
        delta = -qty  # was added → remove
        action = "remove"
    cur.execute("SELECT current_stock FROM products WHERE id = ?", (product_id,))
    current = float(cur.fetchone()["current_stock"])
    new_stock = current + delta
    if new_stock < 0:
        if not force:
            conn.close()
            return False, (
                f"Cannot delete safely. Current stock of **{product_name}** is **{current:.0f}**. "
                f"Deleting this record would {action} **{qty:.0f}** and stock would become **{new_stock:.0f}**. "
                f"Either record more stock first, or use **Force Delete** (stock will be set to 0)."
            )
        new_stock = 0  # force: clamp at zero
    try:
        cur.execute(
            "UPDATE products SET current_stock = ? WHERE id = ?",
            (new_stock, product_id),
        )
        cur.execute("DELETE FROM stock_movements WHERE id = ?", (move_id,))
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()



def get_movement_by_id(move_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM stock_movements WHERE id = ?", (move_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def update_stock_movement(move_id: int, new_qty: float, new_type: str, new_date: str, notes: str = ""):
    """Edit a stock movement: reverse old effect, apply new effect."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT product_id, product_name, movement_type, quantity FROM stock_movements WHERE id = ?",
        (move_id,),
    )
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Movement not found"
    product_id = row["product_id"]
    product_name = row["product_name"]
    old_type = row["movement_type"]
    old_qty = float(row["quantity"])
    new_qty = float(new_qty)
    if new_qty <= 0:
        conn.close()
        return False, "Quantity must be greater than 0"

    remove_types = ("spoilage", "withdrawn", "adjustment_out")

    # Reverse old movement
    if old_type in remove_types:
        reverse_delta = old_qty   # add back
    else:
        reverse_delta = -old_qty  # remove

    # Apply new movement
    if new_type in remove_types:
        apply_delta = -new_qty    # remove
    else:
        apply_delta = new_qty     # add

    cur.execute("SELECT current_stock FROM products WHERE id = ?", (product_id,))
    current = float(cur.fetchone()["current_stock"])
    new_stock = current + reverse_delta + apply_delta
    if new_stock < 0:
        conn.close()
        return False, (
            f"Edit would make stock of **{product_name}** negative "
            f"(would become {new_stock:.0f}). Adjust quantity or type."
        )
    try:
        # Recalculate stock_before / stock_after relative to current after reverse
        stock_before = current + reverse_delta
        stock_after = new_stock
        cur.execute(
            """UPDATE stock_movements
               SET move_date = ?, movement_type = ?, quantity = ?,
                   stock_before = ?, stock_after = ?, notes = ?
               WHERE id = ?""",
            (new_date, new_type, new_qty, stock_before, stock_after, notes, move_id),
        )
        cur.execute(
            "UPDATE products SET current_stock = ? WHERE id = ?",
            (new_stock, product_id),
        )
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def clear_product_stock(product_id: int):
    """Set a single product stock to 0."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE products SET current_stock = 0 WHERE id = ?", (product_id,))
    conn.commit()
    conn.close()


def clear_all_movements():
    """Delete all stock movements. Does not change product stock levels."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM stock_movements")
    conn.commit()
    conn.close()


def get_sale_by_id(sale_id: int):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM sales WHERE id = ?", (sale_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None



def reset_all_stocks():
    """Set every product current_stock to 0. Does not delete sales history."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE products SET current_stock = 0")
    conn.commit()
    conn.close()



# ─── Special Orders ───

def reset_accounts(confirm_text=""):
    """Delete all journal entries and lines. Chart of accounts is kept.
    Pass confirm_text='RESET' to proceed.
    """
    if confirm_text.strip().upper() != "RESET":
        return False, "Type RESET to confirm"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM journal_lines")
        cur.execute("DELETE FROM journal_entries")
        try:
            cur.execute("DELETE FROM sqlite_sequence WHERE name IN ('journal_lines', 'journal_entries')")
        except Exception:
            pass
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()



def add_special_order(order_date, customer_name, product_description, quantity, unit_price, status="Pending", notes=""):
    qty = float(quantity)
    price = float(unit_price)
    if qty <= 0:
        return None, "Quantity must be greater than 0"
    if not product_description.strip():
        return None, "Product description is required"
    amount = qty * price
    uid = f"ORD-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO special_orders
               (unique_id, order_date, customer_name, product_description,
                quantity, unit_price, amount, status, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (uid, order_date, customer_name.strip(), product_description.strip(),
             qty, price, amount, status, notes),
        )
        conn.commit()
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()
    # Recognise revenue when order is completed/paid (not while still Pending)
    st_ok = (status or "").strip().lower() in ("completed", "paid", "delivered", "done")
    if amount > 0 and st_ok:
        try:
            auto_journal_special_order(
                order_date, amount, customer_name, product_description, uid, payment="Cash"
            )
        except Exception:
            pass
    return uid, None


def get_special_orders_df(start_date=None, end_date=None):
    conn = get_conn()
    query = "SELECT * FROM special_orders"
    params = []
    if start_date and end_date:
        query += " WHERE order_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    query += " ORDER BY order_date DESC, created_at DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def update_special_order(order_id, customer_name, product_description, quantity, unit_price, status, notes=""):
    qty = float(quantity)
    price = float(unit_price)
    amount = qty * price
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """UPDATE special_orders SET customer_name=?, product_description=?,
               quantity=?, unit_price=?, amount=?, status=?, notes=? WHERE id=?""",
            (customer_name, product_description, qty, price, amount, status, notes, order_id),
        )
        # fetch unique_id for journal
        cur.execute("SELECT unique_id, order_date FROM special_orders WHERE id = ?", (order_id,))
        row = cur.fetchone()
        conn.commit()
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()
    try:
        st_ok = (status or "").strip().lower() in ("completed", "paid", "delivered", "done")
        if row and amount > 0 and st_ok:
            uid = row["unique_id"]
            if not journal_exists_for_source("special_order", uid):
                auto_journal_special_order(
                    row["order_date"], amount, customer_name, product_description, uid, payment="Cash"
                )
    except Exception:
        pass
    return True, None


def delete_special_order(order_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM special_orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()


# ─── Punch Branch ───
def add_branch_record(record_date, product_id, qty_taken, qty_sold, unit_price=None, notes="", payment_method="Cash"):
    """Record products produced for Punch branch, sold at Punch, and remaining.
    Does NOT touch main branch stock — goods are produced specially for Punch.
    unit_price: selling price at Punch (can differ from main catalogue price).
    """
    prod = get_product_by_id(product_id)
    if not prod:
        return None, "Product not found"
    taken = float(qty_taken)
    sold = float(qty_sold)
    if taken < 0 or sold < 0:
        return None, "Quantities cannot be negative"
    if sold > taken:
        return None, "Sold cannot be more than quantity sent"
    remaining = taken - sold
    # Use Punch selling price if provided, else main catalogue price
    if unit_price is not None and float(unit_price) >= 0:
        price = float(unit_price)
    else:
        price = float(prod["unit_price"])
    amount = sold * price

    uid = f"BR-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO branch_daily
               (unique_id, record_date, product_id, product_name, qty_taken,
                qty_sold, qty_remaining, amount_sold, unit_price, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (uid, record_date, product_id, prod["name"], taken, sold, remaining, amount, price, notes),
        )
        conn.commit()
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()
    # Auto journal Punch sales only when amount sold > 0
    if amount > 0:
        try:
            auto_journal_punch_sale(record_date, amount, prod["name"], uid, payment=payment_method or "Cash")
        except Exception:
            pass
    return uid, None


def get_branch_df(start_date=None, end_date=None):
    conn = get_conn()
    query = "SELECT * FROM branch_daily"
    params = []
    if start_date and end_date:
        query += " WHERE record_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    query += " ORDER BY record_date DESC, created_at DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def update_branch_record(rec_id, qty_taken, qty_sold, unit_price=None, notes=""):
    """Update branch record including Punch selling price. Does not affect main stock."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM branch_daily WHERE id = ?", (rec_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Record not found"
    taken = float(qty_taken)
    sold = float(qty_sold)
    if sold > taken:
        conn.close()
        return False, "Sold cannot be more than quantity sent"
    remaining = taken - sold
    if unit_price is not None and float(unit_price) >= 0:
        price = float(unit_price)
    else:
        price = float(row["unit_price"])
    amount = sold * price
    try:
        cur.execute(
            """UPDATE branch_daily SET qty_taken=?, qty_sold=?, qty_remaining=?,
               amount_sold=?, unit_price=?, notes=? WHERE id=?""",
            (taken, sold, remaining, amount, price, notes, rec_id),
        )
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def delete_branch_record(rec_id, force=False):
    """Delete branch record. Does not affect main stock."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id FROM branch_daily WHERE id = ?", (rec_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Record not found"
    try:
        cur.execute("DELETE FROM branch_daily WHERE id = ?", (rec_id,))
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def clear_all_products():
    """Delete all products so the catalogue can be re-entered.
    Sales and stock history are kept (they store product names as text).
    """
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA foreign_keys = OFF")
        cur.execute("DELETE FROM products")
        try:
            cur.execute("DELETE FROM sqlite_sequence WHERE name = 'products'")
        except Exception:
            pass  # table may not exist
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        try:
            cur.execute("PRAGMA foreign_keys = ON")
        except Exception:
            pass
        conn.close()


def rename_product(product_id, new_name, new_sku=None):
    """Rename a product and update the name in related history tables."""
    new_name = (new_name or "").strip()
    if not new_name:
        return False, "Product name cannot be empty"
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT name FROM products WHERE id = ?", (product_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Product not found"
    old_name = row["name"]
    # Check duplicate name
    cur.execute("SELECT id FROM products WHERE name = ? AND id != ?", (new_name, product_id))
    if cur.fetchone():
        conn.close()
        return False, f"Another product already uses the name '{new_name}'"
    try:
        if new_sku is not None:
            cur.execute(
                "UPDATE products SET name = ?, sku = ? WHERE id = ?",
                (new_name, new_sku.strip(), product_id),
            )
        else:
            cur.execute("UPDATE products SET name = ? WHERE id = ?", (new_name, product_id))
        # Keep history readable with the corrected name
        cur.execute("UPDATE sales SET product_name = ? WHERE product_id = ?", (new_name, product_id))
        cur.execute("UPDATE stock_movements SET product_name = ? WHERE product_id = ?", (new_name, product_id))
        cur.execute("UPDATE branch_daily SET product_name = ? WHERE product_id = ?", (new_name, product_id))
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def delete_single_product(product_id):
    """Delete one product from the catalogue."""
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA foreign_keys = OFF")
        cur.execute("DELETE FROM products WHERE id = ?", (product_id,))
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        try:
            cur.execute("PRAGMA foreign_keys = ON")
        except Exception:
            pass
        conn.close()


def get_products_in_stock_first():
    """Products sorted: those with stock > 0 first, then by name."""
    conn = get_conn()
    df = pd.read_sql_query(
        """SELECT id, name, sku, unit_price, COALESCE(cost_price, 0) AS cost_price, current_stock FROM products
           ORDER BY CASE WHEN current_stock > 0 THEN 0 ELSE 1 END, current_stock DESC, name""",
        conn,
    )
    conn.close()
    return df



# ─── Raw Materials Inventory ───
def get_raw_materials_df():
    conn = get_conn()
    df = pd.read_sql_query(
        """SELECT id, name, unit, current_qty, reorder_level, unit_cost, notes
           FROM raw_materials ORDER BY name""",
        conn,
    )
    conn.close()
    return df


def get_raw_material_by_id(material_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM raw_materials WHERE id = ?", (material_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def add_raw_material(name, unit, reorder_level=0, unit_cost=0, opening_qty=0, notes=""):
    name = (name or "").strip()
    if not name:
        return False, "Name is required"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO raw_materials (name, unit, current_qty, reorder_level, unit_cost, notes)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (name, unit.strip() or "kg", float(opening_qty), float(reorder_level), float(unit_cost), notes),
        )
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def update_raw_material(material_id, name, unit, reorder_level, unit_cost, notes=""):
    name = (name or "").strip()
    if not name:
        return False, "Name is required"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """UPDATE raw_materials SET name=?, unit=?, reorder_level=?, unit_cost=?, notes=?
               WHERE id=?""",
            (name, unit.strip() or "kg", float(reorder_level), float(unit_cost), notes, material_id),
        )
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def delete_raw_material(material_id):
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA foreign_keys = OFF")
        cur.execute("DELETE FROM raw_material_movements WHERE material_id = ?", (material_id,))
        cur.execute("DELETE FROM raw_materials WHERE id = ?", (material_id,))
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        try:
            cur.execute("PRAGMA foreign_keys = ON")
        except Exception:
            pass
        conn.close()


def add_raw_movement(move_date, material_id, movement_type, quantity, unit_cost=None, total_amount=None, notes="", amount_paid=None, payment_method="Cash"):
    """
    movement_type:
      purchase / received  -> increase
      used_in_production   -> decrease
      waste_spoilage       -> decrease
      adjustment_in        -> increase
      adjustment_out       -> decrease

    For purchases: pass total_amount (₦ paid) and quantity → unit_cost is calculated automatically.
    Or pass unit_cost directly.
    """
    mat = get_raw_material_by_id(material_id)
    if not mat:
        return None, "Material not found"
    qty = float(quantity)
    if qty <= 0:
        return None, "Quantity must be greater than 0"
    before = float(mat["current_qty"])
    increase_types = ("purchase", "received", "adjustment_in")
    if movement_type in increase_types:
        after = before + qty
    else:
        after = before - qty
        if after < 0:
            return None, f"Not enough stock. Available: {before:.2f} {mat['unit']}"

    # Purchase cost for this batch (per unit)
    old_cost = float(mat["unit_cost"] or 0)
    if total_amount is not None and float(total_amount) > 0 and qty > 0:
        batch_cost = float(total_amount) / qty
        total = float(total_amount)
    elif unit_cost is not None and float(unit_cost) >= 0:
        batch_cost = float(unit_cost)
        total = qty * batch_cost if movement_type in increase_types else 0.0
    else:
        batch_cost = old_cost
        total = qty * batch_cost if movement_type in increase_types else 0.0

    # Inventory valuation: Weighted Average Cost (WAC)
    # New avg = (old_qty × old_cost + buy_qty × buy_cost) / (old_qty + buy_qty)
    if movement_type in ("purchase", "received") and batch_cost > 0:
        if before > 0 and old_cost > 0:
            new_avg = (before * old_cost + qty * batch_cost) / (before + qty)
        else:
            new_avg = batch_cost
        cost = batch_cost  # this movement records the batch purchase cost
        avg_to_store = new_avg
    else:
        cost = batch_cost if batch_cost > 0 else old_cost
        avg_to_store = old_cost
        # Usage/waste valued at current average cost
        if movement_type not in increase_types:
            total = qty * old_cost

    uid = f"RAW-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            "UPDATE raw_materials SET current_qty = ? WHERE id = ?",
            (after, material_id),
        )
        if movement_type in ("purchase", "received") and avg_to_store > 0:
            cur.execute(
                "UPDATE raw_materials SET unit_cost = ? WHERE id = ?",
                (avg_to_store, material_id),
            )
        cur.execute(
            """INSERT INTO raw_material_movements
               (unique_id, move_date, material_id, material_name, movement_type,
                quantity, qty_before, qty_after, unit_cost, total_cost, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (uid, move_date, material_id, mat["name"], movement_type,
             qty, before, after, cost, total, notes),
        )
        conn.commit()
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()

    # Auto journals
    try:
        if movement_type in ("purchase", "received") and total and float(total) > 0:
            paid = amount_paid if amount_paid is not None else float(total)
            auto_journal_raw_purchase(
                move_date,
                float(total),
                mat["name"],
                uid,
                payment=payment_method or "Cash",
                amount_paid=paid,
            )
        elif movement_type in ("used_in_production", "waste_spoilage"):
            usage_amt = qty * float(old_cost or 0)
            if usage_amt > 0:
                auto_journal_raw_usage(move_date, usage_amt, mat["name"], uid, movement_type)
    except Exception:
        pass
    return uid, None


def get_raw_movements_df(start_date=None, end_date=None):
    conn = get_conn()
    query = "SELECT * FROM raw_material_movements"
    params = []
    if start_date and end_date:
        query += " WHERE move_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    query += " ORDER BY move_date DESC, created_at DESC"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def delete_raw_movement(move_id, force=False):
    """Reverse stock effect, delete movement, reverse related journals."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM raw_material_movements WHERE id = ?", (move_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Movement not found"
    material_id = row["material_id"]
    mtype = row["movement_type"]
    qty = float(row["quantity"])
    move_uid = row["unique_id"]
    increase_types = ("purchase", "received", "adjustment_in")
    if mtype in increase_types:
        delta = -qty
    else:
        delta = qty
    cur.execute("SELECT current_qty, name FROM raw_materials WHERE id = ?", (material_id,))
    mat = cur.fetchone()
    if not mat:
        conn.close()
        return False, "Material not found"
    new_qty = float(mat["current_qty"]) + delta
    if new_qty < 0 and not force:
        conn.close()
        return False, f"Cannot reverse — {mat['name']} would go negative ({new_qty:.2f}). Use force delete."
    if new_qty < 0:
        new_qty = 0
    try:
        cur.execute("UPDATE raw_materials SET current_qty = ? WHERE id = ?", (new_qty, material_id))
        cur.execute("DELETE FROM raw_material_movements WHERE id = ?", (move_id,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()
    try:
        if mtype in ("purchase", "received"):
            reverse_journals_for_source("raw_purchase", move_uid, reason="Purchase deleted")
        elif mtype in ("used_in_production", "waste_spoilage"):
            reverse_journals_for_source("raw_usage", move_uid, reason="Usage deleted")
    except Exception:
        pass
    return True, None



def add_expense(expense_date, category, description, amount, branch="Main", payment_method="Cash", notes=""):
    amt = float(amount)
    if amt <= 0:
        return None, "Amount must be greater than 0"
    uid = f"EXP-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO expenses
               (unique_id, expense_date, category, description, amount, branch, payment_method, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (uid, expense_date, category, description, amt, branch, payment_method, notes),
        )
        conn.commit()
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()
    try:
        auto_journal_expense(expense_date, amt, category, description, uid, payment=payment_method)
    except Exception:
        pass
    return uid, None


def get_expenses_df(start_date=None, end_date=None, branch=None):
    conn = get_conn()
    q = "SELECT * FROM expenses WHERE 1=1"
    params = []
    if start_date and end_date:
        q += " AND expense_date BETWEEN ? AND ?"
        params.extend([start_date, end_date])
    if branch and branch != "All":
        q += " AND branch = ?"
        params.append(branch)
    q += " ORDER BY expense_date DESC, id DESC"
    df = pd.read_sql_query(q, conn, params=params)
    conn.close()
    return df


def delete_expense(exp_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT unique_id FROM expenses WHERE id = ?", (exp_id,))
    row = cur.fetchone()
    uid = row["unique_id"] if row else None
    try:
        cur.execute("DELETE FROM expenses WHERE id = ?", (exp_id,))
        conn.commit()
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()
    if uid:
        try:
            reverse_journals_for_source("expense", uid, reason="Expense deleted")
        except Exception:
            pass
    return True, None


# ─── Staff ───
def add_staff(full_name, role_title, branch, phone, monthly_salary, start_date, notes=""):
    name = (full_name or "").strip()
    if not name:
        return False, "Name is required"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO staff
               (full_name, role_title, branch, phone, monthly_salary, start_date, is_active, notes)
               VALUES (?, ?, ?, ?, ?, ?, 1, ?)""",
            (name, role_title, branch, phone, float(monthly_salary), start_date, notes),
        )
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def update_staff(staff_id, full_name, role_title, branch, phone, monthly_salary, start_date, is_active, notes=""):
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """UPDATE staff SET full_name=?, role_title=?, branch=?, phone=?,
               monthly_salary=?, start_date=?, is_active=?, notes=? WHERE id=?""",
            (full_name.strip(), role_title, branch, phone, float(monthly_salary),
             start_date, 1 if is_active else 0, notes, staff_id),
        )
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def get_staff_df(active_only=False):
    conn = get_conn()
    q = "SELECT * FROM staff"
    if active_only:
        q += " WHERE is_active = 1"
    q += " ORDER BY full_name"
    df = pd.read_sql_query(q, conn)
    conn.close()
    return df


def add_salary_payment(staff_id, pay_date, amount, period_label="", notes="", payment_method="Cash"):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT full_name FROM staff WHERE id = ?", (staff_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return None, "Staff not found"
    amt = float(amount)
    if amt <= 0:
        conn.close()
        return None, "Amount must be greater than 0"
    uid = f"SAL-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    try:
        cur.execute(
            """INSERT INTO salary_payments
               (unique_id, staff_id, staff_name, pay_date, period_label, amount, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (uid, staff_id, row["full_name"], pay_date, period_label, amt, notes),
        )
        conn.commit()
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()
    try:
        auto_journal_salary(pay_date, amt, row["full_name"], uid, payment=payment_method)
    except Exception:
        pass
    return uid, None


def get_salary_payments_df(start_date=None, end_date=None):
    conn = get_conn()
    q = "SELECT * FROM salary_payments"
    params = []
    if start_date and end_date:
        q += " WHERE pay_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    q += " ORDER BY pay_date DESC, id DESC"
    df = pd.read_sql_query(q, conn, params=params)
    conn.close()
    return df


def get_branch_profit_summary(start_date=None, end_date=None):
    """Revenue by branch, expenses by branch, approximate profit (revenue - expenses).
    Note: COGS for finished goods is not fully deducted unless tracked separately.
    """
    sales = get_sales_df()
    if start_date and end_date and not sales.empty:
        sales = sales[(sales["sale_date"] >= start_date) & (sales["sale_date"] <= end_date)]
    main_rev = float(sales["amount"].sum()) if not sales.empty else 0.0

    conn = get_conn()
    bq = "SELECT COALESCE(SUM(amount_sold), 0) AS t FROM branch_daily"
    params = []
    if start_date and end_date:
        bq += " WHERE record_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    punch_rev = float(pd.read_sql_query(bq, conn, params=params).iloc[0]["t"])

    # special orders as main revenue if any amount
    oq = "SELECT COALESCE(SUM(amount), 0) AS t FROM special_orders"
    oparams = []
    if start_date and end_date:
        oq += " WHERE order_date BETWEEN ? AND ?"
        oparams = [start_date, end_date]
    orders_rev = float(pd.read_sql_query(oq, conn, params=oparams).iloc[0]["t"])
    conn.close()

    main_rev = main_rev + orders_rev

    exp = get_expenses_df(start_date, end_date)
    sal = get_salary_payments_df(start_date, end_date)
    main_exp = float(exp[exp["branch"] == "Main"]["amount"].sum()) if not exp.empty else 0.0
    punch_exp = float(exp[exp["branch"] == "Punch"]["amount"].sum()) if not exp.empty else 0.0
    # shared/other expenses split or assigned to Main
    other_exp = float(exp[~exp["branch"].isin(["Main", "Punch"])]["amount"].sum()) if not exp.empty else 0.0
    main_exp += other_exp
    salary_total = float(sal["amount"].sum()) if not sal.empty else 0.0
    # salaries attributed to staff branch if possible — simple: all under Main for now unless we join
    staff = get_staff_df()
    # allocate salary payments proportionally is hard; attribute salary_total to Main by default
    # Better: when paying salary we could store branch — for now add salary to total expenses only

    return {
        "main_revenue": main_rev,
        "punch_revenue": punch_rev,
        "main_expenses": main_exp,
        "punch_expenses": punch_exp,
        "salary_paid": salary_total,
        "total_revenue": main_rev + punch_rev,
        "total_expenses": main_exp + punch_exp + salary_total,
        "main_profit": main_rev - main_exp,  # before shared salary allocation
        "punch_profit": punch_rev - punch_exp,
        "net_profit": main_rev + punch_rev - main_exp - punch_exp - salary_total,
    }



# ─── Accounting Phase 1: Chart of accounts + journals ───
def get_accounts_df(active_only=True):
    conn = get_conn()
    q = "SELECT * FROM accounts"
    if active_only:
        q += " WHERE is_active = 1"
    q += " ORDER BY code"
    df = pd.read_sql_query(q, conn)
    conn.close()
    return df


def get_account_by_code(code: str):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM accounts WHERE code = ? AND is_active = 1", (code,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def post_journal(entry_date, description, lines, source_type="", source_id=""):
    """
    lines: list of dicts {account_code, debit, credit}
    Debits must equal credits.
    """
    if not lines:
        return None, "No journal lines"
    total_dr = sum(float(x.get("debit", 0) or 0) for x in lines)
    total_cr = sum(float(x.get("credit", 0) or 0) for x in lines)
    if abs(total_dr - total_cr) > 0.01:
        return None, f"Journal not balanced: Dr {total_dr:.2f} vs Cr {total_cr:.2f}"
    if total_dr <= 0:
        return None, "Journal amount must be greater than 0"

    resolved = []
    for ln in lines:
        code = str(ln["account_code"])
        acc = get_account_by_code(code)
        if not acc:
            return None, f"Account {code} not found"
        dr = float(ln.get("debit", 0) or 0)
        cr = float(ln.get("credit", 0) or 0)
        if dr < 0 or cr < 0:
            return None, "Debit/credit cannot be negative"
        if dr > 0 and cr > 0:
            return None, "A line cannot have both debit and credit"
        resolved.append((code, acc["name"], dr, cr))

    uid = f"JRN-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            """INSERT INTO journal_entries
               (unique_id, entry_date, description, source_type, source_id)
               VALUES (?, ?, ?, ?, ?)""",
            (uid, entry_date, description, source_type, source_id),
        )
        jid = cur.lastrowid
        for code, name, dr, cr in resolved:
            cur.execute(
                """INSERT INTO journal_lines
                   (journal_id, account_code, account_name, debit, credit)
                   VALUES (?, ?, ?, ?, ?)""",
                (jid, code, name, dr, cr),
            )
        conn.commit()
        return uid, None
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()


def get_journals_df(start_date=None, end_date=None):
    conn = get_conn()
    q = "SELECT * FROM journal_entries"
    params = []
    if start_date and end_date:
        q += " WHERE entry_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    q += " ORDER BY entry_date DESC, id DESC"
    df = pd.read_sql_query(q, conn, params=params)
    conn.close()
    return df


def get_journal_lines_df(journal_id=None):
    conn = get_conn()
    if journal_id:
        df = pd.read_sql_query(
            "SELECT * FROM journal_lines WHERE journal_id = ? ORDER BY id",
            conn,
            params=[journal_id],
        )
    else:
        df = pd.read_sql_query(
            """SELECT jl.*, je.unique_id, je.entry_date, je.description, je.source_type
               FROM journal_lines jl
               JOIN journal_entries je ON je.id = jl.journal_id
               ORDER BY je.entry_date DESC, jl.id""",
            conn,
        )
    conn.close()
    return df


def auto_journal_sale(sale_date, amount, product_name, sale_uid, payment="Cash"):
    """Dr Cash/Bank, Cr Sales - Main"""
    cash_code = payment_account_code(payment)
    lines = [
        {"account_code": cash_code, "debit": amount, "credit": 0},
        {"account_code": "4000", "debit": 0, "credit": amount},
    ]
    return post_journal(
        sale_date,
        f"Sale: {product_name}",
        lines,
        source_type="sale",
        source_id=sale_uid,
    )


def auto_journal_raw_purchase(
    move_date, total_amount, material_name, move_uid, payment="Cash", amount_paid=None
):
    """Dr Inventory; Cr Cash (paid now) and/or Accounts payable (on credit)."""
    total_amount = float(total_amount)
    if total_amount <= 0:
        return None, "No amount to post"
    if amount_paid is None:
        paid = total_amount
    else:
        paid = max(0.0, min(float(amount_paid), total_amount))
    on_credit = total_amount - paid
    cash_code = payment_account_code(payment)
    lines = [{"account_code": "1200", "debit": total_amount, "credit": 0}]
    if paid > 0:
        lines.append({"account_code": cash_code, "debit": 0, "credit": paid})
    if on_credit > 0:
        lines.append({"account_code": "2000", "debit": 0, "credit": on_credit})
    return post_journal(
        move_date,
        f"Purchase: {material_name}"
        + (f" (paid ₦{paid:,.0f}, credit ₦{on_credit:,.0f})" if on_credit > 0 else ""),
        lines,
        source_type="raw_purchase",
        source_id=move_uid,
    )


def auto_journal_supplier_payment(pay_date, amount, supplier_name, pay_uid, payment="Cash"):
    """Pay what is owed: Dr Accounts payable, Cr Cash/Bank."""
    amount = float(amount)
    if amount <= 0:
        return None, "Amount must be greater than 0"
    cash_code = payment_account_code(payment)
    narr = "Supplier payment" + (f": {supplier_name}" if supplier_name else "")
    lines = [
        {"account_code": "2000", "debit": amount, "credit": 0},
        {"account_code": cash_code, "debit": 0, "credit": amount},
    ]
    return post_journal(pay_date, narr, lines, source_type="supplier_payment", source_id=pay_uid)


def auto_journal_loan_received(recv_date, amount, lender, loan_uid, payment="Cash"):
    """Dr Cash/Bank, Cr Loans payable."""
    amount = float(amount)
    if amount <= 0:
        return None, "Amount must be greater than 0"
    cash_code = payment_account_code(payment)
    narr = "Loan received" + (f": {lender}" if lender else "")
    lines = [
        {"account_code": cash_code, "debit": amount, "credit": 0},
        {"account_code": "2100", "debit": 0, "credit": amount},
    ]
    return post_journal(recv_date, narr, lines, source_type="loan_received", source_id=loan_uid)


def auto_journal_loan_repay(pay_date, principal, interest, lender, loan_uid, payment="Cash"):
    """Dr Loans payable (+ interest expense), Cr Cash/Bank."""
    principal = float(principal or 0)
    interest = float(interest or 0)
    total = principal + interest
    if total <= 0:
        return None, "Amount must be greater than 0"
    cash_code = payment_account_code(payment)
    narr = "Loan repayment" + (f": {lender}" if lender else "")
    lines = []
    if principal > 0:
        lines.append({"account_code": "2100", "debit": principal, "credit": 0})
    if interest > 0:
        lines.append({"account_code": "5400", "debit": interest, "credit": 0})
    lines.append({"account_code": cash_code, "debit": 0, "credit": total})
    return post_journal(pay_date, narr, lines, source_type="loan_repay", source_id=loan_uid)


def payment_account_code(payment="Cash"):
    """Map payment method to Cash (1000) or Bank/Transfer (1100)."""
    p = (payment or "Cash").strip().lower()
    if p in ("transfer", "bank", "pos", "card"):
        return "1100"
    return "1000"



def journal_exists_for_source(source_type, source_id):
    if not source_id:
        return False
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT id FROM journal_entries WHERE source_type = ? AND source_id = ? LIMIT 1",
        (source_type, source_id),
    )
    row = cur.fetchone()
    conn.close()
    return row is not None


def reverse_journals_for_source(source_type, source_id, reason="Void"):
    """Post reversing entries for all journals linked to a source, then mark originals.
    Creates equal-and-opposite lines so trial balance stays clean.
    """
    if not source_id:
        return 0, "No source id"
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """SELECT id, unique_id, entry_date, description FROM journal_entries
           WHERE source_type = ? AND source_id = ?
           AND (description IS NULL OR description NOT LIKE 'REV:%')""",
        (source_type, source_id),
    )
    entries = cur.fetchall()
    conn.close()
    if not entries:
        return 0, None
    count = 0
    for ent in entries:
        lines_df = get_journal_lines_df(ent["id"])
        if lines_df is None or lines_df.empty:
            continue
        rev_lines = []
        for _, ln in lines_df.iterrows():
            dr = float(ln["debit"] or 0)
            cr = float(ln["credit"] or 0)
            # swap
            rev_lines.append({
                "account_code": ln["account_code"],
                "debit": cr,
                "credit": dr,
            })
        # filter zero lines
        rev_lines = [x for x in rev_lines if x["debit"] > 0 or x["credit"] > 0]
        if not rev_lines:
            continue
        desc = f"REV: {reason} — {ent['description'] or ent['unique_id']}"
        juid, err = post_journal(
            ent["entry_date"],
            desc,
            rev_lines,
            source_type=f"rev_{source_type}",
            source_id=source_id,
        )
        if not err:
            count += 1
    return count, None


def auto_journal_cogs(sale_date, cogs_amount, product_name, sale_uid):
    """Dr Cost of goods sold, Cr Finished goods inventory."""
    cogs_amount = float(cogs_amount or 0)
    if cogs_amount <= 0:
        return None, "No COGS"
    lines = [
        {"account_code": "5050", "debit": cogs_amount, "credit": 0},
        {"account_code": "1250", "debit": 0, "credit": cogs_amount},
    ]
    return post_journal(
        sale_date,
        f"COGS: {product_name}",
        lines,
        source_type="cogs",
        source_id=sale_uid,
    )


def post_opening_balances(entry_date, balances: dict, notes=""):
    """
    balances keys: cash, bank, raw_inventory, finished_goods, payables, loans
    All amounts >= 0. Posts one balanced journal vs Opening balance equity (3000).
    """
    cash = float(balances.get("cash") or 0)
    bank = float(balances.get("bank") or 0)
    raw_inv = float(balances.get("raw_inventory") or 0)
    fg = float(balances.get("finished_goods") or 0)
    payables = float(balances.get("payables") or 0)
    loans = float(balances.get("loans") or 0)

    lines = []
    if cash > 0:
        lines.append({"account_code": "1000", "debit": cash, "credit": 0})
    if bank > 0:
        lines.append({"account_code": "1100", "debit": bank, "credit": 0})
    if raw_inv > 0:
        lines.append({"account_code": "1200", "debit": raw_inv, "credit": 0})
    if fg > 0:
        lines.append({"account_code": "1250", "debit": fg, "credit": 0})
    if payables > 0:
        lines.append({"account_code": "2000", "debit": 0, "credit": payables})
    if loans > 0:
        lines.append({"account_code": "2100", "debit": 0, "credit": loans})

    total_dr = sum(x["debit"] for x in lines)
    total_cr = sum(x["credit"] for x in lines)
    equity = total_dr - total_cr
    if abs(equity) < 0.01 and total_dr == 0:
        return None, "Enter at least one opening amount"
    if equity > 0:
        lines.append({"account_code": "3000", "debit": 0, "credit": equity})
    elif equity < 0:
        lines.append({"account_code": "3000", "debit": abs(equity), "credit": 0})

    uid = f"OB-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    desc = "Opening balances" + (f" — {notes}" if notes else "")
    return post_journal(entry_date, desc, lines, source_type="opening_balance", source_id=uid)


def get_cogs_summary(start_date=None, end_date=None):
    """Sum COGS from journals and sales revenue for gross profit."""
    conn = get_conn()
    q_cogs = """
        SELECT COALESCE(SUM(jl.debit), 0) AS cogs
        FROM journal_lines jl
        JOIN journal_entries je ON je.id = jl.journal_id
        WHERE jl.account_code = '5050'
    """
    q_sales = """
        SELECT COALESCE(SUM(jl.credit - jl.debit), 0) AS sales
        FROM journal_lines jl
        JOIN journal_entries je ON je.id = jl.journal_id
        WHERE jl.account_code IN ('4000', '4100', '4200')
    """
    params = []
    if start_date and end_date:
        q_cogs += " AND je.entry_date BETWEEN ? AND ?"
        q_sales += " AND je.entry_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    cogs = float(pd.read_sql_query(q_cogs, conn, params=params).iloc[0]["cogs"])
    sales = float(pd.read_sql_query(q_sales, conn, params=params).iloc[0]["sales"])
    conn.close()
    return {"sales": sales, "cogs": cogs, "gross_profit": sales - cogs}



def auto_journal_expense(expense_date, amount, category, description, exp_uid, payment="Cash"):
    """Dr Operating expenses, Cr Cash/Bank"""
    if amount <= 0:
        return None, "No amount"
    cash_code = payment_account_code(payment)
    narr = f"Expense: {category}" + (f" — {description}" if description else "")
    lines = [
        {"account_code": "5200", "debit": amount, "credit": 0},
        {"account_code": cash_code, "debit": 0, "credit": amount},
    ]
    return post_journal(expense_date, narr, lines, source_type="expense", source_id=exp_uid)


def auto_journal_salary(pay_date, amount, staff_name, sal_uid, payment="Cash"):
    """Dr Salaries, Cr Cash/Bank"""
    if amount <= 0:
        return None, "No amount"
    cash_code = payment_account_code(payment)
    lines = [
        {"account_code": "5300", "debit": amount, "credit": 0},
        {"account_code": cash_code, "debit": 0, "credit": amount},
    ]
    return post_journal(
        pay_date,
        f"Salary: {staff_name}",
        lines,
        source_type="salary",
        source_id=sal_uid,
    )


def auto_journal_punch_sale(record_date, amount, product_name, br_uid, payment="Cash"):
    """Dr Cash/Bank, Cr Sales - Punch"""
    if amount <= 0:
        return None, "No amount"
    cash_code = payment_account_code(payment)
    lines = [
        {"account_code": cash_code, "debit": amount, "credit": 0},
        {"account_code": "4100", "debit": 0, "credit": amount},
    ]
    return post_journal(
        record_date,
        f"Punch sale: {product_name}",
        lines,
        source_type="punch_sale",
        source_id=br_uid,
    )


def auto_journal_special_order(order_date, amount, customer, desc, ord_uid, payment="Cash"):
    """Dr Cash/Bank, Cr Sales - Special orders"""
    if amount <= 0:
        return None, "No amount"
    cash_code = payment_account_code(payment)
    narr = f"Special order: {desc}" + (f" ({customer})" if customer else "")
    lines = [
        {"account_code": cash_code, "debit": amount, "credit": 0},
        {"account_code": "4200", "debit": 0, "credit": amount},
    ]
    return post_journal(
        order_date,
        narr,
        lines,
        source_type="special_order",
        source_id=ord_uid,
    )


def auto_journal_raw_usage(move_date, amount, material_name, move_uid, movement_type):
    """Usage or waste: Dr COGS or Spoilage, Cr Raw materials inventory"""
    if amount <= 0:
        return None, "No amount"
    expense_code = "5100" if movement_type == "waste_spoilage" else "5000"
    label = "Waste/spoilage" if movement_type == "waste_spoilage" else "Materials used"
    lines = [
        {"account_code": expense_code, "debit": amount, "credit": 0},
        {"account_code": "1200", "debit": 0, "credit": amount},
    ]
    return post_journal(
        move_date,
        f"{label}: {material_name}",
        lines,
        source_type="raw_usage",
        source_id=move_uid,
    )


def get_ledger_for_account(account_code, start_date=None, end_date=None):
    """Return journal lines for one account with running balance.
    Balance sign: Assets/Expenses increase with debit; Income increases with credit.
    """
    conn = get_conn()
    q = """
        SELECT je.entry_date, je.unique_id, je.description, je.source_type, je.source_id,
               jl.account_code, jl.account_name, jl.debit, jl.credit
        FROM journal_lines jl
        JOIN journal_entries je ON je.id = jl.journal_id
        WHERE jl.account_code = ?
    """
    params = [account_code]
    if start_date and end_date:
        q += " AND je.entry_date BETWEEN ? AND ?"
        params.extend([start_date, end_date])
    q += " ORDER BY je.entry_date ASC, je.id ASC, jl.id ASC"
    df = pd.read_sql_query(q, conn, params=params)
    acc = pd.read_sql_query(
        "SELECT account_type FROM accounts WHERE code = ?", conn, params=[account_code]
    )
    conn.close()
    if df.empty:
        return df
    atype = acc.iloc[0]["account_type"] if not acc.empty else "Asset"
    # Normal debit balance: Asset, Expense
    # Normal credit balance: Income (and Liability/Equity if added later)
    normal_debit = atype in ("Asset", "Expense")  # Liability, Equity, Income = credit normal
    running = 0.0
    balances = []
    for _, r in df.iterrows():
        dr = float(r["debit"] or 0)
        cr = float(r["credit"] or 0)
        if normal_debit:
            running += dr - cr
        else:
            running += cr - dr
        balances.append(running)
    df = df.copy()
    df["balance"] = balances
    return df


def get_trial_balance(start_date=None, end_date=None):
    """Sum debits and credits per account for the period. Only accounts with activity."""
    conn = get_conn()
    q = """
        SELECT jl.account_code, jl.account_name,
               COALESCE(SUM(jl.debit), 0) AS total_debit,
               COALESCE(SUM(jl.credit), 0) AS total_credit
        FROM journal_lines jl
        JOIN journal_entries je ON je.id = jl.journal_id
    """
    params = []
    if start_date and end_date:
        q += " WHERE je.entry_date BETWEEN ? AND ?"
        params = [start_date, end_date]
    q += " GROUP BY jl.account_code, jl.account_name ORDER BY jl.account_code"
    df = pd.read_sql_query(q, conn, params=params)
    # join type
    acc = pd.read_sql_query("SELECT code, account_type FROM accounts", conn)
    conn.close()
    if df.empty:
        return df
    df = df.merge(acc, left_on="account_code", right_on="code", how="left")
    df = df.drop(columns=["code"], errors="ignore")
    # Closing balance column (debit-nature positive if more debits)
    df["balance_debit"] = 0.0
    df["balance_credit"] = 0.0
    for i, r in df.iterrows():
        net = float(r["total_debit"]) - float(r["total_credit"])
        atype = r.get("account_type") or "Asset"
        if atype in ("Asset", "Expense"):
            if net >= 0:
                df.at[i, "balance_debit"] = net
            else:
                df.at[i, "balance_credit"] = abs(net)
        else:
            # Income, Liability, Equity (credit normal)
            if net <= 0:
                df.at[i, "balance_credit"] = abs(net)
            else:
                df.at[i, "balance_debit"] = net
    return df


def to_excel_bytes(dfs: dict) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    return output.getvalue()


# ─────────────────────────────────────────────
# PAGE SETUP
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Funlex Cakes and Cream",
    page_icon="🧁",
    layout="wide",
    initial_sidebar_state="expanded",
)

init_db()

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }

    .block-container {
        padding-top: 1.75rem;
        padding-bottom: 2.5rem;
        max-width: 1100px;
    }

    /* Headers */
    h1, h2, h3 {
        letter-spacing: -0.02em !important;
        color: #1C1C28 !important;
    }
    h1 { font-weight: 800 !important; }
    h2 { font-weight: 700 !important; font-size: 1.45rem !important; }
    h3 { font-weight: 600 !important; }

    /* Metrics as soft cards */
    div[data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #EEE8EF;
        border-radius: 14px;
        padding: 1rem 1.15rem;
        box-shadow: 0 1px 3px rgba(28, 28, 40, 0.04);
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.4rem !important;
        font-weight: 700 !important;
        color: #C2185B !important;
    }
    div[data-testid="stMetricLabel"] {
        font-size: 0.8rem !important;
        color: #6B6B7B !important;
        font-weight: 500 !important;
    }

    /* Buttons */
    .stButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 0.45rem 1.1rem !important;
        border: none !important;
        transition: all 0.15s ease !important;
    }
    .stButton > button[kind="primary"],
    .stButton > button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #C2185B 0%, #AD1457 100%) !important;
        color: #fff !important;
        box-shadow: 0 2px 8px rgba(194, 24, 91, 0.25) !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(194, 24, 91, 0.3) !important;
    }

    /* Inputs */
    .stTextInput input, .stNumberInput input, .stDateInput input,
    .stSelectbox > div > div, .stTextArea textarea {
        border-radius: 10px !important;
        border-color: #E8E4EA !important;
    }
    .stTextInput input:focus, .stNumberInput input:focus {
        border-color: #C2185B !important;
        box-shadow: 0 0 0 2px rgba(194, 24, 91, 0.15) !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.35rem;
        background: #F3EEF2;
        padding: 0.35rem;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 9px !important;
        font-weight: 600 !important;
        padding: 0.45rem 1rem !important;
    }
    .stTabs [aria-selected="true"] {
        background: #FFFFFF !important;
        color: #C2185B !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }

    /* Dataframes */
    div[data-testid="stDataFrame"] {
        border: 1px solid #EEE8EF;
        border-radius: 12px;
        overflow: hidden;
    }

    /* Alerts */
    div[data-testid="stAlert"] {
        border-radius: 12px !important;
    }

    /* Dividers */
    hr {
        margin: 1.25rem 0 !important;
        border-color: #EEE8EF !important;
    }

    /* Forms */
    div[data-testid="stForm"] {
        background: #FFFFFF;
        border: 1px solid #EEE8EF;
        border-radius: 16px;
        padding: 1.25rem 1.4rem 1rem;
        box-shadow: 0 1px 4px rgba(28, 28, 40, 0.04);
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #2A1528 0%, #1F101E 100%) !important;
        border-right: 1px solid #3D2438 !important;
    }
    section[data-testid="stSidebar"] > div {
        padding-top: 1.25rem;
    }
    section[data-testid="stSidebar"] * {
        color: #F5F0F3 !important;
    }
    section[data-testid="stSidebar"] .stRadio label {
        color: #F5F0F3 !important;
        font-weight: 500 !important;
        padding: 0.35rem 0.6rem !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] .stRadio label:hover {
        background: rgba(194, 24, 91, 0.2) !important;
        color: #FFD6E7 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        color: #F8BBD0 !important;
    }
    section[data-testid="stSidebar"] button {
        background: linear-gradient(135deg, #C2185B, #AD1457) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
    }
    section[data-testid="stSidebar"] button:hover {
        background: #AD1457 !important;
    }
    section[data-testid="stSidebar"] [role="radiogroup"] label[data-checked="true"] {
        background: rgba(194, 24, 91, 0.35) !important;
        border-radius: 8px;
    }
    section[data-testid="stSidebar"] hr {
        border-color: #4A2F45 !important;
    }

    /* Caption */
    .stCaption, small {
        color: #6B6B7B !important;
    }

    /* Page title accent bar */
    .fx-page-title {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        margin-bottom: 0.35rem;
    }
    .fx-page-title h2 {
        margin: 0 !important;
    }
    .fx-accent {
        width: 4px;
        height: 1.4rem;
        background: linear-gradient(180deg, #C2185B, #E91E8C);
        border-radius: 4px;
        display: inline-block;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "role" not in st.session_state:
    st.session_state.role = None


def show_logo(width=140):
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=width)


# ─────────────────────────────────────────────
# LOGIN
# ─────────────────────────────────────────────
def login_page():
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        show_logo(160)
        st.markdown(
            """
            <div style="text-align:center; margin-top:0.4rem; margin-bottom:1.2rem;">
                <h1 style="color:#C2185B; font-size:1.6rem; font-weight:800; margin:0.3rem 0 0.15rem 0;
                           letter-spacing:0.4px;">FUNLEX CAKES AND CREAM</h1>
                <p style="color:#AD1457; font-size:0.9rem; font-style:italic; margin:0 0 0.3rem 0;
                          letter-spacing:0.6px; font-weight:500;">The Taste Is The Difference</p>
                <p style="color:#8A8A9A; font-size:0.8rem; margin:0; font-weight:500;">
                    Sales &amp; Stock Management
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        code = st.text_input("Access Code", type="password", placeholder="Enter access code")
        if st.button("Sign in", type="primary", use_container_width=True):
            role = verify_login(code)
            if role:
                st.session_state.authenticated = True
                st.session_state.role = role
                st.rerun()
            else:
                st.error("Invalid access code. Please try again.")


if not st.session_state.authenticated:
    login_page()
    st.stop()


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
role = st.session_state.get("role") or "sales"
is_admin = role == "admin"

with st.sidebar:
    sc1, sc2, sc3 = st.columns([0.5, 2, 0.5])
    with sc2:
        show_logo(110)
    st.markdown(
        """
        <div style="text-align:center; margin-top:0.15rem; margin-bottom:0.5rem;">
            <p style="color:#FFD6E7; font-weight:800; font-size:0.88rem; margin:0.2rem 0 0.1rem 0;
                      letter-spacing:0.4px; line-height:1.25;">FUNLEX CAKES<br/>AND CREAM</p>
            <p style="color:#E91E8C; font-style:italic; font-size:0.68rem; margin:0.15rem 0 0 0;
                      letter-spacing:0.4px;">The Taste Is The Difference</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    role_label = "Admin" if is_admin else "Sales Rep"
    st.markdown(
        f"""<div style="text-align:center; margin:0.5rem 0 0.25rem 0;">
            <span style="background:rgba(194,24,91,0.35); color:#FFD6E7; font-size:0.72rem;
                         font-weight:600; padding:0.25rem 0.7rem; border-radius:20px;">
                {role_label}
            </span>
        </div>""",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    if is_admin:
        menu = [
            "Dashboard",
            "View Stock",
            "Record Sale",
            "Stock Tracker",
            "Raw Materials",
            "Punch Branch",
            "Special Orders",
            "Expenses",
            "Staff & Salaries",
            "Accounts",
            "Daily Totals",
            "Products",
            "Export Excel",
            "Settings",
        ]
    else:
        # Sales rep: view stock + record sales (cannot change stock)
        menu = [
            "View Stock",
            "Record Sale",
        ]

    page = st.radio("Navigation", menu, label_visibility="collapsed")
    st.markdown("---")
    if st.button("Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.role = None
        st.rerun()
    st.caption("© 2026 Funlex Cakes and Cream")

# Safety: sales rep cannot open restricted pages
if not is_admin and page not in ("Record Sale", "View Stock"):
    page = "Record Sale"


# ─────────────────────────────────────────────
# DASHBOARD
# ─────────────────────────────────────────────
if page == "Dashboard":
    st.header("Company Dashboard")
    st.caption(
        "Profit = revenue − expenses − salaries paid. "
        "Finished-goods cost of goods is not fully deducted unless tracked separately."
    )

    today = date.today()
    p1, p2 = st.columns([1, 2])
    with p1:
        period = st.selectbox(
            "Period",
            ["Today", "This month", "Last 30 days", "All time"],
            key="dash_period",
        )
    if period == "Today":
        d0 = d1 = today.isoformat()
    elif period == "This month":
        d0 = today.replace(day=1).isoformat()
        d1 = today.isoformat()
    elif period == "Last 30 days":
        from datetime import timedelta
        d0 = (today - timedelta(days=30)).isoformat()
        d1 = today.isoformat()
    else:
        d0 = d1 = None

    summary = get_branch_profit_summary(d0, d1)
    products = get_all_products()
    sales = get_sales_df()
    staff = get_staff_df(active_only=True)
    low_stock = products[products["current_stock"] <= 5] if not products.empty else pd.DataFrame()
    try:
        raws = get_raw_materials_df()
        low_raw_df = raws[raws["current_qty"] <= raws["reorder_level"]] if not raws.empty else pd.DataFrame()
        low_raw = len(low_raw_df)
        inv_val = float((raws["current_qty"] * raws["unit_cost"]).sum()) if not raws.empty else 0.0
    except Exception:
        low_raw_df = pd.DataFrame()
        low_raw, inv_val = 0, 0.0

    # ── At a glance (top) ──
    st.subheader("At a glance")
    g1, g2, g3, g4, g5 = st.columns(5)
    today_str = today.isoformat()
    today_sales = sales[sales["sale_date"] == today_str] if not sales.empty else pd.DataFrame()
    g1.metric(
        "Today's sales (Main)",
        f"₦{today_sales['amount'].sum():,.0f}" if not today_sales.empty else "₦0",
    )
    g2.metric("Active staff", len(staff))
    g3.metric(
        "Monthly payroll (active)",
        f"₦{staff['monthly_salary'].sum():,.0f}" if not staff.empty else "₦0",
    )
    g4.metric("Low product stock", len(low_stock))
    g5.metric("Raw materials value", f"₦{inv_val:,.0f}")

    # Alerts (no tables)
    if low_raw:
        names = ", ".join(low_raw_df["name"].astype(str).head(8).tolist())
        more = f" (+{low_raw - 8} more)" if low_raw > 8 else ""
        st.warning(
            f"**{low_raw} raw material(s) need reorder:** {names}{more} — open **Raw Materials**."
        )
    if not low_stock.empty:
        names = ", ".join(low_stock["name"].astype(str).head(8).tolist())
        more = f" (+{len(low_stock) - 8} more)" if len(low_stock) > 8 else ""
        st.warning(
            f"**{len(low_stock)} finished product(s) low/out of stock:** {names}{more} — open **View Stock**."
        )

    st.divider()

    # ── Branch performance ──
    st.subheader("Branch performance")
    b1, b2, b3 = st.columns(3)
    with b1:
        st.markdown("**Main branch**")
        st.metric("Revenue", f"₦{summary['main_revenue']:,.0f}")
        st.metric("Expenses", f"₦{summary['main_expenses']:,.0f}")
        st.metric("Profit (approx.)", f"₦{summary['main_profit']:,.0f}")
    with b2:
        st.markdown("**Punch branch**")
        st.metric("Revenue", f"₦{summary['punch_revenue']:,.0f}")
        st.metric("Expenses", f"₦{summary['punch_expenses']:,.0f}")
        st.metric("Profit (approx.)", f"₦{summary['punch_profit']:,.0f}")
    with b3:
        st.markdown("**Company total**")
        st.metric("Total revenue", f"₦{summary['total_revenue']:,.0f}")
        st.metric("Total expenses + salaries", f"₦{summary['total_expenses']:,.0f}")
        st.metric("Net profit (approx.)", f"₦{summary['net_profit']:,.0f}")

    st.caption(
        f"Salaries paid in this period: ₦{summary['salary_paid']:,.0f}. "
        "Record costs under **Expenses** and salary under **Staff & Salaries**."
    )


# ─────────────────────────────────────────────
# RECORD SALE
# ─────────────────────────────────────────────
elif page == "Record Sale":
    st.header("Record New Sale")
    products = get_products_in_stock_first()
    if products.empty:
        st.warning("No products found. Ask Admin to add products first.")
    else:
        in_stock = products[products["current_stock"] > 0]
        st.caption(
            f"Products with stock shown first ({len(in_stock)} in stock of {len(products)} total). "
            "You can change the selling price for this sale."
        )
        with st.form("sale_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                sale_date = st.date_input("Sale Date", value=date.today())
                options = {
                    f"{'✓ ' if r['current_stock'] > 0 else '✗ '}{r['name']}  ·  Stock: {r['current_stock']:.0f}  ·  List ₦{r['unit_price']:,.0f}": r["id"]
                    for _, r in products.iterrows()
                }
                label = st.selectbox("Product (in-stock first)", list(options.keys()))
                product_id = options[label]
            with c2:
                units = st.number_input("Units Sold", min_value=0.0, step=1.0, value=1.0)
                notes = st.text_input("Notes (optional)", placeholder="e.g. walk-in customer")

            prod = get_product_by_id(product_id)
            default_price = float(prod["unit_price"]) if prod else 0.0
            sell_price = st.number_input(
                "Selling Price for this sale (₦)",
                min_value=0.0,
                step=50.0,
                value=default_price,
                help="Change this if the product is sold at a different price than the list price.",
                key="sale_sell_price",
            )
            pay_method = st.selectbox(
                "Payment method",
                ["Cash", "Transfer", "POS"],
                key="sale_pay_method",
                help="Cash → Cash account; Transfer/POS → Bank account in the books.",
            )

            if prod:
                amt = units * sell_price
                close = prod["current_stock"] - units
                st.info(
                    f"**Stock available:** {prod['current_stock']:.0f}  ·  "
                    f"**List price:** ₦{default_price:,.0f}  ·  "
                    f"**This sale price:** ₦{sell_price:,.0f}  ·  "
                    f"**Amount:** ₦{amt:,.0f}  ·  "
                    f"**Stock after sale:** {close:.0f}  ·  "
                    f"**Pay:** {pay_method}"
                )

            if st.form_submit_button("Save Sale", type="primary", use_container_width=True):
                uid, err = add_sale(
                    sale_date.isoformat(),
                    product_id,
                    units,
                    notes,
                    unit_price=sell_price,
                    payment_method=pay_method,
                )
                if err:
                    st.error(err)
                else:
                    st.success(f"Sale saved  ·  Unique ID: **{uid}**")
                    st.balloons()

# ─────────────────────────────────────────────
# VIEW STOCK (read-only — sales reps + admin can open; sales only via menu)
# ─────────────────────────────────────────────
elif page == "View Stock":
    st.header("Available Stock")
    if is_admin:
        st.caption("See what is in stock and what is running low. Use **Stock Tracker** to add production or restock.")
    else:
        st.caption("See what is available and what is running low. Only Admin can change stock.")

    products = get_products_in_stock_first()
    if products.empty:
        st.info("No products in the catalogue yet.")
    else:
        LOW_THRESHOLD = 5  # at or below this = needs attention
        disp = products[["name", "sku", "current_stock", "unit_price"]].copy()
        disp.columns = ["Product", "SKU", "Stock Available", "List Price (₦)"]

        out_stock = disp[disp["Stock Available"] <= 0]
        low_stock = disp[(disp["Stock Available"] > 0) & (disp["Stock Available"] <= LOW_THRESHOLD)]
        ok_stock = disp[disp["Stock Available"] > LOW_THRESHOLD]

        m1, m2, m3 = st.columns(3)
        m1.metric("Out of stock", len(out_stock))
        m2.metric("Running low (≤5)", len(low_stock))
        m3.metric("Healthy stock", len(ok_stock))

        st.divider()

        # Priority: what to produce / restock
        st.subheader("Needs production or restock")
        if out_stock.empty and low_stock.empty:
            st.success("Nothing critical — no products are out of stock or running low.")
        else:
            if not out_stock.empty:
                st.error(f"**Out of stock — produce or restock these ({len(out_stock)})**")
                st.dataframe(out_stock, use_container_width=True, hide_index=True)
            if not low_stock.empty:
                st.warning(f"**Running low — consider producing more ({len(low_stock)})**")
                st.dataframe(low_stock, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("All available stock")
        if ok_stock.empty and low_stock.empty:
            st.info("No products currently in stock.")
        else:
            # Show low + healthy, sorted by stock ascending so low ones visible first
            available = disp[disp["Stock Available"] > 0].sort_values("Stock Available")
            st.dataframe(available, use_container_width=True, hide_index=True)

        st.caption(
            f"Total products: {len(disp)}  ·  "
            f"Out of stock: {len(out_stock)}  ·  "
            f"Running low: {len(low_stock)}  ·  "
            f"Healthy: {len(ok_stock)}"
        )
        if is_admin:
            st.info("To add stock after production, go to **Stock Tracker** → Production / Purchase.")


# ─────────────────────────────────────────────
# STOCK TRACKER
# ─────────────────────────────────────────────

# ─────────────────────────────────────────────
# SPECIAL ORDERS
# ─────────────────────────────────────────────

# ─────────────────────────────────────────────
# RAW MATERIALS INVENTORY
# ─────────────────────────────────────────────
elif page == "Raw Materials":
    if not is_admin:
        st.error("Only Admin can manage raw materials.")
        st.stop()

    st.header("Raw Materials Inventory")
    st.caption(
        "Track flour, butter, sugar and other ingredients. "
        "Valuation method: **Weighted Average Cost (WAC)** — "
        "unit cost updates automatically when you buy at a new price. "
        "Stock value = quantity × average unit cost."
    )

    tab1, tab2, tab3 = st.tabs(["Stock Overview", "Record Movement", "Materials List"])

    materials = get_raw_materials_df()

    with tab1:
        if materials.empty:
            st.info("No raw materials yet. Add them in **Materials List**.")
        else:
            low = materials[materials["current_qty"] <= materials["reorder_level"]]
            ok = materials[materials["current_qty"] > materials["reorder_level"]]
            m1, m2, m3 = st.columns(3)
            m1.metric("Materials", len(materials))
            m2.metric("Need reorder", len(low))
            total_val = (materials["current_qty"] * materials["unit_cost"]).sum()
            m3.metric("Inventory value (₦)", f"{total_val:,.0f}")
            st.caption("Inventory value uses Weighted Average Cost × quantity on hand.")

            st.divider()
            st.subheader("Reorder soon")
            if low.empty:
                st.success("All materials are above reorder level.")
            else:
                show = low[["name", "unit", "current_qty", "reorder_level", "unit_cost"]].copy()
                show.columns = ["Material", "Unit", "In stock", "Reorder at", "Unit cost (₦)"]
                st.warning(f"**{len(low)} material(s)** at or below reorder level")
                st.dataframe(show, use_container_width=True, hide_index=True)

            st.subheader("All materials")
            all_show = materials[["name", "unit", "current_qty", "reorder_level", "unit_cost"]].copy()
            all_show.columns = ["Material", "Unit", "In stock", "Reorder at", "Unit cost (₦)"]
            all_show = all_show.sort_values("In stock")
            st.dataframe(all_show, use_container_width=True, hide_index=True)

            st.divider()
            st.subheader("Recent movements")
            moves = get_raw_movements_df()
            if moves.empty:
                st.caption("No movements recorded yet.")
            else:
                recent = moves.head(20)[
                    ["unique_id", "move_date", "material_name", "movement_type",
                     "quantity", "qty_before", "qty_after", "notes"]
                ].copy()
                type_labels = {
                    "purchase": "Purchase / Received",
                    "received": "Purchase / Received",
                    "used_in_production": "Used in production",
                    "waste_spoilage": "Waste / Spoilage",
                    "adjustment_in": "Adjustment in",
                    "adjustment_out": "Adjustment out",
                }
                recent["movement_type"] = recent["movement_type"].map(
                    lambda x: type_labels.get(x, x)
                )
                recent.columns = [
                    "ID", "Date", "Material", "Type", "Qty", "Before", "After", "Notes"
                ]
                st.dataframe(recent, use_container_width=True, hide_index=True)

    with tab2:
        if materials.empty:
            st.warning("Add materials first in the Materials List tab.")
        else:
            st.subheader("Record ingredient movement")
            with st.form("raw_move_form", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    move_date = st.date_input("Date", value=date.today(), key="raw_move_date")
                    mat_map = {
                        f"{r['name']}  ·  stock {r['current_qty']:.2f} {r['unit']}": r["id"]
                        for _, r in materials.iterrows()
                    }
                    mat_label = st.selectbox("Material", list(mat_map.keys()))
                    material_id = mat_map[mat_label]
                with c2:
                    mtype = st.selectbox(
                        "Type",
                        [
                            "purchase",
                            "used_in_production",
                            "waste_spoilage",
                            "adjustment_in",
                            "adjustment_out",
                        ],
                        format_func=lambda x: {
                            "purchase": "Purchase / Received (add)",
                            "used_in_production": "Used in production (remove)",
                            "waste_spoilage": "Waste / Spoilage (remove)",
                            "adjustment_in": "Adjustment in (add)",
                            "adjustment_out": "Adjustment out (remove)",
                        }[x],
                    )
                mat = get_raw_material_by_id(material_id)
                unit_label = (mat["unit"] if mat else "unit") or "unit"
                st.markdown(f"**Unit of measure:** `{unit_label}` (e.g. kg, litres, crates, cl)")
                qty = st.number_input(
                    f"Quantity (number of {unit_label})",
                    min_value=0.0,
                    step=0.5,
                    value=1.0,
                    help=f"How many {unit_label} moved in this transaction.",
                    key="raw_qty",
                )
                total_paid = 0.0
                total_cost = 0.0
                amount_paid_now = 0.0
                pay_method = "Cash"
                if mtype in ("purchase", "received"):
                    st.markdown("**Purchase details**")
                    total_cost = st.number_input(
                        "Total cost of goods (₦) *",
                        min_value=0.0,
                        step=100.0,
                        value=0.0,
                        help="Full value of this purchase (what the goods cost).",
                        key="raw_total_cost",
                    )
                    p1, p2, p3 = st.columns(3)
                    with p1:
                        amount_paid_now = st.number_input(
                            "Amount paid now (₦)",
                            min_value=0.0,
                            step=100.0,
                            value=0.0,
                            help="What you paid today. Leave 0 if fully on credit.",
                            key="raw_paid_now",
                        )
                    with p2:
                        on_credit = max(0.0, float(total_cost) - float(amount_paid_now))
                        st.metric("On credit (₦)", f"{on_credit:,.0f}")
                    with p3:
                        if qty > 0 and total_cost > 0:
                            st.metric(f"Price per {unit_label}", f"₦{total_cost/qty:,.2f}")
                        else:
                            st.metric(f"Price per {unit_label}", "—")
                    pay_method = st.selectbox(
                        "Payment method (for amount paid now)",
                        ["Cash", "Transfer", "POS"],
                        key="raw_pay_method",
                    )
                    if total_cost > 0:
                        if amount_paid_now > total_cost:
                            st.error("Amount paid now cannot be more than total cost.")
                        elif on_credit > 0 and amount_paid_now > 0:
                            st.info(
                                f"Inventory +₦{total_cost:,.0f}  ·  "
                                f"Cash out ₦{amount_paid_now:,.0f}  ·  "
                                f"**Owed to supplier ₦{on_credit:,.0f}**"
                            )
                        elif on_credit > 0:
                            st.info(
                                f"Fully on credit: Inventory +₦{total_cost:,.0f}  ·  "
                                f"**Payable ₦{on_credit:,.0f}** (no cash out)"
                            )
                        else:
                            st.success(
                                f"Fully paid: {qty:g} {unit_label} @ ₦{total_cost/qty:,.2f} = ₦{total_cost:,.0f}"
                            )
                    else:
                        st.caption(
                            f"Enter quantity ({unit_label}), total cost, and how much you paid now."
                        )
                notes = st.text_input("Notes", placeholder="e.g. supplier name, market")
                if st.form_submit_button("Save Movement", type="primary", use_container_width=True):
                    if mtype in ("purchase", "received"):
                        if total_cost <= 0:
                            st.error("Enter total cost of the goods.")
                        elif amount_paid_now > total_cost:
                            st.error("Amount paid now cannot exceed total cost.")
                        else:
                            uid, err = add_raw_movement(
                                move_date.isoformat(),
                                material_id,
                                mtype,
                                qty,
                                total_amount=total_cost,
                                notes=notes,
                                amount_paid=amount_paid_now,
                                payment_method=pay_method,
                            )
                            if err:
                                st.error(err)
                            else:
                                oc = total_cost - amount_paid_now
                                msg = f"Saved · **{uid}** · {qty:g} {unit_label}"
                                if oc > 0:
                                    msg += f" · On credit ₦{oc:,.0f}"
                                st.success(msg)
                                st.rerun()
                    else:
                        uid, err = add_raw_movement(
                            move_date.isoformat(), material_id, mtype, qty, notes=notes
                        )
                        if err:
                            st.error(err)
                        else:
                            st.success(f"Saved  ·  **{uid}**")
                            st.rerun()

            st.divider()
            st.subheader("Delete a movement")
            moves = get_raw_movements_df()
            if not moves.empty:
                mmap = {
                    f"{r['unique_id']}  ·  {r['material_name']}  ·  {r['movement_type']}  ·  {r['quantity']}": r["id"]
                    for _, r in moves.iterrows()
                }
                sel = st.selectbox("Select movement", list(mmap.keys()), key="del_raw_move")
                dc1, dc2 = st.columns(2)
                with dc1:
                    if st.button("Delete (reverse stock)", key="btn_del_raw"):
                        ok, err = delete_raw_movement(mmap[sel], force=False)
                        if ok:
                            st.success("Deleted and stock reversed.")
                            st.rerun()
                        else:
                            st.error(err)
                with dc2:
                    if st.button("Force delete", key="btn_force_del_raw"):
                        ok, err = delete_raw_movement(mmap[sel], force=True)
                        if ok:
                            st.success("Force deleted.")
                            st.rerun()
                        else:
                            st.error(err)

    with tab3:
        st.subheader("Add material")
        with st.form("add_raw_mat", clear_on_submit=True):
            a1, a2 = st.columns(2)
            with a1:
                n = st.text_input("Name *", placeholder="e.g. Flour")
                u = st.text_input("Unit *", value="kg", placeholder="kg, litres, crates…")
            with a2:
                rq = st.number_input("Reorder level", min_value=0.0, step=1.0, value=5.0)
                uc = st.number_input("Unit cost (₦)", min_value=0.0, step=50.0, value=0.0)
                oq = st.number_input("Opening quantity", min_value=0.0, step=1.0, value=0.0)
            nt = st.text_input("Notes", key="raw_add_notes")
            if st.form_submit_button("Add Material", type="primary"):
                ok, err = add_raw_material(n, u, rq, uc, oq, nt)
                if ok:
                    st.success(f"Added **{n.strip()}**")
                    st.rerun()
                else:
                    st.error(err)

        st.divider()
        st.subheader("Edit / delete material")
        materials = get_raw_materials_df()
        if materials.empty:
            st.caption("No materials to edit.")
        else:
            emap = {r["name"]: r["id"] for _, r in materials.iterrows()}
            esel = st.selectbox("Select material", list(emap.keys()), key="edit_raw_sel")
            cur = materials[materials["id"] == emap[esel]].iloc[0]
            e1, e2 = st.columns(2)
            with e1:
                en = st.text_input("Name", value=cur["name"], key="er_name")
                eu = st.text_input("Unit", value=cur["unit"], key="er_unit")
            with e2:
                er = st.number_input("Reorder level", min_value=0.0, value=float(cur["reorder_level"]), key="er_re")
                ec = st.number_input("Unit cost (₦)", min_value=0.0, value=float(cur["unit_cost"]), key="er_cost")
            enotes = st.text_input("Notes", value=cur["notes"] or "", key="er_notes")
            b1, b2 = st.columns(2)
            with b1:
                if st.button("Save changes", type="primary", key="btn_save_raw"):
                    ok, err = update_raw_material(emap[esel], en, eu, er, ec, enotes)
                    if ok:
                        st.success("Updated.")
                        st.rerun()
                    else:
                        st.error(err)
            with b2:
                if st.button("Delete material", type="secondary", key="btn_del_raw_mat"):
                    ok, err = delete_raw_material(emap[esel])
                    if ok:
                        st.success("Deleted.")
                        st.rerun()
                    else:
                        st.error(err)


elif page == "Special Orders":
    st.header("Special Orders")
    st.caption("Record custom orders produced specially for a customer. Enter details manually.")

    tab1, tab2 = st.tabs(["Record Order", "Order History"])

    with tab1:
        with st.form("order_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                order_date = st.date_input("Order Date", value=date.today())
                customer = st.text_input("Customer Name", placeholder="e.g. Mrs. Ade")
                product_desc = st.text_input(
                    "Product / Description *",
                    placeholder="e.g. 3-tier wedding cake, Red velvet",
                )
            with c2:
                qty = st.number_input("Quantity", min_value=0.0, step=1.0, value=1.0)
                price = st.number_input("Unit Price (₦)", min_value=0.0, step=100.0, value=0.0)
                status = st.selectbox("Status", ["Pending", "Producing", "Ready", "Delivered", "Cancelled"])
            notes = st.text_input("Notes (optional)", placeholder="Delivery date, flavour, etc.")
            st.info(f"**Estimated Amount:** ₦{qty * price:,.0f}")
            if st.form_submit_button("Save Special Order", type="primary", use_container_width=True):
                uid, err = add_special_order(
                    order_date.isoformat(), customer, product_desc, qty, price, status, notes
                )
                if err:
                    st.error(err)
                else:
                    st.success(f"Order saved  ·  Unique ID: **{uid}**")

    with tab2:
        oc1, oc2 = st.columns(2)
        with oc1:
            o_start = st.date_input("From", value=date.today().replace(day=1), key="ord_start")
        with oc2:
            o_end = st.date_input("To", value=date.today(), key="ord_end")
        orders = get_special_orders_df(o_start.isoformat(), o_end.isoformat())
        if orders.empty:
            st.info("No special orders in this period.")
        else:
            st.markdown(
                f"**Orders:** {len(orders)}  ·  **Total value:** ₦{orders['amount'].sum():,.0f}"
            )
            disp = orders[
                ["id", "unique_id", "order_date", "customer_name", "product_description",
                 "quantity", "unit_price", "amount", "status", "notes"]
            ].copy()
            disp.columns = [
                "ID", "Unique ID", "Date", "Customer", "Product", "Qty",
                "Unit Price", "Amount", "Status", "Notes",
            ]
            st.dataframe(disp.drop(columns=["ID"]), use_container_width=True, hide_index=True)

            st.divider()
            st.subheader("Edit or Delete Order")
            omap = {
                f"{r['unique_id']}  ·  {r['product_description'][:40]}  ·  {r['status']}": r["id"]
                for _, r in orders.iterrows()
            }
            sel = st.selectbox("Select order", list(omap.keys()), key="edit_ord_sel")
            oid = omap[sel]
            row = orders[orders["id"] == oid].iloc[0]
            e1, e2 = st.columns(2)
            with e1:
                ed_cust = st.text_input("Customer", value=row["customer_name"] or "", key="ed_cust")
                ed_prod = st.text_input("Product", value=row["product_description"], key="ed_prod")
                ed_qty = st.number_input("Quantity", min_value=0.0, value=float(row["quantity"]), key="ed_oqty")
            with e2:
                ed_price = st.number_input("Unit Price", min_value=0.0, value=float(row["unit_price"]), key="ed_oprice")
                ed_status = st.selectbox(
                    "Status",
                    ["Pending", "Producing", "Ready", "Delivered", "Cancelled"],
                    index=["Pending", "Producing", "Ready", "Delivered", "Cancelled"].index(row["status"])
                    if row["status"] in ["Pending", "Producing", "Ready", "Delivered", "Cancelled"] else 0,
                    key="ed_status",
                )
                ed_notes = st.text_input("Notes", value=row["notes"] or "", key="ed_onotes")
            b1, b2 = st.columns(2)
            with b1:
                if st.button("Save Order Changes", type="primary", key="btn_upd_ord"):
                    ok, err = update_special_order(oid, ed_cust, ed_prod, ed_qty, ed_price, ed_status, ed_notes)
                    if ok:
                        st.success("Order updated.")
                        st.rerun()
                    else:
                        st.error(err)
            with b2:
                if st.button("Delete Order", type="secondary", key="btn_del_ord"):
                    delete_special_order(oid)
                    st.success("Order deleted.")
                    st.rerun()


# ─────────────────────────────────────────────
# PUNCH BRANCH
# ─────────────────────────────────────────────
elif page == "Punch Branch":
    st.header("Funlex Branch – Punch")
    st.caption(
        "Products for Punch are produced specially (not from main stock). "
        "Record quantity sent, quantity sold, remaining, and the selling price at Punch "
        "(can differ from main branch price)."
    )

    tab1, tab2 = st.tabs(["Record Daily Production / Sales", "Branch History"])

    products = get_products_in_stock_first()

    with tab1:
        if products.empty:
            st.warning("No products in main branch. Add products first.")
        else:
            with st.form("branch_form", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    br_date = st.date_input("Date", value=date.today(), key="br_date")
                    options = {
                        f"{r['name']}  ·  ₦{r['unit_price']:,.0f}": r["id"]
                        for _, r in products.iterrows()
                    }
                    label = st.selectbox("Product", list(options.keys()), key="br_prod")
                    product_id = options[label]
                with c2:
                    qty_taken = st.number_input("Qty Produced / Sent to Punch", min_value=0.0, step=1.0, value=1.0)
                    qty_sold = st.number_input("Qty Sold at Punch", min_value=0.0, step=1.0, value=0.0)
                prod = get_product_by_id(product_id)
                default_price = float(prod["unit_price"]) if prod else 0.0
                punch_price = st.number_input(
                    "Selling Price at Punch (₦)",
                    min_value=0.0,
                    step=50.0,
                    value=default_price,
                    key="br_punch_price",
                    help="Punch can sell at a different price from the main branch catalogue.",
                )
                notes = st.text_input("Notes (optional)", key="br_notes")
                pay_method = st.selectbox(
                    "Payment method (for amount sold)",
                    ["Cash", "Transfer", "POS"],
                    key="br_pay",
                )
                remaining = max(0, qty_taken - qty_sold)
                amt = qty_sold * punch_price
                st.info(
                    f"**Remaining at Punch:** {remaining:.0f}  ·  "
                    f"**Amount Sold at Punch:** ₦{amt:,.0f}  "
                    f"(Main catalogue price: ₦{default_price:,.0f} — Punch price can differ)"
                )
                if st.form_submit_button("Save Branch Record", type="primary", use_container_width=True):
                    uid, err = add_branch_record(
                        br_date.isoformat(), product_id, qty_taken, qty_sold, punch_price, notes,
                        payment_method=pay_method,
                    )
                    if err:
                        st.error(err)
                    else:
                        st.success(f"Branch record saved  ·  Unique ID: **{uid}**")

    with tab2:
        bc1, bc2 = st.columns(2)
        with bc1:
            b_start = st.date_input("From", value=date.today().replace(day=1), key="br_start")
        with bc2:
            b_end = st.date_input("To", value=date.today(), key="br_end")
        branch = get_branch_df(b_start.isoformat(), b_end.isoformat())
        if branch.empty:
            st.info("No branch records in this period.")
        else:
            st.markdown(
                f"**Records:** {len(branch)}  ·  "
                f"**Total produced/sent:** {branch['qty_taken'].sum():.0f}  ·  "
                f"**Total sold:** {branch['qty_sold'].sum():.0f}  ·  "
                f"**Revenue at Punch:** ₦{branch['amount_sold'].sum():,.0f}"
            )
            disp = branch[
                ["id", "unique_id", "record_date", "product_name", "qty_taken",
                 "qty_sold", "qty_remaining", "amount_sold", "unit_price", "notes"]
            ].copy()
            disp.columns = [
                "ID", "Unique ID", "Date", "Product", "Produced/Sent", "Sold",
                "Remaining", "Amount Sold (₦)", "Punch Price (₦)", "Notes",
            ]
            st.dataframe(disp.drop(columns=["ID"]), use_container_width=True, hide_index=True)

            st.divider()
            st.subheader("Edit or Delete Branch Record")
            bmap = {
                f"{r['unique_id']}  ·  {r['product_name']}  ·  Sent {r['qty_taken']:.0f} / Sold {r['qty_sold']:.0f}": r["id"]
                for _, r in branch.iterrows()
            }
            sel_b = st.selectbox("Select record", list(bmap.keys()), key="edit_br_sel")
            bid = bmap[sel_b]
            brow = branch[branch["id"] == bid].iloc[0]
            x1, x2, x3 = st.columns(3)
            with x1:
                ed_taken = st.number_input("Qty Produced/Sent", min_value=0.0, value=float(brow["qty_taken"]), key="ed_taken")
            with x2:
                ed_sold = st.number_input("Qty Sold", min_value=0.0, value=float(brow["qty_sold"]), key="ed_sold")
            with x3:
                ed_price = st.number_input(
                    "Selling Price at Punch (₦)",
                    min_value=0.0,
                    step=50.0,
                    value=float(brow["unit_price"]),
                    key="ed_br_price",
                )
            ed_bnotes = st.text_input("Notes", value=brow["notes"] or "", key="ed_bnotes")
            st.caption(
                f"Remaining: **{ed_taken - ed_sold:.0f}**  ·  "
                f"Amount sold: **₦{(ed_taken and ed_sold) * ed_price if False else ed_sold * ed_price:,.0f}**"
            )
            y1, y2 = st.columns(2)
            with y1:
                if st.button("Save Changes", type="primary", key="btn_upd_br"):
                    ok, err = update_branch_record(bid, ed_taken, ed_sold, ed_price, ed_bnotes)
                    if ok:
                        st.success("Branch record updated.")
                        st.rerun()
                    else:
                        st.error(err)
            with y2:
                if st.button("Delete Record", type="secondary", key="btn_del_br"):
                    ok, err = delete_branch_record(bid)
                    if ok:
                        st.success("Branch record deleted.")
                        st.rerun()
                    else:
                        st.error(err)


elif page == "Stock Tracker":
    st.header("Stock Tracker")
    tab1, tab2 = st.tabs(["Record Movement", "Movement History"])
    products = get_all_products()

    with tab1:
        if products.empty:
            st.warning("No products found.")
        else:
            with st.form("stock_form", clear_on_submit=True):
                c1, c2 = st.columns(2)
                with c1:
                    move_date = st.date_input("Date", value=date.today(), key="stk_date")
                    options = {
                        f"{r['name']}  ·  Current: {r['current_stock']:.0f}": r["id"]
                        for _, r in products.iterrows()
                    }
                    label = st.selectbox("Product", list(options.keys()), key="stk_prod")
                    product_id = options[label]
                with c2:
                    move_type = st.selectbox(
                        "Movement Type",
                        ["production", "purchase", "spoilage", "withdrawn", "adjustment_in", "adjustment_out"],
                        format_func=lambda x: {
                            "production": "Production (add)",
                            "purchase": "Purchase (add)",
                            "spoilage": "Spoilage / Bad (remove)",
                            "withdrawn": "Withdrawn / Staff use (remove)",
                            "adjustment_in": "Adjustment In (add)",
                            "adjustment_out": "Adjustment Out (remove)",
                        }[x],
                    )
                    qty = st.number_input("Quantity", min_value=0.0, step=1.0, value=1.0)
                notes = st.text_input("Notes (optional)", key="stk_notes")
                if st.form_submit_button("Save Movement", type="primary", use_container_width=True):
                    uid, err = add_stock_movement(
                        move_date.isoformat(), product_id, move_type, qty, notes
                    )
                    if err:
                        st.error(err)
                    else:
                        st.success(f"Movement saved  ·  Unique ID: **{uid}**")

    with tab2:
        moves = get_stock_movements_df()
        if moves.empty:
            st.info("No stock movements yet.")
        else:
            disp = moves[
                ["id", "unique_id", "move_date", "product_name", "movement_type",
                 "quantity", "stock_before", "stock_after", "notes"]
            ].copy()
            disp.columns = ["ID", "Unique ID", "Date", "Product", "Type", "Qty", "Before", "After", "Notes"]
            st.dataframe(disp.drop(columns=["ID"]), use_container_width=True, hide_index=True)

            st.divider()
            st.subheader("Edit or Delete a Stock Movement")
            move_ids = {
                f"{r['unique_id']}  ·  {r['product_name']}  ·  {r['movement_type']}  ·  {r['quantity']:.0f}": r["id"]
                for _, r in moves.iterrows()
            }
            sel_move = st.selectbox("Select movement to edit/delete", list(move_ids.keys()), key="edit_move_sel")
            move_id = move_ids[sel_move]
            move_row = get_movement_by_id(move_id)

            type_options = ["production", "purchase", "spoilage", "withdrawn", "adjustment_in", "adjustment_out"]
            type_labels = {
                "production": "Production (add)",
                "purchase": "Purchase (add)",
                "spoilage": "Spoilage / Bad (remove)",
                "withdrawn": "Withdrawn / Staff use (remove)",
                "adjustment_in": "Adjustment In (add)",
                "adjustment_out": "Adjustment Out (remove)",
            }

            if move_row:
                ec1, ec2 = st.columns(2)
                with ec1:
                    edit_date = st.date_input(
                        "Date",
                        value=date.fromisoformat(move_row["move_date"]) if move_row["move_date"] else date.today(),
                        key="edit_move_date",
                    )
                    current_type = move_row["movement_type"]
                    type_idx = type_options.index(current_type) if current_type in type_options else 0
                    edit_type = st.selectbox(
                        "Type",
                        type_options,
                        index=type_idx,
                        format_func=lambda x: type_labels.get(x, x),
                        key="edit_move_type",
                    )
                with ec2:
                    edit_qty = st.number_input(
                        "Quantity",
                        min_value=0.0,
                        step=1.0,
                        value=float(move_row["quantity"]),
                        key="edit_move_qty",
                    )
                    edit_notes = st.text_input(
                        "Notes",
                        value=move_row["notes"] or "",
                        key="edit_move_notes",
                    )

                bc1, bc2, bc3 = st.columns(3)
                with bc1:
                    if st.button("Save Changes", type="primary", key="btn_save_move"):
                        ok, err = update_stock_movement(
                            move_id, edit_qty, edit_type, edit_date.isoformat(), edit_notes
                        )
                        if ok:
                            st.success("Movement updated.")
                            st.rerun()
                        else:
                            st.error(err)
                with bc2:
                    if st.button("Delete", type="secondary", key="btn_del_move"):
                        ok, err = delete_stock_movement(move_id, force=False)
                        if ok:
                            st.success("Movement deleted and stock reversed.")
                            st.rerun()
                        else:
                            st.error(err)
                with bc3:
                    if st.button("Force Delete", key="btn_force_del_move"):
                        ok, err = delete_stock_movement(move_id, force=True)
                        if ok:
                            st.success("Force-deleted. Stock set to 0 if needed.")
                            st.rerun()
                        else:
                            st.error(err)

        st.divider()
        st.subheader("Clear Stock / Production")
        st.caption("Use these to set stock back to zero or wipe movement history.")

        products = get_all_products()
        if not products.empty:
            prod_map = {r["name"]: r["id"] for _, r in products.iterrows()}
            clear_prod = st.selectbox(
                "Select product to clear stock to zero",
                list(prod_map.keys()),
                key="clear_prod_sel",
            )
            cc1, cc2 = st.columns(2)
            with cc1:
                if st.button("Set this product stock → 0", type="secondary", key="btn_clear_one"):
                    clear_product_stock(prod_map[clear_prod])
                    st.success(f"Stock for **{clear_prod}** set to 0.")
                    st.rerun()
            with cc2:
                if st.button("Set ALL products stock → 0", key="btn_clear_all_stock"):
                    reset_all_stocks()
                    st.success("All product stocks set to 0.")
                    st.rerun()

        st.markdown("---")
        st.warning("Clearing all movement history only deletes the records. Product stock levels are not changed.")
        if st.button("Clear ALL stock movement history", key="btn_clear_all_moves"):
            clear_all_movements()
            st.success("All stock movement records deleted.")
            st.rerun()


# ─────────────────────────────────────────────
# SALES HISTORY
# ─────────────────────────────────────────────
elif page == "Sales History":
    st.header("Sales History")
    c1, c2 = st.columns(2)
    with c1:
        start = st.date_input("From", value=date.today().replace(day=1), key="hist_start")
    with c2:
        end = st.date_input("To", value=date.today(), key="hist_end")

    sales = get_sales_df(start.isoformat(), end.isoformat())
    if sales.empty:
        st.info("No sales in this period.")
    else:
        st.markdown(
            f"**Revenue:** ₦{sales['amount'].sum():,.0f}  ·  "
            f"**Units:** {sales['units_sold'].sum():,.0f}  ·  "
            f"**Transactions:** {len(sales)}"
        )
        disp = sales[
            ["id", "unique_id", "sale_date", "product_name", "opening_stock",
             "units_sold", "unit_price", "amount", "closing_stock", "notes"]
        ].copy()
        disp.columns = [
            "ID", "Unique ID", "Date", "Product", "Opening", "Units Sold",
            "Unit Price (₦)", "Amount (₦)", "Closing", "Notes",
        ]
        st.dataframe(disp.drop(columns=["ID"]), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Edit or Delete a Sale")
        sale_ids = {f"{r['unique_id']}  ·  {r['product_name']}  ·  {r['units_sold']:.0f} units": r["id"]
                    for _, r in sales.iterrows()}
        selected = st.selectbox("Select sale record", list(sale_ids.keys()), key="edit_sale_sel")
        sale_id = sale_ids[selected]
        sale_row = get_sale_by_id(sale_id)

        ec1, ec2 = st.columns(2)
        with ec1:
            new_units = st.number_input(
                "Units Sold", min_value=0.0, step=1.0,
                value=float(sale_row["units_sold"]) if sale_row else 1.0,
                key="edit_sale_units",
            )
            new_notes = st.text_input(
                "Notes", value=sale_row["notes"] or "" if sale_row else "",
                key="edit_sale_notes",
            )
            if st.button("Save Changes", type="primary", key="btn_update_sale"):
                ok, err = update_sale(sale_id, new_units, new_notes)
                if ok:
                    st.success("Sale updated successfully.")
                    st.rerun()
                else:
                    st.error(err)
        with ec2:
            st.write("")
            st.write("")
            st.warning("Deleting a sale will restore the stock that was deducted.")
            if st.button("Delete this Sale", type="secondary", key="btn_delete_sale"):
                ok, err = delete_sale(sale_id)
                if ok:
                    st.success("Sale deleted and stock restored.")
                    st.rerun()
                else:
                    st.error(err)


# ─────────────────────────────────────────────
# DAILY TOTALS
# ─────────────────────────────────────────────
elif page == "Daily Totals":
    st.header("Daily Sales Totals")
    daily = get_daily_totals()
    if daily.empty:
        st.info("No sales recorded yet.")
    else:
        disp = daily.copy()
        disp["total_amount"] = disp["total_amount"].apply(lambda x: f"₦{x:,.0f}")
        disp.columns = ["Date", "Transactions", "Total Units", "Total Amount"]
        st.dataframe(disp, use_container_width=True, hide_index=True)

        chart_df = daily.copy()
        chart_df["date"] = pd.to_datetime(chart_df["date"])
        st.bar_chart(chart_df.set_index("date")["total_amount"], use_container_width=True)


# ─────────────────────────────────────────────
# PRODUCTS
# ─────────────────────────────────────────────
elif page == "Products":
    st.header("Product Catalogue")
    tab1, tab2, tab3 = st.tabs(["All Products", "Add New Product", "Edit / Rename Product"])

    with tab1:
        products = get_all_products()
        if products.empty:
            st.info("No products yet. Add products in the next tab.")
        else:
            disp = products.copy()
            if "cost_price" not in disp.columns:
                disp["cost_price"] = 0
            disp["unit_price"] = disp["unit_price"].apply(lambda x: f"₦{x:,.0f}")
            disp["cost_price"] = disp["cost_price"].apply(lambda x: f"₦{float(x):,.0f}")
            disp = disp[["id", "name", "sku", "unit_price", "cost_price", "current_stock"]]
            disp.columns = ["ID", "Product", "SKU", "Sell Price", "Cost Price", "Stock"]
            st.dataframe(disp, use_container_width=True, hide_index=True)
            st.caption("Cost price is used for COGS when a sale is recorded (Gross profit = Sales − COGS).")
            st.caption(f"Total products: {len(products)}")

        st.divider()
        st.subheader("Clear All Products")
        st.caption("Removes every product so you can type them again correctly. Sales history is kept.")
        confirm = st.checkbox("Yes, clear the entire product list", key="confirm_clear_prods")
        if st.button("Clear All Products", type="secondary", key="btn_clear_prods"):
            if not confirm:
                st.error("Tick the checkbox first to confirm.")
            else:
                ok, err = clear_all_products()
                if ok:
                    st.success("All products cleared. Open **Add New Product** to re-enter them.")
                    st.rerun()
                else:
                    st.error(f"Could not clear products: {err}")

    with tab2:
        st.subheader("Add a New Product")
        with st.form("new_product", clear_on_submit=True):
            n1, n2 = st.columns(2)
            with n1:
                name = st.text_input("Product Name *", placeholder="e.g. Chocolate Cake")
                sku = st.text_input("SKU (optional)", placeholder="e.g. CC01")
            with n2:
                price = st.number_input("Selling Price (₦) *", min_value=0.0, step=50.0, value=500.0)
                cost = st.number_input(
                    "Cost Price (₦)",
                    min_value=0.0,
                    step=50.0,
                    value=0.0,
                    help="What it costs Funlex to make/buy this item. Used for COGS.",
                )
                stock = st.number_input("Opening Stock", min_value=0.0, step=1.0, value=0.0)
            if st.form_submit_button("Add Product", type="primary", use_container_width=True):
                if not name.strip():
                    st.error("Product name is required.")
                elif price <= 0:
                    st.error("Price must be greater than 0.")
                else:
                    ok, err = add_new_product(name.strip(), sku.strip(), price, stock, cost_price=cost)
                    if ok:
                        st.success(f"Product **{name.strip()}** added successfully.")
                        st.rerun()
                    else:
                        st.error(err)

    with tab3:
        products = get_all_products()
        if products.empty:
            st.info("No products to edit. Add products first.")
        else:
            st.subheader("Rename / Edit Product")
            st.caption("Fix spelling, change price, or delete one product.")
            prod_map = {r["name"]: r["id"] for _, r in products.iterrows()}
            sel = st.selectbox("Select Product", list(prod_map.keys()), key="edit_prod_sel")
            current = products[products["id"] == prod_map[sel]].iloc[0]
            pid = int(prod_map[sel])

            new_name = st.text_input("Product Name (fix spelling here)", value=current["name"], key="edit_pname")
            new_sku = st.text_input("SKU", value=current["sku"] or "", key="edit_psku")
            new_price = st.number_input(
                "Selling Price (₦)",
                min_value=0.0,
                value=float(current["unit_price"]),
                step=50.0,
                key="edit_pprice",
            )
            cur_cost = float(current["cost_price"]) if "cost_price" in current.index else 0.0
            new_cost = st.number_input(
                "Cost Price (₦) — for COGS",
                min_value=0.0,
                value=cur_cost,
                step=50.0,
                key="edit_pcost",
                help="Cost to produce/buy this item. Sales will post COGS = units × cost price.",
            )
            st.caption(f"Current stock: **{current['current_stock']:.0f}** (change stock in Stock Tracker)")

            b1, b2, b3 = st.columns(3)
            with b1:
                if st.button("Save Name & SKU", type="primary", key="btn_rename"):
                    ok, err = rename_product(pid, new_name, new_sku)
                    if ok:
                        update_product_price(pid, new_price, cost_price=new_cost)
                        st.success(f"Product updated to **{new_name.strip()}**")
                        st.rerun()
                    else:
                        st.error(err)
            with b2:
                if st.button("Update Prices", key="btn_upd_price"):
                    if new_price <= 0:
                        st.error("Selling price must be greater than 0.")
                    else:
                        update_product_price(pid, new_price, cost_price=new_cost)
                        st.success(f"Sell ₦{new_price:,.0f} · Cost ₦{new_cost:,.0f}")
                        st.rerun()
            with b3:
                if st.button("Delete This Product", type="secondary", key="btn_del_one_prod"):
                    ok, err = delete_single_product(pid)
                    if ok:
                        st.success(f"Deleted **{current['name']}**")
                        st.rerun()
                    else:
                        st.error(err)


# ─────────────────────────────────────────────
# EXPORT EXCEL
# ─────────────────────────────────────────────
elif page == "Export Excel":
    st.header("Download Data as Excel")
    st.caption("At the end of the month, select the full month range and download everything for analysis.")

    # Quick month selector
    st.subheader("Select Period")
    today = date.today()
    c1, c2, c3 = st.columns(3)
    with c1:
        year = st.number_input("Year", min_value=2024, max_value=2030, value=today.year, step=1)
    with c2:
        month = st.selectbox("Month", list(range(1, 13)), index=today.month - 1,
                             format_func=lambda m: date(2000, m, 1).strftime("%B"))
    with c3:
        st.write("")  # spacer
        st.write("")
        use_custom = st.checkbox("Custom date range", value=False)

    if use_custom:
        c1, c2 = st.columns(2)
        with c1:
            exp_start = st.date_input("From", value=date(today.year, today.month, 1), key="exp_start")
        with c2:
            exp_end = st.date_input("To", value=today, key="exp_end")
    else:
        # Full month
        last_day = monthrange(int(year), int(month))[1]
        exp_start = date(int(year), int(month), 1)
        exp_end = date(int(year), int(month), last_day)
        st.info(f"**Selected period:** {exp_start.strftime('%d %b %Y')}  →  {exp_end.strftime('%d %b %Y')}")

    sales = get_sales_df(exp_start.isoformat(), exp_end.isoformat())
    moves = get_stock_movements_df(exp_start.isoformat(), exp_end.isoformat())
    products = get_all_products()
    daily = get_daily_totals()
    if not daily.empty:
        daily = daily[
            (daily["date"] >= exp_start.isoformat()) & (daily["date"] <= exp_end.isoformat())
        ]

    loss_count = 0
    if not moves.empty:
        loss_count = len(moves[moves["movement_type"].isin(["spoilage", "withdrawn", "adjustment_out"])])
    st.markdown(
        f"Sales: **{len(sales)}**  ·  Movements: **{len(moves)}**  ·  "
        f"Spoilage/Withdrawn: **{loss_count}**  ·  "
        f"Products: **{len(products)}**  ·  Daily totals: **{len(daily)}**"
    )

    if st.button("Generate Excel File", type="primary", use_container_width=True):
        # --- Sales ---
        sales_exp = sales.copy()
        if not sales_exp.empty:
            sales_exp = sales_exp[
                ["unique_id", "sale_date", "product_name", "opening_stock",
                 "units_sold", "unit_price", "amount", "closing_stock", "notes", "created_at"]
            ]
            sales_exp.columns = [
                "Unique ID", "Date", "Product", "Opening Stock", "Units Sold",
                "Unit Price (₦)", "Amount (₦)", "Closing Stock", "Notes", "Recorded At",
            ]

        # --- All Stock Movements (with clear Type labels) ---
        type_labels = {
            "production": "Production",
            "purchase": "Purchase",
            "spoilage": "Spoilage / Bad",
            "withdrawn": "Withdrawn / Staff Use",
            "adjustment_in": "Adjustment In",
            "adjustment_out": "Adjustment Out",
        }
        moves_exp = moves.copy()
        if not moves_exp.empty:
            moves_exp = moves_exp[
                ["unique_id", "move_date", "product_name", "movement_type",
                 "quantity", "stock_before", "stock_after", "notes", "created_at"]
            ].copy()
            moves_exp["movement_type"] = moves_exp["movement_type"].map(
                lambda x: type_labels.get(x, x)
            )
            moves_exp.columns = [
                "Unique ID", "Date", "Product", "Type", "Quantity",
                "Stock Before", "Stock After", "Notes", "Recorded At",
            ]

        # --- Spoilage & Withdrawals only (for accurate loss analysis) ---
        loss_types = ("spoilage", "withdrawn", "adjustment_out")
        loss_df = moves[moves["movement_type"].isin(loss_types)].copy() if not moves.empty else pd.DataFrame()
        if not loss_df.empty:
            loss_exp = loss_df[
                ["unique_id", "move_date", "product_name", "movement_type",
                 "quantity", "stock_before", "stock_after", "notes", "created_at"]
            ].copy()
            loss_exp["movement_type"] = loss_exp["movement_type"].map(
                lambda x: type_labels.get(x, x)
            )
            loss_exp.columns = [
                "Unique ID", "Date", "Product", "Type (Spoilage/Withdrawn)", "Quantity Lost",
                "Stock Before", "Stock After", "Notes", "Recorded At",
            ]
            # Summary by product
            loss_summary = (
                loss_df.groupby("product_name", as_index=False)
                .agg(Total_Quantity_Lost=("quantity", "sum"), Number_of_Records=("quantity", "count"))
            )
            loss_summary.columns = ["Product", "Total Quantity Spoilt/Withdrawn", "Number of Records"]
            loss_summary = loss_summary.sort_values("Total Quantity Spoilt/Withdrawn", ascending=False)
        else:
            loss_exp = pd.DataFrame({"Note": ["No spoilage or withdrawals in this period"]})
            loss_summary = pd.DataFrame({"Note": ["No spoilage or withdrawals in this period"]})

        # --- Products & Daily ---
        prod_exp = products.copy()
        prod_exp.columns = ["ID", "Product Name", "SKU", "Unit Price (₦)", "Current Stock"]

        daily_exp = daily.copy()
        if not daily_exp.empty:
            daily_exp.columns = ["Date", "Transactions", "Total Units", "Total Amount (₦)"]

        orders = get_special_orders_df(exp_start.isoformat(), exp_end.isoformat())
        branch = get_branch_df(exp_start.isoformat(), exp_end.isoformat())

        orders_exp = orders.copy()
        if not orders_exp.empty:
            orders_exp = orders_exp[
                ["unique_id", "order_date", "customer_name", "product_description",
                 "quantity", "unit_price", "amount", "status", "notes", "created_at"]
            ]
            orders_exp.columns = [
                "Unique ID", "Date", "Customer", "Product / Description", "Qty",
                "Unit Price (₦)", "Amount (₦)", "Status", "Notes", "Recorded At",
            ]

        branch_exp = branch.copy()
        if not branch_exp.empty:
            branch_exp = branch_exp[
                ["unique_id", "record_date", "product_name", "qty_taken", "qty_sold",
                 "qty_remaining", "amount_sold", "unit_price", "notes", "created_at"]
            ]
            branch_exp.columns = [
                "Unique ID", "Date", "Product", "Produced/Sent to Punch", "Sold at Punch",
                "Remaining", "Amount Sold (₦)", "Punch Selling Price (₦)", "Notes", "Recorded At",
            ]

        raw_mats = get_raw_materials_df()
        raw_moves = get_raw_movements_df(exp_start.isoformat(), exp_end.isoformat())
        raw_exp = raw_mats.copy()
        if not raw_exp.empty:
            raw_exp = raw_exp[["name", "unit", "current_qty", "reorder_level", "unit_cost", "notes"]]
            raw_exp.columns = ["Material", "Unit", "In Stock", "Reorder Level", "Unit Cost (₦)", "Notes"]
        raw_moves_exp = raw_moves.copy()
        if not raw_moves_exp.empty:
            raw_moves_exp = raw_moves_exp[
                ["unique_id", "move_date", "material_name", "movement_type",
                 "quantity", "qty_before", "qty_after", "unit_cost", "total_cost", "notes"]
            ]
            raw_moves_exp.columns = [
                "Unique ID", "Date", "Material", "Type", "Qty", "Before", "After",
                "Unit Cost", "Total Cost", "Notes",
            ]

        exp_df = get_expenses_df(exp_start.isoformat(), exp_end.isoformat())
        sal_df = get_salary_payments_df(exp_start.isoformat(), exp_end.isoformat())
        staff_df = get_staff_df()

        excel_bytes = to_excel_bytes({
            "Sales": sales_exp if not sales_exp.empty else pd.DataFrame({"Note": ["No sales in range"]}),
            "Stock Movements": moves_exp if not moves_exp.empty else pd.DataFrame({"Note": ["No movements in range"]}),
            "Spoilage & Withdrawals": loss_exp,
            "Loss Summary by Product": loss_summary,
            "Special Orders": orders_exp if not orders_exp.empty else pd.DataFrame({"Note": ["No special orders"]}),
            "Punch Branch": branch_exp if not branch_exp.empty else pd.DataFrame({"Note": ["No branch records"]}),
            "Raw Materials": raw_exp if not raw_exp.empty else pd.DataFrame({"Note": ["No raw materials"]}),
            "Raw Material Movements": raw_moves_exp if not raw_moves_exp.empty else pd.DataFrame({"Note": ["No raw movements"]}),
            "Expenses": exp_df if not exp_df.empty else pd.DataFrame({"Note": ["No expenses"]}),
            "Salary Payments": sal_df if not sal_df.empty else pd.DataFrame({"Note": ["No salary payments"]}),
            "Staff": staff_df if not staff_df.empty else pd.DataFrame({"Note": ["No staff"]}),
            "Products": prod_exp,
            "Daily Totals": daily_exp if not daily_exp.empty else pd.DataFrame({"Note": ["No daily totals"]}),
        })

        filename = f"Funlex_Export_{exp_start}_{exp_end}.xlsx"
        st.download_button(
            label="Download Excel File",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.success("Excel file ready — click the download button above.")
        st.caption(
            "Sheets: Sales · Stock · Spoilage · Special Orders · Punch Branch · Products · Daily Totals"
        )


# ─────────────────────────────────────────────
# SETTINGS
# ─────────────────────────────────────────────

# ─────────────────────────────────────────────
# EXPENSES
# ─────────────────────────────────────────────
elif page == "Expenses":
    if not is_admin:
        st.error("Admin only.")
        st.stop()
    st.header("Daily Expenses")
    st.caption("Record every business cost: rent, fuel, packaging, utilities, transport, etc.")

    EXP_CATS = [
        "Rent", "Electricity", "Fuel / Generator", "Transport", "Packaging",
        "Ingredients (cash purchase)", "Repairs", "Marketing", "Miscellaneous", "Other",
    ]
    tab1, tab2 = st.tabs(["Record Expense", "Expense History"])

    with tab1:
        with st.form("exp_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                ed = st.date_input("Date", value=date.today())
                cat = st.selectbox("Category", EXP_CATS)
                branch = st.selectbox("Branch", ["Main", "Punch", "Shared"])
            with c2:
                amt = st.number_input("Amount (₦)", min_value=0.0, step=100.0, value=0.0)
                pay = st.selectbox("Payment method", ["Cash", "Transfer", "POS", "Other"])
            desc = st.text_input("Description", placeholder="e.g. diesel for generator")
            notes = st.text_input("Notes (optional)")
            if st.form_submit_button("Save Expense", type="primary", use_container_width=True):
                uid, err = add_expense(ed.isoformat(), cat, desc, amt, branch, pay, notes)
                if err:
                    st.error(err)
                else:
                    st.success(f"Expense saved · **{uid}**")
                    st.rerun()

    with tab2:
        f1, f2 = st.columns(2)
        with f1:
            m = st.selectbox("Month filter", ["All", "This month"], key="exp_mf")
        with f2:
            br = st.selectbox("Branch filter", ["All", "Main", "Punch", "Shared"], key="exp_bf")
        if m == "This month":
            t = date.today()
            s0, s1 = t.replace(day=1).isoformat(), t.isoformat()
        else:
            s0 = s1 = None
        exp = get_expenses_df(s0, s1, br if br != "All" else None)
        if exp.empty:
            st.info("No expenses found.")
        else:
            st.metric("Total in view", f"₦{exp['amount'].sum():,.0f}")
            show = exp[["id", "unique_id", "expense_date", "category", "description", "amount", "branch", "payment_method", "notes"]].copy()
            show.columns = ["ID", "UID", "Date", "Category", "Description", "Amount", "Branch", "Method", "Notes"]
            st.dataframe(show, use_container_width=True, hide_index=True)
            del_id = st.number_input("Delete expense ID", min_value=0, step=1, value=0, key="del_exp")
            if st.button("Delete expense") and del_id > 0:
                ok, err = delete_expense(int(del_id))
                if ok:
                    st.success("Deleted.")
                    st.rerun()
                else:
                    st.error(err)


# ─────────────────────────────────────────────
# STAFF & SALARIES
# ─────────────────────────────────────────────
elif page == "Staff & Salaries":
    if not is_admin:
        st.error("Admin only.")
        st.stop()
    st.header("Staff & Salaries")
    st.caption("Keep a staff list with monthly salaries and record salary payments.")

    tab1, tab2, tab3 = st.tabs(["Staff list", "Add / Edit staff", "Record salary payment"])

    staff = get_staff_df()
    with tab1:
        if staff.empty:
            st.info("No staff yet. Add people in the next tab.")
        else:
            active = staff[staff["is_active"] == 1]
            st.metric("Active staff", len(active))
            st.metric("Total monthly salaries (active)", f"₦{active['monthly_salary'].sum():,.0f}" if not active.empty else "₦0")
            show = staff[["id", "full_name", "role_title", "branch", "phone", "monthly_salary", "start_date", "is_active"]].copy()
            show["is_active"] = show["is_active"].map({1: "Active", 0: "Inactive"})
            show.columns = ["ID", "Name", "Role", "Branch", "Phone", "Monthly salary (₦)", "Start date", "Status"]
            st.dataframe(show, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("Add staff")
        with st.form("add_staff_form", clear_on_submit=True):
            a1, a2 = st.columns(2)
            with a1:
                nm = st.text_input("Full name *")
                role = st.text_input("Role / job title", placeholder="e.g. Baker, Sales rep")
                br = st.selectbox("Branch", ["Main", "Punch"], key="st_br")
            with a2:
                phone = st.text_input("Phone")
                sal = st.number_input("Monthly salary (₦)", min_value=0.0, step=1000.0, value=0.0)
                sd = st.date_input("Start date", value=date.today())
            notes = st.text_input("Notes", key="st_notes")
            if st.form_submit_button("Add staff", type="primary"):
                ok, err = add_staff(nm, role, br, phone, sal, sd.isoformat(), notes)
                if ok:
                    st.success(f"Added **{nm.strip()}**")
                    st.rerun()
                else:
                    st.error(err)

        st.divider()
        st.subheader("Edit staff")
        staff = get_staff_df()
        if not staff.empty:
            smap = {f"{r['full_name']} (ID {r['id']})": r["id"] for _, r in staff.iterrows()}
            sel = st.selectbox("Select staff", list(smap.keys()), key="edit_staff_sel")
            sid = smap[sel]
            row = staff[staff["id"] == sid].iloc[0]
            e1, e2 = st.columns(2)
            with e1:
                en = st.text_input("Name", value=row["full_name"], key="es_n")
                er = st.text_input("Role", value=row["role_title"] or "", key="es_r")
                eb = st.selectbox("Branch", ["Main", "Punch"], index=0 if row["branch"] == "Main" else 1, key="es_b")
            with e2:
                ep = st.text_input("Phone", value=row["phone"] or "", key="es_p")
                esal = st.number_input("Monthly salary", min_value=0.0, value=float(row["monthly_salary"]), step=1000.0, key="es_s")
                active = st.checkbox("Active", value=bool(row["is_active"]), key="es_a")
            esd = st.text_input("Start date (YYYY-MM-DD)", value=row["start_date"] or "", key="es_d")
            enotes = st.text_input("Notes", value=row["notes"] or "", key="es_no")
            if st.button("Save staff changes", type="primary"):
                ok, err = update_staff(sid, en, er, eb, ep, esal, esd, active, enotes)
                if ok:
                    st.success("Updated.")
                    st.rerun()
                else:
                    st.error(err)

    with tab3:
        staff = get_staff_df(active_only=True)
        if staff.empty:
            st.warning("Add active staff first.")
        else:
            with st.form("sal_pay_form", clear_on_submit=True):
                smap = {f"{r['full_name']} — ₦{r['monthly_salary']:,.0f}/mo": r["id"] for _, r in staff.iterrows()}
                label = st.selectbox("Staff", list(smap.keys()))
                sid = smap[label]
                default_sal = float(staff[staff["id"] == sid].iloc[0]["monthly_salary"])
                c1, c2 = st.columns(2)
                with c1:
                    pd_ = st.date_input("Payment date", value=date.today())
                    period = st.text_input("Period", value=date.today().strftime("%B %Y"), placeholder="e.g. October 2026")
                with c2:
                    amt = st.number_input("Amount paid (₦)", min_value=0.0, step=1000.0, value=default_sal)
                notes = st.text_input("Notes", key="sal_notes")
                pay_m = st.selectbox("Payment method", ["Cash", "Transfer"], key="sal_pay")
                if st.form_submit_button("Record payment", type="primary", use_container_width=True):
                    uid, err = add_salary_payment(sid, pd_.isoformat(), amt, period, notes, payment_method=pay_m)
                    if err:
                        st.error(err)
                    else:
                        st.success(f"Salary payment saved · **{uid}**")
                        st.rerun()

            st.divider()
            st.subheader("Recent salary payments")
            pays = get_salary_payments_df()
            if pays.empty:
                st.caption("No payments recorded yet.")
            else:
                show = pays.head(20)[["unique_id", "pay_date", "staff_name", "period_label", "amount", "notes"]].copy()
                show.columns = ["UID", "Date", "Staff", "Period", "Amount (₦)", "Notes"]
                st.dataframe(show, use_container_width=True, hide_index=True)



# ─────────────────────────────────────────────
# ACCOUNTS (Phase 1)
# ─────────────────────────────────────────────
elif page == "Accounts":
    if not is_admin:
        st.error("Admin only.")
        st.stop()

    st.header("Accounts")
    st.caption(
        "Sales → Dr Cash / Cr Sales – Main. "
        "Raw material purchases → Dr Inventory / Cr Cash. "
        "Ledger shows each account; trial balance checks that books balance."
    )

    # Shared period filter
    today = date.today()
    period = st.selectbox(
        "Period",
        ["This month", "Last 30 days", "All time"],
        key="acc_period",
    )
    if period == "This month":
        p0, p1 = today.replace(day=1).isoformat(), today.isoformat()
    elif period == "Last 30 days":
        from datetime import timedelta
        p0 = (today - timedelta(days=30)).isoformat()
        p1 = today.isoformat()
    else:
        p0 = p1 = None

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
        ["Chart of accounts", "Journals", "Ledger", "Trial balance", "Payables & Loans", "Opening balances", "Reset accounts"]
    )

    with tab1:
        acc = get_accounts_df(active_only=False)
        if acc.empty:
            st.warning("No accounts found. Restart the app to seed defaults.")
        else:
            show = acc[["code", "name", "account_type", "is_active"]].copy()
            show["is_active"] = show["is_active"].map({1: "Active", 0: "Inactive"})
            show.columns = ["Code", "Account name", "Type", "Status"]
            st.dataframe(show, use_container_width=True, hide_index=True)

    with tab2:
        journals = get_journals_df(p0, p1)
        st.metric("Journal entries", len(journals))
        if journals.empty:
            st.info(
                "No journals yet. Record a **sale** or a **raw material purchase** "
                "(with total amount paid) to create entries."
            )
        else:
            show = journals[["unique_id", "entry_date", "description", "source_type", "source_id"]].copy()
            show.columns = ["Journal ID", "Date", "Description", "Source", "Source ID"]
            st.dataframe(show, use_container_width=True, hide_index=True)

            st.subheader("All journal lines")
            lines = get_journal_lines_df()
            if p0 and p1 and not lines.empty:
                lines = lines[(lines["entry_date"] >= p0) & (lines["entry_date"] <= p1)]
            if not lines.empty:
                lshow = lines[
                    ["entry_date", "unique_id", "account_code", "account_name", "debit", "credit", "description"]
                ].copy()
                lshow.columns = ["Date", "Journal", "Code", "Account", "Debit (₦)", "Credit (₦)", "Narration"]
                st.dataframe(lshow, use_container_width=True, hide_index=True)
                tot_dr = float(lshow["Debit (₦)"].sum())
                tot_cr = float(lshow["Credit (₦)"].sum())
                c1, c2 = st.columns(2)
                c1.metric("Total debits", f"₦{tot_dr:,.0f}")
                c2.metric("Total credits", f"₦{tot_cr:,.0f}")

    with tab3:
        st.subheader("Account ledger")
        st.caption(
            "Pick one account to see every movement and the running balance. "
            "Example: Cash shows −60,000 for flour and +1,000 for a sale."
        )
        accounts = get_accounts_df(active_only=True)
        if accounts.empty:
            st.warning("No accounts.")
        else:
            labels = {
                f"{r['code']} — {r['name']} ({r['account_type']})": r["code"]
                for _, r in accounts.iterrows()
            }
            sel = st.selectbox("Account", list(labels.keys()), key="ledger_acc")
            code = labels[sel]
            led = get_ledger_for_account(code, p0, p1)
            if led.empty:
                st.info("No movements on this account in the selected period.")
            else:
                show = led[
                    ["entry_date", "unique_id", "description", "source_type", "debit", "credit", "balance"]
                ].copy()
                show.columns = [
                    "Date", "Journal", "Description", "Source",
                    "Debit (₦)", "Credit (₦)", "Balance (₦)",
                ]
                st.dataframe(show, use_container_width=True, hide_index=True)
                last_bal = float(led["balance"].iloc[-1])
                st.metric("Closing balance", f"₦{last_bal:,.2f}")
                atype = accounts[accounts["code"] == code].iloc[0]["account_type"]
                if atype in ("Asset", "Expense"):
                    st.caption("Positive balance = debit balance (normal for assets/expenses).")
                else:
                    st.caption("Positive balance = credit balance (normal for income / liabilities).")

    with tab4:
        st.subheader("Trial balance")
        st.caption(
            "Totals per account for the period. "
            "Total debit column should equal total credit column."
        )
        tb = get_trial_balance(p0, p1)
        if tb.empty:
            st.info("No journal activity in this period.")
        else:
            show = tb[
                ["account_code", "account_name", "account_type",
                 "total_debit", "total_credit", "balance_debit", "balance_credit"]
            ].copy()
            show.columns = [
                "Code", "Account", "Type",
                "Period debit (₦)", "Period credit (₦)",
                "Balance Dr (₦)", "Balance Cr (₦)",
            ]
            st.dataframe(show, use_container_width=True, hide_index=True)
            sum_dr = float(tb["total_debit"].sum())
            sum_cr = float(tb["total_credit"].sum())
            bal_dr = float(tb["balance_debit"].sum())
            bal_cr = float(tb["balance_credit"].sum())
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Sum of debits", f"₦{sum_dr:,.0f}")
            m2.metric("Sum of credits", f"₦{sum_cr:,.0f}")
            m3.metric("Balance column Dr", f"₦{bal_dr:,.0f}")
            m4.metric("Balance column Cr", f"₦{bal_cr:,.0f}")
            if abs(sum_dr - sum_cr) < 1:
                st.success("Trial balance is in balance (period debits = period credits).")
            else:
                st.error("Trial balance is out of balance — check journal entries.")
            # Simple income vs expense snapshot
            income = tb[tb["account_type"] == "Income"]
            expense = tb[tb["account_type"] == "Expense"]
            inc_amt = float(income["total_credit"].sum() - income["total_debit"].sum()) if not income.empty else 0.0
            exp_amt = float(expense["total_debit"].sum() - expense["total_credit"].sum()) if not expense.empty else 0.0
            st.divider()
            st.markdown("**From trial balance (this period)**")
            i1, i2, i3 = st.columns(3)
            i1.metric("Income (sales etc.)", f"₦{inc_amt:,.0f}")
            i2.metric("Expenses (posted)", f"₦{exp_amt:,.0f}")
            i3.metric("Income − expenses", f"₦{inc_amt - exp_amt:,.0f}")
            try:
                cogs_s = get_cogs_summary(p0, p1)
                st.markdown("**Gross profit (Sales − COGS)**")
                g1, g2, g3 = st.columns(3)
                g1.metric("Sales (posted)", f"₦{cogs_s['sales']:,.0f}")
                g2.metric("COGS", f"₦{cogs_s['cogs']:,.0f}")
                g3.metric("Gross profit", f"₦{cogs_s['gross_profit']:,.0f}")
                st.caption(
                    "COGS posts when a product has **Cost Price** set and a sale is recorded. "
                    "Set cost prices under **Products**."
                )
            except Exception:
                pass

    with tab5:
        st.subheader("Payables & Loans")
        st.caption(
            "Pay suppliers for goods bought on credit, record loans received, and repay loans. "
            "These post to Accounts payable (2000) and Loans payable (2100)."
        )
        sub1, sub2, sub3 = st.tabs(["Pay supplier", "Loan received", "Repay loan"])

        with sub1:
            st.markdown("**Reduce what you owe suppliers** (Dr Payable · Cr Cash/Bank)")
            with st.form("supplier_pay_form", clear_on_submit=True):
                sp1, sp2 = st.columns(2)
                with sp1:
                    sp_date = st.date_input("Payment date", value=date.today(), key="sp_date")
                    sp_supplier = st.text_input("Supplier name", placeholder="e.g. Flour supplier")
                with sp2:
                    sp_amt = st.number_input("Amount paid (₦)", min_value=0.0, step=100.0, key="sp_amt")
                    sp_pay = st.selectbox("Paid from", ["Cash", "Transfer"], key="sp_pay")
                sp_notes = st.text_input("Notes", key="sp_notes")
                if st.form_submit_button("Record supplier payment", type="primary"):
                    if sp_amt <= 0:
                        st.error("Enter amount.")
                    else:
                        uid = f"SPAY-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
                        juid, err = auto_journal_supplier_payment(
                            sp_date.isoformat(), sp_amt, sp_supplier, uid, payment=sp_pay
                        )
                        if err:
                            st.error(err)
                        else:
                            st.success(f"Supplier payment posted · Journal **{juid}**")
                            st.rerun()

        with sub2:
            st.markdown("**Money borrowed** (Dr Cash/Bank · Cr Loans payable)")
            with st.form("loan_recv_form", clear_on_submit=True):
                lr1, lr2 = st.columns(2)
                with lr1:
                    lr_date = st.date_input("Date received", value=date.today(), key="lr_date")
                    lr_lender = st.text_input("Lender", placeholder="e.g. Bank, family")
                with lr2:
                    lr_amt = st.number_input("Amount received (₦)", min_value=0.0, step=1000.0, key="lr_amt")
                    lr_into = st.selectbox("Received into", ["Cash", "Transfer"], key="lr_into")
                if st.form_submit_button("Record loan received", type="primary"):
                    if lr_amt <= 0:
                        st.error("Enter amount.")
                    else:
                        uid = f"LOAN-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
                        juid, err = auto_journal_loan_received(
                            lr_date.isoformat(), lr_amt, lr_lender, uid, payment=lr_into
                        )
                        if err:
                            st.error(err)
                        else:
                            st.success(f"Loan recorded · Journal **{juid}**")
                            st.rerun()

        with sub3:
            st.markdown("**Repay loan** (Dr Loans payable + interest · Cr Cash/Bank)")
            with st.form("loan_pay_form", clear_on_submit=True):
                lp1, lp2 = st.columns(2)
                with lp1:
                    lp_date = st.date_input("Payment date", value=date.today(), key="lp_date")
                    lp_lender = st.text_input("Lender", key="lp_lender")
                    lp_prin = st.number_input("Principal (₦)", min_value=0.0, step=1000.0, key="lp_prin")
                with lp2:
                    lp_int = st.number_input("Interest (₦)", min_value=0.0, step=100.0, value=0.0, key="lp_int")
                    lp_from = st.selectbox("Paid from", ["Cash", "Transfer"], key="lp_from")
                st.caption(f"Total cash out: ₦{(0):,.0f} — enter principal and interest above.")
                if st.form_submit_button("Record loan repayment", type="primary"):
                    if lp_prin + lp_int <= 0:
                        st.error("Enter principal and/or interest.")
                    else:
                        uid = f"LREP-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
                        juid, err = auto_journal_loan_repay(
                            lp_date.isoformat(), lp_prin, lp_int, lp_lender, uid, payment=lp_from
                        )
                        if err:
                            st.error(err)
                        else:
                            st.success(f"Repayment posted · Journal **{juid}** · Total ₦{lp_prin+lp_int:,.0f}")
                            st.rerun()

        # Quick balances
        st.divider()
        try:
            tb = get_trial_balance(None, None)
            pay = tb[tb["account_code"] == "2000"] if not tb.empty else pd.DataFrame()
            loan = tb[tb["account_code"] == "2100"] if not tb.empty else pd.DataFrame()
            pay_bal = float(pay["balance_credit"].sum() - pay["balance_debit"].sum()) if not pay.empty else 0.0
            # from trial balance columns
            if not pay.empty:
                pay_bal = float(pay.iloc[0]["total_credit"] - pay.iloc[0]["total_debit"])
            if not loan.empty:
                loan_bal = float(loan.iloc[0]["total_credit"] - loan.iloc[0]["total_debit"])
            else:
                loan_bal = 0.0
            m1, m2 = st.columns(2)
            m1.metric("Accounts payable (owed to suppliers)", f"₦{max(0, pay_bal):,.0f}")
            m2.metric("Loans payable (outstanding)", f"₦{max(0, loan_bal):,.0f}")
        except Exception:
            pass

    with tab6:
        st.subheader("Opening balances")
        st.caption(
            "Use once when starting the books. Enter what Funlex already had "
            "(cash, stock value, debts). Posts one balanced journal to Opening balance equity."
        )
        with st.form("opening_bal_form"):
            ob_date = st.date_input("Opening date", value=date.today(), key="ob_date")
            o1, o2 = st.columns(2)
            with o1:
                ob_cash = st.number_input("Cash on hand (₦)", min_value=0.0, step=1000.0, key="ob_cash")
                ob_bank = st.number_input("Bank / transfer balance (₦)", min_value=0.0, step=1000.0, key="ob_bank")
                ob_raw = st.number_input("Raw materials inventory value (₦)", min_value=0.0, step=1000.0, key="ob_raw")
                ob_fg = st.number_input("Finished goods inventory value (₦)", min_value=0.0, step=1000.0, key="ob_fg")
            with o2:
                ob_pay = st.number_input("Accounts payable — suppliers owed (₦)", min_value=0.0, step=1000.0, key="ob_pay")
                ob_loan = st.number_input("Loans payable (₦)", min_value=0.0, step=1000.0, key="ob_loan")
                ob_notes = st.text_input("Notes", key="ob_notes")
            if st.form_submit_button("Post opening balances", type="primary"):
                juid, err = post_opening_balances(
                    ob_date.isoformat(),
                    {
                        "cash": ob_cash,
                        "bank": ob_bank,
                        "raw_inventory": ob_raw,
                        "finished_goods": ob_fg,
                        "payables": ob_pay,
                        "loans": ob_loan,
                    },
                    notes=ob_notes,
                )
                if err:
                    st.error(err)
                else:
                    st.success(f"Opening balances posted · **{juid}**")
                    st.rerun()

    with tab7:
        st.subheader("Reset accounts")
        st.warning(
            "This **deletes all journal entries and journal lines**. "
            "Sales, stock, expenses and other operational data are **not** deleted. "
            "The chart of accounts stays. Use this if you want to start the books from zero."
        )
        st.caption("Type **RESET** below, then confirm.")
        confirm = st.text_input("Type RESET to enable wipe", key="acc_reset_confirm")
        if st.button("Wipe all journals", type="primary", disabled=(confirm.strip().upper() != "RESET")):
            ok, err = reset_accounts(confirm)
            if ok:
                st.success("All journals cleared. New sales and purchases will post fresh entries.")
                st.rerun()
            else:
                st.error(err or "Could not reset.")


elif page == "Settings":
    if not is_admin:
        st.error("Only Admin can access Settings.")
        st.stop()

    st.header("Settings")
    st.subheader("Access Codes (Admin only)")
    st.caption("Set separate codes for Admin and Sales Rep.")

    with st.form("change_admin_code"):
        st.markdown("**Admin access code** (full control)")
        admin_new = st.text_input("New Admin Code", type="password", key="admin_new")
        admin_confirm = st.text_input("Confirm Admin Code", type="password", key="admin_confirm")
        if st.form_submit_button("Update Admin Code", type="primary"):
            if not admin_new or len(admin_new) < 4:
                st.error("Code must be at least 4 characters.")
            elif admin_new != admin_confirm:
                st.error("Codes do not match.")
            else:
                change_role_code("admin", admin_new)
                st.success("Admin code updated.")

    with st.form("change_sales_code"):
        st.markdown("**Sales Rep access code** (record sales only)")
        sales_new = st.text_input("New Sales Rep Code", type="password", key="sales_new")
        sales_confirm = st.text_input("Confirm Sales Rep Code", type="password", key="sales_confirm")
        if st.form_submit_button("Update Sales Rep Code", type="primary"):
            if not sales_new or len(sales_new) < 4:
                st.error("Code must be at least 4 characters.")
            elif sales_new != sales_confirm:
                st.error("Codes do not match.")
            else:
                change_role_code("sales", sales_new)
                st.success("Sales Rep code updated.")

    st.divider()
    st.subheader("Reset All Stock Quantities")
    st.caption("Sets every product stock to 0. Sales history is kept. Admin only.")
    if st.button("Reset All Stocks to Zero", type="secondary"):
        reset_all_stocks()
        st.success("All product stocks set to 0.")
        st.rerun()

    st.divider()
    st.subheader("Reset Accounts (journals only)")
    st.caption(
        "Clears all journal entries so accounting can start fresh. "
        "Does not delete sales, stock, expenses or staff records."
    )
    conf2 = st.text_input("Type RESET to confirm accounts wipe", key="set_acc_reset")
    if st.button("Wipe journals", type="secondary", disabled=(conf2.strip().upper() != "RESET")):
        ok, err = reset_accounts(conf2)
        if ok:
            st.success("Journals cleared.")
            st.rerun()
        else:
            st.error(err or "Failed.")

    st.divider()
    st.subheader("Database Info")
    products = get_all_products()
    sales = get_sales_df()
    moves = get_stock_movements_df()
    st.write(
        f"Products: **{len(products)}**  ·  "
        f"Sales: **{len(sales)}**  ·  "
        f"Stock movements: **{len(moves)}**"
    )
    st.caption(f"Database file: `{DB_PATH.name}`")

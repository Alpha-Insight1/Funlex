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
            current_stock REAL NOT NULL DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
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

    cur.execute("SELECT COUNT(*) FROM products")
    if cur.fetchone()[0] == 0:
        for name, sku, price, stock in DEFAULT_PRODUCTS:
            cur.execute(
                "INSERT INTO products (name, sku, unit_price, current_stock) VALUES (?, ?, ?, ?)",
                (name, sku, price, stock),
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
        "SELECT id, name, sku, unit_price, current_stock FROM products ORDER BY name",
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


def add_sale(sale_date, product_id, units_sold, notes="", unit_price=None):
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
        return uid, None
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()


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


def update_product_price(product_id, new_price):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE products SET unit_price = ? WHERE id = ?", (new_price, product_id))
    conn.commit()
    conn.close()


def add_new_product(name, sku, price, stock):
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO products (name, sku, unit_price, current_stock) VALUES (?, ?, ?, ?)",
            (name, sku, price, stock),
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
    """Delete a sale and restore the stock that was deducted."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT product_id, units_sold FROM sales WHERE id = ?", (sale_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return False, "Sale not found"
    product_id, units = row["product_id"], float(row["units_sold"])
    try:
        # Restore stock
        cur.execute(
            "UPDATE products SET current_stock = current_stock + ? WHERE id = ?",
            (units, product_id),
        )
        cur.execute("DELETE FROM sales WHERE id = ?", (sale_id,))
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


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
        return uid, None
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()


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
        conn.commit()
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def delete_special_order(order_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM special_orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()


# ─── Punch Branch ───
def add_branch_record(record_date, product_id, qty_taken, qty_sold, unit_price=None, notes=""):
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
        return uid, None
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()


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
        """SELECT id, name, sku, unit_price, current_stock FROM products
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


def add_raw_movement(move_date, material_id, movement_type, quantity, unit_cost=None, notes=""):
    """
    movement_type:
      purchase / received  -> increase
      used_in_production   -> decrease
      waste_spoilage       -> decrease
      adjustment_in        -> increase
      adjustment_out       -> decrease
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
    cost = float(unit_cost) if unit_cost is not None and float(unit_cost) >= 0 else float(mat["unit_cost"])
    total = qty * cost if movement_type in increase_types else 0.0
    uid = f"RAW-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:8].upper()}"
    conn = get_conn()
    cur = conn.cursor()
    try:
        cur.execute(
            "UPDATE raw_materials SET current_qty = ? WHERE id = ?",
            (after, material_id),
        )
        # Update unit cost on purchase if provided
        if movement_type in increase_types and unit_cost is not None and float(unit_cost) > 0:
            cur.execute(
                "UPDATE raw_materials SET unit_cost = ? WHERE id = ?",
                (float(unit_cost), material_id),
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
        return uid, None
    except Exception as e:
        conn.rollback()
        return None, str(e)
    finally:
        conn.close()


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
    """Reverse movement effect and delete record."""
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
    increase_types = ("purchase", "received", "adjustment_in")
    # Reverse
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
        return True, None
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


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
    col1, col2, col3 = st.columns([1, 1.35, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="background:#FFFFFF; border:1px solid #EEE8EF; border-radius:20px;
                        padding:2rem 1.75rem 1.75rem; box-shadow:0 8px 32px rgba(28,28,40,0.08);
                        text-align:center;">
            """,
            unsafe_allow_html=True,
        )
        logo_cols = st.columns([1, 2, 1])
        with logo_cols[1]:
            show_logo(160)
        st.markdown(
            """
            <div style="margin-top:0.5rem; margin-bottom:1rem;">
                <h1 style="color:#C2185B; font-size:1.55rem; font-weight:800; margin:0.35rem 0 0.2rem 0;
                           letter-spacing:0.4px;">FUNLEX CAKES AND CREAM</h1>
                <p style="color:#AD1457; font-size:0.88rem; font-style:italic; margin:0 0 0.35rem 0;
                          letter-spacing:0.8px; font-weight:500;">The Taste Is The Difference</p>
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
        st.markdown("</div>", unsafe_allow_html=True)


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
    st.header("Dashboard")
    products = get_all_products()
    sales = get_sales_df()
    daily = get_daily_totals()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Products", len(products))
    c2.metric("Sales Records", len(sales))
    today_str = date.today().isoformat()
    today_sales = sales[sales["sale_date"] == today_str] if not sales.empty else pd.DataFrame()
    c3.metric("Today's Revenue (₦)", f"{today_sales['amount'].sum():,.0f}" if not today_sales.empty else "0")
    low_stock = products[products["current_stock"] <= 5]
    try:
        raws = get_raw_materials_df()
        low_raw = len(raws[raws["current_qty"] <= raws["reorder_level"]]) if not raws.empty else 0
    except Exception:
        low_raw = 0
    c4.metric("Low product stock", len(low_stock))
    # show ingredients needing reorder under metrics if any


    if low_raw:
        st.warning(f"**{low_raw} raw material(s)** at or below reorder level — check **Raw Materials**.")
    st.divider()
    left, right = st.columns(2)

    with left:
        st.subheader("Low / Out of Stock (finished products)")
        if low_stock.empty:
            st.success("All products are well stocked.")
        else:
            st.dataframe(
                low_stock[["name", "current_stock", "unit_price"]].rename(
                    columns={"name": "Product", "current_stock": "Stock", "unit_price": "Price (₦)"}
                ),
                use_container_width=True,
                hide_index=True,
            )

    with right:
        st.subheader("Recent Daily Totals")
        if daily.empty:
            st.info("No sales recorded yet.")
        else:
            disp = daily.head(7).copy()
            disp["total_amount"] = disp["total_amount"].apply(lambda x: f"₦{x:,.0f}")
            st.dataframe(
                disp.rename(
                    columns={
                        "date": "Date",
                        "transactions": "Txns",
                        "total_units": "Units",
                        "total_amount": "Revenue",
                    }
                ),
                use_container_width=True,
                hide_index=True,
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

            if prod:
                amt = units * sell_price
                close = prod["current_stock"] - units
                st.info(
                    f"**Stock available:** {prod['current_stock']:.0f}  ·  "
                    f"**List price:** ₦{default_price:,.0f}  ·  "
                    f"**This sale price:** ₦{sell_price:,.0f}  ·  "
                    f"**Amount:** ₦{amt:,.0f}  ·  "
                    f"**Stock after sale:** {close:.0f}"
                )

            if st.form_submit_button("Save Sale", type="primary", use_container_width=True):
                uid, err = add_sale(
                    sale_date.isoformat(), product_id, units, notes, unit_price=sell_price
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
        "Record purchases, production usage, and waste so you know what to reorder."
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
            m3.metric("Stock value (₦)", f"{total_val:,.0f}")

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
                        f"{r['name']}  ·  {r['current_qty']:.2f} {r['unit']}": r["id"]
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
                    qty = st.number_input("Quantity", min_value=0.0, step=0.5, value=1.0)
                mat = get_raw_material_by_id(material_id)
                default_cost = float(mat["unit_cost"]) if mat else 0.0
                cost = st.number_input(
                    "Unit cost (₦) — for purchases",
                    min_value=0.0,
                    step=50.0,
                    value=default_cost,
                    help="Optional. Updates average reference cost on purchase.",
                )
                notes = st.text_input("Notes", placeholder="e.g. from market, used for cake batch")
                if st.form_submit_button("Save Movement", type="primary", use_container_width=True):
                    uid, err = add_raw_movement(
                        move_date.isoformat(), material_id, mtype, qty, cost, notes
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
                remaining = max(0, qty_taken - qty_sold)
                amt = qty_sold * punch_price
                st.info(
                    f"**Remaining at Punch:** {remaining:.0f}  ·  "
                    f"**Amount Sold at Punch:** ₦{amt:,.0f}  "
                    f"(Main catalogue price: ₦{default_price:,.0f} — Punch price can differ)"
                )
                if st.form_submit_button("Save Branch Record", type="primary", use_container_width=True):
                    uid, err = add_branch_record(
                        br_date.isoformat(), product_id, qty_taken, qty_sold, punch_price, notes
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
            disp["unit_price"] = disp["unit_price"].apply(lambda x: f"₦{x:,.0f}")
            disp.columns = ["ID", "Product", "SKU", "Unit Price", "Current Stock"]
            st.dataframe(disp, use_container_width=True, hide_index=True)
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
                price = st.number_input("Unit Price (₦) *", min_value=0.0, step=50.0, value=500.0)
                stock = st.number_input("Opening Stock", min_value=0.0, step=1.0, value=0.0)
            if st.form_submit_button("Add Product", type="primary", use_container_width=True):
                if not name.strip():
                    st.error("Product name is required.")
                elif price <= 0:
                    st.error("Price must be greater than 0.")
                else:
                    ok, err = add_new_product(name.strip(), sku.strip(), price, stock)
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
                "Unit Price (₦)",
                min_value=0.0,
                value=float(current["unit_price"]),
                step=50.0,
                key="edit_pprice",
            )
            st.caption(f"Current stock: **{current['current_stock']:.0f}** (change stock in Stock Tracker)")

            b1, b2, b3 = st.columns(3)
            with b1:
                if st.button("Save Name & SKU", type="primary", key="btn_rename"):
                    ok, err = rename_product(pid, new_name, new_sku)
                    if ok:
                        if new_price > 0 and float(new_price) != float(current["unit_price"]):
                            update_product_price(pid, new_price)
                        st.success(f"Product updated to **{new_name.strip()}**")
                        st.rerun()
                    else:
                        st.error(err)
            with b2:
                if st.button("Update Price Only", key="btn_upd_price"):
                    if new_price <= 0:
                        st.error("Price must be greater than 0.")
                    else:
                        update_product_price(pid, new_price)
                        st.success(f"Price updated to ₦{new_price:,.0f}")
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

        excel_bytes = to_excel_bytes({
            "Sales": sales_exp if not sales_exp.empty else pd.DataFrame({"Note": ["No sales in range"]}),
            "Stock Movements": moves_exp if not moves_exp.empty else pd.DataFrame({"Note": ["No movements in range"]}),
            "Spoilage & Withdrawals": loss_exp,
            "Loss Summary by Product": loss_summary,
            "Special Orders": orders_exp if not orders_exp.empty else pd.DataFrame({"Note": ["No special orders"]}),
            "Punch Branch": branch_exp if not branch_exp.empty else pd.DataFrame({"Note": ["No branch records"]}),
            "Raw Materials": raw_exp if not raw_exp.empty else pd.DataFrame({"Note": ["No raw materials"]}),
            "Raw Material Movements": raw_moves_exp if not raw_moves_exp.empty else pd.DataFrame({"Note": ["No raw movements"]}),
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

import time
import json
import os
import sys
import logging
import threading
import requests
import pyodbc
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from concurrent.futures import ThreadPoolExecutor, as_completed
import re
import subprocess
import browser_cookie3
import customtkinter as ctk

# ============================================================
# تنظیمات پایه تم روشن مدرن (Light Enterprise Theme)
# ============================================================
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("green")

# ============================================================
# تنظیمات مسیرها - کار در PY و EXE
# ============================================================

def get_app_path():
    """دریافت مسیر اصلی برنامه - کار در EXE و PY"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

APP_PATH = get_app_path()
DESKTOP_PATH = os.path.join(os.path.expanduser("~"), "Desktop")

# --- Configuration ---
# نکته امنیتی: مقادیر محرمانه نسخه قبلی از کد حذف شده‌اند.
# اگر نسخه قبلی منتشر یا در اختیار دیگران قرار گرفته، رمز پایگاه داده، توکن API و کوکی‌ها را عوض کنید.
# تنظیم اتصال پایگاه داده را اینجا وارد کنید.
# قالب: DRIVER={SQL Server};SERVER=نام-سرور;DATABASE=نام-پایگاه-داده;UID=نام-کاربری;PWD=رمز-عبور
# مقادیر نمونه زیر را با اطلاعات واقعی خود جایگزین کنید.
DB_CONN = os.environ.get(
    "SYNC_DB_CONN",
    "DRIVER={SQL Server};SERVER=YOUR_SERVER;DATABASE=YOUR_DATABASE;UID=YOUR_USERNAME;PWD=YOUR_PASSWORD"
)

# --- API Configuration (Default values - will be overridden by config file if exists) ---
# آدرس API فروشگاه را در پنجره «تنظیمات اتصال به API» وارد کنید.
# این مقدار نمونه است و باید با آدرس API سایت خودتان جایگزین شود.
DEFAULT_BASE_URL = "https://YOUR-DOMAIN.example/api/shop/product"
BASE_URL = DEFAULT_BASE_URL

# ============================================================
# API Configuration با کوکی جدید
# ============================================================

# هدرهای ثابت
FIXED_HEADERS = {
    "Authorization": "Bearer YOUR_API_TOKEN_HERE",
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "X-Requested-With": "XMLHttpRequest",
    "X-CSRF-TOKEN": "YOUR_CSRF_TOKEN_HERE",
}

# کوکی پیش‌فرض
DEFAULT_COOKIE = ""  # کوکی نشست را در تنظیمات API وارد کنید؛ کوکی واقعی را در کد قرار ندهید.

# ساخت API_HEADERS کامل با کوکی
API_HEADERS = FIXED_HEADERS.copy()
API_HEADERS["Cookie"] = DEFAULT_COOKIE

# --- Config file path (کنار EXE یا PY) ---
CONFIG_FILE = os.path.join(APP_PATH, "api_config.json")

# --- File Paths (روی دسکتاپ) ---
REPO_FILE = os.path.join(DESKTOP_PATH, "stock_repo.json")
LAST_STOCK_FILE = os.path.join(DESKTOP_PATH, "last_webzi_stock.json")
LAST_DB_FILE = os.path.join(DESKTOP_PATH, "last_db_stock.json")

# ============================================================
# توابع کمکی و پردازش URL و API
# ============================================================

def safe_latin1(val):
    if not val:
        return ""
    return str(val).encode('latin-1', errors='ignore').decode('latin-1')

def derive_product_url(api_url):
    try:
        clean = str(api_url or "").split("?", 1)[0].rstrip("/")
        if "/api/shop/product" in clean:
            prefix = clean.split("/api/shop/product", 1)[0]
            return prefix + "/shop/products/"
        if "/shop/products" in clean:
            return clean.rstrip("/") + "/"
    except Exception:
        pass
    return str(api_url or "")

def derive_api_url(product_url):
    try:
        clean = str(product_url or "").strip().split("?", 1)[0].rstrip("/")
        marker = "/shop/products"
        pos = clean.find(marker)
        if pos >= 0:
            return clean[:pos] + "/api/shop/product"
    except Exception:
        pass
    return ""

def extract_csrf_from_cookie(cookie):
    try:
        m = re.search(r"(?:^|;\s*)CSRF-TOKEN=([^;]+)", str(cookie or ""))
        return safe_latin1(m.group(1)) if m else ""
    except Exception:
        return ""

def parse_api_config(text):
    """تجزیه متن ورودی و استخراج BASE_URL و API_HEADERS"""
    try:
        url_match = re.search(r'BASE_URL\s*=\s*["\'](https?://[^"\']+)["\']', text)
        if not url_match:
            return None, None, None, "❌ آدرس BASE_URL پیدا نشد!"
        
        base_url = url_match.group(1)
        
        headers_match = re.search(r'API_HEADERS\s*=\s*(\{[\s\S]*?\})', text)
        if not headers_match:
            return None, None, None, "❌ هدرهای API_HEADERS پیدا نشد!"
        
        headers_text = headers_match.group(1)
        headers_dict = eval(headers_text)
        
        cookie = headers_dict.get("Cookie", "")
        if not cookie:
            return None, None, None, "❌ کوکی در هدرها پیدا نشد!"
        
        csrf_token = headers_dict.get("X-CSRF-TOKEN", "")
        if not csrf_token:
            csrf_token = extract_csrf_from_cookie(cookie)
            
        return base_url, cookie, csrf_token, None
        
    except Exception as e:
        return None, None, None, f"❌ خطا در تجزیه اطلاعات: {str(e)}"

def load_config():
    """بارگذاری تنظیمات از فایل"""
    global BASE_URL, API_HEADERS, DEFAULT_COOKIE
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                if "BASE_URL" in config:
                    saved_url = config["BASE_URL"]
                    if re.search(r'[^\x00-\x7F]', saved_url):
                        BASE_URL = DEFAULT_BASE_URL
                    else:
                        BASE_URL = saved_url
                if "COOKIE" in config:
                    DEFAULT_COOKIE = config["COOKIE"]
                    API_HEADERS["Cookie"] = DEFAULT_COOKIE
                if "CSRF_TOKEN" in config:
                    FIXED_HEADERS["X-CSRF-TOKEN"] = config["CSRF_TOKEN"]
                    API_HEADERS["X-CSRF-TOKEN"] = config["CSRF_TOKEN"]
                logger.info(f"Config loaded from {CONFIG_FILE}")
                return True
        except Exception as e:
            logger.warning(f"Failed to load config: {e}")
    return False

def save_config(base_url, cookie, csrf_token=None):
    """ذخیره تنظیمات در فایل"""
    try:
        config = {
            "BASE_URL": base_url,
            "COOKIE": cookie
        }
        if csrf_token:
            config["CSRF_TOKEN"] = csrf_token
        
        config_dir = os.path.dirname(CONFIG_FILE)
        if config_dir and not os.path.exists(config_dir):
            os.makedirs(config_dir)
        
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        logger.info(f"Config saved to {CONFIG_FILE}")
        return True
    except PermissionError:
        logger.error(f"Permission denied! Cannot write to {CONFIG_FILE}")
        try:
            fallback_path = os.path.join(DESKTOP_PATH, "api_config.json")
            with open(fallback_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            logger.info(f"Config saved to fallback: {fallback_path}")
            return True
        except:
            return False
    except Exception as e:
        logger.error(f"Failed to save config: {e}")
        return False

def update_api_headers(cookie, csrf_token=None):
    """بروزرسانی هدرها با کوکی و توکن جدید"""
    global API_HEADERS
    API_HEADERS = FIXED_HEADERS.copy()
    API_HEADERS["Cookie"] = cookie
    if csrf_token:
        API_HEADERS["X-CSRF-TOKEN"] = csrf_token

def load_repo():
    if os.path.exists(REPO_FILE):
        try:
            with open(REPO_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_repo(repo):
    try:
        with open(REPO_FILE, 'w', encoding='utf-8') as f:
            json.dump(repo, f, ensure_ascii=False, indent=2)
    except:
        pass

def clean_expired_repo(repo):
    now = datetime.now()
    expired = []
    for code, data in repo.items():
        try:
            entry_time = datetime.fromisoformat(data.get("time", ""))
            if now - entry_time > timedelta(hours=100):
                expired.append(code)
        except:
            expired.append(code)
    for code in expired:
        del repo[code]
    if expired:
        save_repo(repo)
    return repo

def load_last_stock():
    if os.path.exists(LAST_STOCK_FILE):
        try:
            with open(LAST_STOCK_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_last_stock(stock_data):
    try:
        with open(LAST_STOCK_FILE, 'w', encoding='utf-8') as f:
            json.dump(stock_data, f, ensure_ascii=False, indent=2)
    except:
        pass

def load_last_db_stock():
    if os.path.exists(LAST_DB_FILE):
        try:
            with open(LAST_DB_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_last_db_stock(db_data):
    try:
        with open(LAST_DB_FILE, 'w', encoding='utf-8') as f:
            json.dump(db_data, f, ensure_ascii=False, indent=2)
    except:
        pass

# --- Excel Log Setup ---
excel_log_file = None
excel_workbook = None
excel_sheet = None

def init_excel_log():
    global excel_log_file, excel_workbook, excel_sheet
    start_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    excel_log_file = os.path.join(DESKTOP_PATH, f"stock_sync_log_{start_time}.xlsx")
    excel_workbook = openpyxl.Workbook()
    excel_sheet = excel_workbook.active
    excel_sheet.title = "Stock Sync Log"
    
    headers = ["ساعت", "کد کالا", "نام کالا", "موجودی سایت (قدیم)", "موجودی دیتابیس (جدید)", "وضعیت", "نوع محصول", "پیام"]
    excel_sheet.append(headers)
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    for col in range(1, len(headers) + 1):
        cell = excel_sheet.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
    
    excel_sheet.column_dimensions['A'].width = 12
    excel_sheet.column_dimensions['B'].width = 15
    excel_sheet.column_dimensions['C'].width = 40
    excel_sheet.column_dimensions['D'].width = 18
    excel_sheet.column_dimensions['E'].width = 18
    excel_sheet.column_dimensions['F'].width = 20
    excel_sheet.column_dimensions['G'].width = 15
    excel_sheet.column_dimensions['H'].width = 30
    excel_workbook.save(excel_log_file)
    logger.info(f"Excel log file created at: {excel_log_file}")

def add_to_excel_log(code, name, old_stock, new_stock, status, product_type="single", message=""):
    try:
        timestamp = datetime.now().strftime("%H:%M:%S")
        excel_sheet.append([timestamp, code, name, old_stock, new_stock, status, product_type, message])
        last_row = excel_sheet.max_row
        if "موفق" in status or "هماهنگ" in status:
            fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            font = Font(color="006100")
        else:
            fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            font = Font(color="9C0006")
        
        for col in range(1, 9):
            cell = excel_sheet.cell(row=last_row, column=col)
            cell.fill = fill
            cell.font = font
        excel_workbook.save(excel_log_file)
    except Exception as e:
        logger.error(f"Failed to write to Excel: {e}")

class Colors:
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    RESET = '\033[0m'
    BOLD = '\033[1m'

class ColoredFormatter(logging.Formatter):
    def format(self, record):
        if record.levelno == logging.ERROR:
            record.msg = f"{Colors.RED}{Colors.BOLD}[FAILED]{Colors.RESET} {record.msg}"
        elif record.levelno == logging.WARNING:
            record.msg = f"{Colors.YELLOW}[WARNING]{Colors.RESET} {record.msg}"
        elif record.levelno == logging.INFO:
            if "Successfully" in record.msg or "synced" in record.msg:
                record.msg = f"{Colors.GREEN}[SUCCESS]{Colors.RESET} {record.msg}"
            elif "Fetching" in record.msg or "Connecting" in record.msg or "Scanning" in record.msg:
                record.msg = f"{Colors.CYAN}[ACTION]{Colors.RESET} {record.msg}"
            elif "added to repo" in record.msg:
                record.msg = f"{Colors.YELLOW}[REPO]{Colors.RESET} {record.msg}"
            elif "removed from repo" in record.msg:
                record.msg = f"{Colors.GREEN}[REPO]{Colors.RESET} {record.msg}"
            else:
                record.msg = f"{Colors.BLUE}[INFO]{Colors.RESET} {record.msg}"
        return super().format(record)

logger = logging.getLogger("StockSync")
logger.setLevel(logging.INFO)
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(ColoredFormatter('%(message)s'))
logger.addHandler(console_handler)

gui_root = None
main_tree = None
repo_tree = None
gui_status_label = None
main_stats_label = None
repo_stats_label = None
notebook = None
info_text = None
sync_running = True
repo_frame = None

kpi_repo_label = None
kpi_in_label = None
kpi_out_label = None
kpi_fail_label = None
clock_label = None

def update_kpi_counters():
    if not main_tree:
        return
    children = main_tree.get_children()
    total_in = 0
    total_out = 0
    fail = 0
    
    for item in children:
        vals = main_tree.item(item, "values")
        old_val = int(float(vals[2])) if str(vals[2]).replace('.','',1).isdigit() else 0
        new_val = int(float(vals[3])) if str(vals[3]).replace('.','',1).isdigit() else 0
        status = str(vals[5]).strip()
        
        if status == "تغییر موفق":
            diff = new_val - old_val
            if diff > 0:
                total_in += diff
            elif diff < 0:
                total_out += abs(diff)
        else:
            fail += 1

    repo_count = len(load_repo())
    if kpi_repo_label: kpi_repo_label.configure(text=f"{repo_count} کالا")
    if kpi_fail_label: kpi_fail_label.configure(text=f"{fail} مورد")
    if kpi_in_label: kpi_in_label.configure(text=str(total_in))
    if kpi_out_label: kpi_out_label.configure(text=str(total_out))

def update_clock():
    if gui_root and clock_label:
        current_time_str = datetime.now().strftime("%H:%M:%S")
        clock_label.configure(text=current_time_str)
        gui_root.after(1000, update_clock)

def add_to_table(code, name, old_stock, new_stock, status, product_type="single"):
    if gui_root is None or main_tree is None:
        return
    
    def _add():
        timestamp = datetime.now().strftime("%H:%M:%S")
        try:
            old_val = int(float(old_stock))
            new_val = int(float(new_stock))
            diff = new_val - old_val
        except:
            diff = 0
        
        if "موفق" in status or "هماهنگ" in status:
            if diff > 0:
                tag = "increase"
            elif diff < 0:
                tag = "decrease"
            else:
                tag = "success"
        else:
            tag = "failed"
        
        display_name = name if len(name) <= 40 else name[:37] + "..."
        main_tree.insert("", 0, values=(code, display_name, old_stock, new_stock, timestamp, status), tags=(tag,))
        
        success_count = 0
        fail_count = 0
        for item in main_tree.get_children():
            item_status = main_tree.item(item, "values")[5]
            if "موفق" in item_status or "هماهنگ" in item_status:
                success_count += 1
            else:
                fail_count += 1
        
        if main_stats_label:
            main_stats_label.configure(text=f"Total: {success_count + fail_count} | ✅ {success_count} | ❌ {fail_count}")
        
        update_kpi_counters()
        add_to_excel_log(code, name, old_stock, new_stock, status, product_type, "")
    
    gui_root.after(0, _add)

def update_repo_badge():
    if gui_root is None or notebook is None or repo_frame is None:
        return
    try:
        repo = load_repo()
        count = len(repo)
        current_tab = notebook.select()
        is_repo_active = (current_tab == str(repo_frame))
        
        if count > 0 and not is_repo_active:
            notebook.tab(repo_frame, text=f"    📦 مدیریت مخزن سفارشات ({count})    ")
        else:
            notebook.tab(repo_frame, text="    📦 مدیریت مخزن سفارشات    ")
    except Exception:
        pass

def update_repo_table():
    if gui_root is None or repo_tree is None:
        return
    
    def _update():
        for item in repo_tree.get_children():
            repo_tree.delete(item)
        
        repo = load_repo()
        repo = clean_expired_repo(repo)
        
        for code, data in repo.items():
            site_stock = data.get("site_stock", 0)
            repo_stock = data.get("repo_stock", 0)
            entry_time = data.get("time", "")
            name = data.get("name", "N/A")
            
            if site_stock == repo_stock:
                status = "در انتظار صدور فاکتور حسابداری"
            else:
                status = "در انتظار هماهنگی با سایت"
            
            try:
                time_obj = datetime.fromisoformat(entry_time)
                time_str = time_obj.strftime("%H:%M:%S - %Y/%m/%d")
            except:
                time_str = entry_time
            
            repo_tree.insert("", "end", values=(code, name, site_stock, repo_stock, time_str, status), tags=("repo_item",))
        
        if repo_stats_label:
            repo_stats_label.configure(text=f"تعداد در مخزن: {len(repo)}")
        
        update_kpi_counters()
        update_repo_badge()
    
    gui_root.after(0, _update)

def remove_from_repo(code):
    repo = load_repo()
    if code in repo:
        name = repo[code].get("name", "N/A")
        site_stock = repo[code].get("site_stock", 0)
        del repo[code]
        save_repo(repo)
        update_repo_table()
        add_to_table(code, name, site_stock, site_stock, "⚠️ حذف دستی از مخزن", "single")
        logger.info(f"Product {code} manually removed from repo")

def update_status(text):
    if gui_root and gui_status_label:
        gui_root.after(0, lambda: gui_status_label.configure(text=text))

def get_db_stock():
    logger.info("Connecting to SQL Server database...")
    logger.info("Fetching product codes, names and stock quantities from v_PRODUCTS table...")
    
    data = {}
    try:
        conn = pyodbc.connect(DB_CONN)
        cursor = conn.cursor()
        cursor.execute("SELECT CODE, NAME, STOCK FROM v_PRODUCTS WHERE CODE IS NOT NULL AND STOCK IS NOT NULL")
        rows = cursor.fetchall()
        for row in rows:
            code = str(row[0]).strip()
            name = str(row[1]).strip() if row[1] else "N/A"
            stock = int(float(row[2]))
            data[code] = (stock, name)
        cursor.close()
        conn.close()
        logger.info(f"Database scan completed. Retrieved {len(data)} products with stock information.")
    except Exception as e:
        logger.error(f"Database connection or query failed: {e}")
        raise
    return data

def get_product_combinations(product_id):
    try:
        detail_url = f"{BASE_URL}/{product_id}"
        response = requests.get(detail_url, headers=API_HEADERS, timeout=30)
        raw_text = response.text.strip()
        if raw_text.startswith(")]}',"):
            raw_text = raw_text[5:].strip()
        
        detail_data = json.loads(raw_text)
        product_detail = detail_data.get("data", detail_data)
        
        skus = []
        if "combinations" in product_detail and product_detail["combinations"]:
            for combo in product_detail["combinations"]:
                combo_sku = combo.get("sku")
                if combo_sku and combo_sku != "" and combo_sku != "None":
                    combo_qty = combo.get("quantity")
                    if combo_qty is not None and combo_qty != "":
                        combo_title = combo.get("title", product_detail.get("title", "N/A"))
                        skus.append({
                            "sku": str(combo_sku).strip(),
                            "quantity": int(float(combo_qty)),
                            "title": combo_title
                        })
        return skus
    except Exception as e:
        return []

def get_webzi_stock():
    logger.info("Fetching ALL stock from Webzi via API...")
    start_time = time.time()
    
    logger.info("Step 1: Fetching product list...")
    all_products = []
    page = 1
    
    while True:
        params = {"page": str(page)}
        try:
            response = requests.get(BASE_URL, headers=API_HEADERS, params=params, timeout=30)
            raw_text = response.text.strip()
            if raw_text.startswith(")]}',"):
                raw_text = raw_text[5:].strip()
            
            data = json.loads(raw_text)
            products = data.get("items", [])
            
            if not products:
                break
            
            all_products.extend(products)
            page += 1
            
        except Exception as e:
            logger.error(f"API error on page {page}: {e}")
            break
    
    logger.info(f"Found {len(all_products)} total products")
    
    webzi_data = {}
    multi_color_ids = []
    
    for product in all_products:
        custom_per_variation = product.get("custom_per_variation", "0")
        
        if custom_per_variation == "1":
            product_id = product.get("id")
            if product_id:
                multi_color_ids.append(product_id)
        else:
            sku = product.get("sku")
            if sku and sku != "" and sku != "None":
                sku = str(sku).strip()
                raw_quantity = product.get("quantity")
                if raw_quantity is not None and raw_quantity != "":
                    quantity = int(float(raw_quantity))
                    title = product.get("title", "N/A")
                    webzi_data[sku] = (quantity, title, "single")
    
    logger.info(f"Single-color SKUs: {len(webzi_data)}")
    logger.info(f"Multi-color products to check: {len(multi_color_ids)}")
    
    if multi_color_ids:
        logger.info("Step 2: Fetching combinations for multi-color products (multi-threading)...")
        multi_color_skus = []
        max_workers = 10
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(get_product_combinations, pid): pid for pid in multi_color_ids}
            
            for future in as_completed(futures):
                result = future.result()
                for item in result:
                    if item["sku"] not in webzi_data:
                        multi_color_skus.append(item)
                        webzi_data[item["sku"]] = (item["quantity"], item["title"], "multi")
        
        logger.info(f"Multi-color SKUs found: {len(multi_color_skus)}")
    
    total_time = time.time() - start_time
    logger.info(f"Webzi scan completed in {total_time:.2f} seconds. Total SKUs: {len(webzi_data)}")
    
    return webzi_data

def find_and_process_mismatches(db_data, webzi_data):
    repo = load_repo()
    repo = clean_expired_repo(repo)
    last_webzi = load_last_stock()
    last_db = load_last_db_stock()
    
    common_codes = set(db_data.keys()) & set(webzi_data.keys())
    
    for code in list(repo.keys()):
        if code in webzi_data:
            webzi_stock, _, _ = webzi_data[code]
            repo_stock = repo[code].get("repo_stock", 0)
            
            if webzi_stock != repo_stock:
                repo[code]["repo_stock"] = webzi_stock
                repo[code]["site_stock"] = webzi_stock
                repo[code]["time"] = datetime.now().isoformat()
                save_repo(repo)
                logger.info(f"Repo synced with site for {code}: {repo_stock} → {webzi_stock}")
    
    for code in common_codes:
        webzi_stock, _, _ = webzi_data[code]
        db_stock, db_name = db_data[code]
        last_webzi_val = last_webzi.get(code, webzi_stock)
        last_db_val = last_db.get(code, db_stock)
        
        if webzi_stock < last_webzi_val and db_stock == last_db_val and code not in repo:
            repo[code] = {
                "site_stock": webzi_stock,
                "repo_stock": webzi_stock,
                "time": datetime.now().isoformat(),
                "name": db_name
            }
            save_repo(repo)
            logger.info(f"🆕 New order detected! {code} added to repo (stock decreased: {last_webzi_val} → {webzi_stock}, DB unchanged: {db_stock})")
            update_repo_table()
    
    mismatches = []
    
    for code in common_codes:
        db_stock, db_name = db_data[code]
        webzi_stock, webzi_name, product_type = webzi_data[code]
        
        if db_stock != webzi_stock:
            if code in repo:
                logger.info(f"{code}: mismatch ignored (in repo waiting for invoice) | DB:{db_stock} Webzi:{webzi_stock}")
                continue
            
            mismatches.append({
                "code": code,
                "name": db_name,
                "db_stock": db_stock,
                "webzi_stock": webzi_stock,
                "product_type": product_type
            })
    
    logger.info(f"Found {len(mismatches)} mismatches to process")
    
    success_count = 0
    fail_count = 0
    failed_codes = []
    
    if mismatches:
        logger.info(f"\n{Colors.CYAN}{'='*50}{Colors.RESET}")
        logger.info(f"{Colors.CYAN}STARTING BATCH API SYNC: {len(mismatches)} products{Colors.RESET}")
        logger.info(f"{Colors.CYAN}{'='*50}{Colors.RESET}\n")
        
        for idx, item in enumerate(mismatches, 1):
            code = item["code"]
            name = item["name"]
            db_stock = item["db_stock"]
            webzi_stock = item["webzi_stock"]
            product_type = item.get("product_type", "single")
            
            type_label = "چندرنگ" if product_type == "multi" else "تک‌رنگ"
            logger.info(f"[{idx}/{len(mismatches)}] {code} ({name[:30]}) [{type_label}] | Webzi: {webzi_stock} → DB: {db_stock}")
            update_status(f"API updating [{idx}/{len(mismatches)}] - {code}")
            
            success = update_product_via_api(code, db_stock, product_type)
            
            if success:
                success_count += 1
                logger.info(f"[{idx}/{len(mismatches)}] {Colors.GREEN}Product {code} synced: {webzi_stock} → {db_stock}{Colors.RESET}")
                add_to_table(code, name, webzi_stock, db_stock, "تغییر موفق", product_type)
            else:
                fail_count += 1
                failed_codes.append(code)
                logger.error(f"[{idx}/{len(mismatches)}] Product {code} FAILED")
                add_to_table(code, name, webzi_stock, db_stock, "❌ تغییر ناموفق", product_type)
    
    repo = load_repo()
    changed = False
    for code in list(repo.keys()):
        if code in db_data and code in webzi_data:
            db_stock, db_name = db_data[code]
            webzi_stock, _, product_type = webzi_data[code]
            repo_stock = repo[code].get("repo_stock", 0)
            
            if db_stock == webzi_stock == repo_stock:
                name = repo[code].get("name", db_name)
                del repo[code]
                changed = True
                add_to_table(code, name, webzi_stock, db_stock, "✅ هماهنگ شد (خروج از مخزن)", product_type)
                logger.info(f"{code}: removed from repo (fully synced - invoice issued)")
    
    if changed:
        save_repo(repo)
        update_repo_table()
    
    current_webzi_stock = {}
    current_db_stock = {}
    for code in common_codes:
        webzi_stock, _, _ = webzi_data[code]
        db_stock, _ = db_data[code]
        current_webzi_stock[code] = webzi_stock
        current_db_stock[code] = db_stock
    save_last_stock(current_webzi_stock)
    save_last_db_stock(current_db_stock)
    
    logger.info(f"\n{Colors.CYAN}{'='*50}{Colors.RESET}")
    logger.info(f"{Colors.CYAN}SYNC CYCLE COMPLETED{Colors.RESET}")
    logger.info(f"{Colors.CYAN}{'='*50}{Colors.RESET}")
    logger.info(f"Total mismatches: {len(mismatches)}")
    logger.info(f"{Colors.GREEN}Fixed: {success_count}{Colors.RESET}")
    
    if fail_count > 0:
        logger.error(f"Failed: {fail_count} | Codes: {', '.join(failed_codes)}")
    
    update_status(f"Last sync: {datetime.now().strftime('%H:%M:%S')} | ✅ {success_count} | ❌ {fail_count}")
    
    return success_count, fail_count

def update_product_via_api(code, new_stock, product_type="single"):
    try:
        search_params = {
            "filter[global]": code,
            "page": "1",
            "savedSearch": "0",
            "sorting[created_at]": "desc"
        }
        
        response = requests.get(BASE_URL, headers=API_HEADERS, params=search_params)
        raw_text = response.text.strip()
        if raw_text.startswith(")]}',"):
            raw_text = raw_text[5:].strip()
        
        search_data = json.loads(raw_text)
        products_list = search_data.get("items", [])
        
        if not products_list:
            logger.error(f"Product {code} not found via API search")
            return False
        
        product_obj = None
        for p in products_list:
            p_sku = str(p.get("sku", "")).strip()
            if p_sku == code:
                product_obj = p
                break
        
        if not product_obj:
            product_obj = products_list[0]
        
        product_id = product_obj.get("id")
        detail_url = f"{BASE_URL}/{product_id}"
        
        detail_response = requests.get(detail_url, headers=API_HEADERS)
        detail_raw = detail_response.text.strip()
        if detail_raw.startswith(")]}',"):
            detail_raw = detail_raw[5:].strip()
        
        full_product = json.loads(detail_raw).get("data", {})
        
        if product_type == "multi":
            if "combinations" not in full_product or not full_product["combinations"]:
                logger.error(f"Product {product_id} has no combinations")
                return False
            
            combination_found = False
            for combo in full_product["combinations"]:
                if str(combo.get("sku", "")).strip() == code:
                    combo["quantity"] = str(new_stock)
                    combination_found = True
                    break
            
            if not combination_found:
                logger.error(f"Combination {code} not found")
                return False
            
            total_qty = sum(int(c.get("quantity", 0)) for c in full_product["combinations"])
            full_product["quantity"] = str(total_qty)
            full_product["inventory_status"] = "3"
            full_product["custom_per_variation"] = "1"
            
            update_url = f"{BASE_URL}/{product_id}"
            payload = {"data": full_product}
            
            res = requests.post(update_url, json=payload, headers=API_HEADERS)
            
            if res.status_code in [200, 201]:
                verify_res = requests.get(detail_url, headers=API_HEADERS)
                verify_raw = verify_res.text.strip()
                if verify_raw.startswith(")]}',"):
                    verify_raw = verify_raw[5:].strip()
                
                verify_data = json.loads(verify_raw)
                verify_product = verify_data.get("data", verify_data)
                
                for combo in verify_product.get("combinations", []):
                    if str(combo.get("sku", "")).strip() == code:
                        if str(combo.get("quantity", "")) == str(new_stock):
                            return True
                return False
            return False
        
        else:
            payload_data = product_obj.copy()
            
            pricing_fields = [
                "price", 
                "compare_at_price", 
                "sell_price",
                "discount_price", 
                "discount_type", 
                "discount_value", 
                "sale",
                "scheduled_sale", 
                "scheduled_discount_type", 
                "scheduled_discount_value",
                "scheduled_discount_price", 
                "scheduled_discount_start", 
                "scheduled_discount_end",
                "scheduled_discount_has_end"
            ]
            
            for field in pricing_fields:
                if field in full_product:
                    payload_data[field] = full_product[field]
            
            payload_data["quantity"] = str(new_stock)
            payload_data["inventory_status"] = "3"
            
            update_url = f"{BASE_URL}/{product_id}"
            payload = {"data": payload_data}
            
            res = requests.post(update_url, json=payload, headers=API_HEADERS)
            
            if res.status_code in [200, 201]:
                verify_res = requests.get(update_url, headers=API_HEADERS)
                verify_raw = verify_res.text.strip().lstrip(")]}',\n")
                verify_data = json.loads(verify_raw)
                verify_product = verify_data.get("data", verify_data)
                
                if str(verify_product.get("quantity", "")) == str(new_stock):
                    return True
            
            return False
            
    except Exception as e:
        logger.error(f"API error for {code}: {str(e)}")
        return False

def test_api_connection(base_url, cookie, csrf_token=None):
    try:
        test_headers = FIXED_HEADERS.copy()
        test_headers["Cookie"] = safe_latin1(cookie)
        if csrf_token:
            test_headers["X-CSRF-TOKEN"] = safe_latin1(csrf_token)
        
        response = requests.get(base_url, headers=test_headers, params={"page": "1"}, timeout=30)
        
        if response.status_code != 200:
            return False, f"❌ خطا در اتصال - وضعیت: {response.status_code}"
        
        raw_text = response.text.strip()
        
        if raw_text.startswith(")]}',"):
            raw_text = raw_text[5:].strip()
        
        try:
            data = json.loads(raw_text)
        except json.JSONDecodeError:
            return False, "❌ پاسخ دریافتی JSON معتبر نیست"
        
        items = data.get("items")
        
        if items is None:
            return False, "❌ ساختار پاسخ نامعتبر - کلید 'items' وجود ندارد"
        
        if not isinstance(items, list):
            return False, "❌ ساختار پاسخ نامعتبر - 'items' آرایه نیست"
        
        if len(items) == 0:
            return False, "⚠️ اتصال برقرار است اما هیچ محصولی یافت نشد!\n(ممکن است آدرس یا کوکی اشتباه باشد)"
        
        return True, f"✅ اتصال موفق - {len(items)} محصول در صفحه اول دریافت شد"
        
    except requests.exceptions.ConnectionError:
        return False, "❌ خطای اتصال - نمی‌توان به سرور متصل شد"
    except requests.exceptions.Timeout:
        return False, "❌ خطای Timeout - سرور پاسخ نمی‌دهد"
    except Exception as e:
        return False, f"❌ خطا: {str(e)}"

# ============================================================
# تنظیمات اتصال به API
# ============================================================

settings_window = None
config_text = None

def show_settings_dialog():
    global settings_window, config_text, BASE_URL, DEFAULT_COOKIE, API_HEADERS, FIXED_HEADERS
    
    settings_window = ctk.CTkToplevel(gui_root)
    settings_window.title("تنظیمات API")
    settings_window.geometry("900x580")
    settings_window.configure(fg_color="#ffffff")
    settings_window.transient(gui_root)
    settings_window.grab_set()

    ctk.CTkLabel(settings_window, text="⚙️ تنظیمات اتصال به API",
                 font=("Tahoma", 13, "bold"), text_color="#059669").pack(pady=(14, 5))
    ctk.CTkLabel(settings_window, text="راهنما: در الگوی زیر، آدرس API، توکن دسترسی، توکن CSRF و کوکی را با اطلاعات معتبر سایت خود جایگزین کنید. این اطلاعات محرمانه‌اند؛ فایل را منتشر نکنید. برای دریافت اطلاعات می‌توانید از گزینه «استخراج از مرورگر» استفاده کنید.",
                 font=("Tahoma", 9), text_color="#64748b").pack(pady=(0, 12))

    form = ctk.CTkFrame(settings_window, fg_color="#f8fafc", corner_radius=14, border_width=1, border_color="#dbeafe")
    form.pack(fill="both", expand=True, padx=20, pady=5)

    ctk.CTkLabel(form, text="متن تنظیمات API (شامل BASE_URL و API_HEADERS)", font=("Tahoma", 10, "bold"), text_color="#1e293b").pack(anchor="e", padx=18, pady=(16, 4))
    
    config_text = scrolledtext.ScrolledText(form, font=("Consolas", 10), bg="#ffffff", fg="#1e293b", height=13, wrap=tk.NONE, bd=1, relief="solid")
    config_text.pack(fill="both", expand=True, padx=18, pady=(0, 10))

    def text_menu(widget):
        menu = tk.Menu(widget, tearoff=0, font=("Tahoma", 9))
        def paste(event=None):
            try: widget.insert(tk.INSERT, widget.clipboard_get())
            except Exception: pass
            return "break"
        def copy(event=None):
            try:
                val = widget.get(tk.SEL_FIRST, tk.SEL_LAST)
                widget.clipboard_clear(); widget.clipboard_append(val)
            except Exception: pass
            return "break"
        def cut(event=None):
            copy()
            try: widget.delete(tk.SEL_FIRST, tk.SEL_LAST)
            except Exception: pass
            return "break"
        def select_all(event=None): widget.tag_add("sel", "1.0", "end"); return "break"
        menu.add_command(label="✂️ برش (Cut)", command=cut)
        menu.add_command(label="کپی (Copy)", command=copy)
        menu.add_command(label="چسباندن (Paste)", command=paste)
        menu.add_separator()
        menu.add_command(label="انتخاب همه (Select All)", command=select_all)
        widget.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))
        for seq, fn in (("<Control-c>", copy), ("<Control-C>", copy), ("<Control-v>", paste), ("<Control-V>", paste), ("<Control-x>", cut), ("<Control-X>", cut), ("<Control-a>", select_all), ("<Control-A>", select_all)):
            widget.bind(seq, fn)

    text_menu(config_text)

    initial_config_sample = f'BASE_URL = "{BASE_URL}"\n\nAPI_HEADERS = {{\n    "Authorization": "Bearer {FIXED_HEADERS.get("Authorization", "YOUR_API_TOKEN_HERE")}",\n    "Content-Type": "application/json",\n    "Accept": "application/json, text/plain, */*",\n    "X-Requested-With": "XMLHttpRequest",\n    "X-CSRF-TOKEN": "{FIXED_HEADERS.get("X-CSRF-TOKEN", "")}",\n    "Cookie": "{DEFAULT_COOKIE}"\n}}'
    config_text.insert("1.0", initial_config_sample)
    ctk.CTkLabel(form,
        text="محل ورود اطلاعات: BASE_URL = آدرس API | Authorization = توکن دسترسی | X-CSRF-TOKEN = توکن امنیتی | Cookie = کوکی نشست",
        font=("Tahoma", 9), text_color="#b45309", justify="right", wraplength=820
    ).pack(anchor="e", padx=18, pady=(0, 10))

    def current_values():
        content = config_text.get("1.0", tk.END)
        b_url, cookie, csrf, err = parse_api_config(content)
        if err:
            return None, None, None, None, err
        product_url = derive_product_url(b_url)
        return b_url, product_url, cookie, csrf, None

    def set_settings_fields(b_url, cookie, csrf_val):
        formatted_text = f'BASE_URL = "{b_url}"\n\nAPI_HEADERS = {{\n    "Authorization": "Bearer {FIXED_HEADERS.get("Authorization", "YOUR_API_TOKEN_HERE")}",\n    "Content-Type": "application/json",\n    "Accept": "application/json, text/plain, */*",\n    "X-Requested-With": "XMLHttpRequest",\n    "X-CSRF-TOKEN": "{csrf_val}",\n    "Cookie": "{cookie}"\n}}'
        config_text.delete("1.0", tk.END)
        config_text.insert("1.0", formatted_text)

    def execute_attached_or_firefox_extraction(browser_choice):
        target_product_url = derive_product_url(BASE_URL) or "https://YOUR-DOMAIN.example/shop/products/"
        target_base = derive_api_url(target_product_url) or BASE_URL

        if browser_choice == 'chrome':
            from websocket import create_connection
            ws_url = None; actual_url = target_product_url
            try:
                resp = requests.get("http://127.0.0.1:9222/json", timeout=2); targets = resp.json()
                for t in targets:
                    if "mywebzi.ir" in t.get("url", "") or "admin" in t.get("url", ""):
                        ws_url=t.get("webSocketDebuggerUrl"); actual_url=t.get("url"); break
                if not ws_url:
                    for t in targets:
                        if t.get("type") == "page": ws_url=t.get("webSocketDebuggerUrl"); actual_url=t.get("url"); break
            except: pass
            if not ws_url:
                chrome_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                if not os.path.exists(chrome_path):
                    alt=r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                    if os.path.exists(alt): chrome_path=alt
                try:
                    tmp_user_dir=os.path.join(APP_PATH,"chrome_debug_profile"); os.makedirs(tmp_user_dir,exist_ok=True)
                    subprocess.Popen([chrome_path,"--remote-debugging-port=9222","--remote-allow-origins=*",f"--user-data-dir={tmp_user_dir}",target_product_url])
                    for _ in range(10):
                        time.sleep(1.5)
                        try:
                            targets=requests.get("http://127.0.0.1:9222/json",timeout=2).json()
                            for t in targets:
                                if t.get("type")=="page": ws_url=t.get("webSocketDebuggerUrl"); actual_url=t.get("url",target_product_url); break
                            if ws_url: break
                        except: pass
                except Exception as e:
                    messagebox.showerror("خطا", f"نتوانست کروم را اجرا کند: {e}", parent=settings_window); return
            if not ws_url:
                messagebox.showerror("خطا", "پورت 9222 کروم پاسخ نداد. لطفاً کروم را کاملاً بسته و مجدد تلاش کنید.", parent=settings_window); return
            ctrl_win=ctk.CTkToplevel(settings_window); ctrl_win.title("استخراج از کروم"); ctrl_win.geometry("440x220"); ctrl_win.transient(settings_window); ctrl_win.grab_set()
            ctk.CTkLabel(ctrl_win,text="🌐 کروم با قابلیت دیباگ باز شد.\n\nلطفاً در کروم لاگین کرده و وارد صفحه محصولات شوید.\nسپس دکمه زیر را فشار دهید:",font=("Tahoma",10,"bold"),text_color="#1e293b",justify="center").pack(pady=15)
            def do_read_from_chrome():
                nonlocal ws_url, actual_url
                try:
                    targets=requests.get("http://127.0.0.1:9222/json",timeout=3).json(); ws_url=None
                    for t in targets:
                        if "mywebzi.ir" in t.get("url","") or "admin" in t.get("url",""): ws_url=t.get("webSocketDebuggerUrl"); actual_url=t.get("url"); break
                    if not ws_url:
                        messagebox.showerror("خطا","پیدا نکردن تب فعال کروم. مطمئن شوید تب وبزی باز است.",parent=ctrl_win); return
                    ws=create_connection(ws_url); ws.send(json.dumps({"id":1,"method":"Runtime.evaluate","params":{"expression":"document.cookie","returnByValue":True}})); r=json.loads(ws.recv()); ws.close()
                    cookie_str=safe_latin1(r.get("result",{}).get("result",{}).get("value",""))
                    ws=create_connection(ws_url); ws.send(json.dumps({"id":2,"method":"Runtime.evaluate","params":{"expression":"document.querySelector('meta[name=\"csrf-token\"]')?.content || ''","returnByValue":True}})); r=json.loads(ws.recv()); ws.close()
                    csrf_val=safe_latin1(r.get("result",{}).get("result",{}).get("value","")) or extract_csrf_from_cookie(cookie_str)
                    prod_url=actual_url if "/shop/products" in actual_url else target_product_url
                    api_u=derive_api_url(prod_url) or target_base
                    set_settings_fields(api_u, cookie_str, csrf_val)
                    ctrl_win.destroy()
                    ok,msg=test_api_connection(api_u,cookie_str,csrf_val)
                    if ok: messagebox.showinfo("موفقیت کامل","✨ اطلاعات با موفقیت دریافت و ارتباط با سرور برقرار شد!",parent=settings_window)
                    else: messagebox.showwarning("هشدار اتصال",f"⚠️ کوکی خوانده شد، اما تست اتصال شکست خورد:\n{msg}",parent=settings_window)
                except Exception as e: messagebox.showerror("خطا",safe_latin1(str(e)),parent=ctrl_win)
            ctk.CTkButton(ctrl_win,text="✅ لاگین تمام شد، خواندن و تست اتصال",fg_color="#059669",hover_color="#047857",font=("Tahoma",10,"bold"),height=38,corner_radius=10,command=do_read_from_chrome).pack(fill="x",padx=30,pady=5)
        else:
            cookie_str=""; csrf_val=""; actual_url=target_product_url
            try:
                ext_fx=browser_cookie3.firefox(domain_name="mywebzi.ir")
                if not ext_fx: ext_fx=browser_cookie3.firefox()
                if ext_fx:
                    c_map={c.name:c.value for c in ext_fx if "mywebzi" in c.domain or "Webzi" in c.name or "CSRF" in c.name.upper()}
                    if c_map: cookie_str="; ".join([f"{k}={v}" for k,v in c_map.items()]); csrf_val=extract_csrf_from_cookie(cookie_str)
            except: pass
            if cookie_str:
                api_u=derive_api_url(target_product_url) or target_base
                set_settings_fields(api_u, cookie_str, csrf_val)
                ok,msg=test_api_connection(api_u,cookie_str,csrf_val)
                if ok: messagebox.showinfo("موفقیت کامل",f"✨ کوکی‌ها از پروفایل فایرفاکس خوانده شد و ارتباط برقرار شد!\n({msg})",parent=settings_window)
                else: messagebox.showwarning("هشدار",f"⚠️ کوکی خوانده شد ولی تست اتصال خطا داد:\n{msg}",parent=settings_window)
                return
            try:
                from playwright.sync_api import sync_playwright
                ctrl_win=ctk.CTkToplevel(settings_window); ctrl_win.title("مدیریت لاگین - فایرفاکس"); ctrl_win.geometry("440x220"); ctrl_win.transient(settings_window); ctrl_win.grab_set()
                ctk.CTkLabel(ctrl_win,text="🌐 فایرفاکس باز شد.\n\nلطفاً در فایرفاکس لاگین کرده و وارد صفحه محصولات شوید.\nسپس دکمه زیر را فشار دهید:",font=("Tahoma",10,"bold"),text_color="#1e293b",justify="center").pack(pady=15)
                pw_ref={"p_instance":None,"browser":None,"page":None}
                def run_fx_pw():
                    try:
                        p_fx=sync_playwright().start(); pw_ref["p_instance"]=p_fx; browser=p_fx.firefox.launch(headless=False); pw_ref["browser"]=browser; context=browser.new_context(); page=context.new_page(); pw_ref["page"]=page; page.goto(target_product_url)
                    except Exception as e: messagebox.showerror("خطا در فایرفاکس",safe_latin1(str(e)),parent=ctrl_win)
                threading.Thread(target=run_fx_pw,daemon=True).start()
                def do_read_from_fx():
                    page=pw_ref.get("page"); browser=pw_ref.get("browser"); p_inst=pw_ref.get("p_instance")
                    if not page: messagebox.showerror("خطا","صفحه فایرفاکس آماده نیست یا بسته شده است.",parent=ctrl_win); return
                    try:
                        act_url=page.url; c_str=safe_latin1(page.evaluate("document.cookie")); c_val=safe_latin1(page.evaluate("document.querySelector('meta[name=\"csrf-token\"]')?.content || ''")) or extract_csrf_from_cookie(c_str)
                        api_u=derive_api_url(act_url) or target_base
                        set_settings_fields(api_u, c_str, c_val)
                        ctrl_win.destroy()
                        if browser: browser.close()
                        if p_inst: p_inst.stop()
                        ok,msg=test_api_connection(api_u,c_str,c_val)
                        if ok: messagebox.showinfo("موفقیت کامل",f"✨ اطلاعات از فایرفاکس دریافت و ارتباط با سرور برقرار شد!\n({msg})",parent=settings_window)
                        else: messagebox.showwarning("هشدار",f"⚠️ کوکی خوانده شد، اما تست اتصال خطا داد:\n{msg}",parent=settings_window)
                    except Exception as e: messagebox.showerror("خطا",safe_latin1(str(e)),parent=ctrl_win)
                ctk.CTkButton(ctrl_win,text="✅ لاگین تمام شد، خواندن و تست اتصال",fg_color="#059669",hover_color="#047857",font=("Tahoma",10,"bold"),height=38,corner_radius=10,command=do_read_from_fx).pack(fill="x",padx=30,pady=5)
            except Exception as e:
                messagebox.showerror("خطا در فایرفاکس",safe_latin1(str(e)),parent=settings_window)

    def pick_browser_and_extract():
        sel_win=ctk.CTkToplevel(settings_window); sel_win.title("انتخاب مرورگر هدف"); sel_win.geometry("340x200"); sel_win.configure(fg_color="#ffffff"); sel_win.transient(settings_window); sel_win.grab_set()
        ctk.CTkLabel(sel_win,text="مرورگر مورد نظر را انتخاب کنید:",font=("Tahoma",10,"bold"),text_color="#1e293b").pack(pady=12)
        ctk.CTkButton(sel_win,text="🌐 گوگل کروم (اجرای هوشمند با دیباگ)",fg_color="#2563eb",hover_color="#1d4ed8",font=("Tahoma",9,"bold"),height=36,corner_radius=10,command=lambda:[sel_win.destroy(),execute_attached_or_firefox_extraction("chrome")]).pack(fill="x",padx=20,pady=4)
        ctk.CTkButton(sel_win,text="🦊 موزیلا فایرفاکس (Firefox)",fg_color="#ea580c",hover_color="#c2410c",font=("Tahoma",10,"bold"),height=36,corner_radius=10,command=lambda:[sel_win.destroy(),execute_attached_or_firefox_extraction("firefox")]).pack(fill="x",padx=20,pady=4)

    def test_conn():
        b_url, _, cook, csrf, err = current_values()
        if err:
            messagebox.showerror("خطا", err, parent=settings_window)
            return
        success, msg = test_api_connection(b_url, cook, csrf)
        if success:
            messagebox.showinfo("تست اتصال", msg, parent=settings_window)
        else:
            messagebox.showerror("خطا", msg, parent=settings_window)

    def save_conn():
        global BASE_URL, DEFAULT_COOKIE, API_HEADERS, FIXED_HEADERS
        b_url, product_url, cook, csrf, err = current_values()
        if err:
            messagebox.showerror("خطا", err, parent=settings_window)
            return
        BASE_URL = b_url
        DEFAULT_COOKIE = safe_latin1(cook)
        if csrf:
            FIXED_HEADERS["X-CSRF-TOKEN"] = safe_latin1(csrf)
        update_api_headers(DEFAULT_COOKIE, FIXED_HEADERS.get("X-CSRF-TOKEN"))
        save_config(b_url, DEFAULT_COOKIE, FIXED_HEADERS.get("X-CSRF-TOKEN"))
        messagebox.showinfo("موفقیت", "✅ تنظیمات با موفقیت ذخیره شد!", parent=settings_window)
        settings_window.destroy()
        update_info_tab()

    button_frame = ctk.CTkFrame(settings_window, fg_color="transparent")
    button_frame.pack(pady=15)

    ctk.CTkButton(button_frame, text="✨ استخراج از مرورگر", fg_color="#2563eb", hover_color="#1d4ed8", font=("Tahoma", 9, "bold"), corner_radius=8, command=pick_browser_and_extract).pack(side=tk.LEFT, padx=5)
    ctk.CTkButton(button_frame, text="🔍 تست اتصال", fg_color="#cbd5e1", text_color="#1e293b", hover_color="#94a3b8", font=("Tahoma", 9, "bold"), corner_radius=8, command=test_conn).pack(side=tk.LEFT, padx=5)
    ctk.CTkButton(button_frame, text="💾 ذخیره", fg_color="#059669", hover_color="#047857", font=("Tahoma", 9, "bold"), corner_radius=8, command=save_conn).pack(side=tk.LEFT, padx=5)
    ctk.CTkButton(button_frame, text="❌ انصراف", fg_color="#ef4444", hover_color="#dc2626", font=("Tahoma", 9, "bold"), corner_radius=8, command=settings_window.destroy).pack(side=tk.LEFT, padx=5)

def auto_update_cookie_action():
    browser_win = ctk.CTkToplevel(gui_root)
    browser_win.title("انتخاب مرورگر")
    browser_win.geometry("340x220")
    browser_win.configure(fg_color="#ffffff")
    browser_win.transient(gui_root)
    browser_win.grab_set()
    
    ctk.CTkLabel(browser_win, text="🌐 لطفاً مرورگر مورد نظر را انتخاب کنید:", 
                 font=("Tahoma", 11, "bold"), text_color="#1e293b").pack(pady=15)

    def start_fetching(b_name):
        browser_win.destroy()
        
        loading_win = ctk.CTkToplevel(gui_root)
        loading_win.title("لطفاً صبر کنید...")
        loading_win.geometry("350x120")
        loading_win.configure(fg_color="#ffffff")
        loading_win.transient(gui_root)
        loading_win.grab_set()
        
        ctk.CTkLabel(loading_win, text=f"🔄 در حال دریافت هوشمند اطلاعات\nاز {b_name.capitalize()}...", 
                     font=("Tahoma", 10, "bold"), text_color="#059669").pack(expand=True)

        def background_direct_cookie_fetcher(browser_name):
            try:
                domain_name = "mywebzi.ir"
                browser_func = getattr(browser_cookie3, browser_name, browser_cookie3.chrome)
                try:
                    extracted = browser_func(domain_name=domain_name)
                except:
                    extracted = browser_func()

                if not extracted:
                    gui_root.after(0, lambda: [loading_win.destroy(), messagebox.showerror("خطا", f"هیچ کوکی معتبری در {browser_name} پیدا نشد!")])
                    return

                cookies_map = {}
                csrf_val = ""
                found_url = None
                
                for cookie in extracted:
                    if "admin" in str(cookie.path) and found_url is None:
                        parts = cookie.path.split("/shop")
                        if len(parts) > 0:
                            found_url = f"https://{cookie.domain.lstrip('.')}{parts[0]}/api/shop/product"
                    if "Webzi-" in cookie.name or "CSRF" in cookie.name.upper() or "mywebzi" in str(cookie.domain):
                        cookies_map[cookie.name] = cookie.value
                    if "CSRF-TOKEN" in cookie.name.upper() or "csrf" in cookie.name.lower():
                        csrf_val = cookie.value

                if not cookies_map:
                    gui_root.after(0, lambda: [loading_win.destroy(), messagebox.showerror("خطا", "کوکی‌های وبزی یافت نشدند.")])
                    return

                cookie_str = "; ".join([f"{k}={v}" for k, v in cookies_map.items()])
                if not csrf_val:
                    match = re.search(r'CSRF-TOKEN=([^;]+)', cookie_str)
                    if match: csrf_val = match.group(1)

                global DEFAULT_COOKIE, API_HEADERS, FIXED_HEADERS, BASE_URL
                DEFAULT_COOKIE = safe_latin1(cookie_str)
                if csrf_val:
                    FIXED_HEADERS["X-CSRF-TOKEN"] = safe_latin1(csrf_val)
                update_api_headers(DEFAULT_COOKIE, FIXED_HEADERS.get("X-CSRF-TOKEN"))
                if found_url:
                    BASE_URL = found_url

                save_config(BASE_URL, DEFAULT_COOKIE, FIXED_HEADERS.get("X-CSRF-TOKEN"))
                gui_root.after(0, lambda: [loading_win.destroy(), messagebox.showinfo("موفق", f"✅ کوکی و تنظیمات {b_name} با موفقیت بروز شد!"), update_info_tab()])
            except Exception as e:
                gui_root.after(0, lambda: [loading_win.destroy(), messagebox.showerror("خطا", f"خطا: {safe_latin1(str(e))}")])

        threading.Thread(target=background_direct_cookie_fetcher, args=(b_name,), daemon=True).start()

    ctk.CTkButton(browser_win, text="🦊 فایرفاکس (Firefox)", fg_color="#ea580c", hover_color="#c2410c", font=("Tahoma", 10, "bold"), height=35, command=lambda: start_fetching("firefox")).pack(fill="x", padx=30, pady=5)
    ctk.CTkButton(browser_win, text="🌐 گوگل کروم (Chrome)", fg_color="#2563eb", hover_color="#1d4ed8", font=("Tahoma", 10, "bold"), height=35, command=lambda: start_fetching("chrome")).pack(fill="x", padx=30, pady=5)
    ctk.CTkButton(browser_win, text="💻 مایکروسافت اج (Edge)", fg_color="#0891b2", hover_color="#0e7490", font=("Tahoma", 10, "bold"), height=35, command=lambda: start_fetching("edge")).pack(fill="x", padx=30, pady=5)

def process_stock_sync():
    update_status("Fetching database stock...")
    db_data = get_db_stock()
    
    update_status("Fetching Webzi stock via API (multi-threading)...")
    webzi_data = get_webzi_stock()
    
    if not webzi_data:
        logger.warning("No Webzi data fetched. Skipping this cycle.")
        update_status("API fetch failed - skipping cycle")
        return
    
    update_status("Comparing stocks with repo...")
    find_and_process_mismatches(db_data, webzi_data)
    update_repo_table()

def update_info_tab():
    try:
        info_text.delete("1.0", tk.END)
        info_text.insert("1.0", f"📍 BASE_URL:\n{BASE_URL}\n\n🍪 Cookie:\n{DEFAULT_COOKIE}\n\n🔑 X-CSRF-TOKEN:\n{FIXED_HEADERS.get('X-CSRF-TOKEN', 'N/A')}\n\n📁 مسیر فایل تنظیمات:\n{CONFIG_FILE}\n\n✅ وضعیت: {'بارگذاری شده' if os.path.exists(CONFIG_FILE) else 'پیش‌فرض'}\n\n📁 مسیر برنامه:\n{APP_PATH}")
    except:
        pass

# ============================================================
# بخش استارت‌آپ ویندوز
# ============================================================

def add_to_startup():
    try:
        startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        shortcut_path = os.path.join(startup_folder, "StockSync.lnk")
        
        if getattr(sys, 'frozen', False):
            target_path = sys.executable
            arguments = ""
        else:
            target_path = sys.executable
            arguments = f'"{os.path.abspath(__file__)}"'
        
        working_dir = APP_PATH
        
        cmd = (
            f"$WshShell = New-Object -comObject WScript.Shell; "
            f"$Shortcut = $WshShell.CreateShortcut('{shortcut_path}'); "
            f"$Shortcut.TargetPath = '{target_path}'; "
            f"$Shortcut.Arguments = '{arguments}'; "
            f"$Shortcut.WorkingDirectory = '{working_dir}'; "
            f"$Shortcut.Save()"
        )
        subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True, check=True)
        return True
    except Exception as e:
        logger.error(f"❌ خطا در استارت‌آپ: {e}")
        return False

def remove_from_startup():
    try:
        startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
        shortcut_path = os.path.join(startup_folder, "StockSync.lnk")
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            return True
        return True
    except Exception as e:
        logger.error(f"❌ خطا در حذف استارت‌آپ: {e}")
        return False

# ============================================================
# رابط کاربری گرافیکی اصلی
# ============================================================
def create_gui():
    global gui_root, main_tree, repo_tree, gui_status_label, main_stats_label, repo_stats_label, notebook, info_text, repo_frame
    global kpi_repo_label, kpi_in_label, kpi_out_label, kpi_fail_label, clock_label
    
    gui_root = ctk.CTk()
    gui_root.title("Stock Sync Monitor — Enterprise Edition")
    gui_root.geometry("1300x760")
    gui_root.configure(fg_color="#f4f4f6")
    
    BG_CARD, BG_ELEVATED, GRID_HEADER_BG = "#ffffff", "#f8f9fa", "#e2e8f0"
    ACCENT_EMERALD, ACCENT_BLUE, ACCENT_PURPLE, ACCENT_ORANGE = "#059669", "#2563eb", "#7c3aed", "#d97706"
    TEXT_MAIN, TEXT_MUTED, BORDER_COLOR = "#1e293b", "#64748b", "#cbd5e1"

    header_frame = ctk.CTkFrame(gui_root, fg_color=BG_CARD, corner_radius=16, border_width=1, border_color=BORDER_COLOR)
    header_frame.pack(fill="x", padx=20, pady=10)
    
    title_container = ctk.CTkFrame(header_frame, fg_color="transparent")
    title_container.pack(side="right", padx=20, pady=10)
    ctk.CTkLabel(title_container, text="⚡ سامانه هوشمند مانیتورینگ انبار (Enterprise)", font=ctk.CTkFont(family="Tahoma", size=15, weight="bold"), text_color=ACCENT_EMERALD).pack(anchor="e")
    ctk.CTkLabel(title_container, text="مدیریت لحظه‌ای، یکپارچه‌سازی و هماهنگی وب‌سایت با دیتابیس", font=ctk.CTkFont(family="Tahoma", size=10), text_color=TEXT_MUTED).pack(anchor="e", pady=(2, 0))
    
    tools_container = ctk.CTkFrame(header_frame, fg_color="transparent")
    tools_container.pack(side="left", padx=20, pady=10)
    
    def light_btn(parent, text, bg, fg="#ffffff", hover=None, cmd=None):
        return ctk.CTkButton(parent, text=text, fg_color=bg, hover_color=hover, text_color=fg, font=("Tahoma", 11, "bold"), height=35, corner_radius=10, command=cmd)

    light_btn(tools_container, "🌐 دریافت خودکار اطلاعات", "#059669", hover="#047857", cmd=auto_update_cookie_action).pack(side="right", padx=(6, 0))
    light_btn(tools_container, "⚙️ تنظیمات API", "#e2e8f0", fg="#1e293b", hover="#cbd5e1", cmd=show_settings_dialog).pack(side="right", padx=(6, 0))
    
    startup_folder = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')
    is_in_startup = os.path.exists(os.path.join(startup_folder, "StockSync.lnk"))
    
    def toggle_startup():
        nonlocal is_in_startup
        if is_in_startup:
            if remove_from_startup():
                is_in_startup = False
                startup_btn.configure(text="🚀 افزودن به استارت‌آپ", fg_color="#dbeafe", text_color=ACCENT_BLUE)
                messagebox.showinfo("موفق", "برنامه از استارت‌آپ حذف شد.")
        else:
            if add_to_startup():
                is_in_startup = True
                startup_btn.configure(text="🗑️ حذف از استارت‌آپ", fg_color="#fee2e2", text_color="#991b1b")
                messagebox.showinfo("موفق", "برنامه به استارت‌آپ اضافه شد.")
    
    startup_btn = light_btn(tools_container, "🗑️ حذف از استارت‌آپ" if is_in_startup else "🚀 افزودن به استارت‌آپ", "#fee2e2" if is_in_startup else "#dbeafe", fg="#991b1b" if is_in_startup else ACCENT_BLUE, cmd=toggle_startup)
    startup_btn.pack(side="right", padx=(0, 0))

    metrics_grid = ctk.CTkFrame(gui_root, fg_color="transparent")
    metrics_grid.pack(fill="x", padx=20, pady=(0, 8))
    
    def create_light_kpi(parent, title, value_var_ref, status_text, accent_color):
        card = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=14, border_width=1, border_color=BORDER_COLOR)
        card.pack(side="right", fill="x", expand=True, padx=4)
        
        indicator = ctk.CTkFrame(card, fg_color=accent_color, height=3, corner_radius=0)
        indicator.pack(fill="x", padx=1, pady=(1, 0))
        
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=12, pady=8)
        
        ctk.CTkLabel(inner, text=title, font=("Tahoma", 10), text_color=TEXT_MUTED).pack(anchor="e")
        val_frame = ctk.CTkFrame(inner, fg_color="transparent")
        val_frame.pack(fill="x", pady=2)
        
        ctk.CTkLabel(val_frame, text=status_text, font=("Tahoma", 10, "bold"), text_color=accent_color).pack(side="left", pady=2)
        val_lbl = ctk.CTkLabel(val_frame, text=value_var_ref, font=("Tahoma", 18, "bold"), text_color=TEXT_MAIN)
        val_lbl.pack(side="right")
        return val_lbl

    clock_card = ctk.CTkFrame(metrics_grid, fg_color=BG_CARD, corner_radius=14, border_width=1, border_color=BORDER_COLOR)
    clock_card.pack(side="right", fill="x", expand=True, padx=4)
    clock_ind = ctk.CTkFrame(clock_card, fg_color=ACCENT_BLUE, height=3, corner_radius=0)
    clock_ind.pack(fill="x", padx=1, pady=(1, 0))
    clock_inner = ctk.CTkFrame(clock_card, fg_color="transparent")
    clock_inner.pack(fill="both", expand=True, padx=12, pady=8)
    ctk.CTkLabel(clock_inner, text="ساعت جاری سیستم", font=("Tahoma", 10), text_color=TEXT_MUTED).pack(anchor="e")
    clock_val_frame = ctk.CTkFrame(clock_inner, fg_color="transparent")
    clock_val_frame.pack(fill="x", pady=2)
    ctk.CTkLabel(clock_val_frame, text="● آنلاین", font=("Tahoma", 10, "bold"), text_color=ACCENT_BLUE).pack(side="left", pady=2)
    clock_label = ctk.CTkLabel(clock_val_frame, text="00:00:00", font=("Tahoma", 16, "bold"), text_color=TEXT_MAIN)
    clock_label.pack(side="right")

    kpi_repo_label = create_light_kpi(metrics_grid, "وضعیت مخزن سفارشات", "۰ کالا", "در انتظار فاکتور", ACCENT_PURPLE)
    kpi_out_label = create_light_kpi(metrics_grid, "مجموع کالای خروجی", "۰", "خروج از انبار", ACCENT_ORANGE)
    kpi_in_label = create_light_kpi(metrics_grid, "مجموع کالای ورودی", "۰", "ورود به انبار", ACCENT_EMERALD)

    gui_status_label = ctk.CTkLabel(gui_root, text="● سیستم در حال آماده‌سازی...", font=("Tahoma", 9), text_color=TEXT_MUTED)
    gui_status_label.pack(anchor="w", padx=22, pady=(0, 4))

    notebook = ctk.CTkTabview(gui_root, fg_color=BG_CARD, 
                             segmented_button_fg_color="#e2e8f0",
                             segmented_button_selected_color=ACCENT_EMERALD, 
                             segmented_button_selected_hover_color="#047857",
                             segmented_button_unselected_color="#e2e8f0",
                             segmented_button_unselected_hover_color="#cbd5e1",
                             corner_radius=16, border_width=1, border_color=BORDER_COLOR)
    notebook.pack(fill="both", expand=True, padx=20, pady=(0, 8))
    notebook._segmented_button.configure(font=("Tahoma", 12, "bold"), height=40, text_color="#1e293b")
    
    tab_main = notebook.add("    📋 مانیتورینگ زنده و رویدادها    ")
    repo_frame = notebook.add("    📦 مدیریت مخزن سفارشات    ")
    tab_info = notebook.add("    ℹ️ اطلاعات سیستمی API    ")
    
    main_frame = ctk.CTkFrame(tab_main, fg_color=BG_ELEVATED, corner_radius=12, border_width=1, border_color=BORDER_COLOR)
    main_frame.pack(fill="both", expand=True, padx=2, pady=2)
    
    main_stats_label = ctk.CTkLabel(main_frame, text="Total: 0 | ✅ 0 | ❌ 0", font=("Tahoma", 10, "bold"), text_color=TEXT_MAIN)
    main_stats_label.pack(anchor="e", padx=16, pady=(8, 4))
    
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("LightGrid.Treeview", background=BG_ELEVATED, foreground=TEXT_MAIN, rowheight=40, fieldbackground=BG_ELEVATED, borderwidth=0, font=("Tahoma", 10))
    style.configure("LightGrid.Treeview.Heading", background=GRID_HEADER_BG, foreground="#0f172a", borderwidth=0, font=("Tahoma", 11, "bold"))
    
    columns = ("code", "name", "old_stock", "new_stock", "time", "status")
    main_tree = ttk.Treeview(main_frame, columns=columns, show="headings", style="LightGrid.Treeview", height=10)
    for col, txt, w, anc in [("code", "کد کالا (SKU)", 110, "center"), ("name", "شرح کامل کالا / مشخصات محصول", 440, "e"), ("old_stock", "موجودی قبلی سایت", 130, "center"), ("new_stock", "موجودی جدید DB", 130, "center"), ("time", "ساعت ثبت رویداد", 120, "center"), ("status", "وضعیت همگام‌سازی", 220, "center")]:
        main_tree.heading(col, text=txt)
        main_tree.column(col, width=w, anchor=anc)
        
    main_tree.tag_configure("success", background="#d1fae5", foreground="#065f46")
    main_tree.tag_configure("failed", background="#fee2e2", foreground="#991b1b")
    main_tree.tag_configure("increase", background="#e0f2fe", foreground="#0369a1")
    main_tree.tag_configure("decrease", background="#fce7f3", foreground="#be185d")
    
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=main_tree.yview)
    main_tree.configure(yscrollcommand=scrollbar.set)
    main_tree.pack(side="right", fill="both", expand=True, padx=(0, 8), pady=(0, 8))
    scrollbar.pack(side="left", fill="y", padx=(8, 0), pady=(0, 8))

    repo_frame_inner = ctk.CTkFrame(repo_frame, fg_color=BG_ELEVATED, corner_radius=12, border_width=1, border_color=BORDER_COLOR)
    repo_frame_inner.pack(fill="both", expand=True, padx=2, pady=2)
    
    repo_stats_label = ctk.CTkLabel(repo_frame_inner, text="تعداد در مخزن: 0", font=("Tahoma", 10, "bold"), text_color=TEXT_MAIN)
    repo_stats_label.pack(anchor="e", padx=16, pady=(8, 4))
    
    repo_columns = ("code", "name", "site_stock", "repo_stock", "time", "status")
    repo_tree = ttk.Treeview(repo_frame_inner, columns=repo_columns, show="headings", style="LightGrid.Treeview", height=10)
    for col, txt, w, anc in [("code", "کد کالا", 110, "center"), ("name", "نام کالا", 440, "e"), ("site_stock", "موجودی سایت", 130, "center"), ("repo_stock", "موجودی مخزن", 130, "center"), ("time", "زمان ثبت", 150, "center"), ("status", "وضعیت", 220, "center")]:
        repo_tree.heading(col, text=txt)
        repo_tree.column(col, width=w, anchor=anc)
        
    repo_tree.tag_configure("repo_item", background=BG_ELEVATED, foreground=TEXT_MAIN)
    repo_tree.bind("<Double-1>", lambda e: remove_from_repo(repo_tree.item(repo_tree.selection()[0], "values")[0]) if repo_tree.selection() else None)
    
    repo_scrollbar = ttk.Scrollbar(repo_frame_inner, orient="vertical", command=repo_tree.yview)
    repo_tree.configure(yscrollcommand=repo_scrollbar.set)
    repo_tree.pack(side="right", fill="both", expand=True, padx=(0, 8), pady=(0, 8))
    repo_scrollbar.pack(side="left", fill="y", padx=(8, 0), pady=(0, 8))

    info_frame = ctk.CTkFrame(tab_info, fg_color=BG_ELEVATED, corner_radius=12, border_width=1, border_color=BORDER_COLOR)
    info_frame.pack(fill="both", expand=True, padx=2, pady=2)
    
    info_text = scrolledtext.ScrolledText(info_frame, font=("Consolas", 10), bg=BG_ELEVATED, fg=TEXT_MAIN, height=15, wrap=tk.WORD)
    info_text.pack(fill="both", expand=True, padx=10, pady=10)
    
    refresh_info_btn = ctk.CTkButton(info_frame, text="🔄 بروزرسانی اطلاعات", command=update_info_tab, fg_color=ACCENT_EMERALD, text_color="#ffffff", font=("Tahoma", 10, "bold"), height=35)
    refresh_info_btn.pack(pady=10)
    
    update_info_tab()
    update_clock()
    
    ctk.CTkLabel(gui_root, text="سامانه همگام‌سازی موجودی کالا", font=("Tahoma", 9), text_color=TEXT_MUTED).pack(pady=(0, 4))
    return gui_root

def run_sync():
    logger.info(f"\n{Colors.CYAN}{Colors.BOLD}{'=' * 60}{Colors.RESET}")
    logger.info(f"{Colors.CYAN}{Colors.BOLD}سامانه همگام‌سازی موجودی کالا{Colors.RESET}")
    logger.info(f"{Colors.CYAN}{Colors.BOLD}{'=' * 60}{Colors.RESET}\n")
    logger.info(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Mode: همگام‌سازی پایگاه داده و فروشگاه")
    logger.info(f"Check interval: 60 seconds")
    logger.info(f"Excel Log: {excel_log_file}")
    logger.info(f"Repository: {REPO_FILE}")
    logger.info(f"Last stock file: {LAST_STOCK_FILE}")
    logger.info(f"Last DB file: {LAST_DB_FILE}")
    logger.info(f"Config file: {CONFIG_FILE}")
    logger.info(f"BASE_URL: {BASE_URL}")
    logger.info(f"App Path: {APP_PATH}")
    
    update_status("Active - Full Source Detection")
    
    while sync_running:
        try:
            logger.info(f"\n{Colors.CYAN}--- Sync check at {datetime.now().strftime('%H:%M:%S')} ---{Colors.RESET}")
            process_stock_sync()
            logger.info(f"Next check in 60 seconds...")
            time.sleep(60)
        except KeyboardInterrupt:
            logger.info(f"\n{Colors.YELLOW}Shutting down by user request...{Colors.RESET}")
            break
        except Exception as e:
            logger.error(f"Error in sync cycle: {e}")
            update_status(f"Error - Retrying in 60s")
            time.sleep(60)
    
    logger.info(f"{Colors.GREEN}Program terminated.{Colors.RESET}")

if __name__ == "__main__":
    load_config()
    init_excel_log()
    root = create_gui()
    update_repo_table()
    sync_thread = threading.Thread(target=run_sync, daemon=True)
    sync_thread.start()
    root.mainloop()
    sync_running = False
    logger.info("GUI closed. Program terminated.")

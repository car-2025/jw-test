# jw_search_app_v12_edge_fixed9.py  — Part 1/4
# JW.org 抽出アプリ v12 — Edge (fixed9)
# - Google (google.co.jp) 検索経由で site:jw.org を取得（rel/date 各最大50）
# - EdgeDriver バージョン整備済み前提（ユーザーが msedgedriver を更新済み）
# - 自動化フラグ低減 / キャッシュ無効化 / user-data-dir 指定オプションを追加
# - requests -> selenium fallback の堅牢な抽出フロー
# - GUI と Excel 書き出しは次パートで追加

import os
import re
import time
import random
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import requests
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Optional Excel support
try:
    import openpyxl
    from openpyxl import Workbook
except Exception:
    openpyxl = None

# ----------------------------
# Configuration (editable)
# ----------------------------
EDGE_DRIVER_PATH = r"C:\Users\retec\Desktop\jw_test\msedgedriver.exe"
# Use a dedicated user-data-dir to avoid temp profile issues (must exist / writable)
EDGE_USER_DATA_DIR = r"C:\Users\retec\Desktop\jw_test\edge_profile"
BASE_DOMAIN = "https://www.jw.org"
GOOGLE_SEARCH_TPL = "https://www.google.co.jp/search?q=site%3Ajw.org+{}&num={}&start={}"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122 Safari/537.36"}
MAX_PER_MODE = 50
PAGE_STEP = 10
SELENIUM_PAGE_TIMEOUT = 22
EXCEL_PATH = "jw_extracted_fixed9.xlsx"
BACKGROUND_SLEEP = 0.15

# ----------------------------
# Utilities
# ----------------------------
def safe_filename(s: str) -> str:
    if not s:
        return "untitled"
    return re.sub(r'[\\/*?:"<>|]', "_", s)[:120]

def jp_char_count(s: str) -> int:
    return len(re.findall(r'[ぁ-んァ-ヴ一-龠々]', s or ''))

def extract_docid_from_url(url: str):
    if not url:
        return None
    # /d/<digits> pattern (common)
    m = re.search(r'/d/(\d{6,})', url)
    if m:
        try:
            return int(m.group(1))
        except:
            return None
    # trailing numeric segment /12345678/
    m2 = re.search(r'/(\d{6,})/?$', url)
    if m2:
        try:
            return int(m2.group(1))
        except:
            return None
    # any long numeric
    m3 = re.search(r'(\d{7,})', url)
    if m3:
        try:
            return int(m3.group(1))
        except:
            return None
    return None

# ----------------------------
# Excel writer (thread-safe)
# ----------------------------
class ExcelWriter:
    def __init__(self, path=EXCEL_PATH):
        self.path = path
        self._lock = threading.Lock()
        if openpyxl is None:
            return
        if not os.path.exists(self.path):
            wb = Workbook()
            ws = wb.active
            ws.title = "data"
            ws.append(["timestamp", "url", "title", "summary", "body"])
            wb.save(self.path)

    def append(self, row):
        if openpyxl is None:
            print("openpyxl not installed: skipping excel save")
            return
        with self._lock:
            try:
                wb = openpyxl.load_workbook(self.path)
                ws = wb["data"]
                ws.append(row)
                wb.save(self.path)
            except Exception as e:
                print("Excel write error:", e)

# ----------------------------
# make_edge_driver: stronger anti-detection + no-temp-profile fallback
# ----------------------------
def make_edge_driver(headed=True, driver_path=EDGE_DRIVER_PATH, user_data_dir=EDGE_USER_DATA_DIR):
    opts = Options()
    opts.use_chromium = True

    # Anti-detection flags
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    # disable automation controlled blink feature
    opts.add_argument("--disable-blink-features=AutomationControlled")

    # Reduce cache / profile issues
    opts.add_argument("--disable-application-cache")
    opts.add_argument("--disk-cache-size=0")
    opts.add_argument("--disable-gpu-shader-disk-cache")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--lang=ja-JP")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-features=NetworkService,NetworkServiceInProcess")

    # Use an explicit user-data-dir (not system temp) to avoid broken temp profiles
    if not os.path.exists(user_data_dir):
        try:
            os.makedirs(user_data_dir, exist_ok=True)
        except Exception as e:
            print("could not create user_data_dir:", e)

    # point to that profile (prevents creation in Temp with bad perms)
    opts.add_argument(f'--user-data-dir={user_data_dir}')

    # set a randomized window size to look more like a human browser
    width = random.choice([1200, 1280, 1366, 1440])
    height = random.choice([800, 900, 768, 1024])
    opts.add_argument(f"--window-size={width},{height}")

    # Headless is more likely to trigger bot detection — avoid unless explicitly needed
    if not headed:
        # modern headless flag
        opts.add_argument("--headless=new")
        # but still add a few options to reduce detection
        opts.add_argument("--disable-gpu")
    else:
        opts.add_argument("--start-maximized")

    # set a common user agent (not the selenium one)
    # HEADERS constant also used for requests; keep them aligned
    opts.add_argument(f'--user-agent={HEADERS["User-Agent"]}')

    # Create service and driver
    service = Service(driver_path)
    # suppress service logs where possible (selenium 4+ honors arguments)
    try:
        driver = webdriver.Edge(service=service, options=opts)
        # attempt to remove the webdriver attribute from window.navigator if possible
        try:
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                """
            })
        except Exception:
            pass
        driver.set_page_load_timeout(40)
        return driver
    except Exception as e:
        # Last-resort fallback: try without user-data-dir (may create temp profile)
        try:
            print("Primary driver start failed, retrying without user-data-dir —", e)
            opts_fallback = Options()
            opts_fallback.use_chromium = True
            opts_fallback.add_argument("--disable-application-cache")
            opts_fallback.add_argument("--disk-cache-size=0")
            opts_fallback.add_argument("--disable-gpu-shader-disk-cache")
            opts_fallback.add_argument("--disable-dev-shm-usage")
            opts_fallback.add_argument("--no-sandbox")
            opts_fallback.add_argument("--lang=ja-JP")
            opts_fallback.add_argument("--disable-extensions")
            opts_fallback.add_argument("--disable-background-networking")
            opts_fallback.add_argument("--disable-features=NetworkService,NetworkServiceInProcess")
            opts_fallback.add_argument(f'--user-agent={HEADERS["User-Agent"]}')
            if not headed:
                opts_fallback.add_argument("--headless=new")
            else:
                opts_fallback.add_argument("--start-maximized")
            service2 = Service(driver_path)
            driver2 = webdriver.Edge(service=service2, options=opts_fallback)
            try:
                driver2.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
                })
            except Exception:
                pass
            driver2.set_page_load_timeout(40)
            return driver2
        except Exception as e2:
            print("Fallback driver start failed:", e2)
            raise

# End of Part1


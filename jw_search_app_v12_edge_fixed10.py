# jw_search_app_v12_edge_fixed10.py — Part 1/4
# JW.org 自動検索・抽出・要約アプリ v12 fixed10
# - JW.org 公式検索 (/ja/search/?q=...) を直接使用（Google廃止）
# - rel/date 各最大50件（合計100件）取得
# - カテゴリページは排除、本文抽出精度強化
# - EdgeDriver はユーザーが更新済み（142に合わせること推奨）
# - GUI は v12 ベース（選択/解除/要約API欄あり）

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
# Configuration
# ----------------------------
EDGE_DRIVER_PATH = r"C:\Users\retec\Desktop\jw_test\msedgedriver.exe"
EDGE_USER_DATA_DIR = r"C:\Users\retec\Desktop\jw_test\edge_profile_fixed10"
BASE_DOMAIN = "https://www.jw.org"
SEARCH_URL_RELEVANCE_TPL = BASE_DOMAIN + "/ja/search/?q={}&sort=relevance&start={}"
SEARCH_URL_DATE_TPL = BASE_DOMAIN + "/ja/search/?q={}&sort=date&start={}"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122 Safari/537.36"}
MAX_PER_MODE = 50
PAGE_STEP = 10
SELENIUM_PAGE_TIMEOUT = 22
EXCEL_PATH = "jw_extracted_fixed10.xlsx"
BACKGROUND_SLEEP = 0.12

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
    m = re.search(r'/d/(\d{6,})', url)
    if m:
        try:
            return int(m.group(1))
        except:
            return None
    m2 = re.search(r'/(\d{6,})/?$', url)
    if m2:
        try:
            return int(m2.group(1))
        except:
            return None
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
# Edge driver factory — anti-detection & stable profile
# ----------------------------
def make_edge_driver(headed=True, driver_path=EDGE_DRIVER_PATH, user_data_dir=EDGE_USER_DATA_DIR):
    opts = Options()
    opts.use_chromium = True

    # Anti-detection experimental options
    try:
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
    except Exception:
        pass
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

    # Ensure user_data_dir exists to avoid Temp profile issues
    if not os.path.exists(user_data_dir):
        try:
            os.makedirs(user_data_dir, exist_ok=True)
        except Exception as e:
            print("could not create user_data_dir:", e)

    opts.add_argument(f'--user-data-dir={user_data_dir}')

    # randomized window size
    width = random.choice([1200, 1280, 1366, 1440])
    height = random.choice([800, 900, 768, 1024])
    opts.add_argument(f"--window-size={width},{height}")

    if not headed:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
    else:
        opts.add_argument("--start-maximized")

    opts.add_argument(f'--user-agent={HEADERS["User-Agent"]}')

    service = Service(driver_path)
    try:
        driver = webdriver.Edge(service=service, options=opts)
        # attempt to hide webdriver flag
        try:
            driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
            })
        except Exception:
            pass
        driver.set_page_load_timeout(40)
        return driver
    except Exception as e:
        print("Edge driver start failed:", e)
        # fallback without user_data_dir
        try:
            opts2 = Options()
            opts2.use_chromium = True
            opts2.add_argument("--disable-application-cache")
            opts2.add_argument("--disk-cache-size=0")
            opts2.add_argument("--disable-gpu-shader-disk-cache")
            opts2.add_argument("--no-sandbox")
            opts2.add_argument("--lang=ja-JP")
            opts2.add_argument(f'--user-agent={HEADERS["User-Agent"]}')
            if not headed:
                opts2.add_argument("--headless=new")
            else:
                opts2.add_argument("--start-maximized")
            driver2 = webdriver.Edge(service=Service(driver_path), options=opts2)
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

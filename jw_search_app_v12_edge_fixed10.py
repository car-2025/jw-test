# ============================================================
# GoogleFallbackSearcher（fixed10 完全版）
# JW公式検索が reCAPTCHA になるため Google検索へ切替する
#  rel / date のソートを Google 側パラメータで実現
# ============================================================

class JWOrgSearcher:
    """
    JW.org の公式検索ページ（/ja/search/?q=...）を直接開いて
    rel/date の各モードで URL を収集するシンプルなクラス。
    make_edge_driver を使って Edge を起動します。
    """
    def __init__(self, headed=True):
        # make_edge_driver は fixed10 の Part1 にある関数です
        try:
            self.driver = make_edge_driver(headed=headed)
        except Exception:
            # 最低限の起動方法（フォールバック）
            service = Service(EDGE_DRIVER_PATH)
            opts = Options()
            opts.use_chromium = True
            self.driver = webdriver.Edge(service=service, options=opts)
        # 少し余裕を持たせる
        self.driver.set_window_size(1200, 900)
        print("JWOrgSearcher: Edge 起動完了")

    def collect(self, keyword: str, mode: str, max_items: int):
        """
        mode: 'relevance' or 'date'
        返り値: URL のリスト（重複排除済）
        """
        if mode not in ("relevance", "date"):
            mode = "relevance"
        return jw_search_collect(self.driver, keyword, mode, max_items=max_items)

    def close(self):
        try:
            self.driver.quit()
        except:
            pass

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
# jw_search_app_v12_edge_fixed10.py — Part2/4
# === JW.org 公式検索（rel/date）URL収集ロジック ===

# ---------------------------------------------------------
# JW.org 公式検索：正規の検索URLで rel/date ページを巡回してリンク抽出
# ---------------------------------------------------------
def jw_search_collect(driver, keyword: str, mode: str, max_items=50):
    """JW.org 公式検索ページから正規の検索結果のみ抽出する"""
    assert mode in ("relevance", "date")
    tpl = SEARCH_URL_RELEVANCE_TPL if mode == "relevance" else SEARCH_URL_DATE_TPL

    collected = []
    visited_urls = set()
    pages = max(1, (max_items + PAGE_STEP - 1) // PAGE_STEP)

    for idx in range(pages):
        start = idx * PAGE_STEP
        search_url = tpl.format(keyword, start)

        try:
            driver.get(search_url)
        except Exception:
            continue

        # ページ読み込み待機
        try:
            WebDriverWait(driver, SELENIUM_PAGE_TIMEOUT).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "main, body"))
            )
        except Exception:
            time.sleep(1.2)

        time.sleep(1.0 + random.uniform(0.3, 0.8))

        # 検索結果が "該当なし" のケースを検出
        html = driver.page_source
        if "該当する結果は見つかりません" in html or "お探しのページが見つかりません" in html:
            break

        # --- 結果リンク抽出 ---
        anchors = driver.find_elements(By.CSS_SELECTOR, "a[href]")
        for a in anchors:
            try:
                href = a.get_attribute("href")
            except Exception:
                continue
            if not href:
                continue
            if href in visited_urls:
                continue
            visited_urls.add(href)

            # JW.org 内部のみ
            if not href.startswith(BASE_DOMAIN + "/ja/"):
                continue

            # カテゴリページなどは除外
            if any(x in href for x in [
                "/search/?", "/topics/", "/languages/", "/bible/", "/library/",
                "/study-tools/", "/bible-teachings/", "/videos/", "/news/",
                "/whats-new/"
            ]):
                continue

            # 記事 URL 判定
            # パターン： /d/123456789, /YYYYMM とか
            if extract_docid_from_url(href) is None:
                # docid が取れないページは本文が記事ではないのでスキップ
                continue

            collected.append(href)
            if len(collected) >= max_items:
                return collected

        # 次ページが無ければ終了
        if start > 0 and len(collected) == 0:
            # rel=0/date=0 件 → もう出ない
            break

    return collected[:max_items]

# ---------------------------------------------------------
# 本文抽出（requests版）
# ---------------------------------------------------------
def extract_article_body(url: str):
    """JW.org 記事ページの本文を正確に抽出する。カテゴリページは除外される"""
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        if r.status_code != 200:
            return "", ""
        soup = BeautifulSoup(r.text, "html.parser")

        # --- タイトル抽出 ---
        title_el = soup.find("h1")
        title = title_el.get_text(strip=True) if title_el else ""

        # --- 本文抽出 ---
        # パターン1: article[data-article-id]
        body_container = soup.find("article")
        if not body_container:
            # パターン2: div class="content" / "body" / "article-body"
            body_container = soup.find("div", class_=lambda c: c and ("content" in c or "body" in c))

        if not body_container:
            # パターン3: section 内の p
            body_container = soup.find("section")

        if not body_container:
            # fallback: p を全部
            paragraphs = [p.get_text(" ", strip=True) for p in soup.find_all("p")]
            return title, "\n".join(paragraphs)

        # 正規本文（p）抽出
        ps = body_container.find_all("p")
        body = "\n".join([p.get_text(" ", strip=True) for p in ps if p.get_text(strip=True)])

        return title, body

    except Exception:
        return "", ""

# End of Part2
# jw_search_app_v12_edge_fixed10.py — Part3/4
# === GUI + 検索処理 + 本文キャッシュ + 要約API入力欄 ===

class JWAppGUI:
    def __init__(self, master):
        self.master = master
        master.title("JW.org 検索・抽出・要約アプリ v12 — fixed10")
        master.geometry("1300x800")

-        # Selenium 検索器
-        self.searcher = GoogleFallbackSearcher()
+        # Selenium 検索器（JW.org公式検索を使う）
+        self.searcher = JWOrgSearcher()
        self.cached_body = {}     # URL → (title, body)
        self.current_url = None

        # Excel
        self.excel = ExcelWriter()

        # --- UI を構築 ---
        self.build_ui()

    # ---------------------------------------------------------
    # UI 構築
    # ---------------------------------------------------------
    def build_ui(self):
        top = ttk.Frame(self.master, padding=8)
        top.pack(fill="x")

        # 検索語
        ttk.Label(top, text="検索語:").pack(side="left")
        self.ent_keyword = ttk.Entry(top, width=30)
        self.ent_keyword.pack(side="left", padx=5)

        # 件数
        ttk.Label(top, text="関連度 件数:").pack(side="left")
        self.var_rel = tk.IntVar(value=50)
        ttk.Entry(top, textvariable=self.var_rel, width=6).pack(side="left")

        ttk.Label(top, text="新しい順 件数:").pack(side="left")
        self.var_date = tk.IntVar(value=50)
        ttk.Entry(top, textvariable=self.var_date, width=6).pack(side="left")

        ttk.Button(top, text="検索開始", command=self.start_search).pack(side="left", padx=10)

        # 要約API key
        ttk.Label(top, text="要約APIキー:").pack(side="left", padx=8)
        self.ent_api = ttk.Entry(top, width=40)
        self.ent_api.pack(side="left", padx=5)

        # -----------------------------------------------------
        # 左右分割
        pan = ttk.Panedwindow(self.master, orient=tk.HORIZONTAL)
        pan.pack(fill="both", expand=True)

        # --- 左側：URL リスト ---
        left = ttk.Frame(pan, padding=5)
        pan.add(left, weight=1)

        self.tree = ttk.Treeview(left, columns=("url"), show="headings")
        self.tree.heading("url", text="URL（ダブルクリックで本文表示）")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        # 全選択/解除
        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=5)
        ttk.Button(btns, text="全選択", command=self.select_all).pack(side="left", padx=4)
        ttk.Button(btns, text="全解除", command=self.clear_all).pack(side="left", padx=4)

        # --- 右側 ---
        right = ttk.Frame(pan, padding=5)
        pan.add(right, weight=3)

        # 本文表示
        ttk.Label(right, text="本文表示").pack(anchor="w")
        self.txt_article = tk.Text(right, wrap="word", height=25)
        self.txt_article.pack(fill="both", expand=True)

        # 要約ボタン
        ttk.Button(right, text="要約生成", command=self.make_summary).pack(pady=5)

        # 要約表示
        ttk.Label(right, text="要約結果").pack(anchor="w")
        self.txt_summary = tk.Text(right, wrap="word", height=10)
        self.txt_summary.pack(fill="x")

    # ---------------------------------------------------------
    # 全選択 / 全解除
    # ---------------------------------------------------------
    def select_all(self):
        for iid in self.tree.get_children():
            self.tree.selection_add(iid)

    def clear_all(self):
        self.tree.selection_remove(self.tree.get_children())

    # ---------------------------------------------------------
    # 検索開始
    # ---------------------------------------------------------
    def start_search(self):
        kw = self.ent_keyword.get().strip()
        if not kw:
            messagebox.showwarning("警告", "検索語を入力してください")
            return

        rel_n = self.var_rel.get()
        date_n = self.var_date.get()

        self.tree.delete(*self.tree.get_children())
        self.cached_body.clear()
        self.current_url = None

        print("=== 検索開始 ===")

-        # rel
-        rel_urls = self.searcher.google_fetch(kw, "relevance", rel_n)
-        print(f"[Google] rel collected {len(rel_urls)}")
-
-        # date
-        date_urls = self.searcher.google_fetch(kw, "date", date_n)
-        print(f"[Google] date collected {len(date_urls)}")
+        # JW.org 公式検索で取得
+        rel_urls = self.searcher.collect(kw, "relevance", rel_n)
+        print(f"[JW.org] rel collected {len(rel_urls)}")
+
+        date_urls = self.searcher.collect(kw, "date", date_n)
+        print(f"[JW.org] date collected {len(date_urls)}")

        date_urls = self.searcher.collect(kw, "date", date_n)
        print(f"[JW.org] date collected {len(date_urls)}")

        # 重複排除
        all_urls = []
        for u in rel_urls + date_urls:
            if u not in all_urls:
                all_urls.append(u)

        print(f"総取得 URL：{len(all_urls)} 件")

        # GUI に表示
        for url in all_urls:
            self.tree.insert("", "end", values=(url,))

        # バックグラウンド本文取得
        threading.Thread(target=self.fetch_body_background, args=(all_urls,), daemon=True).start()

    # ---------------------------------------------------------
    # バックグラウンド本文取得
    # ---------------------------------------------------------
    def fetch_body_background(self, urls):
        print("=== 本文バックグラウンド取得開始 ===")
        for url in urls:
            if url not in self.cached_body:
                title, body = extract_article_body(url)
                self.cached_body[url] = (title, body)
                time.sleep(0.3)
        print("=== 本文バックグラウンド取得完了 ===")

    # ---------------------------------------------------------
    # URL ダブルクリック → 本文表示
    # ---------------------------------------------------------
    def on_tree_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        url = self.tree.item(sel[0], "values")[0]
        self.current_url = url

        if url not in self.cached_body:
            title, body = extract_article_body(url)
            self.cached_body[url] = (title, body)
        else:
            title, body = self.cached_body[url]

        self.txt_article.delete("1.0", "end")
        self.txt_article.insert(
            "end",
            f"【タイトル】\n{title}\n\n【URL】\n{url}\n\n【本文】\n{body}"
        )

    # ---------------------------------------------------------
    # 要約
    # ---------------------------------------------------------
    def make_summary(self):
        if not self.current_url:
            return

        title, body = self.cached_body.get(self.current_url, ("", ""))
        if not body:
            return

        lines = body.split("\n")
        summary = "。".join(lines[:3]) + "。"

        self.txt_summary.delete("1.0", "end")
        self.txt_summary.insert("end", summary)

        # Excel 出力
        self.excel.append([
            datetime.now().isoformat(),
            self.current_url,
            title,
            summary,
            body
        ])

# End of Part3
# jw_search_app_v12_edge_fixed10.py — Part4/4
# === 起動部（main） ===

def main():
    root = tk.Tk()
    app = JWAppGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

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
# -------------------------------------------------------------
# Part 2/4 — Google 検索収集 / 記事判定 / 本文抽出（requests → Selenium fallback）
# -------------------------------------------------------------

# ----------------------------
# Google検索で jw.org の結果を集める
# ----------------------------
def google_collect_urls(driver, keyword: str, max_items=MAX_PER_MODE):
    """
    google.co.jp を使って site:jw.org + keyword の検索を行い、
    最大 max_items 件の jw.org URL を取得する（重複排除）。
    ページは num=10、start=0,10,20,... 最大5ページ（50件）。
    """
    results = []
    seen = set()
    pages = min((max_items + PAGE_STEP - 1) // PAGE_STEP, 5)

    for p in range(pages):
        start = p * PAGE_STEP
        url = GOOGLE_SEARCH_TPL.format(requests.utils.requote_uri(keyword), PAGE_STEP, start)
        try:
            driver.get(url)
        except Exception as e:
            # navigation failed — try small wait and continue
            print("Google page load failed:", e)
            time.sleep(1.0)
            continue

        # small randomized wait to mimic human
        time.sleep(0.9 + random.random() * 0.6)

        # collect anchor hrefs
        try:
            anchors = driver.find_elements(By.CSS_SELECTOR, "a[href]")
        except Exception:
            anchors = []

        for a in anchors:
            try:
                href = a.get_attribute("href") or ""
            except Exception:
                continue
            # filter Google redirect wrappers
            if href.startswith("https://www.google.co.jp/url?q=") or href.startswith("https://www.google.com/url?q="):
                m = re.search(r'[?&]q=([^&]+)', href)
                if m:
                    href = requests.utils.unquote(m.group(1))
                else:
                    continue
            # ignore non-jw links
            if BASE_DOMAIN not in href:
                continue
            # normalize (remove fragments)
            href = href.split('#')[0].rstrip('/')
            if href in seen:
                continue
            seen.add(href)
            # keep only likely article URLs; is_article_url will be conservative
            if is_article_url(href):
                results.append(href)
                if len(results) >= max_items:
                    return results[:max_items]
        # small delay between pages
        time.sleep(0.4 + random.random() * 0.6)

    return results[:max_items]


# ----------------------------
# 記事URL判定（より保守的）
# ----------------------------
def is_article_url(url: str) -> bool:
    """
    jw.org の記事ページを保守的に判定。
    条件例：
      - BASE_DOMAIN を含む
      - /ja/ を含む（日本語ページ）
      - '/d/' を含む（ドキュメントIDを持つ正式記事）
      - または長い数値を含む末尾セグメント
    除外：
      - /topics/, /languages/, /library/ (一般目次)、/collections/ などの一覧系
    """
    if not url or BASE_DOMAIN not in url:
        return False
    u = url.lower()
    # prefer Japanese pages
    if "/ja/" not in u:
        return False
    # exclude list/index pages
    exclude_patterns = [
        "/topics/", "/languages/", "/collections/", "/library/", "/languages/", "/search?", "/sitemap", "/about/"
    ]
    for ex in exclude_patterns:
        if ex in u:
            return False
    if u.endswith(".pdf"):
        return False
    # accept if contains /d/<digits>
    if re.search(r'/d/\d{6,}', u):
        return True
    # accept if trailing numeric id
    if re.search(r'/\d{6,}/?$', u):
        return True
    # otherwise conservative: reject
    return False


# ----------------------------
# HTML -> タイトル, 本文 抽出ロジック（改良版）
# ----------------------------
ARTICLE_SELECTORS_PRIORITY = [
    'article',
    'div[data-test-id="article-body"]',
    'div[class*="article"]',
    'div[class*="article-body"]',
    'div[class*="content__body"]',
    'main',
    'section'
]

def clean_text_block(text: str) -> str:
    if not text:
        return ''
    text = re.sub(r'\r', '', text)
    lines = [ln.strip() for ln in text.splitlines()]
    out = []
    for ln in lines:
        if not ln:
            out.append('')
            continue
        low = ln.lower()
        if any(k in low for k in ['privacy', 'cookie', 'terms', 'copyright', '利用規約']):
            continue
        if len(ln) < 2:
            continue
        out.append(ln)
    # collapse consecutive blanks
    final = []
    for ln in out:
        if ln == '' and final and final[-1] == '':
            continue
        final.append(ln)
    return '\n'.join(final).strip()

def parse_article_html(html: str):
    """
    HTML文字列からタイトルと本文を返す（title, body）
    本文は優先セレクタ順で長いブロックを採用し、日本語文字数フィルタあり
    """
    soup = BeautifulSoup(html, 'html.parser')
    # title
    title = ''
    h1 = soup.find('h1')
    if h1:
        title = h1.get_text(strip=True)
    elif soup.title:
        title = soup.title.get_text(strip=True)

    candidates = []
    for sel in ARTICLE_SELECTORS_PRIORITY:
        try:
            el = soup.select_one(sel)
        except Exception:
            el = None
        if el:
            txt = el.get_text('\n', strip=True)
            if len(txt) > 200:
                candidates.append(txt)

    # fallback: large divs with many <p>
    if not candidates:
        divs = soup.find_all('div')
        for d in divs:
            ps = d.find_all('p')
            if len(ps) >= 3:
                txt = d.get_text('\n', strip=True)
                if len(txt) > 200:
                    candidates.append(txt)

    # last fallback: all <p>
    if not candidates:
        ps = soup.find_all('p')
        body = '\n'.join([p.get_text(strip=True) for p in ps if p.get_text(strip=True)])
        if len(body) < 200:
            return title or '', ''
        if jp_char_count(body) < 10:
            return title or '', ''
        return title or '', clean_text_block(body)

    # pick longest candidate
    body = max(candidates, key=lambda s: len(s))
    if jp_char_count(body) < 15:
        return title or '', ''
    return title or '', clean_text_block(body)


# ----------------------------
# requestsベース取得（高速） + selenium fallback
# ----------------------------
def extract_article_body_requests(url: str):
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        if r.status_code != 200:
            return '', ''
        return parse_article_html(r.text)
    except Exception:
        return '', ''

def extract_article_body_selenium(driver, url: str):
    try:
        driver.get(url)
        WebDriverWait(driver, SELENIUM_PAGE_TIMEOUT).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body')))
        time.sleep(0.5 + random.random() * 0.6)
        html = driver.page_source
        return parse_article_html(html)
    except Exception as e:
        print("Selenium article fetch failed:", e)
        return '', ''

# End of Part2
# -------------------------------------------------------------
# Part 3/4 — GUI ロジック（検索、バックグラウンド本文取得、Tree操作）
# -------------------------------------------------------------

class ArticleCache:
    """URL → (title, body) を保持"""
    def __init__(self):
        self.data = {}

    def put(self, url, title, body):
        self.data[url] = (title, body)

    def get(self, url):
        return self.data.get(url, ("", ""))

    def has(self, url):
        return url in self.data


class JWSearcher:
    """Google検索を使って記事URLを収集し、必要に応じて Selenium で補完"""
    def __init__(self):
        service = Service(EDGE_DRIVER_PATH)
        self.driver = webdriver.Edge(service=service)
        self.driver.set_window_size(1300, 1000)
        print("EdgeDriver 起動 OK")

    def google_collect(self, keyword, mode, max_items):
        """
        mode = 'rel' | 'date'
        'date' は日付優先：site:jw.org/ja + YYYYフィルタ追加
        """
        if mode == 'rel':
            q = f"site:jw.org/ja {keyword}"
        else:
            q = f"site:jw.org/ja {keyword} 2024 OR 2023 OR 2022 OR 2021"

        print(f"[Google] mode={mode} keyword='{keyword}' → start collecting…")
        urls = google_collect_urls(self.driver, q, max_items=max_items)
        print(f"[Google] collected {len(urls)} items")
        return urls

    def fetch_body(self, url, cache: ArticleCache):
        """cache に無ければ取得（requests → selenium fallback）"""
        if cache.has(url):
            return cache.get(url)

        # try requests
        title, body = extract_article_body_requests(url)
        if title and body:
            cache.put(url, title, body)
            return title, body

        # fallback: selenium
        print("requests failed → Selenium fallback:", url)
        title, body = extract_article_body_selenium(self.driver, url)
        if title and body:
            cache.put(url, title, body)
        return title, body

class JWAppGUI:
    """Tk + Selenium + Google 検索のメインアプリケーション"""
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("JW.org 自動検索・抽出・要約アプリ v12 — fixed9 (Google対応)")
        self.root.geometry("1350x900")

        self.searcher = JWSearcher()
        self.cache = ArticleCache()
        self.excel = ExcelWriter()

        self.current_url = None
        self.tree_items = []  # list of URLs に対応

        self._build_ui()

    # -----------------------------------------------------
    def _build_ui(self):
        top = ttk.Frame(self.root, padding=6)
        top.pack(fill="x")

        ttk.Label(top, text="検索語：").pack(side="left")
        self.ent_kw = ttk.Entry(top, width=30)
        self.ent_kw.pack(side="left", padx=4)

        ttk.Label(top, text="関連度 件数：").pack(side="left", padx=(15,0))
        self.var_rel = tk.IntVar(value=MAX_PER_MODE)
        ttk.Entry(top, textvariable=self.var_rel, width=5).pack(side="left")

        ttk.Label(top, text="新しい順 件数：").pack(side="left", padx=(15,0))
        self.var_date = tk.IntVar(value=MAX_PER_MODE)
        ttk.Entry(top, textvariable=self.var_date, width=5).pack(side="left")

        ttk.Button(top, text="検索開始", command=self.start_search).pack(side="left", padx=10)

        # Selection buttons
        ttk.Button(top, text="全選択", command=self.select_all).pack(side="left", padx=4)
        ttk.Button(top, text="全解除", command=self.clear_selection).pack(side="left", padx=4)

        # -----------------------------------------
        # Paned window
        pan = ttk.Panedwindow(self.root, orient="horizontal")
        pan.pack(fill="both", expand=True)

        # 左ペイン
        left = ttk.Frame(pan)
        pan.add(left, weight=1)

        self.tree = ttk.Treeview(
            left,
            columns=("url",),
            show="headings",
            selectmode="extended"
        )
        self.tree.heading("url", text="抽出されたURL（コピー可）")
        self.tree.column("url", width=450)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_tree_dblclick)

        # copy on ctrl+c
        self.tree.bind("<Control-c>", self.copy_selected_urls)

        # 右ペイン
        right = ttk.Frame(pan)
        pan.add(right, weight=3)

        # 本文表示
        self.txt_body = tk.Text(right, wrap="word")
        self.txt_body.pack(fill="both", expand=True, padx=4, pady=4)

        # 要約ボタン
        ttk.Button(right, text="要約生成", command=self.make_summary).pack(pady=3)

        # 要約表示
        self.txt_summary = tk.Text(right, wrap="word", height=8)
        self.txt_summary.pack(fill="x", padx=4, pady=4)

    # -----------------------------------------------------
    def start_search(self):
        kw = self.ent_kw.get().strip()
        if not kw:
            messagebox.showwarning("警告", "検索語を入力してください。")
            return

        print("=== 検索開始 ===")

        # clear UI
        self.tree.delete(*self.tree.get_children())
        self.cache = ArticleCache()
        self.txt_body.delete("1.0", "end")
        self.txt_summary.delete("1.0", "end")
        self.tree_items = []

        # Google 関連度
        rel_n = self.var_rel.get()
        rel_urls = self.searcher.google_collect(kw, "rel", rel_n)

        # Google 新しい順（日付キーワード追加）
        date_n = self.var_date.get()
        date_urls = self.searcher.google_collect(kw, "date", date_n)

        # combine
        all_urls = rel_urls + [u for u in date_urls if u not in rel_urls]
        print(f"総取得 URL：{len(all_urls)} 件")

        # add to tree
        for u in all_urls:
            iid = self.tree.insert("", "end", values=(u,))
            self.tree_items.append(u)

        # background fetch
        threading.Thread(target=self._background_fetch_bodies, daemon=True).start()

    # -----------------------------------------------------
    def _background_fetch_bodies(self):
        print("=== 本文バックグラウンド取得開始 ===")
        for u in self.tree_items:
            if not self.cache.has(u):
                title, body = self.searcher.fetch_body(u, self.cache)
                if title and body:
                    print(f"[OK] {title[:20]}…")
                else:
                    print(f"[NG] 本文なし：{u}")
        print("=== 本文バックグラウンド取得完了 ===")

    # -----------------------------------------------------
    def on_tree_dblclick(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        url = self.tree.item(sel[0], "values")[0]
        self.current_url = url

        title, body = self.cache.get(url)
        if not body:
            print("cacheなし → 取得中…")
            title, body = self.searcher.fetch_body(url, self.cache)

        self.txt_body.delete("1.0", "end")
        self.txt_body.insert("end", f"【タイトル】\n{title}\n\n【URL】\n{url}\n\n【本文】\n{body}")

    # -----------------------------------------------------
    def copy_selected_urls(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        urls = []
        for iid in sel:
            u = self.tree.item(iid, "values")[0]
            urls.append(u)
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(urls))
        print("コピーしました")

    # -----------------------------------------------------
    def select_all(self):
        self.tree.selection_set(self.tree.get_children())

    def clear_selection(self):
        self.tree.selection_remove(self.tree.get_children())
# -------------------------------------------------------------
# Part 4/4 — 要約生成 + Excel 書き込み + main
# -------------------------------------------------------------

    # -----------------------------------------------------
    def make_summary(self):
        """非常にシンプルな3段落要約（必要に応じてAPI接続に置換可能）"""
        if not self.current_url:
            messagebox.showwarning("警告", "URLが選択されていません。")
            return

        title, body = self.cache.get(self.current_url)
        if not body:
            title, body = self.searcher.fetch_body(self.current_url, self.cache)
            if not body:
                messagebox.showwarning("警告", "本文が取得できませんでした。")
                return

        # 要約（3段落の先頭部分）
        lines = [ln for ln in body.split("\n") if ln.strip()]
        if not lines:
            summary = ""
        else:
            summary = "。".join(lines[:3]) + "。"

        # 表示
        self.txt_summary.delete("1.0", "end")
        self.txt_summary.insert("end", summary)

        # Excel 保存
        row = [
            datetime.now().isoformat(),
            self.current_url,
            title,
            summary,
            body
        ]
        self.excel.append(row)
        print("[Excel] 保存しました:", self.current_url)


# -------------------------------------------------------------
# main
# -------------------------------------------------------------
def main():
    root = tk.Tk()
    app = JWAppGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()


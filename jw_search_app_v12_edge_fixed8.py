# jw_search_app_v12_edge_fixed8.py  — Part 1/4
# JW.org 抽出アプリ v12 — Edge (fixed8, Google-search mode)
# - google.co.jp を使って site:jw.org 検索
# - rel: Google上位 (最大50) / date: docidの降順 (最大50)
# - requests -> Selenium fallback の二段式抽出
# - GUI は v12 ベース、選択/解除/個別トグル、API要約欄あり

import os
import re
import time
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
BASE_DOMAIN = "https://www.jw.org"
# use google.co.jp for Japanese-priority results
GOOGLE_SEARCH_TPL = "https://www.google.co.jp/search?q=site%3Ajw.org+{}&num={}&start={}"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT)"}
PAGE_SIZE = 10  # when iterating google start offsets (we will request num=10 per page)
MAX_PER_MODE = 50  # maximum per rel/date as requested
EXCEL_PATH = "jw_extracted_fixed8.xlsx"
SELENIUM_PAGE_TIMEOUT = 20
BACKGROUND_SLEEP = 0.18

# ----------------------------
# Helpers
# ----------------------------
def safe_filename(s: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", s)[:120]

def jp_char_count(s: str) -> int:
    return len(re.findall(r'[ぁ-んァ-ヴ一-龠々]', s or ''))

def extract_docid_from_url(url: str):
    """
    Try to extract JW.org docid (sequence of digits often present in article URLs).
    Return integer or None.
    Examples of patterns:
      .../d/1200001234/...
      .../wp201912/d/123456789/...
      .../library/.../123456789/  (digits in path)
    """
    if not url:
        return None
    # /d/ followed by digits
    m = re.search(r'/d/(\d{6,})', url)
    if m:
        try:
            return int(m.group(1))
        except:
            return None
    # trailing numeric segment like /123456789/
    m2 = re.search(r'/(\d{6,})/?$', url)
    if m2:
        try:
            return int(m2.group(1))
        except:
            return None
    # other numeric occurrences
    m3 = re.search(r'(\d{7,})', url)
    if m3:
        try:
            return int(m3.group(1))
        except:
            return None
    return None

# ----------------------------
# Excel writer
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
# Selenium driver factory
# ----------------------------
def make_edge_driver(headed=True, driver_path=EDGE_DRIVER_PATH):
    opts = Options()
    opts.use_chromium = True
    # Reduce disk/cache related logs and improve stability
    opts.add_argument("--disable-application-cache")
    opts.add_argument("--disk-cache-size=0")
    opts.add_argument("--disable-gpu-shader-disk-cache")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--lang=ja-JP")
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    # recommended additional flag to avoid some network service logs
    opts.add_argument("--disable-features=NetworkService")
    if not headed:
        opts.add_argument("--headless=new")
    else:
        opts.add_argument("--start-maximized")
    service = Service(driver_path)
    drv = webdriver.Edge(service=service, options=opts)
    drv.set_page_load_timeout(40)
    return drv

# End of Part1
# -------------------------------------------------------------
# Part 2/4 — Google 検索結果 → URL抽出 ＋ 本文抽出ロジック
# -------------------------------------------------------------

# ----------------------------
# Google search: collect jw.org links
# ----------------------------
def google_search_collect(driver, keyword: str, max_items=MAX_PER_MODE):
    """
    google.co.jp を使用し、site:jw.org を対象に最大 max_items 件まで URL を集める。
    - 1ページ num=10
    - start=0,10,20,30,40,… とページ送り
    """
    results = []
    seen = set()

    pages = max_items // 10 + 1
    pages = min(pages, 5)  # Google規制のため最大5ページ（50件）

    for p in range(pages):
        start = p * 10
        url = GOOGLE_SEARCH_TPL.format(keyword, 10, start)

        try:
            driver.get(url)
        except Exception:
            continue

        try:
            # ページ内の検索結果ブロックを待機
            WebDriverWait(driver, SELENIUM_PAGE_TIMEOUT).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a"))
            )
        except Exception:
            time.sleep(1)

        anchors = driver.find_elements(By.CSS_SELECTOR, "a[href]")
        for a in anchors:
            href = a.get_attribute("href") or ""
            if "jw.org" not in href:
                continue
            if href in seen:
                continue
            seen.add(href)

            # Googleの /url?q= 形式を除去
            if href.startswith("https://www.google.co.jp/url?q="):
                m = re.search(r"https://www.google.co.jp/url\\?q=([^&]+)", href)
                if m:
                    href = m.group(1)

            # jw.org の記事URLか判定
            if is_article_url(href):
                results.append(href)
                if len(results) >= max_items:
                    return results

        time.sleep(0.5)

    return results[:max_items]


# ----------------------------
# URL が jw.org の「記事」か判定（カテゴリや索引は除外）
# ----------------------------
def is_article_url(url: str) -> bool:
    if BASE_DOMAIN not in url:
        return False
    # カテゴリ・索引など不要URLの除外
    if any(x in url for x in [
        "/topics/", "/languages/", "/library/", "/videos/",
        "/music/", "/drama/", "/publications/", "/study-bible/"
    ]):
        return False

    # PDFなどは除外
    if url.endswith(".pdf"):
        return False

    return True


# ----------------------------
# 本文抽出（requests → fallback Selenium）
# ----------------------------
def parse_article_html(html: str):
    """HTMLからタイトルと本文を抽出"""
    soup = BeautifulSoup(html, "html.parser")

    # タイトル候補
    title_el = soup.find("h1")
    title = title_el.get_text(strip=True) if title_el else ""

    # 本文候補（複数パターン）
    body = ""

    # pattern A: articleタグ
    art = soup.find("article")
    if art:
        ps = art.find_all("p")
        if ps:
            body = "\n".join(p.get_text(strip=True) for p in ps)
            return title, body

    # pattern B: section
    sec = soup.find("section")
    if sec:
        ps = sec.find_all("p")
        if ps:
            body = "\n".join(p.get_text(strip=True) for p in ps)
            return title, body

    # pattern C: div class に body/content が含まれる領域
    div = soup.find("div", class_=re.compile(r"(body|content|article|main)"))
    if div:
        ps = div.find_all("p")
        if ps:
            body = "\n".join(p.get_text(strip=True) for p in ps)
            return title, body

    # fallback simple
    ps = soup.find_all("p")
    if ps:
        body = "\n".join(p.get_text(strip=True) for p in ps)

    return title, body


def extract_article_body_requests(url: str):
    try:
        html = requests.get(url, headers=HEADERS, timeout=12).text
        return parse_article_html(html)
    except Exception:
        return "", ""


# ----------------------------
# Selenium fallback（requests が空の場合）
# ----------------------------
def extract_article_body_selenium(driver, url: str):
    try:
        driver.get(url)
        WebDriverWait(driver, SELENIUM_PAGE_TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
        )
        html = driver.page_source
        return parse_article_html(html)
    except Exception:
        return "", ""


# End of Part2
# -------------------------------------------------------------
# Part 3/4 — GUI（左右ペイン／URLリスト／本文／要約エリア）
# -------------------------------------------------------------

class JWAppGUI:
    def __init__(self, master):
        self.master = master
        master.title("JW.org 検索・抽出・要約アプリ v12 (Edge, fixed8 Google版)")
        master.geometry("1400x900")

        # 状態管理
        self.driver = make_edge_driver(headed=True)
        self.excel = ExcelWriter()
        self.cache = {}        # url → (title, body)
        self.current_url = None
        self.selected = {}      # url → bool（チェック状態）
        self.api_key = tk.StringVar()  # 要約APIキー

        self._build_ui()

    # GUI構築
    def _build_ui(self):
        # 上部コントロールバー
        top = ttk.Frame(self.master, padding=6)
        top.pack(fill="x")

        ttk.Label(top, text="検索語:").pack(side="left")
        self.ent_kw = ttk.Entry(top, width=30)
        self.ent_kw.pack(side="left", padx=4)

        ttk.Label(top, text="関連度 件数(max50):").pack(side="left")
        self.var_rel = tk.IntVar(value=20)
        ttk.Entry(top, textvariable=self.var_rel, width=6).pack(side="left")

        ttk.Label(top, text="新しい順 件数(max50):").pack(side="left")
        self.var_new = tk.IntVar(value=20)
        ttk.Entry(top, textvariable=self.var_new, width=6).pack(side="left")

        ttk.Button(top, text="検索開始", command=self.start_search).pack(side="left", padx=10)

        # 左右分割
        pan = ttk.Panedwindow(self.master, orient=tk.HORIZONTAL)
        pan.pack(fill="both", expand=True)

        # 左：URLリスト
        left = ttk.Frame(pan)
        pan.add(left, weight=1)

        # チェックボックス画像
        self.img_unchecked = tk.PhotoImage(width=20, height=20)
        self.img_checked = tk.PhotoImage(width=20, height=20)
        # 塗りつぶし
        self.img_unchecked.put(("white",), to=(0, 0, 19, 19))
        self.img_checked.put(("black",), to=(0, 0, 19, 19))

        # Treeview（チェックボックス付）
        self.tree = ttk.Treeview(left, columns=("url"), show="headings", height=25)
        self.tree.heading("url", text="抽出URL（ダブルクリックで表示／コピー可）")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>", self.on_tree_double)

        # ボタン行
        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=3)

        ttk.Button(btns, text="全選択", command=self.select_all).pack(side="left", padx=4)
        ttk.Button(btns, text="全解除", command=self.unselect_all).pack(side="left", padx=4)

        # 右：本文表示エリア＋要約
        right = ttk.Frame(pan)
        pan.add(right, weight=3)

        # 本文表示
        frm_body = ttk.Labelframe(right, text="記事本文")
        frm_body.pack(fill="both", expand=True, padx=4, pady=4)

        self.txt_body = tk.Text(frm_body, wrap="word")
        self.txt_body.pack(fill="both", expand=True)

        # 要約とAPIキー欄
        frm_sum = ttk.Labelframe(right, text="要約＆保存")
        frm_sum.pack(fill="both", expand=False, padx=4, pady=4)

        ttk.Label(frm_sum, text="APIキー:").pack(anchor="w")
        ttk.Entry(frm_sum, textvariable=self.api_key, width=40).pack(anchor="w", padx=4, pady=2)

        self.txt_sum = tk.Text(frm_sum, wrap="word", height=10)
        self.txt_sum.pack(fill="x", padx=4, pady=4)

        ttk.Button(frm_sum, text="選択記事を要約して Excel 保存", command=self.do_summary_all).pack(pady=4)

    # ---------------------------------------------------------
    # 検索開始
    # ---------------------------------------------------------
    def start_search(self):
        kw = self.ent_kw.get().strip()
        if not kw:
            messagebox.showwarning("警告", "検索語を入力してください。")
            return

        rel_n = min(max(self.var_rel.get(), 1), MAX_PER_MODE)
        new_n = min(max(self.var_new.get(), 1), MAX_PER_MODE)

        # リストクリア
        self.tree.delete(*self.tree.get_children())
        self.cache.clear()
        self.selected.clear()

        print("Google検索で抽出を開始…")

        # 関連度順（Google上位）
        rel = google_search_collect(self.driver, kw, rel_n)

        # 新しい順（docid降順）
        all_for_date = google_search_collect(self.driver, kw, MAX_PER_MODE)
        url_docid_pairs = []
        for u in all_for_date:
            docid = extract_docid_from_url(u)
            if docid:
                url_docid_pairs.append((u, docid))

        url_docid_pairs.sort(key=lambda x: x[1], reverse=True)
        date_urls = [u for (u, _) in url_docid_pairs[:new_n]]

        # 重複を除いて結合
        merged = rel + [u for u in date_urls if u not in rel]

        print(f"抽出完了: {len(merged)} 件")
        self.populate_tree(merged)

    # TreeviewへURL挿入
    def populate_tree(self, urls):
        for u in urls:
            self.selected[u] = False
            self.tree.insert("", "end", iid=u, values=(u,))

    # ---------------------------------------------------------
    # Treeview行ダブルクリック → 右側に本文表示
    # ---------------------------------------------------------
    def on_tree_double(self, event):
        item = self.tree.selection()
        if not item:
            return
        url = item[0]
        self.current_url = url

        # まだ取得していない場合は本文抽出
        if url not in self.cache:
            title, body = extract_article_body_requests(url)
            if not body:
                title, body = extract_article_body_selenium(self.driver, url)
            self.cache[url] = (title or "", body or "")

        title, body = self.cache[url]

        self.txt_body.delete("1.0", "end")
        self.txt_body.insert("end", f"【タイトル】\n{title}\n\n【URL】\n{url}\n\n【本文】\n{body}")

    # ---------------------------------------------------------
    # チェック操作
    # ---------------------------------------------------------
    def select_all(self):
        for u in list(self.selected.keys()):
            self.selected[u] = True
        print("全選択")
        self._refresh_selection_states()

    def unselect_all(self):
        for u in list(self.selected.keys()):
            self.selected[u] = False
        print("全解除")
        self._refresh_selection_states()

    def _refresh_selection_states(self):
        # Treeview の背景色で選択状態を見やすく
        for u, sel in self.selected.items():
            try:
                self.tree.item(u, tags=("sel" if sel else "unsel"))
            except:
                pass
        self.tree.tag_configure("sel", background="#d0ffd0")
        self.tree.tag_configure("unsel", background="white")

    # End of Part3
# -------------------------------------------------------------
# Part 4/4 — API要約・保存・起動ルーチン
# -------------------------------------------------------------

# ----------------------------
# API 要約（簡易ラッパー）
# ----------------------------
def call_chatgpt_api(api_key: str, text: str, max_tokens: int = 400):
    """
    OpenAI Chat Completions の簡易呼び出し例。
    ユーザーは有効な API key を設定してください。
    """
    if not api_key or not text:
        return "APIキーまたは本文が空です"
    try:
        endpoint = "https://api.openai.com/v1/chat/completions"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        payload = {
            "model": "gpt-4o-mini",  # 変更可
            "messages": [
                {"role": "system", "content": "あなたは日本語で記事の要約を作るアシスタントです。"},
                {"role": "user", "content": f"次の本文を日本語で3文程度に要約してください：\n\n{text}"}
            ],
            "max_tokens": max_tokens,
            "temperature": 0.2
        }
        r = requests.post(endpoint, headers=headers, json=payload, timeout=60)
        if r.status_code == 200:
            j = r.json()
            choices = j.get("choices") or []
            if choices:
                # Chat completions の形式に依存
                msg = choices[0].get("message") or {}
                return msg.get("content", "") or choices[0].get("text", "")
            return str(j)
        else:
            return f"API error {r.status_code}: {r.text[:400]}"
    except Exception as e:
        return f"API 呼び出し例外: {e}"

# ----------------------------
# JWApp GUI: API要約フック & 保存ロジック
# ----------------------------
def do_api_summary_for_url(app: JWAppGUI, url: str):
    """
    与えられた app インスタンスと URL に対して API 要約を呼ぶ（非同期）。
    """
    art = app.cache.get(url, ("",""))
    if not art or not art[1]:
        return "本文がありません"
    api_key = app.api_key.get().strip()
    if not api_key:
        return "APIキーが未設定です"
    body = art[1]
    res = call_chatgpt_api(api_key, body)
    # UI 更新は main スレッドで
    def update_ui():
        app.txt_sum.delete("1.0", "end")
        app.txt_sum.insert("end", res)
    app.master.after(0, update_ui)
    # also save to excel
    try:
        app.excel.append([datetime.now().isoformat(), url, art[0], res, body])
    except Exception:
        pass
    return "OK"

# Batch operation used by the GUI button "選択記事を要約して Excel 保存"
def batch_summarize_selected(app: JWAppGUI):
    sel_urls = [u for u, v in app.selected.items() if v]
    if not sel_urls:
        messagebox.showinfo("Info", "選択された記事がありません。")
        return
    def worker():
        for u in sel_urls:
            # ensure cached
            if not app.cache.get(u, ("",""))[1]:
                title, body = extract_article_body_requests(u)
                if not body:
                    title, body = extract_article_body_selenium(app.driver, u)
                app.cache[u] = (title or "", body or "")
            # local summary first
            title, body = app.cache[u]
            summary = simple_summary(body, n_sentences=3) if body else ""
            try:
                app.excel.append([datetime.now().isoformat(), u, title, summary, body])
            except Exception:
                pass
        app.master.after(0, lambda: messagebox.showinfo("完了", f"{len(sel_urls)} 件を保存しました。"))
    threading.Thread(target=worker, daemon=True).start()

# Hook the JWAppGUI method to call batch summary
JWAppGUI.do_summary_all = lambda self: batch_summarize_selected(self)
JWAppGUI.api_summarize_single = lambda self: do_api_summary_for_url(self, getattr(self, "current_url", None) or "")

# ----------------------------
# Entrypoint
# ----------------------------
def main():
    root = tk.Tk()
    app = JWAppGUI(root)
    # intercept close to quit driver safely
    def on_close():
        try:
            if hasattr(app, "driver") and app.driver:
                try:
                    app.driver.quit()
                except:
                    pass
        except:
            pass
        root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    main()


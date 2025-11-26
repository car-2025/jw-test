# JW.org 自動検索・抽出・要約アプリ v12 — Edge 対応 fixed5（構文修正版）
# ※ URL 抽出と本文抽出に集中した安定バージョン

import time
import re
import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import requests
from datetime import datetime
import openpyxl
from openpyxl import Workbook
import os

EDGE_DRIVER_PATH = r"C:/Users/retec/Desktop/jw_test/msedgedriver.exe"
BASE_DOMAIN = "https://www.jw.org"
SEARCH_URL_RELEVANCE = BASE_DOMAIN + "/ja/search/?q={}&sort=relevance&start={}"
SEARCH_URL_NEWEST = BASE_DOMAIN + "/ja/search/?q={}&sort=date&start={}"
PAGE_SIZE = 10
HEADERS = {"User-Agent": "Mozilla/5.0"}

# ------------------
# URL が記事ページか判定
# ------------------
def is_article_url(url: str) -> bool:
    if not url.startswith(BASE_DOMAIN):
        return False
    if any(x in url for x in ["/search/?", "/topics/", "/languages/"]):
        return False
    return True

# ------------------
# 記事本文抽出
# ------------------
def extract_article_body(url: str):
    try:
        html = requests.get(url, headers=HEADERS, timeout=15).text
        soup = BeautifulSoup(html, "html.parser")

        title_el = soup.find("h1")
        title = title_el.get_text(strip=True) if title_el else ""

        # 代表的な記事本文コンテナ
        body_container = (
            soup.find("article")
            or soup.find("section")
            or soup.find("div", class_=re.compile("body|content"))
        )

        if not body_container:
            ps = soup.find_all("p")
        else:
            ps = body_container.find_all("p")

        body = "\n".join([p.get_text(strip=True) for p in ps])
        return title, body

    except Exception:
        return "", ""

# ------------------
# Excel 管理
# ------------------
class ExcelWriter:
    def __init__(self, path="jw_output.xlsx"):
        self.path = path
        if not os.path.exists(path):
            wb = Workbook()
            ws = wb.active
            ws.title = "data"
            ws.append(["timestamp", "url", "title", "summary", "body"])
            wb.save(path)

    def append(self, row):
        wb = openpyxl.load_workbook(self.path)
        ws = wb["data"]
        ws.append(row)
        wb.save(self.path)


# ------------------
# Selenium: 検索と URL 収集
# ------------------
class JWSearcher:
    def __init__(self):
        service = Service(EDGE_DRIVER_PATH)
        self.driver = webdriver.Edge(service=service)
        self.driver.set_window_size(1200, 900)
        print("Edge ドライバを起動しました")

    def open(self):
        self.driver.get(BASE_DOMAIN + "/ja/")
        time.sleep(1.5)

    def search_direct(self, keyword: str, sort_mode: str, max_count: int):
        collected = []
        base = SEARCH_URL_RELEVANCE if sort_mode == "relevance" else SEARCH_URL_NEWEST
        pages = max(1, (max_count + PAGE_SIZE - 1) // PAGE_SIZE)

        for i in range(pages):
            url = base.format(keyword, i * PAGE_SIZE)
            self.driver.get(url)
            time.sleep(1.2)

            anchors = self.driver.find_elements(By.CSS_SELECTOR, "a[href]")
            for a in anchors:
                href = a.get_attribute("href")
                if not href:
                    continue
                if href in collected:
                    continue
                if not is_article_url(href):
                    continue

                collected.append(href)
                if len(collected) >= max_count:
                    break

            if len(collected) >= max_count:
                break

        return collected[:max_count]


# ------------------
# GUI
# ------------------
class JWAppEdgeGUI:
    def __init__(self, master):
        self.master = master
        master.title("JW.org 検索・抽出・要約アプリ v12 Edge fixed5")
        master.geometry("1200x750")

        self.searcher = JWSearcher()
        self.excel = ExcelWriter()
        self.cached = {}
        self.current_url = None

        self.build_ui()
        self.searcher.open()

    def build_ui(self):
        top = ttk.Frame(self.master, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="検索語:").pack(side="left")
        self.entry_kw = ttk.Entry(top, width=30)
        self.entry_kw.pack(side="left", padx=5)

        ttk.Label(top, text="関連度 件数:").pack(side="left")
        self.var_rel = tk.IntVar(value=10)
        ttk.Entry(top, textvariable=self.var_rel, width=5).pack(side="left")

        ttk.Label(top, text="新しい順 件数:").pack(side="left")
        self.var_new = tk.IntVar(value=10)
        ttk.Entry(top, textvariable=self.var_new, width=5).pack(side="left")

        ttk.Button(top, text="検索開始", command=self.start_search).pack(side="left", padx=10)

        # 左右分割
        pan = ttk.Panedwindow(self.master, orient=tk.HORIZONTAL)
        pan.pack(fill="both", expand=True)

        left = ttk.Frame(pan)
        pan.add(left, weight=1)

        self.tree = ttk.Treeview(left, columns=("url"), show="headings")
        self.tree.heading("url", text="URL")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_tree_click)

        right = ttk.Frame(pan)
        pan.add(right, weight=3)

        self.txt_article = tk.Text(right, wrap="word")
        self.txt_article.pack(fill="both", expand=True)

        ttk.Button(right, text="要約", command=self.make_summary).pack(pady=4)

        self.txt_summary = tk.Text(right, wrap="word", height=10)
        self.txt_summary.pack(fill="x")

    # ------------------
    def start_search(self):
        kw = self.entry_kw.get().strip()
        if not kw:
            messagebox.showwarning("警告", "検索語を入力してください")
            return

        self.tree.delete(*self.tree.get_children())
        self.cached.clear()

        rel_n = self.var_rel.get()
        new_n = self.var_new.get()

        print("検索を開始します…")

        rel = self.searcher.search_direct(kw, "relevance", rel_n)
        new = self.searcher.search_direct(kw, "date", new_n)

        all_urls = rel + [u for u in new if u not in rel]
        print(f"収集完了: {len(all_urls)} 件")

        for u in all_urls:
            self.tree.insert("", "end", values=(u,))

    # ------------------
    def on_tree_click(self, event):
        item = self.tree.selection()
        if not item:
            return
        url = self.tree.item(item, "values")[0]
        self.current_url = url

        if url not in self.cached:
            print("本文未取得 → requests で取得")
            title, body = extract_article_body(url)
            self.cached[url] = (title, body)
        else:
            title, body = self.cached[url]

        self.txt_article.delete("1.0", "end")
        self.txt_article.insert(
            "end",
            f"【タイトル】\n{title}\n\n【URL】{url}\n\n【本文】\n{body}"
        )

    # ------------------
    def make_summary(self):
        if not self.current_url:
            return
        _, body = self.cached.get(self.current_url, ("", ""))
        if not body:
            return

        lines = body.split("\n")
        summary = "。".join(lines[:3]) + "。"

        self.txt_summary.delete("1.0", "end")
        self.txt_summary.insert("end", summary)

        title, body = self.cached[self.current_url]
        self.excel.append([datetime.now().isoformat(), self.current_url, title, summary, body])


# ------------------
def main():
    root = tk.Tk()
    JWAppEdgeGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

from __future__ import annotations

from datetime import datetime
import os
import re
import time
from typing import Callable, Optional
from urllib.parse import urljoin

from bs4 import BeautifulSoup
import pandas as pd
import requests

BASE_URL = "https://books.toscrape.com/catalogue/"
PAGE_URL_TEMPLATE = "https://books.toscrape.com/catalogue/page-{page}.html"

RATING_MAP = {
    "One": 1,
    "Two": 2,
    "Three": 3,
    "Four": 4,
    "Five": 5,
}


def fetch_page(
    url: str,
    timeout: int = 10,
    max_retries: int = 3,
    retry_delay: float = 2.0,
    log_callback: Optional[Callable[[str], None]] = None,
) -> Optional[str]:
    """指定URLからHTMLを取得する（最大3回リトライ、UTF-8指定）。"""
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(url, timeout=timeout)
            response.raise_for_status()
            response.encoding = "utf-8"
            return response.text
        except requests.RequestException as e:
            msg = f"[警告] 通信エラー ({attempt}/{max_retries}回目): {url} - {e}"
            if log_callback:
                log_callback(msg)
            else:
                print(msg)
            if attempt < max_retries:
                time.sleep(retry_delay)

    err_msg = f"[エラー] 最大リトライ回数({max_retries}回)に達したため取得失敗: {url}"
    if log_callback:
        log_callback(err_msg)
    else:
        print(err_msg)
    return None


def parse_books_from_html(html: str, page_url: str = BASE_URL) -> list[dict]:
    """1ページ分のHTMLから商品リストを抽出する。"""
    soup = BeautifulSoup(html, "html.parser")
    articles = soup.select("article.product_pod")
    books = []

    for article in articles:
        # 1. 商品名 (title) & 5. 商品詳細URL (url)
        a_tag = article.select_one("h3 > a")
        title = a_tag.get("title", "").strip() if a_tag else ""
        href = a_tag.get("href", "") if a_tag else ""
        detail_url = urljoin(page_url, href)

        # 2. 価格 (price)
        price_tag = article.select_one("p.price_color")
        price_text = price_tag.get_text(strip=True) if price_tag else ""
        price_match = re.search(r"(\d+(?:\.\d+)?)", price_text)
        price = float(price_match.group(1)) if price_match else 0.0

        # 3. 在庫状況 (stock)
        stock_tag = article.select_one("p.instock.availability")
        stock = stock_tag.get_text(strip=True) if stock_tag else ""

        # 4. 評価 (rating)
        rating_tag = article.select_one("p.star-rating")
        rating = 0
        if rating_tag:
            classes = rating_tag.get("class", [])
            for cls in classes:
                if cls in RATING_MAP:
                    rating = RATING_MAP[cls]
                    break

        books.append(
            {
                "title": title,
                "price": price,
                "stock": stock,
                "rating": rating,
                "url": detail_url,
            }
        )

    return books


def scrape_books(
    max_pages: int = 5,
    sleep_sec: float = 1.0,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None,
    stop_check: Optional[Callable[[], bool]] = None,
) -> pd.DataFrame:
    """
    指定したページ数（1〜max_pages）を巡回し、
    商品名・価格(float)・在庫状況・評価(int)・詳細URLを取得してDataFrameで返す。
    """
    all_books: list[dict] = []

    for page in range(1, max_pages + 1):
        if stop_check and stop_check():
            msg = "[中断] ユーザー操作によりデータ収集を中断しました。"
            if log_callback:
                log_callback(msg)
            else:
                print(msg)
            break

        url = PAGE_URL_TEMPLATE.format(page=page)
        msg = f"{page} / {max_pages} ページ目を収集中... ({url})"
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

        if progress_callback:
            progress_callback(page, max_pages)

        html = fetch_page(url, log_callback=log_callback)
        if html is None:
            stop_msg = f"[停止] {page}ページ目の取得に失敗したため、処理を安全に停止します。"
            if log_callback:
                log_callback(stop_msg)
            else:
                print(stop_msg)
            break

        page_books = parse_books_from_html(html, page_url=url)
        all_books.extend(page_books)

        # 1ページ取得ごとに必ず1.0秒待機
        time.sleep(sleep_sec)

    columns = ["title", "price", "stock", "rating", "url"]
    return pd.DataFrame(all_books, columns=columns)


def filter_candidates(
    df: pd.DataFrame,
    min_rating: int = 4,
    max_price: float = 30.0,
) -> pd.DataFrame:
    """仕入れ候補の条件に合致するデータを抽出する。"""
    if df.empty:
        return df.copy()

    mask = (
        (df["rating"] >= min_rating)
        & (df["price"] <= max_price)
        & (df["stock"].astype(str).str.contains("In stock", na=False))
    )
    return df[mask].copy()


def export_to_excel(
    df: pd.DataFrame,
    output_dir: str = "output",
    min_rating: int = 4,
    max_price: float = 30.0,
) -> str:
    """全商品一覧と仕入れ候補の2シート構成でExcelファイルに出力する。"""
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"books_data_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)

    candidates_df = filter_candidates(df, min_rating=min_rating, max_price=max_price)

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="全商品一覧", index=False)
        candidates_df.to_excel(writer, sheet_name="仕入れ候補", index=False)

    return filepath


if __name__ == "__main__":
    print("=== スクレイピングテスト実行 (2ページ分) ===")
    df_result = scrape_books(max_pages=2)

    print(f"\n[取得件数]: {len(df_result)} 件")
    print("\n[先頭5件のデータ]:")
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 1000)
    print(df_result.head(5))

    print("\n[データ型確認]:")
    print(df_result.dtypes)

    if not df_result.empty:
        saved_path = export_to_excel(df_result, output_dir="output", min_rating=4, max_price=30.0)
        candidates = filter_candidates(df_result, min_rating=4, max_price=30.0)
        print(f"\n[仕入れ候補件数 (評価>=4, 価格<=30.0, In stock)]: {len(candidates)} 件")
        print(f"[Excel出力完了]: {saved_path}")
    else:
        print("\n[警告] データが0件のためExcel出力をスキップしました。")

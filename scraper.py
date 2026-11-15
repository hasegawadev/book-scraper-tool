from __future__ import annotations

from datetime import datetime
import json
import os
import re
import sys
import threading
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


def get_base_dir() -> str:
    """実行ファイル（.exe）またはスクリプトと同階層のディレクトリを取得する。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


LAST_SCRAPE_FILE = os.path.join(get_base_dir(), "last_scrape.json")


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


def load_last_scrape(filepath: str = LAST_SCRAPE_FILE) -> Optional[dict[str, float]]:
    """前回収集データ（last_scrape.json）を読み込み、{url: price} の辞書を返す。存在しない場合は None。"""
    if not os.path.exists(filepath):
        return None
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            records = json.load(f)
        if not isinstance(records, list):
            return None
        return {
            str(item["url"]): float(item["price"])
            for item in records
            if isinstance(item, dict) and "url" in item and "price" in item
        }
    except Exception:
        return None


def save_last_scrape(
    books: list[dict],
    previous_records: Optional[list[dict]] = None,
    filepath: str = LAST_SCRAPE_FILE,
) -> None:
    """今回取得した全商品データを last_scrape.json に保存する（前回の未巡回ページのデータも保持・更新）。"""
    merged: dict[str, dict] = {}
    if previous_records:
        for item in previous_records:
            if isinstance(item, dict) and "url" in item:
                merged[str(item["url"])] = item
    for book in books:
        if "url" in book:
            merged[str(book["url"])] = {
                "title": book.get("title", ""),
                "price": float(book.get("price", 0.0)),
                "stock": book.get("stock", ""),
                "rating": int(book.get("rating", 0)),
                "url": book.get("url", ""),
            }
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(list(merged.values()), f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def compute_status_diff(url: str, price: float, prev_map: Optional[dict[str, float]]) -> str:
    """前回データの {url: price} と比較して差分状況（status_diff）文字列を返す。"""
    if prev_map is None:
        return "初回取得"
    if url not in prev_map:
        return "NEW"
    prev_price = prev_map[url]
    diff = round(price - prev_price, 2)
    if diff < -1e-6:
        return f"値下がり (-£{abs(diff):.2f})"
    if diff > 1e-6:
        return f"値上がり (+£{diff:.2f})"
    return "変動なし"


def scrape_books(
    max_pages: int = 5,
    sleep_sec: float = 1.0,
    progress_callback: Optional[Callable[[int, int, str, list[dict]], None]] = None,
    log_callback: Optional[Callable[[str], None]] = None,
    stop_event: Optional[threading.Event] = None,
    stop_check: Optional[Callable[[], bool]] = None,
    last_scrape_path: str = LAST_SCRAPE_FILE,
) -> pd.DataFrame:
    """
    指定したページ数（1〜max_pages）を巡回し、
    商品名・価格(float)・在庫状況・評価(int)・詳細URL・差分状況(status_diff)を取得してDataFrameで返す。

    Args:
        max_pages: 取得する最大ページ数（1〜50）
        sleep_sec: 1ページ取得ごとの待機秒数（デフォルト: 1.0秒）
        progress_callback: ページ処理ごとに (current_page, total_pages, message, current_books) を通知する関数
        log_callback: ログメッセージを通知する関数（省略時は print 出力）
        stop_event: 中断制御用の threading.Event オブジェクト（is_set() が True なら安全に停止）
        stop_check: 中断判定用のコールバック関数（True を返したら安全に停止）
        last_scrape_path: 前回データ保存用JSONファイルのパス
    """
    all_books: list[dict] = []

    # 前回データの読み込み
    prev_map = load_last_scrape(last_scrape_path)
    prev_raw: Optional[list[dict]] = None
    if os.path.exists(last_scrape_path):
        try:
            with open(last_scrape_path, "r", encoding="utf-8") as f:
                prev_raw = json.load(f)
        except Exception:
            prev_raw = None

    def _is_stopped() -> bool:
        if stop_event is not None and stop_event.is_set():
            return True
        if stop_check is not None and stop_check():
            return True
        return False

    def _log(message: str) -> None:
        if log_callback:
            log_callback(message)
        else:
            print(message)

    for page in range(1, max_pages + 1):
        if _is_stopped():
            _log("[中断] ユーザー操作によりデータ収集を中断しました。")
            break

        url = PAGE_URL_TEMPLATE.format(page=page)
        status_msg = f"{page} / {max_pages} ページ目を収集中..."
        _log(f"{status_msg} ({url})")

        if progress_callback:
            progress_callback(page - 1, max_pages, status_msg, all_books)

        html = fetch_page(url, log_callback=log_callback)
        if html is None:
            _log(f"[停止] {page}ページ目の取得に失敗したため、処理を安全に停止します。")
            break

        page_books = parse_books_from_html(html, page_url=url)
        price_drop_count = 0
        new_item_count = 0
        for b in page_books:
            diff_str = compute_status_diff(b["url"], b["price"], prev_map)
            b["status_diff"] = diff_str
            if diff_str.startswith("値下がり"):
                price_drop_count += 1
            elif diff_str == "NEW":
                new_item_count += 1

        if price_drop_count > 0 or new_item_count > 0:
            _log(
                f"[完了] {page}ページ目の差分検知: 値下がり {price_drop_count} 件 / NEW {new_item_count} 件"
            )

        all_books.extend(page_books)

        if progress_callback:
            progress_callback(page, max_pages, status_msg, all_books)

        # 1ページ取得ごとに必ず1.0秒待機
        time.sleep(sleep_sec)

    if all_books:
        save_last_scrape(all_books, previous_records=prev_raw, filepath=last_scrape_path)

    columns = ["title", "price", "stock", "rating", "url", "status_diff"]
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
    """全商品一覧と仕入れ候補の2シート構成（status_diff列を含む）でExcelファイルに出力する。"""
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"books_data_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)

    export_df = df.copy()
    if "status_diff" not in export_df.columns:
        export_df["status_diff"] = "初回取得"

    candidates_df = filter_candidates(export_df, min_rating=min_rating, max_price=max_price)

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="全商品一覧", index=False)
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

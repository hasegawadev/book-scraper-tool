from __future__ import annotations

from datetime import datetime
import json
import os
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd

from scraper import (
    PAGE_URL_TEMPLATE,
    export_to_excel,
    fetch_page,
    filter_candidates,
    parse_books_from_html,
)

# 全体テーマの初期設定
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# カラーパレット定義（ダーク / ライト両対応、ダーク基調は #121212 / #1E1E1E / #2E2E2E / #10B981）
COLOR_BG = ("#F3F4F6", "#121212")
COLOR_CARD_BG = ("#FFFFFF", "#1E1E1E")
COLOR_CARD_BORDER = ("#E5E7EB", "#2E2E2E")
COLOR_INPUT_BG = ("#F9FAFB", "#141414")
COLOR_ACCENT = "#10B981"
COLOR_ACCENT_HOVER = "#059669"
COLOR_GHOST_BORDER = "#4B5563"
COLOR_GHOST_HOVER = ("#E5E7EB", "#2A2A2A")
COLOR_DANGER = "#EF4444"
COLOR_DANGER_HOVER = "#DC2626"
COLOR_START_DISABLED = ("#D1D5DB", "#2E2E2E")
COLOR_SECONDARY_BG = ("#E5E7EB", "#262626")
COLOR_SECONDARY_HOVER = ("#D1D5DB", "#333333")
COLOR_TEXT_MAIN = ("#111827", "#F9FAFB")
COLOR_TEXT_SUB = ("#6B7280", "#9CA3AF")

FONT_FAMILY = "Yu Gothic UI"


def get_base_dir() -> str:
    """実行ファイル（.exe）またはスクリプトと同階層のディレクトリを取得する。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_FILE = os.path.join(get_base_dir(), "config.json")
DEFAULT_OUTPUT_DIR = os.path.join(get_base_dir(), "output")


class BookScraperApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__(fg_color=COLOR_BG)
        self.title("ECサイト商品データ自動収集＆条件抽出ツール")
        self.geometry("820x740")
        self.minsize(720, 650)

        # 状態管理フラグ
        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None

        # 設定変数の初期化
        self.max_pages_var = ctk.StringVar(value="5")
        self.min_rating_var = ctk.StringVar(value="4")
        self.max_price_var = ctk.StringVar(value="30.0")
        self.output_dir_var = ctk.StringVar(value=DEFAULT_OUTPUT_DIR)

        # KPI / 進捗ステータス変数
        self.kpi_total_var = ctk.StringVar(value="0 件")
        self.kpi_candidates_var = ctk.StringVar(value="0 件")
        self.kpi_status_var = ctk.StringVar(value="待機中")
        self.progress_detail_var = ctk.StringVar(value="収集開始ボタンを押してスタートしてください")
        self.dark_mode_var = ctk.BooleanVar(value=True)

        # config.json から前回設定を読み込み
        self.load_config()

        # ダッシュボード風UI構築
        self._build_ui()

    def load_config(self) -> None:
        """config.json から前回の出力先フォルダ等の設定を読み込む。"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                saved_dir = data.get("output_dir")
                if saved_dir and isinstance(saved_dir, str):
                    self.output_dir_var.set(saved_dir)
            except Exception:
                pass

    def save_config(self) -> None:
        """現在の出力先フォルダパスを config.json に自動保存する。"""
        data = {
            "output_dir": self.output_dir_var.get().strip() or DEFAULT_OUTPUT_DIR,
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.append_log(f"[警告] 設定ファイルの保存に失敗しました: {e}")

    def toggle_appearance_mode(self) -> None:
        """ダークモード／ライトモードを切り替える。"""
        if self.dark_mode_var.get():
            ctk.set_appearance_mode("dark")
            self.theme_switch.configure(text="ダークモード")
        else:
            ctk.set_appearance_mode("light")
            self.theme_switch.configure(text="ライトモード")

    def _build_ui(self) -> None:
        """モダンSaaSダッシュボード風のUIレイアウトを構築する。"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)

        # フォント定義（すべて Yu Gothic UI に統一）
        font_title = ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold")
        font_section = ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold")
        font_body = ctk.CTkFont(family=FONT_FAMILY, size=13)
        font_body_bold = ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold")
        font_sub = ctk.CTkFont(family=FONT_FAMILY, size=12)
        font_kpi_val = ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold")
        font_kpi_lbl = ctk.CTkFont(family=FONT_FAMILY, size=12)

        # --- 0. ヘッダーエリア（タイトル＋右上ライト／ダークモード切替スイッチ） ---
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 8))
        header_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header_frame,
            text="ECサイト商品データ自動収集ダッシュボード",
            font=font_title,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        self.theme_switch = ctk.CTkSwitch(
            header_frame,
            text="ダークモード",
            variable=self.dark_mode_var,
            onvalue=True,
            offvalue=False,
            command=self.toggle_appearance_mode,
            font=font_sub,
            text_color=COLOR_TEXT_SUB,
            progress_color=COLOR_ACCENT,
        )
        self.theme_switch.grid(row=0, column=1, sticky="e")

        # --- 1. 設定項目カード ---
        settings_card = ctk.CTkFrame(
            self,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        settings_card.grid(row=1, column=0, sticky="ew", padx=24, pady=6)
        settings_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            settings_card,
            text="収集・抽出条件設定",
            font=font_section,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=20, pady=(14, 8))

        # 取得ページ数上限
        ctk.CTkLabel(
            settings_card,
            text="取得ページ数上限 (1〜50):",
            font=font_body,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=1, column=0, sticky="w", padx=(20, 14), pady=5)

        self.pages_entry = ctk.CTkEntry(
            settings_card,
            textvariable=self.max_pages_var,
            width=110,
            height=34,
            corner_radius=8,
            fg_color=COLOR_INPUT_BG,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            font=font_body,
        )
        self.pages_entry.grid(row=1, column=1, sticky="w", pady=5)

        ctk.CTkLabel(
            settings_card,
            text="※全件取得時は 50 を指定（1ページ20件）",
            font=font_sub,
            text_color=COLOR_TEXT_SUB,
        ).grid(row=1, column=2, sticky="w", padx=(12, 20), pady=5)

        # 仕入れ候補の最低評価
        ctk.CTkLabel(
            settings_card,
            text="仕入れ候補の最低評価 (1〜5):",
            font=font_body,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=2, column=0, sticky="w", padx=(20, 14), pady=5)

        self.rating_menu = ctk.CTkOptionMenu(
            settings_card,
            variable=self.min_rating_var,
            values=["1", "2", "3", "4", "5"],
            width=110,
            height=34,
            corner_radius=8,
            fg_color=COLOR_SECONDARY_BG,
            button_color=COLOR_SECONDARY_BG,
            button_hover_color=COLOR_SECONDARY_HOVER,
            text_color=COLOR_TEXT_MAIN,
            dropdown_fg_color=COLOR_CARD_BG,
            dropdown_text_color=COLOR_TEXT_MAIN,
            font=font_body,
            dropdown_font=font_body,
        )
        self.rating_menu.grid(row=2, column=1, sticky="w", pady=5)

        ctk.CTkLabel(
            settings_card,
            text="以上",
            font=font_sub,
            text_color=COLOR_TEXT_SUB,
        ).grid(row=2, column=2, sticky="w", padx=(12, 20), pady=5)

        # 仕入れ候補の上限価格
        ctk.CTkLabel(
            settings_card,
            text="仕入れ候補の上限価格 (£):",
            font=font_body,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=3, column=0, sticky="w", padx=(20, 14), pady=5)

        self.price_entry = ctk.CTkEntry(
            settings_card,
            textvariable=self.max_price_var,
            width=110,
            height=34,
            corner_radius=8,
            fg_color=COLOR_INPUT_BG,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            font=font_body,
        )
        self.price_entry.grid(row=3, column=1, sticky="w", pady=5)

        ctk.CTkLabel(
            settings_card,
            text="以下（かつ在庫「In stock」を含む）",
            font=font_sub,
            text_color=COLOR_TEXT_SUB,
        ).grid(row=3, column=2, sticky="w", padx=(12, 20), pady=5)

        # 出力先フォルダ（横幅いっぱいに広げてパス切れを防止）
        ctk.CTkLabel(
            settings_card,
            text="出力先フォルダ:",
            font=font_body,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=4, column=0, sticky="w", padx=(20, 14), pady=(6, 16))

        dir_subframe = ctk.CTkFrame(settings_card, fg_color="transparent")
        dir_subframe.grid(
            row=4, column=1, columnspan=2, sticky="ew", padx=(0, 20), pady=(6, 16)
        )
        dir_subframe.grid_columnconfigure(0, weight=1)

        self.dir_entry = ctk.CTkEntry(
            dir_subframe,
            textvariable=self.output_dir_var,
            height=34,
            corner_radius=8,
            fg_color=COLOR_INPUT_BG,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            font=font_body,
        )
        self.dir_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.browse_btn = ctk.CTkButton(
            dir_subframe,
            text="参照...",
            width=96,
            height=34,
            corner_radius=8,
            font=font_body,
            fg_color=COLOR_SECONDARY_BG,
            hover_color=COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            command=self.browse_output_dir,
        )
        self.browse_btn.grid(row=0, column=1)

        # --- 2. リアルタイムKPIカード（横並び3分割のミニ統計カード） ---
        kpi_container = ctk.CTkFrame(self, fg_color="transparent")
        kpi_container.grid(row=2, column=0, sticky="ew", padx=24, pady=6)
        kpi_container.grid_columnconfigure((0, 1, 2), weight=1, uniform="kpi")

        # KPI 1: 総取得数
        kpi_card_total = ctk.CTkFrame(
            kpi_container,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        kpi_card_total.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ctk.CTkLabel(
            kpi_card_total,
            text="総取得数",
            font=font_kpi_lbl,
            text_color=COLOR_TEXT_SUB,
            anchor="w",
        ).pack(anchor="w", padx=18, pady=(12, 2))
        ctk.CTkLabel(
            kpi_card_total,
            textvariable=self.kpi_total_var,
            font=font_kpi_val,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        ).pack(anchor="w", padx=18, pady=(0, 12))

        # KPI 2: 仕入れ候補
        kpi_card_cand = ctk.CTkFrame(
            kpi_container,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        kpi_card_cand.grid(row=0, column=1, sticky="ew", padx=6)
        ctk.CTkLabel(
            kpi_card_cand,
            text="仕入れ候補",
            font=font_kpi_lbl,
            text_color=COLOR_TEXT_SUB,
            anchor="w",
        ).pack(anchor="w", padx=18, pady=(12, 2))
        ctk.CTkLabel(
            kpi_card_cand,
            textvariable=self.kpi_candidates_var,
            font=font_kpi_val,
            text_color=COLOR_ACCENT,
            anchor="w",
        ).pack(anchor="w", padx=18, pady=(0, 12))

        # KPI 3: ステータス
        kpi_card_status = ctk.CTkFrame(
            kpi_container,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        kpi_card_status.grid(row=0, column=2, sticky="ew", padx=(6, 0))
        ctk.CTkLabel(
            kpi_card_status,
            text="ステータス",
            font=font_kpi_lbl,
            text_color=COLOR_TEXT_SUB,
            anchor="w",
        ).pack(anchor="w", padx=18, pady=(12, 2))
        self.kpi_status_label = ctk.CTkLabel(
            kpi_card_status,
            textvariable=self.kpi_status_var,
            font=font_kpi_val,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        )
        self.kpi_status_label.pack(anchor="w", padx=18, pady=(0, 12))

        # --- 3. 操作ボタン＆進捗状況カード ---
        progress_card = ctk.CTkFrame(
            self,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        progress_card.grid(row=3, column=0, sticky="ew", padx=24, pady=6)
        progress_card.grid_columnconfigure(0, weight=1)

        action_frame = ctk.CTkFrame(progress_card, fg_color="transparent")
        action_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(14, 10))
        action_frame.grid_columnconfigure(2, weight=1)

        # 収集開始ボタン（エメラルドグリーン）
        self.start_btn = ctk.CTkButton(
            action_frame,
            text="収集開始",
            width=145,
            height=38,
            corner_radius=8,
            font=font_body_bold,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="#FFFFFF",
            command=self.start_scraping,
        )
        self.start_btn.grid(row=0, column=0, padx=(0, 10))

        # 中断ボタン（スマートなゴーストボタン）
        self.stop_btn = ctk.CTkButton(
            action_frame,
            text="中断",
            width=115,
            height=38,
            corner_radius=8,
            font=font_body,
            fg_color="transparent",
            border_width=1,
            border_color=COLOR_GHOST_BORDER,
            hover_color=COLOR_GHOST_HOVER,
            text_color=COLOR_TEXT_MAIN,
            state="disabled",
            command=self.stop_scraping,
        )
        self.stop_btn.grid(row=0, column=1, padx=(0, 10))

        # 出力先フォルダを開くボタン
        self.open_dir_btn = ctk.CTkButton(
            action_frame,
            text="出力先フォルダを開く",
            width=175,
            height=38,
            corner_radius=8,
            font=font_body,
            fg_color=COLOR_SECONDARY_BG,
            hover_color=COLOR_SECONDARY_HOVER,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            command=self.open_output_dir,
        )
        self.open_dir_btn.grid(row=0, column=3, sticky="e")

        # 進捗詳細ラベル＆プログレスバー
        status_header = ctk.CTkFrame(progress_card, fg_color="transparent")
        status_header.grid(row=1, column=0, sticky="ew", padx=20, pady=(2, 4))
        status_header.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            status_header,
            text="進行状況:",
            font=font_section,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.progress_detail_label = ctk.CTkLabel(
            status_header,
            textvariable=self.progress_detail_var,
            font=font_body,
            text_color=COLOR_TEXT_SUB,
            anchor="w",
        )
        self.progress_detail_label.grid(row=0, column=1, sticky="w")

        self.progress_bar = ctk.CTkProgressBar(
            progress_card,
            height=10,
            corner_radius=5,
            fg_color=COLOR_INPUT_BG,
            progress_color=COLOR_ACCENT,
        )
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=20, pady=(4, 16))
        self.progress_bar.set(0.0)

        # --- 4. 実行ログカード ---
        log_card = ctk.CTkFrame(
            self,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        log_card.grid(row=4, column=0, sticky="nsew", padx=24, pady=(6, 20))
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_card,
            text="実行ログ",
            font=font_section,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(12, 6))

        self.log_text = ctk.CTkTextbox(
            log_card,
            font=font_body,
            corner_radius=8,
            fg_color=COLOR_INPUT_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            wrap="word",
            state="disabled",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 16))

    def browse_output_dir(self) -> None:
        """出力先フォルダ選択ダイアログを開き、選択されたパスを保存する。"""
        current_dir = self.output_dir_var.get().strip() or DEFAULT_OUTPUT_DIR
        initial_dir = current_dir if os.path.exists(current_dir) else get_base_dir()
        selected = filedialog.askdirectory(initialdir=initial_dir, title="出力先フォルダを選択")
        if selected:
            self.output_dir_var.set(selected)
            self.save_config()
            self.append_log(f"出力先フォルダを変更しました: {selected}")

    def open_output_dir(self) -> None:
        """出力先フォルダをエクスプローラーで開く。存在しない場合は作成して開く。"""
        out_dir = self.output_dir_var.get().strip() or DEFAULT_OUTPUT_DIR
        try:
            os.makedirs(out_dir, exist_ok=True)
            os.startfile(out_dir)
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダを開けませんでした:\n{e}")

    def append_log(self, message: str) -> None:
        """メインスレッド上で安全に実行ログエリアへメッセージを追記する。"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"

        def _append() -> None:
            self.log_text.configure(state="normal")
            self.log_text.insert(tk.END, line)
            self.log_text.see(tk.END)
            self.log_text.configure(state="disabled")

        self.after(0, _append)

    def update_kpi(
        self,
        total_count: int | None = None,
        candidate_count: int | None = None,
        status_text: str | None = None,
    ) -> None:
        """メインスレッド上で3つのKPIカード（総取得数・仕入れ候補・ステータス）を更新する。"""
        def _update() -> None:
            if total_count is not None:
                self.kpi_total_var.set(f"{total_count} 件")
            if candidate_count is not None:
                self.kpi_candidates_var.set(f"{candidate_count} 件")
            if status_text is not None:
                self.kpi_status_var.set(status_text)

        self.after(0, _update)

    def validate_inputs(self) -> tuple[int, int, float, str] | None:
        """画面入力値を検証して (max_pages, min_rating, max_price, output_dir) を返す。"""
        try:
            max_pages = int(self.max_pages_var.get().strip())
            if not (1 <= max_pages <= 50):
                raise ValueError
        except ValueError:
            messagebox.showwarning("入力エラー", "取得ページ数上限は 1 〜 50 の整数で入力してください。")
            return None

        try:
            min_rating = int(self.min_rating_var.get().strip())
            if not (1 <= min_rating <= 5):
                raise ValueError
        except ValueError:
            messagebox.showwarning("入力エラー", "最低評価は 1 〜 5 の整数で指定してください。")
            return None

        try:
            max_price = float(self.max_price_var.get().strip())
            if max_price < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("入力エラー", "上限価格は 0 以上の数値（例: 30.0）で入力してください。")
            return None

        output_dir = self.output_dir_var.get().strip() or DEFAULT_OUTPUT_DIR
        return max_pages, min_rating, max_price, output_dir

    def _set_running_ui(self, running: bool) -> None:
        """収集実行中かどうかに応じて「収集開始」「中断」ボタンの見た目を切り替える。"""
        if running:
            self.start_btn.configure(state="disabled", fg_color=COLOR_START_DISABLED)
            self.stop_btn.configure(
                state="normal",
                text="中断",
                fg_color=COLOR_DANGER,
                hover_color=COLOR_DANGER_HOVER,
                border_color=COLOR_DANGER,
                text_color="#FFFFFF",
            )
        else:
            self.start_btn.configure(state="normal", fg_color=COLOR_ACCENT)
            self.stop_btn.configure(
                state="disabled",
                text="中断",
                fg_color="transparent",
                hover_color=COLOR_GHOST_HOVER,
                border_color=COLOR_GHOST_BORDER,
                text_color=COLOR_TEXT_MAIN,
            )

    def start_scraping(self) -> None:
        """「収集開始」ボタン押下時の処理。バックグラウンドスレッドで収集を開始する。"""
        validated = self.validate_inputs()
        if validated is None:
            return

        max_pages, min_rating, max_price, output_dir = validated
        self.save_config()

        # UI・KPI状態の初期化
        self.stop_event.clear()
        self._set_running_ui(True)
        self.progress_bar.set(0.0)
        self.update_kpi(total_count=0, candidate_count=0, status_text="収集中")
        self.progress_detail_var.set("データ収集を開始しています...")

        self.append_log(
            f"=== 収集開始: 最大{max_pages}ページ / 条件(評価>={min_rating}, 価格<={max_price}, In stock) ==="
        )

        self.worker_thread = threading.Thread(
            target=self._run_scrape_worker,
            args=(max_pages, min_rating, max_price, output_dir),
            daemon=True,
        )
        self.worker_thread.start()

    def stop_scraping(self) -> None:
        """「中断」ボタン押下時の処理。安全にループを抜けるフラグを立てる。"""
        if self.worker_thread and self.worker_thread.is_alive():
            self.stop_event.set()
            self.stop_btn.configure(
                state="disabled",
                text="中断中...",
                fg_color="transparent",
                border_color=COLOR_GHOST_BORDER,
            )
            self.update_kpi(status_text="中断中")
            self.progress_detail_var.set("中断処理中... 現在のページ処理完了を待っています")
            self.append_log("中断リクエストを受け付けました。安全に停止します...")

    def _run_scrape_worker(
        self, max_pages: int, min_rating: int, max_price: float, output_dir: str
    ) -> None:
        """バックグラウンドスレッドでスクレイピングとKPIリアルタイム更新・Excel出力を実行する。"""
        all_books: list[dict] = []
        total_candidates = 0

        try:
            for page in range(1, max_pages + 1):
                if self.stop_event.is_set():
                    self.append_log("[中断] ユーザー操作によりデータ収集を中断しました。")
                    break

                url = PAGE_URL_TEMPLATE.format(page=page)
                self.append_log(f"{page} / {max_pages} ページ目を収集中... ({url})")

                def _on_page_start(p: int = page) -> None:
                    self.progress_detail_var.set(f"{p} / {max_pages} ページ目を収集中...")
                    self.progress_bar.set((p - 1) / max_pages)

                self.after(0, _on_page_start)

                html = fetch_page(url, log_callback=self.append_log)
                if html is None:
                    self.append_log(
                        f"[停止] {page}ページ目の取得に失敗したため、処理を安全に停止します。"
                    )
                    break

                page_books = parse_books_from_html(html, page_url=url)
                all_books.extend(page_books)

                # 仕入れ候補件数をリアルタイム集計してKPIカードを更新
                page_cand = sum(
                    1
                    for b in page_books
                    if b["rating"] >= min_rating
                    and b["price"] <= max_price
                    and "In stock" in str(b["stock"])
                )
                total_candidates += page_cand
                total_fetched = len(all_books)

                def _on_page_done(
                    p: int = page, tot: int = total_fetched, cand: int = total_candidates
                ) -> None:
                    self.progress_bar.set(p / max_pages)
                    self.kpi_total_var.set(f"{tot} 件")
                    self.kpi_candidates_var.set(f"{cand} 件")

                self.after(0, _on_page_done)

                # 1ページ取得ごとに必ず1.0秒待機
                time.sleep(1.0)

            columns = ["title", "price", "stock", "rating", "url"]
            df = pd.DataFrame(all_books, columns=columns)
            was_stopped = self.stop_event.is_set()

            if not df.empty:
                candidates_df = filter_candidates(
                    df, min_rating=min_rating, max_price=max_price
                )
                saved_path = export_to_excel(
                    df,
                    output_dir=output_dir,
                    min_rating=min_rating,
                    max_price=max_price,
                )
                self.append_log(
                    f"取得件数: 全 {len(df)} 件 / 仕入れ候補: {len(candidates_df)} 件"
                )
                self.append_log(f"Excel出力完了: {saved_path}")

                def _on_done() -> None:
                    self.kpi_total_var.set(f"{len(df)} 件")
                    self.kpi_candidates_var.set(f"{len(candidates_df)} 件")
                    if was_stopped:
                        self.kpi_status_var.set("中断")
                        self.progress_detail_var.set(
                            f"中断完了（取得済み {len(df)} 件をExcelに出力しました）"
                        )
                    else:
                        self.progress_bar.set(1.0)
                        self.kpi_status_var.set("完了")
                        self.progress_detail_var.set(
                            f"完了！ 全 {len(df)} 件（仕入れ候補 {len(candidates_df)} 件）を出力しました"
                        )

                self.after(0, _on_done)
            else:
                self.append_log("[警告] 取得されたデータが 0 件のため、Excel出力をスキップしました。")

                def _on_empty() -> None:
                    if was_stopped:
                        self.kpi_status_var.set("中断")
                        self.progress_detail_var.set("中断完了（取得データ 0 件）")
                    else:
                        self.kpi_status_var.set("停止")
                        self.progress_detail_var.set("エラーまたはデータなしのため終了しました")

                self.after(0, _on_empty)

        except Exception as e:
            self.append_log(f"[エラー] 予期しない例外が発生しました: {e}")

            def _on_error() -> None:
                self.kpi_status_var.set("エラー")
                self.progress_detail_var.set("エラーにより停止しました")

            self.after(0, _on_error)
        finally:
            def _restore_buttons() -> None:
                self._set_running_ui(False)

            self.after(0, _restore_buttons)


def main() -> None:
    app = BookScraperApp()
    app.mainloop()


if __name__ == "__main__":
    main()

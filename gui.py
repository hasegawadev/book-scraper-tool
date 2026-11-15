from __future__ import annotations

from datetime import datetime
import json
import os
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import pandas as pd

from scraper import (
    export_to_excel,
    filter_candidates,
    scrape_books,
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
COLOR_WARNING = "#F59E0B"
COLOR_START_DISABLED = ("#D1D5DB", "#2E2E2E")
COLOR_SECONDARY_BG = ("#E5E7EB", "#262626")
COLOR_SECONDARY_HOVER = ("#D1D5DB", "#333333")
COLOR_TEXT_MAIN = ("#111827", "#F9FAFB")
COLOR_TEXT_SUB = ("#6B7280", "#9CA3AF")

FONT_FAMILY = "Yu Gothic UI"
LOG_PLACEHOLDER_TEXT = "ここに実行ログが表示されます..."


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
        self.geometry("920x860")
        self.minsize(780, 740)

        # 状態管理フラグ
        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.has_log_entries = False
        self.sort_reverse_map: dict[str, bool] = {}

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
        """config.json から前回の設定値（ページ数・最低評価・上限価格・出力先フォルダ）を復元する。"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if not isinstance(data, dict):
                    return
                if "max_pages" in data and data["max_pages"] is not None:
                    self.max_pages_var.set(str(data["max_pages"]))
                if "min_rating" in data and data["min_rating"] is not None:
                    self.min_rating_var.set(str(data["min_rating"]))
                if "max_price" in data and data["max_price"] is not None:
                    self.max_price_var.set(str(data["max_price"]))
                saved_dir = data.get("output_dir")
                if saved_dir and isinstance(saved_dir, str):
                    self.output_dir_var.set(saved_dir)
            except Exception:
                pass

    def save_config(self) -> None:
        """現在の設定値（取得ページ数・最低評価・上限価格・出力先フォルダ）を config.json に保存する。"""
        data = {
            "max_pages": self.max_pages_var.get().strip() or "5",
            "min_rating": self.min_rating_var.get().strip() or "4",
            "max_price": self.max_price_var.get().strip() or "30.0",
            "output_dir": self.output_dir_var.get().strip() or DEFAULT_OUTPUT_DIR,
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.append_log(f"[警告] 設定ファイルの保存に失敗しました: {e}")

    def toggle_appearance_mode(self) -> None:
        """ダークモード／ライトモードを切り替え、Treeviewやログ色も連動させる。"""
        is_dark = self.dark_mode_var.get()
        if is_dark:
            ctk.set_appearance_mode("dark")
            self.theme_switch.configure(text="ダークモード")
        else:
            ctk.set_appearance_mode("light")
            self.theme_switch.configure(text="ライトモード")
        self._apply_treeview_theme(is_dark)
        self._configure_log_tags(is_dark)

    def _apply_treeview_theme(self, is_dark: bool) -> None:
        """ダーク／ライトテーマに合わせて ttk.Treeview の配色を調整する。"""
        style = ttk.Style(self)
        style.theme_use("clam")

        if is_dark:
            bg_color = "#141414"
            fg_color = "#F9FAFB"
            header_bg = "#262626"
            header_fg = "#F9FAFB"
            border_color = "#2E2E2E"
        else:
            bg_color = "#F9FAFB"
            fg_color = "#111827"
            header_bg = "#E5E7EB"
            header_fg = "#111827"
            border_color = "#D1D5DB"

        style.configure(
            "Preview.Treeview",
            background=bg_color,
            foreground=fg_color,
            fieldbackground=bg_color,
            bordercolor=border_color,
            lightcolor=border_color,
            darkcolor=border_color,
            rowheight=26,
            font=(FONT_FAMILY, 10),
        )
        style.map(
            "Preview.Treeview",
            background=[("selected", COLOR_ACCENT)],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            "Preview.Treeview.Heading",
            background=header_bg,
            foreground=header_fg,
            bordercolor=border_color,
            relief="flat",
            font=(FONT_FAMILY, 10, "bold"),
            padding=(6, 5),
        )
        style.map(
            "Preview.Treeview.Heading",
            background=[("active", COLOR_ACCENT)],
            foreground=[("active", "#FFFFFF")],
        )

    def _configure_log_tags(self, is_dark: bool) -> None:
        """CTkTextbox 内部の tk.Text タグによるログ色分けを設定する。"""
        text_widget: tk.Text = self.log_text._textbox
        normal_fg = "#F9FAFB" if is_dark else "#111827"
        placeholder_fg = "#6B7280" if is_dark else "#9CA3AF"
        text_widget.tag_config("normal", foreground=normal_fg)
        text_widget.tag_config("success", foreground=COLOR_ACCENT)
        text_widget.tag_config("warning", foreground=COLOR_WARNING)
        text_widget.tag_config("error", foreground=COLOR_DANGER)
        text_widget.tag_config("placeholder", foreground=placeholder_fg)

    def _show_log_placeholder(self) -> None:
        """初期状態として実行ログにプレースホルダーを表示する。"""
        self.has_log_entries = False
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text._textbox.insert(tk.END, LOG_PLACEHOLDER_TEXT, "placeholder")
        self.log_text.configure(state="disabled")

    def _build_ui(self) -> None:
        """モダンSaaSダッシュボード風のUIレイアウトを構築する。"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)
        self.grid_rowconfigure(5, weight=1)

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
        header_frame.grid(row=0, column=0, sticky="ew", padx=24, pady=(14, 6))
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
        settings_card.grid(row=1, column=0, sticky="ew", padx=24, pady=5)
        settings_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            settings_card,
            text="収集・抽出条件設定",
            font=font_section,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=20, pady=(12, 6))

        # 取得ページ数上限
        ctk.CTkLabel(
            settings_card,
            text="取得ページ数上限 (1〜50):",
            font=font_body,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=1, column=0, sticky="w", padx=(20, 14), pady=4)

        self.pages_entry = ctk.CTkEntry(
            settings_card,
            textvariable=self.max_pages_var,
            width=110,
            height=32,
            corner_radius=8,
            border_width=1,
            fg_color=COLOR_INPUT_BG,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            font=font_body,
        )
        self.pages_entry.grid(row=1, column=1, sticky="w", pady=4)

        ctk.CTkLabel(
            settings_card,
            text="※全件取得時は 50 を指定（1ページ20件）",
            font=font_sub,
            text_color=COLOR_TEXT_SUB,
        ).grid(row=1, column=2, sticky="w", padx=(12, 20), pady=4)

        # 仕入れ候補の最低評価
        ctk.CTkLabel(
            settings_card,
            text="仕入れ候補の最低評価 (1〜5):",
            font=font_body,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=2, column=0, sticky="w", padx=(20, 14), pady=4)

        self.rating_menu = ctk.CTkOptionMenu(
            settings_card,
            variable=self.min_rating_var,
            values=["1", "2", "3", "4", "5"],
            width=110,
            height=32,
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
        self.rating_menu.grid(row=2, column=1, sticky="w", pady=4)

        ctk.CTkLabel(
            settings_card,
            text="以上",
            font=font_sub,
            text_color=COLOR_TEXT_SUB,
        ).grid(row=2, column=2, sticky="w", padx=(12, 20), pady=4)

        # 仕入れ候補の上限価格
        ctk.CTkLabel(
            settings_card,
            text="仕入れ候補の上限価格 (£):",
            font=font_body,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=3, column=0, sticky="w", padx=(20, 14), pady=4)

        self.price_entry = ctk.CTkEntry(
            settings_card,
            textvariable=self.max_price_var,
            width=110,
            height=32,
            corner_radius=8,
            border_width=1,
            fg_color=COLOR_INPUT_BG,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            font=font_body,
        )
        self.price_entry.grid(row=3, column=1, sticky="w", pady=4)

        ctk.CTkLabel(
            settings_card,
            text="以下（かつ在庫「In stock」を含む）",
            font=font_sub,
            text_color=COLOR_TEXT_SUB,
        ).grid(row=3, column=2, sticky="w", padx=(12, 20), pady=4)

        # 出力先フォルダ（横幅いっぱいに広げてパス切れを防止）
        ctk.CTkLabel(
            settings_card,
            text="出力先フォルダ:",
            font=font_body,
            text_color=COLOR_TEXT_MAIN,
        ).grid(row=4, column=0, sticky="w", padx=(20, 14), pady=(4, 14))

        dir_subframe = ctk.CTkFrame(settings_card, fg_color="transparent")
        dir_subframe.grid(
            row=4, column=1, columnspan=2, sticky="ew", padx=(0, 20), pady=(4, 14)
        )
        dir_subframe.grid_columnconfigure(0, weight=1)

        self.dir_entry = ctk.CTkEntry(
            dir_subframe,
            textvariable=self.output_dir_var,
            height=32,
            corner_radius=8,
            border_width=1,
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
            height=32,
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
        kpi_container.grid(row=2, column=0, sticky="ew", padx=24, pady=5)
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
        ).pack(anchor="w", padx=18, pady=(10, 2))
        ctk.CTkLabel(
            kpi_card_total,
            textvariable=self.kpi_total_var,
            font=font_kpi_val,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        ).pack(anchor="w", padx=18, pady=(0, 10))

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
        ).pack(anchor="w", padx=18, pady=(10, 2))
        ctk.CTkLabel(
            kpi_card_cand,
            textvariable=self.kpi_candidates_var,
            font=font_kpi_val,
            text_color=COLOR_ACCENT,
            anchor="w",
        ).pack(anchor="w", padx=18, pady=(0, 10))

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
        ).pack(anchor="w", padx=18, pady=(10, 2))
        self.kpi_status_label = ctk.CTkLabel(
            kpi_card_status,
            textvariable=self.kpi_status_var,
            font=font_kpi_val,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        )
        self.kpi_status_label.pack(anchor="w", padx=18, pady=(0, 10))

        # --- 3. 操作ボタン＆進捗状況カード ---
        progress_card = ctk.CTkFrame(
            self,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        progress_card.grid(row=3, column=0, sticky="ew", padx=24, pady=5)
        progress_card.grid_columnconfigure(0, weight=1)

        action_frame = ctk.CTkFrame(progress_card, fg_color="transparent")
        action_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(12, 8))
        action_frame.grid_columnconfigure(2, weight=1)

        # 収集開始ボタン（エメラルドグリーン）
        self.start_btn = ctk.CTkButton(
            action_frame,
            text="収集開始",
            width=145,
            height=36,
            corner_radius=8,
            font=font_body_bold,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="#FFFFFF",
            command=self.start_scraping,
        )
        self.start_btn.grid(row=0, column=0, padx=(0, 10))

        # 中断ボタン（スマートなゴーストボタン／実行時は赤色強調）
        self.stop_btn = ctk.CTkButton(
            action_frame,
            text="中断",
            width=115,
            height=36,
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
            height=36,
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
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=20, pady=(4, 14))
        self.progress_bar.set(0.0)

        # --- 4. 実行ログカード ---
        log_card = ctk.CTkFrame(
            self,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        log_card.grid(row=4, column=0, sticky="nsew", padx=24, pady=5)
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_card,
            text="実行ログ",
            font=font_section,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(10, 4))

        self.log_text = ctk.CTkTextbox(
            log_card,
            height=115,
            font=font_body,
            corner_radius=8,
            fg_color=COLOR_INPUT_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            text_color=COLOR_TEXT_MAIN,
            wrap="word",
            state="disabled",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 12))
        self._configure_log_tags(is_dark=self.dark_mode_var.get())
        self._show_log_placeholder()

        # --- 5. 仕入れ候補プレビューカード（ttk.Treeview） ---
        preview_card = ctk.CTkFrame(
            self,
            fg_color=COLOR_CARD_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=12,
        )
        preview_card.grid(row=5, column=0, sticky="nsew", padx=24, pady=(5, 16))
        preview_card.grid_columnconfigure(0, weight=1)
        preview_card.grid_rowconfigure(1, weight=1)

        preview_header = ctk.CTkFrame(preview_card, fg_color="transparent")
        preview_header.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 4))
        preview_header.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            preview_header,
            text="仕入れ候補プレビュー",
            font=font_section,
            text_color=COLOR_TEXT_MAIN,
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            preview_header,
            text="※列ヘッダーをクリックすると昇順・降順でソートできます",
            font=font_sub,
            text_color=COLOR_TEXT_SUB,
            anchor="e",
        ).grid(row=0, column=1, sticky="e")

        table_container = ctk.CTkFrame(
            preview_card,
            fg_color=COLOR_INPUT_BG,
            border_width=1,
            border_color=COLOR_CARD_BORDER,
            corner_radius=8,
        )
        table_container.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 12))
        table_container.grid_columnconfigure(0, weight=1)
        table_container.grid_rowconfigure(0, weight=1)

        self._apply_treeview_theme(is_dark=self.dark_mode_var.get())

        columns = ("status_diff", "title", "price", "rating", "stock")
        self.preview_tree = ttk.Treeview(
            table_container,
            columns=columns,
            show="headings",
            style="Preview.Treeview",
            height=6,
        )

        col_configs = [
            ("status_diff", "差分状況", 140, "center"),
            ("title", "商品名", 350, "w"),
            ("price", "価格(£)", 95, "e"),
            ("rating", "評価", 75, "center"),
            ("stock", "在庫状況", 110, "center"),
        ]
        for col_id, heading_text, width, anchor in col_configs:
            self.preview_tree.heading(
                col_id,
                text=heading_text,
                command=lambda c=col_id: self.sort_preview_column(c),
            )
            self.preview_tree.column(
                col_id, width=width, minwidth=60, anchor=anchor, stretch=(col_id == "title")
            )

        tree_scroll = ttk.Scrollbar(
            table_container, orient=tk.VERTICAL, command=self.preview_tree.yview
        )
        self.preview_tree.configure(yscrollcommand=tree_scroll.set)

        self.preview_tree.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=2)
        tree_scroll.grid(row=0, column=1, sticky="ns", padx=(0, 2), pady=2)

    def sort_preview_column(self, col: str) -> None:
        """プレビュー表の指定列をクリックした際に昇順・降順でソートする。"""
        items = list(self.preview_tree.get_children(""))
        if not items:
            return

        reverse = self.sort_reverse_map.get(col, False)

        def _sort_key(item_id: str):
            val = self.preview_tree.set(item_id, col)
            if col == "price":
                m = re.search(r"[-+]?\d+(?:\.\d+)?", str(val))
                return float(m.group(0)) if m else 0.0
            if col == "rating":
                m = re.search(r"\d+", str(val))
                return int(m.group(0)) if m else 0
            if col == "status_diff":
                # 値下がり等の数値差額がある場合は差額順、それ以外は文字列順
                m = re.search(r"[-+]\s*£(\d+(?:\.\d+)?)", str(val))
                if m:
                    sign = -1.0 if "-" in str(val) else 1.0
                    return (0, sign * float(m.group(1)), str(val))
                return (1, 0.0, str(val))
            return str(val).lower()

        items.sort(key=_sort_key, reverse=reverse)
        for index, item_id in enumerate(items):
            self.preview_tree.move(item_id, "", index)

        self.sort_reverse_map[col] = not reverse

        # ヘッダーにソート方向インジケーター（▲/▼）を表示
        base_titles = {
            "status_diff": "差分状況",
            "title": "商品名",
            "price": "価格(£)",
            "rating": "評価",
            "stock": "在庫状況",
        }
        for c_id, title in base_titles.items():
            if c_id == col:
                arrow = " ▼" if reverse else " ▲"
                self.preview_tree.heading(c_id, text=f"{title}{arrow}")
            else:
                self.preview_tree.heading(c_id, text=title)

    def update_preview_table(self, candidates_df: pd.DataFrame) -> None:
        """抽出された仕入れ候補DataFrameをプレビュー表（Treeview）に反映する。"""
        for item_id in self.preview_tree.get_children(""):
            self.preview_tree.delete(item_id)

        if candidates_df.empty:
            return

        for _, row in candidates_df.iterrows():
            status_diff = str(row.get("status_diff", "初回取得"))
            title = str(row.get("title", ""))
            price_val = float(row.get("price", 0.0))
            rating_val = int(row.get("rating", 0))
            stock = str(row.get("stock", ""))

            self.preview_tree.insert(
                "",
                tk.END,
                values=(
                    status_diff,
                    title,
                    f"£{price_val:.2f}",
                    f"★{rating_val} ({rating_val})",
                    stock,
                ),
            )

    def browse_output_dir(self) -> None:
        """出力先フォルダ選択ダイアログを開き、選択されたパスを保存する。"""
        current_dir = self.output_dir_var.get().strip() or DEFAULT_OUTPUT_DIR
        initial_dir = current_dir if os.path.exists(current_dir) else get_base_dir()
        selected = filedialog.askdirectory(initialdir=initial_dir, title="出力先フォルダを選択")
        if selected:
            self.output_dir_var.set(selected)
            self.dir_entry.configure(border_color=COLOR_CARD_BORDER)
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

    def _classify_log_tag(self, message: str) -> str:
        """ログメッセージの内容に応じて色分けタグ（error / warning / success / normal）を判定する。"""
        if any(k in message for k in ("[エラー]", "[停止]", "失敗", "エラー")):
            return "error"
        if any(k in message for k in ("[警告]", "[中断]", "中断リクエスト", "スキップ")):
            return "warning"
        if any(k in message for k in ("[完了]", "Excel出力完了", "値下がり", "=== 収集開始")):
            return "success"
        return "normal"

    def append_log(self, message: str) -> None:
        """メインスレッド上で安全に実行ログエリアへ色分け付きメッセージを追記する。"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        tag = self._classify_log_tag(message)

        def _append() -> None:
            self.log_text.configure(state="normal")
            if not self.has_log_entries:
                self.log_text.delete("1.0", tk.END)
                self.has_log_entries = True
            self.log_text._textbox.insert(tk.END, line, tag)
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

    def _reset_input_borders(self) -> None:
        """入力欄の枠線色を通常のボーダー色に戻す。"""
        self.pages_entry.configure(border_color=COLOR_CARD_BORDER)
        self.price_entry.configure(border_color=COLOR_CARD_BORDER)
        self.dir_entry.configure(border_color=COLOR_CARD_BORDER)

    def validate_inputs(self) -> tuple[int, int, float, str] | None:
        """
        画面入力値を検証する。
        不正な値がある場合はポップアップを出さず、該当入力欄の枠線を赤く（#EF4444）し、
        ステータス欄と進行状況欄にエラー理由を表示する。
        """
        self._reset_input_borders()
        errors: list[str] = []

        max_pages = 5
        try:
            max_pages = int(self.max_pages_var.get().strip())
            if not (1 <= max_pages <= 50):
                raise ValueError
        except ValueError:
            self.pages_entry.configure(border_color=COLOR_DANGER)
            errors.append("取得ページ数上限は 1〜50 の整数で入力してください")

        min_rating = 4
        try:
            min_rating = int(self.min_rating_var.get().strip())
            if not (1 <= min_rating <= 5):
                raise ValueError
        except ValueError:
            errors.append("最低評価は 1〜5 の整数で指定してください")

        max_price = 30.0
        try:
            max_price = float(self.max_price_var.get().strip())
            if max_price < 0:
                raise ValueError
        except ValueError:
            self.price_entry.configure(border_color=COLOR_DANGER)
            errors.append("上限価格は 0 以上の数値（例: 30.0）で入力してください")

        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            self.dir_entry.configure(border_color=COLOR_DANGER)
            errors.append("出力先フォルダパスを入力してください")

        if errors:
            error_msg = "入力エラー: " + " / ".join(errors)
            self.kpi_status_var.set("入力エラー")
            self.progress_detail_var.set(error_msg)
            self.progress_detail_label.configure(text_color=COLOR_DANGER)
            self.append_log(f"[エラー] {error_msg}")
            return None

        self.progress_detail_label.configure(text_color=COLOR_TEXT_SUB)
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
        self.progress_detail_var.set(f"1 / {max_pages} ページ目 取得中...（残り時間を計算中）")

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
        """バックグラウンドスレッドで scraper.scrape_books を呼び出し、進捗・残り時間更新・プレビュー反映・Excel出力を行う。"""
        start_time = time.perf_counter()

        def _on_progress(
            completed_pages: int, total_pages: int, message: str, current_books: list[dict]
        ) -> None:
            elapsed = time.perf_counter() - start_time
            total_fetched = len(current_books)
            total_candidates = sum(
                1
                for b in current_books
                if b["rating"] >= min_rating
                and b["price"] <= max_price
                and "In stock" in str(b["stock"])
            )

            # 処理経過時間から1ページあたりの平均時間を算出し、残り秒数の目安を表示
            if completed_pages > 0:
                avg_sec_per_page = elapsed / completed_pages
                remaining_pages = max(0, total_pages - completed_pages)
                rem_sec = max(0, int(round(avg_sec_per_page * remaining_pages)))
                current_disp = min(completed_pages + 1, total_pages)
                if completed_pages < total_pages:
                    detail_text = (
                        f"{current_disp} / {total_pages} ページ 取得中…（残り約 {rem_sec} 秒）"
                    )
                else:
                    detail_text = f"{total_pages} / {total_pages} ページ 取得完了（残り約 0 秒）"
            else:
                # 1ページ目取得開始時
                est_sec = int(round(total_pages * 1.5))
                detail_text = f"1 / {total_pages} ページ 取得中…（残り約 {est_sec} 秒）"

            def _update_ui() -> None:
                ratio = completed_pages / total_pages if total_pages > 0 else 0.0
                self.progress_bar.set(ratio)
                self.progress_detail_var.set(detail_text)
                self.kpi_total_var.set(f"{total_fetched} 件")
                self.kpi_candidates_var.set(f"{total_candidates} 件")

            self.after(0, _update_ui)

        try:
            df = scrape_books(
                max_pages=max_pages,
                sleep_sec=1.0,
                progress_callback=_on_progress,
                log_callback=self.append_log,
                stop_event=self.stop_event,
            )
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
                    f"[完了] 取得件数: 全 {len(df)} 件 / 仕入れ候補: {len(candidates_df)} 件"
                )
                self.append_log(f"Excel出力完了: {saved_path}")

                def _on_done() -> None:
                    self.kpi_total_var.set(f"{len(df)} 件")
                    self.kpi_candidates_var.set(f"{len(candidates_df)} 件")
                    self.update_preview_table(candidates_df)
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
                    self.update_preview_table(pd.DataFrame())
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

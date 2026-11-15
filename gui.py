from __future__ import annotations

from datetime import datetime
import json
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from scraper import export_to_excel, filter_candidates, scrape_books

# 全体テーマの初期設定
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def get_base_dir() -> str:
    """実行ファイル（.exe）またはスクリプトと同階層のディレクトリを取得する。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_FILE = os.path.join(get_base_dir(), "config.json")
DEFAULT_OUTPUT_DIR = os.path.join(get_base_dir(), "output")


class BookScraperApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("ECサイト商品データ自動収集＆条件抽出ツール")
        self.geometry("760x660")
        self.minsize(680, 580)

        # 状態管理フラグ
        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None

        # 設定変数の初期化
        self.max_pages_var = ctk.StringVar(value="5")
        self.min_rating_var = ctk.StringVar(value="4")
        self.max_price_var = ctk.StringVar(value="30.0")
        self.output_dir_var = ctk.StringVar(value=DEFAULT_OUTPUT_DIR)
        self.status_var = ctk.StringVar(value="待機中")
        self.dark_mode_var = ctk.BooleanVar(value=True)

        # config.json から前回設定を読み込み
        self.load_config()

        # モダンUI構築
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
        """角丸カード型フレーム（CTkFrame）を用いたモダンなGUIレイアウトを構築する。"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        font_title = ctk.CTkFont(family="Yu Gothic UI", size=20, weight="bold")
        font_section = ctk.CTkFont(family="Yu Gothic UI", size=14, weight="bold")
        font_body = ctk.CTkFont(family="Yu Gothic UI", size=13)
        font_sub = ctk.CTkFont(family="Yu Gothic UI", size=12)
        font_log = ctk.CTkFont(family="Consolas", size=12)

        # --- 0. ヘッダーエリア（タイトル＋右上ライト／ダークモード切替スイッチ） ---
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 8))
        header_frame.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            header_frame,
            text="ECサイト商品データ自動収集ツール",
            font=font_title,
            anchor="w",
        )
        title_label.grid(row=0, column=0, sticky="w")

        self.theme_switch = ctk.CTkSwitch(
            header_frame,
            text="ダークモード",
            variable=self.dark_mode_var,
            onvalue=True,
            offvalue=False,
            command=self.toggle_appearance_mode,
            font=font_sub,
        )
        self.theme_switch.grid(row=0, column=1, sticky="e")

        # --- 1. 設定項目カード（CTkFrame） ---
        settings_card = ctk.CTkFrame(self, corner_radius=12)
        settings_card.grid(row=1, column=0, sticky="ew", padx=20, pady=8)
        settings_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            settings_card,
            text="収集・抽出条件設定",
            font=font_section,
            anchor="w",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=18, pady=(14, 8))

        # 取得ページ数上限
        ctk.CTkLabel(
            settings_card, text="取得ページ数上限 (1〜50):", font=font_body
        ).grid(row=1, column=0, sticky="w", padx=(18, 12), pady=6)

        self.pages_entry = ctk.CTkEntry(
            settings_card, textvariable=self.max_pages_var, width=100, font=font_body
        )
        self.pages_entry.grid(row=1, column=1, sticky="w", pady=6)

        ctk.CTkLabel(
            settings_card,
            text="※全件取得時は 50 を指定（1ページ20件）",
            font=font_sub,
            text_color=("gray35", "gray70"),
        ).grid(row=1, column=2, sticky="w", padx=(10, 18), pady=6)

        # 仕入れ候補の最低評価
        ctk.CTkLabel(
            settings_card, text="仕入れ候補の最低評価 (1〜5):", font=font_body
        ).grid(row=2, column=0, sticky="w", padx=(18, 12), pady=6)

        self.rating_menu = ctk.CTkOptionMenu(
            settings_card,
            variable=self.min_rating_var,
            values=["1", "2", "3", "4", "5"],
            width=100,
            font=font_body,
        )
        self.rating_menu.grid(row=2, column=1, sticky="w", pady=6)

        ctk.CTkLabel(
            settings_card,
            text="以上",
            font=font_sub,
            text_color=("gray35", "gray70"),
        ).grid(row=2, column=2, sticky="w", padx=(10, 18), pady=6)

        # 仕入れ候補の上限価格
        ctk.CTkLabel(
            settings_card, text="仕入れ候補の上限価格 (£):", font=font_body
        ).grid(row=3, column=0, sticky="w", padx=(18, 12), pady=6)

        self.price_entry = ctk.CTkEntry(
            settings_card, textvariable=self.max_price_var, width=100, font=font_body
        )
        self.price_entry.grid(row=3, column=1, sticky="w", pady=6)

        ctk.CTkLabel(
            settings_card,
            text="以下（かつ在庫「In stock」を含む）",
            font=font_sub,
            text_color=("gray35", "gray70"),
        ).grid(row=3, column=2, sticky="w", padx=(10, 18), pady=6)

        # 出力先フォルダ選択
        ctk.CTkLabel(
            settings_card, text="出力先フォルダ:", font=font_body
        ).grid(row=4, column=0, sticky="w", padx=(18, 12), pady=(6, 16))

        dir_subframe = ctk.CTkFrame(settings_card, fg_color="transparent")
        dir_subframe.grid(row=4, column=1, columnspan=2, sticky="ew", padx=(0, 18), pady=(6, 16))
        dir_subframe.grid_columnconfigure(0, weight=1)

        self.dir_entry = ctk.CTkEntry(
            dir_subframe, textvariable=self.output_dir_var, font=font_body
        )
        self.dir_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.browse_btn = ctk.CTkButton(
            dir_subframe,
            text="参照...",
            width=90,
            font=font_body,
            fg_color=("gray75", "gray28"),
            hover_color=("gray65", "gray38"),
            text_color=("gray10", "gray95"),
            command=self.browse_output_dir,
        )
        self.browse_btn.grid(row=0, column=1)

        # --- 2. 進捗状況＆操作ボタンカード（CTkFrame） ---
        progress_card = ctk.CTkFrame(self, corner_radius=12)
        progress_card.grid(row=2, column=0, sticky="ew", padx=20, pady=8)
        progress_card.grid_columnconfigure(0, weight=1)

        # アクションボタン行
        action_frame = ctk.CTkFrame(progress_card, fg_color="transparent")
        action_frame.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 10))
        action_frame.grid_columnconfigure(2, weight=1)

        self.start_btn = ctk.CTkButton(
            action_frame,
            text="収集開始",
            width=140,
            height=36,
            font=ctk.CTkFont(family="Yu Gothic UI", size=13, weight="bold"),
            command=self.start_scraping,
        )
        self.start_btn.grid(row=0, column=0, padx=(0, 10))

        self.stop_btn = ctk.CTkButton(
            action_frame,
            text="中断",
            width=110,
            height=36,
            font=font_body,
            fg_color="#c94a4a",
            hover_color="#a83838",
            state="disabled",
            command=self.stop_scraping,
        )
        self.stop_btn.grid(row=0, column=1, padx=(0, 10))

        self.open_dir_btn = ctk.CTkButton(
            action_frame,
            text="出力先フォルダを開く",
            width=170,
            height=36,
            font=font_body,
            fg_color=("gray75", "gray28"),
            hover_color=("gray65", "gray38"),
            text_color=("gray10", "gray95"),
            command=self.open_output_dir,
        )
        self.open_dir_btn.grid(row=0, column=3, sticky="e")

        # 進捗ステータス＆プログレスバー
        status_header = ctk.CTkFrame(progress_card, fg_color="transparent")
        status_header.grid(row=1, column=0, sticky="ew", padx=18, pady=(2, 4))
        status_header.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            status_header, text="進捗状況:", font=font_section
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.status_label = ctk.CTkLabel(
            status_header, textvariable=self.status_var, font=font_body, anchor="w"
        )
        self.status_label.grid(row=0, column=1, sticky="w")

        self.progress_bar = ctk.CTkProgressBar(progress_card, height=12, corner_radius=6)
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=18, pady=(4, 16))
        self.progress_bar.set(0.0)

        # --- 3. 実行ログカード（CTkFrame） ---
        log_card = ctk.CTkFrame(self, corner_radius=12)
        log_card.grid(row=3, column=0, sticky="nsew", padx=20, pady=(8, 18))
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_card, text="実行ログ", font=font_section, anchor="w"
        ).grid(row=0, column=0, sticky="w", padx=18, pady=(12, 6))

        self.log_text = ctk.CTkTextbox(
            log_card,
            font=font_log,
            corner_radius=8,
            wrap="word",
            state="disabled",
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 16))

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

    def update_progress(self, current_page: int, max_pages: int) -> None:
        """メインスレッド上でプログレスバーとステータス表示を更新する。"""
        def _update() -> None:
            ratio = (current_page - 1) / max_pages if max_pages > 0 else 0.0
            self.progress_bar.set(ratio)
            self.status_var.set(f"{current_page} / {max_pages} ページ目を収集中...")

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

    def start_scraping(self) -> None:
        """「収集開始」ボタン押下時の処理。バックグラウンドスレッドで収集を開始する。"""
        validated = self.validate_inputs()
        if validated is None:
            return

        max_pages, min_rating, max_price, output_dir = validated
        self.save_config()

        # UI状態の更新
        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.progress_bar.set(0.0)
        self.status_var.set("データ収集を開始しています...")

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
            self.stop_btn.configure(state="disabled")
            self.status_var.set("中断処理中... 現在のページ処理完了を待っています")
            self.append_log("中断リクエストを受け付けました。安全に停止します...")

    def _run_scrape_worker(
        self, max_pages: int, min_rating: int, max_price: float, output_dir: str
    ) -> None:
        """バックグラウンドスレッドでスクレイピングとExcel出力を実行する。"""
        try:
            df = scrape_books(
                max_pages=max_pages,
                sleep_sec=1.0,
                progress_callback=self.update_progress,
                log_callback=self.append_log,
                stop_check=self.stop_event.is_set,
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
                    f"取得件数: 全 {len(df)} 件 / 仕入れ候補: {len(candidates_df)} 件"
                )
                self.append_log(f"Excel出力完了: {saved_path}")

                def _on_done() -> None:
                    if was_stopped:
                        self.status_var.set(
                            f"中断完了（取得済み {len(df)} 件をExcelに出力しました）"
                        )
                    else:
                        self.progress_bar.set(1.0)
                        self.status_var.set(
                            f"完了！ 全 {len(df)} 件（仕入れ候補 {len(candidates_df)} 件）を出力しました"
                        )

                self.after(0, _on_done)
            else:
                self.append_log("[警告] 取得されたデータが 0 件のため、Excel出力をスキップしました。")

                def _on_empty() -> None:
                    if was_stopped:
                        self.status_var.set("中断完了（取得データ 0 件）")
                    else:
                        self.status_var.set("エラーまたはデータなしのため終了しました")

                self.after(0, _on_empty)

        except Exception as e:
            self.append_log(f"[エラー] 予期しない例外が発生しました: {e}")
            self.after(0, lambda: self.status_var.set("エラーにより停止しました"))
        finally:
            def _restore_buttons() -> None:
                self.start_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")

            self.after(0, _restore_buttons)


def main() -> None:
    app = BookScraperApp()
    app.mainloop()


if __name__ == "__main__":
    main()

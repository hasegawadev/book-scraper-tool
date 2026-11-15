from __future__ import annotations

from datetime import datetime
import json
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from scraper import export_to_excel, filter_candidates, scrape_books


def get_base_dir() -> str:
    """実行ファイル（.exe）またはスクリプトと同階層のディレクトリを取得する。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_FILE = os.path.join(get_base_dir(), "config.json")
DEFAULT_OUTPUT_DIR = os.path.join(get_base_dir(), "output")


class BookScraperApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ECサイト商品データ自動収集＆条件抽出ツール")
        self.root.geometry("680x560")
        self.root.minsize(600, 480)

        # 状態管理フラグ
        self.stop_event = threading.Event()
        self.worker_thread: threading.Thread | None = None

        # 設定変数の初期化
        self.max_pages_var = tk.StringVar(value="5")
        self.min_rating_var = tk.StringVar(value="4")
        self.max_price_var = tk.StringVar(value="30.0")
        self.output_dir_var = tk.StringVar(value=DEFAULT_OUTPUT_DIR)
        self.status_var = tk.StringVar(value="待機中")

        # config.json から前回設定を読み込み
        self.load_config()

        # UI構築
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

    def _build_ui(self) -> None:
        """GUIレイアウトを構築する。"""
        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 1. 抽出・収集設定エリア ---
        settings_frame = ttk.LabelFrame(main_frame, text="収集・抽出条件設定", padding=10)
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        settings_frame.columnconfigure(1, weight=1)

        # 取得ページ数上限
        ttk.Label(settings_frame, text="取得ページ数上限 (1〜50):").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10), pady=4
        )
        pages_spin = ttk.Spinbox(
            settings_frame,
            from_=1,
            to=50,
            textvariable=self.max_pages_var,
            width=10,
        )
        pages_spin.grid(row=0, column=1, sticky=tk.W, pady=4)
        ttk.Label(settings_frame, text="※全件取得時は 50 を指定（1ページ20件）").grid(
            row=0, column=2, sticky=tk.W, padx=(8, 0), pady=4
        )

        # 仕入れ候補の最低評価
        ttk.Label(settings_frame, text="仕入れ候補の最低評価 (1〜5):").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10), pady=4
        )
        rating_combo = ttk.Combobox(
            settings_frame,
            textvariable=self.min_rating_var,
            values=["1", "2", "3", "4", "5"],
            state="readonly",
            width=8,
        )
        rating_combo.grid(row=1, column=1, sticky=tk.W, pady=4)
        ttk.Label(settings_frame, text="以上").grid(
            row=1, column=2, sticky=tk.W, padx=(8, 0), pady=4
        )

        # 仕入れ候補の上限価格
        ttk.Label(settings_frame, text="仕入れ候補の上限価格 (£):").grid(
            row=2, column=0, sticky=tk.W, padx=(0, 10), pady=4
        )
        price_entry = ttk.Entry(settings_frame, textvariable=self.max_price_var, width=12)
        price_entry.grid(row=2, column=1, sticky=tk.W, pady=4)
        ttk.Label(settings_frame, text="以下（かつ在庫「In stock」を含む）").grid(
            row=2, column=2, sticky=tk.W, padx=(8, 0), pady=4
        )

        # 出力先フォルダ選択
        ttk.Label(settings_frame, text="出力先フォルダ:").grid(
            row=3, column=0, sticky=tk.W, padx=(0, 10), pady=4
        )
        dir_frame = ttk.Frame(settings_frame)
        dir_frame.grid(row=3, column=1, columnspan=2, sticky=tk.EW, pady=4)
        dir_frame.columnconfigure(0, weight=1)

        dir_entry = ttk.Entry(dir_frame, textvariable=self.output_dir_var)
        dir_entry.grid(row=0, column=0, sticky=tk.EW, padx=(0, 6))

        browse_btn = ttk.Button(dir_frame, text="参照...", command=self.browse_output_dir)
        browse_btn.grid(row=0, column=1)

        # --- 2. 操作ボタンエリア ---
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_btn = ttk.Button(
            btn_frame, text="収集開始", command=self.start_scraping, width=16
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ttk.Button(
            btn_frame, text="中断", command=self.stop_scraping, state=tk.DISABLED, width=12
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.open_dir_btn = ttk.Button(
            btn_frame, text="出力先フォルダを開く", command=self.open_output_dir, width=20
        )
        self.open_dir_btn.pack(side=tk.RIGHT)

        # --- 3. 進捗表示エリア ---
        progress_frame = ttk.LabelFrame(main_frame, text="進捗状況", padding=10)
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        self.status_label.pack(anchor=tk.W, pady=(0, 6))

        self.progress_bar = ttk.Progressbar(
            progress_frame, orient=tk.HORIZONTAL, mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X)

        # --- 4. 実行ログ表示エリア ---
        log_frame = ttk.LabelFrame(main_frame, text="実行ログ", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=12, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

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
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, line)
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

        self.root.after(0, _append)

    def update_progress(self, current_page: int, max_pages: int) -> None:
        """メインスレッド上でプログレスバーとステータス表示を更新する。"""
        def _update() -> None:
            self.progress_bar["maximum"] = max_pages
            self.progress_bar["value"] = current_page - 1
            self.status_var.set(f"{current_page} / {max_pages} ページ目を収集中...")

        self.root.after(0, _update)

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
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self.progress_bar["maximum"] = max_pages
        self.progress_bar["value"] = 0
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
            self.stop_btn.configure(state=tk.DISABLED)
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
                        self.progress_bar["value"] = max_pages
                        self.status_var.set(
                            f"完了！ 全 {len(df)} 件（仕入れ候補 {len(candidates_df)} 件）を出力しました"
                        )

                self.root.after(0, _on_done)
            else:
                self.append_log("[警告] 取得されたデータが 0 件のため、Excel出力をスキップしました。")

                def _on_empty() -> None:
                    if was_stopped:
                        self.status_var.set("中断完了（取得データ 0 件）")
                    else:
                        self.status_var.set("エラーまたはデータなしのため終了しました")

                self.root.after(0, _on_empty)

        except Exception as e:
            self.append_log(f"[エラー] 予期しない例外が発生しました: {e}")
            self.root.after(0, lambda: self.status_var.set("エラーにより停止しました"))
        finally:
            def _restore_buttons() -> None:
                self.start_btn.configure(state=tk.NORMAL)
                self.stop_btn.configure(state=tk.DISABLED)

            self.root.after(0, _restore_buttons)


def main() -> None:
    root = tk.Tk()
    app = BookScraperApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

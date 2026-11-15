# ECサイト商品データ自動収集＆条件抽出ツール (book-scraper-tool)

対象ECサイト（[Books to Scrape](https://books.toscrape.com/)）から商品データを自動収集し、全商品データおよび指定条件に合致する「仕入れ候補リスト」をExcelファイル（`.xlsx`）へ自動出力するWindows向けデスクトップアプリケーションです。

---

## 主な機能

- **Webスクレイピング自動収集**
  - 指定したページ数（1〜最大50ページ、最大1,000件）を巡回し、各商品の「商品名 (`title`)」「価格 (`price`)」「在庫状況 (`stock`)」「評価 (`rating`: 1〜5)」「商品詳細URL (`url`)」を抽出します。
  - サーバー負荷対策として1ページごとに1秒のウェイト（`time.sleep(1.0)`）を挿入。
  - 通信エラー・タイムアウト時は最大3回まで自動リトライし、異常時は安全に停止するエラーハンドリングを搭載。
- **条件抽出＆Excel 2シート出力**
  - 取得した全データを収録する **`全商品一覧`** シートと、指定条件（最低評価・上限価格・在庫あり `"In stock"`）を満たす商品のみを抽出した **`仕入れ候補`** シートの2シート構成で `.xlsx` ファイルを生成します。
- **直感的なデスクトップGUI (`tkinter`)**
  - 取得ページ数上限・最低評価・上限価格・出力先フォルダをGUIから簡単に変更可能。
  - 前回選択した出力先フォルダを `config.json` に自動保存。
  - `threading` を用いた非同期実行により、収集中の画面フリーズを防止。
  - リアルタイムのプログレスバー・実行ログ表示、安全な「中断」機能、ワンクリックで保存先を開く「出力先フォルダを開く」ボタンを搭載。

---

## 技術スタック

- **言語**: Python 3
- **通信・HTML解析**: `requests`, `beautifulsoup4`
- **データ処理・Excel出力**: `pandas`, `openpyxl`
- **GUI**: `tkinter`（Python標準ライブラリ）
- **実行ファイル化**: `PyInstaller`

---

## ファイル構成

```text
book-scraper-tool/
├── scraper.py         # スクレイピング・データ抽出・Excel出力コアモジュール
├── gui.py             # TkinterデスクトップGUIアプリケーション
├── SPEC.md            # 要件・仕様書
├── requirements.txt   # 依存パッケージ一覧
└── README.md          # 本ドキュメント
```

---

## 開発環境での実行手順

### 1. リポジトリのクローンと仮想環境の準備

```powershell
git clone https://github.com/<your-username>/book-scraper-tool.git
cd book-scraper-tool

python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 2. GUIアプリケーションの起動

```powershell
python gui.py
```

### 3. コアモジュール単体でのテスト実行（CLI・2ページ分）

```powershell
python scraper.py
```

### 4. スタンドアロン実行ファイル（`.exe`）のビルド手順

`PyInstaller` を使用して、Python環境のないWindows環境でも動作する単一実行ファイルを生成できます。

```powershell
pip install pyinstaller
pyinstaller --onefile --noconsole --name "EC商品データ自動収集ツール" gui.py
```

ビルド完了後、`dist/EC商品データ自動収集ツール.exe` が出力されます。

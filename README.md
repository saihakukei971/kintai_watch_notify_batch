# 勤怠表自動転記ツール

勤怠システムからダウンロードしたCSVデータをExcel形式の勤怠表に自動転記するツールです。
Tkinterを使用せず、コマンドライン（CLI）もしくはバッチファイルでの実行に対応しています。

## 機能

- CSVファイルからExcel勤怠表への自動転記
- kintone連携（APIを使用したデータ取得・送信）
- 入力フォルダの監視と自動処理
- 提出状況の確認とSlack/メールでのリマインド通知
- ログ機能（処理履歴・エラーの記録）

## 動作環境

- Python 3.8以上
- Windows（WinPython推奨）またはMac/Linux

## インストール方法

### 1. WinPythonを使用する場合（推奨）

1. [WinPython](https://winpython.github.io/) から最新版の64bit版（Tkなし）をダウンロード
2. ダウンロードしたZIPファイルを適当なフォルダに解凍
3. 解凍したフォルダ内の`WPy64-xxxx\python-x.xx.x.amd64\python.exe`を使用

### 2. 通常のPythonを使用する場合

1. [Python公式サイト](https://www.python.org/downloads/) から最新版をダウンロードしてインストール
2. インストール時に「Add Python to PATH」にチェックを入れる

### 3. 必要なライブラリのインストール

コマンドプロンプト（またはターミナル）で以下のコマンドを実行：

```bash
pip install -r requirements.txt
```

## 使い方

### 1. 基本的な使い方

1. `input` フォルダに勤怠詳細CSVファイルを配置
2. `run.bat` をダブルクリック
3. 処理が完了すると `output` フォルダが自動的に開き、変換されたExcelファイルが表示されます

### 2. コマンドラインでの実行

```bash
# 基本的な実行
python main.py run --file input/勤怠詳細_202502_社員名.csv --template templates/勤怠表雛形_2025年版.xlsx

# ヘルプの表示
python main.py --help
python main.py run --help

# kintoneからデータを取得
python main.py run --mode kintone_pull --app_name "勤怠アプリ" --out_file input/kintone_data.csv

# kintoneにデータを送信
python main.py run --mode kintone_push --file output/集計結果.csv --app_name "勤怠集計アプリ"

# 環境の健全性チェック
python main.py check

# 設定情報の表示
python main.py info
```

### 3. フォルダ監視モード

`run_watcher.bat` をダブルクリックするか、以下のコマンドを実行：

```bash
python main.py watch --directory input --pattern "*.csv"
```

これにより、`input` フォルダを監視し、新しいCSVファイルが追加されると自動的に処理が実行されます。

### 4. kintone連携

1. `.env` ファイルまたは `config/settings.toml` にkintoneの接続情報を設定
2. `run_kintone.bat` をダブルクリック
3. アプリ名を入力（またはバッチファイルの引数として指定）

## 設定

設定は以下の2つのファイルで管理されています：

1. `config/settings.toml` - 一般的な設定
2. `.env` - API トークンなどの機密情報

### 設定例

#### config/settings.toml

```toml
[default]
employee_name = "社員名"
template_path = "templates/勤怠表雛形_2025年版.xlsx"
input_dir = "input"
output_dir = "output"
log_dir = "logs"
```

#### .env

```env
KINTONE_DOMAIN=your-subdomain.cybozu.com
KINTONE_API_TOKEN=your-api-token
SLACK_WEBHOOK_URL=https://hooks.slack.com/services/XXX/YYY/ZZZ
```

## 提出状況確認とリマインド通知

以下のようにして、提出状況を確認し、未提出者にリマインドを送信できます：

1. `config/members.csv` に社員リストを設定
2. `config/settings.toml` に締切日を設定
3. `.env` にSlackのWebhook URLまたはメール設定を追加
4. 以下のコマンドを実行：

```bash
python notifier.py
```

## 定期実行の設定（Windows）

Windowsのタスクスケジューラを使用して、定期的に実行するよう設定できます。

1. タスクスケジューラを開く（スタートメニューから検索）
2. 「基本タスクの作成」をクリック
3. 名前と説明を入力し、「次へ」
4. トリガーを選択（毎日など）し、時間を設定
5. 「プログラムの開始」を選択
6. プログラム/スクリプトに `python.exe` のフルパスを指定
7. 引数に `C:\path\to\your\tool\main.py run --file C:\path\to\your\tool\input\latest.csv` などと指定
8. 「完了」をクリック

## 監視モードの自動起動

1. 「run_watcher.bat」のショートカットを作成
2. Windowsのスタートアップフォルダに配置
   - `Win+R` キーを押して、`shell:startup` と入力して開く
   - 作成したショートカットをコピー

## エラーと対処法

よくあるエラーとその対処法：

| エラー | 原因 | 対処法 |
|--------|------|--------|
| FileNotFoundError | ファイルが存在しない | ファイルパスを確認 |
| PermissionError | ファイルが他のプログラムで開かれている | Excelなど他のプログラムを閉じる |
| UnicodeDecodeError | CSVのエンコーディングが不正 | 正しいエンコーディングを指定 |
| KeyError | CSVに必要なカラムがない | CSVフォーマットを確認 |

詳細なエラーログは `logs` フォルダに保存されています。問題解決の手助けにご利用ください。

## ファイル構成

```
kintai_tool/
├── main.py                # CLIエントリーポイント
├── config.py              # 設定管理
├── utils.py               # 汎用ユーティリティ関数
├── processors/            # 実際の処理ロジック
│   ├── __init__.py
│   ├── csv_processor.py   # CSV処理
│   ├── excel_processor.py # Excel処理
│   └── kintone_client.py  # kintone API連携
├── watcher.py             # ファイル監視処理
├── notifier.py            # 通知機能
├── run.bat                # 基本実行バッチ
├── run_kintone.bat        # kintone連携用バッチ
├── run_watcher.bat        # 監視開始用バッチ
├── .env                   # 環境変数（APIキーなど）
├── config/
│   ├── settings.toml      # 設定ファイル
│   └── members.csv        # 社員リスト
├── input/                 # 入力ファイル配置場所
│   └── .gitkeep
├── output/                # 出力ファイル保存場所
│   └── .gitkeep
├── logs/                  # ログファイル
│   └── .gitkeep
├── templates/             # Excelテンプレート
│   └── 勤怠表雛形_2025年版.xlsx
└── requirements.txt       # 必要なライブラリリスト
```

## 技術スタック

- **typer**: コマンドラインインターフェース（CLI）
- **dynaconf**: 設定管理
- **pandas**: CSV/Excelデータ操作
- **openpyxl**: Excel操作
- **httpx/requests**: API連携
- **watchdog**: ファイル監視
- **loguru**: ロギング
- **rich**: リッチなコンソール出力
- **tqdm**: 進捗表示


## 利用方法の補足（一般ユーザー向け）

### メニュー画面からの操作

`all_in_one.bat` をダブルクリックするとメニュー画面が表示されます。キーボードで番号を入力して操作できます：

1. **勤怠CSV処理（最新ファイル）**: input フォルダにある最新のCSVファイルを自動処理
2. **勤怠CSV処理（ファイル選択）**: ファイル選択ダイアログでCSVファイルを選択して処理
3. **フォルダ監視開始（業務時間）**: 8時間（標準業務時間）フォルダを監視し、新しいファイルがあれば自動処理
4. **kintoneからデータ取得**: kintoneからデータを取得してCSVに保存
5. **kintoneにデータ送信**: CSVファイルをkintoneに送信
6. **提出状況確認とリマインド送信**: 未提出者をチェックして通知を送信
7. **設定情報の表示**: 現在の設定情報を表示
8. **終了**: プログラムを終了

### 監視機能について

`run_watcher.bat` は、業務時間（デフォルト8時間）の間だけフォルダを監視し、その時間が過ぎると自動的に終了します。毎日の業務開始時に実行するだけで、CSV配置による自動処理が可能になります。

### タスクスケジューラ設定

`setup_task.bat` を管理者権限で実行すると、Windowsのタスクスケジューラに自動実行タスクを登録できます。これにより、毎日または平日の指定時間に自動的にリマインド確認・送信が行われます。

## 管理者向け設定・カスタマイズ

### 締切日の設定

締切日は `config/settings.toml` 内の `deadline_day` で設定できます。デフォルトでは毎月25日が設定されています。

```toml
# 締切設定
deadline_day = 25  # 毎月の締切日
```

特定の日付を指定する場合は、代わりに `deadline` を使用します：

```toml
deadline = "2025-05-25"  # 固定の締切日
```

### 従業員名について

`config/settings.toml` の `employee_name` は、CSVファイル名から従業員名を抽出できない場合のデフォルト名として使用されます。

例えば:
- CSVファイル名が「勤怠詳細_202502_山田太郎.csv」なら、山田太郎と認識します
- 命名規則に従っていない場合（「勤怠.csv」など）は、設定された `employee_name` が使用されます

このため、複数の従業員が同じPCを使う場合は、CSVファイル名に名前を含めるよう指導するか、処理前に設定ファイルを更新してください。

### 統合スケジュールタスクについて

`scheduled_tasks.bat` は、以下をすべて一度に実行します：

1. 提出状況の確認
2. リマインド通知の送信（必要な場合）
3. 当日が締切日なら未処理のCSVを一括処理
4. kintoneへのデータ送信（設定されている場合）

このバッチファイルを毎日実行するようスケジュール設定することで、手動操作なしで完全自動運用が可能になります。


# 📘 勤怠管理ツール ユースケース集（統合マニュアル）

このツールは、**CSVからの勤怠データ取り込み、Excel出力、リマインド通知、kintone連携**などを通じて、組織全体の勤怠管理業務を効率化する統合型ワークフロー支援ソリューションです。以下に各職種・部門別の具体的な使用シナリオを示します。

---




# 勤怠管理自動化ツール

## 👩‍💼 管理者のユースケース（人事部・総務部向け）

### 1. 月初の設定準備
* `config/settings.toml` の `deadline`（締切日）を更新
* `config/members.csv` を編集し、新入社員の追加・退職者の削除を反映
* `notifier.py` をタスクスケジューラに登録（未登録の場合）

📅 例：毎月20日と25日の9:00にSlackで未提出リマインド通知

### 2. 自動リマインド管理
- タスクスケジューラによって `notifier.py` が自動実行され、未提出者に Slack またはメールでリマインドが送信される  
- 管理者は通知のコピーを受信し、必要に応じて個別にフォローアップを行う

### 3. 月末の集計作業
- 全従業員の勤怠表が提出された後、`output/` フォルダ内の Excel ファイルを手動で収集・集計  
- もしくは `run_kintone.bat` を実行して、kintone に勤怠情報を一括アップロード  
- 例外ケース（休職者、途中入社など）は手動で対応

## 👨‍💼 一般社員（エンドユーザー）のユースケース

### 1. 月次の勤怠表作成
- 勤怠システム（例：freeeなど）から自分の勤怠データを CSV 形式でダウンロード  
- ダウンロードした CSV を `input/` フォルダに配置  
- `run.bat` をダブルクリックで実行  
- 自動で `output/` フォルダが開き、作成された Excel ファイルを確認  
- 必要に応じて内容を修正後、上長に提出 or 社内共有フォルダに保存

### 2. 複数部下の勤怠管理（部門マネージャー向け）
- `run_watcher.bat` を実行して監視モードを起動  
- 部下が送ってきた複数の CSV を `input/` フォルダにドラッグ＆ドロップ  
- 各ファイルが順に自動処理され、Excel に変換される  
- すべて処理完了後、`output/` フォルダからまとめて部下の勤怠表を取得可能

## 🖥️ 情報システム部門のユースケース

### 1. 初期導入と社内配布
- WinPython などのローカル Python 実行環境を準備し、`requirements.txt` に記載された依存ライブラリをインストール  
- `settings.toml` および `.env` を各社環境にあわせて設定（例：kintone API 認証など）  
- `templates/勤怠表雛形_2025年版.xlsx` を各社内運用に合わせて調整  
- パッケージ全体をネットワーク共有ドライブに配置、または ZIP で一括配布  
- 一般社員・マネージャー向けの簡易操作マニュアルを作成して展開

### 2. 自動化システムの構築
- 社内勤怠システムからの CSV 出力をバッチ化・自動化（例：cron / PowerShell）  
- 出力された CSV を `input/` に転送し、自動で処理される構成を構築  
- `run_kintone.bat` または直接 `kintone_client.py` を用いて、処理結果を API 経由で kintone に反映  
- `logs/` フォルダを定期的にチェックし、エラーや未処理ファイルの早期発見を行う

### 3. メンテナンスと継続運用
- 年度切り替えや制度改定に応じて、Excel テンプレートや `settings.toml` を更新  
- 組織改編があるたびに `members.csv` を見直し、役職・所属変更などを反映  
- Python や外部ライブラリのアップデートに伴う動作検証と適用  
- 障害発生時には `logs/` 内の最新ファイルを確認し、原因特定と対応を迅速に実施

## 💡 統合的な導入価値

このツールは単なる **CSV→Excel 変換ツールではありません**。  
以下のように、**企業の勤怠管理ワークフロー全体を自動化・効率化する統合型ソリューション**として機能します。

- ✅ 勤怠提出の抜け漏れ防止（Slack/メールによる自動リマインド）
- ✅ マネージャーの複数メンバー管理を高速化（監視モード＋自動変換）
- ✅ 情報システム部門との連携による kintone など社内基幹システムとの接続
- ✅ 運用定着を支えるテンプレートカスタマイズ・バッチ配布・ログ監視体制
---

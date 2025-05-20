@echo off
rem kintone連携実行バッチファイル

echo === kintone連携ツール ===

rem タイムスタンプを生成
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "year=%dt:~0,4%"
set "month=%dt:~4,2%"
set "day=%dt:~6,2%"
set "hour=%dt:~8,2%"
set "minute=%dt:~10,2%"
set "second=%dt:~12,2%"
set "timestamp=%year%%month%%day%_%hour%%minute%%second%"

rem アプリ名（引数から取得、または入力を求める）
if "%~1"=="" (
    set /p app_name="kintoneアプリ名を入力してください: "
) else (
    set "app_name=%~1"
)

echo アプリ名: %app_name%
echo タイムスタンプ: %timestamp%

rem 一時CSVファイル名
set "csv_file=input\kintone_data_%timestamp%.csv"

rem kintoneからデータを取得
echo データを取得しています...
python main.py run --mode kintone_pull --app_name "%app_name%" --out_file "%csv_file%"

if %errorlevel% neq 0 (
    echo エラーが発生しました。
    pause
    exit /b 1
)

echo CSVファイルに保存しました: %csv_file%

rem データを処理
echo データを処理しています...
python main.py run --file "%csv_file%" --template "templates\勤怠表雛形_2025年版.xlsx"

if %errorlevel% neq 0 (
    echo 処理中にエラーが発生しました。
) else (
    echo 処理が完了しました。出力フォルダが開きます。
    explorer output
)

pause
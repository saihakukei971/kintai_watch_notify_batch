@echo off
rem 勤怠表自動生成ツール実行バッチファイル

rem 最新のCSVファイルを検索
for /f "delims=" %%a in ('dir /b /a-d /od input\*.csv 2^>nul') do set "latest_csv=%%a"

if not defined latest_csv (
    echo ファイルが見つかりません。input フォルダにCSVファイルを配置してください。
    pause
    exit /b 1
)

echo 最新のCSVファイル: %latest_csv%
echo 処理を開始します...

rem Pythonスクリプトを実行
python main.py run --file "input\%latest_csv%" --template "templates\勤怠表雛形_2025年版.xlsx"

if %errorlevel% neq 0 (
    echo エラーが発生しました。
) else (
    echo 処理が完了しました。出力フォルダが開きます。
    explorer output
)

pause 

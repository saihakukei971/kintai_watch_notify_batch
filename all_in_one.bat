@echo off
color 0A
cls
echo ============================================
echo         勤怠表自動処理ツール メニュー
echo ============================================
echo.

:MENU
echo 操作を選択してください:
echo.
echo 1. 勤怠CSV処理（最新ファイル）
echo 2. 勤怠CSV処理（ファイル選択）
echo 3. フォルダ監視開始（業務時間）
echo 4. kintoneからデータ取得
echo 5. kintoneにデータ送信
echo 6. 提出状況確認とリマインド送信
echo 7. 設定情報の表示
echo 8. 終了
echo.

set /p choice="番号を入力してください (1-8): "

if "%choice%"=="1" goto NORMAL
if "%choice%"=="2" goto SELECT_FILE
if "%choice%"=="3" goto WATCH
if "%choice%"=="4" goto KINTONE_GET
if "%choice%"=="5" goto KINTONE_PUSH
if "%choice%"=="6" goto NOTIFY
if "%choice%"=="7" goto INFO
if "%choice%"=="8" goto END

echo 無効な選択です。もう一度お試しください。
goto MENU

:NORMAL
cls
echo 最新のCSVファイルを処理します...
call run.bat
pause
cls
goto MENU

:SELECT_FILE
cls
echo CSVファイルを選択してください。
set "psCommand="(new-object -COM 'Shell.Application').BrowseForFolder(0,'CSVファイルを選択してください',0,0).self.path""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "folder=%%I"
if "%folder%"=="" goto MENU

set "psCommand="Get-ChildItem -Path '%folder%' -Filter *.csv | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "csv_file=%%I"

if "%csv_file%"=="" (
  echo CSVファイルが見つかりませんでした。
  pause
  cls
  goto MENU
)

echo 選択されたファイル: %csv_file%
echo 処理を開始します...

python main.py run --file "%csv_file%" --template "templates\勤怠表雛形_2025年版.xlsx"

if %errorlevel% neq 0 (
  echo エラーが発生しました。
) else (
  echo 処理が完了しました。出力フォルダが開きます。
  explorer output
)

pause
cls
goto MENU

:WATCH
cls
echo フォルダ監視を開始します（8時間）...
start "フォルダ監視" run_watcher.bat
echo 別ウィンドウで監視を開始しました。
pause
cls
goto MENU

:KINTONE_GET
cls
echo kintoneからデータを取得します。
set /p app_name="kintoneアプリ名を入力してください: "
call run_kintone.bat "%app_name%"
pause
cls
goto MENU

:KINTONE_PUSH
cls
echo kintoneにデータを送信します。
set /p app_name="kintoneアプリ名を入力してください: "

set "psCommand="(new-object -COM 'Shell.Application').BrowseForFolder(0,'送信するCSVファイルを選択してください',0,0).self.path""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "folder=%%I"
if "%folder%"=="" goto MENU

set "psCommand="Get-ChildItem -Path '%folder%' -Filter *.csv | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName""
for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "csv_file=%%I"

if "%csv_file%"=="" (
  echo CSVファイルが見つかりませんでした。
  pause
  cls
  goto MENU
)

echo 選択されたファイル: %csv_file%
echo kintoneアプリ: %app_name%
echo 送信を開始します...

python main.py run --mode kintone_push --file "%csv_file%" --app_name "%app_name%"

if %errorlevel% neq 0 (
  echo エラーが発生しました。
) else (
  echo 送信が完了しました。
)

pause
cls
goto MENU

:NOTIFY
cls
echo 提出状況確認とリマインド送信を実行します...
python notifier.py

if %errorlevel% neq 0 (
  echo エラーが発生しました。
) else (
  echo 処理が完了しました。
)

pause
cls
goto MENU

:INFO
cls
echo 設定情報を表示します...
python main.py info
pause
cls
goto MENU

:END
cls
echo 勤怠表自動処理ツールを終了します。
echo ご利用ありがとうございました。
exit /b 0
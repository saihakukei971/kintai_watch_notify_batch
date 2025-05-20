@echo off
echo === 勤怠管理定期タスク実行 ===
echo 実行日時: %date% %time%

REM ログファイル設定
set log_dir=logs
if not exist %log_dir% mkdir %log_dir%

set log_file=%log_dir%\scheduled_task_%date:~0,4%%date:~5,2%%date:~8,2%.log
echo 実行開始: %date% %time% > %log_file%

REM 1. 提出状況の確認とリマインド送信
echo 提出状況確認とリマインド送信を実行します... >> %log_file%
python notifier.py >> %log_file% 2>&1
if %errorlevel% neq 0 (
  echo [エラー] 提出状況確認に失敗しました: %errorlevel% >> %log_file%
) else (
  echo 提出状況確認が完了しました >> %log_file%
)

REM 2. 当日が締切日の場合は処理を実行
python -c "from datetime import datetime; from config import init_config, get_deadline_date; config=init_config(); deadline=get_deadline_date(); today=datetime.now().strftime('%%Y-%%m-%%d'); print('YES' if deadline == today else 'NO')" > tmp.txt
set /p is_deadline=<tmp.txt
del tmp.txt

if "%is_deadline%"=="YES" (
  echo 締切日のため最終処理を実行します... >> %log_file%
  
  REM 未処理ファイルを一括処理
  for /r input %%f in (*.csv) do (
    echo ファイル処理: %%f >> %log_file%
    python main.py run --file "%%f" >> %log_file% 2>&1
  )
  
  REM kintoneにデータをアップロード（必要な場合）
  if exist "config\.upload_to_kintone" (
    echo kintoneにデータをアップロードします... >> %log_file%
    python main.py run --mode kintone_push --file "output\集計結果.csv" --app_name "勤怠集計" >> %log_file% 2>&1
  )
)

echo 実行終了: %date% %time% >> %log_file%

REM エラーがあれば表示
find /c "[エラー]" %log_file% > nul
if not errorlevel 1 (
  echo エラーが発生しました。詳細はログファイルを確認してください: %log_file%
  pause
) else (
  echo 処理が正常に完了しました。
)
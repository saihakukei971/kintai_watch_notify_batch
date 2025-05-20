@echo off
rem ファイル監視実行バッチファイル（業務時間監視版）

echo === ファイル監視ツール（業務時間監視版） ===
echo inputフォルダを監視し、新しいCSVファイルを検出すると自動的に処理します。
echo 8時間（標準業務時間）後に自動的に終了します。
echo 監視を手動で停止するには Ctrl+C を押してください。

rem Pythonスクリプトを実行（デフォルトで8時間）
python main.py watch --directory input --pattern "*.csv" --hours 8

echo 監視を終了しました。
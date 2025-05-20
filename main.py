#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
勤怠表自動転記ツール - メインエントリーポイント
"""
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional, List

import typer
from rich.console import Console
from rich.traceback import install
from loguru import logger

# 自作モジュールのインポート
from config import Config, init_config
from processors.csv_processor import read_csv, process_data
from processors.excel_processor import write_to_excel
from processors.kintone_client import KintoneClient
from utils import setup_logging, ensure_directories, find_latest_file

# リッチなトレースバックを有効化
install(show_locals=True)
console = Console()

# アプリケーションの初期化
app = typer.Typer(help="勤怠表自動変換ツール", add_completion=False)
conf = init_config()

# ロギングの設定
setup_logging()

@app.callback()
def callback():
    """勤怠表自動変換ツール - 勤怠CSVをExcelに転記します"""
    # コールバックの内容はオプション
    ensure_directories()
    logger.info("勤怠表自動変換ツールを起動しました")


@app.command("run")
def run(
    file: str = typer.Option(..., "--file", "-f", help="処理するCSVファイルのパス"),
    template: Optional[str] = typer.Option(None, "--template", "-t", help="テンプレートExcelファイルのパス"),
    name: Optional[str] = typer.Option(None, "--name", "-n", help="従業員名を指定"),
    mode: str = typer.Option("normal", "--mode", "-m", help="処理モード (normal, kintone_pull, kintone_push)"),
    app_name: Optional[str] = typer.Option(None, "--app_name", help="kintoneアプリ名"),
    out_file: Optional[str] = typer.Option(None, "--out_file", help="出力ファイル名"),
):
    """CSVファイルをExcelの勤怠表に変換します"""
    try:
        # パスの正規化
        file = Path(file).resolve()

        # テンプレートが指定されていない場合はデフォルトを使用
        if not template:
            template = Path(conf.TEMPLATE_PATH).resolve()
        else:
            template = Path(template).resolve()

        # ファイルの存在確認
        if not file.exists() and mode != "kintone_pull":
            logger.error(f"指定されたファイルが存在しません: {file}")
            console.print(f"[bold red]エラー:[/] 指定されたファイルが存在しません: {file}")
            return 1

        if not template.exists():
            logger.error(f"指定されたテンプレートファイルが存在しません: {template}")
            console.print(f"[bold red]エラー:[/] 指定されたテンプレートファイルが存在しません: {template}")
            return 1

        # 従業員名が指定されていない場合はデフォルトを使用
        if not name:
            name = conf.EMPLOYEE_NAME

        # 処理モードに応じた処理を実行
        if mode == "kintone_pull":
            if not app_name:
                logger.error("kintone_pullモードではapp_nameが必須です")
                console.print("[bold red]エラー:[/] kintone_pullモードではapp_nameが必須です")
                return 1

            # kintoneからデータを取得
            kintone = KintoneClient(conf.KINTONE_DOMAIN, conf.KINTONE_API_TOKEN)
            records = kintone.get_records(app_name)

            # 出力ファイル名が指定されていない場合はデフォルトを生成
            if not out_file:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_file = os.path.join(conf.INPUT_DIR, f"kintone_data_{timestamp}.csv")

            # CSVとして保存
            kintone.save_as_csv(records, out_file)
            logger.info(f"kintoneからデータを取得し、{out_file}に保存しました")
            console.print(f"[bold green]成功:[/] kintoneからデータを取得し、{out_file}に保存しました")

            # 取得したCSVを処理対象に設定
            file = Path(out_file).resolve()

        elif mode == "kintone_push":
            if not file.exists():
                logger.error(f"アップロードするファイルが存在しません: {file}")
                console.print(f"[bold red]エラー:[/] アップロードするファイルが存在しません: {file}")
                return 1

            if not app_name:
                logger.error("kintone_pushモードではapp_nameが必須です")
                console.print("[bold red]エラー:[/] kintone_pushモードではapp_nameが必須です")
                return 1

            # CSVからデータを読み込みkintoneにアップロード
            kintone = KintoneClient(conf.KINTONE_DOMAIN, conf.KINTONE_API_TOKEN)
            records = kintone.csv_to_records(str(file))
            result = kintone.add_records(app_name, records)

            logger.info(f"kintoneにデータをアップロードしました: {result}")
            console.print(f"[bold green]成功:[/] kintoneにデータをアップロードしました")
            return 0

        # 通常の処理フロー
        logger.info(f"ファイル処理開始: {file}")
        console.print(f"[bold]処理開始:[/] {file}")

        # CSVデータを読み込み
        with console.status("[bold green]CSVファイルを読み込んでいます..."):
            df = read_csv(str(file))
            logger.info(f"CSVファイル読み込み完了: {len(df)}行")

        # CSVから月情報を取得
        first_date = df["日付"].min()
        year_month = first_date.strftime("%Y%m")

        # データの整形
        with console.status("[bold green]データを処理しています..."):
            df_processed = process_data(df)
            logger.info("データ処理完了")

        # 出力ファイル名を作成
        output_filename = f"勤怠表_{year_month}_{name}.xlsx"
        output_path = os.path.join(conf.OUTPUT_DIR, output_filename)

        # Excelに書き込み
        with console.status("[bold green]Excelファイルに書き込んでいます..."):
            write_to_excel(str(template), output_path, df_processed, str(file))
            logger.info(f"Excelファイル書き込み完了: {output_path}")

        console.print(f"[bold green]✅ 処理完了:[/] 勤怠表を作成しました: {output_path}")
        return 0

    except Exception as e:
        logger.exception(f"処理中にエラーが発生しました: {str(e)}")
        console.print(f"[bold red]エラー:[/] 処理中にエラーが発生しました: {str(e)}")
        return 1


@app.command("watch")
def watch(
    directory: str = typer.Option("input", "--directory", "-d", help="監視するディレクトリ"),
    pattern: str = typer.Option("*.csv", "--pattern", "-p", help="監視するファイルパターン"),
    hours: int = typer.Option(8, "--hours", "-h", help="監視を継続する時間（時間）。デフォルトは8時間（業務時間）"),
):
    """指定されたディレクトリを監視し、新しいファイルが追加されたら自動的に処理します"""
    from watcher import start_watching

    logger.info(f"ディレクトリの監視を開始します: {directory}, パターン: {pattern}, 時間: {hours}時間")
    console.print(f"[bold]ディレクトリの監視を開始します:[/] {directory}, パターン: {pattern}, 時間: {hours}時間")
    console.print("監視を停止するには Ctrl+C を押してください")

    try:
        start_watching(directory, pattern, hours)
    except KeyboardInterrupt:
        logger.info("ユーザーによって監視が停止されました")
        console.print("[bold yellow]監視を停止しました[/]")
    except Exception as e:
        logger.exception(f"監視中にエラーが発生しました: {str(e)}")
        console.print(f"[bold red]エラー:[/] 監視中にエラーが発生しました: {str(e)}")
        return 1

    return 0


@app.command("check")
def check():
    """環境の健全性チェック"""
    issues = []

    # ディレクトリの確認
    dirs_to_check = [conf.INPUT_DIR, conf.OUTPUT_DIR, conf.LOG_DIR, os.path.dirname(conf.TEMPLATE_PATH)]
    for d in dirs_to_check:
        if not os.path.isdir(d):
            issues.append(f"ディレクトリが存在しません: {d}")

    # テンプレートファイルの確認
    if not os.path.isfile(conf.TEMPLATE_PATH):
        issues.append(f"テンプレートファイルが存在しません: {conf.TEMPLATE_PATH}")

    # kintone接続情報の確認
    if not conf.KINTONE_DOMAIN or not conf.KINTONE_API_TOKEN:
        issues.append("kintone接続情報が設定されていません")

    if issues:
        console.print("[bold red]環境チェック結果: 問題が見つかりました[/]")
        for issue in issues:
            console.print(f"  - {issue}")
        return 1

    console.print("[bold green]環境チェック結果: 問題なし[/]")
    return 0


@app.command("info")
def info():
    """設定情報の表示"""
    console.print("[bold]設定情報:[/]")
    console.print(f"  従業員名: {conf.EMPLOYEE_NAME}")
    console.print(f"  テンプレートパス: {conf.TEMPLATE_PATH}")
    console.print(f"  入力ディレクトリ: {conf.INPUT_DIR}")
    console.print(f"  出力ディレクトリ: {conf.OUTPUT_DIR}")
    console.print(f"  ログディレクトリ: {conf.LOG_DIR}")
    console.print(f"  kintoneドメイン: {conf.KINTONE_DOMAIN if conf.KINTONE_DOMAIN else '未設定'}")
    console.print(f"  kintone APIトークン: {'設定済み' if conf.KINTONE_API_TOKEN else '未設定'}")

    # 最新のCSVファイルを検索
    latest_csv = find_latest_file(conf.INPUT_DIR, "*.csv")
    if latest_csv:
        console.print(f"  最新のCSVファイル: {os.path.basename(latest_csv)}")

    return 0


if __name__ == "__main__":
    app()
#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ファイル監視処理モジュール
"""
import os
import time
import subprocess
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, List, Dict

import watchdog.events
import watchdog.observers
from loguru import logger
from rich.console import Console

from config import init_config
from utils import find_latest_file, extract_employee_name_from_filename

console = Console()


class FileHandler(watchdog.events.PatternMatchingEventHandler):
    """ファイル変更イベントハンドラ"""

    def __init__(self, patterns=None, ignore_patterns=None, ignore_directories=True, case_sensitive=False, config=None):
        super().__init__(patterns, ignore_patterns, ignore_directories, case_sensitive)
        self.config = config or init_config()
        self.processing_files = set()  # 処理中のファイル

    def on_created(self, event):
        """ファイル作成イベント"""
        if event.is_directory:
            return

        file_path = event.src_path

        # すでに処理中のファイルはスキップ
        if file_path in self.processing_files:
            return

        # 一時ファイルは無視
        if file_path.endswith('.tmp') or '~$' in file_path or file_path.startswith('.'):
            return

        # ファイルが完全に書き込まれるまで少し待機
        self._wait_for_file_ready(file_path)

        # 処理中フラグを立てる
        self.processing_files.add(file_path)

        try:
            logger.info(f"新しいファイルを検出しました: {file_path}")
            console.print(f"[bold green]新しいファイルを検出:[/] {os.path.basename(file_path)}")

            # 従業員名を取得
            employee_name = extract_employee_name_from_filename(os.path.basename(file_path))
            if not employee_name:
                employee_name = self.config.EMPLOYEE_NAME

            # テンプレートパス
            template_path = self.config.TEMPLATE_PATH

            # 処理を実行
            self._process_file(file_path, template_path, employee_name)

        except Exception as e:
            logger.exception(f"ファイル処理エラー: {str(e)}")
            console.print(f"[bold red]エラー:[/] {str(e)}")

        finally:
            # 処理中フラグを解除
            self.processing_files.discard(file_path)

    def on_modified(self, event):
        """ファイル変更イベント"""
        # 作成イベントと重複するので何もしない
        pass

    def on_moved(self, event):
        """ファイル移動イベント"""
        if event.is_directory:
            return

        # 移動先がパターンにマッチするかチェック
        for pattern in self.patterns:
            if watchdog.events.match_path(pattern, event.dest_path):
                # 移動先ファイルを処理
                self.on_created(watchdog.events.FileCreatedEvent(event.dest_path))
                break

    def _wait_for_file_ready(self, file_path: str, timeout: int = 10, check_interval: float = 0.5):
        """
        ファイルが完全に書き込まれるまで待機

        Args:
            file_path: ファイルパス
            timeout: タイムアウト時間（秒）
            check_interval: チェック間隔（秒）
        """
        start_time = time.time()

        while True:
            # タイムアウトチェック
            if time.time() - start_time > timeout:
                logger.warning(f"ファイル待機タイムアウト: {file_path}")
                return

            try:
                # ファイルサイズが変化しているかチェック
                current_size = os.path.getsize(file_path)
                time.sleep(check_interval)
                new_size = os.path.getsize(file_path)

                # ファイルサイズが安定していれば完了
                if current_size == new_size:
                    return

            except (FileNotFoundError, PermissionError):
                # ファイルがまだ存在しない、またはアクセスできない場合は待機
                time.sleep(check_interval)

    def _process_file(self, file_path: str, template_path: str, employee_name: str):
        """
        ファイルを処理

        Args:
            file_path: 処理するファイルパス
            template_path: テンプレートパス
            employee_name: 従業員名
        """
        # 実行スクリプトのパスを取得
        script_path = Path(__file__).resolve().parent / "main.py"

        # コマンドの構築
        command = [
            "python",
            str(script_path),
            "run",
            "--file", file_path,
            "--template", template_path,
            "--name", employee_name
        ]

        # サブプロセスとして実行
        logger.info(f"処理コマンド: {' '.join(command)}")

        try:
            result = subprocess.run(command, check=True, capture_output=True, text=True)
            logger.info(f"処理完了: {result.stdout}")
            console.print(f"[bold green]処理完了:[/] {os.path.basename(file_path)}")

            # 標準出力からファイルパスを抽出
            output_file = self._extract_output_path(result.stdout)
            if output_file and os.path.exists(output_file):
                console.print(f"[bold]出力ファイル:[/] {os.path.basename(output_file)}")

            return True

        except subprocess.CalledProcessError as e:
            logger.error(f"処理エラー: {e.stderr}")
            console.print(f"[bold red]処理エラー:[/] {e.stderr}")
            return False

    def _extract_output_path(self, output: str) -> Optional[str]:
        """
        標準出力から出力ファイルパスを抽出

        Args:
            output: 標準出力テキスト

        Returns:
            str or None: 抽出したファイルパス
        """
        import re

        # パターン: '✅ 勤怠表を作成しました: {path}'
        match = re.search(r'勤怠表を作成しました: (.+?)$', output, re.MULTILINE)
        if match:
            return match.group(1).strip()

        return None


def start_watching(directory: str, pattern: str = "*.csv", duration_hours: int = 8):
    """
    指定されたディレクトリを監視

    Args:
        directory: 監視するディレクトリ
        pattern: ファイルパターン
        duration_hours: 監視を継続する時間（時間）。デフォルトは8時間（業務時間）
    """
    # 初期化
    config = init_config()

    # ディレクトリの絶対パスを取得
    directory = os.path.abspath(directory)

    # ディレクトリが存在しない場合は作成
    if not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)
        logger.info(f"監視ディレクトリを作成しました: {directory}")

    # 終了時刻の計算
    end_time = None
    if duration_hours:
        end_time = datetime.now() + timedelta(hours=duration_hours)
        logger.info(f"監視時間: {duration_hours}時間 (終了予定: {end_time.strftime('%H:%M:%S')})")
        console.print(f"[bold]監視時間:[/] {duration_hours}時間 (終了予定: {end_time.strftime('%H:%M:%S')})")

    # 既存のファイルを確認
    existing_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if watchdog.events.match_path(pattern, file):
                existing_files.append(os.path.join(root, file))

    if existing_files:
        console.print(f"[bold yellow]既存のファイルが{len(existing_files)}件見つかりました[/]")
        for file in existing_files:
            console.print(f"  - {os.path.basename(file)}")

        # 既存ファイルを処理するかどうか確認
        process_existing = input("既存のファイルを処理しますか？ (y/n): ").strip().lower() == 'y'

        if process_existing:
            handler = FileHandler(patterns=[pattern], config=config)
            for file in existing_files:
                handler.on_created(watchdog.events.FileCreatedEvent(file))

    # イベントハンドラの設定
    event_handler = FileHandler(patterns=[pattern], config=config)
    observer = watchdog.observers.Observer()
    observer.schedule(event_handler, directory, recursive=True)

    # 監視開始
    observer.start()
    logger.info(f"ディレクトリの監視を開始しました: {directory}, パターン: {pattern}")
    console.print(f"[bold green]監視開始:[/] {directory} ({pattern})")
    console.print("監視を停止するには Ctrl+C を押してください")

    try:
        while True:
            # 終了時間のチェック
            if end_time and datetime.now() >= end_time:
                logger.info(f"指定された監視時間({duration_hours}時間)が経過したため終了します")
                console.print(f"[bold yellow]指定された監視時間({duration_hours}時間)が経過したため終了します[/]")
                break

            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()
    logger.info("監視を停止しました")
    console.print("[bold yellow]監視を停止しました[/]")


if __name__ == "__main__":
    from utils import setup_logging

    # ロギング設定
    setup_logging()

    # デフォルトは input ディレクトリを監視（8時間 = 業務時間）
    start_watching("input", "*.csv", 8)
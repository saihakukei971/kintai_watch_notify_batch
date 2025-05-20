#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
汎用ユーティリティ関数
"""
import os
import glob
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from loguru import logger
from rich.console import Console

console = Console()


def setup_logging():
    """ロギング設定を初期化"""
    # 実行ファイルの基準パスを取得
    if getattr(sys, 'frozen', False):
        base_path = Path(sys.executable).parent
    else:
        base_path = Path(__file__).parent

    # ログディレクトリの作成
    log_dir = base_path / 'logs'
    os.makedirs(log_dir, exist_ok=True)

    # タイムスタンプを含むログファイル名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"kintai_tool_{timestamp}.log"

    # ロガーの設定
    logger.remove()  # デフォルトのハンドラーを削除

    # 標準出力へのログ
    logger.add(
        sys.stderr,
        format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
        level="INFO",
        diagnose=True
    )

    # ファイルへのログ
    logger.add(
        str(log_file),
        format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
        level="DEBUG",
        rotation="10 MB",
        compression="zip",
        diagnose=True
    )

    logger.info(f"ログファイル: {log_file}")
    return log_file


def ensure_directories():
    """必要なディレクトリを作成"""
    # 実行ファイルの基準パスを取得
    if getattr(sys, 'frozen', False):
        base_path = Path(sys.executable).parent
    else:
        base_path = Path(__file__).parent

    # 必要なディレクトリ
    dirs = ['input', 'output', 'logs', 'templates', 'config']

    for dir_name in dirs:
        dir_path = base_path / dir_name
        if not dir_path.exists():
            dir_path.mkdir(exist_ok=True)
            logger.info(f"ディレクトリを作成しました: {dir_path}")


def find_latest_file(directory: str, pattern: str) -> Optional[str]:
    """指定されたディレクトリで最新のファイルを検索"""
    files = glob.glob(os.path.join(directory, pattern))
    if not files:
        return None

    # 最新のファイルを返す
    return max(files, key=os.path.getmtime)


def backup_file(file_path: str) -> str:
    """ファイルをバックアップ"""
    if not os.path.exists(file_path):
        return None

    # バックアップファイル名を生成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{file_path}.{timestamp}.bak"

    # コピー
    shutil.copy2(file_path, backup_path)
    logger.info(f"ファイルをバックアップしました: {backup_path}")

    return backup_path


def safe_filename(filename: str) -> str:
    """ファイル名から不正な文字を削除"""
    # Windows で使用できない文字を置換
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')

    # 前後の空白を削除
    filename = filename.strip()

    # ファイル名が空になった場合
    if not filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"file_{timestamp}"

    return filename


def detect_csv_encoding(file_path: str, encodings=('utf-8', 'shift-jis', 'euc-jp', 'iso-2022-jp')) -> str:
    """CSVファイルのエンコーディングを検出"""
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                f.read()
                return encoding
        except UnicodeDecodeError:
            continue

    # デフォルトのエンコーディングを返す
    logger.warning(f"エンコーディングを検出できませんでした: {file_path}, デフォルトの utf-8 を使用します")
    return 'utf-8'


def format_time_str(seconds: float) -> str:
    """秒数を時間表記 (HH:MM) に変換"""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    return f"{hours:02d}:{minutes:02d}"


def parse_time_str(time_str: str) -> float:
    """時間表記 (HH:MM) を秒数に変換"""
    if not time_str or not isinstance(time_str, str):
        return 0.0

    try:
        if ":" in time_str:
            hours, minutes = map(int, time_str.split(':'))
            return hours * 3600 + minutes * 60
        else:
            # 数値のみの場合は時間とみなす
            return float(time_str) * 3600
    except (ValueError, TypeError):
        logger.warning(f"時間文字列のパースに失敗しました: {time_str}")
        return 0.0


def extract_employee_name_from_filename(filename: str) -> Optional[str]:
    """ファイル名から従業員名を抽出"""
    import re

    # パターン: 勤怠詳細_YYYYMM_氏名.csv
    pattern = r'勤怠詳細_\d{6}_(.+?)\.[^\.]+$'
    match = re.search(pattern, filename)

    if match:
        return match.group(1)

    # パターン: 勤怠詳細_氏名_YYYY_MM.csv
    pattern = r'勤怠詳細_(.+?)_\d{4}_\d{2}\.[^\.]+$'
    match = re.search(pattern, filename)

    if match:
        return match.group(1)

    logger.warning(f"ファイル名から従業員名を抽出できませんでした: {filename}")
    return None


def extract_year_month_from_filename(filename: str) -> Optional[tuple]:
    """ファイル名から年月を抽出"""
    import re

    # パターン: 勤怠詳細_YYYYMM_氏名.csv
    pattern = r'勤怠詳細_(\d{4})(\d{2})_'
    match = re.search(pattern, filename)

    if match:
        return (match.group(1), match.group(2))

    # パターン: 勤怠詳細_氏名_YYYY_MM.csv
    pattern = r'勤怠詳細_.*?_(\d{4})_(\d{2})\.'
    match = re.search(pattern, filename)

    if match:
        return (match.group(1), match.group(2))

    logger.warning(f"ファイル名から年月を抽出できませんでした: {filename}")
    return None
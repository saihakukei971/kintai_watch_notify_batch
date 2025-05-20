#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
設定管理モジュール
"""
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from dynaconf import Dynaconf
from loguru import logger
from dotenv import load_dotenv


@dataclass
class Config:
    """設定情報を保持するクラス"""
    EMPLOYEE_NAME: str
    TEMPLATE_PATH: str
    INPUT_DIR: str
    OUTPUT_DIR: str
    LOG_DIR: str
    KINTONE_DOMAIN: Optional[str] = None
    KINTONE_API_TOKEN: Optional[str] = None
    SLACK_WEBHOOK_URL: Optional[str] = None
    DEADLINE: Optional[str] = None
    REMIND_DAYS_BEFORE: int = 5


def get_base_path() -> Path:
    """実行ファイルの基準パスを取得"""
    if getattr(sys, 'frozen', False):
        # exe実行時
        base_path = Path(sys.executable).parent
    else:
        # 通常実行時
        base_path = Path(__file__).parent

    return base_path


def init_config() -> Config:
    """設定を初期化して返す"""
    # 基準パスの取得
    base_path = get_base_path()

    # .envファイルを読み込み (存在する場合)
    env_file = base_path / '.env'
    if env_file.exists():
        load_dotenv(dotenv_path=str(env_file))

    # 設定ファイルのパス
    settings_files = [
        base_path / 'config' / 'settings.toml',
        base_path / '.secrets.toml',
    ]

    # 存在確認
    config_dir = base_path / 'config'
    if not config_dir.exists():
        config_dir.mkdir(exist_ok=True)

    settings_path = settings_files[0]
    if not settings_path.exists():
        # デフォルトの設定ファイルを作成
        create_default_settings(settings_path)

    # Dynaconfで設定を読み込み
    settings = Dynaconf(
        envvar_prefix="KINTAI",
        settings_files=settings_files,
        environments=True,
        load_dotenv=True,
    )

    # Config オブジェクトを作成
    config = Config(
        EMPLOYEE_NAME=settings.get('employee_name', '未設定'),
        TEMPLATE_PATH=settings.get('template_path', str(base_path / 'templates' / '勤怠表雛形_2025年版.xlsx')),
        INPUT_DIR=settings.get('input_dir', str(base_path / 'input')),
        OUTPUT_DIR=settings.get('output_dir', str(base_path / 'output')),
        LOG_DIR=settings.get('log_dir', str(base_path / 'logs')),
        KINTONE_DOMAIN=settings.get('kintone_domain', os.getenv('KINTONE_DOMAIN')),
        KINTONE_API_TOKEN=settings.get('kintone_api_token', os.getenv('KINTONE_API_TOKEN')),
        SLACK_WEBHOOK_URL=settings.get('slack_webhook_url', os.getenv('SLACK_WEBHOOK_URL')),
        DEADLINE=settings.get('deadline', None),
        REMIND_DAYS_BEFORE=int(settings.get('remind_days_before', 5)),
    )

    # 必要なディレクトリがなければ作成
    ensure_directories(config)

    return config


def ensure_directories(config: Config):
    """必要なディレクトリを作成"""
    dirs = [config.INPUT_DIR, config.OUTPUT_DIR, config.LOG_DIR, os.path.dirname(config.TEMPLATE_PATH)]
    for d in dirs:
        os.makedirs(d, exist_ok=True)


def create_default_settings(settings_path: Path):
    """デフォルトの設定ファイルを作成"""
    base_path = get_base_path()

    default_settings = f"""[default]
# 基本設定
employee_name = "社員名"
template_path = "{str(base_path / 'templates' / '勤怠表雛形_2025年版.xlsx').replace('\\\\', '/')}"
input_dir = "{str(base_path / 'input').replace('\\\\', '/')}"
output_dir = "{str(base_path / 'output').replace('\\\\', '/')}"
log_dir = "{str(base_path / 'logs').replace('\\\\', '/')}"

# kintone連携設定
# kintone_domain = "your-subdomain.cybozu.com"
# kintone_api_token = "your-api-token"

# 通知設定
# slack_webhook_url = "https://hooks.slack.com/services/XXX/YYY/ZZZ"
# deadline = "2025-05-25"
# remind_days_before = 5

# 処理設定
csv_encoding = "utf-8"
date_format = "%Y-%m-%d"
"""

    # ディレクトリを作成
    os.makedirs(os.path.dirname(settings_path), exist_ok=True)

    # ファイルに書き込み
    with open(settings_path, 'w', encoding='utf-8') as f:
        f.write(default_settings)

    print(f"デフォルト設定ファイルを作成しました: {settings_path}")


def update_config(name: str = None, template: str = None):
    """設定ファイルを更新"""
    base_path = get_base_path()
    settings_path = base_path / 'config' / 'settings.toml'

    # 設定ファイルが存在しない場合は作成
    if not settings_path.exists():
        create_default_settings(settings_path)

    # 設定ファイルを読み込み
    with open(settings_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 値を更新
    if name:
        if 'employee_name =' in content:
            content = content.replace('employee_name = "' + content.split('employee_name = "')[1].split('"')[0] + '"', f'employee_name = "{name}"')
        else:
            content += f'\nemployee_name = "{name}"'

    if template:
        if 'template_path =' in content:
            content = content.replace('template_path = "' + content.split('template_path = "')[1].split('"')[0] + '"', f'template_path = "{template.replace("\\", "/")}"')
        else:
            content += f'\ntemplate_path = "{template.replace("\\", "/")}"'

    # 設定ファイルを保存
    with open(settings_path, 'w', encoding='utf-8') as f:
        f.write(content)

    logger.info(f"設定ファイルを更新しました: {settings_path}")

def get_deadline_date():
    """今月の締切日を計算（例：毎月25日）"""
    import datetime
    today = datetime.date.today()
    # 当月の25日を締切とする場合
    deadline_day = 25

    # 今月の締切日を作成
    deadline = datetime.date(today.year, today.month, deadline_day)

    # 既に今月の締切日を過ぎていれば翌月の締切日を返す
    if today > deadline:
        if today.month == 12:
            next_month = datetime.date(today.year + 1, 1, deadline_day)
        else:
            next_month = datetime.date(today.year, today.month + 1, deadline_day)
        return next_month.strftime("%Y-%m-%d")

    return deadline.strftime("%Y-%m-%d")


# エクスポート用の変数
INPUT_DIR = None
OUTPUT_DIR = None
LOG_DIR = None
TEMPLATE_PATH = None
DEFAULT_EMPLOYEE_NAME = None

# 初期化時に値を設定
_config = init_config()
INPUT_DIR = _config.INPUT_DIR
OUTPUT_DIR = _config.OUTPUT_DIR
LOG_DIR = _config.LOG_DIR
TEMPLATE_PATH = _config.TEMPLATE_PATH
DEFAULT_EMPLOYEE_NAME = _config.EMPLOYEE_NAME
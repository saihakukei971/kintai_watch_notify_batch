#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
通知機能モジュール
"""
import os
import json
import smtplib
import csv
import time
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
from typing import List, Dict, Optional, Any, Union

import requests
from loguru import logger

from config import Config, init_config, get_deadline_date
from utils import find_latest_file


class Notifier:
    """通知クラス"""

    def __init__(self, config: Config = None):
        """
        初期化

        Args:
            config: 設定オブジェクト (Noneの場合は自動初期化)
        """
        self.config = config or init_config()

    def send_slack_notification(self, message: str, webhook_url: str = None) -> bool:
        """
        Slackに通知を送信

        Args:
            message: 送信するメッセージ
            webhook_url: Webhook URL (Noneの場合は設定から取得)

        Returns:
            bool: 送信成功かどうか
        """
        # Webhook URLの取得
        webhook_url = webhook_url or self.config.SLACK_WEBHOOK_URL

        if not webhook_url:
            logger.error("Slack Webhook URLが設定されていません")
            return False

        # リクエストデータの構築
        payload = {
            "text": message
        }

        try:
            # リクエストの送信
            response = requests.post(
                webhook_url,
                data=json.dumps(payload),
                headers={"Content-Type": "application/json"}
            )

            # レスポンスのチェック
            if response.status_code == 200 and response.text == "ok":
                logger.info("Slack通知を送信しました")
                return True
            else:
                logger.error(f"Slack通知の送信に失敗しました: {response.status_code} {response.text}")
                return False

        except Exception as e:
            logger.exception(f"Slack通知の送信エラー: {str(e)}")
            return False

    def send_email_notification(
        self,
        subject: str,
        message: str,
        to_email: Union[str, List[str]],
        from_email: str = None,
        smtp_server: str = None,
        smtp_port: int = 587,
        smtp_user: str = None,
        smtp_password: str = None
    ) -> bool:
        """
        メール通知を送信

        Args:
            subject: 件名
            message: 本文
            to_email: 宛先メールアドレス (文字列、またはリスト)
            from_email: 送信元メールアドレス
            smtp_server: SMTPサーバー
            smtp_port: SMTPポート
            smtp_user: SMTPユーザー名
            smtp_password: SMTPパスワード

        Returns:
            bool: 送信成功かどうか
        """
        # 設定からの取得
        smtp_server = smtp_server or os.getenv("SMTP_SERVER")
        smtp_port = smtp_port or int(os.getenv("SMTP_PORT", "587"))
        smtp_user = smtp_user or os.getenv("SMTP_USER")
        smtp_password = smtp_password or os.getenv("SMTP_PASSWORD")
        from_email = from_email or os.getenv("EMAIL_FROM")

        # 必須パラメータのチェック
        if not all([smtp_server, smtp_user, smtp_password, from_email]):
            logger.error("メール送信に必要な設定が不足しています")
            return False

        # 宛先のリスト化
        if isinstance(to_email, str):
            to_email = [to_email]

        try:
            # メッセージの作成
            msg = MIMEMultipart()
            msg["Subject"] = subject
            msg["From"] = from_email
            msg["To"] = ", ".join(to_email)

            # 本文を追加
            msg.attach(MIMEText(message, "plain"))

            # SMTPサーバーに接続
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                # TLSを有効化
                server.starttls()

                # ログイン
                server.login(smtp_user, smtp_password)

                # メールを送信
                server.send_message(msg)

            logger.info(f"メール通知を送信しました: {to_email}")
            return True

        except Exception as e:
            logger.exception(f"メール通知の送信エラー: {str(e)}")
            return False

    def check_submissions(
        self,
        members_file: str = None,
        deadline_date: str = None,
        days_before: int = None
    ) -> Dict[str, List[str]]:
        """
        提出状況を確認

        Args:
            members_file: 社員リストファイル (Noneの場合は設定から取得)
            deadline_date: 締切日 (Noneの場合は自動計算)
            days_before: 締切日までの日数 (Noneの場合は設定から取得)

        Returns:
            dict: 未提出者情報
        """
        # 設定値の取得
        members_file = members_file or os.path.join("config", "members.csv")

        # 締切日が指定されていない場合は自動計算
        if not deadline_date:
            deadline_date = get_deadline_date()
            logger.info(f"締切日を自動計算しました: {deadline_date}")

        days_before = days_before or self.config.REMIND_DAYS_BEFORE

        # 必須パラメータのチェック
        if not deadline_date:
            logger.error("締切日が設定されていません")
            return {"error": ["締切日が設定されていません"]}

        if not os.path.exists(members_file):
            logger.error(f"社員リストファイルが存在しません: {members_file}")
            return {"error": [f"社員リストファイルが存在しません: {members_file}"]}

        try:
            # 社員リストの読み込み
            members = self._read_members(members_file)

            # 締切日の解析
            deadline = datetime.strptime(deadline_date, "%Y-%m-%d").date()

            # 現在日付
            today = datetime.now().date()

            # 締切までの日数
            days_to_deadline = (deadline - today).days

            # 締切日を過ぎている場合
            if days_to_deadline < 0:
                logger.warning(f"締切日を過ぎています: {deadline_date}")
                return {"error": [f"締切日を過ぎています: {deadline_date}"]}

            # 提出状況の確認
            if days_to_deadline == 0 or days_to_deadline == days_before:
                # 提出ファイルの確認
                submitted = self._get_submitted_list()

                # 未提出者のリスト
                not_submitted = {}

                for member_id, member_info in members.items():
                    # 提出されているかチェック
                    if member_id not in submitted:
                        # 部署ごとにまとめる
                        department = member_info.get("department", "未所属")

                        if department not in not_submitted:
                            not_submitted[department] = []

                        not_submitted[department].append(f"{member_info['name']} ({member_id})")

                logger.info(f"未提出者: {sum(len(v) for v in not_submitted.values())}名")
                return not_submitted

            # 確認日でない場合
            logger.info(f"締切日まであと{days_to_deadline}日です (確認日: 締切当日 または {days_before}日前)")
            return {}

        except Exception as e:
            logger.exception(f"提出状況の確認エラー: {str(e)}")
            return {"error": [str(e)]}

    def send_reminder(self, not_submitted: Dict[str, List[str]], deadline_date: str = None) -> bool:
        """
        リマインダーを送信

        Args:
            not_submitted: 未提出者情報
            deadline_date: 締切日 (Noneの場合は自動計算)

        Returns:
            bool: 送信成功かどうか
        """
        if not not_submitted or list(not_submitted.keys()) == ["error"]:
            logger.info("送信すべきリマインダーはありません")
            return False

        # 締切日の取得
        if not deadline_date:
            deadline_date = get_deadline_date()

        # 通知済みチェック（部署別）
        filtered_not_submitted = {}

        for department, members in not_submitted.items():
            filtered_members = []

            for member in members:
                # 社員IDを抽出（例: "山田太郎 (1001)" → "1001"）
                import re
                match = re.search(r'\(([^)]+)\)', member)
                if match:
                    member_id = match.group(1)
                    # 24時間以内に通知していない場合のみリストに追加
                    if not self._check_notification_history(member_id, "reminder"):
                        filtered_members.append(member)
                else:
                    # IDが取得できない場合はそのまま追加
                    filtered_members.append(member)

            if filtered_members:
                filtered_not_submitted[department] = filtered_members

        # 通知すべきメンバーがいるかチェック
        if not filtered_not_submitted:
            logger.info("通知すべきメンバーがいません（全員24時間以内に通知済み）")
            return False

        # メッセージの構築
        message = f"⚠️【提出リマインド】⚠️\n\n"
        message += f"以下の方は提出が確認できていません（締切：{deadline_date}）\n\n"

        for department, members in filtered_not_submitted.items():
            message += f"【{department}】\n"
            for member in members:
                message += f"・{member}\n"
            message += "\n"

        message += f"期日までに input/ にCSVを配置してください。"

        # 送信フラグ
        sent = False

        # Slackに送信
        if self.config.SLACK_WEBHOOK_URL:
            slack_result = self.send_slack_notification(message)
            if slack_result:
                sent = True

                # 通知履歴を記録
                for department, members in filtered_not_submitted.items():
                    for member in members:
                        match = re.search(r'\(([^)]+)\)', member)
                        if match:
                            member_id = match.group(1)
                            self._update_notification_history(member_id, "reminder")

        # メールアドレスが設定されていれば送信
        email_to = os.getenv("EMAIL_TO")
        if email_to:
            email_result = self.send_email_notification(
                f"【提出リマインド】締切：{deadline_date}",
                message,
                email_to
            )
            if email_result:
                sent = True

                # 通知履歴を記録（Slackで既に記録していない場合）
                if not self.config.SLACK_WEBHOOK_URL:
                    for department, members in filtered_not_submitted.items():
                        for member in members:
                            match = re.search(r'\(([^)]+)\)', member)
                            if match:
                                member_id = match.group(1)
                                self._update_notification_history(member_id, "reminder")

        if not sent:
            logger.warning("通知先（SlackまたはEmail）が設定されていないか、送信に失敗しました")
            return False

        return True

    def _read_members(self, members_file: str) -> Dict[str, Dict[str, str]]:
        """
        社員リストを読み込み

        Args:
            members_file: 社員リストファイル

        Returns:
            dict: 社員情報
        """
        members = {}

        try:
            with open(members_file, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # IDをキーに
                    member_id = row.get("id", "").strip()

                    if not member_id:
                        continue

                    members[member_id] = {
                        "name": row.get("name", "").strip(),
                        "department": row.get("department", "").strip(),
                        "email": row.get("email", "").strip(),
                    }

            logger.info(f"社員リストを読み込みました: {len(members)}名")
            return members

        except Exception as e:
            logger.exception(f"社員リストの読み込みエラー: {str(e)}")
            return {}

    def _get_submitted_list(self) -> List[str]:
        """
        提出済みリストを取得

        Returns:
            list: 提出済みの社員ID
        """
        submitted = []
        input_dir = self.config.INPUT_DIR

        try:
            # 入力ディレクトリ内のファイルを走査
            for filename in os.listdir(input_dir):
                if filename.endswith(".csv"):
                    # ファイル名から社員IDを抽出
                    import re

                    # パターン: 勤怠詳細_YYYYMM_ID.csv
                    match = re.search(r'勤怠詳細_\d{6}_(.+?)\.csv', filename)
                    if match:
                        submitted.append(match.group(1))
                        continue

                    # パターン: 勤怠詳細_ID_YYYY_MM.csv
                    match = re.search(r'勤怠詳細_(.+?)_\d{4}_\d{2}\.csv', filename)
                    if match:
                        submitted.append(match.group(1))
                        continue

            logger.info(f"提出済みリスト: {len(submitted)}名")
            return submitted

        except Exception as e:
            logger.exception(f"提出済みリストの取得エラー: {str(e)}")
            return []

    def _check_notification_history(self, member_id: str, notification_type: str) -> bool:
        """
        過去24時間以内に通知を送信したかチェック

        Args:
            member_id: 社員ID
            notification_type: 通知タイプ

        Returns:
            bool: 通知済みならTrue
        """
        history_file = os.path.join(self.config.LOG_DIR, "notification_history.json")
        current_time = time.time()
        history = {}

        # 履歴ファイルの読み込み
        if os.path.exists(history_file):
            try:
                with open(history_file, 'r') as f:
                    history = json.load(f)
            except:
                logger.error("通知履歴ファイルの読み込みに失敗しました")

        # 古い履歴の削除（24時間以上前）
        history = {k: v for k, v in history.items() if current_time - v < 86400}

        # 通知履歴のチェック
        key = f"{member_id}_{notification_type}"
        if key in history:
            logger.info(f"{member_id}には24時間以内に通知済み")
            return True

        return False

    def _update_notification_history(self, member_id: str, notification_type: str) -> bool:
        """
        通知履歴を更新

        Args:
            member_id: 社員ID
            notification_type: 通知タイプ

        Returns:
            bool: 更新成功かどうか
        """
        history_file = os.path.join(self.config.LOG_DIR, "notification_history.json")
        current_time = time.time()
        history = {}

        # 履歴ファイルの読み込み
        if os.path.exists(history_file):
            try:
                with open(history_file, 'r') as f:
                    history = json.load(f)
            except:
                logger.error("通知履歴ファイルの読み込みに失敗しました")

        # 古い履歴の削除（24時間以上前）
        history = {k: v for k, v in history.items() if current_time - v < 86400}

        # 履歴の更新
        key = f"{member_id}_{notification_type}"
        history[key] = current_time

        try:
            os.makedirs(os.path.dirname(history_file), exist_ok=True)
            with open(history_file, 'w') as f:
                json.dump(history, f)
            logger.info(f"通知履歴を更新しました: {member_id}")
            return True
        except:
            logger.error("通知履歴ファイルの書き込みに失敗しました")
            return False


def check_and_remind():
    """締切状況をチェックしてリマインドを送信"""
    try:
        notifier = Notifier()
        # 締切日は自動計算
        deadline_date = get_deadline_date()
        not_submitted = notifier.check_submissions(deadline_date=deadline_date)

        if not_submitted and list(not_submitted.keys()) != ["error"]:
            notifier.send_reminder(not_submitted, deadline_date)
            logger.info("リマインダーを送信しました")

        return True

    except Exception as e:
        logger.exception(f"チェック＆リマインドエラー: {str(e)}")
        return False


if __name__ == "__main__":
    from utils import setup_logging

    # ロギング設定
    setup_logging()

    # チェック＆リマインド実行
    check_and_remind()
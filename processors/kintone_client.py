#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
kintone API連携クライアント
"""
import os
import json
import base64
from typing import Dict, List, Any, Optional
import csv
from datetime import datetime

import requests
import pandas as pd
from loguru import logger

from utils import safe_filename


class KintoneClient:
    """kintone APIクライアント"""

    def __init__(self, domain: str, api_token: str = None, username: str = None, password: str = None):
        """
        初期化

        Args:
            domain: kintoneドメイン (例: example.cybozu.com)
            api_token: APIトークン (優先的に使用)
            username: ユーザー名 (APIトークンがない場合)
            password: パスワード (APIトークンがない場合)
        """
        self.domain = domain.rstrip('/')
        if not self.domain.startswith('https://'):
            self.domain = f"https://{self.domain}"

        self.api_token = api_token
        self.username = username
        self.password = password

        # APIのベースURL
        self.base_url = f"{self.domain}/k/v1"

        # アプリID Cache
        self.app_id_cache = {}

        logger.info(f"kintoneクライアントを初期化しました: {self.domain}")

    def _get_headers(self) -> Dict[str, str]:
        """
        API呼び出し用のヘッダーを取得

        Returns:
            dict: ヘッダー情報
        """
        headers = {
            "Content-Type": "application/json",
            "X-HTTP-Method-Override": "GET"
        }

        # API トークン認証
        if self.api_token:
            headers["X-Cybozu-API-Token"] = self.api_token

        # Basic 認証
        elif self.username and self.password:
            auth = f"{self.username}:{self.password}"
            encoded_auth = base64.b64encode(auth.encode()).decode()
            headers["Authorization"] = f"Basic {encoded_auth}"

        else:
            logger.warning("認証情報が設定されていません")

        return headers

    def get_app_id(self, app_name: str) -> Optional[str]:
        """
        アプリ名からアプリIDを取得

        Args:
            app_name: アプリ名

        Returns:
            str or None: アプリID (見つからない場合はNone)
        """
        # キャッシュからの取得を試みる
        if app_name in self.app_id_cache:
            return self.app_id_cache[app_name]

        # アプリ一覧を取得
        url = f"{self.base_url}/apps.json"
        headers = self._get_headers()

        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()

            apps = response.json().get("apps", [])

            # アプリ名でフィルタリング
            for app in apps:
                if app.get("name") == app_name:
                    app_id = str(app.get("appId"))
                    # キャッシュに保存
                    self.app_id_cache[app_name] = app_id
                    logger.info(f"アプリIDを取得しました: {app_name} (ID: {app_id})")
                    return app_id

            logger.warning(f"アプリ名に一致するアプリIDが見つかりません: {app_name}")
            return None

        except Exception as e:
            logger.exception(f"アプリID取得エラー: {str(e)}")
            return None

    def get_records(self, app_name: str, query: str = "", fields: List[str] = None, max_records: int = 500) -> List[Dict[str, Any]]:
        """
        レコードを取得

        Args:
            app_name: アプリ名
            query: クエリ文字列
            fields: 取得するフィールド名のリスト
            max_records: 最大取得レコード数

        Returns:
            list: レコードのリスト
        """
        # アプリIDを取得
        app_id = self.get_app_id(app_name)
        if not app_id:
            logger.error(f"アプリIDが取得できませんでした: {app_name}")
            return []

        # APIエンドポイント
        url = f"{self.base_url}/records.json"
        headers = self._get_headers()

        # クエリパラメータの構築
        params = {
            "app": app_id,
            "totalCount": True
        }

        if query:
            params["query"] = query

        if fields:
            params["fields"] = fields

        all_records = []
        offset = 0
        limit = min(max_records, 500)  # 1回のリクエストで最大500件

        try:
            while True:
                # オフセットとリミットの設定
                current_query = f"{query} limit {limit} offset {offset}" if query else f"limit {limit} offset {offset}"
                params["query"] = current_query

                # API呼び出し
                response = requests.get(url, headers=headers, params=params)
                response.raise_for_status()

                data = response.json()
                records = data.get("records", [])

                if not records:
                    break

                all_records.extend(records)

                # 次のページがあるかチェック
                offset += limit
                if offset >= data.get("totalCount", 0) or len(all_records) >= max_records:
                    break

            logger.info(f"{len(all_records)}件のレコードを取得しました: {app_name}")
            return all_records

        except Exception as e:
            logger.exception(f"レコード取得エラー: {str(e)}")
            return []

    def add_records(self, app_name: str, records: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        レコードを追加

        Args:
            app_name: アプリ名
            records: 追加するレコードのリスト

        Returns:
            dict: API応答
        """
        if not records:
            logger.warning("追加するレコードがありません")
            return {"success": True, "message": "追加するレコードがありません"}

        # アプリIDを取得
        app_id = self.get_app_id(app_name)
        if not app_id:
            logger.error(f"アプリIDが取得できませんでした: {app_name}")
            return {"success": False, "message": f"アプリIDが取得できませんでした: {app_name}"}

        # APIエンドポイント
        url = f"{self.base_url}/records.json"
        headers = self._get_headers()

        # リクエストデータの構築
        req_data = {
            "app": app_id,
            "records": records
        }

        try:
            # API呼び出し
            response = requests.post(url, headers=headers, data=json.dumps(req_data))
            response.raise_for_status()

            result = response.json()
            logger.info(f"{len(records)}件のレコードを追加しました: {app_name}")

            return {"success": True, "data": result}

        except Exception as e:
            logger.exception(f"レコード追加エラー: {str(e)}")
            return {"success": False, "message": str(e)}

    def update_records(self, app_name: str, records: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        レコードを更新

        Args:
            app_name: アプリ名
            records: 更新するレコードのリスト (各レコードにはIDと更新項目を含む)

        Returns:
            dict: API応答
        """
        if not records:
            logger.warning("更新するレコードがありません")
            return {"success": True, "message": "更新するレコードがありません"}

        # アプリIDを取得
        app_id = self.get_app_id(app_name)
        if not app_id:
            logger.error(f"アプリIDが取得できませんでした: {app_name}")
            return {"success": False, "message": f"アプリIDが取得できませんでした: {app_name}"}

        # APIエンドポイント
        url = f"{self.base_url}/records.json"
        headers = self._get_headers()
        headers["X-HTTP-Method-Override"] = "PUT"

        # リクエストデータの構築
        req_data = {
            "app": app_id,
            "records": records
        }

        try:
            # API呼び出し
            response = requests.put(url, headers=headers, data=json.dumps(req_data))
            response.raise_for_status()

            result = response.json()
            logger.info(f"{len(records)}件のレコードを更新しました: {app_name}")

            return {"success": True, "data": result}

        except Exception as e:
            logger.exception(f"レコード更新エラー: {str(e)}")
            return {"success": False, "message": str(e)}

    def delete_records(self, app_name: str, record_ids: List[str]) -> Dict[str, Any]:
        """
        レコードを削除

        Args:
            app_name: アプリ名
            record_ids: 削除するレコードIDのリスト

        Returns:
            dict: API応答
        """
        if not record_ids:
            logger.warning("削除するレコードがありません")
            return {"success": True, "message": "削除するレコードがありません"}

        # アプリIDを取得
        app_id = self.get_app_id(app_name)
        if not app_id:
            logger.error(f"アプリIDが取得できませんでした: {app_name}")
            return {"success": False, "message": f"アプリIDが取得できませんでした: {app_name}"}

        # APIエンドポイント
        url = f"{self.base_url}/records.json"
        headers = self._get_headers()
        headers["X-HTTP-Method-Override"] = "DELETE"

        # リクエストデータの構築
        req_data = {
            "app": app_id,
            "ids": record_ids
        }

        try:
            # API呼び出し
            response = requests.delete(url, headers=headers, data=json.dumps(req_data))
            response.raise_for_status()

            logger.info(f"{len(record_ids)}件のレコードを削除しました: {app_name}")
            return {"success": True}

        except Exception as e:
            logger.exception(f"レコード削除エラー: {str(e)}")
            return {"success": False, "message": str(e)}

    def save_as_csv(self, records: List[Dict[str, Any]], output_file: str, encoding: str = 'utf-8') -> bool:
        """
        レコードをCSVとして保存

        Args:
            records: レコードのリスト
            output_file: 出力ファイルパス
            encoding: エンコーディング

        Returns:
            bool: 保存成功かどうか
        """
        if not records:
            logger.warning("保存するレコードがありません")
            return False

        try:
            # 出力ディレクトリの確認
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
                logger.info(f"出力ディレクトリを作成しました: {output_dir}")

            # フィールド名を取得
            first_record = records[0]
            fieldnames = []

            # 全フィールドを取得
            for field_name, field_data in first_record.items():
                # 通常のフィールド
                if isinstance(field_data, dict) and "value" in field_data:
                    fieldnames.append(field_name)

            # DataFrameに変換
            data_list = []
            for record in records:
                row_data = {}
                for field_name in fieldnames:
                    if field_name in record and "value" in record[field_name]:
                        row_data[field_name] = record[field_name]["value"]
                    else:
                        row_data[field_name] = ""
                data_list.append(row_data)

            # CSVとして保存
            df = pd.DataFrame(data_list)
            df.to_csv(output_file, encoding=encoding, index=False)

            logger.info(f"CSVファイルを保存しました: {output_file}")
            return True

        except Exception as e:
            logger.exception(f"CSV保存エラー: {str(e)}")
            return False

    def csv_to_records(self, csv_file: str, encoding: str = None) -> List[Dict[str, Any]]:
        """
        CSVファイルをkintoneレコード形式に変換

        Args:
            csv_file: CSVファイルパス
            encoding: エンコーディング (自動検出する場合はNone)

        Returns:
            list: kintoneレコード形式のリスト
        """
        try:
            # ファイルの存在確認
            if not os.path.exists(csv_file):
                logger.error(f"ファイルが存在しません: {csv_file}")
                return []

            # エンコーディングを自動検出
            if not encoding:
                from utils import detect_csv_encoding
                encoding = detect_csv_encoding(csv_file)
                logger.info(f"CSVエンコーディング: {encoding}")

            # CSVを読み込み
            df = pd.read_csv(csv_file, encoding=encoding)

            # kintoneレコード形式に変換
            records = []
            for _, row in df.iterrows():
                record = {}
                for col in df.columns:
                    value = row[col]

                    # NaN値はスキップ
                    if pd.isna(value):
                        continue

                    # 日付型判定
                    if isinstance(value, pd.Timestamp):
                        value = value.strftime("%Y-%m-%d")

                    # 辞書形式に変換
                    record[col] = {"value": str(value)}

                records.append(record)

            logger.info(f"{len(records)}件のレコードに変換しました: {csv_file}")
            return records

        except Exception as e:
            logger.exception(f"CSV変換エラー: {str(e)}")
            return []
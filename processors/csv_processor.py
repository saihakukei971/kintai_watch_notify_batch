#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CSV処理モジュール
"""
import os
import pandas as pd
from datetime import datetime
from pathlib import Path
from loguru import logger
from tqdm import tqdm

from utils import detect_csv_encoding, parse_time_str


def read_csv(csv_path: str) -> pd.DataFrame:
    """
    CSVファイルを読み込み、DataFrameとして返す

    Args:
        csv_path: CSVファイルのパス

    Returns:
        DataFrame: 読み込んだデータ

    Raises:
        FileNotFoundError: ファイルが存在しない場合
        ValueError: CSVの形式が不正な場合
    """
    # ファイルの存在確認
    if not os.path.exists(csv_path):
        logger.error(f"ファイルが存在しません: {csv_path}")
        raise FileNotFoundError(f"ファイルが存在しません: {csv_path}")

    # エンコーディングを自動検出
    encoding = detect_csv_encoding(csv_path)
    logger.info(f"CSVエンコーディング: {encoding}")

    try:
        # CSVを読み込み
        df = pd.read_csv(csv_path, encoding=encoding)

        # 必須カラムの確認
        required_columns = ["日付", "始業時刻", "終業時刻", "総勤務時間"]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            logger.error(f"必須カラムがありません: {', '.join(missing_columns)}")
            raise ValueError(f"必須カラムがありません: {', '.join(missing_columns)}")

        # 日付カラムの型変換
        try:
            df["日付"] = pd.to_datetime(df["日付"])
        except Exception as e:
            logger.error(f"日付カラムの変換に失敗しました: {str(e)}")
            raise ValueError(f"日付カラムの変換に失敗しました: {str(e)}")

        # データ行数のログ
        logger.info(f"CSVデータ読み込み完了: {len(df)}行")

        return df

    except Exception as e:
        logger.exception(f"CSVファイルの読み込みに失敗しました: {str(e)}")
        raise


def process_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    勤怠データを整形する

    Args:
        df: 元のDataFrame

    Returns:
        DataFrame: 整形後のDataFrame
    """
    logger.info("データ処理を開始します")

    # カラム一覧をログ
    logger.debug(f"入力データのカラム: {df.columns.tolist()}")

    # 必要なカラムの選択
    columns_to_keep = [
        "日付", "始業時刻", "終業時刻", "総勤務時間",
        "法定内残業", "時間外労働", "深夜労働", "勤怠種別"
    ]

    # 存在するカラムのみを使用
    available_columns = [col for col in columns_to_keep if col in df.columns]

    # 存在しないカラムをログ
    missing_columns = [col for col in columns_to_keep if col not in df.columns]
    if missing_columns:
        logger.warning(f"以下のカラムが存在しないため、処理から除外します: {', '.join(missing_columns)}")

    # データをコピー
    df_filtered = df[available_columns].copy()

    # 勤怠種別の欠損値を処理
    if "勤怠種別" in df_filtered.columns:
        df_filtered["勤怠種別"] = df_filtered["勤怠種別"].fillna("未入力")
    else:
        df_filtered["勤怠種別"] = "通常勤務"

    # 時間を小数時間に変換
    logger.info("時間データを変換しています")

    # 進捗バーに対応
    tqdm.pandas(desc="時間変換")

    # 時間フィールドの変換
    time_fields = ["総勤務時間", "法定内残業", "時間外労働", "深夜労働"]
    for field in time_fields:
        if field in df_filtered.columns:
            df_filtered[field] = df_filtered[field].progress_apply(
                lambda x: parse_time_str(x) / 3600 if pd.notnull(x) else 0
            )

    # 始業・終業時刻の欠損値を処理
    if "始業時刻" in df_filtered.columns:
        df_filtered["始業時刻"] = df_filtered["始業時刻"].fillna("")

    if "終業時刻" in df_filtered.columns:
        df_filtered["終業時刻"] = df_filtered["終業時刻"].fillna("")

    # 必須フィールドの存在確認
    for field in ["日付", "総勤務時間"]:
        if field not in df_filtered.columns:
            logger.error(f"必須フィールドがありません: {field}")
            raise ValueError(f"必須フィールドがありません: {field}")

    # データをログ
    logger.debug(f"処理後のデータ: {len(df_filtered)}行")

    return df_filtered


def save_csv(df: pd.DataFrame, output_path: str, encoding: str = 'utf-8'):
    """
    DataFrameをCSVとして保存

    Args:
        df: 保存するDataFrame
        output_path: 出力先パス
        encoding: エンコーディング
    """
    try:
        # 出力ディレクトリが存在しない場合は作成
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            logger.info(f"出力ディレクトリを作成しました: {output_dir}")

        # CSVとして保存
        df.to_csv(output_path, encoding=encoding, index=False)
        logger.info(f"CSVファイルを保存しました: {output_path}")

        return True

    except Exception as e:
        logger.exception(f"CSVファイルの保存に失敗しました: {str(e)}")
        raise
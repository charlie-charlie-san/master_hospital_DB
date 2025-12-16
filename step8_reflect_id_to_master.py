#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 8: マスターDBへの卸ID反映ツール (step8_reflect_id_to_master.py)
====================================================================
id_mapping.xlsx の紐付け情報をマスターDB (master_db.xlsx) に反映し、
「卸側施設ID」カラムを追加・更新するツール。

【処理フロー】
1. マスターDBの自動バックアップ
2. id_mapping.xlsx から卸側施設IDを取得
3. 同一UIDに複数の卸IDがある場合はカンマ区切りで結合
4. マスターDBにマージして保存
5. 整合性検証

保存先: ~/Desktop/hospital_DB/2_Storage/
バックアップ: ~/Desktop/hospital_DB/9_Archives/
"""

import pandas as pd
import os
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

# ==========================================
# 設定
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
STORAGE_DIR = os.path.join(BASE_DIR, "2_Storage")
ARCHIVE_DIR = os.path.join(BASE_DIR, "9_Archives")

# 対象ファイル
MASTER_FILE = os.path.join(STORAGE_DIR, "master_db.xlsx")
MAPPING_FILE = os.path.join(STORAGE_DIR, "id_mapping.xlsx")

# 卸業者名（このツールで扱う卸）
FIXED_WHOLESALER_NAME = "アスコ"

# 追加するカラム名
NEW_COL_NAME = "卸側施設ID"


# ==========================================
# ユーティリティ関数
# ==========================================
def find_uid_column(df):
    """
    UIDカラムを探す（複数パターン対応）
    """
    uid_patterns = ["自社UID", "施設UID", "UID", "自社ID", "動物病院UID"]
    
    for col in df.columns:
        col_str = str(col).strip()
        if col_str in uid_patterns:
            return col
        if "UID" in col_str.upper():
            return col
    return None


def clean_value(val):
    """
    nan/None/空文字を空文字に統一
    """
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", "null", "nat"]:
        return ""
    return val_str


def clean_id(val):
    """
    IDを文字列として整形（小数点.0を除去）
    例: 12345.0 → "12345"
    """
    if pd.isna(val):
        return ""
    
    # 数値型の場合は整数に変換してから文字列化
    if isinstance(val, float):
        # 小数点以下が0なら整数として扱う
        if val == int(val):
            return str(int(val))
        else:
            return str(val)
    
    val_str = str(val).strip()
    
    if val_str.lower() in ["nan", "none", "null", "nat", ""]:
        return ""
    
    # 文字列でも ".0" で終わっている場合は除去
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    
    return val_str


def verify_backup(original_path, backup_path):
    """
    バックアップが正常に作成されたか検証
    """
    if not os.path.exists(backup_path):
        return False
    
    original_size = os.path.getsize(original_path)
    backup_size = os.path.getsize(backup_path)
    
    # サイズが同じなら成功とみなす
    return original_size == backup_size


# ==========================================
# メイン処理 (Step 8)
# ==========================================
def step8_reflect_id_to_master():
    """
    マスターDBに卸側施設IDを反映する
    """
    root = None
    backup_path = None
    
    try:
        # tkinter初期化
        root = tk.Tk()
        root.withdraw()
        
        print("=" * 60)
        print("Step 8: マスターDBへの卸ID反映ツール")
        print("=" * 60)

        # --------------------------------------
        # 1. ファイル存在チェック
        # --------------------------------------
        if not os.path.exists(MASTER_FILE):
            messagebox.showerror("エラー", f"マスターDBが見つかりません。\n{MASTER_FILE}")
            return
        
        if not os.path.exists(MAPPING_FILE):
            messagebox.showerror("エラー", f"マッピングファイルが見つかりません。\n{MAPPING_FILE}")
            return

        # --------------------------------------
        # 2. 自動バックアップ
        # --------------------------------------
        print("\n📦 バックアップを作成中...")
        
        if not os.path.exists(ARCHIVE_DIR):
            os.makedirs(ARCHIVE_DIR)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(ARCHIVE_DIR, f"master_db_backup_{timestamp}.xlsx")
        
        shutil.copy2(MASTER_FILE, backup_path)
        
        # バックアップ検証
        if not verify_backup(MASTER_FILE, backup_path):
            messagebox.showerror("エラー", "バックアップの作成に失敗しました。処理を中止します。")
            return
        
        print(f"   ✓ {os.path.basename(backup_path)}")

        # --------------------------------------
        # 3. データの読み込み
        # --------------------------------------
        print("\n📖 ファイルを読み込み中...")
        
        df_master = pd.read_excel(MASTER_FILE, engine='openpyxl')
        df_mapping = pd.read_excel(MAPPING_FILE, engine='openpyxl')
        
        master_count = len(df_master)
        mapping_count = len(df_mapping)
        
        print(f"   マスターDB:    {master_count:,} 行")
        print(f"   IDマッピング:  {mapping_count:,} 行")

        # --------------------------------------
        # 4. UIDカラムの検出
        # --------------------------------------
        master_uid_col = find_uid_column(df_master)
        mapping_uid_col = find_uid_column(df_mapping)
        
        if not master_uid_col:
            messagebox.showerror("エラー", "マスターDBにUID列が見つかりません。")
            return
        
        if not mapping_uid_col:
            messagebox.showerror("エラー", "マッピングファイルにUID列が見つかりません。")
            return
        
        print(f"\n   マスターUID列:    {master_uid_col}")
        print(f"   マッピングUID列:  {mapping_uid_col}")

        # --------------------------------------
        # 5. マッピングデータの整理
        # --------------------------------------
        print("\n⚙️  マッピングデータを整理中...")
        
        # 必須カラムチェック
        if "卸側施設ID" not in df_mapping.columns:
            messagebox.showerror("エラー", "マッピングファイルに「卸側施設ID」列がありません。")
            return

        # 卸業者名でフィルタ（指定された卸のみ）
        if "卸業者名" in df_mapping.columns:
            df_mapping_filtered = df_mapping[
                df_mapping["卸業者名"].apply(clean_value) == FIXED_WHOLESALER_NAME
            ].copy()
            print(f"   卸業者「{FIXED_WHOLESALER_NAME}」でフィルタ: {len(df_mapping_filtered):,} 行")
        else:
            df_mapping_filtered = df_mapping.copy()
            print(f"   ※卸業者名列がないため全件対象: {len(df_mapping_filtered):,} 行")

        # クリーニング（卸側施設IDは小数点除去）
        df_mapping_filtered[mapping_uid_col] = df_mapping_filtered[mapping_uid_col].apply(clean_value)
        df_mapping_filtered["卸側施設ID"] = df_mapping_filtered["卸側施設ID"].apply(clean_id)
        
        # 空のデータを除外
        df_map_clean = df_mapping_filtered[
            (df_mapping_filtered[mapping_uid_col] != "") & 
            (df_mapping_filtered["卸側施設ID"] != "")
        ].copy()
        
        print(f"   有効なマッピング: {len(df_map_clean):,} 行")

        # グルーピング（同一UIDに複数の卸IDがある場合はカンマ区切り）
        df_grouped = df_map_clean.groupby(mapping_uid_col)["卸側施設ID"].apply(
            lambda x: ", ".join(sorted(set(x)))
        ).reset_index()
        
        # カラム名を統一
        df_grouped = df_grouped.rename(columns={
            mapping_uid_col: master_uid_col,
            "卸側施設ID": NEW_COL_NAME
        })
        
        unique_uid_count = len(df_grouped)
        print(f"   ユニークUID数: {unique_uid_count:,} 件")

        # --------------------------------------
        # 6. マスターDBへのマージ
        # --------------------------------------
        print("\n⚙️  マスターDBに結合中...")
        
        # 既存の卸側施設ID列があれば削除（更新のため）
        if NEW_COL_NAME in df_master.columns:
            print(f"   ※既存の「{NEW_COL_NAME}」列を更新します")
            df_master = df_master.drop(columns=[NEW_COL_NAME])

        # マージ（Left Join）
        df_merged = pd.merge(
            df_master, 
            df_grouped, 
            on=master_uid_col, 
            how="left"
        )
        
        # NaNを空文字に
        df_merged[NEW_COL_NAME] = df_merged[NEW_COL_NAME].fillna("")

        # --------------------------------------
        # 7. 整合性検証
        # --------------------------------------
        print("\n🔍 整合性を検証中...")
        
        # 行数が変わっていないことを確認
        if len(df_merged) != master_count:
            messagebox.showerror(
                "エラー", 
                f"マージ後の行数が変化しました。\n"
                f"処理前: {master_count} → 処理後: {len(df_merged)}\n"
                f"バックアップから復元してください。"
            )
            return
        
        print(f"   ✓ 行数: {len(df_merged):,} 行（変化なし）")
        
        # カラム数の確認
        expected_cols = len(df_master.columns) + 1
        actual_cols = len(df_merged.columns)
        if actual_cols != expected_cols:
            print(f"   ⚠️ カラム数: {actual_cols} (想定: {expected_cols})")
        else:
            print(f"   ✓ カラム数: {actual_cols} 列")

        # --------------------------------------
        # 8. 保存
        # --------------------------------------
        print("\n💾 保存中...")
        df_merged.to_excel(MASTER_FILE, index=False, engine='openpyxl')
        print(f"   ✓ {os.path.basename(MASTER_FILE)}")

        # --------------------------------------
        # 9. 結果サマリー
        # --------------------------------------
        mapped_count = len(df_merged[df_merged[NEW_COL_NAME] != ""])
        unmapped_count = master_count - mapped_count
        coverage_rate = (mapped_count / master_count * 100) if master_count > 0 else 0
        
        print("\n" + "=" * 60)
        print("【処理結果サマリー】")
        print("=" * 60)
        print(f"  マスターDB総数:     {master_count:,} 件")
        print(f"  マッピング総数:     {mapping_count:,} 件")
        print(f"  → 有効マッピング:  {len(df_map_clean):,} 件")
        print(f"  → ユニークUID:     {unique_uid_count:,} 件")
        print("-" * 60)
        print(f"  紐付け成功:         {mapped_count:,} 件")
        print(f"  未紐付け:           {unmapped_count:,} 件")
        print(f"  カバー率:           {coverage_rate:.1f}%")
        print("=" * 60)
        print(f"\n📁 保存先: {MASTER_FILE}")
        print(f"📦 バックアップ: {backup_path}")

        # 完了メッセージ
        msg = (
            f"✅ マスターDBの更新が完了しました！\n\n"
            f"【結果】\n"
            f"・紐付け成功: {mapped_count:,} 件\n"
            f"・未紐付け: {unmapped_count:,} 件\n"
            f"・カバー率: {coverage_rate:.1f}%\n\n"
            f"カラム「{NEW_COL_NAME}」を追加・更新しました。"
        )
        messagebox.showinfo("成功", msg)
        
        if os.name == 'nt':
            os.startfile(STORAGE_DIR)
        else:
            os.system(f"open '{STORAGE_DIR}'")

    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        
        # バックアップからの復元案内
        restore_msg = ""
        if backup_path and os.path.exists(backup_path):
            restore_msg = f"\n\nバックアップファイル:\n{backup_path}\n\nこのファイルから復元できます。"
        
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n\n{e}{restore_msg}")
        
    finally:
        if root:
            root.destroy()


if __name__ == "__main__":
    step8_reflect_id_to_master()

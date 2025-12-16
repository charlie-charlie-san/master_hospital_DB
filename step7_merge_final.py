#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 7: マッピングデータ統合ツール (step7_merge_final.py)
============================================================
「自動マッチ成功分」と「手動でUIDを埋めた分」を統合し、
最終的な id_mapping.xlsx を生成・更新するツール。

【処理フロー】
1. 自動成功リスト (id_mapping_candidate.xlsx) を読込
2. 手動補完リスト (unmatched_list.xlsx) からUID入力済みの行を抽出
3. 両者を統合
4. 既存のid_mapping.xlsxがあれば追記・重複排除
5. 最終結果を保存

保存先: ~/Desktop/hospital_DB/2_Storage/
"""

import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox

# ==========================================
# 設定
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
STORAGE_DIR = os.path.join(BASE_DIR, "2_Storage")

# 入力ファイル
CANDIDATE_FILE = os.path.join(STORAGE_DIR, "id_mapping_candidate.xlsx")  # 自動成功分
UNMATCHED_FILE = os.path.join(STORAGE_DIR, "unmatched_list.xlsx")        # 手動補完分

# 出力ファイル（最終的な正解データ）
FINAL_MAPPING_FILE = os.path.join(STORAGE_DIR, "id_mapping.xlsx")

# 固定値（補完用）
FIXED_WHOLESALER_NAME = "アスコ"

# 最終出力カラム
FINAL_COLUMNS = ["自社UID", "施設名(確認用)", "卸業者名", "卸側施設ID", "適用開始日"]


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


def standardize_columns(df, source_name="データ"):
    """
    データフレームのカラム名を id_mapping.xlsx の形式に統一する
    重複カラムや複雑なインデックスを安全に処理する
    """
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=FINAL_COLUMNS)

    # 元のDataFrameをコピーしてインデックスをリセット
    df = df.copy()
    df = df.reset_index(drop=True)
    
    # カラム名を文字列に変換（念のため）
    df.columns = [str(c) for c in df.columns]

    # 1. UIDカラムを探す（複数パターン）
    uid_col = None
    uid_patterns = ["自社UID", "施設UID", "UID", "自社ID", "動物病院UID"]
    for col in df.columns:
        if col in uid_patterns or "UID" in col.upper():
            uid_col = col
            break
    
    if uid_col is None:
        print(f"   ⚠️ {source_name}: 'UID' 列が見つかりません。スキップします。")
        return pd.DataFrame(columns=FINAL_COLUMNS)

    # 2. 卸側施設IDを探す
    wholesaler_id_col = None
    for col in ["卸側施設ID", "得意先コード"]:
        if col in df.columns:
            wholesaler_id_col = col
            break

    # 3. 施設名を探す
    name_col = None
    for col in ["施設名(確認用)", "得意先名称", "卸側名称(参考)", "動物病院施設名"]:
        if col in df.columns:
            name_col = col
            break

    # 4. 新しいDataFrameを作成（カラムを1つずつ安全に追加）
    result_rows = []
    
    for idx in range(len(df)):
        row = df.iloc[idx]
        
        # UID取得・クリーン
        uid_val = clean_value(row[uid_col]) if uid_col else ""
        
        # UIDが空ならスキップ
        if uid_val == "":
            continue
        
        # 各カラムの値を取得
        new_row = {
            "自社UID": uid_val,
            "施設名(確認用)": clean_value(row[name_col]) if name_col else "",
            "卸業者名": clean_value(row.get("卸業者名", FIXED_WHOLESALER_NAME)) or FIXED_WHOLESALER_NAME,
            "卸側施設ID": clean_value(row[wholesaler_id_col]) if wholesaler_id_col else "",
            "適用開始日": clean_value(row.get("適用開始日", ""))
        }
        result_rows.append(new_row)
    
    if not result_rows:
        print(f"   ⚠️ {source_name}: UIDが入力されている行がありません。")
        return pd.DataFrame(columns=FINAL_COLUMNS)
    
    return pd.DataFrame(result_rows, columns=FINAL_COLUMNS)


# ==========================================
# メイン処理 (Step 7)
# ==========================================
def step7_merge_final():
    """
    マッピングデータを統合する
    """
    root = None
    
    try:
        # tkinter初期化（messagebox用）
        root = tk.Tk()
        root.withdraw()
        
        print("=" * 60)
        print("Step 7: マッピングデータ統合ツール")
        print("=" * 60)

        # --------------------------------------
        # 1. 自動成功分 (Candidate) の読み込み
        # --------------------------------------
        df_candidate_clean = pd.DataFrame()
        candidate_count = 0
        
        if os.path.exists(CANDIDATE_FILE):
            df_candidate = pd.read_excel(CANDIDATE_FILE, engine='openpyxl')
            candidate_count = len(df_candidate)
            print(f"\n📖 自動成功リスト読込: {candidate_count} 件")
            df_candidate_clean = standardize_columns(df_candidate, "自動成功リスト")
            print(f"   → 有効データ: {len(df_candidate_clean)} 件")
        else:
            print("\n⚠️ 自動成功リストが見つかりません")

        # --------------------------------------
        # 2. 手動補完分 (Unmatched) の読み込み
        # --------------------------------------
        df_manual_clean = pd.DataFrame()
        manual_total = 0
        manual_filled = 0
        
        if os.path.exists(UNMATCHED_FILE):
            print(f"\n📖 手動補完リスト読込中...")
            df_unmatched = pd.read_excel(UNMATCHED_FILE, engine='openpyxl')
            manual_total = len(df_unmatched)
            print(f"   総行数: {manual_total} 件")
            
            # UIDカラムを探す
            uid_col = find_uid_column(df_unmatched)
            
            if uid_col:
                # UIDが空でない行を抽出
                df_unmatched[uid_col] = df_unmatched[uid_col].apply(clean_value)
                df_manual = df_unmatched[df_unmatched[uid_col] != ""].copy()
                manual_filled = len(df_manual)
                
                if manual_filled > 0:
                    # カラム名を統一
                    df_manual = df_manual.rename(columns={uid_col: "自社UID"})
                    df_manual_clean = standardize_columns(df_manual, "手動補完リスト")
                    print(f"   → UID入力済み: {manual_filled} 件")
                else:
                    print("   → UID入力済み: 0 件（まだ手動入力されていません）")
            else:
                print("   ⚠️ 「自社UID」列が見つかりません")
                print("      手動でUID列を追加してから再実行してください")
        else:
            print("\nℹ️ 手動補完リストはありません")

        # --------------------------------------
        # 3. 合体 (Merge)
        # --------------------------------------
        print("\n" + "-" * 40)
        print("⚙️  データを統合中...")
        
        df_new_data = pd.concat([df_candidate_clean, df_manual_clean], ignore_index=True)
        df_new_data = df_new_data.reset_index(drop=True)
        new_data_count = len(df_new_data)

        if df_new_data.empty:
            messagebox.showwarning("警告", "統合すべきデータが1件もありませんでした。")
            return

        print(f"   今回の追加候補: {new_data_count} 件")

        # --------------------------------------
        # 4. 既存ファイルとの統合 & 保存
        # --------------------------------------
        existing_count = 0
        duplicate_count = 0
        
        if os.path.exists(FINAL_MAPPING_FILE):
            print(f"\n🔄 既存マッピングテーブルに追記...")
            df_existing = pd.read_excel(FINAL_MAPPING_FILE, engine='openpyxl')
            existing_count = len(df_existing)
            print(f"   既存データ: {existing_count} 件")
            
            # 既存 + 新規
            df_merged = pd.concat([df_existing, df_new_data], ignore_index=True)
            df_merged = df_merged.reset_index(drop=True)
            
            # 重複排除 (自社UID と 卸側施設ID の組み合わせ、最新を残す)
            before_len = len(df_merged)
            df_merged = df_merged.drop_duplicates(subset=["自社UID", "卸側施設ID"], keep='last')
            df_merged = df_merged.reset_index(drop=True)
            duplicate_count = before_len - len(df_merged)
        else:
            print("\n✨ 新規マッピングテーブルを作成...")
            df_merged = df_new_data.drop_duplicates(subset=["自社UID", "卸側施設ID"], keep='last')
            df_merged = df_merged.reset_index(drop=True)
            duplicate_count = new_data_count - len(df_merged)

        # 卸側施設IDが空の行は削除
        df_merged["卸側施設ID"] = df_merged["卸側施設ID"].apply(clean_value)
        df_merged = df_merged[df_merged["卸側施設ID"] != ""]
        df_merged = df_merged.reset_index(drop=True)
        
        final_count = len(df_merged)

        # 保存
        df_merged.to_excel(FINAL_MAPPING_FILE, index=False, engine='openpyxl')

        # --------------------------------------
        # 5. 結果サマリー
        # --------------------------------------
        print("\n" + "=" * 60)
        print("【統合結果サマリー】")
        print("=" * 60)
        print("【入力データ】")
        print(f"  自動成功リスト:   {len(df_candidate_clean)} 件")
        print(f"  手動補完リスト:   {len(df_manual_clean)} 件")
        print(f"  → 今回の追加:    {new_data_count} 件")
        print("-" * 60)
        print("【統合処理】")
        if existing_count > 0:
            print(f"  既存データ:       {existing_count} 件")
        if duplicate_count > 0:
            print(f"  重複削除:         {duplicate_count} 件")
        print(f"  → 最終登録数:    {final_count} 件")
        print("=" * 60)
        print(f"\n📁 保存先: {FINAL_MAPPING_FILE}")

        # --------------------------------------
        # 完了報告
        # --------------------------------------
        remaining_unmatched = manual_total - manual_filled if manual_total > 0 else 0
        
        msg = (
            f"✅ 統合完了！\n\n"
            f"【今回の追加】\n"
            f"・自動成功: {len(df_candidate_clean)} 件\n"
            f"・手動補完: {len(df_manual_clean)} 件\n\n"
            f"【現在の登録総数】\n"
            f"　{final_count} 件\n\n"
        )
        
        if remaining_unmatched > 0:
            msg += (
                f"⚠️ 未処理: {remaining_unmatched} 件\n"
                f"（unmatched_list.xlsxで手動入力後に再実行）\n\n"
            )
        
        msg += f"保存先:\n{FINAL_MAPPING_FILE}"
        
        messagebox.showinfo("成功", msg)
        
        if os.name == 'nt':
            os.startfile(STORAGE_DIR)
        else:
            os.system(f"open '{STORAGE_DIR}'")

    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n\n{e}")
        
    finally:
        if root:
            root.destroy()


if __name__ == "__main__":
    step7_merge_final()

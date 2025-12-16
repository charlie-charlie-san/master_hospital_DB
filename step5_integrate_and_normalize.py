#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 5: 請求データ統合・正規化ツール (step5_integrate_and_normalize.py)
========================================================================
複数月の請求データを統合し、住所・名称を正規化して
マスターとのマッチング用キーを生成するツール

【機能】
- 複数ファイルを一括選択して統合
- シート自動探索（ヘッダー行を自動検出）
- 得意先コードで重複排除
- 郵便番号の正規化（先頭0対応・ハイフン形式）
- 住所・名称の正規化（マッチング用キー生成）

保存先: ~/Desktop/hospital_DB/2_Storage/
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import mojimoji
import re
from datetime import datetime

# ==========================================
# 設定
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
OUTPUT_DIR = os.path.join(BASE_DIR, "2_Storage")
OUTPUT_FILE = "unique_customer_list_normalized.xlsx"

# 必須カラム（自動探索用）
REQUIRED_COLS_CHECK = ["得意先コード", "得意先名称"]

# 抽出対象カラム
TARGET_COLS = ["得意先コード", "得意先名称", "郵便番号", "住所１", "住所２", "電話番号", "TEL"]

# 正規化設定
CORP_TITLES = [
    "株式会社", "有限会社", "合同会社", "合資会社", "合名会社",
    "医療法人", "医療法人社団", "医療法人財団",
    "社団法人", "財団法人", "一般社団法人", "一般財団法人",
    "公益社団法人", "公益財団法人",
    "社会福祉法人", "学校法人", "宗教法人",
    "NPO法人", "特定非営利活動法人",
    "(株)", "(有)", "（株）", "（有）"
]
KANJI_NUM_MAP = str.maketrans("一二三四五六七八九〇", "1234567890")


# ==========================================
# ユーティリティ関数
# ==========================================
def normalize_postal_code(val):
    """
    郵便番号を正規化する
    - 先頭0が消えないように7桁ゼロ埋め
    - 3桁-4桁のハイフン形式で出力
    """
    if pd.isna(val):
        return ""
    
    if isinstance(val, (int, float)):
        val = str(int(val))
    else:
        val = str(val).strip()
    
    if val.lower() in ["nan", "none", "null", "nat", ""]:
        return ""
    
    val = mojimoji.zen_to_han(val)
    val = val.replace("〒", "").strip()
    digits_only = re.sub(r"[^\d]", "", val)
    
    if len(digits_only) <= 6 and len(digits_only) >= 1:
        digits_only = digits_only.zfill(7)
    
    if len(digits_only) == 7:
        return digits_only[:3] + "-" + digits_only[3:]
    
    return val


def normalize_phone(val):
    """
    電話番号を正規化する
    """
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if val.lower() in ["nan", "none", "null", ""]:
        return ""
    val = mojimoji.zen_to_han(val)
    val = re.sub(r"[ー－―‐–—]", "-", val)
    return val


def normalize_text_for_matching(text):
    """
    住所や名称を「比較しやすい形」に強制変換する（マッチング用）
    """
    if pd.isna(text):
        return ""
    text = str(text)
    
    # 1. 半角全角統一
    text = mojimoji.zen_to_han(text, kana=False)
    text = mojimoji.han_to_zen(text, digit=False, ascii=False)
    
    # 2. 漢数字をアラビア数字に (一丁目 -> 1丁目)
    text = text.translate(KANJI_NUM_MAP)
    
    # 3. 法人格などを削除
    for title in CORP_TITLES:
        text = text.replace(title, "")
    
    # 4. 記号、スペース、ハイフン、丁目番地などをすべて削除
    text = re.sub(r'[\s\-‐－ー―丁目番地号ビル階F棟室]+', '', text)
    
    return text.lower()


# ==========================================
# シート自動探索機能
# ==========================================
def find_data_sheet(excel_path):
    """
    Excel内の全シートを探し、得意先データが含まれるシートを返す
    """
    try:
        xls = pd.ExcelFile(excel_path, engine='openpyxl')
        
        for sheet in xls.sheet_names:
            df_pre = pd.read_excel(excel_path, sheet_name=sheet, header=None, nrows=20, engine='openpyxl')
            
            for i, row in df_pre.iterrows():
                row_text = " ".join(row.astype(str))
                if all(col in row_text for col in REQUIRED_COLS_CHECK):
                    print(f"      ✅ 発見: シート'{sheet}' (ヘッダー: {i+1}行目)")
                    return pd.read_excel(excel_path, sheet_name=sheet, header=i, engine='openpyxl')
        
        return None
        
    except Exception as e:
        print(f"      ❌ 読込エラー: {e}")
        return None


# ==========================================
# メイン処理 (Step 5)
# ==========================================
def step5_integrate_and_normalize():
    """
    請求データを統合し、正規化してマッチング用キーを生成する
    """
    root = None
    
    try:
        print("=" * 60)
        print("Step 5: 請求データ統合・正規化ツール")
        print("=" * 60)

        # 1. 出力先準備
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
            print(f"✓ 出力フォルダ作成: {OUTPUT_DIR}")

        # 2. 複数ファイル選択
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)

        print("\n📂 請求データ(Excel)を【まとめて】選択してください...")
        print("   （8月〜11月など、複数ファイルを選択可能）")
        
        file_paths = filedialog.askopenfilenames(
            title="Step5: 請求データ(複数)を選択",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialdir=os.path.expanduser("~/Desktop")
        )

        if not file_paths:
            print("キャンセルされました。")
            return

        print(f"\n✅ {len(file_paths)} ファイルが選択されました")
        print("-" * 40)

        # 3. 読み込みループ
        all_data_list = []
        file_stats = []
        
        for path in file_paths:
            file_name = os.path.basename(path)
            print(f"\n📖 処理中: {file_name}")
            
            df = find_data_sheet(path)
            
            if df is not None:
                cols_to_keep = [c for c in TARGET_COLS if c in df.columns]
                
                if cols_to_keep:
                    df_filtered = df[cols_to_keep].copy()
                    row_count = len(df_filtered)
                    all_data_list.append(df_filtered)
                    file_stats.append({"file": file_name, "rows": row_count, "status": "OK"})
                    print(f"      📊 {row_count:,} 行を取得")
                else:
                    file_stats.append({"file": file_name, "rows": 0, "status": "カラムなし"})
                    print(f"      ⚠️ 対象カラムが見つかりません")
            else:
                file_stats.append({"file": file_name, "rows": 0, "status": "シートなし"})
                print(f"      ⚠️ スキップ: 有効なデータが見つかりません")

        if not all_data_list:
            messagebox.showerror("エラー", "有効なデータが1つも読み込めませんでした。")
            return

        # 4. 結合
        print("\n" + "-" * 40)
        print("⚙️  全データを結合中...")
        df_combined = pd.concat(all_data_list, ignore_index=True)
        total_rows = len(df_combined)
        print(f"   結合後の全行数: {total_rows:,} 行")

        # 5. 得意先コードがnullの行を除外
        df_combined = df_combined[df_combined["得意先コード"].notna()]
        print(f"   得意先コードあり: {len(df_combined):,} 行")

        # 6. ユニーク化（得意先コードで重複排除、最新を残す）
        print("⚙️  得意先コードで重複を削除中...")
        df_unique = df_combined.drop_duplicates(subset=["得意先コード"], keep='last').copy()
        unique_count = len(df_unique)
        duplicate_count = len(df_combined) - unique_count
        print(f"   重複排除後: {unique_count:,} 行 (削除: {duplicate_count:,} 行)")

        # 7. データ正規化
        print("⚙️  データを正規化中...")
        
        # 郵便番号の正規化
        if "郵便番号" in df_unique.columns:
            df_unique["郵便番号"] = df_unique["郵便番号"].apply(normalize_postal_code)
            print("   ✓ 郵便番号を正規化（7桁ハイフン形式）")
        
        # 電話番号の正規化
        tel_col = None
        for col in ["電話番号", "TEL"]:
            if col in df_unique.columns:
                tel_col = col
                df_unique[col] = df_unique[col].apply(normalize_phone)
                print(f"   ✓ {col}を正規化")
                break

        # 8. 住所フル作成
        addr1 = df_unique.get("住所１", pd.Series([""] * len(df_unique))).fillna("").astype(str)
        addr2 = df_unique.get("住所２", pd.Series([""] * len(df_unique))).fillna("").astype(str)
        df_unique["住所フル"] = addr1 + addr2
        print("   ✓ 住所フルを生成（住所１+住所２）")

        # 9. マッチング用正規化キー生成
        print("🧹 マッチング用の正規化キーを生成中...")
        df_unique["正規化住所キー"] = df_unique["住所フル"].apply(normalize_text_for_matching)
        df_unique["正規化名称キー"] = df_unique["得意先名称"].apply(normalize_text_for_matching)
        print("   ✓ 正規化住所キー、正規化名称キーを生成")

        # 10. 保存
        save_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE)
        df_unique.to_excel(save_path, index=False, engine='openpyxl')

        # 11. 結果サマリー
        print("\n" + "=" * 60)
        print("【処理結果サマリー】")
        print("=" * 60)
        print(f"  入力ファイル数:   {len(file_paths)}")
        print(f"  結合前の全行数:   {total_rows:,} 行")
        print(f"  重複削除数:       {duplicate_count:,} 行")
        print(f"  → 出力行数:      {unique_count:,} 行")
        print("-" * 60)
        print("【ファイル別統計】")
        for stat in file_stats:
            status_icon = "✅" if stat["status"] == "OK" else "⚠️"
            print(f"  {status_icon} {stat['file']}: {stat['rows']:,} 行 ({stat['status']})")
        print("-" * 60)
        print("【生成されたカラム】")
        print("  ・住所フル: 住所１+住所２を結合")
        print("  ・正規化住所キー: マッチング用（空白・記号除去）")
        print("  ・正規化名称キー: マッチング用（法人格除去）")
        print("=" * 60)
        print(f"\n📁 保存先: {save_path}")

        # 完了メッセージ
        msg = (
            f"✅ Step 5 完了！\n\n"
            f"入力ファイル数: {len(file_paths)}\n"
            f"結合前の行数: {total_rows:,}\n"
            f"ユニーク施設数: {unique_count:,}\n\n"
            f"保存先:\n{save_path}\n\n"
            f"★「正規化住所キー」「正規化名称キー」列を使って\n"
            f"  次のStepでマスターと突き合わせます。"
        )
        messagebox.showinfo("成功", msg)
        
        if os.name == 'nt':
            os.startfile(OUTPUT_DIR)
        else:
            os.system(f"open '{OUTPUT_DIR}'")

    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n\n{e}")
        
    finally:
        if root:
            root.destroy()


if __name__ == "__main__":
    step5_integrate_and_normalize()

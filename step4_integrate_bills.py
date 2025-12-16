#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 4: 売上データ統合ツール (step4_integrate_bills.py)
========================================================
複数月の売上データ（Excel）を読み込み、統合して
8月〜11月の売上明細を1つのファイルにまとめるツール

【対象カラム】
売上日, 売上№, 売上行№, 元売上№返品, 元売上行№返品, 
売上取引区分, 区分名称, 商品コード, ＪＡＮコード, 商品名, 
商品規格, 売上数, 売上単価, 売上金額

保存先: ~/Desktop/hospital_DB/work_space/
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime

# ==========================================
# 設定
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
OUTPUT_DIR = os.path.join(BASE_DIR, "2_Storage")
OUTPUT_FILE = "integrated_sales_data.xlsx"

# 必須カラム（これがあるシートを自動で探す）
REQUIRED_COLS_CHECK = ["売上日", "商品コード", "売上金額"]

# 対象カラム（このカラムを抽出）
TARGET_COLS = [
    "売上日",
    "売上№",
    "売上行№",
    "元売上№返品",
    "元売上行№返品",
    "売上取引区分",
    "区分名称",
    "商品コード",
    "ＪＡＮコード",
    "商品名",
    "商品規格",
    "売上数",
    "売上単価",
    "売上金額"
]


# ==========================================
# ユーティリティ関数
# ==========================================
def parse_date(val):
    """
    日付をYYYY/MM/DD形式に変換
    """
    if pd.isna(val):
        return ""
    
    # 既にdatetime型の場合
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%Y/%m/%d")
    
    # 数値型の場合（Excelシリアル値）
    if isinstance(val, (int, float)):
        try:
            if 1 <= val <= 73050:
                parsed = pd.to_datetime(val, unit='D', origin='1899-12-30')
                return parsed.strftime("%Y/%m/%d")
        except:
            pass
        return ""
    
    # 文字列の場合
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", "null", "nat", ""]:
        return ""
    
    try:
        parsed = pd.to_datetime(val_str)
        return parsed.strftime("%Y/%m/%d")
    except:
        return val_str


def clean_numeric(val):
    """
    数値を整形（カンマ区切りなど除去）
    """
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return val
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", "null", ""]:
        return 0
    # カンマを除去
    val_str = val_str.replace(",", "")
    try:
        return float(val_str)
    except:
        return 0


# ==========================================
# シート自動探索機能
# ==========================================
def find_data_sheet(excel_path):
    """
    Excel内の全シートを探し、売上データが含まれるシートを返す
    ヘッダー行も自動検出する
    """
    try:
        xls = pd.ExcelFile(excel_path, engine='openpyxl')
        
        for sheet in xls.sheet_names:
            # 先頭20行を読んでヘッダーを探す
            df_pre = pd.read_excel(excel_path, sheet_name=sheet, header=None, nrows=20, engine='openpyxl')
            
            for i, row in df_pre.iterrows():
                row_text = " ".join(row.astype(str))
                # 必須カラムが含まれているかチェック
                if all(col in row_text for col in REQUIRED_COLS_CHECK):
                    print(f"      ✅ 発見: シート'{sheet}' (ヘッダー: {i+1}行目)")
                    return pd.read_excel(excel_path, sheet_name=sheet, header=i, engine='openpyxl')
        
        return None
        
    except Exception as e:
        print(f"      ❌ 読込エラー: {e}")
        return None


# ==========================================
# メイン処理
# ==========================================
def step4_integrate_bills():
    """
    売上データを統合する
    """
    root = None
    
    try:
        print("=" * 60)
        print("Step 4: 売上データ統合ツール")
        print("=" * 60)

        # 1. 出力先準備
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
            print(f"✓ 出力フォルダ作成: {OUTPUT_DIR}")

        # 2. 複数ファイル選択
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)

        print("\n📂 売上データ(Excel)を【まとめて】選択してください...")
        print("   （8月〜11月など、複数ファイルを選択可能）")
        
        file_paths = filedialog.askopenfilenames(
            title="Step4: 売上データ(複数)を選択",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialdir=os.path.expanduser("~/Desktop")
        )

        if not file_paths:
            print("キャンセルされました。")
            return

        print(f"\n✅ {len(file_paths)} ファイルが選択されました")
        print("-" * 40)

        # 3. 各ファイルを読み込み
        all_data_list = []
        file_stats = []
        
        for path in file_paths:
            file_name = os.path.basename(path)
            print(f"\n📖 処理中: {file_name}")
            
            # シートを自動探索して読み込む
            df = find_data_sheet(path)
            
            if df is not None:
                # 必要な列だけ抽出（存在する列のみ）
                cols_to_keep = [c for c in TARGET_COLS if c in df.columns]
                
                if cols_to_keep:
                    df_filtered = df[cols_to_keep].copy()
                    
                    # ファイル名から月を抽出して追加（参考用）
                    df_filtered["元ファイル"] = file_name
                    
                    row_count = len(df_filtered)
                    all_data_list.append(df_filtered)
                    file_stats.append({"file": file_name, "rows": row_count, "status": "OK"})
                    print(f"      📊 {row_count:,} 行を取得")
                    print(f"      📋 カラム: {cols_to_keep[:5]}...")
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

        # 5. データ正規化
        print("⚙️  データを正規化中...")
        
        # 売上日の正規化
        if "売上日" in df_combined.columns:
            df_combined["売上日"] = df_combined["売上日"].apply(parse_date)
            print("   ✓ 売上日を正規化（YYYY/MM/DD形式）")
        
        # 数値カラムの正規化
        numeric_cols = ["売上数", "売上単価", "売上金額"]
        for col in numeric_cols:
            if col in df_combined.columns:
                df_combined[col] = df_combined[col].apply(clean_numeric)
        print("   ✓ 数値カラムを正規化")

        # 6. 重複チェック（売上№と売上行№の組み合わせ）
        if "売上№" in df_combined.columns and "売上行№" in df_combined.columns:
            before_count = len(df_combined)
            df_combined = df_combined.drop_duplicates(subset=["売上№", "売上行№"], keep='last')
            after_count = len(df_combined)
            duplicate_count = before_count - after_count
            if duplicate_count > 0:
                print(f"   ✓ 重複排除: {duplicate_count:,} 行を削除")
        else:
            duplicate_count = 0

        # 7. 売上日でソート
        if "売上日" in df_combined.columns:
            df_combined = df_combined.sort_values("売上日", ascending=True)
            print("   ✓ 売上日でソート")

        # 8. 集計情報
        total_sales = 0
        if "売上金額" in df_combined.columns:
            total_sales = df_combined["売上金額"].sum()

        # 9. 保存
        save_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE)
        df_combined.to_excel(save_path, index=False, engine='openpyxl')

        # 10. 結果サマリー
        print("\n" + "=" * 60)
        print("【処理結果サマリー】")
        print("=" * 60)
        print(f"  入力ファイル数:   {len(file_paths)}")
        print(f"  結合後の行数:     {total_rows:,} 行")
        if duplicate_count > 0:
            print(f"  重複削除数:       {duplicate_count:,} 行")
        print(f"  → 最終出力行数:  {len(df_combined):,} 行")
        print(f"  売上金額合計:     ¥{total_sales:,.0f}")
        print("-" * 60)
        print("【ファイル別統計】")
        for stat in file_stats:
            status_icon = "✅" if stat["status"] == "OK" else "⚠️"
            print(f"  {status_icon} {stat['file']}: {stat['rows']:,} 行 ({stat['status']})")
        print("=" * 60)
        print(f"\n📁 保存先: {save_path}")

        # 完了メッセージ
        msg = (
            f"✅ 売上データ統合が完了しました！\n\n"
            f"入力ファイル数: {len(file_paths)}\n"
            f"最終行数: {len(df_combined):,}\n"
            f"売上金額合計: ¥{total_sales:,.0f}\n\n"
            f"保存先:\n{save_path}"
        )
        messagebox.showinfo("成功", msg)
        
        # フォルダを開く
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
    step4_integrate_bills()

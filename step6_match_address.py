#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 6: マスターDBマッチングツール (step6_match_address.py)
============================================================
Step 5で生成した正規化リストとマスターDBを突き合わせて
自動マッチングを行い、IDマッピングの候補を生成するツール

【マッチング方法】
1. 住所+施設名先頭2文字 でマッチング
2. 電話番号 でマッチング（方法1で失敗した場合）

保存先: ~/Desktop/hospital_DB/2_Storage/
"""

import pandas as pd
import os
import mojimoji
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# ==========================================
# 設定
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
STORAGE_DIR = os.path.join(BASE_DIR, "2_Storage")

# 入力ファイル (Step 5で作ったもの)
INPUT_LIST_FILE = "unique_customer_list_normalized.xlsx"

# マスターDB
MASTER_FILE = os.path.join(STORAGE_DIR, "master_db.xlsx")

# 卸業者名（固定値）
FIXED_WHOLESALER_NAME = "アスコ"

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
def normalize_text_for_matching(text):
    """
    住所や名称を正規化する（マッチング用）
    """
    if pd.isna(text):
        return ""
    text = str(text)
    text = mojimoji.zen_to_han(text, kana=False)
    text = mojimoji.han_to_zen(text, digit=False, ascii=False)
    text = text.translate(KANJI_NUM_MAP)
    for title in CORP_TITLES:
        text = text.replace(title, "")
    text = re.sub(r'[\s\-‐－ー―丁目番地号ビル階F棟室]+', '', text)
    return text.lower()


def normalize_phone(phone):
    """
    電話番号を正規化（数字のみ抽出）
    """
    if pd.isna(phone):
        return ""
    phone = str(phone)
    phone = mojimoji.zen_to_han(phone)
    phone = re.sub(r'[^\d]', '', phone)
    return phone


def find_uid_column(df):
    """
    UIDカラムを探す（自社UID / 施設UID 両対応）
    """
    for col in df.columns:
        col_str = str(col)
        if "UID" in col_str:
            return col
    return None


# ==========================================
# メイン処理 (Step 6)
# ==========================================
def step6_match_address():
    """
    マスターDBとのマッチングを実行する
    """
    root = None
    
    try:
        print("=" * 60)
        print("Step 6: マスターDBマッチングツール")
        print("=" * 60)

        # 1. マスターDB読み込み
        if not os.path.exists(MASTER_FILE):
            messagebox.showerror("エラー", f"マスターDBが見つかりません。\n{MASTER_FILE}")
            return
        
        print("\n📖 マスターDBを読み込み中...")
        df_master = pd.read_excel(MASTER_FILE, engine='openpyxl')
        print(f"   マスター件数: {len(df_master)} 件")
        
        # UIDカラムを探す
        uid_col = find_uid_column(df_master)
        if not uid_col:
            messagebox.showerror("エラー", "マスターDBにUID列が見つかりません。")
            return
        print(f"   UID列: {uid_col}")
        
        # マスター側の照合キーを作成
        print("   照合キーを作成中...")
        
        # キー1: 正規化住所 + 施設名先頭2文字
        df_master["NormAddr"] = df_master["住所"].apply(normalize_text_for_matching)
        df_master["NormName"] = df_master["動物病院施設名"].apply(normalize_text_for_matching)
        df_master["MatchKey_Addr"] = df_master["NormAddr"] + df_master["NormName"].str[:2]
        
        # キー2: 電話番号（数字のみ）
        if "電話番号" in df_master.columns:
            df_master["MatchKey_Tel"] = df_master["電話番号"].apply(normalize_phone)
        else:
            df_master["MatchKey_Tel"] = ""

        # 2. 請求ユニークリスト読み込み
        input_path = os.path.join(STORAGE_DIR, INPUT_LIST_FILE)
        
        if not os.path.exists(input_path):
            print("\n📂 ファイルが見つからないため、選択してください...")
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            input_path = filedialog.askopenfilename(
                initialdir=STORAGE_DIR,
                title=f"Step5で作ったファイル({INPUT_LIST_FILE})を選択",
                filetypes=[("Excel Files", "*.xlsx")]
            )
            if not input_path:
                print("キャンセルされました。")
                return

        print(f"\n📖 請求リスト読み込み中: {os.path.basename(input_path)}")
        df_unique = pd.read_excel(input_path, engine='openpyxl')
        print(f"   請求リスト件数: {len(df_unique)} 件")

        # 3. マッチング処理
        matched_rows = []
        unmatched_rows = []
        match_stats = {"住所+名前": 0, "電話番号": 0}
        
        print(f"\n⚙️  {len(df_unique)}件のマッチングを実行中...")
        print("   方法1: 住所+施設名 / 方法2: 電話番号")
        print("-" * 40)

        for idx, row in df_unique.iterrows():
            wholesaler_id = row["得意先コード"]
            name_original = row["得意先名称"]
            
            # Step 5で作られたキーを使う（なければその場で生成）
            if "正規化住所キー" in row.index and pd.notna(row["正規化住所キー"]):
                norm_addr = str(row["正規化住所キー"])
            else:
                addr_full = str(row.get("住所フル", ""))
                norm_addr = normalize_text_for_matching(addr_full)

            if "正規化名称キー" in row.index and pd.notna(row["正規化名称キー"]):
                norm_name = str(row["正規化名称キー"])
            else:
                norm_name = normalize_text_for_matching(name_original)

            # マッチング方法1: 住所+名前
            bill_key_addr = norm_addr + norm_name[:2]
            match = df_master[df_master["MatchKey_Addr"] == bill_key_addr]
            match_method = "住所+名前"
            
            # マッチング方法2: 電話番号（方法1で失敗した場合）
            if match.empty:
                bill_tel = ""
                for tel_col in ["電話番号", "TEL"]:
                    if tel_col in row.index and pd.notna(row[tel_col]):
                        bill_tel = normalize_phone(row[tel_col])
                        break
                
                if bill_tel and len(bill_tel) >= 9:
                    match = df_master[df_master["MatchKey_Tel"] == bill_tel]
                    match_method = "電話番号"
            
            if not match.empty:
                # ✅ ヒット！
                master_row = match.iloc[0]
                match_stats[match_method] += 1
                
                matched_rows.append({
                    "自社UID": master_row[uid_col],
                    "施設名(確認用)": master_row["動物病院施設名"], 
                    "卸業者名": FIXED_WHOLESALER_NAME,
                    "卸側施設ID": wholesaler_id,
                    "卸側名称(参考)": name_original,
                    "適用開始日": master_row.get("価格適用開始日", ""),
                    "マッチ方法": match_method
                })
            else:
                # ❌ 失敗
                unmatched_rows.append({
                    "得意先コード": wholesaler_id,
                    "得意先名称": name_original,
                    "住所フル": row.get("住所フル", ""),
                    "正規化キー(参考)": bill_key_addr,
                    "ステータス": "未マッチ"
                })

        # 4. 結果保存
        print("\n📁 結果を保存中...")
        
        # (A) 成功リスト
        if matched_rows:
            df_ok = pd.DataFrame(matched_rows)
            ok_path = os.path.join(STORAGE_DIR, "id_mapping_candidate.xlsx")
            df_ok.to_excel(ok_path, index=False, engine='openpyxl')
            print(f"   ✅ マッチ成功: {len(matched_rows)}件 -> id_mapping_candidate.xlsx")
        else:
            ok_path = os.path.join(STORAGE_DIR, "id_mapping_candidate.xlsx")
            pd.DataFrame().to_excel(ok_path, index=False, engine='openpyxl')

        # (B) 失敗リスト
        if unmatched_rows:
            df_ng = pd.DataFrame(unmatched_rows)
            ng_path = os.path.join(STORAGE_DIR, "unmatched_list.xlsx")
            df_ng.to_excel(ng_path, index=False, engine='openpyxl')
            print(f"   ⚠️ 未マッチ: {len(unmatched_rows)}件 -> unmatched_list.xlsx")

        # 5. 結果サマリー
        total = len(matched_rows) + len(unmatched_rows)
        success_rate = (len(matched_rows) / total * 100) if total > 0 else 0
        
        print("\n" + "=" * 60)
        print("【マッチング結果サマリー】")
        print("=" * 60)
        print(f"  入力件数:       {len(df_unique)} 件")
        print(f"  マッチ成功:     {len(matched_rows)} 件")
        print(f"  未マッチ:       {len(unmatched_rows)} 件")
        print(f"  成功率:         {success_rate:.1f}%")
        print("-" * 60)
        print("【マッチ方法別内訳】")
        print(f"  住所+名前:      {match_stats['住所+名前']} 件")
        print(f"  電話番号:       {match_stats['電話番号']} 件")
        print("=" * 60)
        print(f"\n📁 保存先: {STORAGE_DIR}")

        # 完了メッセージ
        if unmatched_rows:
            msg = (
                f"Step 6 完了！\n\n"
                f"✅ マッチ成功: {len(matched_rows)}件\n"
                f"⚠️ 未マッチ: {len(unmatched_rows)}件\n"
                f"成功率: {success_rate:.1f}%\n\n"
                f"【マッチ方法】\n"
                f"・住所+名前: {match_stats['住所+名前']}件\n"
                f"・電話番号: {match_stats['電話番号']}件\n\n"
                f"「unmatched_list.xlsx」を開き、\n"
                f"手動でUIDを調べて記入してください。"
            )
        else:
            msg = f"完璧です！全{len(matched_rows)}件が自動マッチしました！"

        messagebox.showinfo("Step 6 完了", msg)
        
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
    step6_match_address()

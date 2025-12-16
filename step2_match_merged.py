import pandas as pd
import os
import mojimoji
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# ==========================================
# 設定: hospital_DB 環境
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
STORAGE_DIR = os.path.join(BASE_DIR, "2_Storage")
WORK_DIR = os.path.join(BASE_DIR, "work_space")

# 入力ファイル（さっき作った統合リスト）
INPUT_LIST_FILE = "unique_customer_list_merged.xlsx"

# マスターDB
MASTER_FILE = os.path.join(STORAGE_DIR, "master_db.xlsx")

# 卸業者名（マッピング用固定値）
FIXED_WHOLESALER_NAME = "アスコ"

# 正規化設定
CORP_TITLES = ["株式会社", "有限会社", "合同会社", "医療法人", "社団法人", "(株)", "(有)"]
KANJI_NUM_MAP = str.maketrans("一二三四五六七八九〇", "1234567890")

# ==========================================
# ロジック
# ==========================================
def normalize_text(text):
    """強力な正規化（住所・名称用）"""
    if pd.isna(text): return ""
    text = str(text)
    text = mojimoji.zen_to_han(text, kana=False)
    text = mojimoji.han_to_zen(text, digit=False, ascii=False)
    text = text.translate(KANJI_NUM_MAP)
    # 法人格削除
    for title in CORP_TITLES:
        text = text.replace(title, "")
    # 記号、スペース、ハイフン、丁目番地などをすべて削除して「文字と数字の塊」にする
    text = re.sub(r'[\s\-‐－ー―丁目番地号]+', '', text)
    return text


def normalize_phone(phone):
    """電話番号を正規化（数字のみ抽出）"""
    if pd.isna(phone): return ""
    phone = str(phone)
    phone = mojimoji.zen_to_han(phone)
    # 数字以外を削除
    phone = re.sub(r'[^\d]', '', phone)
    return phone


def main():
    root = None
    try:
        # 1. マスターDB読み込み
        if not os.path.exists(MASTER_FILE):
            messagebox.showerror("エラー", f"マスターDBが見つかりません。\n{MASTER_FILE}")
            return
        
        print("🔄 マスターDBを読み込んでいます...")
        df_master = pd.read_excel(MASTER_FILE)
        
        # マスター側の「照合用キー」を作成
        # キー1: 正規化住所 + 施設名の先頭2文字
        df_master["MatchKey_Addr"] = df_master["住所"].apply(normalize_text) + \
                                df_master["動物病院施設名"].apply(normalize_text).str[:2]
        
        # キー2: 電話番号（数字のみ）
        df_master["MatchKey_Tel"] = df_master["電話番号"].apply(normalize_phone)
        
        print(f"   マスター件数: {len(df_master)} 件")

        # 2. ユニークリスト読み込み
        input_path = os.path.join(WORK_DIR, INPUT_LIST_FILE)
        if not os.path.exists(input_path):
            # 見つからない場合は選択させる
            print("📂 ファイルが見つからないため、選択してください...")
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            input_path = filedialog.askopenfilename(
                initialdir=WORK_DIR,
                title=f"{INPUT_LIST_FILE} を選択",
                filetypes=[("Excel Files", "*.xlsx")]
            )
            if not input_path:
                print("キャンセルされました。")
                return

        print(f"📖 請求リスト読み込み中: {os.path.basename(input_path)}")
        df_unique = pd.read_excel(input_path)

        # 3. マッチング処理
        matched_rows = []
        unmatched_rows = []
        
        print(f"⚙️  {len(df_unique)}件のマッチングを実行中...")
        print("   方法1: 住所+施設名 / 方法2: 電話番号")

        for _, row in df_unique.iterrows():
            # 請求側の住所キー作成
            addr_full = str(row.get("住所フル", ""))
            # もし住所フルが空なら結合して作る
            if not addr_full or addr_full == "nan":
                addr_full = str(row.get("住所１","")) + str(row.get("住所２",""))

            name_full = str(row["得意先名称"])
            wholesaler_id = row["得意先コード"]
            
            # 正規化キー生成
            bill_key_addr = normalize_text(addr_full) + normalize_text(name_full)[:2]
            
            # マッチング方法1: 住所+名前
            match = df_master[df_master["MatchKey_Addr"] == bill_key_addr]
            match_method = "住所+名前"
            
            # マッチング方法2: 電話番号（方法1で見つからない場合）
            if match.empty:
                # 請求データに電話番号があれば使用
                bill_tel = ""
                if "電話番号" in row.index:
                    bill_tel = normalize_phone(row["電話番号"])
                elif "TEL" in row.index:
                    bill_tel = normalize_phone(row["TEL"])
                
                if bill_tel and len(bill_tel) >= 9:  # 9桁以上の電話番号でマッチング
                    match = df_master[df_master["MatchKey_Tel"] == bill_tel]
                    match_method = "電話番号"
            
            if not match.empty:
                # ✅ ヒット！ (自社UIDをゲット)
                master_row = match.iloc[0]
                matched_rows.append({
                    "自社UID": master_row["自社UID"],
                    "施設名(確認用)": master_row["動物病院施設名"], 
                    "卸業者名": FIXED_WHOLESALER_NAME,
                    "卸側施設ID": wholesaler_id,
                    "卸側名称(参考)": name_full,
                    "適用開始日": master_row.get("価格適用開始日", ""),
                    "マッチ方法": match_method
                })
            else:
                # ❌ 失敗 (手動チェック用)
                unmatched_rows.append({
                    "得意先コード": wholesaler_id,
                    "得意先名称": name_full,
                    "住所フル": addr_full,
                    "正規化キー(参考)": bill_key_addr,
                    "ステータス": "未マッチ"
                })

        # 4. 保存
        # (A) 成功リスト -> id_mapping の候補
        if matched_rows:
            df_ok = pd.DataFrame(matched_rows)
            ok_path = os.path.join(WORK_DIR, "id_mapping_candidate.xlsx")
            df_ok.to_excel(ok_path, index=False)
            print(f"✅ 自動マッチ成功: {len(matched_rows)}件 -> {ok_path}")

        # (B) 失敗リスト -> 手動チェック用
        if unmatched_rows:
            df_ng = pd.DataFrame(unmatched_rows)
            ng_path = os.path.join(WORK_DIR, "unmatched_list.xlsx")
            df_ng.to_excel(ng_path, index=False)
            print(f"⚠️ 未マッチデータ: {len(unmatched_rows)}件 -> {ng_path}")
            
            msg = (
                f"処理完了！\n\n"
                f"成功: {len(matched_rows)}件\n"
                f"失敗: {len(unmatched_rows)}件\n\n"
                f"失敗分は '{os.path.basename(ng_path)}' を確認し、\n"
                f"手動でマスターと紐付けてください。"
            )
        else:
            msg = f"完璧です！全{len(matched_rows)}件が自動マッチしました！"

        print(f"\n{'='*50}")
        print(f"【結果サマリー】")
        print(f"  マッチ成功: {len(matched_rows)} 件")
        print(f"  未マッチ:   {len(unmatched_rows)} 件")
        print(f"  成功率:     {len(matched_rows)/(len(matched_rows)+len(unmatched_rows))*100:.1f}%")
        print(f"{'='*50}")

        messagebox.showinfo("完了", msg)
        if os.name == 'nt':
            os.startfile(WORK_DIR)
        else:
            os.system(f"open '{WORK_DIR}'")
            
    except Exception as e:
        print(f"❌ エラー: {e}")
        messagebox.showerror("エラー", str(e))
        
    finally:
        # tkinterのリソース解放
        if root:
            root.destroy()


if __name__ == "__main__":
    main()

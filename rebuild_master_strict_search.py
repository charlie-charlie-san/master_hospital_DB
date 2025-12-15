import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import datetime
import mojimoji
import re
from hashids import Hashids

# ==========================================
# 【設定】保存先: デスクトップ/hospital_DB
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
STORAGE_DIR = os.path.join(BASE_DIR, "2_Storage")
ARCHIVE_DIR = os.path.join(BASE_DIR, "9_Archives")

MASTER_FILE = "master_db.xlsx"
MAPPING_FILE = "id_mapping.xlsx"
FIXED_WHOLESALER_NAME = "アスコ"

ID_SALT = "Financial_System_Secret_Key_2025" 
ID_LENGTH = 6
ID_ALPHABET = "ABCDEFGHJKMNPQRSTVWXYZ23456789"

CORP_TITLES = ["株式会社", "有限会社", "合同会社", "一般社団法人", "公益社団法人", "医療法人", r"\(株\)", r"\(有\)"]
KANJI_NUM_MAP = str.maketrans("一二三四五六七八九〇", "1234567890")

# ==========================================
# ロジック
# ==========================================
def normalize_text(text):
    if pd.isna(text): return ""
    text = str(text)
    text = mojimoji.zen_to_han(text, kana=False)
    text = mojimoji.han_to_zen(text, digit=False, ascii=False)
    for title in CORP_TITLES:
        text = re.sub(title, "", text)
    text = text.translate(KANJI_NUM_MAP)
    text = re.sub(r'[\s\-‐－ー―]+', '', text)
    return text.strip()

def generate_id(index):
    hasher = Hashids(salt=str(ID_SALT), min_length=ID_LENGTH, alphabet=ID_ALPHABET)
    return hasher.encode(index)

def clean_val(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan": return ""
    return val


def normalize_postal_code(val):
    """
    郵便番号を正規化する
    - 先頭0が消えないように7桁ゼロ埋め
    - 3桁-4桁のハイフン形式で出力
    """
    if pd.isna(val):
        return ""
    
    # 数値型の場合は文字列に変換
    if isinstance(val, (int, float)):
        val = str(int(val))
    else:
        val = str(val).strip()
    
    if val.lower() in ["nan", "none", "null", "nat", ""]:
        return ""
    
    # 全角を半角に
    val = mojimoji.zen_to_han(val)
    # 〒マーク削除
    val = val.replace("〒", "").strip()
    # ハイフン・空白を削除して数字のみ取得
    digits_only = re.sub(r"[^\d]", "", val)
    
    # 6桁以下の場合は7桁にゼロ埋め（先頭に0を追加）
    if len(digits_only) <= 6:
        digits_only = digits_only.zfill(7)
    
    # 7桁の場合は XXX-XXXX 形式
    if len(digits_only) == 7:
        return digits_only[:3] + "-" + digits_only[3:]
    
    return val


def parse_date(val):
    """
    日付を解析してYYYY/MM/DD形式の文字列で返す
    - Excelシリアル値（5桁の数値）も正しく変換
    - 変換できない場合は空文字を返す
    """
    try:
        if pd.isna(val):
            return ""
        
        # 既にdatetime型の場合
        if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
            return val.strftime("%Y/%m/%d")
        
        # 数値型の場合（Excelシリアル値）
        if isinstance(val, (int, float)):
            # Excelシリアル値を日付に変換
            # 妥当な範囲（1900年〜2100年）かチェック
            if 1 <= val <= 73050:  # 約1900年〜2100年の範囲
                parsed = pd.to_datetime(val, unit='D', origin='1899-12-30')
                return parsed.strftime("%Y/%m/%d")
            else:
                return ""
        
        # 文字列の場合
        val_str = str(val).strip()
        if val_str.lower() in ["nan", "none", "null", "nat", ""]:
            return ""
        
        # 文字列が5桁の数値のみの場合（Excelシリアル値が文字列として読まれた場合）
        if val_str.isdigit() and 4 <= len(val_str) <= 5:
            serial = int(val_str)
            if 1 <= serial <= 73050:
                parsed = pd.to_datetime(serial, unit='D', origin='1899-12-30')
                return parsed.strftime("%Y/%m/%d")
        
        # 通常の日付文字列として解析
        parsed = pd.to_datetime(val_str)
        # 妥当な年かチェック（1950年〜2100年）
        if parsed.year < 1950 or parsed.year > 2100:
            return ""
        return parsed.strftime("%Y/%m/%d")
    except:
        return ""

def find_correct_dataframe(path):
    """
    全シートを巡回し、「病院住所」かつ「価格適用開始日」が含まれる本命シートを探す
    """
    print(f"🔍 全シートを厳しくスキャン中...")
    try:
        xls = pd.ExcelFile(path)
        sheet_names = xls.sheet_names
        print(f"   シート一覧: {sheet_names}")

        for sheet in sheet_names:
            # 先頭20行だけ読む
            df_pre = pd.read_excel(path, sheet_name=sheet, header=None, nrows=20)
            
            for i, row in df_pre.iterrows():
                row_text = " ".join(row.astype(str))
                
                # ★ここが進化：「住所」だけでなく「価格適用開始日」もあるかチェック！
                if "病院住所" in row_text and "価格適用" in row_text:
                    print(f"   ✅ 本命発見！ シート名: '{sheet}', ヘッダー行: {i+1}行目")
                    return pd.read_excel(path, sheet_name=sheet, header=i)
                
                # 「住所」はあるけど「価格適用」がない場合（惜しいシート）
                elif "病院住所" in row_text:
                    print(f"   ⚠️ スキップ: シート '{sheet}' は住所がありますが、重要項目が足りません。")
        
        print("   ❌ 条件を満たす完全なシートが見つかりませんでした。")
        return None

    except Exception as e:
        print(f"   ⚠️ スキャンエラー: {e}")
        return None

# ==========================================
# メイン処理
# ==========================================
def rebuild_master():
    # フォルダ作成
    for d in [STORAGE_DIR, ARCHIVE_DIR]:
        if not os.path.exists(d): os.makedirs(d)

    # ファイル選択
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    print("📂 入力データ(Excel)を選択してください...")
    input_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not input_path: return

    print(f"📖 ファイル解析開始: {os.path.basename(input_path)}")
    
    try:
        df_input = find_correct_dataframe(input_path)

        if df_input is None:
            raise ValueError("「病院住所」と「価格適用開始日」の両方を持つシートが見つかりませんでした。")

        # カラムマッピング
        COL_MAP = {
            "NAME": "動物病院施設名",
            "LEGAL": "法人名（法人の場合のみ）",
            "REP": "代表者名（漢字）",
            "ZIP": "病院 郵便番号",
            "ADDR": "病院住所",
            "TEL": "病院TEL（電話番号）",
            "EMAIL": "メールアドレス",
            "INVOICE": "適格請求書発行事業者の登録番号",
            "APP_DATE": "契約締結済",
            "PRICE_DATE": "価格適用開始日",
            "CANCEL_DATE": "解約日"
        }

        # 必須カラムチェック
        print(f"   読み込んだ列名: {list(df_input.columns)}")
        missing = [v for k, v in COL_MAP.items() if v not in df_input.columns]
        if missing:
            raise ValueError(f"必須カラムが見つかりません: {missing}")

        master_rows = []
        mapping_rows = []
        processed_keys = set()
        current_seq = 0

        print("⚙️  変換処理中...")
        for i, row in df_input.iterrows():
            addr = row.get(COL_MAP["ADDR"])
            tel = row.get(COL_MAP["TEL"])
            if pd.isna(addr) or pd.isna(tel): continue

            k = normalize_text(addr) + normalize_text(tel)
            if k in processed_keys: continue

            current_seq += 1
            uid = generate_id(current_seq)
            p_date = parse_date(row.get(COL_MAP["PRICE_DATE"]))
            # アクティブフラグ: 価格適用開始日が有効な日付なら1、空なら0
            is_active = 1 if p_date != "" else 0

            master_rows.append({
                "自社UID": uid,
                "動物病院施設名": clean_val(row.get(COL_MAP["NAME"])),
                "法人名": clean_val(row.get(COL_MAP["LEGAL"])),
                "代表者名": clean_val(row.get(COL_MAP["REP"])),
                "郵便番号": normalize_postal_code(row.get(COL_MAP["ZIP"])),  # 修正: 郵便番号正規化
                "住所": clean_val(addr),
                "電話番号": clean_val(tel),
                "メールアドレス": clean_val(row.get(COL_MAP["EMAIL"])),
                "インボイス登録番号": clean_val(row.get(COL_MAP["INVOICE"])),
                "申込日": parse_date(row.get(COL_MAP["APP_DATE"])),
                "価格適用開始日": p_date,
                "アクティブフラグ": is_active,
                "解約日": parse_date(row.get(COL_MAP["CANCEL_DATE"]))
            })

            mapping_rows.append({
                "自社UID": uid,
                "施設名(確認用)": clean_val(row.get(COL_MAP["NAME"])),
                "卸業者名": FIXED_WHOLESALER_NAME,
                "適用開始日": p_date
            })
            processed_keys.add(k)

        # 保存
        if master_rows:
            pd.DataFrame(master_rows).to_excel(os.path.join(STORAGE_DIR, MASTER_FILE), index=False)
            pd.DataFrame(mapping_rows).to_excel(os.path.join(STORAGE_DIR, MAPPING_FILE), index=False)
            
            msg = f"✅ 完了！\n保存先: {STORAGE_DIR}\n件数: {len(master_rows)}件"
            print(msg)
            messagebox.showinfo("成功", msg)
            if os.name == 'nt': os.startfile(STORAGE_DIR)
            else: os.system(f"open '{STORAGE_DIR}'")
        else:
            messagebox.showinfo("結果", "データがありませんでした")

    except Exception as e:
        print(f"❌ エラー: {e}")
        messagebox.showerror("エラー", str(e))

if __name__ == "__main__":
    rebuild_master()
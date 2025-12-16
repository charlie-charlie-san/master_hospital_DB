import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import mojimoji
import re

# ==========================================
# 設定: 保存先は "hospital_DB" の中の "work_space"
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
OUTPUT_DIR = os.path.join(BASE_DIR, "work_space")
OUTPUT_FILE = "unique_customer_list_merged.xlsx"

# 読み込むターゲットシート名
TARGET_SHEET_NAME = "Datalizer1"

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


# ==========================================
# メイン処理
# ==========================================
def extract_unique_multi():
    root = None
    try:
        # 1. 出力先準備
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)

        # 2. 複数ファイル選択ダイアログ
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)

        print("📂 請求データ(Excel)を【8月〜11月分まとめて】選択してください...")
        # 複数ファイル選択モード
        file_paths = filedialog.askopenfilenames(
            title="請求データ(8月,9月,10月,11月)をまとめて選択",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )

        if not file_paths:
            print("キャンセルされました。")
            return

        print(f"✅ {len(file_paths)} 個のファイルが選択されました。結合を開始します...")

        # 3. ループ処理でデータを積み上げる
        all_data_list = []
        # 抽出するカラム（存在しなければスキップするようにします）
        required_cols = ["得意先コード", "得意先名称", "郵便番号", "住所１", "住所２"]

        for path in file_paths:
            file_name = os.path.basename(path)
            print(f"   📖 読み込み中: {file_name} ...")
            
            try:
                # 指定シートのみ読み込み
                df = pd.read_excel(path, sheet_name=TARGET_SHEET_NAME)
                
                # カラムチェック
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    print(f"      ⚠️ スキップ: {file_name} に必要な列がありません {missing}")
                    continue

                # 必要な列だけ抽出してリストに追加
                all_data_list.append(df[required_cols])
                print(f"      ✓ {len(df)} 行を読み込みました")
                
            except ValueError:
                print(f"      ⚠️ スキップ: シート '{TARGET_SHEET_NAME}' が見つかりませんでした。")
            except Exception as e:
                print(f"      ❌ 読込エラー: {file_name} -> {e}")

        if not all_data_list:
            messagebox.showerror("エラー", "有効なデータが1つも読み込めませんでした。")
            return

        # 4. がっちゃんこ（結合）
        print("⚙️  全データを結合中...")
        df_combined = pd.concat(all_data_list, ignore_index=True)

        # 5. ユニーク化（全期間を通しての重複排除）
        print(f"   結合後の全行数: {len(df_combined):,} 行")
        print("⚙️  得意先コードで重複を削除しています...")
        
        # 得意先コードで重複を消す（最後のデータ＝最新を残す）
        df_unique = df_combined.drop_duplicates(subset=["得意先コード"], keep='last')
        
        print(f"   重複排除後: {len(df_unique):,} 行")

        # 6. 郵便番号の正規化（先頭0対応・ハイフン付き）
        print("⚙️  郵便番号を正規化中...")
        df_unique["郵便番号"] = df_unique["郵便番号"].apply(normalize_postal_code)

        # 7. 住所結合（後のマッチング用）
        df_unique["住所フル"] = df_unique["住所１"].fillna("").astype(str) + df_unique["住所２"].fillna("").astype(str)

        # 8. 保存
        save_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE)
        df_unique.to_excel(save_path, index=False)

        # 完了報告
        msg = (
            f"✅ 結合・ユニーク化が完了しました！\n\n"
            f"入力ファイル数: {len(file_paths)}\n"
            f"結合後の行数: {len(df_combined):,} 行\n"
            f"ユニーク施設数: {len(df_unique):,} 件\n\n"
            f"保存先: {save_path}"
        )
        print(msg)
        messagebox.showinfo("成功", msg)
        
        if os.name == 'nt':
            os.startfile(OUTPUT_DIR)
        else:
            os.system(f"open '{OUTPUT_DIR}'")
            
    except Exception as e:
        print(f"❌ エラー: {e}")
        messagebox.showerror("エラー", str(e))
        
    finally:
        # tkinterのリソース解放
        if root:
            root.destroy()


if __name__ == "__main__":
    extract_unique_multi()

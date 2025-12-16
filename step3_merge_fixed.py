import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox

# ==========================================
# 設定
# ==========================================
BASE_DIR = os.path.expanduser("~/Desktop/hospital_DB")
STORAGE_DIR = os.path.join(BASE_DIR, "2_Storage")
WORK_DIR = os.path.join(BASE_DIR, "work_space")

# 入力ファイル
CANDIDATE_FILE = os.path.join(WORK_DIR, "id_mapping_candidate.xlsx")  # 自動成功分
UNMATCHED_FILE = os.path.join(WORK_DIR, "unmatched_list.xlsx")        # 手動補完分

# 出力ファイル（完成形）
FINAL_MAPPING_FILE = os.path.join(STORAGE_DIR, "id_mapping.xlsx")

# 補完用設定
FIXED_WHOLESALER_NAME = "アスコ"

# ==========================================
# ロジック
# ==========================================
def standardize_columns(df, source_type="auto"):
    """
    データフレームのカラム名を「id_mapping.xlsx」の形式に強制統一する
    """
    if df.empty: return pd.DataFrame()

    # 名前変換ルール（左が見つかったら、右の名前に変える）
    rename_map = {
        "得意先コード": "卸側施設ID",   # ★ここが重要！
        "得意先名称": "施設名(確認用)",
        "UID": "自社UID",
        "自社ID": "自社UID",
        "施設UID": "自社UID"  # 施設UIDも対応
    }
    df = df.rename(columns=rename_map)

    # 必要なカラムが足りない場合は、空文字や固定値で埋める
    if "自社UID" not in df.columns:
        print(f"⚠️ {source_type}: '自社UID' 列が見つかりません。スキップします。")
        return pd.DataFrame()

    if "卸業者名" not in df.columns:
        df["卸業者名"] = FIXED_WHOLESALER_NAME

    if "適用開始日" not in df.columns:
        df["適用開始日"] = ""

    # 最終的なカラム構成を定義
    final_cols = ["自社UID", "施設名(確認用)", "卸業者名", "卸側施設ID", "適用開始日"]
    
    # 足りない列を空で作成
    for col in final_cols:
        if col not in df.columns:
            df[col] = ""

    # この順番で抽出して返す
    return df[final_cols]

# ==========================================
# メイン処理
# ==========================================
def main():
    print("🚀 マッピング統合プロセスを開始します (カラム補正版)...")

    # 1. 自動成功リスト(Candidate)の読み込み
    if os.path.exists(CANDIDATE_FILE):
        df_candidate = pd.read_excel(CANDIDATE_FILE)
        print(f"📖 自動成功分を読み込みました: {len(df_candidate)}件")
        
        # ★ここでカラム名を統一！
        df_candidate_clean = standardize_columns(df_candidate, "自動成功リスト")
    else:
        df_candidate_clean = pd.DataFrame()
        print("⚠️ 自動成功リストが見つかりません（0件として進めます）")

    # 2. 手動補完リスト(Unmatched)の読み込み
    df_manual_clean = pd.DataFrame()  # 初期化
    
    if os.path.exists(UNMATCHED_FILE):
        print("📖 手動補完リストを読み込んでいます...")
        df_unmatched = pd.read_excel(UNMATCHED_FILE)
        
        # 手動でUIDを入れた行だけ対象にする
        # 表記ゆれ対応: カラム名を探す（「施設」も追加）
        uid_col = None
        for col in df_unmatched.columns:
            col_str = str(col)
            if "UID" in col_str or "自社" in col_str or "施設" in col_str:
                uid_col = col
                break
        
        if uid_col and not df_unmatched.empty:
            # UIDがある行だけ抜き出す
            df_manual = df_unmatched[df_unmatched[uid_col].notna()].copy()
            # 空文字列も除外
            df_manual = df_manual[df_manual[uid_col].astype(str).str.strip() != ""]
            
            if len(df_manual) > 0:
                # カラム名を統一するために、一時的にリネーム
                df_manual = df_manual.rename(columns={uid_col: "自社UID"})
                
                print(f"✅ 手動入力データ: {len(df_manual)}件")
                
                # ★ここでカラム名を統一！
                df_manual_clean = standardize_columns(df_manual, "手動補完リスト")
            else:
                print("⚠️ 手動補完リストにUID入力済みのデータがありません。")
        else:
            print("⚠️ 手動補完リストに有効な「自社UID」が見つかりません。")
    else:
        print("ℹ️ 手動補完リストが見つかりません（スキップ）")

    # 3. 合体！
    print("⚙️  データを統合中...")
    df_new_data = pd.concat([df_candidate_clean, df_manual_clean], ignore_index=True)

    if len(df_new_data) == 0:
        messagebox.showwarning("警告", "保存すべきデータが1件もありませんでした。")
        return

    # 4. 既存のマッピングテーブル(id_mapping.xlsx)とのマージ
    if os.path.exists(FINAL_MAPPING_FILE):
        print(f"🔄 既存のマッピングテーブルを開いています...")
        df_existing = pd.read_excel(FINAL_MAPPING_FILE)
        
        # 既存データと新データを合体
        df_merged = pd.concat([df_existing, df_new_data], ignore_index=True)
        
        # 重複排除: 「自社UID」と「卸側施設ID」のペアが同じなら、重複とみなして消す（最新を残す）
        before_len = len(df_merged)
        df_merged = df_merged.drop_duplicates(subset=["自社UID", "卸側施設ID"], keep='last')
        after_len = len(df_merged)
        
        print(f"   既存: {len(df_existing)} + 新規: {len(df_new_data)} = 合計: {after_len} (重複削除: {before_len - after_len}件)")
    else:
        print("✨ 新規マッピングテーブルとして作成します...")
        df_merged = df_new_data

    # 5. 保存
    # 卸側施設IDが空の行は、意味がないので念のため削除
    df_merged = df_merged[df_merged["卸側施設ID"].astype(str).str.strip() != ""]
    
    df_merged.to_excel(FINAL_MAPPING_FILE, index=False)
    
    print("\n" + "=" * 50)
    print("【統合結果サマリー】")
    print(f"  自動成功分: {len(df_candidate_clean)} 件")
    print(f"  手動補完分: {len(df_manual_clean)} 件")
    print(f"  合計保存数: {len(df_merged)} 件")
    print("=" * 50)
    
    msg = (
        f"✅ 統合完了！\n\n"
        f"自動成功分: {len(df_candidate_clean)}件\n"
        f"手動補完分: {len(df_manual_clean)}件\n"
        f"現在の登録総数: {len(df_merged)}件\n\n"
        f"保存先: {FINAL_MAPPING_FILE}"
    )
    messagebox.showinfo("成功", msg)
    
    if os.name == 'nt':
        os.startfile(STORAGE_DIR)
    else:
        os.system(f"open '{STORAGE_DIR}'")


if __name__ == "__main__":
    main()

# 複数のエクセルを読み込み集計後に保存する

import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
import pandas as pd
import matplotlib.pyplot as plt
import japanize_matplotlib

#利用可能な列名を保持する
available_cols = []

def get_columns():
    """選択されたファイルから列名を取得"""
    global available_cols
    available_cols.clear()

    #最初のファイルから列名を取得(全ファイルで列名が共通である前提)
    try:
        df_temp = pd.read_excel(filepath_list[0], nrows = 0) #列名だけ取得
        available_cols = df_temp.columns.tolist()
        return available_cols
    except Exception as e:
        messagebox.showerror("列名取得エラー", f"ファイルの列名取得中にエラーが発生しました\n{e}")
        return []
    

filepath_list = []
def select_file():
    """ファイル選択ダイアログボタンが押された時の処理"""
    filepath = filedialog.askopenfilenames(
        title="ファイルを選択してください",
        filetypes=[
            ("Excelファイル", "*.xlsx"), #フィルタ: excelファイル
            ("すべてのファイル", "*.*")],
        initialdir="./",  # 初期表示ディレクトリ (カレントディレクトリ)
    )
    if filepath:
        #重複がないか確認
        for path in filepath:
            if path not in filepath_list:
                filepath_list.append(path)

        select_file_1.delete("1.0", tk.END)
        for key, path in enumerate(filepath_list):
            select_file_1.insert(tk.END, f"{key+1}. {path}\n")
            
        #ファイルが選択されたら列名を取得してドロップダウンを更新
        cols = get_columns()
        combobox_value_empty = [""] + cols #選択肢に空の文字列追加
        column_combobox_1["values"] = cols # 第1集計キー
        column_combobox_2["values"] = combobox_value_empty # 第2集計キー
        value_combobox["values"] = cols #値（集計対象）

        # 初期選択をクリア（任意）
        column_combobox_1.set("")
        column_combobox_2.set("")
        value_combobox.set("")

    else:
        # print("ファイルは選択されませんでした。")
        messagebox.showerror("エラー","ファイルは選択されませんでした。")


result_df = None
def calculat_shop():
    """開いたファイルの集計をする"""
    global result_df

    # ドロップダウンから選択された列名を取得
    col1 = column_combobox_1.get()
    col2 = column_combobox_2.get()
    value_col = value_combobox.get()

    if not col1:
        messagebox.showerror("エラー", "第1集計キーとなる列を選択してください")
        return
    if not value_col:
        messagebox.showerror("エラー", "集計する値の列を選択してください")
        return
    
    groupby_cols = [col1]
    if col2:
        groupby_cols.append(col2)

    if not filepath_list:
        messagebox.showerror("エラー", "ファイルが選択されていません")
        return
    
    # データフレームを入れるリストを作る
    df_list = []
    cols_to_read = groupby_cols + [value_col]
    cols_to_read = list(dict.fromkeys(cols_to_read))
    # エクセルのリストからファイル名を取ってきてread_excelでデータフレームを作る
    for val in filepath_list:
        try:
            df = pd.read_excel(val,usecols=cols_to_read)
            # 欠損あれば削除
            df = df.dropna(subset=cols_to_read)
            df_list.append(df)
            # print(df_list)
        except Exception as e:
            messagebox.showerror("ファイル読み込みエラー", f"ファイル{val}の読み込み中にエラーが発生しました：\n{e}")
            return

    #データフレームを結合する
    total = pd.concat(df_list)
    # print("結合されたデータフレーム：\n", total)

    # 商品分類で集計
    #値の列を数値型に変換する
    total[value_col] = pd.to_numeric(total[value_col], errors="coerce")
    #変換できなかった（数値化できない）行を削除
    total = total.dropna(subset=[value_col])

    result = total.groupby(groupby_cols).sum(numeric_only=True)[value_col]
    # ↑numeric_only=Trueで数値が書いてある列だけ合計するという意味
    print("集計結果：\n", result)

    #担当者と取引先をそれぞれ独立した列とする処理
    # (SeriesからDataFrameにする)
    result_df = result.reset_index()
    messagebox.showinfo("処理完了","集計が完了しました")



def save_file():
    """「別名で保存」ダイアログを表示し、ファイルにデータを保存する関数"""
    global result_df, filepath_list

    if result_df is None:
        messagebox.showerror("保存エラー", "保存するデータがありません\n先に「処理開始」ボタンを押して集計を行ってください")
        return
    
    filepath_save = filedialog.asksaveasfilename(
        title="ファイルを別名で保存",
        defaultextension=".xlsx",
        filetypes=[("Excelファイル", "*.xlsx"),
                   ("すべてのファイル", "*.*")],
        initialdir="./",  # 初期表示ディレクトリ (カレントディレクトリ)
    )

    # ファイルパスが選択された場合（ユーザーが「保存」をクリックした場合）
    if filepath_save:
        try:
            result_df.to_excel(filepath_save, index=True)
            messagebox.showinfo("保存完了", f"ファイルが正常に保存されました:\n{filepath_save}")
            select_file_1.delete("1.0", tk.END)
            filepath_list.clear()
            result_df = None
            column_combobox_1.set("")
            column_combobox_2.set("")
            value_combobox.set("")
            column_combobox_1["values"] = []
            column_combobox_2["values"] = []
            value_combobox["values"] = []

        except Exception as e:
            messagebox.showerror("保存エラー", f"ファイルの保存中にエラーが発生しました:\n{e}")
    else:
        # ファイルパスが選択されなかった場合（ユーザーが「キャンセル」をクリックした場合）
        messagebox.showinfo("保存キャンセル", "ファイルは保存されませんでした。")


#GUI画面
root = tk.Tk()
root.title("支店ごとの集計")
root.minsize(300,400)

select_btn_1 = tk.Button(root, text="ファイルを選択(複数選択可)", command=select_file,width=30)
select_btn_1.grid(row=0,column=0,padx=10,pady=10)

select_file_name_1 = tk.Label(text="選択されたファイルパス：")
select_file_name_1.grid(row=1,column=0,sticky=tk.W,padx=10,pady=10)
select_file_1 = tk.Text(width=70,height=7)
select_file_1.grid(row=2,column=0,sticky=tk.NSEW,padx=10,pady=10)

label_groupby_1 = tk.Label(root, text="第1集計キーを選択（必須）：")
label_groupby_1.grid(row=3, column=0, sticky=tk.W, padx=10, pady=5)
column_combobox_1 = ttk.Combobox(root, state="readonly", values=[])
column_combobox_1.grid(row=4, column=0, sticky=tk.NSEW, padx=10, pady=5)

label_groupby_2 = tk.Label(root, text="第2集計キーを選択（任意）：")
label_groupby_2.grid(row=5, column=0, sticky=tk.W, padx=10, pady=5)
column_combobox_2 = ttk.Combobox(root, state="readonly", values=[])
column_combobox_2.grid(row=6, column=0, sticky=tk.NSEW, padx=10, pady=5)

label_value_col = tk.Label(root, text="集計する値を選択（必須）：")
label_value_col.grid(row=7, column=0, sticky=tk.W, padx=10, pady=5)
value_combobox = ttk.Combobox(root, state="readonly", values=[])
value_combobox.grid(row=8, column=0, sticky=tk.NSEW, padx=10, pady=5)

select_btn_2 = tk.Button(root, text="集計開始", command=calculat_shop, width=15)
select_btn_2.grid(row=9,column=0,padx=10,pady=10)

select_btn_3 = tk.Button(root, text="別名で保存", command=save_file, width=15)
select_btn_3.grid(row=10,column=0,padx=10,pady=10)

#中央寄せ(ウィンドウサイズ変更に対応可)
#列の伸縮設定
root.grid_columnconfigure(0, weight=1) # 0列目 (ラベル側)
root.grid_columnconfigure(1, weight=3) # 1列目 (入力欄側) - こちらをより広げる

#行の伸縮設定
# for i in range(5): # 必要に応じて列数を増やす
#   root.grid_rowconfigure(i, weight=1)

root.mainloop()
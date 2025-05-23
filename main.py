import pandas as pd
import chardet
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk

# ===== 編碼偵測 =====
def detect_encoding(file_path, num_bytes=10000):
    with open(file_path, 'rb') as f:
        raw = f.read(num_bytes)
    result = chardet.detect(raw)
    return result['encoding']

# ===== 分隔符號偵測 =====
def guess_separator(file_path, encoding):
    with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
        line = f.readline()
        if '\t' in line:
            return '\t'
        elif ';' in line:
            return ';'
        elif '，' in line:
            return '，'
        elif ' ' in line:
            return ' ' #這個是空格
        else:
            return ','  # 預設使用英文逗號

# ===== 讀取簽到表並更新欄位選項 =====
def load_attend():
    path = filedialog.askopenfilename(filetypes=[("Teams匯出的參與者CSV檔", "*.csv")])
    file1_path.set(path)
    try:
        encoding = detect_encoding(path)
        sep = guess_separator(path, encoding)
        df = pd.read_csv(path, encoding=encoding, sep=sep, engine='python')
        attend_columns.set(df.columns.tolist())
        update_dropdown(attend_dropdown, attend_columns.get(), left_column)
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取簽到表失敗：\n{e}")

# ===== 讀取通訊錄並更新欄位選項 =====
def load_directory():
    path = filedialog.askopenfilename(filetypes=[("主計通訊錄Excel檔", "*.xlsx")])
    file2_path.set(path)
    try:
        df = pd.read_excel(path)
        directory_columns.set(df.columns.tolist())
        update_dropdown(directory_dropdown, directory_columns.get(), right_column)
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取通訊錄失敗：\n{e}")

# ===== 合併檔案 =====
def merge_files():
    try:
        # 自動偵測編碼與分隔符號
        encoding = detect_encoding(file1_path.get())
        sep = guess_separator(file1_path.get(), encoding)

        # 讀入兩份資料
        df1 = pd.read_csv(file1_path.get(), encoding=encoding, sep=sep, engine='python')
        df2 = pd.read_excel(file2_path.get())

        # 比對欄位
        left_col = left_column.get()
        right_col = right_column.get()

        # 合併資料
        merged = pd.merge(df1, df2, left_on=left_col, right_on=right_col, how="left")

        # 儲存輸出
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if save_path:
            merged.to_excel(save_path, index=False)
            # 🔔 改為顯示自定義成功視窗
            show_success_popup()
    except Exception as e:
        messagebox.showerror("錯誤", f"合併過程發生錯誤：\n{e}")

def show_success_popup():
    popup = tk.Toplevel(root)
    popup.title("合併完成")
    popup.geometry("300x150")
    popup.resizable(False, False)

    tk.Label(popup, text="✅ 合併成功！", font=("Arial", 14)).pack(pady=10)
    tk.Label(popup, text="檔案已成功儲存。").pack()

    # 加入超連結
    def open_link(event):
        webbrowser.open_new("https://github.com/adsa562/DGBAS_TeamsCSV-merge-AddressBookXLSX")  # 修改成你的連結

    link_label = tk.Label(popup, text="🔗Github項目連結", fg="blue", cursor="hand2")
    link_label.pack(pady=5)
    link_label.bind("<Button-1>", open_link)

    # 關閉按鈕
    tk.Button(popup, text="關閉", command=popup.destroy).pack(pady=10)

# ===== 更新下拉式選單 =====
def update_dropdown(menu_widget, options, var):
    menu = menu_widget["menu"]
    menu.delete(0, "end")
    for opt in options:
        menu.add_command(label=opt, command=lambda value=opt: var.set(value))
    if options:
        var.set(options[0])  # 預設第一個欄位

# ===== GUI 初始化 =====
root = tk.Tk()
root.title("Teams簽到表合併主計總處通訊錄工具 v1.0")

file1_path = tk.StringVar()
file2_path = tk.StringVar()

attend_columns = tk.Variable()
directory_columns = tk.Variable()
left_column = tk.StringVar()
right_column = tk.StringVar()

# ===== GUI 元件排列 =====
tk.Label(root, text="簽到表 CSV 檔案：").pack()
tk.Label(root, text="（請於Teams匯出後用excel檔開啟，並只留下姓名及電子信箱欄位再存檔）").pack()
tk.Entry(root, textvariable=file1_path, width=55).pack()
tk.Button(root, text="選擇簽到表 CSV", command=load_attend).pack(pady=5)

tk.Label(root, text="通訊錄 Excel 檔案：").pack()
tk.Label(root, text="（請於「全國主計網」－「主計總處通訊錄」當中匯出excel檔）").pack()
tk.Entry(root, textvariable=file2_path, width=55).pack()
tk.Button(root, text="選擇通訊錄 Excel", command=load_directory).pack(pady=5)

# === 簽到表欄位選擇 ===
attend_frame = tk.Frame(root)
attend_frame.pack(pady=5)
tk.Label(attend_frame, text="簽到表比對欄位：").pack(side="left")
attend_dropdown = tk.OptionMenu(attend_frame, left_column, "")
attend_dropdown.pack(side="left")

# === 通訊錄欄位選擇 ===
directory_frame = tk.Frame(root)
directory_frame.pack(pady=5)
tk.Label(directory_frame, text="通訊錄比對欄位：").pack(side="left")
directory_dropdown = tk.OptionMenu(directory_frame, right_column, "")
directory_dropdown.pack(side="left")

tk.Button(root, text="合併並儲存", command=merge_files, bg="lightgreen").pack(pady=10)

root.mainloop()

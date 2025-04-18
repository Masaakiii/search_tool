import os
import pandas as pd
import warnings
import tkinter as tk
from tkinter import filedialog, messagebox

# 忽略 openpyxl 的警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def browse_folder():
    """讓使用者選擇資料夾"""
    folder = filedialog.askdirectory()
    if folder:
        folder_path_var.set(folder)

def search_keywords():
    """執行搜尋"""
    folder_path = folder_path_var.get()
    keyword1 = keyword1_var.get()
    keyword2 = keyword2_var.get()

    if not folder_path or not keyword1:
        messagebox.showerror("錯誤", "請選擇資料夾並輸入第一個搜尋條件！")
        return

    results = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
            file_path = os.path.join(folder_path, file_name)
            try:
                # 讀取 Excel 檔案
                excel_data = pd.ExcelFile(file_path)
                
                # 遍歷每個工作表，僅搜尋名稱包含 "packing" 的工作表
                for sheet_name in excel_data.sheet_names:
                    if "packing" in sheet_name.lower():  # 忽略大小寫
                        df = excel_data.parse(sheet_name)
                        df = df.astype(str)  # 將所有值轉換為字串
                        
                        # 搜尋第一個關鍵字
                        rows_with_keyword1 = df.isin([keyword1]).any(axis=1)
                        if rows_with_keyword1.any():
                            for index, row in df[rows_with_keyword1].iterrows():
                                # 如果 keyword2 為空，直接輸出結果；否則檢查該行是否包含第二個關鍵字
                                if not keyword2 or keyword2 in row.values:
                                    results.append(
                                        f"檔案: {file_name}, 工作表: {sheet_name}, 行號: {index + 2}, "
                                        f"第一個關鍵字: {keyword1}"
                                        f"{'' if not keyword2 else f', 第二個關鍵字: {keyword2}'}"
                                    )
            except Exception as e:
                results.append(f"無法讀取檔案 {file_name}: {e}")

    # 顯示結果
    if results:
        messagebox.showinfo("搜尋結果", "\n".join(results))
    else:
        messagebox.showinfo("搜尋結果", "未找到符合條件的資料！")

# 建立主視窗
root = tk.Tk()
root.title("Excel 關鍵字搜尋工具")

# 資料夾路徑
folder_path_var = tk.StringVar()
tk.Label(root, text="選擇資料夾:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=folder_path_var, width=40).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="瀏覽", command=browse_folder).grid(row=0, column=2, padx=10, pady=5)

# 第一個關鍵字
keyword1_var = tk.StringVar()
tk.Label(root, text="第一個關鍵字:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=keyword1_var, width=40).grid(row=1, column=1, padx=10, pady=5)

# 第二個關鍵字
keyword2_var = tk.StringVar()
tk.Label(root, text="第二個關鍵字 (可選):").grid(row=2, column=0, padx=10, pady=5, sticky="e")
tk.Entry(root, textvariable=keyword2_var, width=40).grid(row=2, column=1, padx=10, pady=5)

# 搜尋按鈕
tk.Button(root, text="開始搜尋", command=search_keywords).grid(row=3, column=0, columnspan=3, pady=10)

# 啟動主迴圈
root.mainloop()
import os
import pandas as pd
import warnings
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# 忽略 openpyxl 的警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def browse_folder():
    """讓使用者選擇資料夾"""
    folder = filedialog.askdirectory()
    if folder:
        folder_path_var.set(folder)

def save_results_to_file(results):
    """將結果儲存到文字檔案"""
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")],
        title="儲存結果為文字檔"
    )
    if file_path:
        try:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(results)
            messagebox.showinfo("成功", f"結果已儲存到 {file_path}")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法儲存檔案: {e}")

def search_keywords():
    """執行搜尋"""
    folder_path = folder_path_var.get()
    keyword1 = keyword1_var.get()
    keyword2 = keyword2_var.get()

    if not folder_path or not keyword1:
        messagebox.showerror("錯誤", "請選擇資料夾並輸入第一個搜尋條件！")
        return

    grouped_results = {}
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
                                    result = f"行號: {index + 2}, 第一個關鍵字: {keyword1}"
                                    if keyword2:
                                        result += f", 第二個關鍵字: {keyword2}"
                                    
                                    # 分組結果
                                    if file_name not in grouped_results:
                                        grouped_results[file_name] = {}
                                    if sheet_name not in grouped_results[file_name]:
                                        grouped_results[file_name][sheet_name] = []
                                    grouped_results[file_name][sheet_name].append(result)
            except Exception as e:
                grouped_results[file_name] = {"錯誤": [f"無法讀取檔案: {e}"]}

    # 格式化分組結果
    if grouped_results:
        formatted_results = []
        for file_name, sheets in grouped_results.items():
            formatted_results.append(f"檔案: {file_name}")
            for sheet_name, rows in sheets.items():
                formatted_results.append(f"  工作表: {sheet_name}")
                formatted_results.extend([f"    {row}" for row in rows])
            formatted_results.append("")  # 在每個檔案的結果後新增一行空白行
        result_text.delete(1.0, tk.END)  # 清空之前的結果
        formatted_results_str = "\n".join(formatted_results)
        result_text.insert(tk.END, formatted_results_str)  # 插入新結果
        save_button.config(state=tk.NORMAL)  # 啟用儲存按鈕
    else:
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, "未找到符合條件的資料！")
        save_button.config(state=tk.DISABLED)  # 禁用儲存按鈕

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

# 結果顯示區域（可滾動）
result_text = scrolledtext.ScrolledText(root, width=80, height=20, wrap=tk.WORD)
result_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

# 儲存結果按鈕
save_button = tk.Button(root, text="儲存結果", command=lambda: save_results_to_file(result_text.get(1.0, tk.END)), state=tk.DISABLED)
save_button.grid(row=5, column=0, columnspan=3, pady=10)

# 啟動主迴圈
root.mainloop()
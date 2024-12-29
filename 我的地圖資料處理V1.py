import tkinter as tk
from tkinter import filedialog
import subprocess
import sys

# 檢查並安裝所需的套件
def install_package(package_name):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
    except subprocess.CalledProcessError as e:
        print(f"安裝 {package_name} 時發生錯誤: {e}")

# 檢查 pandas 是否已安裝，並在未安裝時安裝
try:
    import pandas as pd
except ImportError:
    print("所需的套件未安裝，正在安裝...")
    install_package('pandas')
    import pandas as pd

# 創建主視窗
root = tk.Tk()
root.title("選擇檔案介面")  # 設置窗口標題
root.geometry("700x550")  # 設置窗口大小
root.configure(bg="#f0f0f0")  # 設置背景顏色

# 設置字體樣式
title_font = ("Arial", 16, "bold")
label_font = ("Arial", 12)
button_font = ("Arial", 12, "bold")

# 顯示標題：excel合成轉換
title_label = tk.Label(root, text="excel合成轉換", font=("Arial", 20, "bold"), fg="black", bg="#f0f0f0")
title_label.grid(row=0, column=0, columnspan=3, pady=15)

# 創建一個框架來組織控件
frame = tk.Frame(root, bg="#f0f0f0")
frame.grid(row=1, column=0, columnspan=3, pady=10)

# 顯示文字
label1 = tk.Label(frame, text="選擇檔案 1：", font=label_font, bg="#f0f0f0")
label1.grid(row=0, column=0, padx=10, pady=10, sticky="w")

# 創建檔案選擇的輸入框 1
entry1 = tk.Entry(frame, width=40, font=label_font)
entry1.grid(row=0, column=1, padx=10, pady=10)

# 創建檔案選擇的按鈕 1
def open_file1():
    file_path = filedialog.askopenfilename(title="選擇檔案 1", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry1.delete(0, tk.END)  # 清空原本的內容
        entry1.insert(0, file_path)  # 在輸入框中插入檔案路徑

button1 = tk.Button(frame, text="選擇檔案", command=open_file1, font=button_font, bg="#b6c8cc", fg="white", relief="raised", width=15)
button1.grid(row=0, column=2, padx=10, pady=10)

# 顯示文字
label2 = tk.Label(frame, text="選擇檔案 2：", font=label_font, bg="#f0f0f0")
label2.grid(row=1, column=0, padx=10, pady=10, sticky="w")

# 創建檔案選擇的輸入框 2
entry2 = tk.Entry(frame, width=40, font=label_font)
entry2.grid(row=1, column=1, padx=10, pady=10)

# 創建檔案選擇的按鈕 2
def open_file2():
    file_path = filedialog.askopenfilename(title="選擇檔案 2", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry2.delete(0, tk.END)  # 清空原本的內容
        entry2.insert(0, file_path)  # 在輸入框中插入檔案路徑

button2 = tk.Button(frame, text="選擇檔案", command=open_file2, font=button_font, bg="#b6c8cc", fg="white", relief="raised", width=15)
button2.grid(row=1, column=2, padx=10, pady=10)

# 訊息框用來顯示讀取的檔案內容
message_box = tk.Text(root, height=8, width=70, font=("Courier", 10), wrap=tk.WORD, bg="#f8f8f8", fg="#333333", relief="sunken")
message_box.grid(row=2, column=0, columnspan=3, padx=20, pady=20)

# 開始按鈕的功能：讀取檔案內容並顯示在訊息框
def start_action():
    global df1, df2  # 確保使用全局變數
    file_path1 = entry1.get()
    file_path2 = entry2.get()

    if file_path1 and file_path2:
        try:
            # 使用 pandas 讀取 Excel 檔案
            df1 = pd.read_excel(file_path1)
            df2 = pd.read_excel(file_path2)

            # 顯示當前的標題行，檢查讀取的標題
            print("檔案 1 標題:", df1.columns)
            print("檔案 2 標題:", df2.columns)

            # 修改標題
            df1.columns = ['項目', '精度', '緯度'] + list(df1.columns[3:])  # 將列標題設置為前3列及其他列
            df2.columns = ['項目', '大地度', '大地緯'] + list(df2.columns[3:])  # 同上

            # 顯示檔案 1 的所有資料
            content1 = df1.to_string(index=False)  # 檔案 1 的所有資料

            # 顯示檔案 2 的資料從第一行開始，並去除 A 欄
            content2 = df2.iloc[:, 1:2].to_string(index=False)  # 去除 A 欄，保留所有資料

            # 合併檔案 1 和檔案 2（插入檔案 2 的 B, C 欄位到檔案 1）
            df_combined = pd.concat([df1.iloc[:, :3], df2.iloc[:, 1:3]], axis=1)  # 保留前三列

            # 顯示內容
            message_box.delete(1.0, tk.END)  # 清空原來的訊息
            message_box.insert(tk.END, f"合併後的資料:\n\n{df_combined.to_string(index=False)}")

            # 隱藏「開始」按鈕，並設置「生成」按鈕文字
            start_button.grid_forget()  # 隱藏開始按鈕
            generate_button.config(text="生成", command=generate_action)  # 設置按鈕文字
            generate_button.grid(row=3, column=1, pady=20)  # 顯示生成按鈕

        except Exception as e:
            message_box.delete(1.0, tk.END)
            message_box.insert(tk.END, f"讀取檔案失敗: {str(e)}")
    else:
        message_box.delete(1.0, tk.END)
        message_box.insert(tk.END, "請選擇兩個檔案！")

# 生成按鈕的功能：儲存修改後的 Excel 檔案
def generate_action():
    global df1, df2  # 確保使用全局變數
    if df1 is not None and df2 is not None:
        try:
            # 只保留前三列 (A, B, C)
            df1 = df1.iloc[:, :3]
            df2 = df2.iloc[:, 1:3]

            # 合併檔案 1 和檔案 2
            df_combined = pd.concat([df1, df2], axis=1)

            # 彈出保存檔案對話框
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                # 儲存合併後的 Excel 檔案
                df_combined.to_excel(save_path, index=False)
                message_box.delete(1.0, tk.END)
                message_box.insert(tk.END, f"檔案已儲存至：{save_path}")

                # 更改按鈕文字為 "再次生成"
                generate_button.config(text="再次生成", command=regenerate_action)
        except Exception as e:
            message_box.delete(1.0, tk.END)
            message_box.insert(tk.END, f"儲存檔案失敗: {str(e)}")

# 再次生成的功能：重新執行儲存操作
def regenerate_action():
    global df1, df2  # 確保使用全局變數
    if df1 is not None and df2 is not None:
        try:
            # 只保留前三列 (A, B, C)
            df1 = df1.iloc[:, :3]
            df2 = df2.iloc[:, 1:3]

            # 合併檔案 1 和檔案 2
            df_combined = pd.concat([df1, df2], axis=1)

            # 彈出保存檔案對話框
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                # 儲存合併後的 Excel 檔案
                df_combined.to_excel(save_path, index=False)
                message_box.delete(1.0, tk.END)
                message_box.insert(tk.END, f"檔案已儲存至：{save_path}")
        except Exception as e:
            message_box.delete(1.0, tk.END)
            message_box.insert(tk.END, f"儲存檔案失敗: {str(e)}")

# 創建「開始」按鈕
start_button = tk.Button(root, text="開始", command=start_action, font=button_font, bg="#FFAF60", fg="white", relief="raised", width=15)
start_button.grid(row=3, column=1, pady=15)

# 創建「生成」按鈕，並預設為隱藏
generate_button = tk.Button(root, text="生成", command=generate_action, font=button_font, bg="#00EC00", fg="white", relief="raised", width=15)
generate_button.grid(row=4, column=1, pady=15)
generate_button.grid_forget()  # 隱藏生成按鈕

# 顯示底部文字
footer_label = tk.Label(root, text="#功能說明:選好檔案後按下開始可以看到預期的樣子，按下生成後可以選擇位置存檔", font=("Arial", 10), bg="#f0f0f0", fg="black")
footer_label.grid(row=5, column=0, columnspan=3, pady=10)  # 放置在窗口的底部

# 啟動 GUI 主循環
root.mainloop()

# 導入所需的函式庫
import webbrowser  # 用於開啟網頁瀏覽器
import sys        # 用於系統相關操作
import tkinter as tk              # 用於創建圖形界面
from tkinter import messagebox    # 用於顯示消息框
import gspread    # 用於操作 Google Sheets
from oauth2client.service_account import ServiceAccountCredentials  # 用於 Google API 認證
from datetime import datetime     # 用於處理日期和時間
import time       # 用於時間相關操作
import google.auth.exceptions    # 用於處理 Google 認證異常
import qrcode     # 用於生成 QR 碼
import os         # 用於操作文件系統
from PIL import Image, ImageTk            # 用於圖像處理
from tkinter import ttk          # 用於創建現代風格的 GUI 元件

# 設定 Google Sheets API 存取範圍
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

try:
    # 嘗試載入 Google API 認證
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    
    try:
        # 嘗試連接指定的 Google Sheets
        sheet = client.open_by_key('1eFc-hJv-ppqnAN3tqm2tgbLj8FEGtId8HTUTLuptMHc').sheet1
        sheet.get_all_values()  # 測試是否能成功讀取資料
    except gspread.exceptions.APIError as e:
        print("試算表存取錯誤！")
        print("請確保：")
        print("1. 試算表 ID 正確")
        print("2. 服務帳號已被添加為試算表的編輯者")
        print(f"詳細錯誤：{str(e)}")
        exit(1)
    except Exception as e:
        print(f"試算表操作錯誤：{str(e)}")
        exit(1)
        
except FileNotFoundError:
    # 處理認證文件不存在的情況
    print("錯誤：找不到 credentials.json 文件！")
    print("請確保您已經下載了新的服務帳號憑證文件，並將其重命名為 credentials.json")
    exit(1)
except Exception as e:
    # 處理其他認證相關錯誤
    print(f"認證錯誤：{str(e)}")
    print("請確保：")
    print("1. credentials.json 文件是最新的")
    print("2. Google Drive API 和 Sheets API 已啟用")
    exit(1)

def generate_qr_code(url):
    """生成 QR Code"""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white")

def show_space_selection():
    """顯示空間選擇視窗的函數"""
    root = tk.Tk()
    root.title("空間選擇")
    
    # 設定視窗大小和位置
    window_width = 400
    window_height = 500  # 增加高度以容納 QR Code
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # 生成 QR Code
    form_url = "https://docs.google.com/forms/d/1zZR20jP6qJTrjCr0KP7E4Vkqfnx5RP1-QinDc0fO8oY/edit?pli=1"
    qr_image = generate_qr_code(form_url)
    
    # 轉換 QR Code 為 Tkinter 可用的格式
    qr_photo = ImageTk.PhotoImage(qr_image)
    
    # 顯示 QR Code
    qr_label = tk.Label(root, image=qr_photo)
    qr_label.image = qr_photo  # 保持參考，防止被垃圾回收
    qr_label.pack(pady=10)
    
    # 創建 UI 元件
    label = ttk.Label(root, text="請選擇空間：")
    label.pack(pady=10)
    
    # 定義可選擇的空間列表
    spaces = ["音樂專業教室", "藝術專業教室", "美學教室", "正修廳", "國際會議", "視聽設備"]
    space_var = tk.StringVar()
    space_combobox = ttk.Combobox(root, textvariable=space_var, values=spaces)
    space_combobox.pack(pady=10)
    
    # 確認按鈕的回調函數
    def on_confirm():
        selected_space = space_var.get()
        if selected_space:
            print(f"已選擇空間：{selected_space}")
            webbrowser.open(form_url)
            root.destroy()
    
    confirm_button = ttk.Button(root, text="確認", command=on_confirm)
    confirm_button.pack(pady=10)
    
    root.mainloop()

def main():
    """主程式函數"""
    # 直接開啟表單
    show_space_selection()
    input("按任意鍵結束程式...")

# 程式入口點
if __name__ == "__main__":
    print(sys.executable)  # 印出 Python 解釋器的路徑
    main()

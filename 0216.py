# 導入所需的函式庫
import webbrowser  # 用於開啟網頁瀏覽器
import sys        # 用於系統相關操作
import cv2        # 用於攝像頭操作和圖像處理
from pyzbar.pyzbar import decode  # 用於解碼 QR 碼
import tkinter as tk              # 用於創建圖形界面
from tkinter import messagebox    # 用於顯示消息框
import gspread    # 用於操作 Google Sheets
from oauth2client.service_account import ServiceAccountCredentials  # 用於 Google API 認證
from datetime import datetime     # 用於處理日期和時間
import time       # 用於時間相關操作
import google.auth.exceptions    # 用於處理 Google 認證異常
import qrcode     # 用於生成 QR 碼
import os         # 用於操作文件系統
from PIL import Image            # 用於圖像處理
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

def scan_qr_code():
    """掃描 QR 碼的函數"""
    # 初始化攝像頭
    cap = cv2.VideoCapture(0)
    
    # 檢查攝像頭是否成功開啟
    if not cap.isOpened():
        print("錯誤：無法開啟攝像頭！")
        print("請確保：")
        print("1. 攝像頭已正確連接")
        print("2. 沒有其他程式正在使用攝像頭")
        print("3. 攝像頭驅動程式已正確安裝")
        return None
    
    # 等待攝像頭初始化
    time.sleep(2)
    
    try:
        while True:
            # 讀取攝像頭畫面
            ret, frame = cap.read()
            
            if not ret or frame is None:
                print("錯誤：無法從攝像頭讀取影像")
                break
                
            # 顯示攝像頭畫面
            cv2.imshow('QR Code Scanner', frame)
            
            # 嘗試解碼 QR 碼
            try:
                decoded_objects = decode(frame)
                for obj in decoded_objects:
                    qr_data = obj.data.decode('utf-8')
                    print(f"掃描到 QR 碼：{qr_data}")
                    cap.release()
                    cv2.destroyAllWindows()
                    return qr_data
            except Exception as e:
                print(f"QR 碼解碼錯誤：{str(e)}")
                continue
            
            # 檢查是否按下 'q' 鍵退出
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
                
    except Exception as e:
        print(f"攝像頭操作錯誤：{str(e)}")
    finally:
        # 釋放資源
        cap.release()
        cv2.destroyAllWindows()
    
    return None

def show_space_selection():
    """顯示空間選擇視窗的函數"""
    root = tk.Tk()
    root.title("空間選擇")
    
    # 設定視窗大小和位置（置中）
    window_width = 400
    window_height = 300
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # 創建 UI 元件
    label = ttk.Label(root, text="請選擇空間：")
    label.pack(pady=20)
    
    # 定義可選擇的空間列表
    spaces = ["音樂專業教室", "藝術專業教室", "美學教室", "正修廳", "國際會議"," 視聽設備"]
    space_var = tk.StringVar()
    space_combobox = ttk.Combobox(root, textvariable=space_var, values=spaces)
    space_combobox.pack(pady=10)
    
    # 確認按鈕的回調函數
    def on_confirm():
        selected_space = space_var.get()
        if selected_space:
            print(f"已選擇空間：{selected_space}")
            root.destroy()
    
    confirm_button = ttk.Button(root, text="確認", command=on_confirm)
    confirm_button.pack(pady=20)
    
    root.mainloop()
    
    return space_var.get()

def generate_and_show_qr_code(data):
    """生成並顯示 QR Code 的函數"""
    try:
        # 創建 QR 碼物件
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        
        # 設定 Google Forms 的 URL
        form_url = "https://docs.google.com/forms/d/1zZR20jP6qJTrjCr0KP7E4Vkqfnx5RP1-QinDc0fO8oY/edit?pli=1"
        qr.add_data(form_url)
        qr.make(fit=True)
        
        # 生成 QR 碼圖片
        qr_image = qr.make_image(fill_color="black", back_color="white")
        
        # 使用時間戳生成唯一的文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"qrcode_{timestamp}.png"
        qr_image.save(filename)
        print(f"QR Code 已生成並保存為：{filename}")
        
        # 顯示生成的 QR 碼圖片
        qr_image.show()
        
        return filename
        
    except Exception as e:
        print(f"生成 QR Code 時發生錯誤：{str(e)}")
        return None

def main():
    """主程式函數"""
    # 生成並顯示 QR Code
    qr_file = generate_and_show_qr_code("space_selection")
    if qr_file:
        print("QR Code 已顯示，請使用手機掃描")
        print("掃描後將在手機瀏覽器中開啟空間選擇表單")
        input("按任意鍵結束程式...")
    else:
        print("QR Code 生成失敗，程式將結束")

# 程式入口點
if __name__ == "__main__":
    print(sys.executable)  # 印出 Python 解釋器的路徑
    main()

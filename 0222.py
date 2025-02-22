# 導入所需的函式庫
import cv2
from pyzbar.pyzbar import decode
import tkinter as tk
from tkinter import messagebox
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import time
import google.auth.exceptions
import qrcode
import os
from PIL import Image  # 新增這行
from tkinter import ttk
import webbrowser  # 新增這行

# 設定 Google Sheets API 憑證
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

try:
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    
    try:
        # 嘗試打開試算表（使用名稱）
        sheet = client.open('空間設備借用紀錄').sheet1
        # 測試是否能訪問試算表
        sheet.get_all_values()
    except gspread.exceptions.SpreadsheetNotFound:
        print("錯誤：找不到指定的試算表！")
        print("請確保：")
        print("1. 試算表名稱正確")
        print("2. 服務帳號已被添加為試算表的編輯者")
        exit(1)
    except gspread.exceptions.APIError as e:
        print("試算表存取錯誤！")
        print("請確保服務帳號已被添加為試算表的編輯者")
        print(f"詳細錯誤：{str(e)}")
        exit(1)
        
except FileNotFoundError:
    print("錯誤：找不到 credentials.json 文件！")
    print("請確保您已經：")
    print("1. 從 Google Cloud Platform 下載了服務帳號憑證")
    print("2. 將憑證文件重命名為 credentials.json")
    print("3. 將文件放在程式的同一個目錄中")
    exit(1)
except Exception as e:
    print(f"認證錯誤：{str(e)}")
    print("請確保：")
    print("1. credentials.json 文件是最新的")
    print("2. Google Drive API 和 Sheets API 已啟用")
    exit(1)

def scan_qr_code():
    # 嘗試打開攝像頭
    cap = cv2.VideoCapture(0)
    
    # 檢查攝像頭是否成功打開
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
            ret, frame = cap.read()
            
            # 檢查是否成功讀取影格
            if not ret or frame is None:
                print("錯誤：無法從攝像頭讀取影像")
                break
                
            # 顯示攝像頭畫面
            cv2.imshow('QR Code Scanner', frame)
            
            # 解碼 QR 碼
            try:
                decoded_objects = decode(frame)
                for obj in decoded_objects:
                    # 獲取 QR 碼內容
                    qr_data = obj.data.decode('utf-8')
                    print(f"掃描到 QR 碼：{qr_data}")
                    cap.release()
                    cv2.destroyAllWindows()
                    return qr_data
            except Exception as e:
                print(f"QR 碼解碼錯誤：{str(e)}")
                continue
            
            # 按 'q' 退出
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
                
    except Exception as e:
        print(f"攝像頭操作錯誤：{str(e)}")
    finally:
        # 確保釋放資源
        cap.release()
        cv2.destroyAllWindows()
    
    return None

def show_space_selection():
    """顯示空間選擇視窗"""
    root = tk.Tk()
    root.title("空間選擇")
    
    # 設定視窗大小和位置
    window_width = 400
    window_height = 300
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    # 創建標籤
    label = ttk.Label(root, text="請選擇空間：")
    label.pack(pady=20)
    
    # 創建空間選擇下拉選單
    spaces = ["音樂專業教室", "藝術專業教室", "美學教室", "正修廳", "國際會議"," 視聽設備"]
    space_var = tk.StringVar()
    space_combobox = ttk.Combobox(root, textvariable=space_var, values=spaces)
    space_combobox.pack(pady=10)
    
    # 確認按鈕
    def on_confirm():
        selected_space = space_var.get()
        if selected_space:
            print(f"已選擇空間：{selected_space}")
            root.destroy()
    
    confirm_button = ttk.Button(root, text="確認", command=on_confirm)
    confirm_button.pack(pady=20)
    
    # 運行視窗
    root.mainloop()
    
    return space_var.get()

def generate_and_show_qr_code(data):
    """生成 QR Code 並顯示"""
    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        
        # 使用 Google Forms 的 URL
        # form_url = "https://docs.google.com/forms/d/e/1FAIpQLSdQPGgbVSZXsJNGN9JgxWfzXhXcXBY0GQEpqZuGg_j_XVOBhA/viewform"
        form_url = "https://docs.google.com/forms/d/e/1FAIpQLSfb13lcqegl1rkzdQDk0-QIgkkcmk6G07SRhK4InyRZhLSqlw/viewform"
        qr.add_data(form_url)
        qr.make(fit=True)
        
        qr_image = qr.make_image(fill_color="black", back_color="white")
        
        # 保存圖片
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"qrcode_{timestamp}.png"
        qr_image.save(filename)
        print(f"QR Code 已生成並保存為：{filename}")
        
        # 顯示圖片
        qr_image.show()
        
        return filename
        
    except Exception as e:
        print(f"生成 QR Code 時發生錯誤：{str(e)}")
        return None

# 主程式
def main():
    # 生成包含 Google Forms URL 的 QR Code
    qr_file = generate_and_show_qr_code("space_selection")
    if qr_file:
        print("QR Code 已顯示，請使用手機掃描")
        print("掃描後將在手機瀏覽器中開啟空間選擇表單")
        input("按任意鍵結束程式...")
    else:
        print("QR Code 生成失敗，程式將結束")

if __name__ == "__main__":
    main()

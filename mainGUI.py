import tkinter as tk
from tkinter import messagebox, ttk  # 新增 ttk 用來畫進度條
import threading                     # 新增 threading 用於背景執行
from utils.get_setting import value_of_key
from utils.script import load_into_chroma_bge_manager, update_peronsal_info_database, search_v3, filter_committee, excel_process_VBA, statistic_committee
from utils.get_setting import setting_data

def execute_mode(mode, current_plan, root_window):
    if current_plan == "產學合作":
        is_industry = True
    elif current_plan == "研究計畫":
        is_industry = False
    else:
        messagebox.showerror("錯誤", "「目前執行計畫」沒有填寫正確")
        return
    
    # ================= 1. 建立載入中視窗 =================
    loading_win = tk.Toplevel(root_window)
    loading_win.title("處理中")
    loading_win.geometry("350x150") # 稍微加高一點放文字
    loading_win.resizable(False, False)
    loading_win.transient(root_window)
    loading_win.grab_set()
    
    tk.Label(loading_win, text=f"正在執行「{mode}」...\n請稍候，這可能需要幾分鐘的時間。", pady=10).pack()
    
    # 【修改】改為 determinate 模式 (真實進度條)
    progress = ttk.Progressbar(loading_win, mode='determinate', length=280)
    progress.pack(pady=5)
    
    # 【新增】用來顯示數字進度的文字標籤
    status_label = tk.Label(loading_win, text="準備中...", fg="blue")
    status_label.pack()

    # ================= 2. 定義更新進度的回呼函式 =================
    def update_progress(current, total, desc):
        def ui_update():
            if total == 0:
                # 如果 total 是 0，代表我們只是想更新文字，進度條切換為來回跑動模式
                progress.config(mode='indeterminate')
                progress.start(15)
                status_label.config(text=desc)
            else:
                # 正常的進度條更新
                progress.config(mode='determinate')
                progress.stop() # 停止來回跑動
                progress['maximum'] = total
                progress['value'] = current
                status_label.config(text=f"處理進度 ({desc}): {current} / {total}")
        
        root_window.after(0, ui_update)

    # ================= 3. 定義任務結束的處理 =================
    def finish_task(status, title, message):
        loading_win.destroy()
        if status == "success":
            messagebox.showinfo(title, message)
        else:
            messagebox.showerror(title, message)

    # ================= 4. 定義背景任務 =================
    def background_task():
        try:
            if mode == '存入資料庫':
                # 【修改】把 update_progress 當作參數傳進去
                load_into_chroma_bge_manager(is_industry, progress_callback=update_progress)
                root_window.after(0, lambda: finish_task("success", "成功", "資料已成功存入資料庫"))
                
            elif mode == '輸出推薦委員':
                # 如果其他函式也有迴圈，也可以比照辦理傳入 progress_callback
                update_peronsal_info_database(is_industry)
                statistic_committee() 
                search_v3(is_industry) 
                filter_committee(is_industry) 
                excel_process_VBA()
                root_window.after(0, lambda: finish_task("success", "成功", "已成功輸出推薦委員"))
                
        except RuntimeError as e:
            error_msg = str(e)
            if error_msg == "server暫時無法使用":
                root_window.after(0, lambda: finish_task("error", "伺服器錯誤", "Ollama 回應超時！"))
            else:
                root_window.after(0, lambda: finish_task("error", "執行錯誤", error_msg))
                
        except Exception as e:
            root_window.after(0, lambda: finish_task("error", "未預期錯誤", f"發生錯誤：{str(e)}"))

    # ================= 5. 啟動背景執行緒 =================
    threading.Thread(target=background_task, daemon=True).start()

def create_gui():
    current_plan = value_of_key("目前執行計畫")
    
    # 建立主視窗
    window = tk.Tk()
    window.title(f"腳本執行器")
    window.geometry("300x200")

    # 標題標籤
    label = tk.Label(window, text=f"請選擇模式 (目前計畫「{current_plan}」)", font=("Arial", 14))
    label.pack(pady=20)

    # 按鈕 - 存入資料庫 (注意：這裡多傳入了 window 參數)
    button_db = tk.Button(window, text="存入資料庫(請確保記憶體足夠)", command=lambda: execute_mode('存入資料庫', current_plan, window))
    button_db.pack(pady=10)

    # 按鈕 - 輸出推薦委員 (注意：這裡多傳入了 window 參數)
    button_committee = tk.Button(window, text="輸出推薦委員", command=lambda: execute_mode('輸出推薦委員', current_plan, window))
    button_committee.pack(pady=10)

    # 啟動主視窗循環
    window.mainloop()

if __name__ == "__main__":
    create_gui()

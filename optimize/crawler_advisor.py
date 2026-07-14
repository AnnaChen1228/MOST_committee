import pandas as pd
import re
import time
import random
import difflib
from tqdm import tqdm
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

FILE_PATH = '../data/RDF_database/committee_all_education.xlsx'
OUTPUT_PATH = '../data/RDF_database/committee_all_education_with_advisor.xlsx'

def load_excel(path, pages):
    data = {}
    for page in pages:
        try:
            data[page] = pd.read_excel(path, sheet_name=page)
        except Exception as e:
            print(f"讀取 Excel 失敗: {e}")
            return None
    return data

# ==========================================
# 1. 輔助函數：字串處理與相似度計算
# ==========================================
def generate_search_query(student_name=None, school_name=None):
    query_parts = []
    if student_name is not None:
        query_parts.append(f'"{student_name}".au')
    
    if school_name:
        temp_string = " or ".join([f'"{name}".asc' for name in school_name])
        query_parts.append(f'({temp_string})') 
    
    return ' and '.join(query_parts)

def normalize_text(text):
    """將校系名稱正規化，去除干擾字眼以提升比對準確率"""
    if not isinstance(text, str) or not text:
        return ""
    text = text.replace("台", "臺")
    # 移除學校相關字眼
    text = text.replace("國立", "").replace("私立", "")
    text = text.replace("大學", "").replace("學院", "").replace("科技", "")
    # 移除系所相關字眼
    text = text.replace("研究所", "").replace("學系", "").replace("碩士班", "").replace("博士班", "").replace("系", "")
    return text.strip()

def get_similarity(a, b):
    """計算兩個字串的相似度 (0.0 ~ 1.0)"""
    norm_a = normalize_text(a)
    norm_b = normalize_text(b)
    if not norm_a or not norm_b:
        return 0.0
    return difflib.SequenceMatcher(None, norm_a, norm_b).ratio()

# ==========================================
# 2. 爬蟲主邏輯 (單筆處理)
# ==========================================
def crawler(row, driver, wait):
    name = str(row['名字']).strip()
    
    # --- 取得目標學校與系所 (用來做綜合相似度比對) ---
    target_phd_school = str(row.get('博士畢業學校', '')).strip()
    target_phd_school = "" if target_phd_school == 'nan' else target_phd_school
    target_phd_dept = str(row.get('博士畢業系所', '')).strip()
    target_phd_dept = "" if target_phd_dept == 'nan' else target_phd_dept
    target_phd_full = target_phd_school + target_phd_dept # 組合校系

    target_ms_school = str(row.get('碩士畢業學校', '')).strip()
    target_ms_school = "" if target_ms_school == 'nan' else target_ms_school
    target_ms_dept = str(row.get('碩士畢業系所', '')).strip()
    target_ms_dept = "" if target_ms_dept == 'nan' else target_ms_dept
    target_ms_full = target_ms_school + target_ms_dept # 組合校系
    
    grad_school = [s for s in [target_phd_school, target_ms_school] if s]
    query = generate_search_query(student_name=name, school_name=grad_school)

    result_data = {
        "名稱" : name,
        "現任學校": str(row.get('學校', '')).strip(),
        "博士畢業學校": "",
        "博士畢業系所": "",
        "博士指導教授": "",
        "博士畢業論文": "",
        "碩士畢業學校": "",
        "碩士畢業系所": "",
        "碩士指導教授": "",
        "碩士畢業論文": "",
        "爬取狀態": "查無資料"
    }

    try:
        # --- A. 回到搜尋頁面 ---
        try:
            cmd_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "指令查詢")))
            cmd_link.click()
        except:
            driver.get('https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/login?o=dwebmge')
            cmd_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "指令查詢")))
            cmd_link.click()

        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='hidden' and @name='qf0' and @value='_hist_']")))
        search_box = driver.find_element(By.ID, "ysearchinput0")
        search_box.clear()
        search_box.send_keys(query)
        
        search_btn = wait.until(EC.element_to_be_clickable((By.ID, "gs32search")))
        driver.execute_script("arguments[0].click();", search_btn)
        
        # --- B. 等待並解析結果 ---
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1.5) 
        
        html = driver.page_source
        if "查無符合條件的資料" in html or "沒有找到符合" in html:
            result_data["爬取狀態"] = "查無資料"
            return result_data
            
        try:
            wait.until(EC.presence_of_element_located((By.ID, "tablefmt1")))
        except:
            result_data["爬取狀態"] = "查無資料"
            return result_data
        
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        table = soup.find(id="tablefmt1")
        
        if table:
            paper_blocks = table.find_all('table', class_='tableoutsimplefmt2')
            
            if not paper_blocks:
                result_data["爬取狀態"] = "查無資料"
            else:
                best_phd_score = -1.0
                best_ms_score = -1.0
                
                for block in paper_blocks:
                    title_elem = block.find('a', class_='slink')
                    paper_title = title_elem.text.strip() if title_elem else "未找到標題"
                    
                    advisor_name = "未找到"
                    school_name = "未找到"
                    dept_name = "未找到"
                    degree_type = None
                    
                    tds = block.find_all('td', class_='std2')
                    for td in tds:
                        td_text = td.text.strip()
                        if "指導教授:" in td_text or "指導教授：" in td_text:
                            advisor_name = td_text.split(':')[1].strip() if ':' in td_text else td_text.split('：')[1].strip()
                            
                        if "／" in td_text and ("博士" in td_text or "碩士" in td_text):
                            parts = td_text.split('／')
                            if len(parts) >= 4:
                                school_name = parts[0].strip()
                                dept_name = parts[1].strip() 
                                if "博士" in parts[3]: degree_type = "博士"
                                elif "碩士" in parts[3]: degree_type = "碩士"

                    # 🌟 核心邏輯：將網頁上的「學校+系所」組合起來，與目標進行比對
                    scraped_full = school_name + dept_name
                    
                    if degree_type == "博士":
                        # 如果有提供目標校系，計算相似度；若無，預設給 0.5 分確保能記錄
                        score = get_similarity(scraped_full, target_phd_full) if target_phd_full else 0.5
                        
                        if score > best_phd_score:
                            best_phd_score = score
                            result_data["博士畢業論文"] = paper_title
                            result_data["博士指導教授"] = advisor_name
                            result_data["博士畢業系所"] = dept_name
                            result_data["博士畢業學校"] = school_name
                            
                    elif degree_type == "碩士" and target_ms_full:
                        score = get_similarity(scraped_full, target_ms_full) if target_ms_full else 0.5
                        
                        if score > best_ms_score:
                            best_ms_score = score
                            result_data["碩士畢業論文"] = paper_title
                            result_data["碩士指導教授"] = advisor_name
                            result_data["碩士畢業系所"] = dept_name
                            result_data["碩士畢業學校"] = school_name
                
                if result_data["博士畢業論文"] or result_data["碩士畢業論文"]:
                    result_data["爬取狀態"] = "成功"
                else:
                    result_data["爬取狀態"] = "未找到符合學位的論文"
        else:
            result_data["爬取狀態"] = "找不到表格"

    except Exception as e:
        html = driver.page_source
        
        # 🔍 判斷是否遇到驗證碼 (根據臺灣碩博士網常見的攔截特徵)
        if "驗證碼" in html or "請輸入圖形上的字元" in html or "captcha" in driver.current_url.lower():
            tqdm.write("\n" + "!"*50)
            tqdm.write("🚨 系統攔截：爬到一半跳出驗證碼了！")
            tqdm.write("👉 請切換到 Chrome 視窗，手動輸入驗證碼並送出。")
            tqdm.write("!"*50)
            
            # 程式會停在這裡，直到你在終端機按下 Enter
            input("\n✅ 完成手動驗證並看到搜尋畫面後，請按 [Enter] 鍵繼續...")
            
            # 標記為 Error，交給外層 main() 的重試機制，在下一回合重新爬取這筆失敗的資料
            result_data["爬取狀態"] = "Error"
        else:
            # 如果不是驗證碼，就印出簡短的錯誤訊息
            error_msg = str(e).split('\n')[0]
            tqdm.write(f"-> 發生預期外錯誤: {error_msg}")
            result_data["爬取狀態"] = "Error"

    return result_data

# ==========================================
# 3. 主程式執行區塊 (加入 5 次 Error 重試機制)
# ==========================================
def main():
    df_dict = load_excel(FILE_PATH, pages=['Sheet1']) 
    if df_dict is None: return
    df = df_dict['Sheet1'] 
    
    print("啟動瀏覽器中...")
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 10)
    
    # 用字典來儲存結果，確保最後輸出的順序與原 DataFrame 一致
    final_results = {}
    
    # 初始狀態：所有的 index 都需要爬取
    pending_indices = list(df.index)
    MAX_RETRIES = 5
    
    try:
        # --- 初始登入與手動驗證 ---
        driver.get('https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/login?o=dwebmge')
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "ysearchinput0")))
        except:
            print("\n" + "!"*50)
            print("🚨 系統攔截：偵測到首頁驗證碼！")
            print("👉 請切換到 Chrome 視窗，手動輸入驗證碼。")
            print("!"*50)
            input("\n✅ 完成手動驗證並看到搜尋畫面後，請按 [Enter] 鍵繼續...")
            wait.until(EC.presence_of_element_located((By.ID, "ysearchinput0")))
            print("🚀 驗證通過，開始批次爬取！\n")

        # 🌟 開始最多 5 回合的重試機制
        for attempt in range(1, MAX_RETRIES + 1):
            if not pending_indices:
                print("🎉 所有資料皆已成功處理完畢（無 Error）！")
                break
                
            print(f"\n=== 🔄 第 {attempt}/{MAX_RETRIES} 回合爬取 (共 {len(pending_indices)} 筆待處理) ===")
            still_pending = [] # 用來記錄這一回合依然 Error 的 index
            
            for idx in tqdm(pending_indices, desc=f"第 {attempt} 回合進度", ncols=100):
                row = df.loc[idx]
                crawl_result = crawler(row, driver, wait)
                
                # 記錄結果 (每次都會覆蓋更新，確保拿到最新的狀態)
                final_results[idx] = crawl_result
                
                # 如果依然是 Error，加入下一回合的重試名單
                if crawl_result["爬取狀態"] == "Error":
                    still_pending.append(idx)
                    
                time.sleep(random.uniform(1.5, 3.5))
            
            # 更新待處理名單，準備進入下一回合
            pending_indices = still_pending
            
            # 如果還有 Error 且還沒到最後一次，稍微休息並重整頁面避免被鎖
            if pending_indices and attempt < MAX_RETRIES:
                print(f"\n⚠️ 本回合仍有 {len(pending_indices)} 筆 Error。休息 10 秒後進行下一次重試...")
                time.sleep(10)
                # 重新回到首頁，重置 Session 狀態，這對破解連續 Error 很有幫助
                driver.get('https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/login?o=dwebmge')
                time.sleep(2)
                
    finally:
        driver.quit()
        print("\n瀏覽器已關閉。")
        
        # 將字典轉換回 List，確保順序與原本的 Excel 完全一致
        results_list = [final_results[i] for i in df.index if i in final_results]
        
        if results_list:
            result_df = pd.DataFrame(results_list)
            try:
                result_df.to_excel(OUTPUT_PATH, index=False)
                print(f"📁 爬取完成！結果已成功儲存至: {OUTPUT_PATH}")
                if pending_indices:
                    print(f"⚠️ 注意：經過 {MAX_RETRIES} 次重試後，仍有 {len(pending_indices)} 筆資料為 Error。")
            except Exception as e:
                print(f"❌ 儲存 Excel 失敗: {e}")

if __name__ == "__main__":
    main()

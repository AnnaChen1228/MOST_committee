# 爬取學歷
import pandas as pd
import requests
import re
import time
import random
from tqdm import tqdm

# : 全局變數設定
URL_BASE = 'https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do'
FILE_PATH = './data/RDF_database/commitee_uni_all.xlsx'
OUTPUT_PATH = './data/RDF_database/commitee_all_education.xlsx'

# 模擬瀏覽器標頭
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
    "Origin": "https://arspb.nstc.gov.tw",
    "Referer": "https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do?action=initSearchList&LANG=chi"
}

def load_excel(path, pages):
    data = {}
    for page in pages:
        try:
            data[page] = pd.read_excel(path, sheet_name=page)
        except Exception as e:
            print(f"讀取 Excel 失敗: {e}")
            return None
    return data

def clean_html_tags(text):
    """移除 HTML 標籤並清理空白"""
    if not text: return ""
    text = re.sub(r'<(script|style).*?>.*?</\1>', '', text, flags=re.DOTALL)
    clean = re.compile('<.*?>')
    text = text.replace('<br>', ' ').replace('<br/>', ' ').replace('&nbsp;', ' ')
    return re.sub(clean, ' ', text).strip()

def crawler(row):
    name = str(row['名字']).strip()
    work_school = str(row['學校']).strip()
    
    # 【優化】這裡註解掉，避免畫面一直捲動，只讓進度條跑
    # tqdm.write(f"搜尋: {name} ({work_school})...", end=" ")

    result_data = {
        "編號": "",  # <--- 新增欄位
        "博士畢業學校": "",
        "博士畢業系所": "",
        "碩士畢業學校": "",
        "碩士畢業系所": "",
        "爬取狀態": "查無資料"
    }

    session = requests.Session()
    
    # 1. 搜尋請求
    payload = {
        "action": "initSearchList",
        "nameChi": name,
        "organ_1": "5",
        "organDesc": work_school,
        "sex": "A",
        "LANG": "chi",
        "isSearch": "1",
        "currentPage": "1",
        "pageSize": "10"
    }

    try:
        res = session.post(URL_BASE, data=payload, headers=HEADERS)
        html_content = res.text
        
        # 2. 解析搜尋列表，抓取 rsNo
        rs_no = None
        
        table_match = re.search(r'<div class="c30Tblist2">.*?<table>(.*?)</table>', html_content, re.DOTALL)
        
        if table_match:
            table_inner = table_match.group(1)
            rows = re.findall(r'<tr.*?>(.*?)</tr>', table_inner, re.DOTALL)
            
            for row in rows:
                cols = re.findall(r'<td.*?>(.*?)</td>', row, re.DOTALL)
                if len(cols) >= 2:
                    name_col = cols[0]
                    organ_col = cols[1]
                    
                    clean_name = clean_html_tags(name_col)
                    clean_organ = clean_html_tags(organ_col)
                    
                    if name in clean_name and work_school in clean_organ:
                        rs_match = re.search(r'rsNo=([a-zA-Z0-9]+)', name_col)
                        if rs_match:
                            rs_no = rs_match.group(1)
                            result_data["編號"] = rs_no  # <--- 存入編號
                            # tqdm.write(f"-> 找到編號: {rs_no}") # 註解掉以保持畫面乾淨
                            break
        
        if not rs_no:
            # tqdm.write("-> 列表無符合")
            result_data["爬取狀態"] = "查無資料"
            return result_data

        # 3. 直接請求「主要學歷」頁面
        edu_url = f"{URL_BASE}?action=initRsm02&rsNo={rs_no}&LANG=chi"
        
        time.sleep(random.uniform(1, 2))
        
        res_detail = session.get(edu_url, headers=HEADERS)
        detail_html = res_detail.text

        # 4. 解析學歷表格
        detail_rows = re.findall(r'<tr.*?>(.*?)</tr>', detail_html, re.DOTALL)
        
        found_edu = False
        for row in detail_rows:
            clean_row = clean_html_tags(row)
            
            if "博士" in clean_row or "碩士" in clean_row:
                if "學位" in clean_row and "畢業學校" in clean_row:
                    continue

                tds = re.findall(r'<td.*?>(.*?)</td>', row, re.DOTALL)
                td_texts = [clean_html_tags(td) for td in tds]
                
                degree_type = ""
                school_str = ""
                dept_str = ""
                
                full_text = " ".join(td_texts)
                if "博士" in full_text: degree_type = "博士"
                elif "碩士" in full_text: degree_type = "碩士"
                
                for txt in td_texts:
                    if txt == degree_type: continue
                    if len(txt) < 2: continue 
                    
                    if any(k in txt for k in ["大學", "學院", "University", "Institute", "School", "College"]):
                        if not school_str:
                            school_str = txt
                    elif any(k in txt for k in ["系", "所", "學程", "Department", "Program", "Ph.D.", "Master"]):
                        if not dept_str:
                            dept_str = txt
                
                if not school_str and len(td_texts) >= 3:
                    pass 

                if degree_type == "博士":
                    result_data["博士畢業學校"] = school_str
                    result_data["博士畢業系所"] = dept_str
                    found_edu = True
                elif degree_type == "碩士":
                    result_data["碩士畢業學校"] = school_str
                    result_data["碩士畢業系所"] = dept_str
                    found_edu = True

        if found_edu:
            result_data["爬取狀態"] = "成功"
        else:
            result_data["爬取狀態"] = "未公開學歷"

    except Exception as e:
        tqdm.write(f"-> Error: {e}") # 錯誤還是要印出來
        result_data["爬取狀態"] = "Error"

    return result_data

def main():
    print("讀取 Excel 中...")
    sheet_name = 'Sheet1'
    excel_data = load_excel(FILE_PATH, [sheet_name])
    
    if not excel_data:
        return

    df = excel_data[sheet_name]
    results = []
    
    stats = {
        "成功": 0,
        "查無資料": 0,
        "未公開學歷": 0,
        "Error": 0
    }

    # 使用 tqdm 顯示進度條，並設定 ncols 讓寬度固定
    for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="爬取進度", ncols=100):
        crawl_result = crawler(row)
        
        status = crawl_result["爬取狀態"]
        if status in stats:
            stats[status] += 1
        else:
            stats["查無資料"] += 1

        row_dict = row.to_dict()
        row_dict.update(crawl_result)
        results.append(row_dict)
        
        time.sleep(random.uniform(2, 4))
    
    print(f"\n存檔中... {OUTPUT_PATH}")
    result_df = pd.DataFrame(results)
    
    # 欄位排序：加入「編號」
    cols_order = ['名字', '學校', '編號', '博士畢業學校', '博士畢業系所', '碩士畢業學校', '碩士畢業系所', '爬取狀態']
    final_cols = [c for c in cols_order if c in result_df.columns] + [c for c in result_df.columns if c not in cols_order]
    
    result_df = result_df[final_cols]
    result_df.to_excel(OUTPUT_PATH, index=False)
    
    print("\n" + "="*30)
    print("【爬取結果統計報告】")
    print(f"總筆數: {len(df)}")
    print(f"✅ 成功抓取學歷: {stats['成功']} 人")
    print(f"❌ 查無此人(搜尋無結果): {stats['查無資料']} 人")
    print(f"⚠️ 未公開學歷(有編號但無資料): {stats['未公開學歷']} 人")
    if stats['Error'] > 0:
        print(f"🚫 發生錯誤: {stats['Error']} 人")
    print("="*30 + "\n")
    print("完成！")

if __name__ == "__main__":
    main()

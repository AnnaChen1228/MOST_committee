import pandas as pd

def load_data(file_path, pages):
    authors = [] 
    print(f"讀取 Excel: {file_path}")
    
    for page in pages:
        try:
            # 讀取 Excel
            project_df = pd.read_excel(file_path, sheet_name=page, dtype=str).fillna("")
        except Exception as e:
            print(f"跳過頁面 {page}: {e}")
            continue

        for index, row in project_df.iterrows():
            author = str(row['計畫主持人']).strip()
            department_full = str(row['機關名稱']).strip()
            job_title = str(row['職稱']).strip()
            # 初始化變數，避免找不到關鍵字時報錯
            school, department = filiter_school(department_full)
            authors.append({
                '名字': author,
                '學校': school,
                '職稱': job_title,
                '系所': department,
                '年份': str(row['年份']).strip() if '年份' in row.keys() else page
            })
            
    return authors

def add_data(authors,author_info, year):
    school, department = filiter_school(author_info['school'])
    authors.append({
        '名字': author_info['name'],
        '職稱': author_info['job_title'],
        '學校': school,
        '系所': department,
        '年份': year
    })
    return authors

def filiter_school(department_full):
    school = department_full
    department = ""
    keywords = ['大學', '院', '博物館', '學校', '法人','中心', '分校']  # 切割關鍵字
    for keyword in keywords:
        if keyword in department_full:
            # split(keyword, 1) 確保只切分第一個出現的關鍵字
            parts = department_full.split(keyword, 1)
            if len(parts) == 2:
                school = parts[0] + keyword  # 將關鍵字加回學校名稱中
                department = parts[1].strip() # 去除系所前後空白
                break
    return school, department


def main():
    data_file = 'data/industry_coop/pass_project_with_abstract_108-114.xlsx'
    years = ['108-114']
    research_store_file_path = 'data/research_proj/115計算機學門審查/pass_project_with_abstract.xlsx'
    research_years = ['108','109','110','111','112','113','114']
    # 1. 獲取資料列表
    result_data = load_data(data_file, years)
    research_result_data = load_data(research_store_file_path, research_years)
    result_data.extend(research_result_data)
    if result_data:
        # 2. 轉換為 DataFrame
        df_result = pd.DataFrame(result_data)
        df_result = df_result.sort_values(by=['名字', '年份'])
        # 3. 指定輸出檔名
        output_filename = './data/RDF_database/commitee_all.xlsx'
        uni_output_filename = './data/RDF_database/commitee_uni_all.xlsx'
        # 4. 存成 Excel (index=False 代表不存入最左邊的 0,1,2... 索引編號)
        df_result.to_excel(output_filename, index=False)
        df_latest = df_result.sort_values(by=['名字', '年份'], ascending=[True, False])
        
        # 針對 '名字' 去除重複，keep='first' 代表保留排序後的第一筆 (也就是年份最大的那一筆)
        df_latest = df_latest.drop_duplicates(subset=['名字'], keep='first')
        df_latest.to_excel(uni_output_filename, index=False)
        
    else:
        print("未讀取到任何資料")

if __name__ == '__main__':
    main()

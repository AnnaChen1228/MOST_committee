import pandas as pd
from typing import List
from utils.get_setting import setting_data, print_setting_data, find_key_path, value_of_key

def get_project_df() -> List[pd.DataFrame]:
    
    # : Field
    apply_project_folder_path = find_key_path("計畫過去申請案件")
    apply_project_file_year = value_of_key("計畫過去申請案件年分範圍")
    statistic_folder_path = find_key_path("統計清單")
    
    # : Variable
    apply_project_excel_file = pd.ExcelFile(apply_project_folder_path)
    
    result_df = {}
    for year in apply_project_file_year:
        
        # - Load the Apply_Project file.
        apply_project_df = pd.read_excel(apply_project_excel_file, year)
        # apply_project_df.columns = apply_project_df.iloc[0]
        
        # - Load the 
        statistic_excel_file = pd.ExcelFile(statistic_folder_path)
        statistic_df = pd.read_excel(statistic_excel_file, f"{year}總計畫清單")
        
        pass_proj_name_list = statistic_df['計畫中文名稱'].to_list()
        
        # - Combine the Apply_Project and Admission_List(Statistic)
        pass_list = []
        for i in range(len(apply_project_df)):
            if apply_project_df.iloc[i]['計畫中文名稱'] in pass_proj_name_list:
                pass_list.append('true')
            else:
                pass_list.append('false')
                
        apply_project_df['通過'] = pass_list  
        result_df[year] = apply_project_df
    all_project_file_with_pass = 'data/research_proj/all_project.xlsx'

    # 確保至少有一個工作表被創建
    with pd.ExcelWriter(all_project_file_with_pass) as writer:
        sheets_created = False
        
        # 嘗試為每個年份創建工作表
        for year in apply_project_file_year:
            try:
                if year in result_df:
                    result_df[year].to_excel(writer, sheet_name=str(year), index=False)
                    sheets_created = True
                    print(f"已創建 {year} 年度的工作表")
                else:
                    print(f"警告: result_df 中找不到 {year} 年度的資料")
            except Exception as e:
                print(f"處理 {year} 年度時出錯: {e}")
        
        # 如果沒有創建任何工作表，則創建一個空白工作表
        if not sheets_created:
            print("沒有找到任何年度的資料，創建一個空白工作表")
            pd.DataFrame().to_excel(writer, sheet_name="空白工作表", index=False)
    print("所有年度的資料已成功寫入到 all_project.xlsx")
    return result_df

def get_industry_coop_proj():
    # 產學計劃
    industry_folder_path = find_key_path("產學過去申請名冊")
    print("industry_folder_path", industry_folder_path)
    xls = pd.ExcelFile(industry_folder_path)
    
    year = ['專題計畫綜合查詢']
    df_list = {}
    for y in year:
        df1 = pd.read_excel(xls, y)
        df_list[y] = df1

    return df_list

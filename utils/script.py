from itertools import chain
import json
import shutil
import tqdm
import chromadb
import re
import pandas as pd
import numpy as np
import ast
import sys
import os
os.environ["TOKENIZERS_PARALLELISM"] = "false"

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from utils.cal_embedding_bge_zh import calculate_docs_embedding_zh, get_embeddings_zh
from utils.load_source_excel import get_project_df, get_industry_coop_proj
from langchain_community.vectorstores.chroma import Chroma
from utils.load_former_manager import get_former_manager

from utils.get_setting import setting_data, print_setting_data, find_key_path, value_of_key
from utils.filter_method import *
from utils.store_vectordb import store_basic_vector_db, store_abstract_vector_db
from utils.generate_abstract import local_generate 
class MissingFieldsException(Exception):
    pass

def load_into_chroma_bge_manager(is_industry=False, progress_callback=None):
    # 讀取黑名單
    blacklist_csv_path = find_key_path("退休或黑名單委員")
    blacklist_df = pd.read_csv(blacklist_csv_path, encoding='utf-8')
    blacklist = blacklist_df['姓名'].tolist()
    print(f"黑名單或是已退休委員:", blacklist)
    
    # 決定資料庫路徑
    if is_industry:
        chroma_basic_path = find_key_path('CHROMA_BASIC_INDUSTRY')
        chroma_abstract_path = find_key_path('CHROMA_ABSTRACT_INDUSTRY')
        db_type = 'industry'
    else:
        chroma_basic_path = find_key_path('CHROMA_BASIC')
        chroma_abstract_path = find_key_path('CHROMA_ABSTRACT')
        db_type = 'research'

    # ================= 新增：清空舊的 Collection =================
    print("正在清理舊的向量資料庫...")
    if os.path.exists(chroma_basic_path):
        try:
            shutil.rmtree(chroma_basic_path)
            print(f"已成功刪除並重置資料夾: {chroma_basic_path}")
        except Exception as e:
            print(f"刪除 {chroma_basic_path} 失敗: {e}")

    # 清理 Abstract 資料庫
    if os.path.exists(chroma_abstract_path):
        try:
            shutil.rmtree(chroma_abstract_path)
            print(f"已成功刪除並重置資料夾: {chroma_abstract_path}")
        except Exception as e:
            print(f"刪除 {chroma_abstract_path} 失敗: {e}")
    # ============================================================

    # 取得對應的 Dataframes
    df_list = get_industry_coop_proj() if is_industry else get_project_df()

    project_data = {} # 用來存放給 store_basic_vector_db 和 store_abstract_vector_db 的資料
    manager_group = {} # 保留原本產出 JSON 的邏輯 (依主持人分組)

    for key, year_data in df_list.items():
        # key 通常是年度 (pass_year)
        # for i in tqdm.tqdm(range(len(year_data)), desc=str(key)):
        total_items = len(year_data)
        for i in tqdm.tqdm(range(len(year_data)), desc=key):
            # 【關鍵加入】如果 GUI 有傳入進度更新函式，就呼叫它來更新畫面
            if progress_callback:
                # 傳入：目前第幾筆(i+1), 總共幾筆, 目前的年份(key)
                progress_callback(i + 1, total_items, str(key))
            row = year_data.iloc[i]
            
            manager = str(row.get('計畫主持人', '')).strip()
            if not manager or manager.lower() == 'nan':
                continue
                
            project_name = str(row.get('計畫中文名稱', '') or '').strip()
            abstract = str(row.get('中文摘要', '') or '').strip()
            keywords = str(row.get('中文關鍵字', '') or '').strip()
            approved = str(row.get('通過', 'false')).strip().lower()
            
            # 產學資料以外，跳過未通過的計畫
            if not is_industry and approved != 'true':  
                continue
            
            # 跳過黑名單
            if manager in blacklist:
                continue
            
            # 呼叫 LLM 進行摘要拆解
            parsed_abstract = {}
            if abstract and abstract.lower() != 'nan':
                try:
                    # 呼叫您寫好的 local_generate
                    response = local_generate(abstract)
                    # 將 LLM 回傳的 JSON 字串轉為 Python 字典
                    parsed_abstract = json.loads(response.message.content)
                except json.JSONDecodeError:
                    print(f"警告: LLM 回傳的格式非有效 JSON，計畫名稱: {project_name}")
                except RuntimeError as e:
                    # 【關鍵修改】如果捕捉到伺服器超時的 RuntimeError，直接往上拋出！
                    # 這樣就會立刻中斷迴圈，並一路傳遞到 execute_mode 顯示錯誤視窗
                    raise e
                except Exception as e:
                    print(f"解析摘要發生錯誤，計畫名稱: {project_name}, 錯誤: {e}")

            # 準備寫入向量資料庫的資料格式 (以 project_name 為 key)
            if project_name:
                project_data[project_name] = {
                    'host_name': manager,
                    'pass_year': key,
                    'title': project_name,
                    'keywords': keywords,
                    'application_directions': parsed_abstract.get('application_directions', ''),
                    'problems_to_solve': parsed_abstract.get('problems_to_solve', ''),
                    'goals_to_achieve': parsed_abstract.get('goals_to_achieve', ''),
                    'methods_to_solve': parsed_abstract.get('methods_to_solve', '')
                }
            
            # 保留原本的 Manager Group 邏輯，用於輸出傳統的 JSON 檔案
            text_for_json = f"{project_name} {abstract} {keywords}\n"
            if manager in manager_group:
                manager_group[manager] += f"\n{text_for_json}"
            else:
                manager_group[manager] = text_for_json

    # 呼叫您自定義的儲存函式 (LangChain Chroma Wrapper)
    print(f"\n開始將 Basic 資料寫入 ChromaDB ({db_type})...")
    if progress_callback:
        # 傳入 0, 0，並寫上提示文字
        progress_callback(0, 0, f"正在將 Basic 資料寫入資料庫...")
        
    store_basic_vector_db(db_type, project_data)

    print(f"開始將 Abstract 資料寫入 ChromaDB ({db_type})...")
    if progress_callback:
        progress_callback(0, 0, f"正在將 Abstract 資料寫入資料庫...")
        
    store_abstract_vector_db(db_type, project_data)

    # 決定 JSON 儲存路徑並儲存 (保留原本的 BGE_MANAGER JSON 輸出)
    bge_manager_key = "BGE_INDUSTRY_MANAGER" if is_industry else "BGE_MANAGER"
    bge_manager_path = find_key_path(bge_manager_key)
        
    with open(bge_manager_path, 'w', encoding='utf-8') as f:
        json.dump(manager_group, f, ensure_ascii=False)
    print(f"已成功儲存主持人 JSON 至: {bge_manager_path}")

    pass_data_store_key = "PASS_INDUSTRY" if is_industry else "PASS_RESEARCH"
    pass_data_store_path = find_key_path(pass_data_store_key)
    # 1. 先將所有資料轉成 DataFrame
    df = pd.DataFrame(list(project_data.values()))

    # 2. 使用 ExcelWriter 來建立多個 Sheet
    # engine='openpyxl' 是寫入 .xlsx 必備的引擎
    with pd.ExcelWriter(pass_data_store_path, engine='openpyxl') as writer:
        
        # 3. 依照 'pass_year' 進行分組 (groupby)
        for year, group_df in df.groupby('pass_year'):
            
            # 確保 Sheet 名稱是字串 (例如 "112", "113")
            sheet_name = str(year)
            
            # 將該年份的資料寫入對應的 Sheet 中
            group_df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"已成功儲存通過案件資料至: {pass_data_store_path}")

def search_v3(is_industry=False):
    
    # 取得計畫相關的欄位值
    tabs = value_of_key("計畫SHEET")

    # 取得輸出 Excel 檔案的資料夾路徑
    output_excel_folder_path = find_key_path("統計表分析")

    # 判斷「研究計畫」與「產學合作」的資料庫路徑
    if is_industry:
        chroma_db_path = find_key_path('CHROMA_INDUSTRY')
        vectorstore = Chroma("CHROMA_INDUSTRY", persist_directory=chroma_db_path, embedding_function=get_embeddings_zh())
        excel_folder_path = find_key_path("產學合作申請名冊")
    else:
        chroma_db_path = find_key_path('CHROMA')
        vectorstore = Chroma("CHROMA", persist_directory=chroma_db_path, embedding_function=get_embeddings_zh())
        excel_folder_path = find_key_path("研究計畫申請名冊")
        
    former_manager = get_former_manager(find_key_path("曾任委員"))
    
    # 要輸出的欄位
    other_fields = value_of_key("計畫相關其他欄位")
    required_fields = []
    keys = [
        "計畫名稱",
        "中文關鍵字",
        "計劃摘要",
        "職稱",
        "申請主持人欄位名稱",
        "申請機構欄位名稱",
        
    ]
    for key in keys:
        temp_value = value_of_key(key)
        if temp_value is not None and temp_value != '':
            required_fields.append(temp_value)
        
    multi_keys = [
        "申請共同主持人"
    ]
    for key in multi_keys:
        temp_value = value_of_key(key)
        for value in temp_value:
            if value not in required_fields and value != '':
                required_fields.append(value)
    
    filter_fields = required_fields + [f for f in other_fields if f not in required_fields]

    RECOMMAND_AMOUNT = 10   # 要推薦的委員數量
    SELECT_AMOUNT = 3       # 要選擇的委員數量
    SELECT_BOX_SYMBOL = ['Y', 'Z', 'AA']

    xls = pd.ExcelFile(excel_folder_path)
    writer = pd.ExcelWriter(output_excel_folder_path, engine='openpyxl')
    
    similarity_record_path = f"./data/output/similarity_record_{value_of_key('FINAL_COMMITTEE')}"
    os.makedirs(os.path.dirname(similarity_record_path), exist_ok=True)
    similarity_df = pd.DataFrame(columns=["query_text", "compared_text", "recommended_manager", "model_name", "similarity_score"])
    
    try:
        project_name_field_name = value_of_key("計畫名稱")
        chinese_keyword_field_name = value_of_key("中文關鍵字")
        abstract_field_name = value_of_key("計劃摘要")
        
        for tab in tabs:
            page_manager_list = []

            # define column name
            df = pd.read_excel(xls, tab)
            
            # 檢查 filter_fields 是否在現有的欄位中
            existing_fields = df.columns.tolist()
            missing_fields = [field for field in filter_fields if field not in existing_fields]
            if missing_fields:
                print("現有欄位:", existing_fields)
                print("應當欄位:", filter_fields)
                raise ValueError("欄位不匹配，程式碼運行停止")  # 引發例外，中止程式碼運行

            df = df[filter_fields]

            for i in range(RECOMMAND_AMOUNT):
                df['推薦委員' + str(i + 1)] = ''
                df['相關分數' + str(i + 1)] = ''
            df['前任委員占比'] = ''
            for i in range(SELECT_AMOUNT):
                df['選取委員' + str(i + 1)] = ''

            # process data
            for i in tqdm.tqdm(range(len(df)), desc=tab):
                manager_list = []
                project_name = df.iloc[i].get(project_name_field_name, '')
                keywords = df.iloc[i].get(chinese_keyword_field_name, '')
                abstract = df.iloc[i].get(abstract_field_name, '')
                
                # 找尋相似度
                current_text_combine = f"{project_name} {keywords} {abstract}"
                documents = vectorstore.similarity_search_with_relevance_scores(
                    current_text_combine,
                    k=RECOMMAND_AMOUNT
                )
                
                # 將搜尋結果寫入 CSV
                for doc, score in documents:
                    recommended_manager = doc.metadata['manager'] 
                    compared_text = doc.page_content  # 可根據需求選擇不同內容
                    model_name = "BGE_ZH"  # 依實際使用的模型名稱
                    
                    # 追加數據
                    new_row = pd.DataFrame([{
                        "query_text": current_text_combine,
                        "compared_text": compared_text,
                        "recommended_manager": recommended_manager,
                        "model_name": model_name,
                        "similarity_score": score
                    }])
                    
                    new_row = new_row.reindex(columns=similarity_df.columns)
                    if not new_row.empty and not new_row.isna().all(axis=None):
                        similarity_df = pd.concat([similarity_df, new_row], ignore_index=True)

        
                # 分數填入 Excel 的動作 (和原程式邏輯相同)
                for j, (doc, score) in enumerate(documents):
                    df.loc[df.index[i], '推薦委員' + str(j + 1)] = doc.metadata['manager']
                    manager_list.append(doc.metadata['manager'])
                    df.loc[df.index[i], '相關分數' + str(j + 1)] = score

                page_manager_list.append(manager_list)
                df.loc[df.index[i], '前任委員占比'] = len([x for x in manager_list if x in former_manager]) / RECOMMAND_AMOUNT

            df.to_excel(writer, sheet_name=tab, index=False)

            # setup dropdown list
            workbook = writer.book
            worksheet = workbook[tab]

            for j in range(SELECT_AMOUNT):
                for i, manager_list in enumerate(page_manager_list):
                    data_range = ','.join(manager_list)
                    dv = DataValidation(type="list", formula1=f'"{data_range}"', allow_blank=True)
                    dv.add(SELECT_BOX_SYMBOL[j] + str(i + 2))
                    worksheet.add_data_validation(dv)

            highligh_former_manager(writer, tab, former_manager, output_excel_folder_path)
            draw_color_for_similarity_score(writer, tab, output_excel_folder_path)

    except Exception as e:
        if not writer.book.sheetnames:
            print("ERROR")
            writer.book.create_sheet(title="Error")
        raise  # 重新引發異常以停止程式

    finally:
        # 確保 ExcelWriter 正常關閉
        writer.close()  
        with pd.ExcelWriter(similarity_record_path, engine='openpyxl') as similarity_writer:
            similarity_df.to_excel(similarity_writer, sheet_name="Similarity Records", index=False)
            
            
def draw_color_for_similarity_score(writer, tab, output_excel):
    
    from openpyxl.formatting.rule import ColorScaleRule
    
    SIMILARITY_SCORE_RANGE = '$E$2:$w$1000'
    workbook = writer.book
    worksheet = workbook[tab]
    rule = ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F9F900")
    worksheet.conditional_formatting.add(SIMILARITY_SCORE_RANGE, rule)
    workbook.save(output_excel)

def highligh_former_manager(writer, tab, former_manager, output_excel):
    
    from openpyxl.formatting import Rule
    from openpyxl.styles.differential import DifferentialStyle
    
    RECOMMAND_MANAGER_SYMBOL = ['D','F','H','J','L','N','P','R','T','V']
    workbook = writer.book
    worksheet = workbook[tab]
    redFill = PatternFill(start_color='FFA042', end_color='FFA042', fill_type='solid')

    for s in RECOMMAND_MANAGER_SYMBOL:
        col1 = worksheet[s]
        for i, cell in enumerate(col1):
            cell_value = cell.value
            if cell_value in former_manager:
                rule = Rule(type="cellIs", operator="equal", formula=[f'"{cell_value}"'], dxf=DifferentialStyle(fill=redFill))
                # rule = Rule(type="expression", operator="equal", formula=[f'"{cell_value}"'], dxf=DifferentialStyle(fill=redFill))
                worksheet.conditional_formatting.add(f'{s}{str(i+1)}', rule)

    workbook.save(output_excel)

def statistic_committee():
    
    apply_project_file_year = value_of_key("計畫過去申請案件年分範圍")
    
    statistic_folder_path = find_key_path("統計清單") 
    statistic_excel_file = pd.ExcelFile(statistic_folder_path)
    
    industry_folder_path = find_key_path("產學過去申請名冊")
    industry_data = pd.read_excel(industry_folder_path)
    
    past_apply_project_path = find_key_path("計畫過去申請案件")
    past_apply_project_file = pd.ExcelFile(past_apply_project_path)
    
    newest_person_path = find_key_path("暫存最新人才資料庫")
    newest_person_df = pd.read_excel(newest_person_path)
    
    # committee_study_folder_path = find_key_path("碩博士論文_RDF")
    # committee_study_data = pd.read_excel(committee_study_folder_path)
    
    #@ 處理委員的所有相關學校名單: 名稱 - 年份 - 學校 - 職稱
    committee_person_RDF = []
    
    # #: 碩博士論文
    # for index, row in tqdm.tqdm(committee_study_data.iterrows(), desc="碩博士論文"):
    #     # 先檢查欄位是否存在
    #     name = row.get('學生姓名', '')
    #     year = row.get('畢業學年度', '') if '畢業學年度' in row else ''
    #     institution = row.get('畢業學校', '') if '畢業學校' in row else ''

    #     # 確保數據不為 NaN（轉為字串或預設值）
    #     year = str(year) if pd.notna(year) else ''
    #     institution = str(institution) if pd.notna(institution) else ''

    #     committee_person_RDF.append({
    #         '名稱': name,
    #         '年份': year,
    #         '機關名稱': institution,
    #         '職稱': ""
    #     })
    
    #: 暫存最新人才資料庫
    for index, row in newest_person_df.iterrows():
        committee_person_RDF.append({
            '名稱': row['名稱'],
            '年份': int(row['年份']),
            '機關名稱': row['機關名稱'],
            '職稱': row['職稱'],
            "來源": row['來源']
        })

    #: 研究計劃（申請案件）
    for year in apply_project_file_year:
        current_sheet = year
        past_apply_project_df = pd.read_excel(past_apply_project_file, current_sheet)
        for index, row in tqdm.tqdm(past_apply_project_df.iterrows(), desc=f"{current_sheet}"):
            committee_person_RDF.append({
                '名稱': row['計畫主持人'],
                '年份': year,
                '機關名稱': row['機關名稱'],
                '職稱': row.get('職稱', '教授'),
                "來源": f"研究計劃（申請案件）- {current_sheet}"
            })
        
        
    #: 研究計劃（統計案件）
    for year in apply_project_file_year:
        current_sheet = f"{year}總計畫清單"
        statistic_df = pd.read_excel(statistic_excel_file, current_sheet)
        for index, row in tqdm.tqdm(statistic_df.iterrows(), desc=f"{current_sheet}"):
            committee_person_RDF.append({
                '名稱': row['計畫主持人'],
                '年份': year,
                '機關名稱': row['機關名稱'],
                '職稱': row.get('職稱', '教授'),
                "來源": f"研究計劃（統計案件）- {current_sheet}"
            })
            
    #: 產學合作
    for index, row in industry_data.iterrows():
        committee_person_RDF.append({
            '名稱': row['計畫主持人'],
            '年份': row["計畫編號"][:3] if not pd.isna(row["計畫編號"]) else "",
            '機關名稱': row['單位名稱'],
            '職稱': row['職稱'],
            "來源": f"產學合作 - 序號:{row['序號']}"
        })
    
    committee_person_RDF_df = pd.DataFrame(committee_person_RDF)
    mask = (committee_person_RDF_df['機關名稱'].notna()) & (committee_person_RDF_df['機關名稱'] != '')
    committee_person_RDF_df = committee_person_RDF_df[mask]
    committee_person_RDF_df[['學校', '系所']] = committee_person_RDF_df['機關名稱'].apply(split_institution)
    committee_person_RDF_df = committee_person_RDF_df.sort_values(by=["名稱"])
    committee_person_RDF_df.to_excel(find_key_path("統計清單人才資料_RDF"), index=False)
    
    from utils.filter_method import extract_max_year
    
    committee_person_RDF_df['年份'] = committee_person_RDF_df['年份'].apply(extract_max_year)
    committee_person_RDF_df['年份'] = committee_person_RDF_df['年份'].fillna(0) 
    committee_person_RDF_df['年份'] = committee_person_RDF_df['年份'].astype(int)
    
    unique_person_RDF_df = committee_person_RDF_df.loc[committee_person_RDF_df.groupby('名稱')['年份'].idxmax()]
    unique_person_RDF_df.to_excel(find_key_path("統計清單人才資料_RDF_UNI"), index=False)

def filter_committee(is_industry=False):
    
    #: Load the data
    crawler_RDF_folder_path = find_key_path("碩博士論文_RDF")
    crawler_RDF_data = pd.read_excel(crawler_RDF_folder_path)
    
    statistical_analysis_folder_path = find_key_path("統計表分析") 
    statistical_analysis_file = pd.ExcelFile(statistical_analysis_folder_path)
    
    if is_industry: apply_list_folder_path = find_key_path("產學合作申請名冊")
    else: apply_list_folder_path = find_key_path("研究計畫申請名冊") 
        
    apply_list_file = pd.ExcelFile(apply_list_folder_path)
    
    committee_person_path = find_key_path("統計清單人才資料_RDF_UNI")
    committee_person_RDF_df = pd.read_excel(committee_person_path)
    
    #- Strategy
    writer = pd.ExcelWriter(find_key_path("過濾相近後統計表"), engine='openpyxl')
    
    #@ 審查委員不能與計劃申請學校有關
    for sheet in statistical_analysis_file.sheet_names:
        current_sheet_statistical_excel_data = pd.read_excel(statistical_analysis_file, sheet_name=sheet)
        result_dict = []
        all_apply_members = current_sheet_statistical_excel_data[value_of_key("申請主持人欄位名稱")].to_list() # 所有申請人
        
        for index, statistical_row in current_sheet_statistical_excel_data.iterrows():
        # ~ 每個統計表的 row.
        
            # = 審查委員的背景
            committee_person_dict = []
            for index_of_committee in range(1, 11):
            # ~ 推薦委員 10 人
                
                # 委員過去待過的學校
                been_list = []
                find_temp_df = committee_person_RDF_df[committee_person_RDF_df["名稱"] == statistical_row[f'推薦委員{index_of_committee}']]
                for index, row in find_temp_df.iterrows(): 
                    been_list.append(row["學校"])
                been_list = list(set(been_list)) 
                
                # 委員過去畢業的學校
                graduate_list = []
                relate_school = find_crawler_person_relative_school(f'推薦委員{index_of_committee}', crawler_RDF_data)
                graduate_list.extend(list(set(relate_school)))
                
                # 委員職稱
                committee_person_dict.append({
                    "委員名稱": statistical_row[f'推薦委員{index_of_committee}'],
                    "委員曾就職學校": been_list,
                    "委員過去畢業學校": graduate_list,
                    "職稱": find_temp_df['職稱']
                })
                
            # = 申請學校 + 主持人學校 + 共同主持人學校
            '''
            如果有重疊，就會被歸類到 Filtered Members 裏
            反之則留在 Remaining Members
            至於 Filter Reasons 則紀錄每個被篩掉委員的具體原因
            '''
            total_committee_person_dict_result = {
                'Filtered Members': [],
                'Remaining Members': [],
                'Filter Reasons': {}
            }
            
            for sheet in apply_list_file.sheet_names: 
                current_sheet_apply_excel_data = pd.read_excel(apply_list_file, sheet_name=sheet)
                find_temp_df = current_sheet_apply_excel_data[
                    current_sheet_apply_excel_data[value_of_key("計畫名稱")] == statistical_row[value_of_key("計畫名稱")]
                ]
                
                for index, row in find_temp_df.iterrows(): 
                    joint_person_list = row.get(value_of_key("申請共同主持人"), pd.Series()).tolist()
                    joint_department_list = row.get(value_of_key("申請共同機構欄位名稱"), pd.Series()).tolist() 
                    common_person_dict = extract_text_in_parentheses(joint_person_list)
                    common_department_dict = extract_text_in_parentheses(joint_department_list)
                    
                    common_joint_list = common_person_dict + common_department_dict
                    
                    # 找到關聯性
                    project_manager_school = list([find_crawler_person_relative_school(name, crawler_RDF_data) for name, department in common_joint_list])
                    apply_school = {
                        "申請人名稱": row.get(value_of_key("申請主持人欄位名稱"), ''),
                        "申請人職稱": row.get(value_of_key("職稱"), ''),
                        "計畫申請學校": split_institution(row.get(value_of_key("申請機構欄位名稱"), ''))[0], 
                        "共同計畫主持的學校": [split_institution(department)[0] for name, department in common_joint_list],
                        "計畫主持人過去畢業的學校": find_crawler_person_relative_school(value_of_key("申請主持人欄位名稱"), crawler_RDF_data),
                        "共同主持人過去畢業的學校": list(chain.from_iterable(project_manager_school)), 
                    }
                    
                    #~ 審查委員不能與計劃申請學校(包含共同主持人)有關
                    # print("apply_school\n", apply_school)
                    # print("committee_person_dict\n", committee_person_dict)
                    filter_pairs = [("計畫申請學校", "委員曾就職學校"), ("共同計畫主持的學校", "委員曾就職學校")]
                    TITLE_RESTRICTIONS = {
                        "助理教授": ["教授", "研究員"], # 助理教授不能審教授或研究員
                        "助研究員": ["教授", "研究員"] # 助研究員不能審教授或研究員
                    }
                    current_committee_person_dict_result = filter_committee_advanced(
                        apply_school, 
                        committee_person_dict, 
                        filter_pairs, 
                        all_apply_members, 
                        TITLE_RESTRICTIONS,
                        whether_to_execute_the_option= {
                            "是否過濾申請人": True if is_industry else False,
                            "是否過濾相同學校": True,
                            "是否過濾職稱": True,
                            "是否過濾掉自身" : True
                        }
                    )
                    total_committee_person_dict_result = merge_committee_advanced(total_committee_person_dict_result, current_committee_person_dict_result)
                    
                    
                if len(find_temp_df) > 0: break #= 找不到東西，跳掉
        
            #- Input Selector
            # final_committee_person_list = [item for item in committee_person_dict["Remaining Members"][:3]]
            # for index, name in enumerate(final_committee_person_list):
            #     statistical_row[f"選取委員{index+1}"] = name
            
            #- Reason
            statistical_row["篩掉人員"] = total_committee_person_dict_result["Filtered Members"]
            statistical_row["篩選原因"] = total_committee_person_dict_result["Filter Reasons"]
            
            result_dict.append(statistical_row)
            
        pd.DataFrame(result_dict).to_excel(writer, sheet_name=sheet, index=False)
    writer.close()

def filter_committee_remove(is_industry=False):
    #: 1. 預先載入資料 (移出迴圈以提升效能)
    crawler_RDF_folder_path = find_key_path("碩博士論文_RDF")
    crawler_RDF_data = pd.read_excel(crawler_RDF_folder_path)
    
    statistical_analysis_folder_path = find_key_path("統計表分析") 
    statistical_analysis_file = pd.ExcelFile(statistical_analysis_folder_path)
    
    # 根據是否為產學合作選擇路徑
    apply_list_folder_path = find_key_path("產學合作申請名冊") if is_industry else find_key_path("研究計畫申請名冊")
    apply_list_file = pd.ExcelFile(apply_list_folder_path)
    
    committee_person_path = find_key_path("統計清單人才資料_RDF_UNI")
    committee_person_RDF_df = pd.read_excel(committee_person_path)
    
    # 建立寫入器
    writer = pd.ExcelWriter(find_key_path("過濾相近後統計表"), engine='openpyxl')
    
    #@ 2. 開始處理每個分頁
    for sheet in statistical_analysis_file.sheet_names:
        current_sheet_statistical_excel_data = pd.read_excel(statistical_analysis_file, sheet_name=sheet)
        result_dict = []
        
        # 取得所有申請人清單
        all_apply_members = current_sheet_statistical_excel_data[value_of_key("申請主持人欄位名稱")].to_list()
        
        for index, statistical_row in current_sheet_statistical_excel_data.iterrows():
            # = 收集 30 位委員的背景資料
            committee_person_candidates = []
            for i in range(1, 31):  # 調整為 30 位
                col_name = f'推薦委員{i}'
                if col_name not in statistical_row or pd.isna(statistical_row[col_name]):
                    continue
                
                name = statistical_row[col_name]
                
                # 委員過去就職學校
                find_temp_df = committee_person_RDF_df[committee_person_RDF_df["名稱"] == name]
                been_list = list(set(find_temp_df["學校"].tolist()))
                
                # 委員畢業學校
                graduate_list = list(set(find_crawler_person_relative_school(name, crawler_RDF_data)))
                
                committee_person_candidates.append({
                    "委員名稱": name,
                    "委員曾就職學校": been_list,
                    "委員過去畢業學校": graduate_list,
                    "職稱": find_temp_df['職稱'].iloc[0] if not find_temp_df.empty else "未知"
                })
            
            # = 取得申請案相關資訊 (學校、主持人等)
            total_filter_result = {
                'Filtered Members': [],
                'Remaining Members': [],
                'Filter Reasons': {}
            }
            
            # 遍歷申請名冊找到對應計畫
            for app_sheet in apply_list_file.sheet_names:
                current_app_data = pd.read_excel(apply_list_file, sheet_name=app_sheet)
                target_project = current_app_data[current_app_data[value_of_key("計畫名稱")] == statistical_row[value_of_key("計畫名稱")]]
                
                if target_project.empty:
                    continue

                for _, app_row in target_project.iterrows():
                    # 提取共同主持人資訊
                    joint_person_list = app_row.get(value_of_key("申請共同主持人"), [])
                    joint_dept_list = app_row.get(value_of_key("申請共同機構欄位名稱"), [])
                    
                    common_joint_list = extract_text_in_parentheses(joint_person_list) + extract_text_in_parentheses(joint_dept_list)
                    
                    # 建立申請端學校清單
                    apply_school_info = {
                        "申請人名稱": app_row.get(value_of_key("申請主持人欄位名稱"), ''),
                        "計畫申請學校": split_institution(app_row.get(value_of_key("申請機構欄位名稱"), ''))[0],
                        "共同計畫主持的學校": [split_institution(dept)[0] for _, dept in common_joint_list],
                        "計畫主持人過去畢業的學校": find_crawler_person_relative_school(app_row.get(value_of_key("申請主持人欄位名稱")), crawler_RDF_data),
                    }

                    # 執行進階過濾
                    filter_pairs = [("計畫申請學校", "委員曾就職學校"), ("共同計畫主持的學校", "委員曾就職學校")]
                    TITLE_RESTRICTIONS = {"助理教授": ["教授", "研究員"], "助研究員": ["教授", "研究員"]}
                    
                    current_res = filter_committee_advanced(
                        apply_school_info, 
                        committee_person_candidates, 
                        filter_pairs, 
                        all_apply_members, 
                        TITLE_RESTRICTIONS,
                        whether_to_execute_the_option={
                            "是否過濾申請人": True if is_industry else False,
                            "是否過濾相同學校": True,
                            "是否過濾職稱": True,
                            "是否過濾掉自身": True
                        }
                    )
                    total_filter_result = merge_committee_advanced(total_filter_result, current_res)
                break 

            # = 3. 挑選前 10 位合格委員並寫回 Row
            qualified_members = total_filter_result["Remaining Members"]
            for i in range(1, 11):
                # 如果合格人數不足 10 位，後面會填入空值或提示
                statistical_row[f"選取委員{i}"] = qualified_members[i-1] if i <= len(qualified_members) else ""
            
            # 紀錄篩掉的原因供查核
            statistical_row["篩掉人員"] = str(total_filter_result["Filtered Members"])
            statistical_row["篩選原因"] = str(total_filter_result["Filter Reasons"])
            
            result_dict.append(statistical_row)
            
        # 儲存該分頁結果
        pd.DataFrame(result_dict).to_excel(writer, sheet_name=sheet, index=False)
    
    writer.close()
    print("過濾完成，結果已儲存。")

def load_data(file_path):
    """
        讀取 Excel 檔案並回傳 workbook 和 worksheet.
    """
    workbook = openpyxl.load_workbook(file_path)
    worksheet = workbook.active
    return workbook, worksheet

def generate_letters_excel(start_index, gap, count):
    def index_to_excel_column(index):
        column = ""
        while index > 0:
            index -= 1  # 將索引調整為 0 基礎
            column = chr(index % 26 + 65) + column
            index //= 26
        return column

    letters = []
    for i in range(count):
        letter_index = start_index + i * gap
        letters.append(index_to_excel_column(letter_index))
    return letters

def add_comments(target_ws, data_ws, columns_to_comment):
    """
        在目標工作表上添加註釋
    """
    
    # columns_to_comment = ['D', 'F', 'H', 'J', 'L', 'N', 'P', 'R', 'T', 'V']  # 需要添加註釋的欄位

    # 建立名稱與詳細資訊的對應字典
    # ['名稱', '年份', '機關名稱', '職稱', '來源', '學校', '系所']
    name_to_details = {
    data_ws.cell(row=i, column=1).value: \
        f"名稱: {data_ws.cell(row=i, column=1).value}\n"
        f"年份: {data_ws.cell(row=i, column=2).value}\n"
        f"機關: {data_ws.cell(row=i, column=3).value}\n"
        f"職稱: {data_ws.cell(row=i, column=4).value}\n"
        # f"來源: {data_ws.cell(row=i, column=5).value}"
        for i in range(2, data_ws.max_row + 1)
    }
    
    headers = [cell.value for cell in data_ws[1]]  # 取得第一列的標題名稱
    source_index = headers.index('來源') + 1  # Excel 的索引從 1 開始
    print("來源欄位的索引位置:", source_index)
    print("欄位名稱:", headers)

    # 在每個指定欄位添加註釋
    for col in columns_to_comment:
        for cell in target_ws[col]:
            if cell.value in name_to_details:
                comment_text = name_to_details[cell.value]
                comment = openpyxl.comments.Comment(comment_text, "Python Script")
                comment.width = 350  # 設置寬度
                comment.height = 100  # 設置高度
                cell.comment = comment
                
def excel_process_VBA():
    from openpyxl.utils.cell import column_index_from_string
    from openpyxl.utils import get_column_letter

    gap = 2
    range_count = 10
    
    #: load the excel data.
    talent_workbook, talent_sheet = load_data(find_key_path("統計清單人才資料_RDF"))
    committee_workbook, committee_sheet = load_data(find_key_path("過濾相近後統計表"))
    out_count = committee_sheet.max_column - ( range_count * gap ) - 6
    
    start_index = out_count + 1 
    letter_index = generate_letters_excel(start_index, gap, range_count) # start_index = Excel 26 進制的索引
    # timeline = [start_index + i * gap for i in range(range_count)]  
    
    add_comments(committee_sheet, talent_sheet, columns_to_comment=letter_index)
    pink_fill = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid') #= 填色
    
    header_value = committee_sheet.cell(row=1, column=start_index).value
    print("[起頭] 第 {} 欄之標題為：{}".format(start_index, header_value))

    # 檢查 AB, 滿足條件改色
    for row in committee_sheet.iter_rows(min_row=2, max_col=committee_sheet.max_column):
        filter_list = ast.literal_eval(row[-2].value)
        
        #- 若有重複的的部分進行圖色（篩選委員）
        for col_letter in letter_index:
            col_index = column_index_from_string(col_letter) - 1
            if row[col_index].value in filter_list:
                row[col_index].fill = pink_fill

    # 保存
    committee_workbook.save(find_key_path("FINAL_COMMITTEE"))
    print(f"[結束] 已經保存至: {find_key_path('FINAL_COMMITTEE')}")

from datetime import datetime
def update_peronsal_info_database(is_industry=False):
    '''
        更新暫存最新人才資料庫
    '''
    
    # 讀取最新 Excel 檔案
    if is_industry:
        apply_list_folder_path = find_key_path("產學合作申請名冊")
        file_name = value_of_key("產學合作申請名冊")
    else:
        apply_list_folder_path = find_key_path("研究計畫申請名冊") 
        file_name = value_of_key("研究計畫申請名冊")  # 修正錯誤的 `file_name(file_name)`

    apply_list_file = pd.ExcelFile(apply_list_folder_path)
    
    # 讀取暫存人才資料庫，若不存在則建立新的 DataFrame
    personal_info_database_path = find_key_path("暫存最新人才資料庫")

    if os.path.exists(personal_info_database_path):
        personal_info_database = pd.read_excel(personal_info_database_path)
    else:
        personal_info_database = pd.DataFrame(columns=['名稱', '年份', '機關名稱', '職稱', '來源'])

    new_data = []

    for sheet in apply_list_file.sheet_names: 
        current_sheet_apply_excel_data = pd.read_excel(apply_list_file, sheet_name=sheet)
        for _, row in current_sheet_apply_excel_data.iterrows():
            new_row = {
                '名稱': row.get(value_of_key('申請主持人欄位名稱')), 
                '年份': datetime.now().year - 1911,  # 直接填入現在的年份
                '機關名稱': row.get(value_of_key('申請機構欄位名稱'), ''),
                '職稱': row.get(value_of_key('職稱'), ''),
                '來源': file_name,
            }
            new_data.append(new_row)

    # 將新數據轉換為 DataFrame
    new_data_df = pd.DataFrame(new_data)

    # 合併數據，並以 "名稱" 為主鍵，保留最新年份的資料
    updated_database = pd.concat([personal_info_database, new_data_df]).sort_values(by=['名稱', '年份'], ascending=[True, False])
    updated_database = updated_database.drop_duplicates(subset=['名稱'], keep='first')  # 只保留最新的年份

    # 儲存回 Excel
    updated_database.to_excel(personal_info_database_path, index=False)
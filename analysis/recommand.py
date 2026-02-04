import pandas as pd
import os

def load_data(file_path, pages):
    page_data = {}  # 初始化字典
    target_cols = ['project', 'manager', 'similarity_score', 'recommended_manager', 'school']
    for page in pages:
        try:
            df = pd.read_excel(file_path, sheet_name=page, dtype=str).fillna("")
            page_data[page] = df[target_cols]
            
        except KeyError as e:
            print(f"頁面 {page} 缺少欄位: {e}")
            continue
        except Exception as e:
            print(f"跳過頁面 {page}: {e}")
            continue
        
    return page_data

def save_results_to_excel(results, output_filename="data/output/recommendation_results.xlsx"):
    excel_rows = []

    for project, managers in results.items():
        # 1. 每一列 (Row) 的基本資料
        row_data = {
            "Project": project,
            'Manager': managers['manager'],
            "School": managers['school']
        }
        
        # 2. 動態產生推薦欄位
        # managers 是一個 list: [('王小明', 0.8), ('李大同', 0.6)...]
        for i, (manager_name, score) in enumerate(managers['recommand_managers']):
            if i == 10:
                break
            rank = i + 1  # 排名從 1 開始
            
            # 建立欄位名稱: Recommend 1, Score 1, Recommend 2, Score 2...
            row_data[f"Recommend {rank}"] = manager_name
            row_data[f"Score {rank}"] = score
            
        excel_rows.append(row_data)

    # 3. 轉成 DataFrame
    df = pd.DataFrame(excel_rows)

    # 4. 存成 Excel
    df.to_excel(output_filename, index=False)
    print(f"檔案已儲存至: {output_filename}")

def group_by_project_dict(datas, pages):
    results = {}
    
    for page in pages:
        df = datas[page]
        rows = df.to_dict('records') 
        for row in rows:
            project_name = row.get('project')
            
            if not project_name:
                continue
            if project_name not in results:
                results[project_name] = {}
            if page not in results[project_name]:
                results[project_name][page] = [] 
            results[project_name][page].append(row)
            
    return results

def calculate_score(data):
    # 定義權重表
    weights = {
        'title': 1,
        'keyword': 1,
        'application_directions': 3,
        'problems_to_solve': 3,
        'goals_to_achieve': 1,
        'methods_to_solve': 1
    }
    total_weight_sum = sum(weights.values())
    final_results = {}
    for project_name, pages_data in data.items():
        manager_scores = {} # 用來暫存該專案下每位教授的分數
        
        for page, rows in pages_data.items():
            weight = weights.get(page, 0)
            
            if weight == 0:
                continue

            for row in rows:
                manager = row.get('recommended_manager')
                score_str = row.get('similarity_score', 0)
                
                if not manager:
                    continue

                try:
                    score = float(score_str)
                except ValueError:
                    score = 0
                
                weighted_score = score * weight
                
                if manager not in manager_scores:
                    manager_scores[manager] = 0
                
                manager_scores[manager] += weighted_score
        
        for manager in manager_scores:
            manager_scores[manager] = manager_scores[manager] / total_weight_sum
        sorted_managers = sorted(manager_scores.items(), key=lambda x: x[1], reverse=True)
        
        final_results[project_name] = {
            'manager': row.get('manager'),
            'school': row.get('school'),
            'recommand_managers':sorted_managers,
            
        }

    return final_results

def filter_committee_advanced(
        schools_info, 
        committee_members, 
        filter_pairs, 
        apply_member_list=None, 
        TITLE_RESTRICTIONS={},
        whether_to_execute_the_option={
            "是否過濾申請人": True,
            "是否過濾相同學校": True,
            "是否過濾職稱": True
        }
    ):
    """
    進階過濾委員名單，根據具體的配對關係進行過濾，並提供過濾的具體原因。

    :param schools_info: 包含學校相關資訊的字典
    :param committee_members: 包含委員相關資訊的列表
    :param filter_pairs: 列表，包含過濾配對條件，例如 [("申請學校", "就職學校")]
    :param apply_member_list: 申請人名單，若提供則優先過濾
    :param TITLE_RESTRICTIONS: 職稱過濾規則的字典
    
    :return: 一個字典，包含過濾前後的委員名單和未過濾的委員名單，以及過濾原因
    """
    
    filtered_members = set()
    filter_reasons = {}
    # 1. (若篩選委員有申請人，則刪除) => 如果有提供 apply_member_list，先將該列表中的委員優先篩選掉
    if apply_member_list and whether_to_execute_the_option["是否過濾申請人"]:
        for member in committee_members:
            if member['委員名稱'] in apply_member_list:
                filtered_members.add(member['委員名稱'])
                filter_reasons[member['委員名稱']] = f"委員名稱 {member['委員名稱']} 出現在申請人之中"

    #  2. 根據配對條件進行過濾（例如 (計畫申請學校, 委員曾就職學校) 等）
    if whether_to_execute_the_option["是否過濾相同學校"]:
        for school_type, member_field in filter_pairs:
            if school_type in schools_info and schools_info[school_type]:
                school_list = schools_info[school_type] if isinstance(schools_info[school_type], list) else [schools_info[school_type]]
                print(schools_info["申請人名稱"])
                print(school_list)
                for member in committee_members:
                    print(member['委員名稱'])
                    print(member[member_field])
                    matching_schools = [school for school in member[member_field] if school in school_list and school]
                    if matching_schools:
                        filtered_members.add(member['委員名稱'])
                        filter_reasons[member['委員名稱']] = f"{school_type} 與 {member_field} ({', '.join(matching_schools)}) 重疊"
                print('----')
    # 3. 根據職稱進行過濾
    if whether_to_execute_the_option["是否過濾職稱"]:
        for member in committee_members:
            if member['委員名稱'] in filtered_members:
                continue  # 若已被篩選，不再處理
            
            applicant_title = str(schools_info.get("申請人職稱", "")).strip()
            member_title = str(member.get("職稱", "")).strip()

            # 若該職稱有過濾規則，且申請人職稱在排除名單中
            if member_title in TITLE_RESTRICTIONS and applicant_title in TITLE_RESTRICTIONS[member_title]:
                filtered_members.add(member['委員名稱'])
                filter_reasons[member['委員名稱']] = f"{member_title} 不能審查 {applicant_title}"

    # 創建過濾後的委員名單
    remaining_members = [member['委員名稱'] for member in committee_members if member['委員名稱'] not in filtered_members]

    # 返回結果
    return {
        'Filtered Members': list(filtered_members),
        'Remaining Members': remaining_members,
        'Filter Reasons': filter_reasons
    }
    
def get_clean_school_name(school):
    # 定義關鍵字列表 (注意：順序很重要，程式會以先找到的為主)
    keywords = ['大學', '院', '博物館', '學校', '法人']
    
    # 預設結果等於原始名稱 (萬一都沒對到關鍵字，就保留原樣)
    final_name = school

    for kw in keywords:
        if kw in school:
            final_name = school.split(kw)[0] + kw
            break 
            
    return final_name

def main():
    input_file = 'data/output/result_score.xlsx'
    
    if not os.path.exists(input_file):
        print(f"找不到檔案: {input_file}")
        return

    pages = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    
    # 1. 讀取資料
    org_data = load_data(input_file, pages) 
    # print(org_data)
    project_data = group_by_project_dict(org_data,pages)
    # print(len(integrate_data['前瞻多模式多色階半色調技術之探究']['title']))
    result = calculate_score(project_data)
    # print(result['前瞻多模式多色階半色調技術之探究'])
    save_results_to_excel(result)
    blacklist_csv_path = 'data/retiree_blacklist.csv'
    blacklist_df = pd.read_csv(blacklist_csv_path, encoding='utf-8')
    blacklist = blacklist_df['姓名'].tolist()
    print(f"黑名單或是已退休委員:", blacklist)

    person_path = 'data/RDF_database/統計清單人才資料_RDF_UNI.xlsx'
    newest_person_df = pd.read_excel(person_path)
    print(f"統計清單人才資料:", newest_person_df)
    committee_person_dict = newest_person_df.set_index('名稱').to_dict('index')
    filter_pairs = [("計畫申請學校", "委員曾就職學校")]
    TITLE_RESTRICTIONS = {
                        "助理教授": ["教授", "研究員"], # 助理教授不能審教授或研究員
                        "助研究員": ["教授", "研究員"] # 助研究員不能審教授或研究員
                    }
    all_apply_members = []

    for project in result:
        apply_school = {
                        "申請人名稱": result[project]['manager'],
                        "計畫申請學校": get_clean_school_name(result[project]['school'])
                    }
        current_committee_person_dict_result = filter_committee_advanced(
            apply_school, 
            committee_person_dict, 
            filter_pairs, 
            all_apply_members, 
            TITLE_RESTRICTIONS,
            whether_to_execute_the_option= {
                "是否過濾申請人": False,
                "是否過濾相同學校": True,
                "是否過濾職稱": True,
                "是否過濾掉自身" : True
            }
        )
        print(current_committee_person_dict_result)
    

if __name__ == "__main__":
    main()
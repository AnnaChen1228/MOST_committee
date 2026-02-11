import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
# 假設 save_commitee_data 是你自己的模組，確保它在同目錄下
from save_commitee_data import filiter_school 

# --- 1. 讀取申請資料 (維持你的邏輯) ---
def load_apply_data(file_path, pages):
    apply_dicts = {} 
    print(f"讀取申請資料: {file_path}")
    
    for page in pages:
        try:
            apply_project_df = pd.read_excel(file_path, sheet_name=page, dtype=str).fillna("")
        except Exception as e:
            print(f"跳過頁面 {page}: {e}")
            continue

        for index, row in apply_project_df.iterrows():
            title = str(row['計畫中文名稱']).strip()
            if not title: continue
            
            # 主持人學校清洗
            school, department = filiter_school(row['機關名稱'])
            
            rawCoManagers = row['共同主持人']
            coSchools = []
            coDepartments = []
            coManagers = []

            if pd.notna(rawCoManagers) and str(rawCoManagers).strip():
                orgCoMangers = str(rawCoManagers).split(';')
                
                for comanager in orgCoMangers:
                    comanager = comanager.strip()
                    if not comanager: continue 
                    
                    # Regex 拆分
                    match = re.match(r'^(.*?)[(（](.*)[)）]', comanager)
                    
                    if match:
                        name = match.group(1).strip()
                        raw_school_str = match.group(2).strip()
                        # 這裡記得也要清洗學校名稱，以便跟委員資料庫比對
                        coschool, codepartment = filiter_school(raw_school_str)
                    else:
                        name = comanager
                        coschool = ""
                        codepartment = ''
                    
                    coManagers.append(name)
                    coSchools.append(coschool)
                    coDepartments.append(codepartment)

            apply_dicts[title] = {
                'manager': str(row['計畫主持人']),
                'title': title,
                'managerSchool': school,
                'managerDepartment': department,
                'coManger': coManagers, # List of names
                'coSchool': coSchools   # List of schools
            }
    return apply_dicts

# --- 2. 讀取評分結果 (維持不變) ---
def load_score_data(file_path, pages):
    page_data = {}
    target_cols = ['project', 'manager', 'similarity_score', 'recommended_manager']
    
    for page in pages:
        try:
            df = pd.read_excel(file_path, sheet_name=page, dtype=str).fillna("")
            available_cols = [c for c in target_cols if c in df.columns]
            page_data[page] = df[available_cols]
        except Exception as e:
            print(f"跳過頁面 {page}: {e}")
            continue
    return page_data

def group_by_project_dict(datas, pages):
    results = {}
    for page in pages:
        if page not in datas: continue
        df = datas[page]
        rows = df.to_dict('records') 
        for row in rows:
            project_name = row.get('project')
            if not project_name: continue
            if project_name not in results:
                results[project_name] = {}
            if page not in results[project_name]:
                results[project_name][page] = [] 
            results[project_name][page].append(row)
    return results

# --- 3. 計算加權分數 (維持不變) ---
def calculate_score(data):
    weights = {
        'title': 1, 'keyword': 1,
        'application_directions': 3, 'problems_to_solve': 3,
        'goals_to_achieve': 1, 'methods_to_solve': 1
    }
    total_weight_sum = sum(weights.values())
    final_results = {}

    for project_name, pages_data in data.items():
        manager_scores = {}
        applicant_manager = ""
        # 這裡的 applicant_school 稍後會從 apply_data 覆蓋更準確的值
        
        for page, rows in pages_data.items():
            weight = weights.get(page, 0)
            if weight == 0: continue

            for row in rows:
                if not applicant_manager:
                    applicant_manager = row.get('manager', '')

                rec_manager = row.get('recommended_manager')
                score_str = row.get('similarity_score', 0)
                
                if not rec_manager: continue

                try:
                    score = float(score_str)
                except ValueError:
                    score = 0
                
                weighted_score = score * weight
                
                if rec_manager not in manager_scores:
                    manager_scores[rec_manager] = 0
                
                manager_scores[rec_manager] += weighted_score
        
        candidates_list = []
        for manager, raw_score in manager_scores.items():
            final_score = raw_score / total_weight_sum
            candidates_list.append({
                'name': manager,
                'score': final_score
            })
            
        candidates_list.sort(key=lambda x: x['score'], reverse=True)
        
        final_results[project_name] = {
            'manager': applicant_manager,
            'school': "", # 預設空，稍後填入
            'candidates': candidates_list 
        }

    return final_results

# --- ★★★ 4. 存檔並上色 (新增共同主持人邏輯) ★★★ ---
def save_results_with_highlight(results, output_filename, reviewer_school_map, blacklist):
    """
    顏色定義：
    - 黑名單: 灰色 (D9D9D9)
    - 本人: 淺綠色 (E2EFDA)
    - 共同主持人 / 共同主持人學校: 淺藍色 (BDD7EE)  <-- 新增
    - 主持人同校: 淺紅色 (FCE4D6)
    """
    excel_rows = []
    sorted_projects = sorted(results.keys())

    # --- 第一步：寫入 Excel ---
    for project in sorted_projects:
        info = results[project]
        row_data = {
            "Project": project,
            'Manager': info['manager'],
            "School": info['school'] # 主持人學校
        }
        
        candidates = info.get('candidates', [])
        
        for i, candidate in enumerate(candidates):
            if i >= 10: break 
            rank = i + 1
            cand_name = candidate['name']
            cand_school = reviewer_school_map.get(cand_name, "")
            
            row_data[f"Recommend {rank}"] = cand_name
            row_data[f"Score {rank}"] = candidate['score']
            row_data[f"Rec_School {rank}"] = cand_school
            
        excel_rows.append(row_data)

    df = pd.DataFrame(excel_rows)
    df.to_excel(output_filename, index=False)
    
    # --- 第二步：上色 ---
    print(f"正在為 {output_filename} 進行上色標記...")
    wb = load_workbook(output_filename)
    ws = wb.active
    
    # 定義顏色
    fill_blacklist = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") # 灰
    fill_person = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")    # 綠 (本人)
    fill_school = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")    # 紅 (同校)
    fill_copi = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")      # 藍 (共同主持人相關)

    for i, project in enumerate(sorted_projects):
        row_idx = i + 2 
        info = results[project]
        
        applicant_name = info['manager']
        applicant_school = info['school']
        
        # 取得共同主持人資訊 (如果沒有則為空list)
        co_managers = info.get('co_managers', [])
        co_schools = info.get('co_schools', [])
        
        candidates = info.get('candidates', [])
        
        for k, cand in enumerate(candidates):
            if k >= 10: break
            
            cand_name = cand['name']
            cand_school = reviewer_school_map.get(cand_name, "")
            
            target_fill = None
            
            # --- 判斷邏輯 (順序可依需求調整) ---
            
            # 1. 黑名單
            if cand_name in blacklist:
                target_fill = fill_blacklist
            
            # 2. 本人
            elif cand_name == applicant_name:
                target_fill = fill_person
            
            # 3. 共同主持人相關 (本人 OR 學校) -> 淺藍色
            elif (cand_name in co_managers) or (cand_school and cand_school in co_schools):
                target_fill = fill_copi
                
            # 4. 主持人同校 -> 淺紅色
            elif applicant_school and cand_school and applicant_school == cand_school:
                target_fill = fill_school
            
            # 執行上色
            if target_fill:
                base_col = 4 + (k * 3)
                ws.cell(row=row_idx, column=base_col).fill = target_fill
                ws.cell(row=row_idx, column=base_col+1).fill = target_fill
                ws.cell(row=row_idx, column=base_col+2).fill = target_fill

    wb.save(output_filename)
    print(f"檔案已儲存並上色: {output_filename}")

# --- 5. 存檔乾淨名單 ---
def save_results_clean(results, output_filename, key_name='final_candidates'):
    excel_rows = []
    sorted_projects = sorted(results.keys())

    for project in sorted_projects:
        info = results[project]
        row_data = {
            "Project": project,
            'Manager': info['manager'],
            "School": info['school']
        }
        
        valid_candidates = info.get(key_name, [])
        
        for i, candidate in enumerate(valid_candidates):
            if i >= 10: break 
            rank = i + 1
            row_data[f"Recommend {rank}"] = candidate['name']
            row_data[f"Score {rank}"] = candidate['score']
            row_data[f"Rec_School {rank}"] = candidate.get('school', '') 
            
        excel_rows.append(row_data)

    df = pd.DataFrame(excel_rows)
    df.to_excel(output_filename, index=False)
    print(f"檔案已儲存至: {output_filename}")

def main():
    apply_path = 'data/research_proj/115計算機學門審查/apply_project_with_abstract.xlsx'
    apply_years = ['115']
    input_score_file = 'data/output/result_score.xlsx'
    committee_uni_file = 'data/RDF_database/commitee_uni.xlsx'
    blacklist_csv_path = 'data/retiree_blacklist.csv'
    
    # 1. 讀取申請資料
    if not os.path.exists(apply_path):
        print(f"找不到申請檔案: {apply_path}")
        return
    apply_data = load_apply_data(apply_path, apply_years)

    # 2. 讀取評分 & 計算
    if not os.path.exists(input_score_file):
        print(f"找不到評分檔案: {input_score_file}")
        return
    print("正在讀取評分資料...")
    pages = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    org_data = load_score_data(input_score_file, pages) 
    project_data = group_by_project_dict(org_data, pages)
    result = calculate_score(project_data)

    # ★★★ 關鍵步驟：將申請資料(學校、共同主持人) 合併到 result 中 ★★★
    print("正在合併申請人與共同主持人資訊...")
    for project_name, info in result.items():
        if project_name in apply_data:
            app_info = apply_data[project_name]
            info['school'] = app_info['managerSchool']  # 更新正確的主持人學校
            info['co_managers'] = app_info['coManger']  # 注入共同主持人名單
            info['co_schools'] = app_info['coSchool']   # 注入共同主持人學校
        else:
            # 如果找不到對應的申請資料，給空值避免報錯
            info['co_managers'] = []
            info['co_schools'] = []

    # 3. 載入外部資料 (委員學校、黑名單)
    print(f"正在讀取委員學校資料: {committee_uni_file}")
    reviewer_school_map = {}
    if os.path.exists(committee_uni_file):
        try:
            commitee_df = pd.read_excel(committee_uni_file, dtype=str).fillna("")
            for _, row in commitee_df.iterrows():
                name = str(row['名字']).strip()
                # 建議這裡也可以用 filiter_school 清洗一下，確保比對準確
                # school, _ = filiter_school(str(row['學校']))
                school = str(row['學校']).strip() 
                if name:
                    reviewer_school_map[name] = school
            print(f"已載入委員學校資料: {len(reviewer_school_map)} 筆")
        except Exception as e:
            print(f"讀取委員資料失敗: {e}")

    blacklist = []
    if os.path.exists(blacklist_csv_path):
        try:
            blacklist_df = pd.read_csv(blacklist_csv_path, encoding='utf-8')
            blacklist = blacklist_df['姓名'].astype(str).tolist()
            print(f"已載入黑名單: {len(blacklist)} 人")
        except Exception as e:
            print(f"讀取黑名單失敗: {e}")

    # 4. 儲存原始名單 (含過濾標記 - 包含共同主持人)
    print("正在儲存原始名單 (含過濾標記)...")
    save_results_with_highlight(result, 'data/output/recommendation_results_org_colored.xlsx', reviewer_school_map, blacklist)

    # 5. 執行過濾 (產生乾淨名單)
    print("正在執行過濾...")
    for project_name, info in result.items():
        applicant_name = info['manager']
        applicant_school = info['school']
        co_managers = info.get('co_managers', [])
        co_schools = info.get('co_schools', [])
        
        filtered_candidates = []
        
        for cand in info['candidates']:
            cand_name = cand['name']
            
            # (A) 黑名單
            if cand_name in blacklist: continue
            
            # (B) 本人
            if cand_name == applicant_name: continue
            
            # (C) 共同主持人本人
            if cand_name in co_managers: continue
            
            # 取得委員學校
            cand_school_raw = reviewer_school_map.get(cand_name, "")
            cand['school'] = cand_school_raw 
            
            if cand_school_raw:
                # (D) 主持人同校
                if applicant_school and applicant_school == cand_school_raw:
                    continue
                
                # (E) 共同主持人同校
                if cand_school_raw in co_schools:
                    continue
            
            filtered_candidates.append(cand)
        
        info['final_candidates'] = filtered_candidates

    # 6. 存一份過濾後的乾淨檔案
    print("正在儲存過濾後名單...")
    save_results_clean(result, 'data/output/recommendation_results_filter.xlsx', key_name='final_candidates')

if __name__ == "__main__":
    main()

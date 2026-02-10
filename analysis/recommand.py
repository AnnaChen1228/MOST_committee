import pandas as pd
import os

# --- 1. 讀取評分結果 ---
def load_data(file_path, pages):
    page_data = {}
    target_cols = ['project', 'manager', 'similarity_score', 'recommended_manager', 'school']
    
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

# --- 2. 計算加權分數 ---
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
        applicant_school = ""
        
        for page, rows in pages_data.items():
            weight = weights.get(page, 0)
            if weight == 0: continue

            for row in rows:
                if not applicant_manager:
                    applicant_manager = row.get('manager', '')
                    applicant_school = row.get('school', '')

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
            
        # 依照分數由大到小排序
        candidates_list.sort(key=lambda x: x['score'], reverse=True)
        
        final_results[project_name] = {
            'manager': applicant_manager,
            'school': applicant_school,
            'candidates': candidates_list # 這是未過濾的原始名單
        }

    return final_results

# --- 3. 存檔函式 (修改：增加 key_name 參數) ---
def save_results_to_excel(results, output_filename, key_name='final_candidates'):
    """
    key_name: 指定要從 results 裡讀取哪個名單
              'candidates' -> 原始名單
              'final_candidates' -> 過濾後名單
    """
    excel_rows = []

    for project, info in results.items():
        row_data = {
            "Project": project,
            'Manager': info['manager'],
            "School": info['school']
        }
        
        # 根據傳入的 key_name 抓取對應的名單
        valid_candidates = info.get(key_name, [])
        
        for i, candidate in enumerate(valid_candidates):
            if i >= 10: break # 只取前 10 名
            rank = i + 1
            row_data[f"Recommend {rank}"] = candidate['name']
            row_data[f"Score {rank}"] = candidate['score']
            row_data[f"Rec_School {rank}"] = candidate.get('school', '') 
            
        excel_rows.append(row_data)

    df = pd.DataFrame(excel_rows)
    df.to_excel(output_filename, index=False)
    print(f"檔案已儲存至: {output_filename}")

def main():
    input_score_file = 'data/output/result_score.xlsx'
    committee_uni_file = 'data/RDF_database/commitee_uni.xlsx'
    blacklist_csv_path = 'data/retiree_blacklist.csv'
    
    if not os.path.exists(input_score_file):
        print(f"找不到評分檔案: {input_score_file}")
        return

    # 1. 計算分數
    print("正在讀取評分資料...")
    pages = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    org_data = load_data(input_score_file, pages) 
    project_data = group_by_project_dict(org_data, pages)
    result = calculate_score(project_data)

    # 2. 載入外部資料
    print(f"正在讀取委員學校資料: {committee_uni_file}")
    reviewer_school_map = {}
    if os.path.exists(committee_uni_file):
        try:
            commitee_df = pd.read_excel(committee_uni_file, dtype=str).fillna("")
            for _, row in commitee_df.iterrows():
                name = str(row['名字']).strip()
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

    # ★★★ 3. 先存一份原始檔 (未過濾) ★★★
    # 指定 key_name='candidates'，因為這時候還沒有 final_candidates
    print("正在儲存原始名單...")
    save_results_to_excel(result, 'data/output/recommendation_results_org.xlsx', key_name='candidates')

    # 4. 執行過濾
    print("正在執行過濾 (黑名單、本人、同校迴避)...")
    
    for project_name, info in result.items():
        applicant_name = info['manager']
        applicant_school = info['school']
        
        filtered_candidates = []
        
        # 遍歷原始名單
        for cand in info['candidates']:
            cand_name = cand['name']
            
            # (A) 黑名單
            if cand_name in blacklist:
                continue
                
            # (B) 本人
            if cand_name == applicant_name:
                continue
            
            # (C) 同校
            cand_school_raw = reviewer_school_map.get(cand_name, "")
            cand['school'] = cand_school_raw 
            
            if applicant_school and cand['school']:
                # 這裡建議還是加上簡單的清洗，避免 "國立台灣大學" != "台灣大學" 的問題
                # 如果確定資料很乾淨可以不用
                if applicant_school == cand['school']:
                    continue
            
            filtered_candidates.append(cand)
        
        # 將過濾結果存入 final_candidates
        info['final_candidates'] = filtered_candidates

    # ★★★ 5. 存一份過濾後的檔案 ★★★
    # 預設 key_name='final_candidates'
    print("正在儲存過濾後名單...")
    save_results_to_excel(result, 'data/output/recommendation_results_filter.xlsx', key_name='final_candidates')

if __name__ == "__main__":
    main()

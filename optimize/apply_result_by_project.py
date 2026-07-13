import pandas as pd
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma
import os
import argparse
import chromadb
import tqdm

# 各計畫類型對應的向量資料庫路徑（研究計畫 / 產學合作）
DB_PATHS = {
    'research': {
        'basic': 'database/vectorstore_basic_research_by_project',
        'abstract': 'database/vectorstore_abstract_research_by_project',
    },
    'industry': {
        'basic': 'database/vectorstore_basic_industry_by_project',
        'abstract': 'database/vectorstore_abstract_industry_by_project',
    },
}

# 一律比對全部欄位（名稱、關鍵字、摘要拆解出的四個欄位）
ALL_FIELDS = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']

# --- 1. 設定模型 ---
model_name = 'BAAI/bge-large-zh-v1.5'
model_kwargs = {'device': 'cpu'}
encode_kwargs = {'normalize_embeddings': True}

print("正在載入模型...")
model = HuggingFaceEmbeddings(
    model_name=model_name,
    model_kwargs=model_kwargs,
    encode_kwargs=encode_kwargs,
)

# --- 2. 讀取申請計畫資料 ---
def load_data(file_path, pages):
    apply_dicts = {} 
    print(f"讀取 Excel: {file_path}")
    
    for page in pages:
        try:
            apply_project_df = pd.read_excel(file_path, sheet_name=page, dtype=str).fillna("")
        except Exception as e:
            print(f"跳過頁面 {page}: {e}")
            continue

        for index, row in apply_project_df.iterrows():
            title = str(row.get('計畫名稱') or row.get('計畫中文名稱') or "").strip()
            if not title: continue

            # ★ 修正：使用 .get() 兼容不同 Excel 可能的欄位名稱
            try:
                keywords_str = row.get('keywords') or row.get('中文關鍵字') or ""
                if isinstance(keywords_str, list):
                    keywords = keywords_str
                elif keywords_str:
                    keywords_str = keywords_str.replace('，', ',').replace('；', ',').replace(';', ',').replace('、', ',').replace('。', ',')
                    keywords_str = keywords_str.replace('\n', ',').replace('\r', ',')
                    keywords = keywords_str.split(',')
                    keywords = [k.strip() for k in keywords if k.strip()]
                else:
                    keywords = []
            except:
                keywords = []
                
            apply_dicts[title] = {
                'manager': str(row.get('主持人', '')),
                'title': title,
                'keywords': keywords,
                'abstract': row.get('中文摘要', ''),
                'application_directions': row.get('application_directions', ''),
                'problems_to_solve': row.get('problems_to_solve', ''),
                'goals_to_achieve': row.get('goals_to_achieve', ''),
                'methods_to_solve': row.get('methods_to_solve', '')
            }
    return apply_dicts

# --- 3. 儲存結果 ---
def save_data(file_path, data_list, sheet_name):
    if not data_list:
        return

    df = pd.DataFrame(data_list)

    if os.path.exists(file_path):
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

# --- 4. 搜尋函式 ---
def search(vectordb, query_text, RECOMMAND_AMOUNT=50): # 稍微抓多一點備用
    try:
        documents = vectordb.similarity_search_with_relevance_scores(
            query_text,
            k=RECOMMAND_AMOUNT,
            score_threshold=0.1
        )
        return documents
    except Exception as e:
        print(f"搜尋錯誤: {e}")
        return []

def main():
    parser = argparse.ArgumentParser(description='比對當次申請案件與向量資料庫中的過往通過案件，計算相似度分數（以專案為單位）')
    parser.add_argument('--input', default='data/research_proj/115計算機學門審查/新興大專生計畫/115 新興 大專生計畫-智慧計算.xlsx',
                        help='當次申請案件的 Excel 檔')
    parser.add_argument('--sheets', nargs='+', default=['工作表1'],
                        help='申請案件要讀取的工作表(sheet)名稱')
    parser.add_argument('--type', choices=['research', 'industry'], default='research',
                        help='計畫類型，決定向量資料庫預設路徑（research=研究計畫, industry=產學合作）')
    parser.add_argument('--basic-db', default=None,
                        help='Title / Keywords 向量資料庫路徑（未指定時依 --type 自動決定，需與 store 步驟一致）')
    parser.add_argument('--abstract-db', default=None,
                        help='摘要四欄位向量資料庫路徑（未指定時依 --type 自動決定，需與 store 步驟一致）')
    parser.add_argument('--output', default='data/research_proj/115計算機學門審查/新興大專生計畫/result_score(115 新興 大專生計畫-智慧計算_by_project_research).xlsx',
                        help='相似度分數輸出的 Excel 檔')
    args = parser.parse_args()

    store_file_path = args.input
    years = args.sheets

    # 1. 載入申請資料
    apply_dicts = load_data(store_file_path, years)
    print(len(apply_dicts))
    # 2. 設定資料庫路徑與分類（未指定則依 --type 帶出預設）
    path_basic = args.basic_db or DB_PATHS[args.type]['basic']
    path_abstract = args.abstract_db or DB_PATHS[args.type]['abstract']

    collections_map = {
        'title': path_basic,
        'keywords': path_basic,
        'application_directions': path_abstract,
        'problems_to_solve': path_abstract,
        'goals_to_achieve': path_abstract,
        'methods_to_solve': path_abstract
    }

    # 一律比對全部欄位
    target_collections = ALL_FIELDS
    print("正在連線至向量資料庫...")
    client_basic = chromadb.PersistentClient(path=path_basic)
    client_abstract = chromadb.PersistentClient(path=path_abstract)

    output_file = args.output
    if os.path.exists(output_file):
        try:
            os.remove(output_file)
            print(f"🗑️ 已刪除舊的輸出檔案: {output_file}")
        except Exception as e:
            print(f"⚠️ 無法刪除舊檔: {e}")
            return 

    # 3. 開始搜尋
    for col_name in target_collections:
        db_path = collections_map.get(col_name)
        current_client = client_basic if db_path == path_basic else client_abstract
            
        try:
            vectorstore = Chroma(
                client=current_client,
                collection_name=col_name,
                embedding_function=model,
            )
        except Exception as e:
            print(f"⚠️ 無法載入 Collection '{col_name}'，跳過。")
            continue

        results_list = []
        
        for project_title, project_info in tqdm.tqdm(apply_dicts.items(), desc=f"Searching {col_name}"):
            
            query_data = project_info.get(col_name)
            
            if isinstance(query_data, list):
                query_text = ",".join(query_data)
            elif isinstance(query_data, str):
                query_text = query_data
            else:
                query_text = ""

            if not query_text or len(query_text.strip()) < 2:
                continue

            documents = search(vectorstore, query_text)
            best_candidates = {} 
            
            for doc, raw_score in documents:
                score = round(raw_score, 6)
                
                # ★ 核心修改：取得推薦的「計畫名稱」與「主持人」
                recommended_title = doc.metadata.get('title', 'Unknown Project')
                recommended_manager = doc.metadata.get('manager', 'Unknown Manager')
                
                candidate_info = {
                    "apply_project": project_title,            # 申請案名稱
                    "manager": project_info['manager'],  # 申請人
                    "query_text": query_text,                  # 搜尋文本
                    "recommended_project": recommended_title,  # ★ 推薦的歷史計畫名稱
                    "recommended_manager": recommended_manager,# 該歷史計畫的主持人
                    "similarity_score": score,                 # 相似度分數
                    "matched_content": doc.page_content,       # 具體比對到的文字
                    "collection_field": col_name               # 比對的欄位
                }
                
                # ★ 核心修改：以「計畫名稱 (recommended_title)」作為去重依據
                if recommended_title in best_candidates:
                    if score > best_candidates[recommended_title]['similarity_score']:
                        best_candidates[recommended_title] = candidate_info
                else:
                    best_candidates[recommended_title] = candidate_info
            
            # 依照分數排序，每個申請案取前 30 個「最相似的計畫」
            candidates_sorted = sorted(best_candidates.values(), key=lambda x: x['similarity_score'], reverse=True)
            final_candidates = candidates_sorted[:30] 
            
            results_list.extend(final_candidates)
            
        # 將該類別的結果存入對應的 Excel Sheet
        save_data(output_file, results_list, col_name)

    print("\n🎉 全部搜尋完成！")

if __name__ == "__main__":
    main()

import pandas as pd
from langchain_community.embeddings import HuggingFaceEmbeddings # 改回這個比較穩定
from langchain_community.vectorstores import Chroma
import os
import chromadb
import tqdm

# --- 1. 設定模型 (需與存檔時一致) ---
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
            title = str(row['計畫中文名稱']).strip()
            if not title: continue

            # 處理關鍵字
            try:
                keywords_str = row['中文關鍵字']
                if keywords_str:
                    keywords_str = keywords_str.replace('，', ',').replace('；', ',').replace(';', ',').replace('、', ',').replace('。', ',')
                    keywords_str = keywords_str.replace('\n', ',').replace('\r', ',')
                    keywords = keywords_str.split('\n')
                    keywords = [k.strip() for k in keywords if k.strip()]
                else:
                    keywords = []
            except:
                keywords = []

            apply_dicts[title] = {
                'manager': str(row['計畫主持人']),
                'title': title,
                'keywords': keywords, # List
                'abstract': row['中文摘要'],
                'application_directions': row.get('application_directions', ''),
                'problems_to_solve': row.get('problems_to_solve', ''),
                'goals_to_achieve': row.get('goals_to_achieve', ''),
                'methods_to_solve': row.get('methods_to_solve', ''),
                'school': row['機關名稱']
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
def search(vectordb, query_text, RECOMMAND_AMOUNT=30):
    # 使用 similarity_search_with_score (Cosine Distance)
    # Chroma 預設是距離 (越小越好)，但 LangChain 的 relevance_score 會轉成 (0~1 越大越好)
    try:
        documents = vectordb.similarity_search_with_relevance_scores(
            query_text,
            k=RECOMMAND_AMOUNT,
            score_threshold=0.1 # 過濾掉太不相關的
        )
        return documents
    except Exception as e:
        print(f"搜尋錯誤: {e}")
        return []

def main():
    store_file_path = 'data/research_proj/115計算機學門審查/apply_project_with_abstract.xlsx'
    years = ['115'] # 假設這是申請年度
    
    # 1. 載入申請資料
    apply_dicts = load_data(store_file_path, years)
    
    # 2. 設定資料庫路徑與分類
    path_basic = "database/vectorstore_basic"
    path_abstract = "database/vectorstore_abstract"
    
    # 定義哪些欄位去哪個資料庫找
    collections_map = {
        'title': path_basic,
        'keywords': path_basic,
        'application_directions': path_abstract,
        'problems_to_solve': path_abstract,
        'goals_to_achieve': path_abstract,
        'methods_to_solve': path_abstract
    }
    
    # 您想要搜尋的欄位列表
    target_collections = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    
    # 初始化 Clients (只連線一次)
    print("正在連線至向量資料庫...")
    client_basic = chromadb.PersistentClient(path=path_basic)
    client_abstract = chromadb.PersistentClient(path=path_abstract)
    
    output_file = 'data/output/result_score.xlsx'
    if os.path.exists(output_file):
        try:
            os.remove(output_file)
            print(f"🗑️ 已刪除舊的輸出檔案: {output_file}")
        except Exception as e:
            print(f"⚠️ 無法刪除舊檔 (請檢查是否被 Excel 開啟中): {e}")
            return # 如果刪不掉就停止，避免寫入錯誤
        
    # 3. 開始搜尋
    for col_name in target_collections:
        # print(f"\n🔍 正在處理欄位: {col_name}")
        
        # 決定使用哪個 Client
        db_path = collections_map.get(col_name)
        if db_path == path_basic:
            current_client = client_basic
        else:
            current_client = client_abstract
            
        # 載入該欄位的 VectorStore
        try:
            vectorstore = Chroma(
                client=current_client,
                collection_name=col_name,
                embedding_function=model,
            )
        except Exception as e:
            print(f"⚠️ 無法載入 Collection '{col_name}' (可能不存在)，跳過。錯誤: {e}")
            continue

        results_list = []
        
        # 遍歷每一個申請計畫
        for project_title, project_info in tqdm.tqdm(apply_dicts.items(), desc=f"Searching {col_name}"):
            
            # 取得查詢文字
            query_data = project_info.get(col_name)
            
            # 處理 list 類型的 keywords
            if isinstance(query_data, list):
                query_text = ",".join(query_data)
            elif isinstance(query_data, str):
                query_text = query_data
            else:
                query_text = ""

            # 若內容太短或為空則跳過
            if not query_text or len(query_text.strip()) < 2:
                continue

            # 執行搜尋
            documents = search(vectorstore, query_text)
            
            # --- 去重與分數比對邏輯 ---
            # 這裡的邏輯是：
            # 1. 針對同一個申請案，搜尋出來的結果可能有多個是同一位教授 (因為該教授有多個過往計畫)。
            # 2. 我們只保留該位教授「分數最高」的那一筆紀錄。
            
            best_candidates = {} 
            
            for doc, score in documents:
                recommended_manager = doc.metadata.get('manager', 'Unknown')
                # 建立結果物件
                candidate_info = {
                    "project": project_title,           # 申請的計畫
                    "manager": project_info['manager'], # 申請人
                    "query_text": query_text,     # 查詢內容 (截短方便閱讀)
                    "matched_content": doc.page_content,
                    "recommended_manager": recommended_manager, # 推薦的審查人
                    "similarity_score": score,
                    "matched_doc_id": doc.metadata.get('title', 'N/A'), # 對應到的過往計畫標題(如果是Abstract)
                    "school": project_info['school'],
                    "collection_field": col_name
                }

                # 比對分數：如果這位教授已經在名單內，保留分數較高的那次
                if recommended_manager in best_candidates:
                    if score > best_candidates[recommended_manager]['similarity_score']:
                        best_candidates[recommended_manager] = candidate_info
                else:
                    best_candidates[recommended_manager] = candidate_info
            
            # 將整理好的最佳名單加入總表
            results_list.extend(list(best_candidates.values()))
            
        save_data(output_file, results_list, col_name)
        # print(f"✅ {col_name} 搜尋完成並存檔。")

    print("\n🎉 全部搜尋完成！")

if __name__ == "__main__":
    main()

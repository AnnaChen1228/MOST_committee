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
            title = str(row.get('計畫名稱') or row.get('計畫中文名稱') or "").strip()
            if not title: continue

            # 處理關鍵字
            try:
                keywords_str = row['中文關鍵字']
                if isinstance(keywords_str, list):
                    keywords = keywords_str
                elif keywords_str:
                    keywords_str = keywords_str.replace('，', ',').replace('；', ',').replace(';', ',').replace('、', ',').replace('。', ',')
                    keywords_str = keywords_str.replace('\n', ',').replace('\r', ',')
                    keywords = keywords_str.split('\n')
                    keywords = [k.strip() for k in keywords if k.strip()]
                else:
                    keywords = []
            except:
                keywords = []
            apply_dicts[title] = {
                'manager': str(row['主持人']),
                'title': title,
                'keywords': keywords, # List
                'abstract': row['中文摘要'],
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
    store_file_path = 'data/research_proj/115計算機學門審查/瑞典/apply_project_with_abstract(計畫書資料).xlsx'
    years = ['計畫書資料'] # 假設這是申請年度 
    # years = ['115-1電子資通領域(大產學)初審推薦名冊']
    # store_file_path = "data/industry_coop/apply_project_with_abstract(電子資通).xlsx"
    
    # 1. 載入申請資料
    apply_dicts = load_data(store_file_path, years)
    # # 2. 設定資料庫路徑與分類
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
    
    # 確保輸出檔案路徑正確
    output_file = 'data/research_proj/115計算機學門審查/瑞典/result_score(計畫書資料_research).xlsx'
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
        if db_path == path_basic:
            current_client = client_basic
        else:
            current_client = client_abstract
            
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
        
        # ★ 建議 2: 強制對申請案排序 (sorted)，確保每次執行的順序一模一樣
        # sorted_projects = sorted(apply_dicts.items(), key=lambda x: x[0])
        
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

            # 執行搜尋 (這裡會抓 50 筆)
            documents = search(vectorstore, query_text)
            best_candidates = {} 
            
            for doc, raw_score in documents:
                # ★ 建議 3: 收到分數立刻四捨五入 (小數點後 6 位)
                # 這是解決「同一台電腦結果不同」最關鍵的一步
                score = round(raw_score, 6)
                
                recommended_manager = doc.metadata.get('manager', 'Unknown')
                
                candidate_info = {
                    "project": project_title,
                    "manager": project_info['manager'],
                    "query_text": query_text,
                    "matched_content": doc.page_content,
                    "recommended_manager": recommended_manager,
                    "similarity_score": score, # 存入四捨五入後的分數
                    "matched_doc_id": doc.metadata.get('title', 'N/A'),
                    "collection_field": col_name
                }
                # 比對分數邏輯
                if recommended_manager in best_candidates:
                    # 如果分數比較高，就更新
                    if score > best_candidates[recommended_manager]['similarity_score']:
                        best_candidates[recommended_manager] = candidate_info
                    # ★ 建議 4: 如果分數「一樣」，則不動作 (保留原本的)
                    # 因為我們已經四捨五入過了，所以這裡的「一樣」是真的由數值決定，而不是浮點數誤差
                else:
                    best_candidates[recommended_manager] = candidate_info
            
            # ★ 建議 5: 搜尋完後，這裡再把所有候選人依照分數排序，只取前 30 名存檔
            # 這樣比直接讓 Chroma 截斷更穩定
            candidates_sorted = sorted(best_candidates.values(), key=lambda x: x['similarity_score'], reverse=True)
            final_candidates = candidates_sorted[:30] # 只取前 30 名寫入 Excel
            
            results_list.extend(final_candidates)
            
        save_data(output_file, results_list, col_name)

    print("\n🎉 全部搜尋完成！")

if __name__ == "__main__":
    main()

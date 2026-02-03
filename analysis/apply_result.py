import pandas as pd
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma
from langchain.schema import Document
import os
import chromadb
import tqdm
from langchain_community.embeddings import HuggingFaceBgeEmbeddings

# pass_project_with_ab apply_project_with_ab

model_name = 'BAAI/bge-large-zh-v1.5'
model_kwargs = {'device':'cpu'}
encode_kwargs = {'normalize_embeddings':True}
model = HuggingFaceBgeEmbeddings(
    model_name = model_name,
    model_kwargs = model_kwargs,
    encode_kwargs = encode_kwargs,
    # query_instruction="Represent this sentence for searching relevant passages:"
)

def load_data(file_path,pages):
    applpy_dicts = {}  # key: title, value: dict with titles, abstracts, keywords, name, project, value: abstract
    for pase in pages:
        apply_project_df = pd.read_excel(file_path, sheet_name=pase)
        for index, row in apply_project_df.iterrows():
            try:
                keywords_str = row['中文關鍵字']
                if pd.notna(keywords_str) and isinstance(keywords_str, str):
                    keywords_str = keywords_str.replace('，', ',').replace('；', ',').replace(';', ',').replace('、', ',').replace('。', ',')
                    keywords_str = keywords_str.replace('\n', ',').replace('\r', ',')
                    keywords = keywords_str.split(',')
                    keywords = [k.strip() for k in keywords if k.strip()]
                else:
                    keywords = []
            except Exception as e:
                print(f"處理 {row['計畫中文名稱']} 的關鍵字時出錯: {e}")
                keywords = []
            applpy_dicts[row['計畫中文名稱']] = {
                'manager': row['計畫主持人'],
                'title': row['計畫中文名稱'],
                'keywords': keywords,
                'abstract': row['中文摘要'],
                'application_directions': row['application_directions'],
                'problems_to_solve': row['problems_to_solve'],
                'goals_to_achieve': row['goals_to_achieve'],
                'methods_to_solve': row['methods_to_solve']
            }

        
    return applpy_dicts

def save_data(file_path, data_list, sheet_name):
    # 1. 如果資料是空的，就不存，避免報錯
    if not data_list:
        print(f"Sheet {sheet_name} has no data, skipping save.")
        return

    # 2. 將 List of Dicts 轉成 DataFrame
    df = pd.DataFrame(data_list)
    
    # 3. 判斷檔案是否存在，決定寫入模式
    if os.path.exists(file_path):
        # 檔案存在，用 'a' (append) 模式加入新 Sheet
        # if_sheet_exists='replace' 需要 pandas >= 1.3.0
        with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        # 檔案不存在，用 'w' (write) 模式建立新檔
        with pd.ExcelWriter(file_path, mode='w', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
    # print(f"Saved sheet: {sheet_name} to {file_path}")

def load_vector_db(chroma_db_path, collection_names):
    vectorstores = {}
    for name in collection_names:
        vectorstores[name] = Chroma(
            name,
            persist_directory=chroma_db_path,
            embedding_function=model
        )
    return vectorstores

def search(vectordb, data, RECOMMAND_AMOUNT=30):
    documents = vectordb.similarity_search_with_relevance_scores(
        data,
        k=RECOMMAND_AMOUNT
    )      
    return documents

def main():
    store_file_ptah = 'data/research_proj/115計算機學門審查/apply_project_with_abstract.xlsx'
    years = ['115']
    
    print("Loading Data...")
    # 這裡只回傳一個 apply_dicts 即可
    apply_dicts = load_data(store_file_ptah, years)
    
    collection_names = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    # 
    chroma_db_path = "database/vectorstore_weight_v3"
    
    print("Loading Vector DB...")
    vectorstores = load_vector_db(chroma_db_path, collection_names)
    
    final_results = {} 
    
    output_file = 'output/result_score.xlsx'
    # 1. 外層迴圈：遍歷不同欄位 (title, keywords, problems_to_solve...)
    for col_name in collection_names:
        
        print(f"Processing collection: {col_name}")
        final_results[col_name] = []
        vectorstore = vectorstores[col_name]
        
        # 2. 內層迴圈：遍歷每一個申請計畫
        # key 是計畫名稱 (title), value 是該計畫的所有欄位資料
        for project_title, project_info in tqdm.tqdm(apply_dicts.items(), desc=f"Searching {col_name}"):
            
            # --- 準備 Query Text ---
            query_data = project_info.get(col_name)
            
            # 如果是 keywords，它是一個 list，需要轉成字串
            if col_name == 'keywords' and isinstance(query_data, list):
                query_text = "\n".join(query_data)
            # 如果是其他欄位，確保它是字串
            elif isinstance(query_data, str):
                query_text = query_data
            else:
                query_text = ""

            # 如果內容是空的，就跳過
            if not query_text or not query_text.strip():
                continue

            # --- 搜尋 ---
            documents = search(vectorstore, query_text)
            
            # --- 去重邏輯 (同一個 Manager 取最高分) ---
            best_candidates = {} 
            
            for doc, score in documents:
                recommended_manager = doc.metadata.get('manager', 'Unknown')
                
                # 建立結果物件
                candidate_info = {
                    "project": project_title,    # 查詢的計畫名稱
                    "manager": project_info['manager'], # 查詢的計畫主持人
                    "query_text": query_text,
                    "compared_text": doc.page_content,
                    "recommended_manager": recommended_manager, # 推薦的審查委員
                    "similarity_score": score,
                    "collection_field": col_name
                }

                # 比較分數邏輯
                if recommended_manager in best_candidates:
                    if score > best_candidates[recommended_manager]['similarity_score']:
                        best_candidates[recommended_manager] = candidate_info
                else:
                    best_candidates[recommended_manager] = candidate_info
            
            # 將這個計畫針對這個欄位的最佳推薦名單加入總表
            final_results[col_name].extend(list(best_candidates.values()))
            save_data(output_file,final_results[col_name],col_name)

    # 簡單檢查輸出
    if 'title' in final_results and final_results['title']:
        print("\nSample Result (Title):")
        print(final_results['title'][0])

if __name__ == "__main__":
    main()
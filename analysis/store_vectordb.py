import pandas as pd
from langchain_community.embeddings import HuggingFaceEmbeddings
import chromadb # 直接使用 chromadb 原生套件
import os
import shutil
import tqdm

# --- 1. 設定 Embedding 模型 ---
model_name = 'BAAI/bge-large-zh-v1.5'
model_kwargs = {'device': 'cpu'}
encode_kwargs = {'normalize_embeddings': True}

print("正在載入模型...")
model = HuggingFaceEmbeddings(
    model_name=model_name,
    model_kwargs=model_kwargs,
    encode_kwargs=encode_kwargs,
)

# --- 2. 讀取資料 (維持原樣) ---
def load_data(file_path, pages):
    authors_dict = {}  
    projects_dict = {}  
    print(f"讀取 Excel: {file_path}")
    
    for page in pages:
        try:
            pass_project_df = pd.read_excel(file_path, sheet_name=page, dtype=str).fillna("")
        except Exception as e:
            print(f"跳過頁面 {page}: {e}")
            continue

        for index, row in pass_project_df.iterrows():
            author = str(row['計畫主持人']).strip()
            if not author: continue
            
            if author not in authors_dict:
                authors_dict[author] = {'title': [], 'keywords': []}
            
            title = str(row['計畫中文名稱']).strip()
            try:
                keywords_str = row['中文關鍵字']
                if keywords_str:
                    keywords_str = keywords_str.replace('，', ',').replace('；', ',').replace(';', ',').replace('、', ',').replace('。', ',')
                    keywords_str = keywords_str.replace('\n', ',').replace('\r', ',')
                    keywords = keywords_str.split(',')
                    keywords = [k.strip() for k in keywords if k.strip()]
                else:
                    keywords = []
            except:
                keywords = []
            
            authors_dict[author]['title'].append(title)
            authors_dict[author]['keywords'].extend(keywords)
            
            if title not in projects_dict:
                projects_dict[title] = {}
            projects_dict[title]['abstract'] = row['中文摘要']
            projects_dict[title]['application_directions'] = row.get('application_directions', '')
            projects_dict[title]['problems_to_solve'] = row.get('problems_to_solve', '')
            projects_dict[title]['goals_to_achieve'] = row.get('goals_to_achieve', '')
            projects_dict[title]['methods_to_solve'] = row.get('methods_to_solve', '')
            projects_dict[title]['manager'] = author

    for author in authors_dict:
        authors_dict[author]['keywords'] = list(set(authors_dict[author]['keywords']))
        
    return authors_dict, projects_dict

# --- 3. 儲存向量資料庫 (手動暴力版) ---
def store_vector_db(chroma_db_path, authors_dict, projects_dict, batch_size=50):
    # 強制刪除舊資料庫
    if os.path.exists(chroma_db_path):
        print(f"正在刪除舊資料庫 (確保乾淨): {chroma_db_path}")
        try:
            shutil.rmtree(chroma_db_path)
        except Exception as e:
            print(f"⚠️ 無法刪除資料夾 (可能被佔用): {e}")
            return

    os.makedirs(chroma_db_path, exist_ok=True)
    
    # 初始化 ChromaDB Client
    client = chromadb.PersistentClient(path=chroma_db_path)
    
    collection_names = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    abstract_names = ['application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    
    # 準備資料容器
    data_buckets = {name: {'ids': [], 'docs': [], 'metas': []} for name in collection_names}

    print("正在整理資料...")
    # 整理 Authors
    for key, value in tqdm.tqdm(authors_dict.items(), desc=f"正在整理資料..."):
        manager = key # 您的 ID
        
        title_text = "\n".join([t for t in value.get('title', []) if t])
        if title_text:
            data_buckets['title']['ids'].append(manager)
            data_buckets['title']['docs'].append(title_text)
            data_buckets['title']['metas'].append({'manager': manager})
            
        keywords_text = "\n".join([t for t in value.get('keywords', []) if t])
        if keywords_text:
            data_buckets['keywords']['ids'].append(manager)
            data_buckets['keywords']['docs'].append(keywords_text)
            data_buckets['keywords']['metas'].append({'manager': manager})

    # 整理 Projects
    for key, value in tqdm.tqdm(projects_dict.items(), desc=f"正在整理資料..."):
        title = key # 您的 ID
        manager = value.get('manager', '')
        meta = {'title': title, 'manager': manager}
        
        for name in abstract_names:
            text_content = value.get(name)
            if text_content and isinstance(text_content, str) and text_content.strip():
                data_buckets[name]['ids'].append(title)
                data_buckets[name]['docs'].append(text_content)
                data_buckets[name]['metas'].append(meta)

    # --- 核心修改：手動計算並寫入 ---
    print("🚀 開始計算向量並寫入 ChromaDB...")
    
    for name, data in data_buckets.items():
        ids = data['ids']
        docs = data['docs']
        metas = data['metas']
        total = len(ids)
        
        if total == 0:
            continue
            
        # 建立 Collection (指定 cosine 距離)
        collection = client.create_collection(name=name, metadata={"hnsw:space": "cosine"})
        
        print(f"正在處理 Collection: {name} (共 {total} 筆)")
        
        # 批次處理
        for i in tqdm.tqdm(range(0, total, batch_size), desc=f"Store {name}"):
            batch_ids = ids[i : i + batch_size]
            batch_docs = docs[i : i + batch_size]
            batch_metas = metas[i : i + batch_size]
            
            # 1. 手動計算向量 (這裡保證會有向量！)
            try:
                batch_embeddings = model.embed_documents(batch_docs)
            except Exception as e:
                print(f"❌ 計算向量失敗: {e}")
                continue
            
            # 檢查向量是否真的算出來了
            if not batch_embeddings or len(batch_embeddings) == 0:
                print("❌ 警告：算出空向量！")
                continue
                
            # 2. 手動寫入 (明確傳入 embeddings)
            collection.upsert(
                ids=batch_ids,
                documents=batch_docs,
                metadatas=batch_metas,
                embeddings=batch_embeddings # <--- 關鍵：絕對不能少
            )
            
    print("✅ 所有資料寫入完成。")

def main():
    store_file_path = 'data/research_proj/115計算機學門審查/pass_project_with_abstract.xlsx'
    if not os.path.exists(store_file_path):
        print(f"找不到檔案: {store_file_path}")
        return

    years = ['108','109','110','111','112','113','114']
    authors_dict, projects_dict = load_data(store_file_path, years)
    
    chroma_db_path = "database/vectorstore_weight_v3"
    store_vector_db(chroma_db_path, authors_dict, projects_dict)
    
    print('---finish---')

if __name__ == "__main__":
    main()

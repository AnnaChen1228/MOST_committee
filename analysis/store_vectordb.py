import pandas as pd
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma
from langchain_core.documents import Document
import os
import shutil
import tqdm

# --- 1. 設定 Embedding 模型 ---
model_name = 'BAAI/bge-large-zh-v1.5'
model_kwargs = {'device': 'cpu'}
encode_kwargs = {'normalize_embeddings': True}

print("正在載入模型...")
embedding_model = HuggingFaceEmbeddings(
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
            
            title = str(row.get('計畫名稱') or row.get('計畫中文名稱') or "").strip()
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
            else:
                print(title)
            projects_dict[title]['abstract'] = row['中文摘要']
            projects_dict[title]['application_directions'] = row.get('application_directions', '')
            projects_dict[title]['problems_to_solve'] = row.get('problems_to_solve', '')
            projects_dict[title]['goals_to_achieve'] = row.get('goals_to_achieve', '')
            projects_dict[title]['methods_to_solve'] = row.get('methods_to_solve', '')
            projects_dict[title]['manager'] = author

    for author in authors_dict:
        authors_dict[author]['keywords'] = list(set(authors_dict[author]['keywords']))
        
    return authors_dict, projects_dict

# --- 3A. 儲存 Basic 資料庫 (Title, Keywords) ---
def store_basic_db(chroma_db_path, authors_dict, embedding_model):
    print(f"\n🔵 正在處理 Basic 資料庫 (Title/Keywords) -> {chroma_db_path}")
    
    # 清理舊資料夾
    if os.path.exists(chroma_db_path):
        try:
            shutil.rmtree(chroma_db_path)
            print("  已刪除舊資料夾")
        except Exception as e:
            print(f"  ⚠️ 無法刪除資料夾: {e}")
            return

    # 準備資料
    data_store = {
        'title': {'docs': [], 'ids': []},
        'keywords': {'docs': [], 'ids': []}
    }

    # 整理 Authors 資料
    for manager, value in tqdm.tqdm(authors_dict.items(), desc="Processing Basic Info"):
        # Title
        title_text = "\n".join([t for t in value.get('title', []) if t])
        if title_text:
            doc = Document(page_content=title_text, metadata={'manager': manager, 'type': 'title'})
            data_store['title']['docs'].append(doc)
            data_store['title']['ids'].append(manager)

        # Keywords
        keywords_text = "\n".join([t for t in value.get('keywords', []) if t])
        if keywords_text:
            doc = Document(page_content=keywords_text, metadata={'manager': manager, 'type': 'keywords'})
            data_store['keywords']['docs'].append(doc)
            data_store['keywords']['ids'].append(manager)

    # 寫入 ChromaDB
    for col_name, data in data_store.items():
        docs = data['docs']
        ids = data['ids']
        
        if not docs: continue
            
        print(f"  正在建立 Collection: {col_name} (共 {len(docs)} 筆)...")
        try:
            Chroma.from_documents(
                documents=docs,
                embedding=embedding_model,
                persist_directory=chroma_db_path,
                collection_name=col_name,
                ids=ids
            )
            print(f"  ✅ {col_name} 儲存成功。")
        except Exception as e:
            print(f"  ❌ {col_name} 儲存失敗: {e}")

# --- 3B. 儲存 Abstract 資料庫 (長文本) ---
def store_abstract_db(chroma_db_path, projects_dict, embedding_model):
    print(f"\n🟠 正在處理 Abstract 資料庫 (長文本) -> {chroma_db_path}")
    
    # 清理舊資料夾
    if os.path.exists(chroma_db_path):
        try:
            shutil.rmtree(chroma_db_path)
            print("  已刪除舊資料夾")
        except Exception as e:
            print(f"  ⚠️ 無法刪除資料夾: {e}")
            return

    # 準備資料
    data_store = {
        'application_directions': {'docs': [], 'ids': []},
        'problems_to_solve': {'docs': [], 'ids': []},
        'goals_to_achieve': {'docs': [], 'ids': []},
        'methods_to_solve': {'docs': [], 'ids': []}
    }
    
    abstract_names = ['application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']

    # 整理 Projects 資料 (這裡我幫你把原本註解掉的迴圈打開了)
    for title, value in tqdm.tqdm(projects_dict.items(), desc="Processing Abstracts"):
        manager = value.get('manager', '')
        
        for name in abstract_names:
            text_content = value.get(name)
            if text_content and isinstance(text_content, str) and text_content.strip():
                
                doc = Document(
                    page_content=text_content,
                    metadata={'title': title, 'manager': manager, 'field': name}
                )
                data_store[name]['docs'].append(doc)
                data_store[name]['ids'].append(title)

    # 寫入 ChromaDB
    for col_name, data in data_store.items():
        docs = data['docs']
        ids = data['ids']
        
        if not docs: continue
            
        print(f"  正在建立 Collection: {col_name} (共 {len(docs)} 筆)...")
        try:
            Chroma.from_documents(
                documents=docs,
                embedding=embedding_model,
                persist_directory=chroma_db_path,
                collection_name=col_name,
                ids=ids
            )
            print(f"  ✅ {col_name} 儲存成功。")
        except Exception as e:
            print(f"  ❌ {col_name} 儲存失敗: {e}")

def main():
    industry_store_file_path = 'data/industry_coop/pass_project_with_abstract.xlsx'
    research_store_file_path = 'data/research_proj/115計算機學門審查/pass_project_with_abstract.xlsx'
    if not os.path.exists(industry_store_file_path):
        print(f"找不到檔案: {industry_store_file_path}")
        return
    if not os.path.exists(research_store_file_path):
        print(f"找不到檔案: {research_store_file_path}")
        return
    industry_years = ['專題計畫綜合查詢']
    research_years = ['108','109','110','111','112','113','114']
    # 1. 讀取資料
    industry_authors_dict, industry_projects_dict = load_data(industry_store_file_path, industry_years)
    research_authors_dict, research_projects_dict = load_data(research_store_file_path, research_years)

    all_authors_dict = industry_authors_dict.copy()
    all_authors_dict.update(research_authors_dict)

    all_projects_dict = industry_projects_dict.copy()
    all_projects_dict.update(research_projects_dict)
    # # 2. 定義兩個不同的資料庫路徑
    path_basic = "database/vectorstore_basic_all"       # 存 Title, Keywords
    path_abstract = "database/vectorstore_abstract_all" # 存 Abstracts
    
    # # 3. 分別執行儲存
    store_basic_db(path_basic, all_authors_dict, embedding_model)
    store_abstract_db(path_abstract, all_projects_dict, embedding_model)
    
    print('\n🎉 --- All Finished ---')

if __name__ == "__main__":
    main()

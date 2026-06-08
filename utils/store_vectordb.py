import pandas as pd
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Chroma
from langchain_core.documents import Document
from utils.get_setting import setting_data, print_setting_data, find_key_path, value_of_key

model_kwargs = {'device': 'cpu'}
encode_kwargs = {'normalize_embeddings': True}

embedding_model = HuggingFaceEmbeddings(
    model_name=value_of_key('EMBEDDING_MODEL_NAME'),
    model_kwargs=model_kwargs,
    encode_kwargs=encode_kwargs,
)

def generate_store_datas(project_data, key_names):
    data_store = {}
    for name in key_names:
        data_store[name] = {'docs': [], 'ids': []}

    for key, value in project_data.items():
        manager = value.get('host_name', '')
        pass_year = str(value.get('pass_year', ''))
        for name in key_names:
            text_content = value.get(name)
            if text_content and isinstance(text_content, str) and text_content.strip():
                
                doc = Document(
                    page_content=text_content,
                    metadata={'title': key, 'manager': manager, 'field': name}
                )
                data_store[name]['docs'].append(doc) #因為批次所以才合再一起
                data_store[name]['ids'].append(f'{key}_{manager}_{pass_year}') #因為批次所以才合再一起
    return data_store

def store_basic_vector_db(type, project_data):
    if type == 'research':
        chroma_db_path = find_key_path('CHROMA_BASIC')
    elif type == 'industry':
        chroma_db_path = find_key_path('CHROMA_BASIC_INDUSTRY')
    else:
        raise ValueError("Invalid type. Must be 'research' or 'industry'.")
    
    data_store = generate_store_datas(project_data, ['title', 'keywords'])
    for name, data in data_store.items():
        docs = data['docs']
        ids = data['ids']
        if not docs: 
            continue
        vector_db = Chroma(
            persist_directory=chroma_db_path,
            embedding_function=embedding_model,
            collection_name=name
        )
        vector_db.add_documents(documents=docs, ids=ids)

def store_abstract_vector_db(type, project_data):
    if type == 'research':
        chroma_db_path = find_key_path('CHROMA_ABSTRACT')
    elif type == 'industry':
        chroma_db_path = find_key_path('CHROMA_ABSTRACT_INDUSTRY')
    else:
        raise ValueError("Invalid type. Must be 'research' or 'industry'.")
    
    data_store = generate_store_datas(project_data, ['application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve'])
    for name, data in data_store.items():
        docs = data['docs']
        ids = data['ids']
        if not docs: 
            continue
        vector_db = Chroma(
            persist_directory=chroma_db_path,
            embedding_function=embedding_model,
            collection_name=name
        )
        vector_db.add_documents(documents=docs, ids=ids)

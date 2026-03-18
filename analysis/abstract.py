import ollama
import pandas as pd
import json
import time
from datetime import datetime, timedelta
from tqdm.auto import tqdm
import os

ollama_client = ollama.Client(host='http://localhost:1228')

prompt = '''
你是一位專門分析科研計畫摘要的助理。請從提供的計畫摘要中準確擷取以下四個關鍵要素：

1. 應用方向：此計畫的預期應用領域或產業方向
2. 欲解決問題：此計畫試圖解決的具體問題或挑戰
3. 達成目標：此計畫希望實現的具體目標或成果
4. 解決方法：此計畫採用的技術、方法或途徑

請嚴格遵守以下規則：
- 直接從摘要中擷取相關文字與句子，不要改寫或添加你的解釋，僅提取原文內容
- 如有多個相關內容，請用"；"分隔
- 若摘要中未明確提及某項內容，該欄位請填寫"摘要未明確說明"
- 確保擷取的內容完整，不要隨意截斷句子
- 僅輸出JSON格式的結果，不要有其他說明文字

輸出格式必須為有效的JSON：
{
    "application_directions": "從摘要中擷取的應用方向",
    "problems_to_solve": "從摘要中擷取的欲解決問題",
    "goals_to_achieve": "從摘要中擷取的達成目標",
    "methods_to_solve": "從摘要中擷取的解決方法"
}

'''

schema = {
    "type": "object",
    "properties": {
        "application_directions": {"type": "string","description": "此計畫的應用方向為何"},
        "problems_to_solve": {"type": "string","description": "此計畫預計要解決的問題為何"},
        "goals_to_achieve": {"type": "string","description": "此計畫預計要達成的目標為何"},
        "methods_to_solve": {"type": "string","description": "此計畫預計要使用的解決方法為何"}
    },
    "required": [
        "application_directions",
        "problems_to_solve",
        "goals_to_achieve",
        "methods_to_solve",
    ],
}

def local_generate(model,systemPrompt,userPrompt,schema):
    messages = [
        {"role": "system", "content": systemPrompt},
        {"role": "user", "content": userPrompt}
    ]
    response = ollama_client.chat(
        model=model,
        messages=messages,
        format=schema,
        options={"num_ctx": 128000}
    )
    return response

def generate_by_file(file,years,output_file):
    if not os.path.exists(output_file):
        print(f"創建新的輸出文件: {output_file}")
        # 創建一個空白的 Excel 文件
        with pd.ExcelWriter(output_file, mode='w', engine='openpyxl') as writer:
            pd.DataFrame().to_excel(writer, sheet_name='初始化', index=False)
    for year in years:
        year_start_time = time.time()
        start_time = time.time()
        print(f"\n開始處理 {year} 年度資料...")
        
        apply_project_df = pd.read_excel(file, year, engine='openpyxl') #審查資料
        apply_project_df = apply_project_df.fillna('')
        total_projects = len(apply_project_df)
        print(f"共有 {total_projects} 個項目需要處理")
        # 創建一個新的 DataFrame 來存儲結果
        results_df = pd.DataFrame()
        
        # 使用 tqdm 創建進度條
        with tqdm(total=total_projects, desc=f"處理 {year} 年度資料", unit="項目") as pbar:
            # 記錄處理時間以計算平均速度
            processing_times = []
            
            # 遍歷每個項目
            for i, project in enumerate(apply_project_df.itertuples()):
                # if i < 120:
                #     continue
                item_start_time = time.time()
                # 獲取原始資料
                project_dict = project._asdict()
                
                # 移除索引欄位（如果不需要）
                if 'Index' in project_dict:
                    del project_dict['Index']
                # 獲取摘要
                abstract = project_dict['中文摘要']
                if isinstance(abstract, str):
                    abstract = abstract.replace('_x000D', '')  # 清理特殊字符
                if not abstract:
                    continue
                # 使用 LLM 生成結果
                llama_response = local_generate('gpt-oss:120b', prompt, abstract, schema)
                raw_data = json.loads(llama_response.message.content)
                
                # 將 LLM 結果添加到項目字典中
                project_dict["application_directions"] = raw_data.get("application_directions", "")
                project_dict["problems_to_solve"] = raw_data.get("problems_to_solve", "")
                project_dict["goals_to_achieve"] = raw_data.get("goals_to_achieve", "")
                project_dict["methods_to_solve"] = raw_data.get("methods_to_solve", "")
                project_dict["keywords"] = raw_data.get("keywords", "")
                # 將項目添加到結果 DataFrame
                results_df = pd.concat([results_df, pd.DataFrame([project_dict])], ignore_index=True)
                
                # 計算此項目處理時間
                item_time = time.time() - item_start_time
                processing_times.append(item_time)
                
                # 更新進度條
                pbar.update(1)
                
                # 計算並顯示平均處理時間和預計完成時間
                avg_time = sum(processing_times) / len(processing_times)
                remaining_items = total_projects - (i + 1)
                est_remaining_time = avg_time * remaining_items
                est_completion_time = datetime.now() + timedelta(seconds=est_remaining_time)
                
                # 更新進度條描述
                pbar.set_postfix({
                    "平均時間": f"{avg_time:.2f}秒/項", 
                    "預計完成": est_completion_time.strftime("%H:%M:%S")
                })
                
                # 每處理 10 個項目保存一次（防止意外丟失數據）
                if (i + 1) % 10 == 0:
                    with pd.ExcelWriter(output_file, mode='a', engine='openpyxl', 
                                    if_sheet_exists='replace') as writer:
                        results_df.to_excel(writer, sheet_name=year, index=False)
                    pbar.set_postfix({
                        "平均時間": f"{avg_time:.2f}秒/項", 
                        "預計完成": est_completion_time.strftime("%H:%M:%S"),
                        "已保存": f"{i+1}項"
                    })

        # 計算年度處理總時間
        year_time = time.time() - year_start_time
        year_time_str = str(timedelta(seconds=int(year_time)))
        
        # 最終保存完整結果
        with pd.ExcelWriter(output_file, mode='a', engine='openpyxl', 
                        if_sheet_exists='replace') as writer:
            results_df.to_excel(writer, sheet_name=year, index=False)
        
        print(f"{year} 年度處理完成，共 {len(results_df)} 個項目，耗時: {year_time_str}")
        
        # 計算平均處理時間
        avg_time_per_item = year_time / total_projects
        print(f"平均每項目處理時間: {avg_time_per_item:.2f} 秒")

    # 總處理時間
    total_time = time.time() - start_time
    total_time_str = str(timedelta(seconds=int(total_time)))
    print(f"\n全部處理完成，總耗時: {total_time_str}")

def main():
    # print('---start pass project abstract generation---')
    # path_file_path = 'data/industry_coop/108-114核定計畫名單(含關鍵字與摘要)20260223_產學計畫案.xlsx'
    # pass_project_excel_file = pd.ExcelFile(path_file_path)
    # years = ['108-114']
    # pass_output_file = "data/industry_coop/pass_project_with_abstract_108-114.xlsx"
    # generate_by_file(pass_project_excel_file, years, pass_output_file)

    print('---start apply project abstract generation---')
    apply_file_path = 'data/research_proj/115計算機學門審查/附件一_115技專預推名冊(E41)智慧計算.xlsx'
    apply_project_excel_file = pd.ExcelFile(apply_file_path)
    apply_years = ['預推名冊(E41智慧計算)']
    apply_output_file = "data/research_proj/115計算機學門審查/apply_project_with_abstract(附件一_115技專預推名冊(E41)智慧計算).xlsx"
    generate_by_file(apply_project_excel_file, apply_years, apply_output_file)
if __name__ == "__main__":
    main()



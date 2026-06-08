import ollama
import base64
import concurrent.futures
from utils.get_setting import setting_data, print_setting_data, find_key_path, value_of_key

ollama_client = ollama.Client(
    host=value_of_key('OLLAMA_HOST'),  # 替換成你的 Ollama 伺服器地址和端口
)
# auth_str = f"{value_of_key('OLLAMA_USER')}:{value_of_key('OLLAMA_PASSWORD')}"
# b64_auth = base64.b64encode(auth_str.encode('utf-8')).decode('utf-8')
# ollama_client._client.headers["Authorization"] = f"Basic {b64_auth}"
MODEL_NAME = value_of_key('LLM_MODEL_NAME')

SYSTEMPROMPT = '''
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

SCHEMA = {
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

def local_generate(abstract_text, timeout_seconds=60):
    userPrompt = f"請分析以下科研計畫摘要，並擷取四個關鍵要素：\n\n{abstract_text}"
    messages = [
        {"role": "system", "content": SYSTEMPROMPT},
        {"role": "user", "content": userPrompt}
    ]
    
    if not MODEL_NAME:
        raise ValueError("模型名稱未設定！請檢查 .yaml 檔案。")

    def _call_ollama():
        return ollama_client.chat(
            model=MODEL_NAME,
            messages=messages,
            format=SCHEMA,
            options={"num_ctx": 128000}
        )

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(_call_ollama)
            response = future.result(timeout=timeout_seconds)
            return response
            
    except concurrent.futures.TimeoutError:
        # 發生超時，直接拋出錯誤中斷程式
        raise RuntimeError("server暫時無法使用")
        
    except Exception as e:
        # 發生其他錯誤，也拋出
        raise RuntimeError(f"Ollama 伺服器發生錯誤: {e}")
import yaml
import os

# - 打開 YAML 檔案並讀取數據
setting_data = None
with open('./setting.yaml', 'r', encoding='utf-8') as file:
    setting_data = yaml.safe_load(file)

# - 本次審查專屬的輸出資料夾（執行時由 GUI 設定）
#   設定後，下列「每次審查產出的結果/過程檔案」會改存到此資料夾，
#   其餘（如人才資料庫等需跨次保留的檔案）維持原本位置。
_RUN_OUTPUT_DIR = None
_RUN_OUTPUT_KEYS = {
    "統計表分析",        # 推薦表統合與分析
    "過濾相近後統計表",
    "FINAL_COMMITTEE",   # 最終推薦表
    "申請計畫拆解結果",
}

def set_run_output_dir(path):
    '''設定本次審查的輸出資料夾，若不存在則建立。'''
    global _RUN_OUTPUT_DIR
    _RUN_OUTPUT_DIR = path
    if path and not os.path.exists(path):
        os.makedirs(path)

def get_run_output_dir():
    '''取得本次審查的輸出資料夾（未設定時回傳 None）。'''
    return _RUN_OUTPUT_DIR

# - 找到路徑過程
def find_key_path_list(dictionary, target_key, path=''):
    for key, value in dictionary.items():
        
        if key == target_key:
            temp_path = path + str(value)
            result_list = temp_path.split(" -> ")
            return result_list
        
        elif isinstance(value, dict):
            result = find_key_path_list(value, target_key, path + key + ' -> ')
            if result: return result
        
    return []

# - 輸出路徑
def find_key_path(target_key):
    path_list = find_key_path_list(setting_data, target_key)

    # 本次審查的結果/過程檔案 → 改存到本次審查專屬資料夾
    if _RUN_OUTPUT_DIR and target_key in _RUN_OUTPUT_KEYS and path_list:
        result_text = os.path.join(_RUN_OUTPUT_DIR, path_list[-1])
    else:
        result_text = "."
        for node in path_list[1:]:
            result_text = result_text + "/" + node

    directory = os.path.dirname(result_text)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)

    return result_text

# - 直接找到資料 value
def value_of_key(target_key, dictionary=setting_data):
    
    for key, value in dictionary.items():
        if key == target_key: return value
        elif isinstance(value, dict):
            result = value_of_key(target_key, value)
            if result: return result
        
    return ""

# - 輸出設定數據
def print_setting_data():
    
    print("v ================================ DATA SETTING ================================")
    def print_dict(d, indent=0):
        for key, value in d.items():
            print('  ' * indent + str(key))
            if isinstance(value, dict):
                print_dict(value, indent+1)
            else:
                print('  ' * (indent+1) + str(value))
    
    # 使用遞迴函數打印 settings 字典
    print_dict(setting_data)
    print("^ ================================ DATA SETTING ================================")
    
    
    
if __name__ == '__main__':
    print_setting_data()
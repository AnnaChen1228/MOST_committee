import argparse
from utils.get_setting import value_of_key
from utils.script import load_into_chroma_bge_manager, search_v3, filter_committee, excel_process_VBA, statistic_committee, update_peronsal_info_database
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="onnxruntime")

def main():
    
    # : parser setting
    parser = argparse.ArgumentParser(description="腳本執行器")
    parser.add_argument('--choose_mode', default='輸出推薦委員', type=str, help='選擇模式（"存入資料庫", "輸出推薦委員"）', choices=['存入資料庫', '輸出推薦委員'])
    parser.add_argument('--is_industry', type=str, default='False', help='是否為產學計劃')
    parser.add_argument('--is_load_chroma_bge', type=str, default='False', help='是否重新匯入資料')
    args = parser.parse_args()
    
    if value_of_key("目前執行計畫") == "產學合作": is_industry = True
    elif value_of_key("目前執行計畫") == "研究計畫": is_industry = False
    else: 
        print("「目前執行計畫」沒有填寫正確（產學合作｜研究計畫）")
        return
    
    # 匯入資料庫
    if args.choose_mode == '存入資料庫':    
        load_into_chroma_bge_manager(is_industry)
    elif args.choose_mode == '輸出推薦委員':
        update_peronsal_info_database(is_industry)
        statistic_committee() #= output: 統計清單人才資料_RDF
        search_v3(is_industry) 
        filter_committee(is_industry) #= 篩選人員
        excel_process_VBA()
        print("已成功輸出推薦委員")
        
    # 

if __name__ == "__main__":
    
    main()
    
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def load_data(path):
    data = pd.read_excel(path, dtype=str).fillna("")
    return data

def load_data_by_page(path,pages):
    data = {}
    for page in pages:
        data[page] = pd.read_excel(path,sheet_name=page,dtype=str).fillna("")
    return data
def save_data(data,path):
    df = pd.DataFrame(list(data.items()), columns=['委員姓名', '推薦次數'])
    
    # 2. 依照推薦次數由高到低排序 (選配，但建議執行)
    df = df.sort_values(by='推薦次數', ascending=False)
    
    # 3. 儲存為 CSV
    # encoding='utf-8-sig' 是確保 Excel 開啟中文不亂碼的關鍵
    df.to_csv(path, index=False, encoding='utf-8-sig')
    
    print(f"✅ 資料已成功儲存至: {path}")
          
def count_recommendations(df):
    recommend_df = df.filter(regex='Recommend|推薦委員')
    all_names = recommend_df.stack()
    name_counts = all_names.value_counts().to_dict()
    return name_counts
    
def count_fig(count_dict, average_score):
    # --- Mac 中文設定 ---
    plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False

    # 1. 取得數值並計算統計量
    counts = list(count_dict.values())
    if not counts: return
    
    mean_val = np.mean(counts)  # 平均值
    max_val = max(counts)      # 最大值
    total_people = len(counts)  # 總人數

    # 2. 定義區間 (Bins)
    bins = list(range(0, max_val + 6, 5)) 
    labels = [f"{bins[i]+1}-{bins[i+1]}" for i in range(len(bins)-1)]
    
    # 3. 分組統計
    df_counts = pd.cut(counts, bins=bins, labels=labels, right=True)
    dist = df_counts.value_counts().sort_index()
    
    # 4. 繪圖
    plt.figure(figsize=(14, 6))
    bars = plt.bar(dist.index.astype(str), dist.values, color='teal', edgecolor='black', alpha=0.7)
    
    # 在長條上方顯示人數
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), 
                 ha='center', va='bottom', fontweight='bold')

    # 5. 顯示平均值與統計資訊
    # 在圖表內建立一個文字框
    stats_text = f"總人數: {total_people} 人\n平均推薦: {mean_val:.2f} 次\n最大次數: {max_val} 次\n平均相似度分數: {average_score}"
    plt.gca().text(0.95, 0.95, stats_text, transform=plt.gca().transAxes,
                   fontsize=12, verticalalignment='top', horizontalalignment='right',
                   bbox=dict(boxstyle='round', facecolor='white', alpha=0.5))

    # 設定標題與標籤
    plt.title(f'委員被推薦次數分佈 (平均值: {mean_val:.2f} 次)', fontsize=16)
    plt.xlabel('被推薦次數區間', fontsize=12)
    plt.ylabel('委員人數', fontsize=12)
    plt.grid(axis='y', linestyle=':', alpha=0.5)
    
    plt.tight_layout()
    plt.show()

def calculate_average_similarity(df):
    """提取所有分數欄位並計算總平均值"""
    # 1. 篩選出所有名稱包含 'Score' 的欄位
    score_df = df.filter(regex='Score|相關分數|similarity_score')
    
    if score_df.empty:
        print("⚠️ 找不到包含 'Score' 的分數欄位，請檢查 Excel 標題。")
        return 0
    
    # 2. 強制轉換為數值型態
    # errors='coerce' 會把無法轉換的文字(如空字串)變成 NaN
    score_df_numeric = score_df.apply(pd.to_numeric, errors='coerce')
    
    # 3. 攤平並排除 NaN 以及 0 (代表沒分數的資料)
    all_scores = score_df_numeric.stack()
    all_scores = all_scores[all_scores > 0]
    
    if all_scores.empty:
        print("⚠️ 分數欄位中沒有有效的數值資料。")
        return 0
        
    # 4. 計算平均值
    avg_score = all_scores.mean()
    return avg_score

def compare_overlap(df1, df2):
    """比較兩個 DataFrame 每一行的推薦名單重疊比例與分類個數"""
    # 1. 提取推薦委員的欄位
    rec1 = df1.filter(regex='Recommend|推薦委員')
    rec2 = df2.filter(regex='Recommend|推薦委員')

    num_rows = min(len(rec1), len(rec2))
    ratios = []

    for i in range(num_rows):
        # 清理資料：轉字串、去空格、排除無效值
        names1 = set([str(n).strip() for n in rec1.iloc[i] if str(n).strip() not in ["", "0", "nan", "None"]])
        names2 = set([str(n).strip() for n in rec2.iloc[i] if str(n).strip() not in ["", "0", "nan", "None"]])

        if not names1 and not names2:
            ratios.append(1.0)  # 兩邊都沒人，視為一致
            continue
        
        intersection = names1.intersection(names2)
        union = names1.union(names2)
        
        # 計算 Jaccard 相似度
        ratio = len(intersection) / len(union) if union else 0
        ratios.append(ratio)

    # 2. 統計分類個數
    full_match = sum(1 for r in ratios if r == 1.0)
    no_match = sum(1 for r in ratios if r == 0.0)
    partial_match = sum(1 for r in ratios if 0.0 < r < 1.0)
    
    avg_overlap = np.mean(ratios) if ratios else 0

    return {
        "avg_overlap": avg_overlap,
        "full_match": full_match,
        "no_match": no_match,
        "partial_match": partial_match,
        "total_cases": num_rows,
        "all_ratios": ratios
    }

def main():
    input_path = 'data/output/recommendation_results_org.xlsx'
    org_path = 'data/output/(勿對外公開資料或流傳)108-115年智慧計算學門大批專題計畫申請案件(含中英文摘要及關鍵字)_推薦表統合_VBA.xlsx'
    data = load_data(input_path)
    org_data = load_data(org_path)
    recommand_count = count_recommendations(data)
    average_score = calculate_average_similarity(data)
    print(average_score)
    save_data(recommand_count,'data/output/recommand_count.csv')
    count_fig(recommand_count, average_score)
    # org_recommand_count = count_recommendations(org_data)
    # org_average_score = calculate_average_similarity(org_data)
    # print(org_average_score)
    # save_data(org_recommand_count,'data/output/org_recommand_count.csv')
    # count_fig(org_recommand_count, org_average_score)

    # result = compare_overlap(data, org_data)

    # print("\n" + "═"*40)
    # print(f"🔍 推薦名單比對摘要報告")
    # print("═"*40)
    # print(f"總分析案件數： {result['total_cases']} 案")
    # print(f"平均重疊比例： {result['avg_overlap']:.2%}")
    # print("-" * 40)
    # print(f"✅ 完全一致 (100%)： {result['full_match']} 案")
    # print(f"⚠️ 部分重疊 (1-99%)： {result['partial_match']} 案")
    # print(f"❌ 完全不同 (0%)  ： {result['no_match']} 案")
    # print("═"*40 + "\n")

    # item_data_file = 'data/output/result_score.xlsx'
    # pages = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    # item_data = load_data_by_page(item_data_file,pages)
    # avg = {}
    # for key,value in item_data.items():
    #     avg[key] = calculate_average_similarity(value)
    # print(avg)

if __name__ == '__main__':
    main()
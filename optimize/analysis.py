import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import seaborn as sns
import platform
import os
from datetime import datetime

# --- 1. 資料讀取與寫入 ---

def load_data(path):
    try:
        data = pd.read_excel(path, dtype=str).fillna("")
        return data
    except FileNotFoundError:
        print(f"❌ 找不到檔案: {path}")
        return pd.DataFrame()

def load_data_by_page(path, pages):
    data = {}
    try:
        xls = pd.ExcelFile(path)
        for page in pages:
            if page in xls.sheet_names:
                data[page] = pd.read_excel(xls, sheet_name=page, dtype=str).fillna("")
            else:
                print(f"⚠️ 警告: Sheet '{page}' 不存在於檔案中。")
                data[page] = pd.DataFrame()
    except FileNotFoundError:
        print(f"❌ 找不到檔案: {path}")
    return data

def save_data(data, path):
    directory = os.path.dirname(path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)

    df = pd.DataFrame(list(data.items()), columns=['委員姓名', '推薦次數'])
    df = df.sort_values(by='推薦次數', ascending=False)
    df.to_csv(path, index=False, encoding='utf-8-sig')
    print(f"✅ CSV 資料已儲存: {path}")

# --- 2. 統計計算 ---

def count_recommendations(df):
    if df.empty: return {}
    
    # 1. 先抓出所有包含 "Recommend" 或 "推薦委員" 的欄位
    # 這時候會包含 "推薦委員" 和 "推薦委員學校"
    potential_cols = df.filter(regex='Recommend|推薦委員|選取委員').columns
    
    # 2. 設定要排除的關鍵字
    # 只要欄位名稱裡有 "學校" 或 "School"，就踢掉
    exclude_keywords = ['學校', 'School', '單位', 'Unit'] 
    
    # 3. 進行過濾：保留「不包含」排除關鍵字的欄位
    target_cols = [c for c in potential_cols if not any(bad_word in c for bad_word in exclude_keywords)]
    
    # --- 除錯用：你可以把下面這行取消註解，看看最後抓到了哪些欄位 ---
    # print(f"最終使用的欄位: {target_cols}")
    
    # 4. 只使用過濾後的欄位來取資料
    recommend_df = df[target_cols]
    
    all_names = recommend_df.stack()
    
    # 轉字串並去除空白
    all_names = all_names.astype(str).str.strip()
    
    # 過濾無效值
    valid_names = all_names[~all_names.isin(["", "nan", "NaN", "0", "None"])]
    
    name_counts = valid_names.value_counts().to_dict()
    return name_counts


def calculate_average_similarity(df):
    if df.empty: return 0
    score_df = df.filter(regex='Score|相關分數|similarity_score')
    if score_df.empty:
        return 0
    score_df_numeric = score_df.apply(pd.to_numeric, errors='coerce')
    all_scores = score_df_numeric.stack()
    all_scores = all_scores[all_scores > 0]
    if all_scores.empty: return 0
    return all_scores.mean()

# --- 3. 視覺化繪圖 ---

def count_fig(count_dict, average_score, title_prefix="", save_path=None):
    """(單一資料集) 繪製推薦次數分佈圖 (雙軸)"""
    if platform.system() == 'Darwin':
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
    else:
        plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    counts = list(count_dict.values())
    if not counts: return
    
    mean_val = np.mean(counts)
    max_val = max(counts)
    total_people = len(counts)

    step = 10
    if max_val > 50: step = 10
    bins = list(range(0, max_val + step + 1, step)) 
    labels = [f"{bins[i]+1}-{bins[i+1]}" for i in range(len(bins)-1)]
    
    df_counts = pd.cut(counts, bins=bins, labels=labels, right=True)
    dist = df_counts.value_counts().sort_index()
    
    x_labels = dist.index.astype(str)
    y_bar = dist.values
    y_cumulative = np.cumsum(y_bar) / total_people * 100

    fig, ax1 = plt.subplots(figsize=(20, 7))

    color_bar = '#4c72b0'
    bars = ax1.bar(x_labels, y_bar, color=color_bar, alpha=0.8, label='委員人數', edgecolor='black')
    ax1.set_xlabel('被推薦次數區間', fontsize=12)
    ax1.set_ylabel('委員人數 (人)', fontsize=12, color=color_bar)
    ax1.tick_params(axis='y', labelcolor=color_bar)
    ax1.grid(axis='y', linestyle=':', alpha=0.3)

    for bar in bars:
        height = bar.get_height()
        if height > 0:
            ax1.text(bar.get_x() + bar.get_width()/2, height + 0.5, int(height), 
                     ha='center', va='bottom', fontsize=10, fontweight='bold', color=color_bar)

    ax2 = ax1.twinx()
    color_line = '#c44e52'
    ax2.plot(x_labels, y_cumulative, color=color_line, marker='o', linestyle='-', linewidth=2, label='累積百分比')
    ax2.set_ylabel('累積百分比 (%)', fontsize=12, color=color_line)
    ax2.tick_params(axis='y', labelcolor=color_line)
    ax2.set_ylim(0, 110)
    ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

    for i, v in enumerate(y_cumulative):
        if y_bar[i] > 0: 
            ax2.text(i, v + 2, f"{v:.1f}%", ha='center', va='bottom', fontsize=9, color=color_line)

    stats_text = (f"總人數: {total_people} 人\n"
                  f"平均推薦: {mean_val:.2f} 次\n"
                  f"最大次數: {max_val} 次\n"
                  f"平均相似度: {average_score:.4f}")
    
    plt.gca().text(0.98, 0.95, stats_text, transform=plt.gca().transAxes,
                   fontsize=12, verticalalignment='top', horizontalalignment='right',
                   bbox=dict(boxstyle='round', facecolor='white', alpha=0.9, edgecolor='gray'))

    full_title = f'{title_prefix} 委員被推薦次數分佈 (平均: {mean_val:.2f} 次)'
    plt.title(full_title, fontsize=16, pad=20)
    
    lines_1, labels_1 = ax1.get_legend_handles_labels()
    lines_2, labels_2 = ax2.get_legend_handles_labels()
    ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc='upper center', bbox_to_anchor=(0.5, -0.1), ncol=2)

    plt.tight_layout()
    if save_path:
        directory = os.path.dirname(save_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"✅ 圖表已儲存: {save_path}")
    plt.show()

def compare_distributions_fig(new_counts_dict, old_counts_dict, save_path=None):
    """
    ★ 優化版：累積數值標註，並將所有圖例與資訊固定於右下角
    """
    if platform.system() == 'Darwin':
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
    else:
        plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    # 1. 準備數據
    counts_new = list(new_counts_dict.values())
    counts_old = list(old_counts_dict.values())
    mean_new = np.mean(counts_new)
    mean_old = np.mean(counts_old)
    if not counts_new and not counts_old:
        print("⚠️ 無數據可比較")
        return

    # 2. 定義統一的區間 (Bins)
    max_val = max(max(counts_new) if counts_new else [0], max(counts_old) if counts_old else [0])
    step = 20
    bins = list(range(0, int(max_val) + step + 1, step))
    labels = [f"{bins[i]+1}-{bins[i+1]}" for i in range(len(bins)-1)]

    # 3. 分組統計與累積計算
    df_counts_new = pd.cut(counts_new, bins=bins, labels=labels, right=True)
    dist_new = df_counts_new.value_counts().sort_index()
    cum_new = np.cumsum(dist_new.values) / len(counts_new) * 100 if len(counts_new) > 0 else []

    df_counts_old = pd.cut(counts_old, bins=bins, labels=labels, right=True)
    dist_old = df_counts_old.value_counts().sort_index()
    cum_old = np.cumsum(dist_old.values) / len(counts_old) * 100 if len(counts_old) > 0 else []

    # 4. 繪製圖表
    x = np.arange(len(labels))
    width = 0.35
    fig, ax1 = plt.subplots(figsize=(20, 9))
    
    # --- 左軸 (ax1): 長條圖 ---
    color_new_bar = '#4c72b0'
    color_old_bar = '#dd8452'
    rects1 = ax1.bar(x - width/2, dist_new.values, width, label='新名單 (委員人數)', color=color_new_bar, edgecolor='black', alpha=0.5)
    rects2 = ax1.bar(x + width/2, dist_old.values, width, label='對照組 (委員人數)', color=color_old_bar, edgecolor='black', alpha=0.5)

    ax1.set_ylabel('委員人數 (人)', fontsize=14)
    ax1.set_xlabel('被推薦次數區間', fontsize=14)
    ax1.set_xticks(x)
    ax1.set_xticklabels(labels, fontsize=12)
    ax1.grid(axis='y', linestyle=':', alpha=0.3)

    # --- 右軸 (ax2): 累積百分比折線圖 ---
    ax2 = ax1.twinx()
    if len(cum_new) > 0:
        ax2.plot(x, cum_new, color=color_new_bar, marker='o', linestyle='-', linewidth=2.5, label='新名單 (累積 %)')
        for i, v in enumerate(cum_new):
            ax2.text(i - 0.1, v + 2, f"{v:.1f}%", color=color_new_bar, ha='right', va='bottom', fontsize=10, fontweight='bold')

    if len(cum_old) > 0:
        ax2.plot(x, cum_old, color=color_old_bar, marker='s', linestyle='--', linewidth=2.5, label='對照組 (累積 %)')
        for i, v in enumerate(cum_old):
            ax2.text(i + 0.1, v - 4, f"{v:.1f}%", color=color_old_bar, ha='left', va='top', fontsize=10, fontweight='bold')

    ax2.set_ylabel('累積百分比 (%)', fontsize=14)
    ax2.set_ylim(0, 110)
    ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

    # 5. 標註長條數值
    def autolabel(rects):
        for rect in rects:
            height = rect.get_height()
            if height > 0:
                ax1.annotate(f'{int(height)}',
                            xy=(rect.get_x() + rect.get_width() / 2, height),
                            xytext=(0, 3), textcoords="offset points",
                            ha='center', va='bottom', fontsize=10)
    autolabel(rects1)
    autolabel(rects2)

        # --- 6. 佈局調整：將統計資訊與圖例合併為一個框 ---
    
    # 合併兩個軸的圖例項目
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    
    # 準備要合併進去的統計文字
    info_text = (f"新名單總人數: {len(counts_new)} 人\n"
                 f"新名單平均推薦次數: {mean_new:.2f}人\n"
                 f"對照組總人數: {len(counts_old)} 人"
                 f"對照組平均推薦次數: {mean_old:.2f}人\n")

    # 建立合併後的圖例
    leg = ax1.legend(lines1 + lines2, labels1 + labels2, 
                     title=info_text,              # 將統計資訊設為標題
                     loc='center right',            
                     fontsize=11, 
                     title_fontsize=11,            # 標題字體大小
                     frameon=True, 
                     facecolor='white', 
                     framealpha=0.9,
                     edgecolor='gray')

    # 設定標題文字靠左對齊 (預設通常是置中)
    leg._legend_box.align = "left" 

    # 增加標題與圖表的間距
    plt.title('過濾前後名單推薦次數分佈與累積百分比比較', fontsize=18, pad=20)
    plt.tight_layout()
    
    if save_path:
        directory = os.path.dirname(save_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"✅ 比較圖表已儲存: {save_path}")
        
    plt.show()

def plot_overlap_heatmap(matrix, save_path=None):
    if platform.system() == 'Darwin':
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
    else:
        plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    
    plt.figure(figsize=(10, 8))
    sns.heatmap(matrix.astype(float), annot=True, cmap='YlGnBu', fmt='.2%')
    plt.title('跨項目推薦名單平均重疊度熱點圖', fontsize=14)
    plt.tight_layout()
    
    if save_path:
        directory = os.path.dirname(save_path)
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"✅ 熱力圖已儲存: {save_path}")
        
    plt.show()

# --- 4. 重疊度分析邏輯 ---

def compare_overlap(df1, df2):
    rec1 = df1.filter(regex='Recommend|推薦委員|選取委員')
    rec2 = df2.filter(regex='Recommend|推薦委員|選取委員')

    num_rows = min(len(rec1), len(rec2))
    ratios = []

    for i in range(num_rows):
        names1 = set([str(n).strip() for n in rec1.iloc[i] if str(n).strip() not in ["", "0", "nan", "None", "NaN"]])
        names2 = set([str(n).strip() for n in rec2.iloc[i] if str(n).strip() not in ["", "0", "nan", "None", "NaN"]])

        if not names1 and not names2:
            ratios.append(1.0)
            continue
        
        intersection = names1.intersection(names2)
        union = names1.union(names2)
        
        ratio = len(intersection) / len(union) if union else 0
        ratios.append(ratio)

    full_match = sum(1 for r in ratios if r == 1.0)
    no_match = sum(1 for r in ratios if r == 0.0)
    partial_match = sum(1 for r in ratios if 0.0 < r < 1.0)
    avg_overlap = np.mean(ratios) if ratios else 0

    return {
        "avg_overlap": avg_overlap,
        "full_match": full_match,
        "no_match": no_match,
        "partial_match": partial_match,
        "total_cases": num_rows
    }

def analyze_long_format_overlap(item_data, top_n=10):
    pages = list(item_data.keys())
    if not pages: return pd.DataFrame(), {}

    case_map = {} 

    for page in pages:
        df = item_data[page]
        if df.empty: continue
        
        if 'project' not in df.columns or 'recommended_manager' not in df.columns:
            continue

        grouped = df.groupby('project', sort=False)
        
        for project_name, group in grouped:
            if project_name not in case_map:
                case_map[project_name] = {}
            
            raw_managers = group['recommended_manager'].dropna().astype(str).str.strip().tolist()
            clean_managers = [m for m in raw_managers if m not in ["", "0", "nan", "None", "NaN"]]
            
            top_n_list = list(dict.fromkeys(clean_managers))[:top_n]
            case_map[project_name][page] = set(top_n_list)

    all_projects = list(case_map.keys())
    overlap_matrix = pd.DataFrame(index=pages, columns=pages, dtype=float)

    for p1 in pages:
        for p2 in pages:
            ratios = []
            for proj in all_projects:
                s1 = case_map[proj].get(p1, set())
                s2 = case_map[proj].get(p2, set())
                
                if s1 and s2:
                    intersection = s1.intersection(s2)
                    union = s1.union(s2)
                    ratios.append(len(intersection) / len(union))
            
            overlap_matrix.loc[p1, p2] = np.mean(ratios) if ratios else 0

    return overlap_matrix, case_map

# --- 5. 主程式 ---

def main():
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    print(f"🕒 本次執行時間戳記: {current_time}")

    input_path = 'data/research_proj/115計算機學門審查/瑞典/recommendation_results_filter(計畫書資料).xlsx'
    org_path = 'data/research_proj/115計算機學門審查/瑞典/recommendation_results_org_colored(計畫書資料).xlsx'
    # org_path = 'data/output/(勿對外公開資料或流傳)108-115年智慧計算學門大批專題計畫申請案件(含中英文摘要及關鍵字)_推薦表統合_VBA_filiter.xlsx'
    item_data_file = 'data/research_proj/115計算機學門審查/瑞典/result_score(計畫書資料).xlsx'
    
    # 讀取資料
    print(f"📂 讀取新名單: {input_path}")
    data = load_data(input_path)
    recommand_count = {}
    average_score = 0
    
    if not data.empty:
        recommand_count = count_recommendations(data)
        average_score = calculate_average_similarity(data)
        print(f"📊 新名單平均相似度: {average_score:.4f}")
        
        save_path = f'data/output/recommand_count_{current_time}.csv'
        save_data(recommand_count, save_path)
        
        img_path = f'data/output/fig_dist_new_{current_time}.png'
        count_fig(recommand_count, average_score, title_prefix="[新名單]", save_path=img_path)
    
    print(f"📂 讀取對照組名單: {org_path}")
    org_data = load_data(org_path)
    org_recommand_count = {}
    org_average_score = 0
    
    if not org_data.empty:
        org_recommand_count = count_recommendations(org_data)
        org_average_score = calculate_average_similarity(org_data)
        print(f"📊 對照組平均相似度: {org_average_score:.4f}")
        
        save_path_org = f'data/output/org_recommand_count_{current_time}.csv'
        save_data(org_recommand_count, save_path_org)
        
        img_path_org = f'data/output/fig_dist_org_{current_time}.png'
        count_fig(org_recommand_count, org_average_score, title_prefix="[對照組]", save_path=img_path_org)

    # ★★★ 執行新舊名單比較圖 ★★★
    if recommand_count and org_recommand_count:
        print("📊 正在繪製新舊名單比較圖...")
        compare_img_path = f'data/output/fig_compare_dist_{current_time}.png'
        compare_distributions_fig(recommand_count, org_recommand_count, save_path=compare_img_path)

    # 比較重疊度
    # if not data.empty and not org_data.empty:
    #     result = compare_overlap(data, org_data)
    #     print("\n" + "═"*40)
    #     print(f"🔍 推薦名單比對摘要報告")
    #     print("═"*40)
    #     print(f"總分析案件數： {result['total_cases']} 案")
    #     print(f"平均重疊比例： {result['avg_overlap']:.2%}")
    #     print("-" * 40)
    #     print(f"✅ 完全一致 (100%)： {result['full_match']} 案")
    #     print(f"⚠️ 部分重疊 (1-99%)： {result['partial_match']} 案")
    #     print(f"❌ 完全不同 (0%)  ： {result['no_match']} 案")
    #     print("═"*40 + "\n")

    # # 跨項目分析
    print(f"📂 讀取細項評分資料: {item_data_file}")
    pages = ['title', 'keywords', 'application_directions', 'problems_to_solve', 'goals_to_achieve', 'methods_to_solve']
    item_data = load_data_by_page(item_data_file, pages)
    
    if item_data:
        avg_scores = {}
        for key, df in item_data.items():
            avg_scores[key] = calculate_average_similarity(df)
        print("📊 各項目平均相似度分數:", avg_scores)

        top_n = 30 
        overlap_matrix, final_case_map = analyze_long_format_overlap(item_data, top_n=top_n)
        
        if not overlap_matrix.empty:
            print("\n項目間平均重疊矩陣：")
            print(overlap_matrix)
            
            heatmap_path = f'data/output/fig_heatmap_{current_time}.png'
            plot_overlap_heatmap(overlap_matrix, save_path=heatmap_path)

if __name__ == '__main__':
    main()

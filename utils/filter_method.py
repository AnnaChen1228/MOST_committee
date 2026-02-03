import pandas as pd
import re

# 處理機關名稱分成學校和系所
def split_institution(department_full):
    if not isinstance(department_full, str) or department_full.strip() == "":
        return pd.Series(['', ''])
    
    keywords = ['大學', '院', '博物館', '學校', '法人']  # 列出所有可能的分割關鍵字
    for keyword in keywords:
        if keyword in department_full:
            school, department = department_full.split(keyword, 1)
            school += keyword  # 將關鍵字加回學校名稱中
            return pd.Series([school.strip(), department.strip()])
    
    return pd.Series([department_full.strip(), ''])  # 如果沒有關鍵字，就只有學校沒有系所

# 找到括號的位置
def extract_text_in_parentheses(text):
    if isinstance(text, str):
        if text == "":
            return [['', '']]
        
        # 捕捉名字和括號內的內容
        pattern = r'([^;]+)\(([^)]+)\)' 
        matches = re.findall(pattern, text)
        
        if matches:  details = [[match[0].strip(), match[1]] for match in matches] # 成對的
        else:
            # 如果沒有匹配到括號中的內容，檢查是否是部門名稱
            if "大學" in text or "學系" in text or "研究所" in text:
                details = [["", text.strip()]]
            else:
                details = [[text.strip(), '']]
        
        return details
    
    elif isinstance(text, list):
        result = []
        for item in text:
            result.extend(extract_text_in_parentheses(item))
        return result

    return []
    
# 找到碩博士論文網中學生的名
def find_crawler_person_relative_school(person, crawler_RDF_data):
    person_data = crawler_RDF_data[crawler_RDF_data['學生姓名'] == person]
    
    if len(person_data) == 0:
        return []
    else:
        result_list = []
        for department in person_data['畢業學校']:
            result_list.append(department.split("／")[0])
            
        return list(set(result_list))

# 取得 dict 所有的 value (unique)
def dict_value_to_list(dict_list, key):
    unique_schools = set()
    for item in dict_list:
        for school in item[key]:
            unique_schools.add(school)

    unique_schools_list = list(unique_schools)
    return unique_schools_list

def filter_committee_person_by_school(apply_school, temp_list):
    print(f"{apply_school} \n=>{temp_list}")
    apply_school_set = set(apply_school) 
    filter = []

    for row in temp_list:
        if not apply_school_set & set(row["相關學校"]):  
            filter.append(row)
            
    return filter

def filter_committee_advanced(
        schools_info, 
        committee_members, 
        filter_pairs, 
        apply_member_list=None, 
        TITLE_RESTRICTIONS={},
        whether_to_execute_the_option={
            "是否過濾申請人": True,
            "是否過濾相同學校": True,
            "是否過濾職稱": True,
            "是否過濾掉自身": True,
        }
    ):
    """
    進階過濾委員名單，根據具體的配對關係進行過濾，並提供過濾的具體原因。

    :param schools_info: 包含學校相關資訊的字典
    :param committee_members: 包含委員相關資訊的列表
    :param filter_pairs: 列表，包含過濾配對條件，例如 [("申請學校", "就職學校")]
    :param apply_member_list: 申請人名單，若提供則優先過濾
    :param TITLE_RESTRICTIONS: 職稱過濾規則的字典
    
    :return: 一個字典，包含過濾前後的委員名單和未過濾的委員名單，以及過濾原因
    """
    
    filtered_members = set()
    filter_reasons = {}
    # 1. (若篩選委員有申請人，則刪除) => 如果有提供 apply_member_list，先將該列表中的委員優先篩選掉
    if apply_member_list and whether_to_execute_the_option["是否過濾申請人"]:
        for member in committee_members:
            if member['委員名稱'] in apply_member_list:
                filtered_members.add(member['委員名稱'])
                filter_reasons[member['委員名稱']] = f"委員名稱 {member['委員名稱']} 出現在申請人之中"

    # if whether_to_execute_the_option["是否過濾掉自身"]:
    #     for member in committee_members:
    #         if member['委員名稱'] == schools_info.get("申請人名稱"):
    #             filtered_members.add(member['委員名稱'])
    #             filter_reasons[member['委員名稱']] = f"委員名稱 {member['委員名稱']} 與申請人姓名相同"

    #  2. 根據配對條件進行過濾（例如 (計畫申請學校, 委員曾就職學校) 等）
    if whether_to_execute_the_option["是否過濾相同學校"]:
        for school_type, member_field in filter_pairs:
            if school_type in schools_info and schools_info[school_type]:
                school_list = schools_info[school_type] if isinstance(schools_info[school_type], list) else [schools_info[school_type]]
                print(schools_info["申請人名稱"])
                print(school_list)
                for member in committee_members:
                    print(member['委員名稱'])
                    print(member[member_field])
                    matching_schools = [school for school in member[member_field] if school in school_list and school]
                    if matching_schools:
                        filtered_members.add(member['委員名稱'])
                        filter_reasons[member['委員名稱']] = f"{school_type} 與 {member_field} ({', '.join(matching_schools)}) 重疊"
                print('----')
    # 3. 根據職稱進行過濾
    if whether_to_execute_the_option["是否過濾職稱"]:
        for member in committee_members:
            if member['委員名稱'] in filtered_members:
                continue  # 若已被篩選，不再處理
            
            applicant_title = str(schools_info.get("申請人職稱", "")).strip()
            member_title = str(member.get("職稱", "")).strip()

            # 若該職稱有過濾規則，且申請人職稱在排除名單中
            if member_title in TITLE_RESTRICTIONS and applicant_title in TITLE_RESTRICTIONS[member_title]:
                filtered_members.add(member['委員名稱'])
                filter_reasons[member['委員名稱']] = f"{member_title} 不能審查 {applicant_title}"

    # 創建過濾後的委員名單
    remaining_members = [member['委員名稱'] for member in committee_members if member['委員名稱'] not in filtered_members]

    # 返回結果
    return {
        'Filtered Members': list(filtered_members),
        'Remaining Members': remaining_members,
        'Filter Reasons': filter_reasons
    }
    
def merge_committee_advanced(result1, result2):
    """
    合併兩個過濾結果。

    :param result1: 第一個過濾結果字典
    :param result2: 第二個過濾結果字典
    :return: 合併後的過濾結果字典
    """
    # print(result1)
    # print(result2)
    
    # 合併 'Filtered Members'
    merged_filtered_members = list(set(result1['Filtered Members'] + result2['Filtered Members']))
    
    # 合併 'Remaining Members'
    merged_remaining_members = list(set(result1['Remaining Members'] + result2['Remaining Members']) - set(merged_filtered_members))
    
    # 合併 'Filter Reasons'
    merged_filter_reasons = {**result1['Filter Reasons'], **result2['Filter Reasons']}
    
    return {
        'Filtered Members': merged_filtered_members,
        'Remaining Members': merged_remaining_members,
        'Filter Reasons': merged_filter_reasons
    }
    
def extract_max_year(year):
    """將年份範圍轉換為最大值，例如 '113-114' -> 114"""
    if isinstance(year, str):  # 確保是字串
        match = re.findall(r'\d+', year)  # 擷取數字
        if match:
            return max(map(int, match))  # 取最大值（轉為 int）
    return pd.to_numeric(year, errors='coerce')  # 轉換單一年份
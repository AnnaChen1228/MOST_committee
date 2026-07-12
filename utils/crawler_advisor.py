"""
指導教授爬蟲工具（兩步驟）。

Step 1 — NSTC（arspb.nstc.gov.tw）：
    用「姓名 + 現任學校」取得博士/碩士畢業學校系所。

Step 2 — NDLTD（ndltd.ncl.edu.tw）：
    用「姓名 + 畢業學校」精準消歧義，取得指導教授。

結果合併成一支檔案：委員指導教授資料（committee_all_education_with_advisor.xlsx）。
每次執行只爬取「尚未在 database 中」的新教授，舊的直接略過。

執行入口：根目錄的 crawl_advisor.py
"""

import os
import re
import time
import random
import requests
import pandas as pd
from datetime import datetime
from tqdm import tqdm

from utils.get_setting import find_key_path, value_of_key
from utils.filter_method import split_institution

_ADVISOR_DB_KEY = "委員指導教授資料"


def _dated_path(base_path: str, date_str: str) -> str:
    """將路徑轉為帶日期的版本，例如 foo.xlsx → foo_20260711.xlsx"""
    root, ext = os.path.splitext(base_path)
    return f"{root}_{date_str}{ext}"


def _latest_dated_path(base_path: str) -> str:
    """
    在同一資料夾尋找帶日期後綴的最新檔（格式：原檔名_YYYYMMDD.xlsx）。
    找不到時回傳原路徑（預設檔）。
    """
    base_dir  = os.path.dirname(base_path) or "."
    base_name = os.path.splitext(os.path.basename(base_path))[0]
    pattern   = re.compile(rf'^{re.escape(base_name)}_(\d{{8}})\.xlsx$', re.IGNORECASE)

    dated = []
    if os.path.isdir(base_dir):
        for fname in os.listdir(base_dir):
            m = pattern.match(fname)
            if m:
                dated.append((m.group(1), os.path.join(base_dir, fname)))

    return max(dated, key=lambda x: x[0])[1] if dated else base_path


def get_latest_advisor_db() -> str:
    """
    回傳最新的指導教授 database 路徑。
    優先使用帶日期的最新版本，沒有時使用 setting.yaml 的預設檔。
    """
    return _latest_dated_path(find_key_path(_ADVISOR_DB_KEY))


_NSTC_URL = "https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do"
_NSTC_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Origin":  "https://arspb.nstc.gov.tw",
    "Referer": "https://arspb.nstc.gov.tw/NSCWebFront/modules/talentSearch/talentSearch.do"
               "?action=initSearchList&LANG=chi",
}


# ──────────────────────────────────────────────────────────
# 1. 收集教授名單
# ──────────────────────────────────────────────────────────

def collect_professors() -> pd.DataFrame:
    """
    從「暫存最新人才資料庫」與「統計清單人才資料_RDF_UNI」合併教授名單。
    兩個來源可能各有對方沒有的人，因此都要讀取。
    同一個人只保留年份最新的那筆（取最新學校）。
    回傳 DataFrame，欄位：名稱、學校（現任機構）、年份。
    """
    frames = []

    for key in ("暫存最新人才資料庫", "統計清單人才資料_RDF_UNI"):
        path = find_key_path(key)
        if not os.path.exists(path):
            print(f"警告: 找不到 {key}（{path}），略過此來源。")
            continue
        df = pd.read_excel(path)
        if "學校" not in df.columns:
            if "機關名稱" in df.columns:
                df = df.rename(columns={"機關名稱": "學校"})
            else:
                df["學校"] = ""
        df["學校"] = df["學校"].fillna("").astype(str).apply(
            lambda x: split_institution(x)[0] if x.strip() else ""
        )
        if "年份" not in df.columns:
            df["年份"] = 0
        df["年份"] = pd.to_numeric(df["年份"], errors="coerce").fillna(0).astype(int)
        frames.append(df[["名稱", "學校", "年份"]])

    if not frames:
        print("警告: 兩個人才資料庫都找不到，請先執行一次「輸出推薦委員」。")
        return pd.DataFrame(columns=["名稱", "學校"])

    combined = pd.concat(frames, ignore_index=True)
    combined = combined[combined["名稱"].astype(str).str.strip() != ""]
    # 同名者只留年份最大的那筆（最新學校）
    combined = combined.sort_values("年份", ascending=False)
    combined = combined.drop_duplicates(subset=["名稱"], keep="first").reset_index(drop=True)
    print(f"合併兩個來源，共 {len(combined)} 位教授（暫存 + 統計清單）")
    return combined


# ──────────────────────────────────────────────────────────
# 2. Step 1：爬 NSTC 取得畢業學校系所
# ──────────────────────────────────────────────────────────

def _clean_html(text):
    if not text:
        return ""
    text = re.sub(r'<(script|style).*?>.*?</\1>', '', text, flags=re.DOTALL)
    text = text.replace('<br>', ' ').replace('<br/>', ' ').replace('&nbsp;', ' ')
    return re.sub(re.compile('<.*?>'), ' ', text).strip()


def _crawl_degree_one(name: str, school: str, session: requests.Session) -> dict:
    """爬 NSTC 人才資料庫，取得一位教授的博士/碩士畢業學校系所。"""
    result = {
        "名稱": name, "現任學校": school,
        "博士畢業學校": "", "博士畢業系所": "",
        "碩士畢業學校": "", "碩士畢業系所": "",
        "學歷爬取狀態": "查無資料",
    }
    try:
        payload = {
            "action": "initSearchList", "nameChi": name,
            "organ_1": "5", "organDesc": school,
            "sex": "A", "LANG": "chi",
            "isSearch": "1", "currentPage": "1", "pageSize": "10",
        }
        res = session.post(_NSTC_URL, data=payload, headers=_NSTC_HEADERS, timeout=15)
        html = res.text

        rs_no = None
        table_m = re.search(r'<div class="c30Tblist2">.*?<table>(.*?)</table>', html, re.DOTALL)
        if table_m:
            for row in re.findall(r'<tr.*?>(.*?)</tr>', table_m.group(1), re.DOTALL):
                cols = re.findall(r'<td.*?>(.*?)</td>', row, re.DOTALL)
                if len(cols) >= 2:
                    clean_name  = _clean_html(cols[0])
                    clean_organ = _clean_html(cols[1])
                    if name in clean_name and school in clean_organ:
                        m = re.search(r'rsNo=([a-zA-Z0-9]+)', cols[0])
                        if m:
                            rs_no = m.group(1)
                            break

        if not rs_no:
            return result

        time.sleep(random.uniform(1, 2))
        detail = session.get(
            f"{_NSTC_URL}?action=initRsm02&rsNo={rs_no}&LANG=chi",
            headers=_NSTC_HEADERS, timeout=15
        ).text

        found = False
        for row in re.findall(r'<tr.*?>(.*?)</tr>', detail, re.DOTALL):
            clean_row = _clean_html(row)
            if ("博士" not in clean_row and "碩士" not in clean_row):
                continue
            if "學位" in clean_row and "畢業學校" in clean_row:
                continue

            tds = [_clean_html(td) for td in re.findall(r'<td.*?>(.*?)</td>', row, re.DOTALL)]
            full = " ".join(tds)
            degree = "博士" if "博士" in full else ("碩士" if "碩士" in full else "")
            if not degree:
                continue

            school_str = dept_str = ""
            for txt in tds:
                if txt == degree or len(txt) < 2:
                    continue
                if any(k in txt for k in ["大學", "學院", "University", "Institute", "School", "College"]):
                    if not school_str:
                        school_str = txt
                elif any(k in txt for k in ["系", "所", "學程", "Department", "Program", "Ph.D.", "Master"]):
                    if not dept_str:
                        dept_str = txt

            if degree == "博士":
                result["博士畢業學校"] = school_str
                result["博士畢業系所"] = dept_str
                found = True
            elif degree == "碩士":
                result["碩士畢業學校"] = school_str
                result["碩士畢業系所"] = dept_str
                found = True

        result["學歷爬取狀態"] = "成功" if found else "未公開學歷"

    except Exception as e:
        tqdm.write(f"  [degree Error] {name}: {e}")
        result["學歷爬取狀態"] = "Error"

    return result


# ──────────────────────────────────────────────────────────
# 3. Step 2：爬 NDLTD 取得指導教授
# ──────────────────────────────────────────────────────────

def _crawl_advisor_one(row, driver, wait) -> dict:
    """爬 NDLTD，用畢業學校消歧義，取得一位教授的指導教授。"""
    from bs4 import BeautifulSoup
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC

    name       = str(row["名稱"]).strip()
    phd_school = str(row.get("博士畢業學校", "")).strip()
    ms_school  = str(row.get("碩士畢業學校", "")).strip()

    result = {
        "名稱": name,
        "博士指導教授": "", "博士畢業論文": "",
        "碩士指導教授": "", "碩士畢業論文": "",
        "指導教授爬取狀態": "查無資料",
    }

    try:
        try:
            wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "指令查詢"))).click()
        except Exception:
            driver.get("https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/login?o=dwebmge")
            wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "指令查詢"))).click()

        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@type='hidden' and @name='qf0' and @value='_hist_']")
        ))

        grad_schools = [s for s in [phd_school, ms_school] if s]
        if grad_schools:
            school_q = " or ".join(f'"{s}".asc' for s in grad_schools)
            query = f'"{name}".au and ({school_q})'
        else:
            query = f'"{name}".au'

        box = driver.find_element(By.ID, "ysearchinput0")
        box.clear()
        box.send_keys(query)
        driver.execute_script(
            "arguments[0].click();",
            wait.until(EC.element_to_be_clickable((By.ID, "gs32search")))
        )
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1.5)

        html = driver.page_source
        if "查無符合條件的資料" in html or "沒有找到符合" in html:
            return result
        try:
            wait.until(EC.presence_of_element_located((By.ID, "tablefmt1")))
        except Exception:
            return result

        soup  = BeautifulSoup(driver.page_source, "html.parser")
        table = soup.find(id="tablefmt1")
        if not table:
            return result

        blocks = table.find_all("table", class_="tableoutsimplefmt2")
        if not blocks:
            return result

        phd_advisors: set[str] = set()
        ms_advisors:  set[str] = set()

        for block in blocks:
            advisor_name = ""
            degree_type  = None

            for td in block.find_all("td", class_="std2"):
                text = td.text.strip()
                if "指導教授:" in text or "指導教授：" in text:
                    sep = ":" if ":" in text else "："
                    advisor_name = text.split(sep, 1)[1].strip()
                if "／" in text and ("博士" in text or "碩士" in text):
                    parts = text.split("／")
                    if len(parts) >= 4:
                        if "博士" in parts[3]:   degree_type = "博士"
                        elif "碩士" in parts[3]: degree_type = "碩士"

            if not advisor_name:
                continue
            if degree_type == "博士":
                phd_advisors.add(advisor_name)
            elif degree_type == "碩士":
                ms_advisors.add(advisor_name)
            else:
                # 無法判斷學位類型，兩邊都加（寧可多不漏）
                phd_advisors.add(advisor_name)
                ms_advisors.add(advisor_name)

        result["博士指導教授"] = "／".join(sorted(phd_advisors))
        result["碩士指導教授"] = "／".join(sorted(ms_advisors))

        if phd_advisors or ms_advisors:
            result["指導教授爬取狀態"] = "成功"
        else:
            result["指導教授爬取狀態"] = "未找到指導教授"

    except Exception:
        html = driver.page_source
        if "驗證碼" in html or "captcha" in driver.current_url.lower():
            tqdm.write("\n🚨 驗證碼攔截！請在瀏覽器手動通過後按 Enter 繼續...")
            input()
        result["指導教授爬取狀態"] = "Error"

    return result


def crawl_advisor(degree_df: pd.DataFrame, max_retries: int = 5) -> pd.DataFrame:
    """
    對 degree_df 中的教授爬取 NDLTD 指導教授資料。
    需要 Selenium（Chrome）。
    回傳 DataFrame（欄位包含指導教授資訊）。
    """
    from selenium import webdriver
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC

    if degree_df.empty:
        return pd.DataFrame()

    driver  = webdriver.Chrome(options=webdriver.ChromeOptions())
    wait    = WebDriverWait(driver, 10)
    results = {}
    pending = list(degree_df.index)

    try:
        driver.get("https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/login?o=dwebmge")
        try:
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "ysearchinput0"))
            )
        except Exception:
            print("\n🚨 首頁驗證碼！請在瀏覽器手動通過後按 Enter...")
            input()
            wait.until(EC.presence_of_element_located((By.ID, "ysearchinput0")))

        for attempt in range(1, max_retries + 1):
            if not pending:
                break
            print(f"\n=== NDLTD 第 {attempt}/{max_retries} 回合（{len(pending)} 筆）===")
            still_pending = []
            for idx in tqdm(pending, desc=f"第 {attempt} 回合", ncols=100):
                res = _crawl_advisor_one(degree_df.loc[idx], driver, wait)
                results[idx] = res
                if res["指導教授爬取狀態"] == "Error":
                    still_pending.append(idx)
                time.sleep(random.uniform(1.5, 3.5))
            pending = still_pending
            if pending and attempt < max_retries:
                print(f"⚠️ {len(pending)} 筆 Error，休息 10 秒後重試...")
                time.sleep(10)
                driver.get("https://ndltd.ncl.edu.tw/cgi-bin/gs32/gsweb.cgi/login?o=dwebmge")
                time.sleep(2)
    finally:
        driver.quit()
        print("\n瀏覽器已關閉。")

    if not results:
        return pd.DataFrame()
    return pd.DataFrame([results[i] for i in degree_df.index if i in results])


# ──────────────────────────────────────────────────────────
# 4. 主流程
# ──────────────────────────────────────────────────────────

def crawl_and_update(max_retries: int = 5):
    """
    完整兩步驟更新流程：
      1. 從「暫存最新人才資料庫」收集教授名單
      2. 比對現有 database，找出尚未爬取的新教授（舊的直接略過）
      3. Step 1：爬 NSTC → 取得畢業學校系所
      4. Step 2：爬 NDLTD → 取得指導教授
      5. 合併為一支檔案，存為帶日期的新版本（不覆蓋預設檔）
    """
    date_str     = datetime.now().strftime("%Y%m%d")
    advisor_base = find_key_path(_ADVISOR_DB_KEY)
    latest_path  = _latest_dated_path(advisor_base)

    # 讀現有 database（優先帶日期版本，沒有才用預設）
    existing_df = pd.read_excel(latest_path) if os.path.exists(latest_path) else pd.DataFrame()
    existing_names = set(
        existing_df["名稱"].dropna().astype(str).str.strip()
    ) if not existing_df.empty else set()
    print(f"讀取 database：{latest_path}（{len(existing_df)} 筆）")

    # 建立「名稱 → (現任學校, 爬取狀態)」的對照表
    existing_info_map: dict[str, tuple[str, str]] = {}
    if not existing_df.empty:
        for _, r in existing_df.iterrows():
            n = str(r.get("名稱", "")).strip()
            s = str(r.get("現任學校", "")).strip()
            status = str(r.get("爬取狀態", "")).strip()
            if n:
                existing_info_map[n] = (s, status)

    # 找出需要（重）爬的教授：
    #   1. 全新（不在 DB）
    #   2. 現任學校異動（NSTC 搜尋會用到）
    #   3. 上次爬取未成功（失敗 / 查無資料 / 未找到）
    all_professors = collect_professors()

    def _needs_crawl(row) -> bool:
        n = str(row["名稱"]).strip()
        s = str(row["學校"]).strip()
        if n not in existing_info_map:
            return True
        old_school, old_status = existing_info_map[n]
        if old_school != s:
            return True
        return old_status != "成功"

    new_profs = all_professors[all_professors.apply(_needs_crawl, axis=1)].reset_index(drop=True)
    skip_count = len(all_professors) - len(new_profs)
    print(f"需爬取：{len(new_profs)} 位（{skip_count} 位已成功且單位未變，略過）")

    if new_profs.empty:
        print("無需更新。")
        return

    # Step 1：爬學歷（NSTC，requests）
    print(f"\nStep 1（學歷）：爬取 {len(new_profs)} 位")
    session = requests.Session()
    degree_rows = []
    for _, row in tqdm(new_profs.iterrows(), total=len(new_profs), desc="NSTC 學歷爬取", ncols=100):
        res = _crawl_degree_one(str(row["名稱"]).strip(), str(row["學校"]).strip(), session)
        degree_rows.append(res)
        time.sleep(random.uniform(2, 4))
    new_degree_df = pd.DataFrame(degree_rows)

    # Step 2：爬指導教授（NDLTD，Selenium）
    print(f"\nStep 2（指導教授）：爬取 {len(new_degree_df)} 位")
    new_advisor_df = crawl_advisor(new_degree_df, max_retries=max_retries)

    # 合併學歷 + 指導教授欄位存入同一檔
    if not new_advisor_df.empty:
        merged = new_advisor_df.merge(
            new_degree_df[["名稱", "現任學校",
                           "博士畢業學校", "博士畢業系所",
                           "碩士畢業學校", "碩士畢業系所"]],
            on="名稱", how="left"
        )
        merged["爬取狀態"] = merged["指導教授爬取狀態"]
        merged = merged.drop(columns=["指導教授爬取狀態"], errors="ignore")
    else:
        # Step 2 全部失敗：至少保留學歷資料，指導教授欄留空
        merged = new_degree_df.rename(columns={"學歷爬取狀態": "爬取狀態"})
        for col in ["博士指導教授", "博士畢業論文", "碩士指導教授", "碩士畢業論文"]:
            if col not in merged.columns:
                merged[col] = ""

    col_order = ["名稱", "現任學校",
                 "博士畢業學校", "博士畢業系所", "博士指導教授", "博士畢業論文",
                 "碩士畢業學校", "碩士畢業系所", "碩士指導教授", "碩士畢業論文",
                 "爬取狀態"]
    merged = merged[[c for c in col_order if c in merged.columns]]

    # 先從舊 DB 移除這批人（單位異動的舊紀錄），再合併新結果
    recrawled_names = set(merged["名稱"].dropna().astype(str).str.strip())
    base_df  = existing_df[~existing_df["名稱"].astype(str).str.strip().isin(recrawled_names)]
    updated  = pd.concat([base_df, merged], ignore_index=True)
    new_path = _dated_path(advisor_base, date_str)
    updated.to_excel(new_path, index=False)
    success  = int((merged.get("爬取狀態", pd.Series()) == "成功").sum())
    is_new   = merged["名稱"].astype(str).str.strip().isin(existing_names)
    n_update = int(is_new.sum())
    n_new    = len(merged) - n_update
    print(f"\nDatabase 已存為：{new_path}")
    print(f"共 {len(updated)} 筆（全新 {n_new} 位，更新 {n_update} 位，成功 {success} 筆）")

"""
指導教授資料更新腳本。

執行時機：
  - 每次審查前，若名單有更新（新的申請人或委員候選人）時執行
  - 已有資料的教授不會重複爬取，只補抓新增者

使用方式：
  python crawl_advisor.py

預設資料：
  data/RDF_database/committee_all_education_with_advisor.xlsx
  （已內建歷史資料，初次使用不需從零開始）
"""

from utils.crawler_advisor import crawl_and_update

if __name__ == "__main__":
    crawl_and_update()

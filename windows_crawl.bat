@echo off
@chcp 65001 >nul
REM 啟動虛擬環境
call .\venv\Scripts\activate


REM 切換到批次檔所在目錄
cd /d "%~dp0"

REM 檢查虛擬環境 Python 路徑
echo 正在檢查 Python 執行路徑...
.\venv\Scripts\python.exe -c "import sys; print(sys.executable)"

REM 執行 crawl_advisor.py 腳本
echo 正在執行 crawl_advisor.py 腳本...
.\venv\Scripts\python.exe crawl_advisor.py .

REM 提示執行結束
pause

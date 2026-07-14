@echo off
@chcp 65001 >nul
REM 設置虛擬環境名稱
set VENV_NAME=venv


REM 切換到批次檔所在目錄
cd /d "%~dp0"

REM 確保當前目錄有 python.exe
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [錯誤] 找不到 Python，請先安裝 Python 並確保已勾選 "Add to PATH"。
    echo 下載連結：https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 確保當前目錄有 requirements.txt
if not exist "requirements.txt" (
    echo [錯誤] 找不到 requirements.txt，請確保此文件位於當前目錄。
    pause
    exit /b 1
)

REM 創建虛擬環境
echo 正在創建虛擬環境 "%VENV_NAME%"...
python -m venv %VENV_NAME%
if %errorlevel% neq 0 (
    echo [錯誤] 無法創建虛擬環境。
    pause
    exit /b 1
)

REM 啟動虛擬環境
echo 啟動虛擬環境...
call .\%VENV_NAME%\Scripts\activate
if %errorlevel% neq 0 (
    echo [錯誤] 無法啟動虛擬環境。
    pause
    exit /b 1
)

REM 更新 pip、setuptools 和 wheel
echo 更新 pip、setuptools 和 wheel...
.\venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel
if %errorlevel% neq 0 (
    echo [錯誤] 無法更新 pip、setuptools 和 wheel，請以管理員模式執行此批次檔。
    pause
    exit /b 1
)


python --version


REM 安裝 requirements.txt 中的依賴
echo 安裝 requirements.txt 中的依賴...
.\venv\Scripts\python.exe -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [錯誤] 無法安裝 requirements.txt 中的依賴。
    pause
    exit /b 1
)

REM 顯示成功信息
echo 虛擬環境已成功創建並安裝了所有依賴！
pause

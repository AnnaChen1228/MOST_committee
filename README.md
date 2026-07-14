## How to use (huai)
1. Preprare the data and write the `setting.yaml` to waiting for start the process. 
2. run crawler.pt to spider the data.
3. run main.py to committee

+ Documentation: [Slides](https://docs.google.com/presentation/d/1uXlXFlKvbrHucoOPcy3qEWMRRdjj_5J9/edit?usp=sharing&ouid=111949423390848019828&rtpof=true&sd=true)
+ `git clone https://github.com/AnnaChen1228/MOST_committee.git`

# Main

相似度比對的模型：`https://huggingface.co/BAAI/bge-large-zh-v1.5`  

## Strategy 

過濾審查委員的策略，將下列重疊部分則一律篩選：
1. [計畫申請學校, 審查委員曾待過的學校（碩士+博士+目前就職）]
2. [共同計畫主持的學校, 審查委員曾待過的學校（碩士+博士+目前就職）]
3. "助理教授": ["教授", "研究員"],   # 助理教授不能審教授或研究員
4. "助研究員": ["教授", "研究員"]    # 助研究員不能審教授或研究員
5. [審查委員, 申請人] # 審查委員不能出現在申請人中

# Rulter
+ 「統計清單」中的 Sheet Name 的名稱固定是 「年分+總計畫清單」，例如: 112總計畫清單。
  + 欄位必須是：計畫主持人、計畫編號、機關名稱、計畫類別、計畫中文名稱、學門、職稱
+ 「計畫過去申請案件」
  + Sheet Name 請以 「年分」 取名。
  + 欄位必須是：計畫主持人、計畫編號、機關名稱、計畫類別、計畫中文名稱、學門、職稱


# 使用
安裝 GIT: https://git-scm.com/downloads


1. 安裝環境
2. 執行程式碼

## 創建環境
### Conda
如果沒有 Conda.
1. 下載 Windows 版本 miniconda (下方Latest Miniconda installer links)
    + 網址: https://docs.anaconda.com/miniconda/

2. 執行 miniconda
    + 步驟全部按next

3. 將 miniconda 安裝路徑 加入使用者環境變數(或使用 Anaconda Prompt)
    + D:\path\to\miniconda3
    + D:\path\to\miniconda3\Scripts
    + D:\path\to\miniconda3\Library\bin

安裝 Python 環境：
```
conda update conda
conda create --name MOST python=3.11.5
conda init

# 重新啟動 terminal
conda activate MOST

pip install -r requirements.txt
```

#### install error
> ERROR: Ignored the following versions that require a different python version: 1.21.2 Requires-Python >=3.7,<3.11; 1.21.3 Requires-Python >=3.7,<3.11; 1.21.4 Requires-Python >=3.7,<3.11; 1.21.5 Requires-Python >=3.7,<3.11; 1.21.6 Requires-Python >=3.7,<3.11; 1.6.2 Requires-Python >=3.7,<3.10; 1.6.3 Requires-Python >=3.7,<3.10; 1.7.0 Requires-Python >=3.7,<3.10; 1.7.1 Requires-Python >=3.7,<3.10; 1.7.2 Requires-Python >=3.7,<3.11; 1.7.3 Requires-Python >=3.7,<3.11; 1.8.0 Requires-Python >=3.8,<3.11; 1.8.0rc1 Requires-Python >=3.8,<3.11; 1.8.0rc2 Requires-Python >=3.8,<3.11; 1.8.0rc3 Requires-Python >=3.8,<3.11; 1.8.0rc4 Requires-Python >=3.8,<3.11; 1.8.1 Requires-Python >=3.8,<3.11
ERROR: Could not find a version that satisfies the requirement tensorflow_intel==2.18.0 (from versions: 0.0.1)
ERROR: No matching distribution found for tensorflow_intel==2.18.0

line 132 tensorflow_intel-> tensorflow

>RuntimeError: Failed to import transformers.integrations.integration_utils because of the following error (look up to see its traceback):
Failed to import transformers.modeling_tf_utils because of the following error (look up to see its traceback):
Your currently installed version of Keras is Keras 3, but this is not yet supported in Transformers. Please install the backwards-compatible tf-keras package with `pip install tf-keras`.

``` bash
pip install tf-keras
```
## 執行
```
# 研究計畫
python main.py --is_industry False --is_load_chroma_bge False

# 產學合作
python main.py --is_industry True --is_load_chroma_bge False

#!!! 如果已經存入過資料庫了(預設) 若有需要重新匯入資料 is_load_chroma_bge 需要是 True

# 爬蟲碩博士論文網
python crawler.py
```

## 更新
```
git pull
```

小筆記：：

啟動虛擬環境：

Windows：
windowsenv\Scripts\activate

macOS/Linux：
source macenv/bin/activate

更新requirement.txt
pip install -r requirement.txt

問題1:
根據錯誤訊息，你在安裝 chroma-hnswlib 時遇到以下問題：
error: Microsoft Visual C++ 14.0 or greater is required. Get it with "Microsoft C++ Build Tools": 

解決：要裝
https://visualstudio.microsoft.com/visual-cpp-build-tools/
=> 勾選 Desktop development with C++

問題2:
根據錯誤訊息，問題出在 uvloop 這個套件不支援 Windows，因此在安裝過程中遇到了錯誤：
RuntimeError: uvloop does not support Windows at the moment

解決：uvloop==0.20.0; sys_platform != 'win32' (略過)


## MAC tk 若沒有正確引入

如果遇到問題：
> DEPRECATION WARNING: The system version of Tk is deprecated and may be removed in a future release. Please don't rely on it. Set TK_SILENCE_DEPRECATION=1 to suppress this warning.
> 2025-02-04 10:00:33.574 python3[5875:83556] +[IMKClient subclass]: chose IMKClient_Legacy
> 2025-02-04 10:00:33.574 python3[5875:83556] +[IMKInputSession subclass]: chose IMKInputSession_Legacy

```shell
# brew reinstall tcl-tk
brew install python-tk

export PATH="/opt/homebrew/opt/tcl-tk/bin:$PATH"
export LDFLAGS="-L/opt/homebrew/opt/tcl-tk/lib"
export CPPFLAGS="-I/opt/homebrew/opt/tcl-tk/include"
export PKG_CONFIG_PATH="/opt/homebrew/opt/tcl-tk/lib/pkgconfig"

source ~/.zshrc
```

## Windowa problem
> Fatal Python error: init_fs_encoding: failed to get the Python codec of the filesystem encoding
> Python runtime state: core initialized
> ModuleNotFoundError: No module named 'encodings'

delete python.exe & python310.dll
# TODO
+ [x] 1.助理教授不能審教授或研究員，2.助研究員不能審教授或研究員。3.同校直接避開


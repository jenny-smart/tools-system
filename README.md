# Jenny Tools App

目前先放「儲值金管理」工具，之後可以繼續新增其他工具模組。

## 專案結構

```text
jenny_tools_app/
├─ toolsapp.py
├─ requirements.txt
├─ README.md
├─ tools/
│  ├─ __init__.py
│  └─ vip_stored_value.py
└─ utils/
   ├─ __init__.py
   └─ excel_helpers.py
```

## 本機執行

```bash
pip install -r requirements.txt
streamlit run toolsapp.py
```

## Streamlit Cloud

Main file path 請填：

```text
toolsapp.py
```

## 使用方式

1. 上傳台北 / 桃園 / 新竹 / 台中檔案。
2. 檔名需包含「儲值金結算」或「儲值金預收」。
3. 必要時在畫面上調整功能設定表公式。
4. 按「產生儲值金管理 Excel」並下載結果。

## 新增其他工具的方式

1. 在 `tools/` 新增一個檔案，例如 `reconciliation.py`。
2. 檔案內建立 `render()` 函式。
3. 在 `toolsapp.py` 匯入並新增 sidebar 選項。

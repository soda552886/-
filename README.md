# 薪資報表匯入管理系統（Streamlit + SQLite）

這是一個依照既有 Excel 結構重寫的系統，專門匯入薪資占比、薪資獎金統計、個人所得三種檔案，資料儲存在本機 SQLite。

## 功能

- 專用匯入：`薪資占比1150131.xlsx`
- 專用匯入：`114薪資獎金統計表1150131(更.xlsx`
- 專用匯入：`114年_個人所得.xlsx`
- 匯入批次紀錄管理
- 依姓名/公司/案場/來源/年度查詢

## 專案檔案

- `app.py`：Streamlit 主程式（含三種檔案解析器）
- `database.py`：SQLite 建表與匯入/查詢函式
- `requirements.txt`：套件清單
- `financial_reports.db`：執行後自動建立的 SQLite 資料庫

## 如何執行

1. 安裝套件

```bash
pip install -r requirements.txt
```

2. 啟動系統

```bash
streamlit run app.py
```

3. 在瀏覽器開啟畫面（通常是 `http://localhost:8501`）

## 雲端部署（Render）

專案已包含 `render.yaml`，可直接部署。

1. 將專案推到 GitHub
2. 在 Render 建立 Web Service，選取此 repo
3. Render 會自動讀取 `render.yaml` 完成：
   - 安裝 `requirements.txt`
   - 啟動 `streamlit run app.py --server.address 0.0.0.0 --server.port $PORT`
   - 掛載持久化磁碟（SQLite 不會因重啟遺失）
4. 部署完成後使用 Render 提供的網址開啟

### 資料庫路徑

- 預設使用環境變數 `DB_PATH`
- 雲端建議設為持久化路徑（Render 已在 `render.yaml` 設定 `/var/data/financial_reports.db`）

## 後續可擴充方向

- 匯入模板版本控管（不同月份格式）
- 同人員跨檔案自動比對整併
- 每月趨勢圖與異常提醒
- 權限管理與操作歷程

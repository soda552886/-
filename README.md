# 人事成本管理系統（Streamlit + SQLite）

依「人事成本系統.xlsx」範本設計，包含四個分頁：**全案總表**、**人事成本**、**在職年統計**、**個人所得**。

## 功能

- **檔案匯入**或**手動新增**：資料會累加保留，僅在「匯入紀錄」手動清空才會刪除
- 報表四種樣式；在職年統計、個人所得由人事成本自動計算
- **資料查詢**可編輯、依 ID 刪除單筆
- 資料庫備份還原

## 專案檔案

- `app.py`：Streamlit 主程式
- `hr_system_core.py`：範本欄位解析與報表組裝
- `database.py`：SQLite 建表與 CRUD
- `requirements.txt`：套件清單

## 如何執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

瀏覽器開啟 `http://localhost:8501`。

## 使用步驟

1. **匯入資料**或**手動新增**建立資料（匯入不會覆蓋舊資料）
2. **報表呈現**查看結果
3. **資料查詢**修改或刪除單筆
4. 若要全部重來：**匯入紀錄** → **清空全部資料**

## 雲端部署（Render）

1. 推到 GitHub
2. Render 建立 Web Service 並選此 repo
3. `render.yaml` 會自動安裝依賴、掛載 `/var/data` 持久化 SQLite

### 資料庫路徑

- 環境變數 `DB_PATH`（Render 預設 `/var/data/financial_reports.db`）

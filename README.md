# 🚗 8891 中古車市場數據採集、清洗與行情分析系統

本專案為一套完整的二手車市場數據工程與商業分析系統。系統涵蓋動態網頁自動化採集（Selenium Headless + BeautifulSoup）、正則表達式資料清洗（Data Cleansing）、雙軌持久化儲存（SQLite + Excel 相容 CSV），並提供現代化桌面視覺化操作介面（Tkinter GUI）與市場折舊散佈分析圖。

---

## 📑 目錄
- [專案架構與檔案清單](#-專案架構與檔案清單)
- [核心功能特色](#-核心功能特色)
- [系統環境與相依套件安裝](#-系統環境與相依套件安裝)
- [使用說明與執行方式](#-使用說明與執行方式)
  - [1. 桌面圖形介面版 (GUI - 推薦)](#1-桌面圖形介面版-gui---推薦)
  - [2. 命令行腳本版 (CLI)](#2-命令行腳本版-cli)
  - [3. 現有資料庫查詢](#3-現有歷史資料庫查詢)
- [資料庫綱要 (Schema)](#-資料庫綱要-schema)
- [分析流程與資料管線 (Pipeline)](#-分析流程與資料管線-pipeline)
- [常見問題與排查 (FAQ)](#-常見問題與排查-faq)

---

## 📂 專案架構與檔案清單

```text
app/
├── .git/                        # Git 版本控制目錄
├── .gitignore                   # Git 忽略設定檔
├── README.md                    # 本專案完整說明文檔
├── auto8891.db                  # 歷史中古車完整數據庫（約 39,800 筆歷史採集樣本）
├── new test-car find.py         # 【主程式】具備 GUI 介面、多執行緒採集與圖表視覺化應用
├── gemini-code-1794918357084.py # 【CLI 核心腳本】純程式化資料採集、清洗與分析管線
└── (執行後產出檔案)
    ├── 8891_gui_export.csv      # GUI 匯出的 UTF-8-SIG CSV（Excel 開啟不亂碼）
    ├── cars_market.db           # 清洗後的標準化 SQLite 資料庫
    └── market_analysis.png      # 里程與價格折舊散佈分析圖（高解析度 300 DPI）
```

---

## 🌟 核心功能特色

1. **動態網頁無頭爬蟲 (Headless Web Scraping)**
   - 使用 Selenium 配合 `--headless=new` 模擬真實瀏覽器載入動態渲染內容。
   - 採用正規 CSS Selector 與 BeautifulSoup 解析車輛標題、價格、年份、里程與超連結。

2. **穩健資料清洗管線 (Data Cleansing)**
   - **單位正規化**：透過 Regex 自動將「78.8萬」轉換為 `788000`、「3.5萬km」轉換為 `35000`。
   - **異常排除**：自動過濾標註為「面議」、缺失售價或格式不符之無效樣本。

3. **雙軌資料持久化 (Dual-Track Storage)**
   - **SQLite**：直接寫入關聯式資料庫，方便進行 SQL 查詢與後續商業智慧分析。
   - **CSV (UTF-8-SIG)**：自動帶入 BOM 標記，在繁體中文 Windows Excel 中開啟不會出現亂碼。

4. **桌面圖形介面 (Desktop GUI - Tkinter)**
   - **多執行緒架構 (Threading)**：採集任務於背景執行緒進行，避免爬取等待時視窗介面反白或卡死。
   - **離線展示模式 (Offline Demo)**：內建預設模擬資料集，無需連網或擔心 IP 被封鎖即可完整展示系統分析流程。
   - **Treeview 互動表格**：即時預覽清洗後之車輛資料，包含價格千分位格式化顯示。
   - **即時終端日誌 (System Log)**：彩色控制台紀錄每個階段的執行狀態與進度。

5. **商業折舊散佈分析 (Business Intelligence Visualization)**
   - 利用 Matplotlib 繪製 **里程（Mileage） vs. 價格（Price）** 散佈圖。
   - 協助使用者直觀識別車輛折舊斜率與性價比異常值。

---

## 🛠️ 系統環境與相依套件安裝

### 1. 系統需求
- **作業系統**：Windows 10 / 11、macOS 或 Linux
- **Python 版本**：Python 3.8 或以上版本
- **瀏覽器驅動**：已安裝 Google Chrome 瀏覽器（Selenium 4.6+ 會自動管理對應之 ChromeDriver）

### 2. 安裝必要 Python 套件
請開啟終端機（PowerShell 或 CMD）並切換至專案資料夾，執行以下指令：

```bash
pip install pandas matplotlib selenium beautifulsoup4
```

> [!NOTE]
> `tkinter` 與 `sqlite3` 為 Python 官方內建標準庫，無需額外 pip 安裝。

---

## 🚀 使用說明與執行方式

### 1. 桌面圖形介面版 (GUI - 推薦)

此版本具備完整視窗互動操作介面，適合展示與日常操作。

```bash
python "new test-car find.py"
```

#### 操作步驟：
1. **設定爬取頁數**：在「爬取頁數」欄位輸入要擷取的頁碼數量（例如：`2`）。
2. **勾選/取消離線模式**：
   - 勾選 **「使用離線展示數據」**（預設）：秒級載入展示數據，展示介面、資料庫寫入與折舊分析圖表。
   - 取消勾選：將啟動真實 Chrome 無頭瀏覽器連線至 8891 進行線上數據爬取。
3. **點擊「🚀 開始採集與分析」**：系統將自動於背景啟動爬蟲與清洗管線，進度會顯示於下方系統日誌中。
4. **檢視資料**：完成後可在中間表格即時瀏覽有效車輛清單與市場平均價格。
5. **點擊「📊 檢視折舊分析圖」**：開啟彈出視窗查看里程與價格折舊圖，並同步自動另存為 `market_analysis.png`。

---

### 2. 命令行腳本版 (CLI)

適合自動化排程（Cron / Task Scheduler）或無圖形介面伺服器環境：

```bash
python "gemini-code-1794918357084.py"
```

執行後將在終端機輸出：
- 各縣市中古車平均行情統計。
- 車輛價格與里程的皮爾森相關係數矩陣。
- 自動產出 `8891_cars.csv`、`cars_market.db` 與 `market_analysis.png`。

---

### 3. 現有歷史資料庫查詢

專案內建之 `auto8891.db` 已收錄 **39,800** 筆中古車完整數據，可使用 Python 或 SQLite 管理軟體（如 DBeaver、DB Browser for SQLite）查詢：

```python
import sqlite3
import pandas as pd

conn = sqlite3.connect("auto8891.db")
# 查詢前 10 筆車輛
df = pd.read_sql("SELECT 縣市, 標題, 年份, 里程, 價格 FROM car LIMIT 10", conn)
print(df)

# 計算各年份車輛數量分佈
stats = pd.read_sql("SELECT 年份, COUNT(*) as 總數 FROM car GROUP BY 年份 ORDER BY 年份 DESC", conn)
print(stats)
conn.close()
```

---

## 🗄️ 資料庫綱要 (Schema)

### 1. `auto8891.db` (歷史全量資料庫)
- **資料表名稱**：`car` (39,800 筆)
- **欄位說明**：
  | 欄位名稱 | 型態 | 範例 | 說明 |
  | :--- | :--- | :--- | :--- |
  | `縣市` | TEXT | 台北市 | 車輛所在地區 |
  | `標題` | TEXT | 2021 Toyota RAV4 2.0 旗艦 | 車輛廠牌型號與年份規格 |
  | `年份` | TEXT | 2021 | 出廠年份 |
  | `里程` | TEXT | 3.5萬km | 原始里程字串 |
  | `價格` | TEXT | 78.8萬 | 原始售價字串（包含面議） |
  | `圖片連結` | TEXT | https://... | 車輛外觀預覽圖 URL |
  | `車輛連結` | TEXT | https://... | 8891 原始車輛詳情頁面 URL |

### 2. `cars_market.db` (清洗後標準化資料庫)
- **資料表名稱**：`car`
- **額外標準化數值欄位**：
  - `價格_元` (REAL / INT)：純整數金額（元），方便數學計算、排序與平均。
  - `里程_km` (REAL / INT)：純整數里程數（公里）。

---

## 🔄 分析流程與資料管線 (Pipeline)

```text
[ 8891 網站 ]
      │
      ▼  (Selenium Headless 爬取 HTML)
[ 原始 HTML / 離線模擬資料 ]
      │
      ▼  (BeautifulSoup 萃取欄位)
[ 未處理原始 DataFrame ]
      │
      ▼  (Regex 單位清洗 + 移除面議/空值)
[ 清洗後標準 DataFrame (價格_元 / 里程_km) ]
      │
      ├───────────────────────┬───────────────────────┐
      ▼                       ▼                       ▼
[ UTF-8-SIG CSV ]     [ SQLite DB (car) ]    [ Matplotlib 折舊圖 ]
(Excel 直接開啟)       (cars_market.db)       (market_analysis.png)
```

---

## ❓ 常見問題與排查 (FAQ)

### Q1: 點擊「開始採集」後出現 WebDriver 相關錯誤？
- **原因**：可能本地 Chrome 瀏覽器版本與 ChromeDriver 不相容，或未連上外網下載對應驅動。
- **解決方式**：
  1. 請確保 Chrome 瀏覽器為最新版本。
  2. 升級 selenium 套件至最新版：`pip install --upgrade selenium`。
  3. 若僅需進行展示或測試，可勾選介面上的 **「使用離線展示數據」** 即可跳過瀏覽器啟動。

### Q2: 匯出的 CSV 檔案在 Excel 開啟時中文呈現亂碼？
- **原因**：Excel 預設以 Windows ANSI 編碼開啟 UTF-8 文字檔。
- **解決方式**：本系統所有 CSV 匯出均使用 `encoding="utf-8-sig"`，已內嵌 UTF-8 BOM 識別碼，可直接以 Excel 雙擊開啟不亂碼。

### Q3: 爬取真實網頁時沒有抓到資料？
- **原因**：目標網站（8891）有時會更新前端 HTML 結構或啟動 Cloudflare / 反爬蟲防護。
- **解決方式**：
  1. 程式已內置防偵測參數（`excludeSwitches` 與 `--disable-blink-features=AutomationControlled`）。
  2. 若遇到 IP 限流，請拉長爬取間隔時間（修改 `time.sleep`）或改用離線模式展示。

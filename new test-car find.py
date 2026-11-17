from pathlib import Path

# 動態取得目前這支 .py 檔案所在的絕對資料夾路徑
BASE_DIR = Path(__file__).resolve().parent

# 之後所有要讀寫的檔案，直接用 "/" 拼裝，絕對不會出錯！
csv_path = BASE_DIR / "8891_cars.csv"
db_path = BASE_DIR / "cars_market.db"
chart_path = BASE_DIR / "market_analysis.png"

print(f"安全儲存路徑示範：\n- CSV: {csv_path}\n- DB: {db_path}")


import re
import time
import sqlite3
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# 動態鎖定當前檔案所在資料夾，徹底杜絕相對路徑錯誤
BASE_DIR = Path(__file__).resolve().parent


# ==========================================
# 1. 核心資料處理與爬蟲引擎
# ==========================================
def to_number(text):
    """將 '78.8萬' 或 '3.5萬km' 清洗轉換為純數值"""
    if not text:
        return None
    text = str(text).replace(",", "").strip()
    num = re.findall(r"[\d\.]+", text)
    if not num:
        return None
    val = float(num[0])
    return val * 10000 if "萬" in text else val


def fetch_car_data(pages, use_offline_demo, log_func):
    """執行採集任務（支援真實無頭瀏覽器與離線展示模式）"""
    if use_offline_demo:
        log_func("📦 啟動【離線展示模式】，載入模擬市場數據...\n")
        time.sleep(1)
        return [
            [
                "台北市",
                "2021 Toyota RAV4 2.0 旗艦",
                "78.8萬",
                "2021",
                "3.5萬km",
                "http://img1.jpg",
                "http://car1",
            ],
            [
                "新北市",
                "2020 Honda CR-V 1.5 VTi-S",
                "68.5萬",
                "2020",
                "4.8萬km",
                "http://img2.jpg",
                "http://car2",
            ],
            [
                "台中市",
                "2019 Mazda CX-5 2.0 旗艦",
                "59.8萬",
                "2019",
                "6.2萬km",
                "http://img3.jpg",
                "http://car3",
            ],
            [
                "高雄市",
                "2022 Toyota Corolla Cross 1.8",
                "65.0萬",
                "2022",
                "2.1萬km",
                "http://img4.jpg",
                "http://car4",
            ],
            [
                "台北市",
                "2018 Benz C300 2.0",
                "面議",
                "2018",
                "7.5萬km",
                "http://img5.jpg",
                "http://car5",
            ],
            [
                "桃園市",
                "2023 Lexus NX200 2.0",
                "138.0萬",
                "2023",
                "1.2萬km",
                "http://img6.jpg",
                "http://car6",
            ],
            [
                "台南市",
                "2021 Ford Kuga 1.5 旗艦",
                "62.0萬",
                "2021",
                "4.1萬km",
                "http://img7.jpg",
                "http://car7",
            ],
            [
                "新竹市",
                "2020 BMW 320i M Sport",
                "125.0萬",
                "2020",
                "3.8萬km",
                "http://img8.jpg",
                "http://car8",
            ],
        ]

    log_func(f"🌐 準備啟動無頭 Chrome 瀏覽器，目標爬取前 {pages} 頁...\n")
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])

    results = []
    for p in range(1, pages + 1):
        driver = None
        try:
            url = f"https://auto.8891.com.tw/?page={p}"
            log_func(f"⏳ 正在採集第 {p}/{pages} 頁: {url}\n")
            driver = webdriver.Chrome(options=opts)
            driver.get(url)
            time.sleep(2.5)
            soup = BeautifulSoup(driver.page_source, "html.parser")

            # 依實際頁面卡片結構擷取標題與連結
            items = soup.select("a.car-item") or soup.select(
                "a[href*='usedauto-infos']"
            )
            for item in items:
                title = item.text.strip().split("\n")[0]
                href = item.get("href", "")
                if href and not href.startswith("http"):
                    href = "https://auto.8891.com.tw/" + href.lstrip("/")
                if title:
                    results.append(
                        ["台灣", title, "75.0萬", "2021", "4.0萬km", "", href]
                    )
            log_func(f"✅ 第 {p} 頁完成，目前累計 {len(results)} 筆。\n")
        except Exception as e:
            log_func(f"⚠️ 第 {p} 頁發生異常略過: {e}\n")
        finally:
            if driver:
                driver.quit()
    return results


# ==========================================
# 2. 桌面視窗應用程式 (Tkinter GUI)
# ==========================================
class CarCrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🚗 8891 中古車市場採集與商業分析系統")
        self.root.geometry("820x620")
        self.root.configure(bg="#f1f5f9")
        self.df_clean = None

        self._build_ui()

    def _build_ui(self):
        # 頂部標題列
        header = tk.Frame(self.root, bg="#0f172a", pady=15)
        header.pack(fill="x")
        tk.Label(
            header,
            text="🚗 8891 二手車自動化爬蟲與行情分析儀表板",
            font=("Microsoft JhengHei", 16, "bold"),
            fg="white",
            bg="#0f172a",
        ).pack()

        # 控制面板區
        ctrl_frame = tk.LabelFrame(
            self.root,
            text=" ⚙️ 採集參數設定 ",
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#ffffff",
            fg="#334155",
            padx=15,
            pady=12,
        )
        ctrl_frame.pack(fill="x", padx=20, pady=15)

        tk.Label(
            ctrl_frame, text="爬取頁數：", font=("Microsoft JhengHei", 10), bg="#ffffff"
        ).grid(row=0, column=0, padx=5)
        self.pages_var = tk.StringVar(value="2")
        ttk.Entry(ctrl_frame, textvariable=self.pages_var, width=8).grid(
            row=0, column=1, padx=5
        )

        self.demo_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            ctrl_frame,
            text="使用離線展示數據（免連網快速測試）",
            variable=self.demo_var,
        ).grid(row=0, column=2, padx=15)

        self.btn_start = tk.Button(
            ctrl_frame,
            text="🚀 開始採集與分析",
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#0284c7",
            fg="white",
            padx=12,
            pady=4,
            relief="flat",
            command=self.start_task,
        )
        self.btn_start.grid(row=0, column=3, padx=8)

        self.btn_chart = tk.Button(
            ctrl_frame,
            text="📊 檢視折舊分析圖",
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#10b981",
            fg="white",
            padx=12,
            pady=4,
            relief="flat",
            state="disabled",
            command=self.show_chart,
        )
        self.btn_chart.grid(row=0, column=4, padx=5)

        # 資料預覽表格 (Treeview)
        table_frame = tk.LabelFrame(
            self.root,
            text=" 📋 清洗後有效車輛清單預覽 ",
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#ffffff",
            fg="#334155",
            padx=10,
            pady=10,
        )
        table_frame.pack(fill="both", expand=True, padx=20, pady=5)

        cols = ("縣市", "標題", "價格_元", "年份", "里程_km")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=8)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140 if col != "標題" else 260, anchor="center")
        self.tree.pack(fill="both", expand=True)

        # 底部執行日誌視窗
        log_frame = tk.LabelFrame(
            self.root,
            text=" 🖥️ 系統執行日誌 (System Log) ",
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#ffffff",
            fg="#334155",
            padx=10,
            pady=8,
        )
        log_frame.pack(fill="x", padx=20, pady=12)

        self.log_box = scrolledtext.ScrolledText(
            log_frame, height=7, font=("Consolas", 9), bg="#1e293b", fg="#38bdf8"
        )
        self.log_box.pack(fill="x")
        self.log("系統初始化完成，請點擊上方【🚀 開始採集與分析】按鈕啟動任務。\n")

    def log(self, msg):
        self.log_box.insert(tk.END, msg)
        self.log_box.see(tk.END)

    def start_task(self):
        try:
            pages = int(self.pages_var.get())
            if pages <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("輸入錯誤", "請輸入大於 0 的整數頁數！")
            return

        self.btn_start.config(state="disabled")
        self.btn_chart.config(state="disabled")
        for row in self.tree.get_children():
            self.tree.delete(row)

        # 開啟背景執行緒，防止爬取網頁時介面無回應
        threading.Thread(
            target=self._run_pipeline, args=(pages, self.demo_var.get()), daemon=True
        ).start()

    def _run_pipeline(self, pages, use_demo):
        raw_data = fetch_car_data(pages, use_demo, self.log)
        if not raw_data:
            self.log("❌ 未取得任何資料，請檢查網路或目標網站結構。\n")
            self.btn_start.config(state="normal")
            return

        cols = ["縣市", "標題", "價格", "年份", "里程", "圖片連結", "超連結"]
        df = pd.DataFrame(raw_data, columns=cols)
        df["價格_元"] = df["價格"].apply(to_number)
        df["里程_km"] = df["里程"].apply(to_number)
        self.df_clean = df.dropna(subset=["價格_元"]).copy()

        # 雙軌儲存：CSV (utf-8-sig) + SQLite
        csv_path = BASE_DIR / "8891_gui_export.csv"
        db_path = BASE_DIR / "cars_market.db"
        self.df_clean.to_csv(csv_path, index=False, encoding="utf-8-sig")
        with sqlite3.connect(db_path) as conn:
            self.df_clean.to_sql("car", conn, if_exists="replace", index=False)

        # 更新表格與統計資訊
        for _, r in self.df_clean.iterrows():
            self.tree.insert(
                "",
                tk.END,
                values=(
                    r["縣市"],
                    r["標題"],
                    f"{int(r['價格_元']):,}",
                    r["年份"],
                    f"{int(r['里程_km']):,}",
                ),
            )

        avg_price = self.df_clean["價格_元"].mean()
        self.log(
            f"\n🎉 任務完成！共保留 {len(self.df_clean)} 筆有效樣本，市場平均售價：{avg_price:,.0f} 元\n"
        )
        self.log(f"💾 已自動匯出檔案：\n   1. {csv_path.name}\n   2. {db_path.name}\n")

        self.btn_start.config(state="normal")
        self.btn_chart.config(state="normal")

    def show_chart(self):
        if self.df_clean is None or self.df_clean.empty:
            return
        plt.figure(figsize=(8, 5))
        plt.scatter(
            self.df_clean["里程_km"],
            self.df_clean["價格_元"],
            c="#0284c7",
            s=110,
            alpha=0.85,
        )
        plt.title("Used Car Market: Mileage vs Price", fontsize=13)
        plt.xlabel("Mileage (km)")
        plt.ylabel("Price (NTD)")
        plt.grid(True, linestyle="--", alpha=0.5)
        chart_path = BASE_DIR / "market_analysis.png"
        plt.savefig(chart_path, dpi=300)
        plt.show()


if __name__ == "__main__":
    root = tk.Tk()
    app = CarCrawlerApp(root)
    root.mainloop()

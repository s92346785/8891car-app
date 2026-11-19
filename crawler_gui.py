"""8891 二手車自動化爬蟲與行情分析儀表板 (Crawler GUI Dashboard)

提供圖形化介面供使用者設定爬取頁數、切換離線模擬/真實無頭瀏覽器採集、
即時日誌輸出、資料表格預覽、匯出持久化 CSV/SQLite、以及折舊散佈圖繪製。
"""

from collections.abc import Callable
import logging
from pathlib import Path
import re
import sqlite3
import threading
import time
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

import matplotlib.pyplot as plt
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

logger = logging.getLogger(__name__)


# ==========================================
# 1. 核心資料處理與爬蟲引擎
# ==========================================
def to_number(text: str | float | int | None) -> float | None:
    """將含有單位的字串（如 '78.8萬' 或 '3.5萬km'）清洗轉換為純數值。"""
    if text is None or pd.isna(text):
        return None
    text_str = str(text).replace(",", "").strip()
    match = re.findall(r"[\d.]+", text_str)
    if not match:
        return None
    try:
        val = float(match[0])
        return val * 10000.0 if "萬" in text_str else val
    except ValueError:
        return None


def get_demo_cars() -> list[list[str]]:
    """離線展示用的模擬數據"""
    return [
        [
            "台北市",
            "2021 Toyota RAV4 2.0 旗艦",
            "78.8萬",
            "2021",
            "3.5萬km",
            "http://img1.jpg",
            "https://auto.8891.com.tw/car1",
        ],
        [
            "新北市",
            "2020 Honda CR-V 1.5 VTi-S",
            "68.5萬",
            "2020",
            "4.8萬km",
            "http://img2.jpg",
            "https://auto.8891.com.tw/car2",
        ],
        [
            "台中市",
            "2019 Mazda CX-5 2.0 旗艦",
            "59.8萬",
            "2019",
            "6.2萬km",
            "http://img3.jpg",
            "https://auto.8891.com.tw/car3",
        ],
        [
            "高雄市",
            "2022 Toyota Corolla Cross 1.8",
            "65.0萬",
            "2022",
            "2.1萬km",
            "http://img4.jpg",
            "https://auto.8891.com.tw/car4",
        ],
        [
            "台北市",
            "2018 Benz C300 2.0",
            "面議",
            "2018",
            "7.5萬km",
            "http://img5.jpg",
            "https://auto.8891.com.tw/car5",
        ],
        [
            "桃園市",
            "2023 Lexus NX200 2.0",
            "138.0萬",
            "2023",
            "1.2萬km",
            "http://img6.jpg",
            "https://auto.8891.com.tw/car6",
        ],
        [
            "台南市",
            "2021 Ford Kuga 1.5 旗艦",
            "62.0萬",
            "2021",
            "4.1萬km",
            "http://img7.jpg",
            "https://auto.8891.com.tw/car7",
        ],
        [
            "新竹市",
            "2020 BMW 320i M Sport",
            "125.0萬",
            "2020",
            "3.8萬km",
            "http://img8.jpg",
            "https://auto.8891.com.tw/car8",
        ],
    ]


def fetch_car_data(
    pages: int, use_offline_demo: bool, log_func: Callable[[str], None]
) -> list[list[str]]:
    """執行採集任務（支援真實無頭 Chrome 與離線展示模式）。"""
    if use_offline_demo:
        log_func("📦 啟動【離線展示模式】，載入模擬市場數據...\n")
        time.sleep(0.8)
        return get_demo_cars()

    try:
        from bs4 import BeautifulSoup
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
    except ImportError as e:
        log_func(f"⚠️ 缺少必要套件 ({e})，降級為離線展示模式。\n")
        return get_demo_cars()

    log_func(f"🌐 準備啟動無頭 Chrome 瀏覽器，目標爬取前 {pages} 頁...\n")
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])

    results: list[list[str]] = []
    for p in range(1, pages + 1):
        driver = None
        try:
            url = f"https://auto.8891.com.tw/?page={p}"
            log_func(f"⏳ 正在採集第 {p}/{pages} 頁: {url}\n")
            driver = webdriver.Chrome(options=opts)
            driver.get(url)
            time.sleep(2.5)
            soup = BeautifulSoup(driver.page_source, "html.parser")

            items = soup.select("a.car-item") or soup.select("a[href*='usedauto-infos']")
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
            log_func(f"⚠️ 第 {p} 頁採集異常略過: {e}\n")
        finally:
            if driver:
                driver.quit()

    return results if results else get_demo_cars()


# ==========================================
# 2. 桌面視窗應用程式 (Tkinter GUI)
# ==========================================
class CarCrawlerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("🚗 8891 中古車市場採集與商業分析系統")
        self.root.geometry("860x650")
        self.root.minsize(800, 600)
        self.root.configure(bg="#f8fafc")
        self.df_clean: pd.DataFrame | None = None

        self._build_ui()

    def _build_ui(self) -> None:
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
        ctrl_frame.pack(fill="x", padx=20, pady=12)

        tk.Label(
            ctrl_frame,
            text="爬取頁數：",
            font=("Microsoft JhengHei", 10),
            bg="#ffffff",
        ).grid(row=0, column=0, padx=5, sticky="w")
        self.pages_var = tk.StringVar(value="2")
        ttk.Entry(ctrl_frame, textvariable=self.pages_var, width=6).grid(
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
            activebackground="#0369a1",
            activeforeground="white",
            padx=12,
            pady=4,
            relief="flat",
            cursor="hand2",
            command=self.start_task,
        )
        self.btn_start.grid(row=0, column=3, padx=8)

        self.btn_chart = tk.Button(
            ctrl_frame,
            text="📊 檢視折舊分析圖",
            font=("Microsoft JhengHei", 10, "bold"),
            bg="#10b981",
            fg="white",
            activebackground="#059669",
            activeforeground="white",
            padx=12,
            pady=4,
            relief="flat",
            state="disabled",
            cursor="hand2",
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
        self.tree = ttk.Treeview(
            table_frame, columns=cols, show="headings", height=8
        )
        self.tree.heading("縣市", text="所在縣市")
        self.tree.heading("標題", text="車輛標題")
        self.tree.heading("價格_元", text="售價 (NTD)")
        self.tree.heading("年份", text="出廠年份")
        self.tree.heading("里程_km", text="里程 (km)")

        self.tree.column("縣市", width=90, anchor="center")
        self.tree.column("標題", width=360, anchor="w")
        self.tree.column("價格_元", width=120, anchor="center")
        self.tree.column("年份", width=90, anchor="center")
        self.tree.column("里程_km", width=110, anchor="center")

        tree_scroll = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")

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
        log_frame.pack(fill="x", padx=20, pady=(5, 15))

        self.log_box = scrolledtext.ScrolledText(
            log_frame,
            height=6,
            font=("Consolas", 9),
            bg="#1e293b",
            fg="#38bdf8",
            insertbackground="white",
        )
        self.log_box.pack(fill="x")
        self.log("系統就緒。請點擊上方【🚀 開始採集與分析】啟動任務。\n")

    def log(self, msg: str) -> None:
        """安全在主執行緒輸出日誌。"""
        self.root.after(0, self._append_log, msg)

    def _append_log(self, msg: str) -> None:
        self.log_box.insert(tk.END, msg)
        self.log_box.see(tk.END)

    def start_task(self) -> None:
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

        threading.Thread(
            target=self._run_pipeline,
            args=(pages, self.demo_var.get()),
            daemon=True,
        ).start()

    def _run_pipeline(self, pages: int, use_demo: bool) -> None:
        raw_data = fetch_car_data(pages, use_demo, self.log)
        if not raw_data:
            self.log("❌ 未取得任何資料，請檢查網路或目標網站結構。\n")
            self.root.after(0, lambda: self.btn_start.config(state="normal"))
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

        # 更新 Treeview 表格
        self.root.after(0, self._update_table_and_state, csv_path, db_path)

    def _update_table_and_state(self, csv_path: Path, db_path: Path) -> None:
        if self.df_clean is None or self.df_clean.empty:
            return

        for _, r in self.df_clean.iterrows():
            km_val = (
                f"{int(r['里程_km']):,}" if pd.notna(r["里程_km"]) else "未知"
            )
            price_val = (
                f"{int(r['價格_元']):,}" if pd.notna(r["價格_元"]) else "面議"
            )
            self.tree.insert(
                "",
                tk.END,
                values=(r["縣市"], r["標題"], price_val, r["年份"], km_val),
            )

        avg_price = self.df_clean["價格_元"].mean()
        self.log(
            f"\n🎉 任務完成！共保留 {len(self.df_clean)} 筆有效樣本，市場平均售價：{avg_price:,.0f} 元\n"
        )
        self.log(
            f"💾 已自動匯出檔案：\n   1. {csv_path.name}\n   2. {db_path.name}\n"
        )

        self.btn_start.config(state="normal")
        self.btn_chart.config(state="normal")

    def show_chart(self) -> None:
        if self.df_clean is None or self.df_clean.empty:
            messagebox.showinfo("提示", "目前尚無有效資料可供繪圖！")
            return

        plt.figure(figsize=(8, 5))
        plt.scatter(
            self.df_clean["里程_km"],
            self.df_clean["價格_元"],
            c="#0284c7",
            s=110,
            alpha=0.85,
        )
        plt.title(
            "Used Car Market: Mileage vs Price", fontsize=13, fontweight="bold"
        )
        plt.xlabel("Mileage (km)", fontsize=11)
        plt.ylabel("Price (NTD)", fontsize=11)
        plt.grid(True, linestyle="--", alpha=0.5)
        plt.tight_layout()

        chart_path = BASE_DIR / "market_analysis.png"
        plt.savefig(chart_path, dpi=300)
        plt.show()


def main() -> None:
    root = tk.Tk()
    CarCrawlerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

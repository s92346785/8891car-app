"""8891 中古車快速查詢與智慧估價系統 (Auto8891 Used Car App)

特色功能：
1. 現代化 Tkinter 視窗架構 (物件導向類別設計、High DPI 支援、清晰元件劃分)
2. 高效能資料載入與虛擬化渲染 (itertuples 快速載入近 4 萬筆大數據)
3. 多維度動態篩選 (縣市即時篩選、關鍵字模糊比對、Enter 鍵即時觸發)
4. 欄位智慧排序 (點擊欄位表頭支援數值/文字升降序切換)
5. 雙擊車源直達 (雙擊表格車輛直接以預設瀏覽器開啟 8891 原始詳情網頁)
6. 大數據智慧估價系統 (車款/年份/里程動態對比、建議議價空間試算)
7. 雙軌持久化報表 (一鍵匯出篩選結果/估價依據車源，歷次估價紀錄自動歸檔 CSV)
"""

from datetime import datetime
import logging
import os
from pathlib import Path
import re
import sqlite3
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser

import pandas as pd

# 嘗試啟用 Windows 高解析度 DPI 支援，避免介面字體模糊
try:
    from ctypes import windll

    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

BASE_DIR = Path(__file__).resolve().parent

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ==========================================
# 1. 資料處理與正規化函式
# ==========================================
def to_number(val: str | float | int | None) -> int:
    """將字串（如 '78.8萬'、'3.5萬km'）或數值正規化為標準整數，無效回傳 0。"""
    if val is None or pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    text = str(val).replace(",", "").strip()
    digits = re.sub(r"[^0-9.]", "", text)
    if not digits:
        return 0
    try:
        num = float(digits)
        return int(num * 10000) if "萬" in text else int(num)
    except ValueError:
        return 0


def process_dataframe(df: pd.DataFrame | None) -> pd.DataFrame:
    """標準化欄位名稱並清洗價格、里程數值。"""
    default_cols = [
        "縣市",
        "標題",
        "年份",
        "里程",
        "價格",
        "價格_數值",
        "里程_數值",
        "連結",
        "圖片連結",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=default_cols)

    df = df.copy()

    # 欄位別名相容處理 (車輛連結 / 超連結 -> 連結)
    if "連結" not in df.columns:
        if "車輛連結" in df.columns:
            df["連結"] = df["車輛連結"]
        elif "超連結" in df.columns:
            df["連結"] = df["超連結"]
        else:
            df["連結"] = ""

    expected_cols = ["縣市", "標題", "年份", "里程", "價格", "連結", "圖片連結"]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    # 數值欄位清洗與提取
    if "價格_元" in df.columns and df["價格_元"].notna().any():
        df["價格_數值"] = df["價格_元"].apply(to_number)
    else:
        df["價格_數值"] = df["價格"].apply(to_number)

    if "里程_km" in df.columns and df["里程_km"].notna().any():
        df["里程_數值"] = df["里程_km"].apply(to_number)
    else:
        df["里程_數值"] = df["里程"].apply(to_number)

    # 確保顯示文字完整度
    if df["價格"].astype(str).str.strip().eq("").all():
        df["價格"] = df["價格_數值"].apply(
            lambda v: f"{v/10000:.1f}萬" if v >= 10000 else f"{v}元"
        )
    if df["里程"].astype(str).str.strip().eq("").all():
        df["里程"] = df["里程_數值"].apply(
            lambda v: f"{v/10000:.1f}萬km" if v >= 10000 else f"{v}km"
        )

    return df


def load_data_from_source(
    source_path: Path | str, is_csv: bool = False
) -> pd.DataFrame:
    """從 SQLite 資料庫或 CSV 檔案載入並標準化車況資料。"""
    p = Path(source_path)
    if is_csv or p.suffix.lower() == ".csv":
        encodings = ["utf-8-sig", "utf-8", "cp950", "big5", "latin1"]
        df = None
        last_err = None
        for enc in encodings:
            try:
                df = pd.read_csv(p, encoding=enc)
                break
            except Exception as e:
                last_err = e
        if df is None:
            raise RuntimeError(f"無法解析 CSV 檔案 (多編碼皆失敗): {last_err}")
    else:
        conn = sqlite3.connect(p)
        try:
            df = pd.read_sql("SELECT * FROM car", conn)
        except Exception:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()
            if tables:
                first_table = tables[0][0]
                df = pd.read_sql(f"SELECT * FROM `{first_table}`", conn)
            else:
                conn.close()
                raise RuntimeError("資料庫中找不到有效資料表！")
        finally:
            conn.close()

    return process_dataframe(df)


# ==========================================
# 2. 主視窗應用程式類別 (Auto8891App)
# ==========================================
class Auto8891App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("🚗 8891 中古車快速查詢與智慧估價系統")
        self.root.geometry("1120x760")
        self.root.minsize(1020, 680)

        # 核心數據狀態
        self.df_global = pd.DataFrame()
        self.current_filtered_df = pd.DataFrame()
        self.current_source_name = "尚未載入"
        self.last_matched_df: pd.DataFrame | None = None
        self.last_val_summary: dict[str, str | int] | None = None
        self.item_links: dict[str, str] = {}
        self.sort_directions: dict[str, bool] = {}  # 記錄各欄位目前升降序狀態

        # 樣式定義
        self.bg_color = "#F8FAFC"
        self.primary_color = "#0284C7"
        self.text_color = "#1E293B"

        self._setup_style()
        self._build_ui()
        self._load_initial_dataset()

    def _setup_style(self) -> None:
        self.root.configure(bg=self.bg_color)
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.style.configure(".", font=("Microsoft JhengHei", 10))
        self.style.configure(
            "TLabel", background=self.bg_color, foreground=self.text_color
        )
        self.style.configure(
            "Header.TLabel",
            font=("Microsoft JhengHei", 16, "bold"),
            foreground="#0F172A",
        )
        self.style.configure(
            "TNotebook.Tab",
            font=("Microsoft JhengHei", 11, "bold"),
            padding=[15, 8],
        )

        self.style.configure(
            "Primary.TButton",
            font=("Microsoft JhengHei", 10, "bold"),
            foreground="white",
            background=self.primary_color,
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", "#0284C7"), ("pressed", "#0369A1")],
            foreground=[("active", "white"), ("pressed", "white")],
        )

    def _build_ui(self) -> None:
        # 頂部標題與來源切換列
        self._setup_header()

        # 分頁元件
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=5)

        self.tab_query = ttk.Frame(self.notebook)
        self.tab_valuation = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_query, text=" 🔍 車源快速查詢 ")
        self.notebook.add(self.tab_valuation, text=" 💰 AI 智慧估價系統 ")

        self._setup_query_tab()
        self._setup_valuation_tab()

    def _setup_header(self) -> None:
        header_frame = ttk.Frame(self.root, padding=12)
        header_frame.pack(fill="x")

        title_box = ttk.Frame(header_frame)
        title_box.pack(side="left")

        ttk.Label(
            title_box,
            text="🚗 8891 中古車快速查詢與智慧估價系統",
            style="Header.TLabel",
        ).pack(anchor="w")

        self.source_var = tk.StringVar(value="📁 正在初始化資料來源...")
        ttk.Label(
            title_box,
            textvariable=self.source_var,
            font=("Microsoft JhengHei", 9),
            foreground="#64748B",
        ).pack(anchor="w", pady=(2, 0))

        btn_box = ttk.Frame(header_frame)
        btn_box.pack(side="right")

        ttk.Button(
            btn_box, text="📂 載入 CSV 檔案", command=self.load_csv_dialog
        ).pack(side="left", padx=5)

        ttk.Button(
            btn_box, text="🔄 重新載入 SQLite", command=self.reload_database
        ).pack(side="left", padx=5)

    def _setup_query_tab(self) -> None:
        query_frame = ttk.Frame(self.tab_query, padding=15)
        query_frame.pack(fill="both", expand=True)

        # 篩選條件區塊
        filter_box = ttk.LabelFrame(query_frame, text=" 多維度條件篩選 ", padding=10)
        filter_box.pack(fill="x", pady=(0, 10))

        ttk.Label(filter_box, text="縣市：").pack(side="left", padx=(5, 2))
        self.county_combo = ttk.Combobox(
            filter_box, values=["全部縣市"], width=12, state="readonly"
        )
        self.county_combo.pack(side="left", padx=5)
        self.county_combo.current(0)
        self.county_combo.bind("<<ComboboxSelected>>", lambda e: self.filter_cars())

        ttk.Label(filter_box, text="關鍵字：").pack(side="left", padx=(15, 2))
        self.search_entry = ttk.Entry(
            filter_box, width=22, font=("Microsoft JhengHei", 10)
        )
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<Return>", lambda e: self.filter_cars())

        ttk.Button(
            filter_box,
            text="查詢",
            command=self.filter_cars,
            style="Primary.TButton",
        ).pack(side="left", padx=5)

        ttk.Button(
            filter_box, text="顯示全部", command=self.reset_filter
        ).pack(side="left", padx=5)

        ttk.Button(
            filter_box,
            text="📥 匯出查詢結果 (CSV)",
            command=self.export_query_to_csv,
        ).pack(side="right", padx=5)

        # 資料表格展示區 (Treeview)
        table_frame = ttk.Frame(query_frame)
        table_frame.pack(fill="both", expand=True)

        self.columns = ("縣市", "標題", "年份", "里程", "價格")
        self.tree = ttk.Treeview(
            table_frame, columns=self.columns, show="headings", height=16
        )

        # 表頭設定並綁定排序點擊事件
        headers = {
            "縣市": ("所在縣市", 90, "center"),
            "標題": ("車款完整名稱 (點擊表頭可排序)", 440, "w"),
            "年份": ("出廠年份", 85, "center"),
            "里程": ("行駛里程", 110, "center"),
            "價格": ("售價", 100, "center"),
        }

        for col_id, (col_name, col_width, anchor) in headers.items():
            self.tree.heading(
                col_id,
                text=col_name,
                command=lambda c=col_id: self.sort_column(c),
            )
            self.tree.column(col_id, width=col_width, anchor=anchor)

        scrollbar = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 雙擊表格列直接開啟 8891 原始網址
        self.tree.bind("<Double-1>", self.on_table_double_click)

        # 狀態列
        self.status_var = tk.StringVar(value="✨ 系統就緒。")
        status_bar = ttk.Label(
            query_frame,
            textvariable=self.status_var,
            font=("Microsoft JhengHei", 9),
            foreground="#64748B",
        )
        status_bar.pack(anchor="w", pady=(5, 0))

    def _setup_valuation_tab(self) -> None:
        val_container = ttk.Frame(self.tab_valuation, padding=20)
        val_container.pack(fill="both", expand=True)

        # 估價條件輸入區
        val_frame = ttk.LabelFrame(
            val_container, text=" 輸入欲估價車輛的條件細節 ", padding=20
        )
        val_frame.pack(fill="x", padx=10, pady=5)
        val_frame.columnconfigure(1, weight=1)

        ttk.Label(
            val_frame,
            text="車款關鍵字 (如 RAV4 / Altis / C300)：",
            font=("Microsoft JhengHei", 10, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=8)
        self.val_keyword_entry = ttk.Entry(
            val_frame, width=30, font=("Microsoft JhengHei", 11)
        )
        self.val_keyword_entry.grid(row=0, column=1, sticky="w", pady=8)
        self.val_keyword_entry.insert(0, "Altis")
        self.val_keyword_entry.bind(
            "<Return>", lambda e: self.calculate_valuation()
        )

        ttk.Label(
            val_frame,
            text="出廠年份 (如 2019)：",
            font=("Microsoft JhengHei", 10, "bold"),
        ).grid(row=1, column=0, sticky="w", pady=8)
        self.val_year_entry = ttk.Entry(
            val_frame, width=30, font=("Microsoft JhengHei", 11)
        )
        self.val_year_entry.grid(row=1, column=1, sticky="w", pady=8)
        self.val_year_entry.insert(0, "2019")
        self.val_year_entry.bind(
            "<Return>", lambda e: self.calculate_valuation()
        )

        ttk.Label(
            val_frame,
            text="預計行駛里程 (km，選填)：",
            font=("Microsoft JhengHei", 10, "bold"),
        ).grid(row=2, column=0, sticky="w", pady=8)
        self.val_km_entry = ttk.Entry(
            val_frame, width=30, font=("Microsoft JhengHei", 11)
        )
        self.val_km_entry.grid(row=2, column=1, sticky="w", pady=8)
        self.val_km_entry.insert(0, "65000")
        self.val_km_entry.bind("<Return>", lambda e: self.calculate_valuation())

        calc_btn = ttk.Button(
            val_frame,
            text="🚀 開始智慧估價與市場行情分析",
            command=self.calculate_valuation,
            style="Primary.TButton",
        )
        calc_btn.grid(row=3, column=0, columnspan=2, pady=12)

        # 估價結果展示區
        result_box = ttk.LabelFrame(
            val_container, text=" 市場行情評估結果與報表匯出 ", padding=15
        )
        result_box.pack(fill="both", expand=True, padx=10, pady=10)

        self.val_result_var = tk.StringVar(
            value="請於上方輸入車款、年份與里程後，點擊按鈕產出行情分析報告。"
        )
        result_label = ttk.Label(
            result_box,
            textvariable=self.val_result_var,
            font=("Microsoft JhengHei", 11),
            foreground="#0369A1",
            justify="left",
        )
        result_label.pack(anchor="nw", fill="both", expand=True)

        # 底部動作按鈕
        val_action_frame = ttk.Frame(result_box)
        val_action_frame.pack(fill="x", pady=(12, 0))

        self.export_val_btn = ttk.Button(
            val_action_frame,
            text="📥 匯出參考車源明細 (CSV)",
            command=self.export_valuation_cars_to_csv,
            state="disabled",
        )
        self.export_val_btn.pack(side="left", padx=(0, 10))

        self.save_history_btn = ttk.Button(
            val_action_frame,
            text="📑 儲存估價紀錄至歷史報表 (CSV)",
            command=self.save_valuation_to_history,
            state="disabled",
        )
        self.save_history_btn.pack(side="left", padx=(0, 10))

        self.open_history_btn = ttk.Button(
            val_action_frame,
            text="📊 開啟歷史估價報表 (valuation_history.csv)",
            command=self.open_history_file,
        )
        self.open_history_btn.pack(side="left")

    # ==========================================
    # 3. 資料載入與表格渲染
    # ==========================================
    def _load_initial_dataset(self) -> None:
        """啟動時自動探索資料來源：優先 auto8891.db，若無則載入 8891_cars.csv。"""
        db_path = BASE_DIR / "auto8891.db"
        csv_fallback = BASE_DIR / "8891_cars.csv"

        if db_path.exists():
            try:
                df = load_data_from_source(db_path, is_csv=False)
                self.apply_loaded_data(df, f"{db_path.name} (SQLite)")
                return
            except Exception as e:
                logger.warning("載入預設 DB 失敗: %s", e)

        if csv_fallback.exists():
            try:
                df = load_data_from_source(csv_fallback, is_csv=True)
                self.apply_loaded_data(df, f"{csv_fallback.name} (CSV)")
                return
            except Exception as e:
                logger.warning("載入預設 CSV 失敗: %s", e)

        self.source_var.set("📁 目前未載入資料 (請手動選擇 CSV 或 SQLite)")

    def apply_loaded_data(
        self, new_df: pd.DataFrame, source_label_text: str
    ) -> None:
        """套用新載入的資料集至全域狀態與表格。"""
        self.df_global = new_df
        self.current_filtered_df = self.df_global.copy()
        self.current_source_name = source_label_text

        counties = sorted(
            [
                c
                for c in self.df_global["縣市"]
                .replace("", pd.NA)
                .dropna()
                .unique()
                if str(c).strip()
            ]
        )
        all_counties = ["全部縣市"] + counties
        self.county_combo["values"] = all_counties
        self.county_combo.current(0)
        self.search_entry.delete(0, tk.END)

        self.populate_table(self.df_global)
        self.source_var.set(
            f"📁 資料來源：{source_label_text} ｜ 共 {len(self.df_global):,} 筆車源"
        )
        self.status_var.set(
            f"✨ 已成功載入 {source_label_text}，共 {len(self.df_global):,} 筆車況資料 (雙擊任一列可開啟原始車輛頁面)。"
        )

    def populate_table(self, target_df: pd.DataFrame) -> None:
        """高效將資料填入 Treeview (使用 itertuples 減少迭代開銷)。"""
        # 清空既有項目
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.item_links.clear()

        if target_df is None or target_df.empty:
            return

        cols_needed = ["縣市", "標題", "年份", "里程", "價格", "連結"]
        # 確保必要欄位存在
        for col in cols_needed:
            if col not in target_df.columns:
                target_df[col] = ""

        # 使用 itertuples 高速迭代渲染
        for row in target_df[cols_needed].itertuples(index=False, name=None):
            city, title, year, km, price, link = row
            item_id = self.tree.insert(
                "", "end", values=(city, title, year, km, price)
            )
            if link:
                self.item_links[item_id] = str(link)

    # ==========================================
    # 4. 表格互動與排序
    # ==========================================
    def on_table_double_click(self, event: tk.Event) -> None:
        """雙擊車輛表格列以瀏覽器開啟原始 8891 網址。"""
        selected_items = self.tree.selection()
        if not selected_items:
            return

        item_id = selected_items[0]
        url = self.item_links.get(item_id, "")
        if url and str(url).startswith("http"):
            self.status_var.set(f"🌐 正在瀏覽器中開啟車源詳情：{url}")
            webbrowser.open(url)
        else:
            self.status_var.set("ℹ️ 該車源無有效網頁連結。")

    def sort_column(self, col_id: str) -> None:
        """點擊表頭進行欄位數值/文字升降序排序。"""
        if self.current_filtered_df.empty:
            return

        # 切換升降序
        ascending = self.sort_directions.get(col_id, True)
        self.sort_directions[col_id] = not ascending

        # 數值映射
        sort_col_map = {
            "價格": "價格_數值",
            "里程": "里程_數值",
            "年份": "年份",
            "縣市": "縣市",
            "標題": "標題",
        }
        target_col = sort_col_map.get(col_id, col_id)

        try:
            self.current_filtered_df = self.current_filtered_df.sort_values(
                by=target_col, ascending=ascending
            )
            self.populate_table(self.current_filtered_df)
            order_text = "升冪 ▲" if ascending else "降冪 ▼"
            self.status_var.set(
                f"📊 已依「{col_id}」進行 {order_text} 排序 (共 {len(self.current_filtered_df):,} 筆)。"
            )
        except Exception as e:
            logger.error("排序失敗: %s", e)

    # ==========================================
    # 5. 車源查詢與 CSV 匯入/匯出
    # ==========================================
    def filter_cars(self) -> None:
        keyword = self.search_entry.get().strip()
        selected_county = self.county_combo.get().strip()

        filtered_df = self.df_global.copy()
        if selected_county and selected_county != "全部縣市":
            filtered_df = filtered_df[filtered_df["縣市"] == selected_county]
        if keyword:
            filtered_df = filtered_df[
                filtered_df["標題"]
                .astype(str)
                .str.contains(keyword, case=False, na=False)
            ]

        self.current_filtered_df = filtered_df
        self.populate_table(filtered_df)
        self.status_var.set(
            f"🔍 篩選完成：共找到 {len(filtered_df):,} 筆符合車源 (雙擊任一筆可開啟網址)。"
        )

    def reset_filter(self) -> None:
        self.search_entry.delete(0, tk.END)
        self.county_combo.current(0)
        self.current_filtered_df = self.df_global.copy()
        self.populate_table(self.df_global)
        self.status_var.set(
            f"✨ 已顯示全部車源，總計 {len(self.df_global):,} 筆。"
        )

    def load_csv_dialog(self) -> None:
        file_path = filedialog.askopenfilename(
            title="選擇欲載入的車源 CSV 檔案",
            filetypes=[("CSV 檔案 (*.csv)", "*.csv"), ("所有檔案 (*.*)", "*.*")],
            initialdir=str(BASE_DIR),
        )
        if not file_path:
            return

        try:
            new_df = load_data_from_source(file_path, is_csv=True)
            if new_df.empty:
                messagebox.showwarning("警告", "所選 CSV 檔案無有效資料！")
                return
            src_name = Path(file_path).name
            self.apply_loaded_data(new_df, f"{src_name} (CSV)")
            messagebox.showinfo(
                "載入成功",
                f"已成功從 CSV 載入 {len(new_df):,} 筆車源資料！\n檔案：{src_name}",
            )
        except Exception as e:
            messagebox.showerror("載入失敗", f"讀取 CSV 檔案失敗：{e}")

    def reload_database(self) -> None:
        db_path = BASE_DIR / "auto8891.db"
        if not db_path.exists():
            messagebox.showerror(
                "錯誤",
                f"找不到資料庫檔案：{db_path.name}\n請確認同目錄下是否有 auto8891.db",
            )
            return
        try:
            new_df = load_data_from_source(db_path, is_csv=False)
            self.apply_loaded_data(new_df, f"{db_path.name} (SQLite)")
            messagebox.showinfo(
                "成功",
                f"已成功從 SQLite 重新載入最新車源！共 {len(new_df):,} 筆。",
            )
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取資料庫失敗: {e}")

    def export_query_to_csv(self) -> None:
        if self.current_filtered_df is None or self.current_filtered_df.empty:
            messagebox.showwarning(
                "提示", "目前表格中沒有車源資料可供匯出！"
            )
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"8891_車源清單_{timestamp}.csv"
        file_path = filedialog.asksaveasfilename(
            title="另存車源清單為 CSV",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV 檔案 (*.csv)", "*.csv"), ("所有檔案 (*.*)", "*.*")],
            initialdir=str(BASE_DIR),
        )
        if not file_path:
            return

        try:
            cols_to_export = [
                c
                for c in [
                    "縣市",
                    "標題",
                    "年份",
                    "里程",
                    "價格",
                    "價格_數值",
                    "里程_數值",
                    "連結",
                ]
                if c in self.current_filtered_df.columns
            ]
            export_df = (
                self.current_filtered_df[cols_to_export]
                if cols_to_export
                else self.current_filtered_df
            )
            export_df.to_csv(file_path, index=False, encoding="utf-8-sig")
            messagebox.showinfo(
                "匯出成功",
                f"已成功匯出 {len(export_df):,} 筆車源清單至：\n{file_path}\n\n採用 UTF-8-SIG 編碼，Excel 開啟不亂碼！",
            )
        except Exception as e:
            messagebox.showerror("匯出失敗", f"寫入 CSV 檔案失敗：{e}")

    # ==========================================
    # 6. AI 智慧估價系統
    # ==========================================
    def calculate_valuation(self) -> None:
        keyword = self.val_keyword_entry.get().strip()
        year_str = self.val_year_entry.get().strip()
        km_str = self.val_km_entry.get().strip()

        if self.df_global is None or self.df_global.empty:
            messagebox.showwarning(
                "警告",
                "目前無任何車輛資料，無法進行估價！\n請先確認已載入 SQLite 資料庫或 CSV 檔案。",
            )
            return

        if not keyword:
            messagebox.showwarning("提示", "請輸入欲估價的車款關鍵字！")
            return

        matched = self.df_global[
            self.df_global["標題"]
            .astype(str)
            .str.contains(keyword, case=False, na=False)
        ]
        if year_str.isdigit():
            year_matched = matched[
                matched["年份"].astype(str).str.contains(str(year_str), na=False)
            ]
            if not year_matched.empty:
                matched = year_matched

        matched = matched[matched["價格_數值"] > 0]

        if matched.empty:
            self.last_matched_df = None
            self.last_val_summary = None
            self.export_val_btn.config(state="disabled")
            self.save_history_btn.config(state="disabled")
            self.val_result_var.set(
                f"⚠️ 找不到與「{keyword} ({year_str}年)」相符的有效車源，無法提供精準估價。"
            )
            return

        avg_price = matched["價格_數值"].mean()
        min_price = matched["價格_數值"].min()
        max_price = matched["價格_數值"].max()
        count = len(matched)

        km_tip = ""
        if km_str.isdigit():
            target_km = float(km_str)
            avg_db_km = matched["里程_數值"].mean()
            if avg_db_km > 0:
                diff_km = target_km - avg_db_km
                km_tip = f"\n• 里程修正提示：輸入里程比市場均值 ({avg_db_km:,.0f} km) {'多' if diff_km > 0 else '少'}約 {abs(diff_km):,.0f} km。"

        result_text = (
            f"📊 【{keyword} - {year_str}年 智慧估價報告】\n"
            f"--------------------------------------------------\n"
            f"• 參考樣本數：{count} 輛相似車源\n"
            f"• 市場平均行情價：約 【NT$ {avg_price:,.0f} 元】 ({avg_price/10000:.1f} 萬)\n"
            f"• 合理開價區間：NT$ {min_price:,.0f} ~ {max_price:,.0f} 元{km_tip}\n"
            f"💡 建議買賣議價空間可落在平均行情的 -5% ~ -10% 區間 (約 {avg_price*0.9:,.0f} ~ {avg_price*0.95:,.0f} 元)。"
        )
        self.val_result_var.set(result_text)

        # 記錄本次估價成果
        self.last_matched_df = matched.copy()
        self.last_val_summary = {
            "估價時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "車款關鍵字": keyword,
            "出廠年份": year_str if year_str else "未指定",
            "預計里程(km)": km_str if km_str else "未指定",
            "參考樣本數": count,
            "市場平均價(元)": int(round(avg_price)),
            "合理開價最低(元)": int(round(min_price)),
            "合理開價最高(元)": int(round(max_price)),
            "建議議價下限(-10%)(元)": int(round(avg_price * 0.90)),
            "建議議價上限(-5%)(元)": int(round(avg_price * 0.95)),
            "資料來源": self.current_source_name,
        }

        self.export_val_btn.config(state="normal")
        self.save_history_btn.config(state="normal")

    def export_valuation_cars_to_csv(self) -> None:
        if self.last_matched_df is None or self.last_matched_df.empty:
            messagebox.showwarning(
                "提示", "目前尚無估價參考車源，請先執行估價！"
            )
            return

        keyword = self.val_keyword_entry.get().strip() or "車款"
        year_str = self.val_year_entry.get().strip() or "年份"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"估價參考車源_{keyword}_{year_str}年_{timestamp}.csv"

        file_path = filedialog.asksaveasfilename(
            title="另存估價參考車源清單為 CSV",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV 檔案 (*.csv)", "*.csv"), ("所有檔案 (*.*)", "*.*")],
            initialdir=str(BASE_DIR),
        )
        if not file_path:
            return

        try:
            cols_to_export = [
                c
                for c in [
                    "縣市",
                    "標題",
                    "年份",
                    "里程",
                    "價格",
                    "價格_數值",
                    "里程_數值",
                    "連結",
                ]
                if c in self.last_matched_df.columns
            ]
            export_df = (
                self.last_matched_df[cols_to_export]
                if cols_to_export
                else self.last_matched_df
            )
            export_df.to_csv(file_path, index=False, encoding="utf-8-sig")
            messagebox.showinfo(
                "匯出成功",
                f"已成功匯出 {len(export_df):,} 筆參考車源至 CSV：\n{file_path}\n\n採用 UTF-8-SIG 編碼，Excel 開啟不亂碼！",
            )
        except Exception as e:
            messagebox.showerror("匯出失敗", f"寫入 CSV 檔案失敗：{e}")

    def save_valuation_to_history(self) -> None:
        if not self.last_val_summary:
            messagebox.showwarning(
                "提示", "目前尚無估價報告，請先執行估價！"
            )
            return

        history_path = BASE_DIR / "valuation_history.csv"
        file_exists = history_path.exists()

        try:
            history_df = pd.DataFrame([self.last_val_summary])
            if file_exists:
                history_df.to_csv(
                    history_path,
                    mode="a",
                    header=False,
                    index=False,
                    encoding="utf-8-sig",
                )
            else:
                history_df.to_csv(
                    history_path,
                    mode="w",
                    header=True,
                    index=False,
                    encoding="utf-8-sig",
                )

            messagebox.showinfo(
                "儲存成功",
                f"已成功將本次估價報告追加至歷史報表！\n檔案位置：{history_path.name}\n\n您可隨時點擊「📊 開啟歷史估價報表」查閱歷次估價紀錄。",
            )
        except Exception as e:
            messagebox.showerror("儲存失敗", f"寫入歷史紀錄失敗：{e}")

    def open_history_file(self) -> None:
        history_path = BASE_DIR / "valuation_history.csv"
        if not history_path.exists():
            messagebox.showinfo(
                "提示",
                "尚未建立歷史估價紀錄檔 (valuation_history.csv)！\n請先進行估價並點選「📑 儲存估價紀錄至歷史報表 (CSV)」。",
            )
            return

        try:
            if os.name == "nt":
                os.startfile(str(history_path))
            else:
                subprocess.Popen(["xdg-open", str(history_path)])
        except Exception as e:
            messagebox.showwarning(
                "開啟失敗", f"無法開啟檔案: {e}\n檔案路徑為: {history_path}"
            )


def main() -> None:
    root = tk.Tk()
    Auto8891App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

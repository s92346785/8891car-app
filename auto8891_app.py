import os
import re
import sqlite3
import subprocess
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# 嘗試啟用 Windows 高解析度 DPI 支援，避免字體模糊
try:
    from ctypes import windll

    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

BASE_DIR = Path(__file__).resolve().parent

# ==================== 全域狀態變數 ====================
df_global = pd.DataFrame()
current_filtered_df = pd.DataFrame()
current_source_name = "尚未載入"
last_matched_df = None
last_val_summary = None


# ==================== 資料清洗與核心函式 ====================
def to_number(s):
    """將字串（如 '78.8萬'、'3.5萬km'）或數值正規化為標準整數，無效回傳 0"""
    if pd.isna(s):
        return 0
    if isinstance(s, (int, float)):
        return int(s)
    if not isinstance(s, str):
        return 0
    digits = re.sub(r"[^0-9.]", "", s)
    if not digits:
        return 0
    try:
        num = float(digits)
        return int(num * 10000) if "萬" in s else int(num)
    except ValueError:
        return 0


def process_dataframe(df):
    """標準化欄位名稱並完成價格與里程數值清洗"""
    if df is None or df.empty:
        return pd.DataFrame(
            columns=[
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
        )

    df = df.copy()

    # 欄位別名相容處理 (超連結 / 車輛連結 -> 連結)
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

    # 優先使用既有數值欄位 (價格_元 / 里程_km)，否則自文字清洗
    if "價格_元" in df.columns and df["價格_元"].notna().any():
        df["價格_數值"] = df["價格_元"].apply(to_number)
    else:
        df["價格_數值"] = df["價格"].apply(to_number)

    if "里程_km" in df.columns and df["里程_km"].notna().any():
        df["里程_數值"] = df["里程_km"].apply(to_number)
    else:
        df["里程_數值"] = df["里程"].apply(to_number)

    # 確保顯示文字格式齊全
    if df["價格"].astype(str).str.strip().eq("").all():
        df["價格"] = df["價格_數值"].apply(
            lambda v: f"{v/10000:.1f}萬" if v >= 10000 else f"{v}元"
        )
    if df["里程"].astype(str).str.strip().eq("").all():
        df["里程"] = df["里程_數值"].apply(
            lambda v: f"{v/10000:.1f}萬km" if v >= 10000 else f"{v}km"
        )

    return df


def load_data_from_source(source_path, is_csv=False):
    """從 SQLite 資料庫或 CSV 檔案載入並標準化車況資料"""
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
            raise Exception(f"無法解析 CSV 檔案 (多編碼皆失敗): {last_err}")
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
                raise Exception("資料庫中找不到有效資料表！")
        conn.close()

    return process_dataframe(df)


def load_initial_data():
    """啟動時自動探索資料來源：優先 SQLite，若無則自動載入 CSV 備用"""
    global current_source_name
    db_path = BASE_DIR / "auto8891.db"
    csv_fallback = BASE_DIR / "8891_cars.csv"

    if db_path.exists():
        try:
            df = load_data_from_source(db_path, is_csv=False)
            current_source_name = f"{db_path.name} (SQLite)"
            return df
        except Exception as e:
            print(f"嘗試載入預設 DB 失敗: {e}")

    if csv_fallback.exists():
        try:
            df = load_data_from_source(csv_fallback, is_csv=True)
            current_source_name = f"{csv_fallback.name} (CSV)"
            return df
        except Exception as e:
            print(f"嘗試載入預設 CSV 失敗: {e}")

    current_source_name = "未載入資料 (請手動選擇)"
    return pd.DataFrame(
        columns=["縣市", "標題", "年份", "里程", "價格", "價格_數值", "里程_數值"]
    )


# ==================== UI 資料更新與表格處理 ====================
def populate_table(target_df, tree_widget):
    for item in tree_widget.get_children():
        tree_widget.delete(item)

    if target_df is not None and not target_df.empty:
        for _, row in target_df.iterrows():
            tree_widget.insert(
                "",
                "end",
                values=(
                    row["縣市"],
                    row["標題"],
                    row["年份"],
                    row["里程"],
                    row["價格"],
                ),
            )


def apply_loaded_data(new_df, source_label_text):
    global df_global, current_filtered_df, all_counties, current_source_name
    df_global = new_df
    current_filtered_df = df_global.copy()
    current_source_name = source_label_text

    counties = sorted(
        [
            c
            for c in df_global["縣市"].replace("", pd.NA).dropna().unique()
            if str(c).strip()
        ]
    )
    all_counties = ["全部縣市"] + counties
    county_combo["values"] = all_counties
    county_combo.current(0)
    search_entry.delete(0, tk.END)

    populate_table(df_global, tree)
    source_var.set(
        f"📁 資料來源：{source_label_text} ｜ 共 {len(df_global):,} 筆車源"
    )
    status_var.set(
        f"✨ 已成功載入 {source_label_text}，共 {len(df_global):,} 筆車況資料。"
    )


def load_csv_dialog():
    """透過檔案選擇器載入使用者指定的 CSV 車況資料"""
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
        apply_loaded_data(new_df, f"{src_name} (CSV)")
        messagebox.showinfo(
            "載入成功",
            f"已成功從 CSV 載入 {len(new_df):,} 筆車源資料！\n檔案：{src_name}",
        )
    except Exception as e:
        messagebox.showerror("載入失敗", f"讀取 CSV 檔案失敗：{e}")


def reload_database():
    """重新從 SQLite 資料庫 (auto8891.db) 讀取最新數據"""
    db_path = BASE_DIR / "auto8891.db"
    if not db_path.exists():
        messagebox.showerror(
            "錯誤",
            f"找不到資料庫檔案：{db_path.name}\n請確認同目錄下是否有 auto8891.db",
        )
        return
    try:
        new_df = load_data_from_source(db_path, is_csv=False)
        apply_loaded_data(new_df, f"{db_path.name} (SQLite)")
        messagebox.showinfo(
            "成功",
            f"已成功從 SQLite 重新載入最新車源！共 {len(new_df):,} 筆。",
        )
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取資料庫失敗: {e}")


# ==================== 分頁一：車源查詢與篩選 ====================
def filter_cars():
    global current_filtered_df
    keyword = search_entry.get().strip()
    selected_county = county_combo.get().strip()

    filtered_df = df_global.copy()
    if selected_county and selected_county != "全部縣市":
        filtered_df = filtered_df[filtered_df["縣市"] == selected_county]
    if keyword:
        filtered_df = filtered_df[
            filtered_df["標題"].astype(str).str.contains(keyword, case=False, na=False)
        ]

    current_filtered_df = filtered_df
    populate_table(filtered_df, tree)
    status_var.set(f"🔍 篩選完成：共找到 {len(filtered_df)} 筆符合的車源。")


def reset_filter():
    global current_filtered_df
    search_entry.delete(0, tk.END)
    county_combo.current(0)
    current_filtered_df = df_global.copy()
    populate_table(df_global, tree)
    status_var.set(f"✨ 已顯示全部車源，總計 {len(df_global)} 筆。")


def export_query_to_csv():
    """匯出當前查詢篩選結果至 CSV"""
    if current_filtered_df is None or current_filtered_df.empty:
        messagebox.showwarning("提示", "目前表格中沒有車源資料可供匯出！")
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
            if c in current_filtered_df.columns
        ]
        export_df = (
            current_filtered_df[cols_to_export]
            if cols_to_export
            else current_filtered_df
        )
        export_df.to_csv(file_path, index=False, encoding="utf-8-sig")
        messagebox.showinfo(
            "匯出成功",
            f"已成功匯出 {len(export_df):,} 筆車源清單至：\n{file_path}\n\n採用 UTF-8-SIG 編碼，Excel 開啟不亂碼！",
        )
    except Exception as e:
        messagebox.showerror("匯出失敗", f"寫入 CSV 檔案失敗：{e}")


# ==================== 分頁二：AI 智慧估價與 CSV 匯出 ====================
def calculate_valuation():
    global last_matched_df, last_val_summary
    keyword = val_keyword_entry.get().strip()
    year_str = val_year_entry.get().strip()
    km_str = val_km_entry.get().strip()

    if df_global is None or df_global.empty:
        messagebox.showwarning(
            "警告",
            "目前無任何車輛資料，無法進行估價！\n請先確認已載入 SQLite 資料庫或 CSV 檔案。",
        )
        return

    if not keyword:
        messagebox.showwarning("提示", "請輸入欲估價的車款關鍵字！")
        return

    matched = df_global[
        df_global["標題"].astype(str).str.contains(keyword, case=False, na=False)
    ]
    if year_str.isdigit():
        year_matched = matched[
            matched["年份"].astype(str).str.contains(str(year_str), na=False)
        ]
        if not year_matched.empty:
            matched = year_matched

    matched = matched[matched["價格_數值"] > 0]

    if matched.empty:
        last_matched_df = None
        last_val_summary = None
        export_val_btn.config(state="disabled")
        save_history_btn.config(state="disabled")
        val_result_var.set(
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
    val_result_var.set(result_text)

    # 記錄本次估價成果供 CSV 匯出與歷史保存
    last_matched_df = matched.copy()
    last_val_summary = {
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
        "資料來源": current_source_name,
    }

    export_val_btn.config(state="normal")
    save_history_btn.config(state="normal")


def export_valuation_cars_to_csv():
    """匯出本次估價參考的所有車源明細至 CSV 檔案"""
    global last_matched_df
    if last_matched_df is None or last_matched_df.empty:
        messagebox.showwarning("提示", "目前尚無估價參考車源，請先執行估價！")
        return

    keyword = val_keyword_entry.get().strip() or "車款"
    year_str = val_year_entry.get().strip() or "年份"
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
            if c in last_matched_df.columns
        ]
        export_df = (
            last_matched_df[cols_to_export]
            if cols_to_export
            else last_matched_df
        )
        export_df.to_csv(file_path, index=False, encoding="utf-8-sig")
        messagebox.showinfo(
            "匯出成功",
            f"已成功匯出 {len(export_df):,} 筆參考車源至 CSV：\n{file_path}\n\n採用 UTF-8-SIG 編碼，Excel 開啟不亂碼！",
        )
    except Exception as e:
        messagebox.showerror("匯出失敗", f"寫入 CSV 檔案失敗：{e}")


def save_valuation_to_history():
    """將本次估價報告摘要追加儲存至 valuation_history.csv"""
    global last_val_summary
    if not last_val_summary:
        messagebox.showwarning("提示", "目前尚無估價報告，請先執行估價！")
        return

    history_path = BASE_DIR / "valuation_history.csv"
    file_exists = history_path.exists()

    try:
        history_df = pd.DataFrame([last_val_summary])
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


def open_history_file():
    """以作業系統預設程式開啟估價歷史紀錄 CSV"""
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


# ==================== 主程式視窗建構 ====================
root = tk.Tk()
root.title("🚗 8891 中古車快速查詢與智慧估價系統")
root.geometry("1100x740")
root.minsize(1000, 660)

# 載入初始資料 (優先 SQLite，自動探索備用 CSV)
df_global = load_initial_data()
current_filtered_df = df_global.copy()
all_counties = (
    ["全部縣市"]
    + sorted(
        [
            c
            for c in df_global["縣市"].replace("", pd.NA).dropna().unique()
            if str(c).strip()
        ]
    )
    if not df_global.empty
    else ["全部縣市"]
)

# 樣式設定
style = ttk.Style()
style.theme_use("clam")

BG_COLOR = "#F8FAFC"
PRIMARY_COLOR = "#0284C7"
TEXT_COLOR = "#1E293B"

root.configure(bg=BG_COLOR)

style.configure(".", font=("Microsoft JhengHei", 10))
style.configure("TLabel", background=BG_COLOR, foreground=TEXT_COLOR)
style.configure(
    "Header.TLabel",
    font=("Microsoft JhengHei", 16, "bold"),
    foreground="#0F172A",
)
style.configure(
    "TNotebook.Tab", font=("Microsoft JhengHei", 11, "bold"), padding=[15, 8]
)

style.configure(
    "Primary.TButton",
    font=("Microsoft JhengHei", 10, "bold"),
    foreground="white",
    background=PRIMARY_COLOR,
)
style.map(
    "Primary.TButton",
    background=[("active", "#0284C7"), ("pressed", "#0369A1")],
    foreground=[("active", "white"), ("pressed", "white")],
)

# 頂部標題與資料來源控制列
header_frame = ttk.Frame(root, padding=12)
header_frame.pack(fill="x")

title_box = ttk.Frame(header_frame)
title_box.pack(side="left")

ttk.Label(
    title_box, text="🚗 8891 中古車快速查詢與智慧估價系統", style="Header.TLabel"
).pack(anchor="w")

source_var = tk.StringVar(
    value=f"📁 資料來源：{current_source_name} ｜ 共 {len(df_global):,} 筆車源"
)
source_label = ttk.Label(
    title_box,
    textvariable=source_var,
    font=("Microsoft JhengHei", 9),
    foreground="#64748B",
)
source_label.pack(anchor="w", pady=(2, 0))

btn_box = ttk.Frame(header_frame)
btn_box.pack(side="right")

load_csv_btn = ttk.Button(btn_box, text="📂 載入 CSV 檔案", command=load_csv_dialog)
load_csv_btn.pack(side="left", padx=5)

reload_btn = ttk.Button(
    btn_box, text="🔄 重新載入 SQLite", command=reload_database
)
reload_btn.pack(side="left", padx=5)

# 分頁標籤容器
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=15, pady=5)

tab_query = ttk.Frame(notebook)
tab_valuation = ttk.Frame(notebook)

notebook.add(tab_query, text=" 🔍 車源快速查詢 ")
notebook.add(tab_valuation, text=" 💰 AI 智慧估價系統 ")

# ==================== 分頁一佈局：車源查詢 ====================
query_frame = ttk.Frame(tab_query, padding=15)
query_frame.pack(fill="both", expand=True)

filter_box = ttk.LabelFrame(query_frame, text=" 多維度條件篩選 ", padding=10)
filter_box.pack(fill="x", pady=(0, 10))

ttk.Label(filter_box, text="縣市：").pack(side="left", padx=(5, 2))
county_combo = ttk.Combobox(
    filter_box, values=all_counties, width=12, state="readonly"
)
county_combo.pack(side="left", padx=5)
county_combo.current(0)

ttk.Label(filter_box, text="關鍵字：").pack(side="left", padx=(15, 2))
search_entry = ttk.Entry(filter_box, width=22, font=("Microsoft JhengHei", 10))
search_entry.pack(side="left", padx=5)
search_entry.bind("<Return>", lambda event: filter_cars())

search_btn = ttk.Button(
    filter_box, text="查詢", command=filter_cars, style="Primary.TButton"
)
search_btn.pack(side="left", padx=5)

reset_btn = ttk.Button(filter_box, text="顯示全部", command=reset_filter)
reset_btn.pack(side="left", padx=5)

export_query_btn = ttk.Button(
    filter_box, text="📥 匯出查詢結果 (CSV)", command=export_query_to_csv
)
export_query_btn.pack(side="right", padx=5)

# 查詢表格展示區
table_frame = ttk.Frame(query_frame)
table_frame.pack(fill="both", expand=True)

columns = ("縣市", "標題", "年份", "里程", "價格")
tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=16)

tree.heading("縣市", text="所在縣市")
tree.heading("標題", text="車款完整名稱")
tree.heading("年份", text="出廠年份")
tree.heading("里程", text="行駛里程")
tree.heading("價格", text="售價")

tree.column("縣市", width=90, anchor="center")
tree.column("標題", width=420, anchor="w")
tree.column("年份", width=80, anchor="center")
tree.column("里程", width=110, anchor="center")
tree.column("價格", width=100, anchor="center")

scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)

tree.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

status_var = tk.StringVar(
    value=f"✨ 系統就緒，目前載入 {len(df_global):,} 筆車源資料。"
)
status_bar = ttk.Label(
    query_frame,
    textvariable=status_var,
    font=("Microsoft JhengHei", 9),
    foreground="#64748B",
)
status_bar.pack(anchor="w", pady=(5, 0))

populate_table(df_global, tree)

# ==================== 分頁二佈局：智慧估價系統 ====================
val_container = ttk.Frame(tab_valuation, padding=20)
val_container.pack(fill="both", expand=True)

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
val_keyword_entry = ttk.Entry(
    val_frame, width=30, font=("Microsoft JhengHei", 11)
)
val_keyword_entry.grid(row=0, column=1, sticky="w", pady=8)
val_keyword_entry.insert(0, "Altis")
val_keyword_entry.bind("<Return>", lambda event: calculate_valuation())

ttk.Label(
    val_frame, text="出廠年份 (如 2019)：", font=("Microsoft JhengHei", 10, "bold")
).grid(row=1, column=0, sticky="w", pady=8)
val_year_entry = ttk.Entry(
    val_frame, width=30, font=("Microsoft JhengHei", 11)
)
val_year_entry.grid(row=1, column=1, sticky="w", pady=8)
val_year_entry.insert(0, "2019")
val_year_entry.bind("<Return>", lambda event: calculate_valuation())

ttk.Label(
    val_frame,
    text="預計行駛里程 (km，選填)：",
    font=("Microsoft JhengHei", 10, "bold"),
).grid(row=2, column=0, sticky="w", pady=8)
val_km_entry = ttk.Entry(val_frame, width=30, font=("Microsoft JhengHei", 11))
val_km_entry.grid(row=2, column=1, sticky="w", pady=8)
val_km_entry.insert(0, "65000")
val_km_entry.bind("<Return>", lambda event: calculate_valuation())

calc_btn = ttk.Button(
    val_frame,
    text="🚀 開始智慧估價與市場行情分析",
    command=calculate_valuation,
    style="Primary.TButton",
)
calc_btn.grid(row=3, column=0, columnspan=2, pady=12)

# 估價報告與 CSV 匯出操作區
result_box = ttk.LabelFrame(
    val_container, text=" 市場行情評估結果與報表匯出 ", padding=15
)
result_box.pack(fill="both", expand=True, padx=10, pady=10)

val_result_var = tk.StringVar(
    value="請於上方輸入車款、年份與里程後，點擊按鈕產出行情分析報告。"
)
result_label = ttk.Label(
    result_box,
    textvariable=val_result_var,
    font=("Microsoft JhengHei", 11),
    foreground="#0369A1",
    justify="left",
)
result_label.pack(anchor="nw", fill="both", expand=True)

# 底部 CSV 動作按鈕列
val_action_frame = ttk.Frame(result_box)
val_action_frame.pack(fill="x", pady=(12, 0))

export_val_btn = ttk.Button(
    val_action_frame,
    text="📥 匯出參考車源明細 (CSV)",
    command=export_valuation_cars_to_csv,
    state="disabled",
)
export_val_btn.pack(side="left", padx=(0, 10))

save_history_btn = ttk.Button(
    val_action_frame,
    text="📑 儲存估價紀錄至歷史報表 (CSV)",
    command=save_valuation_to_history,
    state="disabled",
)
save_history_btn.pack(side="left", padx=(0, 10))

open_history_btn = ttk.Button(
    val_action_frame,
    text="📊 開啟歷史估價報表 (valuation_history.csv)",
    command=open_history_file,
)
open_history_btn.pack(side="left")

if __name__ == "__main__":
    root.mainloop()

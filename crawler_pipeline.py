"""8891 中古車爬蟲採集與商業分析自動化管線 (Crawler & Data Pipeline)

功能包含：
1. 資料清洗轉換 (字串轉純數值)
2. 爬蟲採集 (無頭瀏覽器動態渲染 / 離線模擬資料)
3. 雙軌持久化儲存 (UTF-8-SIG CSV 與 SQLite 資料庫)
4. 商業統計指標分析與行情折舊散佈圖繪製
"""

import argparse
import logging
from pathlib import Path
import re
import sqlite3
import time

import matplotlib.pyplot as plt
import pandas as pd

# 設定基礎目錄
BASE_DIR = Path(__file__).resolve().parent

# 日誌格式設定
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ==========================================
# 1. 資料清洗模組 (Data Engineering)
# ==========================================
def to_number(text: str | float | int | None) -> float | None:
    """將含有單位的字串（如 '78.8萬' 或 '3.5萬km'）清洗轉換為純數值。

    異常或無法解析時回傳 None。
    """
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


# ==========================================
# 2. 動態爬蟲採集模組 (Headless Scraping)
# ==========================================
def get_mock_car_data() -> list[list[str]]:
    """提供標準離線模擬展示數據，確保無網路或反爬阻擋環境亦可完整驗證管線。"""
    return [
        [
            "台北市",
            "2021 Toyota RAV4 2.0",
            "78.8萬",
            "2021",
            "3.5萬km",
            "http://img1.jpg",
            "https://auto.8891.com.tw/car1",
        ],
        [
            "新北市",
            "2020 Honda CR-V 1.5",
            "68.5萬",
            "2020",
            "4.8萬km",
            "http://img2.jpg",
            "https://auto.8891.com.tw/car2",
        ],
        [
            "台中市",
            "2019 Mazda CX-5 2.0",
            "59.8萬",
            "2019",
            "6.2萬km",
            "http://img3.jpg",
            "https://auto.8891.com.tw/car3",
        ],
        [
            "高雄市",
            "2022 Toyota CC 1.8",
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
            "2023 Lexus NX200",
            "138.0萬",
            "2023",
            "1.2萬km",
            "http://img6.jpg",
            "https://auto.8891.com.tw/car6",
        ],
    ]


def fetch_car_data(target_url: str | None = None) -> list[list[str]]:
    """採集車輛資訊。若未提供 target_url，則使用離線模擬數據。"""
    if not target_url:
        logger.info("📦 採用離線模擬資料集進行分析...")
        return get_mock_car_data()

    try:
        from bs4 import BeautifulSoup
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
    except ImportError as e:
        logger.error("未安裝 selenium 或 beautifulsoup4 套件，將降級為離線展示資料: %s", e)
        return get_mock_car_data()

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    results: list[list[str]] = []
    driver = None
    try:
        logger.info("🌐 啟動無頭 Chrome 爬取目標網址：%s", target_url)
        driver = webdriver.Chrome(options=options)
        driver.get(target_url)
        time.sleep(2.5)  # 等待前端動態渲染

        soup = BeautifulSoup(driver.page_source, "html.parser")
        items = soup.select("a.car-item") or soup.select("a[href*='usedauto-infos']")
        for item in items:
            title = item.text.strip().split("\n")[0]
            href = item.get("href", "")
            if href and not href.startswith("http"):
                href = "https://auto.8891.com.tw/" + href.lstrip("/")
            if title:
                results.append(["台灣", title, "面議", "2020", "0km", "", href])
        logger.info("採集完成，獲取 %d 筆資料", len(results))
    except Exception as e:
        logger.warning("⚠️ 爬取過程發生異常: %s，將使用離線模擬資料", e)
        return get_mock_car_data()
    finally:
        if driver:
            driver.quit()

    return results if results else get_mock_car_data()


# ==========================================
# 3. 儲存與商業分析管線 (Pipeline & BI)
# ==========================================
def run_pipeline(
    target_url: str | None = None,
    csv_filename: str = "8891_cars.csv",
    db_filename: str = "cars_market.db",
    chart_filename: str = "market_analysis.png",
) -> None:
    """執行完整資料處理管線：採集 -> 清洗 -> 儲存 -> 分析製圖。"""
    raw_data = fetch_car_data(target_url)
    cols = ["縣市", "標題", "價格", "年份", "里程", "圖片連結", "超連結"]
    df = pd.DataFrame(raw_data, columns=cols)

    # 清洗數值並過濾面議/無效價格
    df["價格_元"] = df["價格"].apply(to_number)
    df["里程_km"] = df["里程"].apply(to_number)
    df_clean = df.dropna(subset=["價格_元"]).copy()

    if df_clean.empty:
        logger.warning("⚠️ 清洗後無有效車輛數據！")
        return

    csv_path = BASE_DIR / csv_filename
    db_path = BASE_DIR / db_filename
    chart_path = BASE_DIR / chart_filename

    # 持久化儲存 (UTF-8-SIG CSV 與 SQLite)
    df_clean.to_csv(csv_path, index=False, encoding="utf-8-sig")
    with sqlite3.connect(db_path) as conn:
        df_clean.to_sql("car", conn, if_exists="replace", index=False)

    logger.info("✅ 數據已持久化儲存至：%s 及 %s", csv_path.name, db_path.name)

    # 輸出統計摘要
    print("\n" + "=" * 50)
    print("📊 【各縣市行情統計】")
    city_stat = df_clean.groupby("縣市")["價格_元"].agg(
        筆數="count", 平均價格="mean"
    )
    city_stat["平均價格"] = city_stat["平均價格"].map("{:,.0f} 元".format)
    print(city_stat)

    valid_mileage = df_clean.dropna(subset=["里程_km"])
    if len(valid_mileage) > 1:
        corr = valid_mileage[["價格_元", "里程_km"]].corr().iloc[0, 1]
        print(f"\n📈 【價格與里程相關係數】：{corr:.4f}")

    # 繪製行情折舊散佈圖
    plt.figure(figsize=(8, 5))
    plt.scatter(
        df_clean["里程_km"],
        df_clean["價格_元"],
        c="#0284c7",
        s=100,
        alpha=0.85,
        edgecolors="none",
    )
    plt.title("Used Car Market: Mileage vs Price", fontsize=13, fontweight="bold")
    plt.xlabel("Mileage (km)", fontsize=11)
    plt.ylabel("Price (NTD)", fontsize=11)
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.tight_layout()
    plt.savefig(chart_path, dpi=300)
    plt.close()

    print(f"\n📈 市場折舊分析圖已輸出至：{chart_path.name}")
    print("=" * 50 + "\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="8891 二手車資料採集與商業分析管線")
    parser.add_argument(
        "--url",
        type=str,
        default=None,
        help="指定爬取的 8891 網址（若未提供則使用離線模擬資料）",
    )
    args = parser.parse_args()
    run_pipeline(target_url=args.url)


if __name__ == "__main__":
    main()

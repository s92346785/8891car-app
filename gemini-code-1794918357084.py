import re
import time
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options


# ==========================================
# 1. 資料清洗模組 (Data Engineering)
# ==========================================
def to_number(text):
    """將 '78.8萬' 或 '3.5萬km' 清洗轉換為純數值，異常回傳 None"""
    if not text:
        return None
    text = str(text).replace(",", "").strip()
    num = re.findall(r"[\d\.]+", text)
    if not num:
        return None
    val = float(num[0])
    return val * 10000 if "萬" in text else val


# ==========================================
# 2. 動態爬蟲採集模組 (Headless Scraping)
# ==========================================
def fetch_car_data(target_url=None):
    """啟動無頭瀏覽器採集單頁車輛資訊 (含生命週期管理)"""
    if not target_url:
        # 預設離線備用測試資料 (確保無網路環境亦可展示完整分析流程)
        return [
            [
                "台北市",
                "2021 Toyota RAV4 2.0",
                "78.8萬",
                "2021",
                "3.5萬km",
                "http://img1.jpg",
                "http://car1",
            ],
            [
                "新北市",
                "2020 Honda CR-V 1.5",
                "68.5萬",
                "2020",
                "4.8萬km",
                "http://img2.jpg",
                "http://car2",
            ],
            [
                "台中市",
                "2019 Mazda CX-5 2.0",
                "59.8萬",
                "2019",
                "6.2萬km",
                "http://img3.jpg",
                "http://car3",
            ],
            [
                "高雄市",
                "2022 Toyota CC 1.8",
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
                "2023 Lexus NX200",
                "138.0萬",
                "2023",
                "1.2萬km",
                "http://img6.jpg",
                "http://car6",
            ],
        ]

    options = Options()
    options.add_argument("--headless")  # 無頭模式背景執行
    driver = webdriver.Chrome(options=options)
    results = []
    try:
        driver.get(target_url)
        time.sleep(2)  # 等待前端動態渲染
        soup = BeautifulSoup(driver.page_source, "html.parser")
        # 依實際 DOM 結構提取欄位 (此處以標準結構示意)
        for item in soup.select("a.car-item"):
            title = item.text.strip()
            href = "https://auto.8891.com.tw/" + item.get("href", "").lstrip("/")
            results.append(["未知", title, "面議", "2020", "0km", "", href])
    except Exception as e:
        print(f"⚠️ 採集異常略過: {e}")
    finally:
        driver.quit()  # 確保釋放記憶體資源
    return results


# ==========================================
# 3. 儲存與商業分析管線 (Pipeline & BI)
# ==========================================
def run_pipeline():
    # (A) 取得數據並轉為 DataFrame
    raw_data = fetch_car_data()  # 若要實機爬取可傳入真實網址
    cols = ["縣市", "標題", "價格", "年份", "里程", "圖片連結", "超連結"]
    df = pd.DataFrame(raw_data, columns=cols)

    # (B) 數值轉換與過濾「面議」無效樣本
    df["價格_元"] = df["價格"].apply(to_number)
    df["里程_km"] = df["里程"].apply(to_number)
    df_clean = df.dropna(subset=["價格_元"]).copy()

    # (C) 雙軌持久化儲存：匯出防亂碼 CSV 與 SQLite 資料庫
    df_clean.to_csv("8891_cars.csv", index=False, encoding="utf-8-sig")
    with sqlite3.connect("cars_market.db") as conn:
        df_clean.to_sql("car", conn, if_exists="replace", index=False)

    # (D) 商業統計洞察與折舊散佈圖產出
    print(
        "📊 【各縣市行情統計】\n",
        df_clean.groupby("縣市")["價格_元"].agg(["count", "mean"]),
    )
    print("\n📈 【價格與里程相關係數】\n", df_clean[["價格_元", "里程_km"]].corr())

    plt.figure(figsize=(8, 5))
    plt.scatter(df_clean["里程_km"], df_clean["價格_元"], c="#0284c7", s=100, alpha=0.8)
    plt.title("Used Car Market: Mileage vs Price", fontsize=13)
    plt.xlabel("Mileage (km)")
    plt.ylabel("Price (NTD)")
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.savefig("market_analysis.png", dpi=300)
    plt.close()
    print(
        "\n✅ 管線執行完畢！已產出 8891_cars.csv、cars_market.db 與 market_analysis.png"
    )


if __name__ == "__main__":
    run_pipeline()

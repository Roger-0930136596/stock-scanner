from flask import Flask, render_template_string, redirect
import urllib.request
import json
import pandas as pd
import datetime
import threading
import schedule
import time
import ssl

app = Flask(__name__)
last_updated = ""
result_data = []

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

def get_twse(date_str):
    url = f"https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date={date_str}&type=ALL"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, context=ctx, timeout=20) as res:
        return json.loads(res.read().decode("utf-8"))

def get_stock_data(date_obj):
    """抓取指定日期的股票資料，回傳 DataFrame 或 None"""
    date_str = date_obj.strftime("%Y%m%d")
    try:
        data = get_twse(date_str)
        tables = data.get("tables", [])
        for t in tables:
            if "每日收盤行情" in t.get("title", ""):
                rows = t.get("data", [])
                fields = t.get("fields", [])
                if rows and fields:
                    df = pd.DataFrame(rows, columns=fields)
                    return df
    except Exception as e:
        print(f"抓取 {date_str} 失敗：{e}")
    return None

def find_latest_trading_day(start_date, max_days=30):
    """從 start_date 往前找，找到有資料的最近交易日"""
    for i in range(max_days):
        d = start_date - datetime.timedelta(days=i)
        # 跳過週六(5)、週日(6)
        if d.weekday() >= 5:
            continue
        print(f"嘗試日期：{d.strftime('%Y%m%d')}")
        df = get_stock_data(d)
        if df is not None and not df.empty:
            print(f"找到交易日：{d.strftime('%Y%m%d')}")
            return d, df
    return None, None

def fetch_data():
    global last_updated, result_data
    try:
        today = datetime.date.today()
        print(f"\n開始選股，今天日期：{today}")

        # 找今日（或最近交易日）資料
        today_date, df = find_latest_trading_day(today)
        if df is None:
            last_updated = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "（找不到近期資料）"
            return

        # 找成交量欄位
        vol_col = next((c for c in df.columns if "成交" in c and "股" in c), df.columns[2])
        close_col = next((c for c in df.columns if "收盤" in c), None)
        ud_col = next((c for c in df.columns if "漲跌" in c and "價差" not in c and "幅" not in c), None)
        id_col = df.columns[0]
        name_col = df.columns[1]

        df["今日張數"] = pd.to_numeric(
            df[vol_col].astype(str).str.replace(",", ""), errors="coerce"
        ).fillna(0).astype(int) // 1000

        print(f"今日資料：{today_date}，共 {len(df)} 筆")

        # 找昨日資料（從今日往前一天開始找）
        yesterday_start = today_date - datetime.timedelta(days=1)
        yest_date, df2_raw = find_latest_trading_day(yesterday_start)

        if df2_raw is None:
            last_updated = "找不到昨日資料"
            return

        df2_raw["昨日張數"] = pd.to_numeric(
            df2_raw[vol_col].astype(str).str.replace(",", ""), errors="coerce"
        ).fillna(0).astype(int) // 1000

        df2 = df2_raw[[id_col, "昨日張數"]].copy()
        print(f"昨日資料：{yest_date}，共 {len(df2)} 筆")

        # 合併比對
        merged = df.merge(df2, on=id_col)
        
        # 排除權證、認購、認售（代號開頭為0的ETF保留，但6碼的權證排除）
        merged = merged[~merged[name_col].str.contains("認購|認售|權證|購|售", na=False)]
        merged = merged[merged[id_col].str.len() <= 4]  # 只保留4碼以內的股票代號
        
        merged = merged[merged["昨日張數"] >= 500]
        merged = merged[merged["今日張數"] > 0]
        merged["量比"] = (merged["今日張數"] / merged["昨日張數"]).round(2)
        result = merged[merged["量比"] >= 1.5].copy()
        result = result.sort_values("量比", ascending=False)

        output = []
        for _, row in result.iterrows():
            output.append({
                "股票代號": row[id_col],
                "股票名稱": row[name_col],
                "昨日張數": int(row["昨日張數"]),
                "今日張數": int(row["今日張數"]),
                "量比": float(row["量比"]),
                "收盤價": row[close_col] if close_col else "",
                "漲跌": row[ud_col] if ud_col else "",
            })

        result_data = output
        last_updated = f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}（今日：{today_date}，昨日：{yest_date}）"
        print(f"選股完成，共 {len(result_data)} 支")

    except Exception as e:
        last_updated = f"錯誤：{e}"
        print(f"錯誤：{e}")

HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>台股爆量選股</title>
<style>
  body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
  h1 { color: #333; }
  .updated { color: #888; font-size: 14px; margin-bottom: 20px; }
  table { width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }
  th { background: #2c3e50; color: white; padding: 12px; text-align: center; }
  td { padding: 10px 12px; text-align: center; border-bottom: 1px solid #eee; }
  tr:hover { background: #f9f9f9; }
  .up { color: #e74c3c; font-weight: bold; }
  .down { color: #27ae60; font-weight: bold; }
  .ratio { background: #e8f5e9; color: #2e7d32; font-weight: bold; padding: 4px 10px; border-radius: 12px; }
  .btn { display: inline-block; margin-top: 10px; padding: 10px 24px; background: #2c3e50; color: white; border-radius: 6px; border: none; font-size: 15px; cursor: pointer; }
  .btn:hover { background: #34495e; }
  .count { margin: 12px 0; font-size: 15px; }
</style>
</head>
<body>
<h1>台股爆量選股（量比 &gt; 1.5倍）</h1>
<div class="updated">最後更新：{{ updated }}</div>
<form method="post" action="/refresh">
  <button class="btn" type="submit">立即重新抓取</button>
</form>
<br>
{% if data %}
<div class="count">共找到 <strong>{{ data|length }}</strong> 支爆量股票</div>
<table>
  <tr><th>獲利</th><th>K線</th><th>代號</th><th>名稱</th><th>昨日張數</th><th>今日張數</th><th>量比</th><th>收盤價</th><th>漲跌</th></tr>
  {% for row in data %}
  <tr>
    <td><a href="https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID={{ row['股票代號'] }}" target="_blank" style="color:#185FA5; text-decoration:none; font-size:13px;">獲利</a></td>
    <td><a href="https://tw.stock.yahoo.com/quote/{{ row['股票代號'] }}.TW/technical-analysis" target="_blank" style="color:#185FA5; text-decoration:none; font-size:13px;">K線</a></td>
    <td>{{ row['股票代號'] }}</td>
    <td>{{ row['股票名稱'] }}</td>
    <td>{{ "{:,}".format(row['昨日張數']) }}</td>
    <td>{{ "{:,}".format(row['今日張數']) }}</td>
    <td><span class="ratio">{{ row['量比'] }}x</span></td>
    <td>{{ row['收盤價'] }}</td>
    <td class="{{ 'up' if row['漲跌'] and '+' in row['漲跌']|string else 'down' }}">{{ row['漲跌']|striptags }}</td>
  </tr>
  {% endfor %}
</table>
{% else %}
<p>目前無資料，請按「立即重新抓取」，或等待今日收盤後再查看。</p>
{% endif %}
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(HTML, data=result_data, updated=last_updated or "尚未更新")

@app.route("/refresh", methods=["POST"])
def refresh():
    fetch_data()
    return redirect("/")

def run_schedule():
    schedule.every().day.at("13:00").do(fetch_data)
    while True:
        schedule.run_pending()
        time.sleep(30)

if __name__ == "__main__":
    t = threading.Thread(target=run_schedule, daemon=True)
    t.start()
    fetch_data()
    app.run(host="0.0.0.0", port=5000)
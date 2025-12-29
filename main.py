import requests
import pandas as pd
from datetime import datetime, timedelta
import time
import io
import random
import os

def get_weekly_public_announcements():
    # 1. 計算日期範圍：上週一 至 上週日
    today = datetime.now()
    # 取得上週一 (Today - weekday - 7)
    last_monday = today - timedelta(days=today.weekday() + 7)
    last_sunday = last_monday + timedelta(days=6)
    
    # 轉成民國年格式字串
    s_y, s_m, s_d = str(last_monday.year - 1911), str(last_monday.month), str(last_monday.day)
    e_y, e_m, e_d = str(last_sunday.year - 1911), str(last_sunday.month), str(last_sunday.day)
    
    print(f"📅 執行爬取區間: {last_monday.date()} ~ {last_sunday.date()}")
    
    url = "https://mopsov.twse.com.tw/mops/web/ajax_t05st02"
    
    # pub: 公開發行, sii: 上市, otc: 上櫃, rotc: 興櫃
    market_types = {'pub': '公開發行', 'sii': '上市', 'otc': '上櫃', 'rotc': '興櫃'}
    all_data = []

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded'
    }

    for market_code, market_name in market_types.items():
        print(f"🔎 正在掃描市場: {market_name} ({market_code})...")
        payload = {
            'encodeURIComponent': '1', 'step': '1', 'firstin': '1', 'off': '1',
            'year': s_y, 'month': s_m, 'day': s_d,
            'year2': e_y, 'month2': e_m, 'day2': e_d,
            'typek': market_code, 'co_id': '', 'spoke_time': '1',
        }
        
        try:
            r = requests.post(url, data=payload, headers=headers)
            r.encoding = 'utf8'
            time.sleep(random.uniform(2, 5)) # 稍微久一點避免被擋
            
            if "查無資料" in r.text:
                continue
                
            dfs = pd.read_html(io.StringIO(r.text))
            for df in dfs:
                if any(col in str(df.columns) for col in ['公司代號', '主旨', '案由']):
                    df['市場類別'] = market_name
                    all_data.append(df)
                    break
        except Exception as e:
            print(f"   - {market_name} 錯誤: {e}")

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        keywords = ['資金貸與', '背書保證', '會計師', '更換', '解任', '委任']
        
        subject_col = None
        for col in final_df.columns:
            if '主旨' in str(col) or '案由' in str(col):
                subject_col = col
                break
        
        if subject_col:
            mask = final_df[subject_col].astype(str).apply(lambda x: any(k in x for k in keywords))
            filtered_df = final_df[mask]
            
            # 設定輸出檔名
            filename = f"weekly_report_{last_sunday.date()}.xlsx"
            
            # 重要：確保輸出目錄存在 (GitHub Actions 有時需要)
            filtered_df.to_excel(filename, index=False)
            print(f"✅ 檔案已產生: {filename}")
        else:
            print("❌ 找不到主旨欄位")
    else:
        print("❌ 本週無資料")

if __name__ == "__main__":
    get_weekly_public_announcements()

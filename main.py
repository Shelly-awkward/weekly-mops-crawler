import requests
import pandas as pd
from datetime import datetime, timedelta
import time
import io
import random
import os

def get_weekly_public_announcements():
    # 1. 計算日期範圍
    today = datetime.now()
    last_monday = today - timedelta(days=today.weekday() + 7)
    last_sunday = last_monday + timedelta(days=6)
    
    s_y, s_m, s_d = str(last_monday.year - 1911), str(last_monday.month), str(last_monday.day)
    e_y, e_m, e_d = str(last_sunday.year - 1911), str(last_sunday.month), str(last_sunday.day)
    
    date_range_str = f"{last_monday.date()} ~ {last_sunday.date()}"
    print(f"📅 執行爬取區間: {date_range_str}")
    
    url = "https://mopsov.twse.com.tw/mops/web/ajax_t05st02"
    market_types = {'pub': '公開發行', 'sii': '上市', 'otc': '上櫃', 'rotc': '興櫃'}
    
    all_data = []
    log_messages = [] # 用來記錄執行狀況存入 Excel

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded'
    }

    for market_code, market_name in market_types.items():
        print(f"🔎 掃描: {market_name} ({market_code})...")
        payload = {
            'encodeURIComponent': '1', 'step': '1', 'firstin': '1', 'off': '1',
            'year': s_y, 'month': s_m, 'day': s_d,
            'year2': e_y, 'month2': e_m, 'day2': e_d,
            'typek': market_code, 'co_id': '', 'spoke_time': '1',
        }
        
        try:
            r = requests.post(url, data=payload, headers=headers, timeout=30)
            r.encoding = 'utf8'
            
            if "查無資料" in r.text:
                msg = f"{market_name}: 官方回傳查無資料"
                print(msg)
                log_messages.append(msg)
                continue
                
            # 嘗試解析表格
            try:
                dfs = pd.read_html(io.StringIO(r.text))
            except ValueError:
                msg = f"{market_name}: 無法解析 HTML 表格 (可能是被擋 IP 或格式改變)"
                print(msg)
                log_messages.append(msg)
                continue

            found_table = False
            for df in dfs:
                if any(col in str(df.columns) for col in ['公司代號', '主旨', '案由']):
                    df['市場類別'] = market_name
                    # 轉成字串避免合併錯誤
                    df = df.astype(str)
                    all_data.append(df)
                    found_table = True
                    log_messages.append(f"{market_name}: 成功取得 {len(df)} 筆原始資料")
                    break
            
            if not found_table:
                log_messages.append(f"{market_name}: 有回應但找不到目標表格")
                
            time.sleep(random.uniform(3, 6)) # 延長休息時間
            
        except Exception as e:
            err_msg = f"{market_name} 連線錯誤: {str(e)}"
            print(err_msg)
            log_messages.append(err_msg)

    # 準備輸出
    filename = f"weekly_report_{last_sunday.date()}.xlsx"
    
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
            
            if not filtered_df.empty:
                filtered_df.to_excel(filename, index=False)
                print(f"✅ 成功產出資料: {filename}")
            else:
                # 抓到了但沒有符合關鍵字的
                pd.DataFrame({'狀態': ['有抓到資料，但無符合關鍵字(資金貸與/背書/會計師)之公告'], '檢查區間': [date_range_str], '執行紀錄': [' | '.join(log_messages)]}).to_excel(filename, index=False)
                print(f"⚠️ 無符合關鍵字資料，已產出除錯報表: {filename}")
        else:
             pd.DataFrame({'狀態': ['找不到主旨欄位'], '執行紀錄': [' | '.join(log_messages)]}).to_excel(filename, index=False)
    else:
        # 完全沒抓到資料 (可能被擋)
        pd.DataFrame({'狀態': ['完全無資料 (可能被 MOPS 封鎖 IP)'], '檢查區間': [date_range_str], '執行紀錄': [' | '.join(log_messages)]}).to_excel(filename, index=False)
        print(f"❌ 無資料，已產出錯誤紀錄表: {filename}")

if __name__ == "__main__":
    get_weekly_public_announcements()

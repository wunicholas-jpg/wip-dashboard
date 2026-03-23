import streamlit as st
import pandas as pd

# 設定網頁標題與寬度
st.set_page_config(page_title="WIP 智能管理系統", layout="wide")
st.title("🏭 OSAT ATE FT WIP 智能管理系統")
st.markdown("### 第一階段：資料自動清洗與結構化 (ETL)")

# 1. 建立檔案上傳區塊
uploaded_file = st.file_uploader("📥 請上傳 OSAT 提供的 WIP CSV 檔案 (例如: ZC13 WIP 2026MP.xlsx - 20260323.csv)", type=['csv'])

if uploaded_file is not None:
    # ---------------------------------------------------------
    # 模組 A：讀取原始資料
    # ---------------------------------------------------------
    st.subheader("🔍 1. 原始資料 (OSAT 匯出格式)")
    # 讀取 CSV
    df_raw = pd.read_csv(uploaded_file)
    st.write("這是一張「寬表」，人類容易看，但 AI 和系統很難做加總與關聯分析：")
    st.dataframe(df_raw.head(10), use_container_width=True)

    # ---------------------------------------------------------
    # 模組 B：開始清洗資料 (Option A 的核心邏輯)
    # ---------------------------------------------------------
    st.subheader("✨ 2. 清洗後的資料 (AI 可讀的標準資料庫格式)")
    
    with st.spinner('資料清洗中...'):
        # 1. 將第一欄重新命名為 'Station' (站點)
        df_raw.rename(columns={df_raw.columns[0]: 'Station'}, inplace=True)
        
        # 2. 過濾掉不要的雜亂列 (例如時間列 08:00:00、Sum、機台配置等)
        # 確保 Station 欄位是字串才能進行過濾
        df_raw['Station'] = df_raw['Station'].astype(str)
        exclude_keywords = ['Process Step', 'Sum', 'TTL', 'Cum', '機台配置', 'nan', 'NaN']
        df_clean = df_raw[~df_raw['Station'].isin(exclude_keywords)].copy()
        
        # 3. 找出所有屬於「日期」的欄位 (這裡簡單判斷欄位名稱是否有 '202' 年份)
        date_cols = [col for col in df_clean.columns if '202' in str(col)]
        
        # 4. 關鍵轉換：用 melt 將「寬表」轉為「長表」(扁平化)
        df_melted = df_clean.melt(
            id_vars=['Station'], 
            value_vars=date_cols, 
            var_name='Date', 
            value_name='WIP_Qty'
        )
        
        # 5. 清理數值：把空值或無法轉換的字串變成 0
        df_melted['WIP_Qty'] = pd.to_numeric(df_melted['WIP_Qty'], errors='coerce').fillna(0)
        
        # 6. 依據日期與站點排序，讓資料更整齊
        df_melted = df_melted.sort_values(by=['Date', 'Station']).reset_index(drop=True)

    st.write("我們將矩陣轉換成了 `[站點, 日期, WIP數量]` 的標準格式，這也是未來丟給 AI Agent 的格式：")
    st.dataframe(df_melted, use_container_width=True)

    # ---------------------------------------------------------
    # 模組 C：簡單視覺化驗證 (讓你知道清洗成功了)
    # ---------------------------------------------------------
    st.subheader("📊 3. 快速驗證：最新日期的各站點 WIP 狀況")
    
    # 抓取資料中最新的一天
    latest_date = df_melted['Date'].max()
    st.write(f"當前顯示日期：**{latest_date}**")
    
    # 篩選最新一天的資料，並排除 WIP 為 0 的站點讓圖表更乾淨
    df_latest = df_melted[(df_melted['Date'] == latest_date) & (df_melted['WIP_Qty'] > 0)]
    
    # 使用 Streamlit 內建的圖表快速呈現
    st.bar_chart(data=df_latest, x='Station', y='WIP_Qty', use_container_width=True)
    
    st.success("🎉 資料清洗與轉換成功！這份乾淨的 DataFrame (df_melted) 已經準備好在下一步餵給 AI Agent 了。")

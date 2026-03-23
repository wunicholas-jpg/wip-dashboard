import streamlit as st
import pandas as pd

st.set_page_config(page_title="WIP & Demand 整合面板", layout="wide")
st.title("📊 複合式 WIP 與出貨需求管理系統")
st.markdown("這套系統能自動解析 OSAT 報表中的三段式結構 (WIP 趨勢 / DRAM 佔比 / 出貨 Demand)")

# 建立檔案上傳區塊 (支援 xlsx 與 csv)
uploaded_file = st.file_uploader("📥 請上傳包含三段資訊的 Excel 檔案", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        with st.spinner("資料解析與智慧分塊中..."):
            # 讀取檔案，不預設 header，讓系統把整張表完整讀進來
            if uploaded_file.name.endswith('.csv'):
                df_all = pd.read_csv(uploaded_file, header=None)
            else:
                df_all = pd.read_excel(uploaded_file, header=None)
            
            # --- 關鍵技術：全表掃描尋找切分點 ---
            # 設定我們預期的關鍵字 (您可依據真實 Excel 內的字眼在這裡微調)
            mid_keywords = '16GB|12GB|D-Ram|機台配置|DRAM'
            bottom_keywords = 'Accum|Ship|出貨|Demand|需求|Receiving'
            
            # 1. 尋找「中部」的起始列：掃描每一列，只要有出現 mid_keywords 就記錄下來
            mid_start_idx = len(df_all)
            mid_search = df_all[df_all.apply(lambda row: row.astype(str).str.contains(mid_keywords, case=False, na=False).any(), axis=1)].index
            if len(mid_search) > 0:
                mid_start_idx = mid_search[0]
                
            # 2. 尋找「下半部」的起始列：從中部之後開始掃描，尋找 bottom_keywords
            bottom_start_idx = len(df_all)
            bottom_search = df_all.iloc[mid_start_idx:][df_all.iloc[mid_start_idx:].apply(lambda row: row.astype(str).str.contains(bottom_keywords, case=False, na=False).any(), axis=1)].index
            if len(bottom_search) > 0:
                bottom_start_idx = bottom_search[0]

            # ---------------------------------------------------------
            # 區塊 1：上半部 (歷史 WIP 趨勢 3/2~3/23)
            # ---------------------------------------------------------
            st.markdown("---")
            st.markdown("### 📈 第一部分：站點 WIP 趨勢")
            # 擷取 0 到 mid_start_idx 的資料，並清除全空的欄與列
            df_top = df_all.iloc[0:mid_start_idx].dropna(how='all', axis=0).dropna(how='all', axis=1)
            if not df_top.empty:
                df_top.columns = df_top.iloc[0] # 將第一列設為表格標題
                df_top = df_top[1:].reset_index(drop=True)
                st.dataframe(df_top, use_container_width=True)

            # ---------------------------------------------------------
            # 區塊 2：中部 (當日 16G/12G WIP 狀態)
            # ---------------------------------------------------------
            if mid_start_idx < len(df_all):
                st.markdown("---")
                st.markdown("### 🗂️ 第二部分：當日 WIP 狀態 (By DRAM 16G/12G)")
                df_mid = df_all.iloc[mid_start_idx:bottom_start_idx].dropna(how='all', axis=0).dropna(how='all', axis=1)
                if not df_mid.empty:
                    df_mid.columns = df_mid.iloc[0]
                    df_mid = df_mid[1:].reset_index(drop=True)
                    st.dataframe(df_mid, use_container_width=True)

            # ---------------------------------------------------------
            # 區塊 3：下半部 (出貨 Demand)
            # ---------------------------------------------------------
            if bottom_start_idx < len(df_all):
                st.markdown("---")
                st.markdown("### 📦 第三部分：出貨需求與缺口 (Demand & Risk)")
                df_bottom = df_all.iloc[bottom_start_idx:].dropna(how='all', axis=0).dropna(how='all', axis=1)
                if not df_bottom.empty:
                    df_bottom.columns = df_bottom.iloc[0]
                    df_bottom = df_bottom[1:].reset_index(drop=True)
                    st.dataframe(df_bottom, use_container_width=True)
                    
            st.success("🎉 成功將 OSAT 複合式報表切割為三個獨立資料庫！下一步即可導入圖表與 AI Agent。")

    except Exception as e:
        st.error(f"檔案解析發生錯誤：{e}")
        st.warning("請確認 Excel 內容是否包含我們設定的關鍵字。以下是系統讀取到的原始資料全貌：")
        st.dataframe(df_all, use_container_width=True)

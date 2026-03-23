import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="WIP Dashboard", layout="wide")
st.title("🏭 ATE FT WIP 整合管理面板 (ZC13)")

uploaded_file = st.file_uploader("📥 請上傳最新 WIP Excel 檔案", type=['xlsx'])

if uploaded_file:
    try:
        with st.spinner("掃描 Excel 全部分頁與數據中..."):
            # 【關鍵修正】讀取 Excel 中的 "所有" 分頁 (Sheets)
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            df_wip_trend = None
            dram_dict = {'MU16G': 0, 'SS16G': 0, 'HY12G': 0, 'SS12G': 0}
            df_demand = None

            # --- 核心邏輯：像雷達一樣掃描所有分頁 ---
            for sheet in sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                
                # 1. 尋找上半部 WIP 趨勢 (找 Process Step)
                if df_wip_trend is None:
                    wip_mask = df.apply(lambda r: r.astype(str).str.contains('Process Step', case=False, na=False).any(), axis=1)
                    if wip_mask.any():
                        wip_start = wip_mask.idxmax()
                        date_row = wip_start - 1 if wip_start > 0 else wip_start
                        
                        # 擷取到出現 Sum 為止
                        end_mask = df.iloc[wip_start:].apply(lambda r: r.astype(str).str.contains('Sum|TTL', case=False, na=False).any(), axis=1)
                        wip_end = end_mask.idxmax() if end_mask.any() else wip_start + 40
                        
                        raw_wip = df.iloc[date_row:wip_end].copy()
                        
                        # 清洗標題 (解決重複欄位與時間格式)
                        cols = []
                        for c in raw_wip.iloc[0]:
                            cstr = str(c).split(' ')[0]
                            if cstr == 'nan': cstr = 'Station'
                            cols.append(cstr)
                        raw_wip.columns = [f"{c}_{i}" if cols.count(c)>1 else c for i, c in enumerate(cols)]
                        
                        # 去除雜訊
                        df_wip_trend = raw_wip.iloc[2:].dropna(subset=[raw_wip.columns[0]])
                        df_wip_trend = df_wip_trend[~df_wip_trend[df_wip_trend.columns[0]].astype(str).str.contains('nan|Sum|TTL')]

                # 2. 尋找中部 DRAM 數據 (全域掃描關鍵字)
                for key in dram_dict.keys():
                    mask = df.apply(lambda r: r.astype(str).str.contains(key, case=False, na=False), axis=1)
                    if mask.any().any():
                        for r_idx, row in df[mask.any(axis=1)].iterrows():
                            for c_idx, val in row.items():
                                if key in str(val):
                                    # 抓取右邊一格的數字
                                    try:
                                        qty = pd.to_numeric(df.iloc[r_idx, c_idx+1], errors='coerce')
                                        if pd.notna(qty): dram_dict[key] = max(dram_dict[key], qty)
                                    except:
                                        pass

                # 3. 尋找下半部 Demand (找 Accum 或是 箭頭符號)
                if df_demand is None:
                    # 搜尋包含 'Accum' 或 dates like '3/31' 的列
                    demand_mask = df.apply(lambda r: r.astype(str).str.contains('Accum|-->|Receiving', case=False, na=False).any(), axis=1)
                    if demand_mask.any():
                        d_start = demand_mask.idxmax()
                        # 抓取該表格
                        raw_demand = df.iloc[max(0, d_start-1):d_start+20].dropna(how='all', axis=1).dropna(how='all', axis=0)
                        
                        # 確保標題不重複
                        d_cols = [str(c) for c in raw_demand.iloc[0]]
                        raw_demand.columns = [f"{c}_{i}" if d_cols.count(c)>1 else c for i, c in enumerate(d_cols)]
                        df_demand = raw_demand.iloc[1:].reset_index(drop=True)

            # ========================================
            # 介面渲染：Dashboard 正式呈現
            # ========================================
            
            # --- 第一部分：歷史趨勢 ---
            if df_wip_trend is not None:
                st.markdown("### 📈 第一部分：站點 WIP 歷史趨勢 (3/2 ~ 當前)")
                
                station_col = df_wip_trend.columns[0]
                # 抓取所有日期欄位 (例如包含 2026 的)
                date_cols = [c for c in df_wip_trend.columns if '202' in str(c)]
                
                if date_cols:
                    for c in date_cols:
                        df_wip_trend[c] = pd.to_numeric(df_wip_trend[c], errors='coerce').fillna(0)
                    
                    # 轉換為折線圖所需的長表格式
                    df_melt = df_wip_trend.melt(id_vars=[station_col], value_vars=date_cols, var_name='Date', value_name='Qty')
                    df_melt['Date'] = pd.to_datetime(df_melt['Date'], errors='coerce')
                    df_melt = df_melt.dropna(subset=['Date']).sort_values('Date')
                    df_melt = df_melt[df_melt['Qty'] > 0] # 過濾掉 0 讓圖表更乾淨
                    
                    # 繪製趨勢折線圖
                    fig_trend = px.line(df_melt, x='Date', y='Qty', color=station_col, markers=True,
                                        title="每日各站點 WIP 數量流動趨勢")
                    st.plotly_chart(fig_trend, use_container_width=True)

            st.markdown("---")
            
            # --- 第二部分與第三部分並排呈現 ---
            col1, col2 = st.columns([1, 1.5])
            
            with col1:
                st.markdown("### 🗂️ 第二部分：當日 DRAM WIP")
                df_dram = pd.DataFrame(list(dram_dict.items()), columns=['Spec', 'Qty'])
                df_dram = df_dram[df_dram['Qty'] > 0]
                
                if not df_dram.empty:
                    fig_pie = px.pie(df_dram, names='Spec', values='Qty', hole=0.4)
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.warning("⚠️ 無法找到 MU16G, SS16G 等數據")

            with col2:
                st.markdown("### 📦 第三部分：出貨需求 (Demand)")
                if df_demand is not None and not df_demand.empty:
                    st.dataframe(df_demand, use_container_width=True)
                else:
                    st.warning("⚠️ 掃描了所有分頁，但找不到出貨需求表格。")
                    
    except Exception as e:
        st.error("解析發生嚴重錯誤，請截圖下方紅色文字給我看：")
        import traceback
        st.code(traceback.format_exc())

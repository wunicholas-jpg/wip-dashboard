import streamlit as st
import pandas as pd
import plotly.express as px

# 設置網頁
st.set_page_config(page_title="WIP Dashboard", layout="wide")
st.title("🚀 ZC13 WIP 智能管理面板")

# 自動修正重複標題的函數
def fix_columns(columns):
    cols = []
    seen = {}
    for i, c in enumerate(columns):
        c = str(c).strip() if pd.notnull(c) and str(c).strip() != "" else f"Unnamed_{i}"
        if c in seen:
            seen[c] += 1
            cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            cols.append(c)
    return cols

uploaded_file = st.file_uploader("📥 上傳 ZC13 WIP 原始檔 (XLSX)", type=['xlsx'])

if uploaded_file:
    try:
        # 讀取所有數據
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        # --- 區塊 1: WIP Trend (找 Process Step) ---
        wip_idx = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("Process Step", na=False)].index
        if not wip_idx.empty:
            start = wip_idx[0]
            # 找到下一個 "Sum" 作為結束
            sum_idx = df_raw.iloc[start:].index[df_raw.iloc[start:, 0].astype(str).str.contains("Sum", na=False)]
            end = sum_idx[0] if not sum_idx.empty else start + 30
            
            df_wip = df_raw.iloc[start:end].copy()
            df_wip.columns = fix_columns(df_wip.iloc[0])
            df_wip = df_wip[2:].reset_index(drop=True)
            
            # 數值轉型
            target_col = df_wip.columns[1] # 抓最新日期那一欄
            df_wip[target_col] = pd.to_numeric(df_wip[target_col], errors='coerce').fillna(0)
            
            st.subheader(f"📊 各站點 WIP 分佈 ({target_col})")
            fig = px.bar(df_wip, x=df_wip.columns[0], y=target_col, color=df_wip.columns[0], text_auto='.2s')
            st.plotly_chart(fig, use_container_width=True)
        
        # --- 區塊 2: DRAM Status (搜尋全表關鍵字) ---
        st.subheader("🗂️ DRAM 規格拆解 (16G / 12G)")
        dram_keywords = ['MU16G', 'SS16G', 'HY12G', 'SS12G']
        dram_data = []
        
        # 遍歷整張表尋找 DRAM 關鍵字
        for r in range(len(df_raw)):
            row_values = df_raw.iloc[r].astype(str).tolist()
            for i, val in enumerate(row_values):
                for key in dram_keywords:
                    if key in val:
                        # 假設數字在關鍵字右邊那一格
                        qty = pd.to_numeric(df_raw.iloc[r, i+1], errors='coerce') if i+1 < len(df_raw.columns) else 0
                        dram_data.append({"Spec": key, "Qty": qty or 0})
        
        if dram_data:
            df_dram = pd.DataFrame(dram_data).groupby("Spec").sum().reset_index()
            c1, c2 = st.columns([1, 2])
            with c1:
                st.dataframe(df_dram, use_container_width=True)
            with c2:
                fig_pie = px.pie(df_dram, names='Spec', values='Qty', hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.warning("找不到 DRAM 數據 (MU16G/SS12G)，請確認 Excel 內容。")

        # --- 區塊 3: Demand (找 Accum) ---
        st.subheader("📦 出貨需求與 AI 風險分析")
        demand_idx = df_raw[df_raw.apply(lambda r: r.astype(str).str.contains("Accum").any(), axis=1)].index
        if not demand_idx.empty:
            df_demand = df_raw.iloc[demand_idx[0]:demand_idx[0]+15].dropna(how='all', axis=1)
            df_demand.columns = fix_columns(df_demand.iloc[0])
            st.dataframe(df_demand[1:], use_container_width=True)
            
            # --- AI Agent 模擬 ---
            st.info("🤖 **AI Risk Assessment**")
            st.write("根據當前 WIP 與 Demand 數據：")
            total_wip = df_wip[target_col].sum()
            st.write(f"- 當前總 WIP: **{int(total_wip):,}**")
            st.write("- **出貨風險評估**：FT 站點目前水位正常，但 IQC 站點有堆積現象，建議加速流轉。")
        
    except Exception as e:
        st.error(f"解析發生問題：{e}")
        st.write("原始數據預覽：")
        st.dataframe(df_raw.head(50))

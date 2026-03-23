import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="ZC13 WIP Dashboard", layout="wide")

st.title("🚀 ATE FT WIP 智能管理面板")
st.markdown("---")

uploaded_file = st.file_uploader("📥 上傳 OSAT WIP 原始檔 (XLSX/CSV)", type=['xlsx', 'csv'])

def clean_column_names(columns):
    """處理重複與格式混亂的標題"""
    return [str(c).split(' ')[0] if pd.notnull(c) else f"Unnamed_{i}" for i, c in enumerate(columns)]

if uploaded_file:
    # 讀取整張表，不設標題
    df_raw = pd.read_excel(uploaded_file, header=None) if uploaded_file.name.endswith('xlsx') else pd.read_csv(uploaded_file, header=None)
    
    # --- 1. 定位 WIP Trend (Top) ---
    # 尋找 "Process Step" 作為開頭
    wip_start_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("Process Step", na=False)].index[0]
    # 假設 WIP 表格到 "Sum" 結束
    wip_end_row = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("Sum", na=False)].index[0]
    
    df_wip = df_raw.iloc[wip_start_row:wip_end_row].dropna(how='all', axis=1)
    df_wip.columns = clean_column_names(df_wip.iloc[0])
    df_wip = df_wip[2:].reset_index(drop=True) # 避開 Process Step 和時間列

    # --- 2. 定位 DRAM Breakdown (Middle) ---
    # 搜尋全表尋找 "D-Ram" 或 "MU16G"
    mask = df_raw.apply(lambda row: row.astype(str).str.contains('D-Ram|MU16G|SS16G', case=False).any(), axis=1)
    dram_indices = df_raw[mask].index
    
    if len(dram_indices) > 0:
        dram_row = dram_indices[0]
        # 抓取 DRAM 所在的區塊 (通常在關鍵字周圍)
        df_dram = df_raw.iloc[dram_row-1 : dram_row+5, :].dropna(how='all', axis=1).dropna(how='all', axis=0)
        df_dram.columns = ["Type", "DRAM_Spec", "WIP_Qty", "Remain", "Other"][:len(df_dram.columns)]
    else:
        df_dram = pd.DataFrame()

    # --- 3. 定位 Demand (Bottom) ---
    demand_indices = df_raw[df_raw.iloc[:, 0].astype(str).str.contains("Receiving|Ship|Demand", na=False)].index
    if len(demand_indices) > 0:
        df_demand = df_raw.iloc[demand_indices[0]:].dropna(how='all', axis=1).dropna(how='all', axis=0)
        df_demand.columns = clean_column_names(df_demand.iloc[0])
        df_demand = df_demand[1:]
    else:
        df_demand = pd.DataFrame()

    # ================= 介面呈現 =================
    
    # 第一區：關鍵數據卡片
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        total_wip = pd.to_numeric(df_wip.iloc[:, 1], errors='coerce').sum()
        st.metric("Total WIP (Current)", f"{int(total_wip):,}")
    with col2:
        # 假設第一站是 IQC
        iqc_qty = pd.to_numeric(df_wip[df_wip.iloc[:,0].str.contains("IQC", na=False)].iloc[:,1], errors='coerce').sum()
        st.metric("IQC Status", f"{int(iqc_qty):,}")
    with col3:
        st.metric("Risk Status", "Normal" if total_wip > 50000 else "Critical", delta_color="inverse")

    st.markdown("### 📊 WIP 站點分佈與趨勢")
    
    # 畫出當前 WIP 分佈圖
    fig_wip = px.bar(df_wip, x=df_wip.columns[0], y=df_wip.columns[1], 
                     title="Current WIP by Station", labels={df_wip.columns[1]: 'Quantity'})
    st.plotly_chart(fig_wip, use_container_width=True)

    c1, c2 = st.columns(2)
    
    with c1:
        st.markdown("### 🗂️ DRAM 規格拆解 (16G/12G)")
        if not df_dram.empty:
            # 整理 DRAM 數據畫圓餅圖
            fig_pie = px.pie(df_dram, names="DRAM_Spec", values="WIP_Qty", hole=0.4)
            st.plotly_chart(fig_pie)
        else:
            st.warning("找不到 DRAM 數據區塊，請確認關鍵字是否為 'D-Ram' 或 'MU16G'")

    with c2:
        st.markdown("### 📦 出貨需求對比")
        if not df_demand.empty:
            st.dataframe(df_demand, use_container_width=True)
        else:
            st.info("尚無出貨需求數據")

    # --- AI Agent 區塊 ---
    st.markdown("---")
    st.markdown("### 🤖 AI Agent 決策助理")
    user_q = st.text_input("您可以問我關於 WIP 的問題：", placeholder="例如：16G 的 WIP 夠應付本週需求嗎？")
    
    if user_q:
        with st.chat_message("assistant"):
            # 這裡先用簡單邏輯模擬，下一步可以串接真正的 Gemini/GPT
            if "16G" in user_q:
                qty_16g = df_dram[df_dram['DRAM_Spec'].str.contains('16G', na=False)]['WIP_Qty'].sum()
                st.write(f"目前的 16G WIP 總數為 {qty_16g:,}。根據出貨需求表格，本週目標為 47,000，目前看來**風險較低**。")
            else:
                st.write("我已經分析了當前 WIP 數據，目前 FT1 站點累積較多，可能是潛在瓶頸。")

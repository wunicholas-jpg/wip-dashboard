import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="WIP Management System", layout="wide")

st.title("🏭 ATE FT WIP 智能管理面板 (ZC13)")

uploaded_file = st.file_uploader("📥 請上傳最新 WIP Excel 檔案", type=['xlsx', 'csv'])

def get_unique_cols(df_cols):
    """強力修正重複標題與 NaN 標題問題"""
    new_cols = []
    seen = {}
    for i, val in enumerate(df_cols):
        val = str(val).strip() if pd.notnull(val) else f"Unnamed_{i}"
        if val in seen:
            seen[val] += 1
            new_cols.append(f"{val}_{seen[val]}")
        else:
            seen[val] = 0
            new_cols.append(val)
    return new_cols

if uploaded_file:
    try:
        # 1. 讀取原始數據 (不設 header)
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)

        # --- 第一區塊：WIP Trend (搜尋 "Process Step") ---
        wip_header_row = df_raw[df_raw.apply(lambda r: r.astype(str).str.contains("Process Step").any(), axis=1)].index
        if not wip_header_row.empty:
            start_idx = wip_header_row[0]
            # 往後抓 30 列左右，直到遇到 "Sum" 或空白
            df_wip = df_raw.iloc[start_idx:start_idx+40].dropna(how='all', axis=1).dropna(how='all', axis=0)
            df_wip.columns = get_unique_cols(df_wip.iloc[0]) # 修正重複標題
            df_wip = df_wip[2:].reset_index(drop=True) # 跳過標題與時間列
            
            # 清洗 WIP 數據，把文字轉數字
            main_col = df_wip.columns[0] # 'Process Step'
            val_col = df_wip.columns[1]  # '2026-03-23' (最新一天)
            df_wip[val_col] = pd.to_numeric(df_wip[val_col], errors='coerce').fillna(0)
            df_wip = df_wip[df_wip[val_col] > 0] # 只看有數值的站點
        else:
            st.error("找不到 'Process Step' 關鍵字，請確認檔案格式")

        # --- 第二區塊：DRAM Status (搜尋 "MU16G/SS12G") ---
        dram_keywords = ['MU16G', 'SS16G', 'HY12G', 'SS12G']
        dram_rows = df_raw[df_raw.apply(lambda r: r.astype(str).str.contains('|'.join(dram_keywords)).any(), axis=1)]
        
        dram_data = []
        for idx, row in dram_rows.iterrows():
            # 邏輯：在該行尋找第一個數值
            row_list = row.tolist()
            for i, val in enumerate(row_list):
                if str(val) in dram_keywords:
                    qty = row_list[i+1] if i+1 < len(row_list) else 0
                    dram_data.append({'Spec': str(val), 'Qty': pd.to_numeric(qty, errors='coerce') or 0})
        
        df_dram = pd.DataFrame(dram_data)

        # --- 第三區塊：Demand (從另一張表或底部搜尋) ---
        # 這裡示範從底部搜尋包含 "Qty" 和 "Accum" 的區塊
        demand_row = df_raw[df_raw.apply(lambda r: r.astype(str).str.contains("Accum").any(), axis=1)].index
        if not demand_row.empty:
            df_demand = df_raw.iloc[demand_row[0]:demand_row[0]+10].dropna(how='all', axis=1)
            df_demand.columns = get_unique_cols(df_demand.iloc[0])
            df_demand = df_demand[1:].reset_index(drop=True)
        else:
            df_demand = pd.DataFrame()

        # ==================== Dashboard Layout ====================
        
        # 1. 指標卡
        m1, m2, m3 = st.columns(3)
        m1.metric("Total WIP", f"{int(df_wip[val_col].sum()):,}")
        m2.metric("16G Total", f"{int(df_dram[df_dram['Spec'].str.contains('16G')]['Qty'].sum()):,}")
        m3.metric("12G Total", f"{int(df_dram[df_dram['Spec'].str.contains('12G')]['Qty'].sum()):,}")

        st.markdown("---")
        
        # 2. 圖表區
        c1, c2 = st.columns([2, 1])
        with c1:
            st.subheader(f"📊 各站點 WIP 分佈 ({val_col})")
            fig_bar = px.bar(df_wip, x=main_col, y=val_col, color=main_col, text_auto='.2s')
            st.plotly_chart(fig_bar, use_container_width=True)
        
        with c2:
            st.subheader("🍩 DRAM 規格佔比")
            if not df_dram.empty:
                fig_pie = px.pie(df_dram, names='Spec', values='Qty', hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)
            else:
                st.warning("無法解析 DRAM 數據區塊")

        # 3. 出貨需求與 AI Agent
        st.markdown("---")
        st.subheader("📦 Shipment Demand & AI Risk Analysis")
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            st.write("最新 Demand 列表:")
            st.dataframe(df_demand, use_container_width=True)
            
        with col_right:
            st.info("🤖 **AI Agent 評估報告**")
            # 這裡串接簡單邏輯，未來可換成真正的 LLM
            current_ft = df_wip[df_wip[main_col].str.contains("FT", na=False)][val_col].sum()
            st.write(f"1. **產能缺口**：當前 FT 站點總計有 {int(current_ft):,} WIP。")
            st.write("2. **風險分析**：對比下週一 50K 的出貨目標，目前缺口約 8K，請確認 IQC/LS 站點流速。")
            if st.button("點擊生成詳細風險報告"):
                st.write("正在分析歷史良率與機台 OEE... (此功能需掛載 AI Agent)")

    except Exception as e:
        st.error(f"解析失敗: {e}")
        st.write("請確認您的 Excel 檔案結構，或檢查第一欄關鍵字。")

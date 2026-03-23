import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="WIP FT Management", layout="wide")

st.title("🏭 ATE FT WIP 智能管理面板 (ZC13)")

# 標準 17 站點順序
TARGET_STATIONS = [
    "Receive from TSMC", "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", 
    "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1(FTA)", "LS4", "Bake", 
    "TR", "FQC", "PACK", "MP ship"
]

def to_num(x):
    try:
        val = str(x).replace(',', '').strip()
        return float(val) if val not in ['nan', '', 'None'] else 0.0
    except:
        return 0.0

uploaded_file = st.file_uploader("📥 請上傳最新 ZC13 WIP 原始檔 (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # =========================================================
        # 1. 第一部分：History_WIP (趨勢與瘦身表格)
        # =========================================================
        if "History_WIP" in xls.sheet_names:
            st.markdown("### 📈 第一部分：站點 WIP 歷史趨勢分析")
            df_hist_raw = pd.read_excel(xls, sheet_name="History_WIP", header=None)
            
            # 抓取日期
            raw_dates = [str(d).split(' ')[0] for d in df_hist_raw.iloc[0, 1:] if pd.notnull(d)]
            
            hist_list = []
            for i in range(len(df_hist_raw)):
                name = str(df_hist_raw.iloc[i, 0]).strip()
                if name in TARGET_STATIONS:
                    vals = [to_num(v) for v in df_hist_raw.iloc[i, 1:len(raw_dates)+1]]
                    hist_list.append([name] + vals)
            
            df_hist = pd.DataFrame(hist_list, columns=["Station"] + raw_dates)
            
            # --- 圖像化改進：增加篩選器避免線條太多 ---
            selected_stations = st.multiselect(
                "選擇要觀察的站點 (預設顯示關鍵測試站):", 
                TARGET_STATIONS, 
                default=["FT CORR", "FT1", "SLT", "PACK"]
            )
            
            df_melt = df_hist.melt(id_vars="Station", var_name="Date", value_name="Qty")
            df_melt["Date"] = pd.to_datetime(df_melt["Date"])
            df_plot = df_melt[df_melt["Station"].isin(selected_stations)]
            
            fig_trend = px.line(df_plot, x="Date", y="Qty", color="Station", markers=True, height=500,
                               title="WIP 歷史流動趨勢 (可自由縮放)")
            st.plotly_chart(fig_trend, use_container_width=True)
            
            # --- 表格縮小：放進 Expander ---
            with st.expander("查看完整 WIP 歷史數據表格 (Raw Data)"):
                st.dataframe(df_hist.style.format(precision=0), use_container_width=True)

        # =========================================================
        # 2. 第二部分：Current_WIP (修正數據抓取錯誤)
        # =========================================================
        st.markdown("---")
        if "Current_WIP" in xls.sheet_names:
            st.markdown("### 🗂️ 第二部分：Current WIP 當前站點分佈")
            df_curr_raw = pd.read_excel(xls, sheet_name="Current_WIP", header=None)
            
            # 定義 DRAM 分類
            specs = ["MU16G", "SS16G", "HY12G", "SS12G"]
            dram_details = []
            
            # 遍歷尋找各規格的數據行
            for i in range(len(df_curr_raw)):
                row_label = str(df_curr_raw.iloc[i, 1]).strip() # Memory Config 欄位
                if row_label in specs:
                    d_type = "16G" if "16G" in row_label else "12G"
                    # 對應站點從 Col 2 (Receiving) 開始，共 17 站 (部分表可能到 Col 18)
                    for j, st_name in enumerate(TARGET_STATIONS[1:]): 
                        qty = to_num(df_curr_raw.iloc[i, j+2])
                        dram_details.append({"DRAM": d_type, "Spec": row_label, "Station": st_name, "Qty": qty})
            
            df_curr = pd.DataFrame(dram_details)
            
            if not df_curr.empty:
                # 圖像化：今日各站點數據
                fig_curr = px.bar(df_curr, x="Station", y="Qty", color="Spec", 
                                 title="今日各站點 16G/12G 實際庫存分佈", text_auto='.2s', height=500)
                st.plotly_chart(fig_curr, use_container_width=True)
            else:
                st.error("無法正確抓取 Current_WIP 數據，請檢查 Memory Config 名稱是否正確。")

        # =========================================================
        # 3. 第三部分：Ship Demand (完整日期與地點顯示)
        # =========================================================
        st.markdown("---")
        if "Ship Demand" in xls.sheet_names:
            st.markdown("### 📦 第三部分：Shipment Demand 出貨排程與風險評估")
            df_dem_raw = pd.read_excel(xls, sheet_name="Ship Demand", header=None)
            
            # 抓取出貨日期 (Row 3, Col 5 以後)
            d_dates = [str(d).split(' ')[0] for d in df_dem_raw.iloc[3, 5:15] if pd.notnull(d)]
            
            demand_data = []
            for i in range(len(df_dem_raw)):
                spec = str(df_dem_raw.iloc[i, 1]).strip()
                place = str(df_dem_raw.iloc[i, 4]).strip() # Shipping Place
                
                if any(x in spec for x in specs):
                    for idx, d_date in enumerate(d_dates):
                        qty = to_num(df_dem_raw.iloc[i, 5+idx])
                        if qty > 0:
                            demand_data.append({"Date": d_date, "Spec": spec, "Place": place, "Qty": qty})
            
            df_demand = pd.DataFrame(demand_data)
            
            if not df_demand.empty:
                col_L, col_R = st.columns([2, 1])
                with col_L:
                    fig_demand = px.bar(df_demand, x="Date", y="Qty", color="Spec", pattern_shape="Place",
                                       title="3/25 ~ 5/07 出貨預測 (by 地點)", barmode="stack", height=500)
                    st.plotly_chart(fig_demand, use_container_width=True)
                
                with col_R:
                    st.info("🤖 **AI Agent 出貨風險診斷**")
                    # 計算近期(第一筆日期)的總需求
                    first_date = df_demand["Date"].iloc[0]
                    short_term_demand = df_demand[df_demand["Date"] == first_date]["Qty"].sum()
                    
                    # 計算現有可出貨 WIP (PACK + MP Ship)
                    ship_ready_wip = df_curr[df_curr["Station"].isin(["PACK", "MP ship"])]["Qty"].sum()
                    
                    st.write(f"- **下次出貨日期**: {first_date}")
                    st.write(f"- **該日總需求量**: {int(short_term_demand):,}")
                    st.write(f"- **成品區現有量**: {int(ship_ready_wip):,}")
                    
                    if ship_ready_wip < short_term_demand:
                        st.error(f"🚨 **Risk**: 成品庫存不足！尚差 {int(short_term_demand - ship_ready_wip):,} 顆。")
                        # 追蹤上一站 FT1
                        ft1_wip = df_curr[df_curr["Station"]=="FT1"]["Qty"].sum()
                        st.warning(f"💡 觀察 FT1 站點尚有 {int(ft1_wip):,} 顆，建議加速流轉至成品區。")
                    else:
                        st.success("✅ **Normal**: 現有水位充足，無立即 Risk。")

    except Exception as e:
        st.error(f"系統解析失敗，請確認檔案格式是否異動。錯誤訊息: {e}")

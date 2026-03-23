import streamlit as st
import pandas as pd
import plotly.express as px

# 網頁寬度全開
st.set_page_config(page_title="WIP Management V4", layout="wide")

st.title("🏭 ZC13 ATE FT 智能管理儀表板 (工控級 V4)")
st.markdown("---")

# 定義 17 個標準站點
TARGET_STATIONS = [
    "Receive from TSMC", "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", 
    "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1(FTA)", "LS4", "Bake", 
    "TR", "FQC", "PACK", "MP ship"
]

def to_num(x):
    try:
        val = str(x).replace(',', '').strip()
        return float(val) if val not in ['nan', '', 'None', '#REF!'] else 0.0
    except:
        return 0.0

uploaded_file = st.file_uploader("📥 請上傳 ZC13 WIP 原始 Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # =========================================================
        # 1. 第一部分：History_WIP (趨勢圖像化改進)
        # =========================================================
        if "History_WIP" in xls.sheet_names:
            st.subheader("📈 第一部分：WIP 歷史趨勢與總量監測")
            df_hist_raw = pd.read_excel(xls, sheet_name="History_WIP", header=None)
            
            # 抓取日期並強制排序
            raw_dates = [str(d).split(' ')[0] for d in df_hist_raw.iloc[0, 1:] if pd.notnull(d)]
            
            hist_list = []
            for i in range(len(df_hist_raw)):
                name = str(df_hist_raw.iloc[i, 0]).strip()
                if name in TARGET_STATIONS:
                    vals = [to_num(df_hist_raw.iloc[i, idx+1]) for idx, d in enumerate(raw_dates)]
                    hist_list.append([name] + vals)
            
            df_hist = pd.DataFrame(hist_list, columns=["Station"] + raw_dates)
            
            # 圖像化：改用堆疊柱狀圖，更直觀看到每日總產能
            df_melt = df_hist.melt(id_vars="Station", var_name="Date", value_name="Qty")
            df_melt = df_melt[df_melt["Qty"] > 0] # 只顯示有數值的
            
            fig_hist = px.bar(df_melt, x="Date", y="Qty", color="Station", 
                             title="每日 WIP 站點分佈 (堆疊圖可看總量與比例)",
                             barmode="stack", height=500, text_auto='.2s')
            fig_hist.update_xaxes(type='category') # 確保所有日期節點都顯示
            st.plotly_chart(fig_hist, use_container_width=True)
            
            with st.expander("📄 查看 History WIP 原始數據表"):
                st.dataframe(df_hist, use_container_width=False)

        # =========================================================
        # 2. 第二部分：Current_WIP (直觀分佈圖)
        # =========================================================
        st.markdown("---")
        if "Current_WIP" in xls.sheet_names:
            st.subheader("🗂️ 第二部分：Current WIP 當前庫存分析")
            df_curr_raw = pd.read_excel(xls, sheet_name="Current_WIP", header=None)
            
            curr_data = []
            # 遍歷抓取 16G/12G 各項數據
            for i in range(len(df_curr_raw)):
                config = str(df_curr_raw.iloc[i, 1]).strip()
                dram_type = str(df_curr_raw.iloc[i, 0]).strip()
                if config in ["MU16G", "SS16G", "HY12G", "SS12G"]:
                    for j, st_name in enumerate(TARGET_STATIONS[1:]): # 從 Receiving 開始
                        qty = to_num(df_curr_raw.iloc[i, j+2])
                        curr_data.append({"DRAM": dram_type.replace(" demand", ""), "Config": config, "Station": st_name, "Qty": qty})
            
            df_curr = pd.DataFrame(curr_data)
            
            c1, c2 = st.columns([1, 2])
            with c1:
                # 使用圓環圖顯示 DRAM 總量佔比
                fig_pie = px.pie(df_curr, values="Qty", names="DRAM", hole=.4, title="16G vs 12G 總佔比")
                st.plotly_chart(fig_pie, use_container_width=True)
            with c2:
                # 使用分組柱狀圖看站點詳情
                fig_curr_bar = px.bar(df_curr, x="Station", y="Qty", color="Config", 
                                     title="當前各站點 Config 分佈", barmode="group", text_auto='.2s')
                st.plotly_chart(fig_curr_bar, use_container_width=True)
            
            with st.expander("📄 查看 Current WIP 原始數據表"):
                st.dataframe(df_curr, use_container_width=True)

        # =========================================================
        # 3. 第三部分：Ship Demand (補齊 FIHVN 資訊)
        # =========================================================
        st.markdown("---")
        if "Ship Demand" in xls.sheet_names:
            st.subheader("📦 第三部分：Shipment Demand 出貨排程 (FIHCN/FIHVN/HKDC)")
            df_dem_raw = pd.read_excel(xls, sheet_name="Ship Demand", header=None)
            
            # 抓取日期 (Row 3, Col 5 以後)
            d_dates = [str(d).split(' ')[0] for d in df_dem_raw.iloc[3, 5:15] if pd.notnull(d)]
            
            demand_rows = []
            for i in range(len(df_dem_raw)):
                config = str(df_dem_raw.iloc[i, 1]).strip()
                place = str(df_dem_raw.iloc[i, 4]).strip() # Shipping place
                
                if place in ["FIHCN", "FIHVN", "HKDC"]:
                    # 向上尋找 DRAM 類型
                    d_type = "16G" if "16" in config else "12G"
                    for idx, d_date in enumerate(d_dates):
                        qty = to_num(df_dem_raw.iloc[i, 5+idx])
                        if qty > 0:
                            demand_rows.append({"Date": d_date, "DRAM": d_type, "Config": config, "Place": place, "Qty": qty})
            
            df_demand = pd.DataFrame(demand_rows)
            
            # 圖像化：依照出貨地 (Place) 著色
            fig_dem = px.bar(df_demand, x="Date", y="Qty", color="Place", facet_col="DRAM",
                            title="未來出貨預測 (By 地點 & DRAM)", barmode="group", text_auto='.2s')
            st.plotly_chart(fig_dem, use_container_width=True)
            
            with st.expander("📄 查看 Ship Demand 原始數據表"):
                st.dataframe(df_demand, use_container_width=True)

            # =========================================================
            # AI Agent：整合風險診斷
            # =========================================================
            st.info("🤖 **AI Agent 出貨風險診斷報告**")
            col_a, col_b = st.columns(2)
            
            with col_a:
                # 診斷 16G
                total_16g_wip = df_curr[df_curr["DRAM"]=="16GB"]["Qty"].sum()
                total_16g_demand = df_demand[df_demand["DRAM"]=="16G"]["Qty"].sum()
                st.write(f"**16G 狀態分析：**")
                st.write(f"- 總庫存: {int(total_16g_wip):,} / 總需求: {int(total_16g_demand):,}")
                if total_16g_wip < total_16g_demand:
                    st.error(f"⚠️ 16G 缺口: {int(total_16g_demand - total_16g_wip):,} 顆，請確認 IQC 水位。")
                else:
                    st.success("✅ 16G 庫存足以覆蓋至 5/07 的所有需求。")
            
            with col_b:
                # 診斷地點風險
                fihvn_demand = df_demand[df_demand["Place"]=="FIHVN"]["Qty"].sum()
                st.write(f"**特定地點分析 (FIHVN)：**")
                st.write(f"- FIHVN 總出貨需求: {int(fihvn_demand):,}")
                if fihvn_demand > 0:
                    st.warning(f"💡 注意：FIHVN 有 {int(fihvn_demand):,} 的出貨排程，請優先檢核該分批的 Label 與 Packing Spec。")

    except Exception as e:
        st.error(f"解析失敗，請確認 Excel 結構。錯誤訊息: {e}")

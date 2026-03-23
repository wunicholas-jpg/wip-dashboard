import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# 基礎網頁設定
st.set_page_config(page_title="WIP FT Management", layout="wide")
st.title("🏭 ATE FT WIP 智能管理面板 (ZC13)")

# 定義 17 個標準站點
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

uploaded_file = st.file_uploader("📥 請上傳重整後的 ZC13 WIP 原始檔 (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # ---------------------------------------------------------
        # 1. 第一部分：History_WIP (趨勢分析)
        # ---------------------------------------------------------
        if "History_WIP" in xls.sheet_names:
            df_hist_raw = pd.read_excel(xls, sheet_name="History_WIP", header=None)
            
            # 抓取日期 (Row 0, 從 Col 1 開始)
            raw_dates = df_hist_raw.iloc[0, 1:].tolist()
            # 格式化日期：移除時間，只保留日期
            clean_dates = []
            for d in raw_dates:
                if pd.notnull(d):
                    dt = pd.to_datetime(d)
                    clean_dates.append(dt.strftime('%Y-%m-%d'))
            
            # 抓取站點數據
            hist_list = []
            for i in range(len(df_hist_raw)):
                st_name = str(df_hist_raw.iloc[i, 0]).strip()
                if st_name in TARGET_STATIONS:
                    vals = [to_num(v) for v in df_hist_raw.iloc[i, 1:len(clean_dates)+1]]
                    hist_list.append([st_name] + vals)
            
            df_hist = pd.DataFrame(hist_list, columns=["Station"] + clean_dates)
            
            st.subheader("📈 第一部分：站點 WIP 歷史紀錄 (日期由新到舊)")
            # 呈現表格 (日期已經是由新到舊)
            st.dataframe(df_hist, use_container_width=True)
            
            # 圖像化：歷史趨勢折線圖
            df_melt = df_hist.melt(id_vars="Station", var_name="Date", value_name="Qty")
            fig_trend = px.line(df_melt, x="Date", y="Qty", color="Station", markers=True,
                               title="各站點 WIP 歷史流動趨勢", height=450)
            st.plotly_chart(fig_trend, use_container_width=True)
        
        # ---------------------------------------------------------
        # 2. 第二部分：Current_WIP (當前分佈)
        # ---------------------------------------------------------
        st.markdown("---")
        if "Current_WIP" in xls.sheet_names:
            df_curr_raw = pd.read_excel(xls, sheet_name="Current_WIP", header=None)
            st.subheader("🗂️ 第二部分：Current WIP 站點狀態 (by DRAM Type)")
            
            # 根據您的 CSV 結構，DRAM 在 Col 0, 數據在特定行
            # 邏輯：尋找包含 '16GB' 與 '12GB' 的總計列 (Row 4 與 Row 7 附近)
            dram_data = []
            for i in range(len(df_curr_raw)):
                label = str(df_curr_raw.iloc[i, 0]).strip()
                if label in ["16GB", "12GB"]:
                    # 對應各站點 (Col 2 到 Col 18 左右)
                    for j, st_name in enumerate(TARGET_STATIONS[1:]): # 跳過第一站
                        qty = to_num(df_curr_raw.iloc[i, j+2])
                        dram_data.append({"DRAM": label, "Station": st_name, "Qty": qty})
            
            df_curr = pd.DataFrame(dram_data)
            
            # 圖像化：分組柱狀圖
            fig_curr = px.bar(df_curr, x="Station", y="Qty", color="DRAM", barmode="group",
                             text_auto='.2s', title="當前 16G/12G 站點負載分佈")
            st.plotly_chart(fig_curr, use_container_width=True)

        # ---------------------------------------------------------
        # 3. 第三部分：Ship Demand (出貨風險分析)
        # ---------------------------------------------------------
        st.markdown("---")
        if "Ship Demand" in xls.sheet_names:
            df_demand_raw = pd.read_excel(xls, sheet_name="Ship Demand", header=None)
            st.subheader("📦 第三部分：Shipment Demand 出貨需求排程")
            
            # 抓取日期標題 (Row 3, Col 5 以後)
            d_dates = [str(d).split(' ')[0] for d in df_demand_raw.iloc[3, 5:12] if pd.notnull(d)]
            
            demand_rows = []
            # 掃描 16G/12G 行
            for i in range(len(df_demand_raw)):
                spec = str(df_demand_raw.iloc[i, 1]).strip() # MU16G, SS16G...
                place = str(df_demand_raw.iloc[i, 4]).strip() # FIHCN...
                if any(x in spec for x in ["16G", "12G"]):
                    d_type = "16G" if "16G" in spec else "12G"
                    for idx, d_date in enumerate(d_dates):
                        qty = to_num(df_demand_raw.iloc[i, 5+idx])
                        if qty > 0:
                            demand_rows.append({"Date": d_date, "Type": d_type, "Spec": spec, "Place": place, "Qty": qty})
            
            df_demand = pd.DataFrame(demand_rows)
            
            col_L, col_R = st.columns([1.5, 1])
            with col_L:
                fig_ship = px.bar(df_demand, x="Date", y="Qty", color="Spec", pattern_shape="Place",
                                 title="3/25 ~ 5/07 出貨預測 (by Shipping Place)")
                st.plotly_chart(fig_ship, use_container_width=True)
            
            with col_R:
                st.info("🤖 **AI Agent 風險評估助理**")
                # 簡單邏輯：抓取 16G 現有的 FT1 + FT CORR 數量
                if not df_curr.empty:
                    ft_wip_16g = df_curr[(df_curr["DRAM"]=="16GB") & (df_curr["Station"].str.contains("FT"))]["Qty"].sum()
                    next_demand_16g = df_demand[df_demand["Type"]=="16G"]["Qty"].iloc[0] if not df_demand.empty else 0
                    
                    st.write(f"1. **今日 16G FT 庫存**: {int(ft_wip_16g):,}")
                    st.write(f"2. **下週出貨需求**: {int(next_demand_16g):,}")
                    
                    if ft_wip_16g < next_demand_16g:
                        st.error("🚨 **Risk detected**: 16G 庫存不足以支應下次出貨！建議優先處理 FT1 站點。")
                    else:
                        st.success("✅ **Normal**: 現有 WIP 水位正常，無立即出貨風險。")
                    
                    st.write("---")
                    # 瓶頸分析
                    bottleneck = df_curr.loc[df_curr['Qty'].idxmax()]
                    st.warning(f"💡 目前最大堆積站點：**{bottleneck['Station']}**")

    except Exception as e:
        st.error(f"解析發生錯誤：{e}")
        st.info("請確保 Excel 內的工作表名稱正確：'History_WIP', 'Current_WIP', 'Ship Demand'")

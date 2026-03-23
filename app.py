import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ZC13 WIP & Demand Dashboard", layout="wide")
st.title("🏭 ZC13 ATE FT WIP 智能管理面板")

# 17 個標準站點定義
TARGET_STATIONS = [
    "Receive from TSMC", "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", 
    "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1(FTA)", "LS4", "Bake", 
    "TR", "FQC", "PACK", "MP ship"
]

def to_num(x):
    try:
        return float(str(x).replace(',', '').strip())
    except:
        return 0.0

uploaded_file = st.file_uploader("📥 請上傳原始 ZC13 WIP (.xlsx) 檔案", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        # 讀取主要 WIP Sheet (假設是第一張)
        df_wip_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
        
        # ---------------------------------------------------------
        # 1. 第一部分：歷史趨勢 (全日期抓取)
        # ---------------------------------------------------------
        st.subheader("📈 第一部分：站點 WIP 歷史趨勢 (3/02 ~ 3/23)")
        
        # 尋找所有日期列 (第一列)
        date_row = df_wip_raw.iloc[0].astype(str)
        date_indices = [i for i, val in enumerate(date_row) if "2026" in val]
        dates = [date_row[i].split(' ')[0] for i in date_indices]
        
        hist_data = []
        for i in range(len(df_wip_raw)):
            st_name = str(df_wip_raw.iloc[i, 0]).strip()
            if st_name in TARGET_STATIONS:
                row_vals = [to_num(df_wip_raw.iloc[i, idx]) for idx in date_indices]
                hist_data.append([st_name] + row_vals)
        
        df_hist = pd.DataFrame(hist_data, columns=["Station"] + dates)
        
        # 繪製折線圖 (按日期排序)
        df_melt = df_hist.melt(id_vars="Station", var_name="Date", value_name="Qty")
        df_melt["Date"] = pd.to_datetime(df_melt["Date"])
        df_melt = df_melt.sort_values("Date")
        
        fig_trend = px.line(df_melt, x="Date", y="Qty", color="Station", markers=True, height=500)
        st.plotly_chart(fig_trend, use_container_width=True)

        # ---------------------------------------------------------
        # 2. 第二部分：今日 WIP 狀態 (by 站點 by DRAM)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader(f"🗂️ 第二部分：今日 (3/23) 各站點數據 (16G vs 12G)")
        
        # 尋找包含 'delta' 的需求分頁
        delta_sheet_name = [s for s in xls.sheet_names if 'delta' in s.lower()][0]
        df_delta = pd.read_excel(xls, sheet_name=delta_sheet_name, header=None)
        
        # 根據 Excel 實際位置: 16G 在左側 (Row 8, Col 1-16), 12G 在右側 (Row 8, Col 31-46)
        dram_compare = []
        for i, st_name in enumerate(TARGET_STATIONS[1:]): # 跳過第一站
            qty_16g = to_num(df_delta.iloc[8, i+1])
            qty_12g = to_num(df_delta.iloc[8, i+31])
            dram_compare.append({"Station": st_name, "DRAM": "16G", "Qty": qty_16g})
            dram_compare.append({"Station": st_name, "DRAM": "12G", "Qty": qty_12g})
            
        df_comp = pd.DataFrame(dram_compare)
        
        # 圖像化各站點 WIP 分佈
        fig_bar = px.bar(df_comp, x="Station", y="Qty", color="DRAM", barmode="group", text_auto='.2s')
        st.plotly_chart(fig_bar, use_container_width=True)

        # ---------------------------------------------------------
        # 3. 第三部分：Ship Demand (3/25 ~ 5/07)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader("📦 第三部分：Shipment Demand by DRAM & Shipping Place")
        
        # 16G Demand: Col 25(Date), 26(Qty), 27(Place/Accum) - 抓取 3/25 之後
        d16 = df_delta.iloc[2:30, [25, 26, 27]].dropna(subset=[25])
        d16.columns = ["Date", "Qty", "Place"]; d16["DRAM"] = "16G"
        
        # 12G Demand: Col 55(Date), 56(Qty), 57(Place/Accum)
        d12 = df_delta.iloc[2:30, [55, 56, 57]].dropna(subset=[55])
        d12.columns = ["Date", "Qty", "Place"]; d12["DRAM"] = "12G"
        
        df_demand = pd.concat([d16, d12])
        df_demand["Qty"] = df_demand["Qty"].apply(to_num)
        df_demand["Date"] = df_demand["Date"].astype(str)
        
        # 過濾掉非日期字串 (如 Accum.)
        df_demand = df_demand[df_demand["Date"].str.contains('202|3/|4/|5/')]
        
        col_L, col_R = st.columns([1.5, 1])
        with col_L:
            st.write("📊 未來出貨需求排程 (by Shipping Place):")
            fig_demand = px.bar(df_demand, x="Date", y="Qty", color="DRAM", 
                                pattern_shape="Place", barmode="stack", title="3/25 ~ 5/07 出貨預測")
            st.plotly_chart(fig_demand, use_container_width=True)
        
        with col_R:
            st.write("📋 需求清單明細:")
            st.dataframe(df_demand, use_container_width=True, height=400)
            
            # --- AI Agent Risk 評估 ---
            st.info("🤖 **AI Risk 分析建議**")
            total_16g_wip = df_comp[df_comp["DRAM"]=="16G"]["Qty"].sum()
            target_16g_ship = df_demand[(df_demand["DRAM"]=="16G")]["Qty"].sum()
            
            st.write(f"1. **當前 16G 總 WIP**: {int(total_16g_wip):,}")
            st.write(f"2. **4月底前 16G 總需求**: {int(target_16g_ship):,}")
            
            if total_16g_wip < target_16g_ship:
                st.error("🚨 **Risk detected**: 16G 庫存不足以支應至五月之出貨，請協調 OSAT 加班。")
            else:
                st.success("✅ **Normal**: 現有 16G 水位足以支應近期出貨。")

    except Exception as e:
        st.error(f"系統解析失敗: {e}")
        st.write("請確保上傳的檔案包含 'Station -ship delta' 分頁且格式符合 2026MP 範本。")

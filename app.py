import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="WIP & Demand Dashboard", layout="wide")
st.title("🏭 ZC13 ATE FT WIP 智能管理面板")

# 17 個標準站點
TARGET_STATIONS = [
    "Receive from TSMC", "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", 
    "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1(FTA)", "LS4", "Bake", 
    "TR", "FQC", "PACK", "MP ship"
]

def to_num(x):
    try:
        val = str(x).replace(',', '').strip()
        return float(val) if val != 'nan' and val != '' else 0.0
    except:
        return 0.0

uploaded_file = st.file_uploader("📥 請上傳 ZC13 WIP 原始檔 (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # ---------------------------------------------------------
        # 第一部分：WIP 歷史趨勢 (由第一張 Sheet 讀取)
        # ---------------------------------------------------------
        df_wip_raw = pd.read_excel(xls, sheet_name=0, header=None)
        
        # 抓取日期欄位 (Row 0) - 自動過濾掉全空欄位
        date_row = df_wip_raw.iloc[0]
        date_indices = [i for i, v in enumerate(date_row) if "2026" in str(v)]
        # 格式化日期：只保留 YYYY-MM-DD
        clean_dates = [str(date_row[i]).split(' ')[0] for i in date_indices]

        hist_data = []
        for i in range(len(df_wip_raw)):
            row_label = str(df_wip_raw.iloc[i, 0]).strip()
            if row_label in TARGET_STATIONS:
                vals = [to_num(df_wip_raw.iloc[i, idx]) for idx in date_indices]
                hist_data.append([row_label] + vals)
        
        df_hist = pd.DataFrame(hist_data, columns=["Station"] + clean_dates)

        st.subheader("📈 第一部分：站點 WIP 歷史趨勢 (3/02 ~ 當前)")
        df_melt = df_hist.melt(id_vars="Station", var_name="Date", value_name="Qty")
        df_melt["Date"] = pd.to_datetime(df_melt["Date"])
        df_melt = df_melt.sort_values("Date")
        
        fig_trend = px.line(df_melt, x="Date", y="Qty", color="Station", markers=True, height=500)
        st.plotly_chart(fig_trend, use_container_width=True)

        # ---------------------------------------------------------
        # 第二部分：今日 WIP 狀態 by DRAM (從第一張表下方抓取)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader(f"🗂️ 第二部分：今日 ({clean_dates[0]}) 各站點數據 (by DRAM)")
        
        dram_keywords = {"MU16G": "16G", "SS16G": "16G", "HY12G": "12G", "SS12G": "12G"}
        dram_compare = []
        
        # 尋找第一張表下方的 DRAM 站點統計區塊
        for r in range(len(df_wip_raw)):
            row_label = str(df_wip_raw.iloc[r, 82]).strip() # 根據檔案，DRAM 字樣通常在 Col 81/82
            if row_label in dram_keywords:
                d_type = dram_keywords[row_label]
                # 抓取該行對應當日 (Col 83/84 附近的數據)
                qty = to_num(df_wip_raw.iloc[r, 83])
                # 我們把這些資訊對應回站點 (此部分需根據您底下的表格結構動態調整)
                dram_compare.append({"Spec": row_label, "DRAM": d_type, "Qty": qty})

        # 如果上方邏輯抓不到，則從全表掃描所有站點 Qty
        fig_bar = px.bar(df_hist, x="Station", y=clean_dates[0], color="Station", text_auto='.2s', title="今日各站點總負載")
        st.plotly_chart(fig_bar, use_container_width=True)

        # ---------------------------------------------------------
        # 第三部分：Ship Demand (由第二張 Sheet 讀取)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader("📦 第三部分：Shipment Demand by DRAM & Shipping Place")
        
        if len(xls.sheet_names) > 1:
            df_demand_raw = pd.read_excel(xls, sheet_name=1, header=None)
            
            # 定義出貨日期欄位 (Row 4, Col 5 以後)
            demand_dates = [str(v).split(' ')[0] for v in df_demand_raw.iloc[4, 5:12]]
            
            demand_rows = []
            for r in range(5, 25): # 掃描出貨清單區域
                spec = str(df_demand_raw.iloc[r, 1]).strip()
                place = str(df_demand_raw.iloc[r, 4]).strip()
                if spec in ["MU16G", "SS16G", "HY12G", "SS12G"]:
                    d_type = "16G" if "16G" in spec else "12G"
                    for i, d_date in enumerate(demand_dates):
                        qty = to_num(df_demand_raw.iloc[r, 5 + i])
                        if qty > 0:
                            demand_rows.append({"Date": d_date, "DRAM": d_type, "Spec": spec, "Place": place, "Qty": qty})
            
            df_demand = pd.DataFrame(demand_rows)
            
            if not df_demand.empty:
                colL, colR = st.columns([2, 1])
                with colL:
                    fig_demand = px.bar(df_demand, x="Date", y="Qty", color="Spec", 
                                        pattern_shape="Place", barmode="stack", title="3/25 ~ 5/07 出貨預測")
                    st.plotly_chart(fig_demand, use_container_width=True)
                with colR:
                    st.write("📋 詳細需求清單:")
                    st.dataframe(df_demand, use_container_width=True, height=400)
                    
                    # --- AI Agent Risk 評估 ---
                    st.info("🤖 **AI Agent 決策建議**")
                    today_wip = df_hist[clean_dates[0]].sum()
                    next_ship = df_demand.groupby("Date")["Qty"].sum().iloc[0]
                    st.write(f"- **今日總庫存**: {int(today_wip):,}")
                    st.write(f"- **下次出貨 (3/25)**: {int(next_ship):,}")
                    if today_wip < next_ship:
                        st.error("🚨 **Risk detected**: 總 WIP 不足以支應下次出貨。")
                    else:
                        st.success("✅ **Normal**: 水位充足。")
        else:
            st.warning("找不到第二張出貨需求分頁 (Ship Demand)")

    except Exception as e:
        st.error(f"解析發生錯誤：{e}")

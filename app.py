import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="ZC13 WIP Dashboard", layout="wide")

st.title("🏭 ATE FT WIP 智能管理系統 (ZC13)")
st.markdown("---")

# 定義 17 個標準站點
TARGET_STATIONS = [
    "Receive from TSMC", "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", 
    "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1(FTA)", "LS4", "Bake", 
    "TR", "FQC", "PACK", "MP ship"
]

# 上傳區
uploaded_wip = st.file_uploader("📥 上傳今日 WIP 檔案 (如: 20260323.csv)", type=['csv', 'xlsx'])
uploaded_demand = st.file_uploader("📥 上傳出貨需求檔案 (如: Station -ship delta.csv)", type=['csv', 'xlsx'])

def clean_val(x):
    try:
        return float(str(x).replace(',', ''))
    except:
        return 0.0

if uploaded_wip and uploaded_demand:
    try:
        # 1. 讀取 WIP 與 趨勢 (Part 1)
        df_wip_raw = pd.read_csv(uploaded_wip, header=None)
        
        # 提取日期 (Row 0) 與 數據 (Row 1-17)
        dates = df_wip_raw.iloc[0, 1:17].tolist() # 3/23 到 3/02
        dates = [str(d).split(' ')[0] for d in dates]
        
        wip_data = []
        for i in range(1, 18):
            st_name = str(df_wip_raw.iloc[i, 0]).strip()
            row_vals = [clean_val(v) for v in df_wip_raw.iloc[i, 1:17].tolist()]
            wip_data.append([st_name] + row_vals)
            
        df_history = pd.DataFrame(wip_data, columns=["Station"] + dates)
        
        # --- Part 1 UI: 歷史趨勢折線圖 ---
        st.subheader("📈 第一部分：站點 WIP 歷史趨勢 (3/02 ~ 3/23)")
        df_melted = df_history.melt(id_vars="Station", var_name="Date", value_name="Qty")
        df_melted["Date"] = pd.to_datetime(df_melted["Date"])
        
        fig_line = px.line(df_melted, x="Date", y="Qty", color="Station", markers=True, 
                          height=500, title="各站點歷史 WIP 流動")
        fig_line.update_layout(xaxis={'categoryorder':'category ascending'})
        st.plotly_chart(fig_line, use_container_width=True)
        
        # --- Part 2: 今日 WIP 狀態 by DRAM ---
        st.markdown("---")
        st.subheader(f"🗂️ 第二部分：今日 ({dates[0]}) 各站點 WIP 狀況 (16G vs 12G)")
        
        # 讀取 Station -ship delta 抓取 DRAM 站點分佈
        # 假設 Row 8 是最新數據
        df_delta = pd.read_csv(uploaded_demand, header=None)
        
        # 16G 站點 Qty (Cols 1-16, Row 8)
        # 12G 站點 Qty (Cols 31-46, Row 8)
        dram_compare = []
        for i, st_name in enumerate(TARGET_STATIONS[1:]): # 跳過第一個
            val_16g = clean_val(df_delta.iloc[8, i+1])
            val_12g = clean_val(df_delta.iloc[8, i+31])
            dram_compare.append({"Station": st_name, "DRAM": "16G", "Qty": val_16g})
            dram_compare.append({"Station": st_name, "DRAM": "12G", "Qty": val_12g})
        
        df_comp = pd.DataFrame(dram_compare)
        fig_bar = px.bar(df_comp, x="Station", y="Qty", color="DRAM", barmode="group",
                         text_auto='.2s', title="今日各站點 16G/12G 分佈")
        st.plotly_chart(fig_bar, use_container_width=True)

        # --- Part 3: 出貨需求 Demand ---
        st.markdown("---")
        st.subheader("📦 第三部分：出貨需求與 AI 風險分析")
        
        # 提取 16G Demand (Cols 25, 26) 與 12G Demand (Cols 55, 56)
        demand_16g = df_delta.iloc[2:6, [25, 26]].dropna()
        demand_16g.columns = ["Date", "Qty"]
        demand_16g["DRAM"] = "16G"
        
        demand_12g = df_delta.iloc[2:6, [55, 56]].dropna()
        demand_12g.columns = ["Date", "Qty"]
        demand_12g["DRAM"] = "12G"
        
        df_demand_total = pd.concat([demand_16g, demand_12g])
        df_demand_total["Qty"] = df_demand_total["Qty"].apply(clean_val)
        
        c1, c2 = st.columns([1, 1])
        with c1:
            st.write("📋 出貨清單 (Upcoming Demand):")
            st.dataframe(df_demand_total, use_container_width=True)
            fig_demand = px.bar(df_demand_total, x="Date", y="Qty", color="DRAM", 
                                barmode="group", title="未來出貨需求排程")
            st.plotly_chart(fig_demand, use_container_width=True)
            
        with c2:
            st.info("🤖 **AI Agent 近況回覆與風險評估**")
            # 簡易風險分析邏輯
            total_16g_wip = df_comp[df_comp["DRAM"] == "16G"]["Qty"].sum()
            next_16g_demand = df_demand_total[df_demand_total["DRAM"]=="16G"]["Qty"].iloc[0]
            
            st.write(f"1. **庫存狀態**：目前 16G 總 WIP 為 **{int(total_16g_wip):,}**。")
            st.write(f"2. **需求比對**：下一個出貨目標為 **{int(next_16g_demand):,}**。")
            
            if total_16g_wip < next_16g_demand:
                st.error("🚨 **Risk!!** 現有 WIP 不足以支應下一次出貨目標。建議檢查 IQC/LS1 站點。")
            else:
                st.success("✅ **Normal**：目前 WIP 水位足以應付近期需求。")
            
            st.write("---")
            st.write("**瓶頸分析**：")
            bottleneck = df_comp.loc[df_comp['Qty'].idxmax()]
            st.warning(f"當前主要 WIP 積壓在站點: **{bottleneck['Station']}** ({int(bottleneck['Qty']):,})")

    except Exception as e:
        st.error(f"系統解析檔案時發生錯誤: {e}")
        st.write("請確保上傳的 CSV 檔案格式與 3/23 及 Station -ship delta 的結構一致。")

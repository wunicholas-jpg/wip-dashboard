import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 基礎設定與配色 ---
st.set_page_config(page_title="KYEC WIP Dashboard V7", layout="wide")

G_BLUE = "#4285F4"
G_RED = "#EA4335"    # 警示用
G_YELLOW = "#FBBC05"
G_GREEN = "#34A853"
G_GRAY = "#70757a"

# 16 個標準站點流向
FLOW_STATIONS = [
    "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", "LS QC2", 
    "SLT", "LS QC3", "FT2 Corr", "FT2 (FTA)", "LS4 QC4", 
    "Bake", "T&R", "FQC", "PACK", "MP Ship"
]

def to_num(x):
    try:
        if pd.isna(x) or str(x).strip() in ['', '#REF!', 'None', 'NaN']: return 0.0
        return float(str(x).replace(',', '').strip())
    except: return 0.0

st.title("🏭 ZC13 ATE FT 智能管理儀表板 (V7)")

# 0) KYEC End-End 流程圖 (專業方塊)
st.markdown("### 🔄 KYEC End-to-End Production Flow")
flow_html = f"""
<div style="display: flex; flex-wrap: wrap; gap: 10px; align-items: center; justify-content: center; padding: 15px; background: #f8f9fa; border-radius: 10px;">
    {"".join([f'<div style="background: {G_BLUE}; color: white; padding: 8px 12px; border-radius: 4px; font-weight: bold; font-size: 12px;">{s}</div>' + (' <span style="color: #666;">➔</span> ' if i < len(FLOW_STATIONS)-1 else '') for i, s in enumerate(FLOW_STATIONS)])}
</div>
"""
st.markdown(flow_html, unsafe_allow_html=True)
st.markdown("---")

uploaded_file = st.file_uploader("📥 上傳最新 ZC13 WIP 原始 Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        dram_types = ["MU16G", "SS16G", "HY12G", "SS12G"]
        color_map = {"MU16G": G_BLUE, "SS16G": G_GREEN, "HY12G": G_YELLOW, "SS12G": G_GRAY}

        # =========================================================
        # 1. 第一部分：History_WIP (僅 Show 7 天)
        # =========================================================
        if "History_WIP" in xls.sheet_names:
            st.subheader("📈 第一部分：WIP 歷史紀錄 (近 7 日監測)")
            df_h = pd.read_excel(xls, sheet_name="History_WIP", header=None)
            h_dates = [str(d).split(' ')[0] for d in df_h.iloc[0, 1:] if pd.notnull(d)]
            
            h_list = []
            for i in range(1, len(df_h)):
                name = str(df_h.iloc[i, 0]).strip()
                if name in FLOW_STATIONS or "TSMC" in name:
                    vals = [to_num(df_h.iloc[i, j+1]) for j in range(len(h_dates))]
                    h_list.append([name, sum(vals)] + vals)
            
            df_hist_full = pd.DataFrame(h_list, columns=["Station", "Total_Sum"] + h_dates)
            
            # 僅選擇最近 7 天顯示於圖表
            recent_7_dates = h_dates[:7]
            df_recent = df_hist_full[["Station"] + recent_7_dates]
            
            top_5 = df_hist_full.sort_values("Total_Sum", ascending=False).head(5)["Station"].tolist()
            selected = st.multiselect("選擇站點 (圖表僅顯示近 7 日數據):", df_hist_full["Station"].unique(), default=top_5)
            
            df_melt = df_recent.melt(id_vars="Station", var_name="Date", value_name="Qty")
            df_plot = df_melt[df_melt["Station"].isin(selected)]
            
            fig_h = px.bar(df_plot, x="Date", y="Qty", color="Station", barmode="group",
                          color_discrete_sequence=[G_BLUE, G_GREEN, G_YELLOW, "#673AB7", "#00ACC1"],
                          height=400, title="近 7 日站點產能變化")
            fig_h.update_xaxes(type='category')
            st.plotly_chart(fig_h, use_container_width=True)
            
            with st.expander("📄 查看完整 History WIP 數據表 (Raw Data)"):
                st.dataframe(df_hist_full.drop(columns="Total_Sum"), use_container_width=True)

        # =========================================================
        # 2. 第二部分：Current_WIP (依 Flow 順序排序)
        # =========================================================
        st.markdown("---")
        if "Current_WIP" in xls.sheet_names:
            st.subheader("🗂️ 第二部分：Current WIP 當前各規格分佈")
            df_c = pd.read_excel(xls, sheet_name="Current_WIP", header=None)
            
            curr_data = []
            for i in range(len(df_c)):
                row_label = str(df_c.iloc[i, 0]).strip()
                if row_label in dram_types:
                    # 依 Flow 順序抓取
                    for j, s_name in enumerate(FLOW_STATIONS):
                        # 搜尋該 Sheet 是否有對應站點的數據 (假設數據列在標題列之後)
                        qty = to_num(df_c.iloc[i, j+1]) # 根據 V6 提供座標
                        curr_data.append({"DRAM Type": row_label, "Station": s_name, "Qty": qty})
            
            df_curr = pd.DataFrame(curr_data)
            # 強制排序 Station 依 FLOW_STATIONS 順序
            df_curr['Station'] = pd.Categorical(df_curr['Station'], categories=FLOW_STATIONS, ordered=True)
            df_curr = df_curr.sort_values('Station')

            fig_c = px.bar(df_curr, x="Station", y="Qty", color="DRAM Type", 
                          color_discrete_map=color_map,
                          title="當前 WIP 水位 (依生產流程排序)", barmode="group", text_auto='.2s')
            st.plotly_chart(fig_c, use_container_width=True)
            
            with st.expander("📄 查看 Current WIP 數據表"):
                st.dataframe(df_curr.pivot(index='DRAM Type', columns='Station', values='Qty'), use_container_width=True)

        # =========================================================
        # 3. 第三部分：Ship Demand (FIHCN/FIHVN 分離, DRAM 多彩)
        # =========================================================
        st.markdown("---")
        if "Ship Demand" in xls.sheet_names:
            st.subheader("📦 第三部分：Shipment Demand 出貨排程預估")
            df_d = pd.read_excel(xls, sheet_name="Ship Demand", header=None)
            d_dates = [str(d).split(' ')[0] for d in df_d.iloc[3, 5:15] if pd.notnull(d)]
            
            d_list = []
            for i in range(len(df_d)):
                spec = str(df_d.iloc[i, 1]).strip() # MU16G...
                place = str(df_d.iloc[i, 4]).strip() # FIHCN/FIHVN
                if spec in dram_types:
                    for idx, d_date in enumerate(d_dates):
                        qty = to_num(df_d.iloc[i, 5+idx])
                        if qty > 0:
                            d_list.append({"Date": d_date, "DRAM Type": spec, "Place": place, "Qty": qty})
            
            df_demand = pd.DataFrame(d_list)
            
            # 分別顯示各 DRAM 的直方圖，並區分 FIHCN/FIHVN
            tabs = st.tabs(dram_types)
            for i, spec in enumerate(dram_types):
                with tabs[i]:
                    df_spec = df_demand[df_demand["DRAM Type"] == spec]
                    if not df_spec.empty:
                        fig_spec = px.bar(df_spec, x="Date", y="Qty", color="Place", barmode="group",
                                         color_discrete_map={"FIHCN": G_BLUE, "FIHVN": G_GREEN, "HKDC": G_YELLOW},
                                         title=f"📊 {spec} 出貨預測 (By Place)", text_auto='.2s')
                        st.plotly_chart(fig_spec, use_container_width=True)
                    else:
                        st.write(f"目前沒有 {spec} 的出貨需求。")

        # =========================================================
        # 4. AI Agent：出貨缺口圖表呈現 (庫存遞減邏輯)
        # =========================================================
        st.markdown("---")
        st.error("🤖 **AI Agent 出貨達成率與缺口分析 (依 PACK 站點計算)**")
        
        if not df_curr.empty and not df_demand.empty:
            # 初始 PACK 庫存
            pack_stock = df_curr[df_curr["Station"] == "PACK"].groupby("DRAM Type")["Qty"].sum().to_dict()
            unique_dates = sorted(df_demand["Date"].unique())
            
            for spec in dram_types:
                st.markdown(f"#### 🔍 分析: {spec}")
                current_bal = pack_stock.get(spec, 0)
                analysis_rows = []
                
                for d_date in unique_dates:
                    d_qty = df_demand[(df_demand["Date"] == d_date) & (df_demand["DRAM Type"] == spec)]["Qty"].sum()
                    if d_qty == 0: continue
                    
                    old_bal = current_bal
                    current_bal -= d_qty
                    
                    # 警訊邏輯
                    status = "✅ OK" if current_bal >= 0 else f"🚨 GAP: {int(abs(current_bal)):,}"
                    
                    analysis_rows.append({
                        "Ship Date": d_date,
                        "Initial PACK Stock": int(old_bal),
                        "Demand Qty": int(d_qty),
                        "Expected End Balance": int(current_bal),
                        "Status": status
                    })
                
                if analysis_rows:
                    res_df = pd.DataFrame(analysis_rows)
                    c1, c2 = st.columns([1.5, 1])
                    with c1:
                        # 缺口圖表：顯示庫存餘額走勢
                        fig_gap = go.Figure()
                        fig_gap.add_trace(go.Scatter(x=res_df["Ship Date"], y=res_df["Expected End Balance"], 
                                                  mode='lines+markers', name='Stock Runway',
                                                  line=dict(color=G_BLUE if current_bal >=0 else G_RED)))
                        fig_gap.add_hline(y=0, line_dash="dash", line_color="black")
                        fig_gap.update_layout(title=f"{spec} 庫存消耗走勢圖", height=300)
                        st.plotly_chart(fig_gap, use_container_width=True)
                    with c2:
                        st.table(res_df.style.applymap(lambda x: f'color: {G_RED}' if '🚨' in str(x) else 'color: black', subset=['Status']))
                else:
                    st.write("該規格無出貨計畫。")

    except Exception as e:
        st.error(f"系統解析異常: {e}")

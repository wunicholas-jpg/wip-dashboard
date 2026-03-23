import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. 網頁基礎設定與 Google 配色 ---
st.set_page_config(page_title="KYEC WIP Dashboard V9", layout="wide")

G_BLUE = "#4285F4"
G_GREEN = "#34A853"
G_YELLOW = "#FBBC05"
G_GRAY = "#70757a"
G_RED = "#EA4335" # 僅用於警告

# 定義 16 個標準站點順序
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

st.title("🏭 ZC13 ATE FT 智能管理儀表板 (V9 終極版)")
st.markdown("---")

# 0) KYEC End-to-End Production Flow (專業方塊圖)
st.markdown("### 🔄 KYEC End-to-End Production Flow")
flow_html = f"""
<div style="display: flex; flex-wrap: wrap; gap: 8px; align-items: center; justify-content: center; padding: 20px; background: #ffffff; border: 1px solid #e0e0e0; border-radius: 8px;">
    {"".join([f'<div style="background: {G_BLUE}; color: white; padding: 10px 15px; border-radius: 5px; font-weight: bold; font-size: 13px; box-shadow: 2px 2px 5px rgba(0,0,0,0.1);">{s}</div>' + (' <b style="color: #4285F4; font-size: 18px;">➔</b> ' if i < len(FLOW_STATIONS)-1 else '') for i, s in enumerate(FLOW_STATIONS)])}
</div>
"""
st.markdown(flow_html, unsafe_allow_html=True)
st.markdown("---")

uploaded_file = st.file_uploader("📥 請上傳最新版本 ZC13 WIP 原始 Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        specs = ["MU16G", "SS16G", "HY12G", "SS12G"]

        # =========================================================
        # 1. 第一部分：History_WIP (僅 Show 7 天)
        # =========================================================
        if "History_WIP" in xls.sheet_names:
            st.subheader("📈 第一部分：WIP 歷史演進 (近 7 日數據)")
            df_h = pd.read_excel(xls, sheet_name="History_WIP", header=None)
            h_dates = [str(d).split(' ')[0] for d in df_h.iloc[0, 1:] if pd.notnull(d)]
            
            h_list = []
            for i in range(1, len(df_h)):
                name = str(df_h.iloc[i, 0]).strip()
                if name in FLOW_STATIONS or "TSMC" in name:
                    vals = [to_num(df_h.iloc[i, j+1]) for j in range(len(h_dates))]
                    h_list.append([name, sum(vals)] + vals)
            
            df_hist_full = pd.DataFrame(h_list, columns=["Station", "Total"] + h_dates)
            recent_dates = h_dates[:7] # 只取最近 7 天
            
            # 預設前五大站點
            top_5 = df_hist_full.sort_values("Total", ascending=False).head(5)["Station"].tolist()
            selected = st.multiselect("選擇觀察站點 (圖表僅顯示近 7 日):", df_hist_full["Station"].unique(), default=top_5)
            
            df_plot = df_hist_full[df_hist_full["Station"].isin(selected)][["Station"] + recent_dates]
            df_melt = df_plot.melt(id_vars="Station", var_name="Date", value_name="Qty")
            
            fig_h = px.bar(df_melt, x="Date", y="Qty", color="Station", barmode="group",
                          color_discrete_sequence=[G_BLUE, G_GREEN, G_YELLOW, G_GRAY, "#9C27B0"],
                          height=400, title="各站點水位歷史變動 (近 7 日)")
            fig_h.update_xaxes(type='category')
            st.plotly_chart(fig_h, use_container_width=True)
            
            with st.expander("📄 查看完整歷史數據 (Raw Data)"):
                st.dataframe(df_hist_full.drop(columns="Total"), use_container_width=True)

        # =========================================================
        # 2. 第二部分：Current_WIP (依 Flow 排序, DRAM Type 修正)
        # =========================================================
        st.markdown("---")
        if "Current_WIP" in xls.sheet_names:
            st.subheader("🗂️ 第二部分：Current WIP 當前各規格分佈")
            # 讀取時不設 header 以便精確解析
            df_c_raw = pd.read_excel(xls, sheet_name="Current_WIP", header=None)
            
            curr_data = []
            for i in range(len(df_c_raw)):
                label = str(df_c_raw.iloc[i, 0]).strip()
                if label in specs:
                    for j, s_name in enumerate(FLOW_STATIONS):
                        # 數據從 Col 1 開始對應 Receiving
                        qty = to_num(df_c_raw.iloc[i, j+1])
                        curr_data.append({"DRAM Type": label, "Station": s_name, "Qty": qty})
            
            df_curr = pd.DataFrame(curr_data)
            df_curr['Station'] = pd.Categorical(df_curr['Station'], categories=FLOW_STATIONS, ordered=True)
            df_curr = df_curr.sort_values('Station')

            fig_c = px.bar(df_curr, x="Station", y="Qty", color="DRAM Type", 
                          color_discrete_map={"MU16G": G_BLUE, "SS16G": G_GREEN, "HY12G": G_YELLOW, "SS12G": G_GRAY},
                          title="當前站點水位 (依生產流程排序)", barmode="group", text_auto='.2s')
            st.plotly_chart(fig_c, use_container_width=True)
            
            with st.expander("📄 查看 Current WIP 數據表"):
                # Pivot 表格
                df_pivot = df_curr.pivot(index='DRAM Type', columns='Station', values='Qty')
                # 重新排序列
                df_pivot = df_pivot[FLOW_STATIONS]
                st.dataframe(df_pivot, use_container_width=True)

        # =========================================================
        # 3. 第三部分：Ship Demand (FIHCN/FIHVN 分離解析)
        # =========================================================
        st.markdown("---")
        if "Ship Demand" in xls.sheet_names:
            st.subheader("📦 第三部分：Shipment Demand 出貨需求 (By Location & DRAM)")
            df_d_raw = pd.read_excel(xls, sheet_name="Ship Demand", header=None)
            
            # 日期在 Row 3, Col 3 以後
            d_dates = [str(d).split(' ')[0] for d in df_d_raw.iloc[3, 3:12] if pd.notnull(d)]
            
            demand_rows = []
            current_spec = ""
            for i in range(4, len(df_d_raw)):
                row_spec = str(df_d_raw.iloc[i, 1]).strip()
                if row_spec in specs:
                    current_spec = row_spec
                
                place = str(df_d_raw.iloc[i, 2]).strip()
                if place in ["FIHCN", "FIHVN", "HKDC"]:
                    for idx, d_date in enumerate(d_dates):
                        qty = to_num(df_d_raw.iloc[i, 3+idx])
                        if qty > 0:
                            demand_rows.append({
                                "Date": d_date, 
                                "DRAM Type": current_spec, 
                                "Place": place, 
                                "Qty": qty
                            })
            
            df_demand = pd.DataFrame(demand_rows)
            
            # 分 Tab 顯示不同 DRAM 的出貨排程
            d_tabs = st.tabs(specs)
            for i, spec in enumerate(specs):
                with d_tabs[i]:
                    df_spec_d = df_demand[df_demand["DRAM Type"] == spec]
                    if not df_spec_d.empty:
                        fig_spec_d = px.bar(df_spec_d, x="Date", y="Qty", color="Place", 
                                           barmode="group", # 分開 FIHCN/FIHVN
                                           color_discrete_map={"FIHCN": G_BLUE, "FIHVN": G_GREEN, "HKDC": G_YELLOW},
                                           title=f"{spec} 各地點出貨需求細節")
                        st.plotly_chart(fig_spec_d, use_container_width=True)
                    else:
                        st.write(f"目前沒有 {spec} 的出貨計畫。")

        # =========================================================
        # 4. AI Agent：出貨缺口視覺化 (餘額消耗邏輯修正)
        # =========================================================
        st.markdown("---")
        st.error("🤖 **AI Agent 出貨缺口診斷 (庫存 Runway 遞減分析)**")
        
        if not df_curr.empty and not df_demand.empty:
            # 取得 PACK 站點初始量
            pack_stock = df_curr[df_curr["Station"] == "PACK"].set_index("DRAM Type")["Qty"].to_dict()
            unique_dates = sorted(df_demand["Date"].unique())
            
            for spec in specs:
                st.markdown(f"#### 🔍 分析對象: {spec}")
                current_runway = pack_stock.get(spec, 0)
                analysis_results = []
                
                for d_date in unique_dates:
                    d_qty = df_demand[(df_demand["Date"] == d_date) & (df_demand["DRAM Type"] == spec)]["Qty"].sum()
                    if d_qty == 0: continue
                    
                    old_bal = current_runway
                    current_runway -= d_qty
                    gap = 0 if current_runway >= 0 else abs(current_runway)
                    
                    status = "✅ 充足" if current_runway >= 0 else f"🚨 GAP: {int(abs(current_runway)):,}"
                    
                    analysis_results.append({
                        "Ship Date": d_date,
                        "Initial Stock": int(old_bal),
                        "Demand Qty": int(d_qty),
                        "End Balance": int(current_runway),
                        "Status": status
                    })
                
                if analysis_results:
                    res_df = pd.DataFrame(analysis_results)
                    c1, c2 = st.columns([2, 1])
                    with c1:
                        # 更好的 UX：使用柱狀圖顯示庫存剩餘量，負值變紅
                        res_df["Color"] = res_df["End Balance"].apply(lambda x: G_RED if x < 0 else G_BLUE)
                        fig_runway = px.bar(res_df, x="Ship Date", y="End Balance", 
                                           title=f"{spec} 庫存水位預測 (負值代表缺料)",
                                           text_auto='.2s')
                        fig_runway.update_traces(marker_color=res_df["Color"])
                        fig_runway.add_hline(y=0, line_dash="dash", line_color="black")
                        st.plotly_chart(fig_runway, use_container_width=True)
                    with c2:
                        def style_gap(val):
                            color = G_RED if '🚨' in str(val) else 'black'
                            return f'color: {color}; font-weight: bold'
                        st.table(res_df.style.applymap(style_gap, subset=['Status']))
                else:
                    st.write(f"無 {spec} 需求分析。")

    except Exception as e:
        st.error(f"解析失敗: {e}")

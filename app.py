import streamlit as st
import pandas as pd
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.graph_objects as go

st.set_page_config(page_title="KYEC WIP Dashboard V6", layout="wide")

# --- Google 配色定義 ---
G_BLUE = "#4285F4"
G_GREEN = "#34A853"
G_YELLOW = "#FBBC05"
G_GRAY = "#70757a"
G_RED = "#EA4335"  # 僅用於警示

st.title("🏭 ZC13 ATE FT 智能管理儀表板 (V6 專業版)")
st.markdown("---")

# 16 個標準站點流程顯示
st.info("🔄 **KYEC E2E Flow:** Receiving ➔ IQC ➔ LS1 QC1 ➔ FT CORR ➔ FT1 ➔ LS QC2 ➔ SLT ➔ LS QC3 ➔ FT2 Corr ➔ EQC1 (FTA) ➔ LS4 ➔ Bake ➔ TR ➔ FQC ➔ PACK ➔ MP Ship")

def to_num(x):
    try:
        if pd.isna(x) or str(x).strip() in ['', '#REF!', 'None', 'NaN']: return 0.0
        return float(str(x).replace(',', '').strip())
    except: return 0.0

uploaded_file = st.file_uploader("📥 請上傳最新版本 ZC13 WIP 原始 Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        target_configs = ["MU16G", "SS16G", "HY12G", "SS12G"]
        stations_list = ["Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1 (FTA)", "LS4", "Bake", "TR", "FQC", "PACK", "MP Ship"]

        # =========================================================
        # 1. 第一部分：History_WIP (分組柱狀圖)
        # =========================================================
        if "History_WIP" in xls.sheet_names:
            st.subheader("📈 第一部分：WIP 歷史紀錄 (日期由新到舊)")
            df_h = pd.read_excel(xls, sheet_name="History_WIP", header=None)
            h_dates = [str(d).split(' ')[0] for d in df_h.iloc[0, 1:] if pd.notnull(d)]
            
            h_list = []
            for i in range(1, len(df_h)):
                name = str(df_h.iloc[i, 0]).strip()
                if name in stations_list or "TSMC" in name:
                    vals = [to_num(df_h.iloc[i, j+1]) for j in range(len(h_dates))]
                    h_list.append([name, sum(vals)] + vals)
            
            df_hist = pd.DataFrame(h_list, columns=["Station", "Total_Sum"] + h_dates)
            top_5 = df_hist.sort_values("Total_Sum", ascending=False).head(5)["Station"].tolist()
            
            selected = st.multiselect("選擇觀察站點 (預設 Top 5):", df_hist["Station"].unique(), default=top_5)
            
            df_melt = df_hist.drop(columns="Total_Sum").melt(id_vars="Station", var_name="Date", value_name="Qty")
            df_plot = df_melt[df_melt["Station"].isin(selected)]
            
            # 使用分組柱狀圖取代折線圖，更清楚
            fig_h = px.bar(df_plot, x="Date", y="Qty", color="Station", barmode="group",
                          color_discrete_sequence=[G_BLUE, G_GREEN, G_YELLOW, G_GRAY, "#673AB7"],
                          height=400, title="各站點歷史 Qty 演進")
            fig_h.update_xaxes(type='category')
            st.plotly_chart(fig_h, use_container_width=True)
            
            with st.expander("📄 查看 History WIP 數據表"):
                st.dataframe(df_hist.drop(columns="Total_Sum"), use_container_width=False)

        # =========================================================
        # 2. 第二部分：Current_WIP (Google 配色分佈)
        # =========================================================
        st.markdown("---")
        if "Current_WIP" in xls.sheet_names:
            st.subheader("🗂️ 第二部分：Current WIP 當前各規格分佈")
            df_c = pd.read_excel(xls, sheet_name="Current_WIP", header=None)
            
            curr_data = []
            for i in range(len(df_c)):
                row_label = str(df_c.iloc[i, 0]).strip()
                if row_label in target_configs:
                    for j, s_name in enumerate(stations_list):
                        qty = to_num(df_c.iloc[i, j+1])
                        curr_data.append({"Spec": row_label, "Station": s_name, "Qty": qty})
            
            df_curr = pd.DataFrame(curr_data)
            
            # 使用 Google 配色
            fig_c = px.bar(df_curr, x="Station", y="Qty", color="Spec", 
                          color_discrete_map={"MU16G": G_BLUE, "SS16G": G_GREEN, "HY12G": G_YELLOW, "SS12G": G_GRAY},
                          title="當前站點水位 (Google 企業配色)", barmode="group", text_auto='.2s')
            st.plotly_chart(fig_c, use_container_width=True)
            
            with st.expander("📄 查看 Current WIP 數據表"):
                st.dataframe(df_curr.pivot(index='Spec', columns='Station', values='Qty'), use_container_width=True)

        # =========================================================
        # 3. 第三部分：Ship Demand (各 DRAM 獨立直方圖, 包含 FIHVN)
        # =========================================================
        st.markdown("---")
        if "Ship Demand" in xls.sheet_names:
            st.subheader("📦 第三部分：Shipment Demand 出貨排程 (各規格獨立視圖)")
            df_d = pd.read_excel(xls, sheet_name="Ship Demand", header=None)
            d_dates = [str(d).split(' ')[0] for d in df_d.iloc[3, 5:15] if pd.notnull(d)]
            
            d_list = []
            for i in range(len(df_d)):
                spec = str(df_d.iloc[i, 1]).strip()
                place = str(df_d.iloc[i, 4]).strip()
                if spec in target_configs:
                    for idx, d_date in enumerate(d_dates):
                        qty = to_num(df_d.iloc[i, 5+idx])
                        if qty > 0:
                            d_list.append({"Date": d_date, "Spec": spec, "Place": place, "Qty": qty})
            
            df_demand = pd.DataFrame(d_list)
            
            # 分別顯示各 DRAM 規格的直方圖 (非堆疊)
            for spec in target_configs:
                df_spec = df_demand[df_demand["Spec"] == spec]
                if not df_spec.empty:
                    fig_spec = px.bar(df_spec, x="Date", y="Qty", color="Place", barmode="group",
                                     color_discrete_map={"FIHCN": G_BLUE, "FIHVN": G_GREEN, "HKDC": G_YELLOW},
                                     title=f"📊 {spec} 出貨排程預測 (含 FIHVN)", text_auto='.2s', height=300)
                    st.plotly_chart(fig_spec, use_container_width=True)

        # =========================================================
        # 4. AI Agent：PACK 庫存消耗邏輯 (動態扣除法)
        # =========================================================
        st.markdown("---")
        st.info("🤖 **AI Agent 出貨缺口診斷報告 (庫存遞減邏輯)**")
        
        if not df_curr.empty and not df_demand.empty:
            # 獲取 PACK 站點初始量
            df_pack_init = df_curr[df_curr["Station"] == "PACK"].groupby("Spec")["Qty"].sum().to_dict()
            
            unique_dates = sorted(df_demand["Date"].unique())
            
            for spec in target_configs:
                st.write(f"🔍 **分析規格: {spec}**")
                current_balance = df_pack_init.get(spec, 0)
                analysis_rows = []
                
                for d_date in unique_dates:
                    demand_qty = df_demand[(df_demand["Date"] == d_date) & (df_demand["Spec"] == spec)]["Qty"].sum()
                    
                    if demand_qty == 0: continue
                    
                    old_balance = current_balance
                    current_balance -= demand_qty
                    gap = 0 if current_balance >= 0 else abs(current_balance)
                    
                    status_text = "✅ 充足" if current_balance >= 0 else f"🚨 缺口: {int(gap):,}"
                    
                    analysis_rows.append({
                        "出貨日期": d_date,
                        "期初 PACK 庫存": int(old_balance),
                        "本次需求量": int(demand_qty),
                        "期末預估餘額": int(current_balance),
                        "狀態評估": status_text
                    })
                
                # 呈現分析表格
                if analysis_rows:
                    res_df = pd.DataFrame(analysis_rows)
                    # 只有警示狀態才用顏色標註
                    def highlight_risk(val):
                        color = G_RED if '🚨' in str(val) else 'black'
                        return f'color: {color}'
                    st.table(res_df.style.applymap(highlight_risk, subset=['狀態評估']))
                else:
                    st.write("該規格無出貨需求。")

    except Exception as e:
        st.error(f"解析失敗，請確認檔案格式。錯誤: {e}")

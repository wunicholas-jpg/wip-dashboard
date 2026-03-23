import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ZC13 WIP Dashboard V5", layout="wide")

st.title("🏭 ZC13 ATE FT 智能管理儀表板 (V5 專業版)")
st.markdown("---")

# 標準 17 站點清單
STATIONS = [
    "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", "LS QC2", "SLT", 
    "LS QC3", "FT2 Corr", "EQC1 (FTA)", "LS4", "Bake", "TR", "FQC", "PACK", "MP Ship"
]

def to_num(x):
    try:
        if pd.isna(x) or str(x).strip() in ['', '#REF!', 'None', 'NaN']: return 0.0
        return float(str(x).replace(',', '').strip())
    except: return 0.0

uploaded_file = st.file_uploader("📥 上傳 ZC13 WIP 原始 Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # =========================================================
        # 1. 第一部分：History_WIP (趨勢與自動篩選)
        # =========================================================
        if "History_WIP" in xls.sheet_names:
            st.subheader("📈 第一部分：WIP 歷史趨勢 (自動聚焦 Top 5)")
            df_h = pd.read_excel(xls, sheet_name="History_WIP", header=None)
            h_dates = [str(d).split(' ')[0] for d in df_h.iloc[0, 1:] if pd.notnull(d)]
            
            h_list = []
            for i in range(1, len(df_h)):
                name = str(df_h.iloc[i, 0]).strip()
                if name in STATIONS or name == "Receive from TSMC":
                    vals = [to_num(df_h.iloc[i, j+1]) for j in range(len(h_dates))]
                    h_list.append([name, sum(vals)] + vals)
            
            df_hist = pd.DataFrame(h_list, columns=["Station", "Total_Sum"] + h_dates)
            
            # 找出前五大佔比站點
            top_5 = df_hist.sort_values("Total_Sum", ascending=False).head(5)["Station"].tolist()
            
            selected = st.multiselect("選擇站點 (預設前五大):", df_hist["Station"].unique(), default=top_5)
            
            df_melt = df_hist.drop(columns="Total_Sum").melt(id_vars="Station", var_name="Date", value_name="Qty")
            df_plot = df_melt[df_melt["Station"].isin(selected)]
            
            fig_h = px.line(df_plot, x="Date", y="Qty", color="Station", markers=True, height=450)
            fig_h.update_xaxes(type='category') # 強制顯示所有日期
            st.plotly_chart(fig_h, use_container_width=True)
            
            with st.expander("📄 查看歷史數據 (瘦身表格)"):
                st.dataframe(df_hist.drop(columns="Total_Sum"), use_container_width=True)

        # =========================================================
        # 2. 第二部分：Current_WIP (精準座標解析)
        # =========================================================
        st.markdown("---")
        if "Current_WIP" in xls.sheet_names:
            st.subheader("🗂️ 第二部分：Current WIP 當前各規格分佈")
            df_c = pd.read_excel(xls, sheet_name="Current_WIP", header=None)
            
            curr_data = []
            target_configs = ["MU16G", "SS16G", "HY12G", "SS12G"]
            
            # 根據您的 CSV 結構，DRAM 名稱在第 0 欄
            for i in range(len(df_c)):
                row_label = str(df_c.iloc[i, 0]).strip()
                if row_label in target_configs:
                    # 站點數據從 Col 1 開始
                    for j, s_name in enumerate(STATIONS):
                        qty = to_num(df_c.iloc[i, j+1])
                        curr_data.append({"Spec": row_label, "Station": s_name, "Qty": qty})
            
            df_curr = pd.DataFrame(curr_data)
            
            # 橫向矩陣圖
            fig_c = px.bar(df_curr, x="Station", y="Qty", color="Spec", 
                          title="今日各站點 16G/12G 實際分布", barmode="group", text_auto='.2s')
            st.plotly_chart(fig_c, use_container_width=True)
            
            with st.expander("📄 查看 Current WIP 原始數據表"):
                st.dataframe(df_curr.pivot(index='Spec', columns='Station', values='Qty'), use_container_width=True)

        # =========================================================
        # 3. 第三部分：Ship Demand (DRAM 分配與視覺化)
        # =========================================================
        st.markdown("---")
        if "Ship Demand" in xls.sheet_names:
            st.subheader("📦 第三部分：Shipment Demand 出貨排程分析")
            df_d = pd.read_excel(xls, sheet_name="Ship Demand", header=None)
            
            # 日期在 Row 3, Col 5 以後
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
            
            # UX 優化：按 Spec 著色
            fig_d = px.bar(df_demand, x="Date", y="Qty", color="Spec", pattern_shape="Place",
                          title="未來出貨預測 (按 DRAM 規格分類)", barmode="stack", text_auto='.2s')
            st.plotly_chart(fig_d, use_container_width=True)
            
            with st.expander("📄 查看 Ship Demand 原始數據表"):
                st.dataframe(df_demand, use_container_width=True)

        # =========================================================
        # 4. AI Agent：PACK 站點與出貨缺口分析
        # =========================================================
        st.markdown("---")
        st.info("🤖 **AI Agent 出貨達成率診斷報告**")
        
        if not df_curr.empty and not df_demand.empty:
            # 抓取 PACK 站點的數據
            df_pack = df_curr[df_curr["Station"] == "PACK"]
            
            # 近期兩次出貨
            upcoming_dates = sorted(df_demand["Date"].unique())[:2]
            
            cols = st.columns(len(upcoming_dates))
            for idx, u_date in enumerate(upcoming_dates):
                with cols[idx]:
                    st.write(f"📅 **目標日期: {u_date}**")
                    target_d = df_demand[df_demand["Date"] == u_date]
                    
                    results = []
                    for spec in target_configs:
                        demand_qty = target_d[target_d["Spec"] == spec]["Qty"].sum()
                        pack_qty = df_pack[df_pack["Spec"] == spec]["Qty"].sum()
                        gap = pack_qty - demand_qty
                        status = "✅ OK" if gap >= 0 else f"🚨 缺口: {int(abs(gap)):,}"
                        results.append({"DRAM Spec": spec, "PACK 庫存": int(pack_qty), "需求量": int(demand_qty), "狀態": status})
                    
                    st.table(pd.DataFrame(results))
            
            st.warning("💡 **AI 建議**：若 PACK 站點不足，請立即追蹤 FT1 與 FQC 站點之進度，確保及時轉入 PACK。")

    except Exception as e:
        st.error(f"系統解析失敗。請確認 Sub-sheet 名稱與結構正確。錯誤訊息: {e}")

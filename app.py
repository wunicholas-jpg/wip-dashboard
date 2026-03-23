import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# --- 1. Page Configuration & Professional Color Palette ---
st.set_page_config(page_title="KYEC WIP E2E Management Dashboard", layout="wide")

G_BLUE = "#4285F4"
G_GREEN = "#34A853"
G_YELLOW = "#FBBC05"
G_GRAY = "#70757a"
G_RED = "#EA4335" # Alert Only

DRAM_COLORS = {"MU16G": G_BLUE, "SS16G": G_GREEN, "HY12G": G_YELLOW, "SS12G": G_GRAY}

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

def format_date(d):
    """Force date string to YYYY-MM-DD format only"""
    try:
        return str(pd.to_datetime(d).date())
    except:
        return str(d).split(' ')[0]

st.title("📊 KYEC WIP E2E Management Dashboard")
st.markdown("---")

# 0) Production Flow Visualization
st.markdown("### 🔄 Production Flow: End-to-End Process")
flow_html = f"""
<div style="display: flex; flex-wrap: wrap; gap: 8px; align-items: center; justify-content: center; padding: 20px; background: #ffffff; border: 1px solid #e0e0e0; border-radius: 8px;">
    {"".join([f'<div style="background: {G_BLUE}; color: white; padding: 10px 15px; border-radius: 5px; font-weight: bold; font-size: 13px; box-shadow: 2px 2px 5px rgba(0,0,0,0.1);">{s}</div>' + (' <b style="color: #4285F4; font-size: 18px;">➔</b> ' if i < len(FLOW_STATIONS)-1 else '') for i, s in enumerate(FLOW_STATIONS)])}
</div>
"""
st.markdown(flow_html, unsafe_allow_html=True)
st.markdown("---")

uploaded_file = st.file_uploader("📥 Upload ZC13 WIP Master File (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        specs = ["MU16G", "SS16G", "HY12G", "SS12G"]

        # =========================================================
        # Part 1: Historical WIP Evolution
        # =========================================================
        if "History_WIP" in xls.sheet_names:
            st.subheader("📈 Part 1: WIP Historical Trends (Last 7 Days)")
            df_h = pd.read_excel(xls, sheet_name="History_WIP", header=None)
            h_dates = [format_date(d) for d in df_h.iloc[0, 1:] if pd.notnull(d)]
            
            h_list = []
            for i in range(1, len(df_h)):
                name = str(df_h.iloc[i, 0]).strip()
                if name in FLOW_STATIONS or "TSMC" in name:
                    vals = [to_num(df_h.iloc[i, j+1]) for j in range(len(h_dates))]
                    h_list.append([name, sum(vals)] + vals)
            
            df_hist_full = pd.DataFrame(h_list, columns=["Station", "Total"] + h_dates)
            recent_dates = h_dates[:7]
            
            top_5 = df_hist_full.sort_values("Total", ascending=False).head(5)["Station"].tolist()
            selected = st.multiselect("Select Stations:", df_hist_full["Station"].unique(), default=top_5)
            
            df_plot = df_hist_full[df_hist_full["Station"].isin(selected)][["Station"] + recent_dates]
            df_melt = df_plot.melt(id_vars="Station", var_name="Date", value_name="Qty")
            
            fig_h = px.bar(df_melt, x="Date", y="Qty", color="Station", barmode="group",
                          color_discrete_sequence=[G_BLUE, G_GREEN, G_YELLOW, G_GRAY, "#9C27B0"],
                          height=400, title="Daily Station Level Changes")
            fig_h.update_xaxes(type='category')
            st.plotly_chart(fig_h, use_container_width=True)

        # =========================================================
        # Part 2: Current WIP Status
        # =========================================================
        st.markdown("---")
        if "Current_WIP" in xls.sheet_names:
            st.subheader("🗂️ Part 2: Current WIP Distribution by DRAM Type")
            df_c_raw = pd.read_excel(xls, sheet_name="Current_WIP", header=None)
            curr_data = []
            for i in range(len(df_c_raw)):
                label = str(df_c_raw.iloc[i, 0]).strip()
                if label in specs:
                    for j, s_name in enumerate(FLOW_STATIONS):
                        qty = to_num(df_c_raw.iloc[i, j+1])
                        curr_data.append({"DRAM Type": label, "Station": s_name, "Qty": qty})
            
            df_curr = pd.DataFrame(curr_data)
            df_curr['Station'] = pd.Categorical(df_curr['Station'], categories=FLOW_STATIONS, ordered=True)
            df_curr = df_curr.sort_values('Station')
            
            fig_c = px.bar(df_curr, x="Station", y="Qty", color="DRAM Type", 
                          color_discrete_map=DRAM_COLORS, barmode="group", text_auto='.2s')
            st.plotly_chart(fig_c, use_container_width=True)

        # =========================================================
        # Part 3: Shipment Requirement
        # =========================================================
        st.markdown("---")
        if "Ship Demand" in xls.sheet_names:
            st.subheader("📦 Part 3: Shipment Requirement (By Location & DRAM Type)")
            df_d_raw = pd.read_excel(xls, sheet_name="Ship Demand", header=None)
            d_dates = [format_date(d) for d in df_d_raw.iloc[3, 3:12] if pd.notnull(d)]
            
            demand_rows = []
            current_spec = ""
            for i in range(4, len(df_d_raw)):
                row_spec = str(df_d_raw.iloc[i, 1]).strip()
                if row_spec in specs: current_spec = row_spec
                place = str(df_d_raw.iloc[i, 2]).strip()
                if place in ["FIHCN", "FIHVN", "HKDC"]:
                    for idx, d_date in enumerate(d_dates):
                        qty = to_num(df_d_raw.iloc[i, 3+idx])
                        if qty > 0:
                            demand_rows.append({"Date": d_date, "DRAM Type": current_spec, "Place": place, "Qty": qty})
            
            df_demand = pd.DataFrame(demand_rows)
            d_tabs = st.tabs(specs)
            for i, spec in enumerate(specs):
                with d_tabs[i]:
                    df_spec_d = df_demand[df_demand["DRAM Type"] == spec]
                    if not df_spec_d.empty:
                        fig_spec_d = px.bar(df_spec_d, x="Date", y="Qty", color="Place", 
                                           barmode="group", text_auto='.3s', # Added Data Labels
                                           color_discrete_map={"FIHCN": G_BLUE, "FIHVN": G_GREEN, "HKDC": G_YELLOW},
                                           title=f"{spec} Shipment Requirement by Location")
                        st.plotly_chart(fig_spec_d, use_container_width=True)

        # =========================================================
        # Part 4: AI Agent Analysis (Inventory Runway Model)
        # =========================================================
        st.markdown("---")
        st.error("🤖 AI Agent: Shipment Gap Analysis (PACK Stock Runway)")
        
        overall_summary = []
        if not df_curr.empty and not df_demand.empty:
            pack_stock = df_curr[df_curr["Station"] == "PACK"].set_index("DRAM Type")["Qty"].to_dict()
            unique_dates = sorted(df_demand["Date"].unique())
            
            for spec in specs:
                st.markdown(f"#### 🔍 Runway Analysis: {spec}")
                current_runway = pack_stock.get(spec, 0)
                analysis_results = []
                for d_date in unique_dates:
                    d_qty = df_demand[(df_demand["Date"] == d_date) & (df_demand["DRAM Type"] == spec)]["Qty"].sum()
                    if d_qty == 0: continue
                    old_bal = current_runway
                    current_runway -= d_qty
                    status = "✅ Sufficient" if current_runway >= 0 else f"🚨 GAP: {int(abs(current_runway)):,}"
                    if current_runway < 0: overall_summary.append(f"CRITICAL: {spec} shortage of {int(abs(current_runway)):,} on {d_date}.")
                    analysis_results.append({"Ship Date": d_date, "Initial Stock": int(old_bal), "Demand Qty": int(d_qty), "End Balance": int(current_runway), "Status": status})
                
                if analysis_results:
                    res_df = pd.DataFrame(analysis_results)
                    c1, c2 = st.columns([2, 1])
                    with c1:
                        res_df["BarColor"] = res_df["End Balance"].apply(lambda x: G_RED if x < 0 else G_BLUE)
                        fig_runway = px.bar(res_df, x="Ship Date", y="End Balance", text_auto='.2s', title=f"{spec} Balance Runway")
                        fig_runway.update_traces(marker_color=res_df["BarColor"])
                        st.plotly_chart(fig_runway, use_container_width=True)
                    with c2:
                        st.table(res_df.drop(columns=["BarColor"]))

            st.info("💡 AI Decision Summary (TL;DR)")
            if overall_summary:
                for msg in overall_summary: st.write(f"• {msg}")
            else: st.success("No immediate shortages detected.")

        # =========================================================
        # NEXT STEP: AI Chat Interrogator
        # =========================================================
        st.markdown("---")
        st.subheader("💬 Interactive AI Data Interrogator")
        user_query = st.text_input("Ask a question about the WIP or Demand data:", placeholder="e.g., Which DRAM has the most severe shortage?")
        
        if user_query:
            with st.chat_message("assistant"):
                # Basic context injection logic
                if "shortage" in user_query.lower() or "risk" in user_query.lower():
                    if overall_summary:
                        st.write(f"Based on the runway analysis, the primary risks are: {', '.join(overall_summary)}. I recommend focusing on moving WIP to the PACK station for these specs.")
                    else:
                        st.write("Current analysis shows no shortages in the PACK station. Your inventory runway is healthy.")
                elif "total" in user_query.lower():
                    total_wip = df_curr["Qty"].sum()
                    st.write(f"The total WIP across all stations is currently {int(total_wip):,}. {df_curr.groupby('DRAM Type')['Qty'].sum().to_dict()} is the breakdown by DRAM.")
                else:
                    st.write("I am monitoring the E2E flow. Your biggest WIP accumulation is currently at the " + str(df_curr.loc[df_curr['Qty'].idxmax(), 'Station']) + " station.")

    except Exception as e:
        st.error(f"Analysis Failed: {e}")

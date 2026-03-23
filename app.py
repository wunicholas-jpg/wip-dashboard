import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="WIP Dashboard", layout="wide")

st.title("🏭 ATE FT WIP 智能管理系統 (ZC13)")
st.info("請直接上傳原始 Excel 檔案，系統將自動解析三個區塊。")

# 17 個標準站點定義
TARGET_STATIONS = [
    "Receive from TSMC", "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", 
    "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1(FTA)", "LS4", "Bake", 
    "TR", "FQC", "PACK", "MP ship"
]

def clean_num(x):
    try:
        return float(str(x).replace(',', '').strip())
    except:
        return 0.0

# --- 唯一上傳按鈕 ---
uploaded_file = st.file_uploader("📥 請上傳 ZC13 WIP 原始 Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        # 使用 ExcelFile 讀取，以便同時操作多個 Sheet
        xls = pd.ExcelFile(uploaded_file)
        
        # ---------------------------------------------------------
        # 第一部分：WIP 歷史趨勢 (由第一個 Sheet 讀取)
        # ---------------------------------------------------------
        df_wip_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
        
        # 抓取日期 (Row 0)
        dates = [str(d).split(' ')[0] for d in df_wip_raw.iloc[0, 1:17].tolist() if pd.notnull(d)]
        
        # 抓取 17 站點數據
        wip_history = []
        for i in range(len(df_wip_raw)):
            row_name = str(df_wip_raw.iloc[i, 0]).strip()
            if row_name in TARGET_STATIONS:
                vals = [clean_num(v) for v in df_wip_raw.iloc[i, 1:len(dates)+1].tolist()]
                wip_history.append([row_name] + vals)
        
        df_hist = pd.DataFrame(wip_history, columns=["Station"] + dates)
        
        st.subheader("📈 第一部分：站點 WIP 歷史趨勢 (3/02 ~ 3/23)")
        # 轉換為圖表格式
        df_melt = df_hist.melt(id_vars="Station", var_name="Date", value_name="Qty")
        df_melt["Date"] = pd.to_datetime(df_melt["Date"])
        
        fig_trend = px.line(df_melt, x="Date", y="Qty", color="Station", markers=True, height=500)
        st.plotly_chart(fig_trend, use_container_width=True)

        # ---------------------------------------------------------
        # 第二部分：今日 WIP 狀態 by DRAM (由 'Station -ship delta' 讀取)
        # ---------------------------------------------------------
        st.markdown("---")
        latest_date = dates[0]
        st.subheader(f"🗂️ 第二部分：今日 ({latest_date}) WIP 狀態 by DRAM (16G vs 12G)")
        
        # 自動尋找包含 delta 的 Sheet
        delta_sheet = [s for s in xls.sheet_names if 'delta' in s.lower()][0]
        df_delta = pd.read_excel(xls, sheet_name=delta_sheet, header=None)
        
        # 依照您的 Excel 座標抓取今日數據 (Row 8)
        # 16G: Col 1~16, 12G: Col 31~46
        dram_comp = []
        for i, st_name in enumerate(TARGET_STATIONS[1:]): # 跳過 TSMC 站
            val_16g = clean_num(df_delta.iloc[8, i+1])
            val_12g = clean_num(df_delta.iloc[8, i+31])
            dram_comp.append({"Station": st_name, "DRAM": "16G", "Qty": val_16g})
            dram_comp.append({"Station": st_name, "DRAM": "12G", "Qty": val_12g})
            
        df_comp = pd.DataFrame(dram_comp)
        fig_bar = px.bar(df_comp, x="Station", y="Qty", color="DRAM", barmode="group", text_auto='.2s')
        st.plotly_chart(fig_bar, use_container_width=True)

        # ---------------------------------------------------------
        # 第三部分：出貨需求 (由 'Station -ship delta' 讀取)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader("📦 第三部分：出貨需求與 AI 風險分析")
        
        # 抓取 Demand (座標為 16G: Col 25,26 / 12G: Col 55,56)
        d16 = df_delta.iloc[2:6, [25, 26]].dropna()
        d16.columns = ["Date", "Qty"]; d16["DRAM"] = "16G"
        d12 = df_delta.iloc[2:6, [55, 56]].dropna()
        d12.columns = ["Date", "Qty"]; d12["DRAM"] = "12G"
        
        df_demand = pd.concat([d16, d12])
        df_demand["Qty"] = df_demand["Qty"].apply(clean_num)
        
        col_L, col_R = st.columns([1, 1])
        with col_L:
            st.write("📋 未來出貨需求排程:")
            fig_demand = px.bar(df_demand, x="Date", y="Qty", color="DRAM", barmode="group")
            st.plotly_chart(fig_demand, use_container_width=True)
            
        with col_R:
            st.info("🤖 **AI Agent 風險評估**")
            # 抓取 FT1 的 16G 今日 WIP 與下一次需求比對
            ft1_16g = df_comp[(df_comp["Station"]=="FT1") & (df_comp["DRAM"]=="16G")]["Qty"].sum()
            target_16g = df_demand[df_demand["DRAM"]=="16G"]["Qty"].iloc[0]
            
            st.write(f"- **16G FT1 庫存**: {int(ft1_16g):,}")
            st.write(f"- **下一次出貨需求**: {int(target_16g):,}")
            
            if ft1_16g < target_16g:
                st.error(f"❌ **Risk**: 16G 在 FT1 的 WIP 不足以應付下次出貨。")
            else:
                st.success(f"✅ **Normal**: 現有庫存足以支應下次出貨。")

    except Exception as e:
        st.error(f"解析發生問題：{e}")
        st.write("請確保上傳的是原始 .xlsx 檔案且包含 WIP 與 Station -ship delta 分頁。")

import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="WIP Dashboard", layout="wide")
st.title("🚀 ATE FT WIP 智能管理面板 (ZC13)")

# --- 定義您要求的 17 個標準站點 ---
TARGET_STATIONS = [
    "Receive from TSMC", "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", 
    "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1(FTA)", "LS4", "Bake", 
    "TR", "FQC", "PACK", "MP ship"
]

uploaded_file = st.file_uploader("📥 請上傳 ZC13 WIP 原始檔 (XLSX)", type=['xlsx'])

if uploaded_file:
    try:
        # 讀取 Excel 的所有 Sheet
        xls = pd.ExcelFile(uploaded_file)
        # 假設第一張 Sheet 是最新的 WIP 資料
        df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
        
        # --- 1. 第一部分：歷史趨勢 (上半部) ---
        st.subheader("📈 第一部分：站點 WIP 歷史趨勢 (左新右舊)")
        
        # 第一列是日期，第二列開始是數據
        dates = df_raw.iloc[0, 1:].tolist()
        # 尋找 Process Step 這一列作為索引
        wip_data_rows = []
        for index, row in df_raw.iterrows():
            station_name = str(row[0]).strip()
            if station_name in TARGET_STATIONS:
                # 抓取該站點對應所有日期的數量
                values = pd.to_numeric(row[1:len(dates)+1], errors='coerce').fillna(0).tolist()
                wip_data_rows.append([station_name] + values)
        
        # 建立歷史趨勢 Dataframe
        # 標題格式：[Station, 2026-03-23, 2026-03-20, ...]
        clean_date_cols = [str(d).split(' ')[0] for d in dates]
        df_history = pd.DataFrame(wip_data_rows, columns=['Station'] + clean_date_cols)
        
        # 呈現表格
        st.write("歷史 WIP 紀錄 (由右至左日期更新):")
        st.dataframe(df_history, use_container_width=True)

        # --- 2. 第二部分：當日狀態 by DRAM (中部) ---
        st.markdown("---")
        st.subheader(f"🗂️ 第二部分：當日 ({clean_date_cols[0]}) WIP 狀態 by DRAM")
        
        # 這裡需要區分 16G 與 12G
        # 邏輯：掃描全表找出 16G/12G 關鍵字所在的列，通常在 WIP 表格下方
        dram_status = []
        for index, row in df_raw.iterrows():
            row_str = str(row[0]).strip()
            # 判斷是哪種規格
            dram_type = ""
            if "MU16" in row_str or "SS16" in row_str: dram_type = "16G"
            elif "HY12" in row_str or "SS12" in row_str: dram_type = "12G"
            
            if dram_type:
                # 這裡假設我們要抓取第一欄(今日)的 WIP
                qty = pd.to_numeric(row[1], errors='coerce') or 0
                dram_status.append({"Spec": row_str, "Type": dram_type, "Today_WIP": qty})
        
        if dram_status:
            df_dram = pd.DataFrame(dram_status)
            c1, c2 = st.columns([1, 1])
            with c1:
                st.write("DRAM WIP 統計表:")
                st.table(df_dram)
            with c2:
                fig_pie = px.pie(df_dram, names='Spec', values='Today_WIP', color='Type', hole=0.4, title="16G vs 12G 佔比")
                st.plotly_chart(fig_pie, use_container_width=True)
        
        # --- 3. 第三部分：出貨需求與風險 (下半部) ---
        st.markdown("---")
        st.subheader("📦 第三部分：出貨需求 (Demand)")
        
        # 尋找名稱包含 "ship" 或 "delta" 的 Sheet
        demand_sheet = [s for s in xls.sheet_names if 'ship' in s.lower() or 'delta' in s.lower()]
        if demand_sheet:
            df_demand_raw = pd.read_excel(xls, sheet_name=demand_sheet[0], header=None)
            # 簡單處理：抓取包含 Qty 或日期格式的區塊
            st.write(f"從分頁 `{demand_sheet[0]}` 提取的需求資訊：")
            st.dataframe(df_demand_raw.dropna(how='all', axis=0).dropna(how='all', axis=1).head(20), use_container_width=True)
            
            # 💡 簡易 Risk 分析
            today_ft_wip = df_history[df_history['Station'].str.contains("FT1|FT CORR", na=False)][clean_date_cols[0]].sum()
            st.info(f"💡 **AI 產線觀測**：今日 FT 總 WIP 為 **{int(today_ft_wip):,}**。請對照下方 Demand 評估出貨 Risk。")
        else:
            st.warning("⚠️ 找不到出貨需求相關的分頁 (Sheet名稱需包含 ship 或 delta)")

    except Exception as e:
        st.error(f"解析發生錯誤：{e}")
        st.info("提示：請確保您的第一張 Sheet 包含站點名稱（如 IQC, FT1 等），且第一列為日期。")

import streamlit as st
import pandas as pd
import plotly.express as px

# 1. 網頁基礎設定
st.set_page_config(page_title="ATE FT WIP Dashboard", layout="wide")
st.title("🏭 ZC13 ATE FT WIP 智能管理面板")

# 定位需要的 17 個標準站點
TARGET_STATIONS = [
    "Receive from TSMC", "Receiving", "IQC", "LS1 QC1", "FT CORR", "FT1", 
    "LS QC2", "SLT", "LS QC3", "FT2 Corr", "EQC1(FTA)", "LS4", "Bake", 
    "TR", "FQC", "PACK", "MP ship"
]

# 2. 檔案上傳
uploaded_file = st.file_uploader("📥 請上傳 ZC13 WIP 原始 Excel 檔案 (.xlsx)", type=['xlsx'])

def fix_col_names(df):
    """處理 Excel 標題重複與時間格式問題"""
    new_cols = []
    for i, col in enumerate(df.iloc[0]):
        c_str = str(col).split(' ')[0] if pd.notnull(col) else f"Empty_{i}"
        new_cols.append(c_str)
    df.columns = new_cols
    return df.iloc[1:].reset_index(drop=True)

if uploaded_file:
    try:
        # 讀取 Excel 所有 Sheet
        xls = pd.ExcelFile(uploaded_file)
        
        # --- 【第一部分：WIP 歷史趨勢】 ---
        # 預設讀取第一張工作表
        df_main = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
        
        st.subheader("📈 第一部分：站點 WIP 歷史趨勢 (左新右舊)")
        
        # 提取日期標題 (從 Row 0 提取)
        date_headers = ["Station"] + [str(d).split(' ')[0] for d in df_main.iloc[0, 1:] if pd.notnull(d)]
        
        # 篩選 17 個站點的資料
        wip_rows = []
        for _, row in df_main.iterrows():
            st_name = str(row[0]).strip()
            if st_name in TARGET_STATIONS:
                # 抓取數值數據 (與日期數量對應)
                vals = row[1:len(date_headers)].tolist()
                wip_rows.append([st_name] + vals)
        
        df_history = pd.DataFrame(wip_rows, columns=date_headers)
        st.dataframe(df_history, use_container_width=True)

        # --- 【第二部分：今日 DRAM 狀態】 ---
        st.markdown("---")
        st.subheader(f"🗂️ 第二部分：當日 ({date_headers[1]}) WIP 狀態 by DRAM")
        
        # 在全表中搜尋 DRAM 關鍵字
        dram_info = []
        dram_mapping = {
            "MU16G": "16G", "SS16G": "16G",
            "HY12G": "12G", "SS12G": "12G"
        }
        
        # 搜尋邏輯：遍歷所有 Cell，找到 MU16G/SS16G/HY12G/SS12G
        for r in range(len(df_main)):
            for c in range(len(df_main.columns)):
                cell_val = str(df_main.iloc[r, c])
                for key, d_type in dram_mapping.items():
                    if key in cell_val:
                        # 抓取右邊一格的數值 (通常是當日 WIP)
                        try:
                            val = pd.to_numeric(df_main.iloc[r, c+1], errors='coerce')
                            if pd.notna(val):
                                dram_info.append({"Spec": key, "Type": d_type, "WIP": val})
                        except: pass

        if dram_info:
            df_dram = pd.DataFrame(dram_info).drop_duplicates(subset=['Spec'])
            c1, c2 = st.columns([1, 1])
            with c1:
                st.table(df_dram)
            with c2:
                fig_dram = px.pie(df_dram, names='Spec', values='WIP', color='Type', 
                                  hole=0.4, title="16G vs 12G 佔比")
                st.plotly_chart(fig_dram, use_container_width=True)
        else:
            st.warning("無法在表中定位 DRAM 數據 (MU16G/SS12G 等)，請確認 Excel 內容。")

        # --- 【第三部分：出貨需求 Demand】 ---
        st.markdown("---")
        st.subheader("📦 第三部分：出貨需求 (Demand Analysis)")
        
        # 尋找名稱包含 "ship" 或 "delta" 的工作表
        demand_sheet_name = [s for s in xls.sheet_names if 'ship' in s.lower() or 'delta' in s.lower()]
        
        if demand_sheet_name:
            df_demand = pd.read_excel(xls, sheet_name=demand_sheet_name[0], header=None)
            # 過濾空白欄位並呈現
            df_demand_clean = df_demand.dropna(how='all', axis=0).dropna(how='all', axis=1)
            st.write(f"數據來源分頁：`{demand_sheet_name[0]}`")
            st.dataframe(df_demand_clean.head(20), use_container_width=True)
            
            # --- AI Risk 分析預留區 ---
            st.info("🤖 **AI Agent 觀察報告**")
            # 簡單邏輯：抓取 FT1 的今日 WIP
            ft1_wip = df_history[df_history['Station'] == 'FT1'][date_headers[1]].values
            ft1_val = ft1_wip[0] if len(ft1_wip)>0 else 0
            st.write(f"1. 目前 **FT1** 站點可用 WIP 為：**{int(ft1_val):,}**。")
            st.write("2. 正在比對下半部 Demand 日期... (若 FT1 數量小於下週一需求，AI 將標示為 Risk)")
        else:
            st.info("找不到獨立的 Demand 分頁，顯示原始表底部內容：")
            st.dataframe(df_main.tail(20), use_container_width=True)

    except Exception as e:
        st.error(f"解析發生錯誤：{e}")

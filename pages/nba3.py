import io
import streamlit as st
import mysql.connector
import pandas as pd

# 設定頁面語系與基本配置
st.set_page_config(page_title="NBA 球員數據查詢系統", layout="wide")

# 1. 從 st.secrets 取得 MySQL 連線資訊
try:
    db_config = st.secrets["mysql"]
except KeyError:
    st.error("請確認已在 `.streamlit/secrets.toml` 中正確設定 [mysql] 相關欄位。")
    st.stop()

# 建立資料庫連線與讀取資料的函式（加上快取避免重複撈取）
@st.cache_data(ttl=600)
def load_nba_data():
    try:
        conn = mysql.connector.connect(
            host=db_config["host"],
            port=int(db_config["port"]),
            user=db_config["user"],
            password=db_config["password"],
            database=db_config["database"]
        )
        
        # 讀取對應的三張表
        query_c = "SELECT * FROM career_summaries"
        query_p = "SELECT * FROM players"
        query_t = "SELECT * FROM teams"
        
        df_c = pd.read_sql(query_c, conn)
        df_p = pd.read_sql(query_p, conn)
        df_t = pd.read_sql(query_t, conn)
        
        conn.close()
        
        # 2. 資料合併 (Merge)
        # c 和 p 用 personId 做 merge 稱為 cp
        df_cp = pd.merge(df_c, df_p, on="personId", how="inner")
        # cp 和 t 用 teamId 做 merge 稱為 cpt
        df_cpt = pd.merge(df_cp, df_t, on="teamId", how="inner")
        
        return df_cpt
    except Exception as e:
        st.error(f"資料庫連線或讀取失敗: {e}")
        return pd.DataFrame()

# 載入核心資料
df_cpt = load_nba_data()

# 欄位中文名稱對照字典
COLUMN_MAPPING = {
    'personId': '球員 ID',
    'tpp': '三分球命中率',
    'ftp': '罰球命中率',
    'fgp': '投籃命中率',
    'ppg': '場均得分',
    'rpg': '場均籃板',
    'apg': '場均助攻',
    'bpg': '場均阻攻',
    'mpg': '場均上場時間',
    'spg': '場均抄截',
    'assists': '助攻總數',
    'blocks': '阻攻總數',
    'steals': '抄截總數',
    'turnovers': '失誤總數',
    'offReb': '進攻籃板',
    'defReb': '防守籃板',
    'totReb': '總籃板',
    'fgm': '投籃命中數',
    'fga': '投籃出手數',
    'tpm': '三分球命中數',
    'tpa': '三分球出手數',
    'fta': '罰球出手數',
    'pFouls': '個人犯規數',
    'points': '總得分',
    'gamesPlayed': '出賽場數',
    'gamesStarted': '先發場數',
    'plusMinus': '正負值(淨勝分)',
    'min': '總上埸分鐘數',
    'dd2': '雙十次數(兩雙)',
    'td3': '大三元次數(三雙)'
}

# 3. 用 Streamlit 建立分頁
tabs = ["球員隊伍數據查詢", "系統說明"]
active_tab = st.sidebar.radio("導覽選單", tabs)

if active_tab == "球員隊伍數據查詢":
    st.title("🏀 NBA 球員與隊伍數據查詢")
    
    if not df_cpt.empty:
        # 檢查欄位中是否存在 'nickname'
        if 'nickname' in df_cpt.columns:
            # 4. 從 cpt 欄位 nickname 蒐集 unique 值成為下拉式選單選項
            unique_nicknames = sorted(df_cpt['nickname'].dropna().unique())
            
            selected_nickname = st.selectbox(
                "請選擇球隊暱稱 (Nickname):", 
                options=unique_nicknames,
                index=0
            )
            
            # 5. 根據選項結果篩選資料
            filtered_df = df_cpt[df_cpt['nickname'] == selected_nickname].copy()
            
            # 6. 欄位名稱做中文轉換
            # 將現有字典中有的欄位替換，其餘欄位保持原樣
            display_df = filtered_df.rename(columns=COLUMN_MAPPING)
            
            st.subheader(f"📊 {selected_nickname} 的篩選結果 (共 {len(display_df)} 筆資料)")
            st.dataframe(display_df, use_container_width=True)
            
            # 7. 下載成 隊伍暱稱.xlsx
            # 使用 io.BytesIO 將 Excel 輸出至記憶體中供使用者下載
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                display_df.to_excel(writer, index=False, sheet_name='Data')
            processed_data = output.getvalue()
            
            st.download_button(
                label=f"📥 下載 Excel 檔案 ({selected_nickname}.xlsx)",
                data=processed_data,
                file_name=f"{selected_nickname}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("合併後的資料表中未包含 'nickname' 欄位，請檢查資料表欄位設計。")
    else:
        st.info("暫無資料，請確認資料庫與連線狀態。")

elif active_tab == "系統說明":
    st.title("ℹ️ 系統說明")
    st.markdown("""
    本系統自動整合 NBA 資料庫中的三張實體表：
    1. `career_summaries` (c) - 球員生涯總計
    2. `players` (p) - 球員基本資料
    3. `teams` (t) - 球隊資料
    
    **運作邏輯：**
    * 系統啟動時會從 `st.secrets` 安全地載入 Aiven Cloud MySQL 的憑證。
    * 透過欄位 `personId` 與 `teamId` 進行內部合併 (Inner Join)。
    * 使用者可透過球隊的 `nickname` 進行即時數據篩選。
    * 提供中文欄位對照轉換，並支援一鍵導出 Excel (.xlsx) 功能。
    """)

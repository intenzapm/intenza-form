# app.py
# ------------------------------------------------------------
# 安裝：pip install streamlit pandas gspread google-auth gspread-dataframe xlsxwriter plotly

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime
from io import BytesIO
import plotly.express as px
import xlsxwriter  # 這其實不是直接import而是安裝套件


# ===== 基本設定 =====
SHEET_ID = '14cBokPO-kPKSG0hR254ApacCWN74Eus-HSd1ScfOfO4'
RESPONSES_SHEET = 'Responses'
MACHINES_SHEET  = 'Machines'
QUESTIONS_SHEET = 'Questions'
SCOPE = ['https://www.googleapis.com/auth/spreadsheets']

st.set_page_config(layout='wide')
st.markdown("<h1 style='text-align: center; color: #4CAF50;'>INTENZA 人因評估系統（設定表版）</h1>", unsafe_allow_html=True)

# ===== 初始化 Google Sheet 客戶端 =====
try:
    credentials = Credentials.from_service_account_info(st.secrets['gcp_service_account'], scopes=SCOPE)
    gc = gspread.authorize(credentials)
    sh = gc.open_by_key(SHEET_ID)
except Exception as e:
    st.error(f"❌ 無法連線到 Google Sheet：{e}")
    st.stop()

# ===== 工具函式 =====
def _get_or_create_worksheet(sheet, title, rows=1000, cols=26):
    try:
        return sheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return sheet.add_worksheet(title=title, rows=rows, cols=cols)
    
def create_excel(df_input: pd.DataFrame, sheet_name='資料'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        (df_input if not df_input.empty else pd.DataFrame()).to_excel(writer, index=False, sheet_name=sheet_name)
        workbook  = writer.book
        worksheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#4CAF50', 
            'font_color': 'white', 
            'align': 'center'
        })
        if not df_input.empty:
            for col_num, value in enumerate(df_input.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)
        worksheet.freeze_panes(1, 0)
    output.seek(0)
    return output

def _safe_get_all_records(ws):
    """
    安全讀取成 DataFrame；如果沒有欄位名稱，自動補預設欄位。
    """
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    default_header = ['測試者', '機器代碼', '區塊', '項目', 'Pass/NG', 'Note', '分數', '日期時間']

    header = values[0]
    # 判斷第一列是否像欄位名稱
    if not any(col in header for col in ['項目', '區塊', '分數']):
        # 沒有明顯欄位名稱 → 補預設
        return pd.DataFrame(values, columns=default_header[:len(values[0])])
    else:
        return pd.DataFrame(values[1:], columns=header)

def _ensure_columns(df: pd.DataFrame, required: dict, warn_prefix: str = ''):
    """
    確保 df 擁有 required 指定欄位；缺欄位會補上並警告，但不報錯。
    required: {col_name: default_value}
    """
    if df is None or df.empty:
        df = pd.DataFrame(columns=list(required.keys()))
    for col, default_val in required.items():
        if col not in df.columns:
            if warn_prefix:
                st.warning(f"⚠️ {warn_prefix} 缺少「{col}」欄位，已使用預設值。")
            df[col] = default_val
    df = df.fillna('')
    return df


@st.cache_data(ttl=300, show_spinner=False)
def load_settings():
    ws_m = _get_or_create_worksheet(sh, MACHINES_SHEET)
    ws_q = _get_or_create_worksheet(sh, QUESTIONS_SHEET)

    machines_df  = _safe_get_all_records(ws_m)
    questions_df = _safe_get_all_records(ws_q)

    machines_df = _ensure_columns(machines_df, {'系列名稱': '未分類', '機器代碼': ''})
    machines_df = machines_df[machines_df['機器代碼'].astype(str).str.strip() != '']

    questions_df = _ensure_columns(questions_df, {'區塊分類': '未分類', '問題內容': '', '適用機器代碼': ''})
    questions_df = questions_df[questions_df['問題內容'].astype(str).str.strip() != '']

    if not machines_df.empty:
        series_list = machines_df['系列名稱'].unique().tolist()
        machine_dict = {s: machines_df[machines_df['系列名稱'] == s]['機器代碼'].tolist() for s in series_list}
    else:
        series_list = []
        machine_dict = {}

    return machines_df.reset_index(drop=True), questions_df.reset_index(drop=True), series_list, machine_dict

def get_questions_for_machine(questions_df: pd.DataFrame, current_machine: str):
    sections = {}
    if questions_df is None or questions_df.empty:
        return sections
    qdf = questions_df.copy()
    mask = (qdf['適用機器代碼'].astype(str).str.strip() == '') | \
           (qdf['適用機器代碼'].astype(str).str.strip() == str(current_machine))
    qdf = qdf[mask]
    for sec in qdf['區塊分類'].unique():
        items = qdf[qdf['區塊分類'] == sec]['問題內容'].tolist()
        if items:
            sections[sec] = items
    return sections

@st.cache_data(ttl=60, show_spinner=False)
def load_responses_df():
    ws_resp = _get_or_create_worksheet(sh, RESPONSES_SHEET)
    df = _safe_get_all_records(ws_resp)
    df = _ensure_columns(df, {
        '測試者': '', '機器代碼': '', '區塊': '', '項目': '',
        'Pass/NG': '', 'Note': '', '分數': '', '日期時間': ''
    })
    return df

def compute_series_progress(responses_df, tester_name, selected_series, machine_dict):
    in_series = machine_dict.get(selected_series, []) if selected_series else []
    if not in_series:
        return 0, 0, [], []
    if responses_df.empty:
        return 0, len(in_series), [], in_series
    df = responses_df.copy()
    if tester_name:
        df = df[df['測試者'] == tester_name]
    completed_machines = df[df['項目'] == '整體評分']['機器代碼'].astype(str).str.strip().unique().tolist()
    done_list = [m for m in in_series if m in completed_machines]
    remaining_list = [m for m in in_series if m not in done_list]
    return len(done_list), len(in_series), done_list, remaining_list

# ===== 載入設定 =====
c1, c2 = st.sidebar.columns(2)
if c1.button('🔄 重新載入設定'):
    load_settings.clear()
machines_df, questions_df, series_list, machine_dict = load_settings()

# ===== 初始化 session state =====
if 'records' not in st.session_state:
    st.session_state.records = []
if 'current_machine_index' not in st.session_state:
    st.session_state.current_machine_index = 0
if 'tester_name' not in st.session_state:
    st.session_state.tester_name = ''
if 'selected_series' not in st.session_state:
    st.session_state.selected_series = None

# ===== 模式選擇 =====
app_mode = st.sidebar.selectbox('選擇功能', ['表單填寫工具', '分析工具'])
fill_mode = st.sidebar.radio('填寫模式', ['逐台模式', '自由切換模式'], index=0)

series_options = ['<未選擇>'] + series_list
selected_series_sidebar = st.sidebar.selectbox('系列', series_options)

if selected_series_sidebar != st.session_state.selected_series:
    st.session_state.current_machine_index = 0
    st.session_state.selected_series = selected_series_sidebar if selected_series_sidebar != '<未選擇>' else None

if fill_mode == '逐台模式':
    codes = machine_dict.get(st.session_state.selected_series, [])
    current_machine = codes[st.session_state.current_machine_index] if codes and st.session_state.selected_series else None
else:
    selected_machine_sidebar = st.sidebar.selectbox('機器', ['<未選擇>'] + machine_dict.get(selected_series_sidebar, []))
    current_machine = selected_machine_sidebar if selected_machine_sidebar != '<未選擇>' else None

responses_df = load_responses_df()

# ===== 表單填寫工具 =====
if app_mode == '表單填寫工具':
    if not st.session_state.tester_name:
        name_input = st.text_input('請輸入測試者姓名')
        if st.button('✅ 確認姓名'):
            if name_input.strip():
                st.session_state.tester_name = name_input.strip()
                st.rerun()
        st.stop()

    if not current_machine:
        st.info("ℹ️ 請先選擇系列與機器。")
        st.stop()

    EVALUATION_SECTIONS = get_questions_for_machine(questions_df, current_machine)
    data_list = []
    date_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    overall_questions = EVALUATION_SECTIONS.pop('整體評估', None)

    for section, items in EVALUATION_SECTIONS.items():
        st.subheader(f'🔹 {section}')
        for item in items:
            key_result = f'{section}_{item}_result'
            col1, col2 = st.columns(2)
            with col1:
                if st.button('✅ Pass', key=f'{section}_{item}_pass'):
                    st.session_state[key_result] = 'Pass'
            with col2:
                if st.button('❌ NG', key=f'{section}_{item}_ng'):
                    st.session_state[key_result] = 'NG'
            current_selection = st.session_state.get(key_result)
            note = st.text_input(f'{item} Note', key=f'{section}_{item}_note', value='')
            data_list.append({
                '測試者': st.session_state.tester_name,
                '機器代碼': current_machine,
                '區塊': section,
                '項目': item,
                'Pass/NG': current_selection if current_selection else '未選擇',
                'Note': note,
                '分數': None,
                '日期時間': date_str
            })

    if overall_questions and any('整體評分' in str(q) for q in overall_questions):
        st.subheader("🔹 整體評估")
        score = st.radio('⭐ 整體評分（1~5分）', [1, 2, 3, 4, 5], index=2)
        data_list.append({
            '測試者': st.session_state.tester_name,
            '機器代碼': current_machine,
            '區塊': '整體評估',
            '項目': '整體評分',
            'Pass/NG': 'N/A',
            'Note': '',
            '分數': int(score),
            '日期時間': date_str
        })

    if st.button('✅ 完成本機台並儲存'):
        ws_resp = _get_or_create_worksheet(sh, RESPONSES_SHEET)
        existing_values = ws_resp.get_all_values()
        have_header = len(existing_values) >= 1

        df = pd.DataFrame(data_list)
        if not have_header:
            ws_resp.update([df.columns.tolist()] + df.values.tolist())
        else:
            set_with_dataframe(ws_resp, df, row=len(existing_values)+1, include_index=False, include_column_header=False)

        load_responses_df.clear()

        if fill_mode == '逐台模式':
            codes = machine_dict.get(st.session_state.selected_series, [])
            if current_machine in codes:
                idx = codes.index(current_machine)
                st.session_state.current_machine_index = idx + 1
            st.success("✅ 已儲存，進入下一台..." if st.session_state.current_machine_index < len(codes) else "🎉 本系列已完成！")
        else:
            st.success("✅ 已儲存，可自由切換其他機台。")
        st.rerun()


# ===== 分析工具 =====
elif app_mode == '分析工具':
    ws_resp = _get_or_create_worksheet(sh, RESPONSES_SHEET)
    all_data = _safe_get_all_records(ws_resp)

    if all_data.empty:
        st.warning("⚠️ 尚無資料可分析。")
        st.stop()

    # 確保欄位存在
    for col in ['Pass/NG','分數','項目','區塊','機器代碼','日期時間','測試者','Note']:
        if col not in all_data.columns:
            all_data[col] = ''

    # 整體評分資料準備
    score_data = all_data[all_data['項目'] == '整體評分'].copy()
    score_data['分數'] = score_data['分數'].astype(str).str.strip()
    score_data['整體評分'] = pd.to_numeric(score_data['分數'], errors='coerce')

    # NG 統計
    ng_data = all_data[all_data['Pass/NG'] == 'NG'].copy()

    # 機器清單與區塊順序
    if machines_df.empty:
        MACHINE_CODES_ALL = sorted(all_data['機器代碼'].unique().tolist())
    else:
        MACHINE_CODES_ALL = machines_df['機器代碼'].unique().tolist()

    if questions_df.empty:
        SECTION_ORDER = sorted(all_data['區塊'].unique().tolist())
    else:
        SECTION_ORDER = questions_df['區塊分類'].unique().tolist()
        if '整體評估' not in SECTION_ORDER:
            SECTION_ORDER = list(SECTION_ORDER) + ['整體評估']

    # 通過率及總體評分統計
    summary_list = []
    for machine in MACHINE_CODES_ALL:
        machine_df = all_data[all_data['機器代碼'] == machine]
        for section in SECTION_ORDER:
            sec_df = machine_df[machine_df['區塊'] == section]
            if sec_df.empty:
                continue
            pass_count = (sec_df['Pass/NG'] == 'Pass').sum()
            ng_count = (sec_df['Pass/NG'] == 'NG').sum()
            total = pass_count + ng_count
            pass_rate = (pass_count / total * 100) if total > 0 else None

            summary_list.append({
                '區塊': section,
                '項目': '通過率 (%)',
                machine: f"{pass_rate:.1f}%" if pass_rate is not None else 'N/A'
            })

        avg_score = score_data[score_data['機器代碼'] == machine]['整體評分'].mean()
        summary_list.append({
            '區塊': '整體評估',
            '項目': '總體評分',
            machine: f"{avg_score:.1f}" if pd.notna(avg_score) else 'N/A'
        })

    # NG 排行
    if not ng_data.empty:
        ng_summary = ng_data.groupby(['機器代碼', '區塊', '項目']).size().reset_index(name='NG次數')
        for machine in MACHINE_CODES_ALL:
            machine_ng = ng_summary[ng_summary['機器代碼'] == machine].sort_values('NG次數', ascending=False)
            for _, row in machine_ng.iterrows():
                summary_list.append({
                    '區塊': f"NG：{row['區塊']}",
                    '項目': row['項目'],
                    machine: f"{row['NG次數']} 次"
                })

    summary_df = pd.DataFrame(summary_list) if summary_list else pd.DataFrame(columns=['區塊','項目'] + MACHINE_CODES_ALL)

    for machine in MACHINE_CODES_ALL:
        if machine not in summary_df.columns:
            summary_df[machine] = None

    final_df = summary_df.pivot_table(index=['區塊', '項目'], values=MACHINE_CODES_ALL, aggfunc='first').reset_index()

    # 顯示分析結果
    st.markdown("### 📊 分析結果預覽")
    st.dataframe(final_df)

    # 總體評分排行榜
    if not score_data.empty:
        avg_scores = score_data.groupby('機器代碼')['整體評分'].mean().reset_index()
        fig_score = px.bar(
            avg_scores,
            x='機器代碼',
            y='整體評分',
            title='⭐ 總體評分排行榜',
            text='整體評分',
            color='整體評分',
            color_continuous_scale=['red', 'yellow', 'green']
        )
        fig_score.update_traces(textposition='outside', textfont_size=14)
        fig_score.update_layout(height=500, bargap=0.2)
        st.plotly_chart(fig_score)
    else:
        st.info("ℹ️ 目前沒有『整體評分』資料。")

    # NG 條形圖
    if not ng_data.empty:
        ng_notes = ng_data.copy()

        if 'Note' not in ng_notes.columns:
            ng_notes['Note'] = ''

        ng_notes = ng_notes.merge(
            all_data[['機器代碼', '項目', 'Note']].drop_duplicates(),
            on=['機器代碼', '項目'],
            how='left',
            suffixes=('', '_from_all_data')
        )
        ng_notes['Note'] = ng_notes['Note'].fillna('')
        ng_notes.loc[ng_notes['Note_from_all_data'].notna(), 'Note'] = ng_notes['Note_from_all_data']
        ng_notes = ng_notes.drop(columns=['Note_from_all_data'])

        ng_notes['項目_型號'] = ng_notes['項目'] + '｜' + ng_notes['機器代碼']

        ng_agg = ng_notes.groupby('項目_型號').agg({
            'Pass/NG': 'count',
            'Note': lambda x: '; '.join(sorted(set([str(v) for v in x if str(v).strip() != ''])))
        }).reset_index().rename(columns={'Pass/NG': 'NG次數'})

        ng_agg['備註長度'] = ng_agg['Note'].apply(lambda x: len(x or ''))
        ng_agg = ng_agg.sort_values(['NG次數', '備註長度'], ascending=[False, False])
        category_order = ng_agg['項目_型號'].tolist()

        fig_ng = px.bar(
            ng_agg,
            x='NG次數',
            y='項目_型號',
            orientation='h',
            title='❌ 所有 NG 項目（項目｜機器代碼，合併備註）',
            color='NG次數',
            color_continuous_scale='Reds',
            custom_data=['Note']
        )
        fig_ng.update_traces(hovertemplate='%{y}<br>NG次數: %{x}<br>備註: %{customdata[0]}', text=None)
        fig_ng.update_layout(
            yaxis=dict(categoryorder='array', categoryarray=category_order, automargin=True),
            margin=dict(l=300, r=20, t=50, b=50),
            height=600
        )
        st.plotly_chart(fig_ng)

    # 下載 Excel
    st.sidebar.download_button(
        '📥 下載分析報告 Excel',
        create_excel(final_df, sheet_name='分析報告'),
        file_name=f'分析報告_INTEZA_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'
    )
    st.sidebar.download_button(
        '📥 下載全部 Responses',
        create_excel(all_data, sheet_name='Responses'),
        file_name=f'全部資料_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'
    )

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from datetime import datetime
from io import BytesIO
import plotly.express as px
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import seaborn as sns


# ===== 基本設定 =====
SHEET_ID = '14cBokPO-kPKSG0hR254ApacCWN74Eus-HSd1ScfOfO4'
SHEET_NAME = '工作表1'
ANALYSIS_SHEET_NAME = '分析報告'
scope = ['https://www.googleapis.com/auth/spreadsheets']

# ===== 初始化 Google Sheet 客戶端 =====
credentials = Credentials.from_service_account_info(st.secrets['gcp_service_account'], scopes=scope)
gc = gspread.authorize(credentials)
sh = gc.open_by_key(SHEET_ID)
worksheet = sh.worksheet(SHEET_NAME)

ZL_MACHINES = ['ZL-01', 'ZL-02', 'ZL-03', 'ZL-04', 'ZL-05', 'ZL-07', 'ZL-08', 'ZL-09', 'ZL-10', 'ZL-11']
DL_MACHINES = ['DL-03', 'DL-04', 'DL-05', 'DL-10', 'DL-13']

FIBO_QUESTIONS = {
    'DL-03': ['覺得整體重量會太輕嗎？'],
    'DL-04': ['覺得輕的好還是重的好？'],
    'ZL-01': ['座椅目前夠低嗎？'],
    'ZL-02': ['椅背會太低嗎？'],
    'ZL-07': ['腰帶會很不舒服嗎？'],
    'ZL-08': ['會覺得很難上機嗎？'],
    'ZL-09': ['壓腿滾筒會不會太硬很不舒服？']
}

EVALUATION_SECTIONS = {
    '觸感體驗': ['座位調整重量片是否方便？', '整體動作是否穩定有質感？', '承靠部位是否舒適？', '抓握部分是否符合手感？'],
    '人因調整': ['把手調整是否容易？', '承靠墊位置是否符合需求？', '坐墊位置是否調整方便？', '握把／踏板位置與角度是否符合需求？', '使用時關節是否可對齊軸點？'],
    '力線評估': ['起始重量是否恰當？', '動作過程中重量變化是否流暢？'],
    '運動軌跡': ['是否能完成全行程訓練？', '關節活動角度是否自然？', '運動軌跡是否能完全刺激目標肌群？'],
    '心理感受': ['使用後的滿意度如何？', '是否有願意推薦給他人的意願？'],
    '價值感受': ['你認為我們品牌在傳遞什麼形象？', '你估算這台機器價值多少？']
}

st.set_page_config(layout='wide')
st.markdown("<h1 style='text-align: center; color: #4CAF50;'>INTENZA 人因評估系統</h1>", unsafe_allow_html=True)

app_mode = st.sidebar.selectbox('選擇功能', ['表單填寫工具', '分析工具'])

# 初始化 session state
if 'records' not in st.session_state:
    st.session_state.records = []
if 'current_machine_index' not in st.session_state:
    st.session_state.current_machine_index = 0
if 'tester_name' not in st.session_state:
    st.session_state.tester_name = ''
if 'selected_series' not in st.session_state:
    st.session_state.selected_series = None

MACHINE_CODES = []
current_machine = None
if st.session_state.selected_series:
    MACHINE_CODES = ZL_MACHINES if st.session_state.selected_series == 'ZL 系列' else DL_MACHINES
    if st.session_state.current_machine_index < len(MACHINE_CODES):
        current_machine = MACHINE_CODES[st.session_state.current_machine_index]
        
    # 👉 這段是我們新增的
    selected_machine = st.sidebar.selectbox('📍 手動選擇要填寫的機器（可選）', ['<不選擇>'] + MACHINE_CODES)
    if selected_machine != '<不選擇>':
        current_machine = selected_machine


st.sidebar.success(f"✅ 目前測試者姓名：{st.session_state.tester_name or '未輸入'}")
if current_machine:
    st.sidebar.info(f"🚀 目前進行機台：{current_machine}")

# 顯示系列完成度
zl_completed = len([m for m in set([r['機器代碼'] for r in st.session_state.records]) if m in ZL_MACHINES])
dl_completed = len([m for m in set([r['機器代碼'] for r in st.session_state.records]) if m in DL_MACHINES])

st.sidebar.write(f"📊 ZL 系列完成度：{zl_completed} / {len(ZL_MACHINES)}")
st.sidebar.write(f"📊 DL 系列完成度：{dl_completed} / {len(DL_MACHINES)}")

# 下載 Google Sheet 今天資料
try:
    all_data = pd.DataFrame(worksheet.get_all_records())
    all_data['日期時間'] = pd.to_datetime(all_data['日期時間'], errors='coerce')

    def create_all_data_excel(df_input):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_input.to_excel(writer, index=False, sheet_name='全部資料')
            workbook = writer.book
            worksheet_xl = writer.sheets['全部資料']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#4CAF50', 'font_color': 'white', 'align': 'center'})
            for col_num, value in enumerate(df_input.columns.values):
                worksheet_xl.write(0, col_num, value, header_format)
                worksheet_xl.set_column(col_num, col_num, 20)
            worksheet_xl.freeze_panes(1, 0)
        output.seek(0)
        return output

    st.sidebar.download_button(
        '📥 下載全部資料 (Google Sheet)',
        create_all_data_excel(all_data),
        file_name=f'全部資料_{datetime.now().strftime("%Y%m%d")}.xlsx'
    )
except Exception:
    st.sidebar.write('Google Sheet 尚無資料或讀取失敗')


# 下載 Session 資料
if st.session_state.records:
    df_session = pd.DataFrame(st.session_state.records)

    def create_session_excel(df_input):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_input.to_excel(writer, index=False, sheet_name='Session資料')
            workbook = writer.book
            worksheet_xl = writer.sheets['Session資料']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#4CAF50', 'font_color': 'white', 'align': 'center'})
            for col_num, value in enumerate(df_input.columns.values):
                worksheet_xl.write(0, col_num, value, header_format)
                worksheet_xl.set_column(col_num, col_num, 20)
            worksheet_xl.freeze_panes(1, 0)
        output.seek(0)
        return output

    st.sidebar.download_button(
        '💾 下載目前測試者資料 (Session)',
        create_session_excel(df_session),
        file_name=f'Session資料_{st.session_state.tester_name}_{datetime.now().strftime("%Y%m%d")}.xlsx'
    )
else:
    st.sidebar.write('目前沒有 Session 資料可下載')

if app_mode == '表單填寫工具':
    all_machines = ZL_MACHINES + DL_MACHINES
    completed_machines = sorted(set([r['機器代碼'] for r in st.session_state.records]), key=lambda x: all_machines.index(x))

    if st.session_state.tester_name == '':
        tester_input = st.text_input('請輸入測試者姓名')
        if st.button('✅ 確認提交姓名'):
            if tester_input.strip() != '':
                st.session_state.tester_name = tester_input.strip()
                st.rerun()
            else:
                st.warning('請先輸入姓名再提交')
        st.stop()
    else:
        if st.button('🔄 重新輸入姓名'):
            st.session_state.tester_name = ''
            st.session_state.selected_series = None
            st.session_state.current_machine_index = 0
            st.rerun()

    if st.session_state.selected_series is None:
        series_choice = st.radio('請選擇要開始的系列', ['ZL 系列', 'DL 系列'])
        if st.button('✅ 確認系列'):
            st.session_state.selected_series = series_choice
            st.session_state.current_machine_index = 0
            st.rerun()
        st.stop()

    if current_machine is None:
        st.success(f'🎉 {st.session_state.selected_series} 填寫完成！請至側邊欄下載資料或選擇另一系列繼續填寫')
        if st.sidebar.button('🔄 切換系列／重新開始'):
            st.session_state.selected_series = None
            st.session_state.current_machine_index = 0
            st.rerun()

    else:
        data_list = []
        date_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        for section, items in EVALUATION_SECTIONS.items():
            st.subheader(f'🔹 {section}')
            section_notes = []

            for item in items:
                key_result = f'{section}_{item}_result'

                # 這裡不再主動設定 st.session_state[key_result] = None
                st.markdown(f"**{item}**")
                col1, col2 = st.columns(2)
                with col1:
                    if st.button('✅ Pass', key=f'{section}_{item}_pass'):
                        st.session_state[key_result] = 'Pass'
                with col2:
                    if st.button('❌ NG', key=f'{section}_{item}_ng'):
                        st.session_state[key_result] = 'NG'

                current_selection = st.session_state.get(key_result)
                if current_selection:
                    st.write(f"👉 已選擇：**{current_selection}**")

                note = st.text_input(f'{item} Note', key=f'{section}_{item}_note', value='')
                if note.strip() != '':
                    section_notes.append(f'{item}: {note}')

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

            combined_note = '; '.join(section_notes)
            summary_note = st.text_area(
                f'💬 {section} 區塊總結 Note（以下為細項 Note 整理供參考）\n{combined_note}',
                key=f'{section}_summary_note',
                value=''
            )
            data_list.append({
                '測試者': st.session_state.tester_name,
                '機器代碼': current_machine,
                '區塊': section,
                '項目': '區塊總結 Note',
                'Pass/NG': 'N/A',
                'Note': summary_note,
                '分數': None,
                '日期時間': date_str
            })

        if current_machine in FIBO_QUESTIONS:
            st.subheader('🔹 Fibo問題追蹤')
            for item in FIBO_QUESTIONS[current_machine]:
                display_item = f'{item} （Fibo問題）'
                key_result = f'Fibo_{item}_result'

                # 這裡同樣不主動設定 None
                st.markdown(f"**{display_item}**")
                col1, col2 = st.columns(2)
                with col1:
                    if st.button('✅ Pass', key=f'Fibo_{item}_pass'):
                        st.session_state[key_result] = 'Pass'
                with col2:
                    if st.button('❌ NG', key=f'Fibo_{item}_ng'):
                        st.session_state[key_result] = 'NG'

                current_selection = st.session_state.get(key_result)
                if current_selection:
                    st.write(f"👉 已選擇：**{current_selection}**")

                note = st.text_input(f'{display_item} Note', key=f'Fibo_{item}_note', value='')
                data_list.append({
                    '測試者': st.session_state.tester_name,
                    '機器代碼': current_machine,
                    '區塊': 'Fibo問題追蹤',
                    '項目': display_item,
                    'Pass/NG': current_selection if current_selection else '未選擇',
                    'Note': note,
                    '分數': None,
                    '日期時間': date_str
                })

        score = st.radio('⭐ 整體評分（1~5分）', [1, 2, 3, 4, 5], index=2)
        data_list.append({
            '測試者': st.session_state.tester_name,
            '機器代碼': current_machine,
            '區塊': '整體評估',
            '項目': '整體評分',
            'Pass/NG': 'N/A',
            'Note': '',
            '分數': score,
            '日期時間': date_str
        })

        if st.button('✅ 完成本機台並儲存，進入下一台'):
            st.session_state.records.extend(data_list)
            df = pd.DataFrame(data_list)
            existing_rows = len(worksheet.get_all_values())
            set_with_dataframe(
                worksheet,
                df,
                row=existing_rows + 1,
                include_index=False,
                include_column_header=False
            )

            # 強化版清理：只要 key 名含有 _result、_note、_summary_note 就刪掉
            for key in list(st.session_state.keys()):
                if '_result' in key or '_note' in key or '_summary_note' in key:
                    del st.session_state[key]

            st.session_state.current_machine_index += 1
            st.success("已儲存到 Google Sheet，正在切換到下一台...")
            st.rerun()

elif app_mode == '分析工具':
    try:
        raw_values = worksheet.get_all_values()
        if not raw_values:
            st.warning("⚠️ Google Sheet 尚無資料可分析。")
            st.stop()
        header = raw_values[0]
        if '' in header:
            header = [col if col != '' else f'Unnamed_{i}' for i, col in enumerate(header)]
        if len(header) != len(set(header)):
            st.error(f"❌ Google Sheet header 有重複值：{header}")
            st.stop()
        all_data = pd.DataFrame(raw_values[1:], columns=header)
    except Exception as e:
        st.error(f"❌ Google Sheet 讀取失敗：{e}")
        st.stop()

    st.success(f"✅ 從 Google Sheet 讀取 {len(all_data)} 筆資料！")

    df = all_data.copy()
    ng_data = df[df['Pass/NG'] == 'NG']
    score_data = df[df['項目'] == '整體評分'].copy()
    score_data['整體評分'] = pd.to_numeric(score_data['分數'], errors='coerce')

    summary_list = []
    SECTION_ORDER = list(EVALUATION_SECTIONS.keys()) + ['Fibo問題追蹤', '整體評估']
    MACHINE_CODES_ALL = ZL_MACHINES + DL_MACHINES

    for machine in MACHINE_CODES_ALL:
        machine_df = df[df['機器代碼'] == machine]
        for section in SECTION_ORDER:
            sec_df = machine_df[machine_df['區塊'] == section]
            if sec_df.empty:
                continue
            pass_count = (sec_df['Pass/NG'] == 'Pass').sum()
            ng_count = (sec_df['Pass/NG'] == 'NG').sum()
            total = pass_count + ng_count
            pass_rate = (pass_count / total * 100) if total > 0 else None
            notes = sec_df[(sec_df['項目'] == '區塊總結 Note') & (sec_df['Note'] != '')]
            combined_notes = '; '.join([f"{n}（{t}）" for n, t in zip(notes['Note'], notes['測試者'])])

            summary_list.append({
                '區塊': section,
                '項目': '通過率 (%)',
                machine: f"{pass_rate:.1f}%" if pass_rate is not None else 'N/A'
            })
            summary_list.append({
                '區塊': section,
                '項目': '區塊總結 Note',
                machine: combined_notes if combined_notes else '無'
            })

        avg_score = score_data[score_data['機器代碼'] == machine]['整體評分'].mean()
        summary_list.append({
            '區塊': '整體評估',
            '項目': '總體評分',
            machine: f"{avg_score:.1f}" if not pd.isna(avg_score) else 'N/A'
        })

    ng_summary = ng_data.groupby(['機器代碼', '區塊', '項目']).size().reset_index(name='NG次數')
    for machine in MACHINE_CODES_ALL:
        machine_ng = ng_summary[ng_summary['機器代碼'] == machine].sort_values('NG次數', ascending=False)
        for _, row in machine_ng.iterrows():
            summary_list.append({
                '區塊': f"NG：{row['區塊']}",
                '項目': row['項目'],
                machine: f"{row['NG次數']} 次"
            })

    summary_df = pd.DataFrame(summary_list)
    for machine in MACHINE_CODES_ALL:
        if machine not in summary_df.columns:
            summary_df[machine] = None

    final_df = summary_df.pivot_table(index=['區塊', '項目'], values=MACHINE_CODES_ALL, aggfunc='first').reset_index()
    ng_sections = sorted([s for s in final_df['區塊'].unique() if s.startswith('NG：')])
    section_order_full = SECTION_ORDER + ng_sections
    final_df['區塊'] = pd.Categorical(final_df['區塊'], categories=section_order_full, ordered=True)
    final_df = final_df.sort_values(['區塊', '項目']).reset_index(drop=True)

    # ========== 視覺化部分 ==========
    st.markdown("### 📊 分析結果預覽")
    st.dataframe(final_df)

    # 總體評分排行榜（紅到綠）
    avg_scores = score_data.groupby('機器代碼')['整體評分'].mean().reset_index()
    
    fig_score = px.bar(
        avg_scores,
        x='機器代碼',
        y='整體評分',
        title='⭐ 總體評分排行榜',
        text='整體評分',  # 要顯示的數值
        color='整體評分',
        color_continuous_scale=['red', 'yellow', 'green']
    )
    
    # 文字放到柱狀圖上方 + 放大字體
    fig_score.update_traces(
        textposition='outside',
        textfont_size=14
    )
    
    # 如果項目多，可以適當調整寬度、間距
    fig_score.update_layout(
        height=500,
        bargap=0.2
    )
    
    st.plotly_chart(fig_score)

    # 準備 ng_notes（保證前面有 ng_summary 和 df）
    ng_notes = ng_summary.merge(
        df[['機器代碼', '項目', 'Note']].drop_duplicates(),
        on=['機器代碼', '項目'],
        how='left'
    )
    
    # 把順序改為：項目｜機器代碼
    ng_notes['項目_型號'] = ng_notes['項目'] + '｜' + ng_notes['機器代碼']
    
    # 合併備註
    ng_agg = ng_notes.groupby('項目_型號').agg({
        'NG次數': 'sum',
        'Note': lambda x: '; '.join(sorted(set(x.dropna().unique()))) if x.notna().any() else ''
    }).reset_index()
    
    # 計算備註長度（作為次要排序依據）
    ng_agg['備註長度'] = ng_agg['Note'].apply(lambda x: len(x))
    
    # 按 NG次數 大到小，再按 備註長度 大到小排序
    ng_agg = ng_agg.sort_values(['NG次數', '備註長度'], ascending=[False, False])
    
    # ✅ 保留原順序，不反轉（最多的會在 plotly 的最下方）
    category_order = ng_agg['項目_型號'].tolist()
    
    # 畫圖
    fig_ng = px.bar(
        ng_agg,
        x='NG次數',
        y='項目_型號',
        orientation='h',
        title='❌ 所有 NG 項目（項目｜機器代碼，合併備註，按秩序排序）',
        color='NG次數',
        color_continuous_scale='Reds',
        custom_data=['Note']
    )
    
    # 設定 hover 內容（圖上不顯示數值）
    fig_ng.update_traces(
        hovertemplate='%{y}<br>NG次數: %{x}<br>備註: %{customdata[0]}',
        text=None
    )
    
    # 設定 y 軸順序、左側留空間
    fig_ng.update_layout(
        yaxis=dict(
            categoryorder='array',
            categoryarray=category_order,
            automargin=True
        ),
        margin=dict(l=300, r=20, t=50, b=50),
        height=600
    )
    
    # 顯示圖表
    st.plotly_chart(fig_ng)





    # 下載分析報告 Excel
    def create_analysis_excel(df_input):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_input.to_excel(writer, index=False, sheet_name='分析報告')
            workbook = writer.book
            worksheet = writer.sheets['分析報告']
            header_format = workbook.add_format({
                'bold': True, 'bg_color': '#4CAF50',
                'font_color': 'white', 'align': 'center'
            })
            for col_num, value in enumerate(df_input.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)
            worksheet.freeze_panes(1, 0)
        output.seek(0)
        return output

    st.sidebar.download_button(
        '📥 下載分析報告 Excel',
        create_analysis_excel(final_df),
        file_name=f'分析報告_INTEZA_{pd.Timestamp.now().strftime("%Y%m%d")}.xlsx'
    )

try:
    test_values = worksheet.get_all_values()
    st.success("✅ 成功連線到 Google Sheet！共有 {} 筆資料".format(len(test_values)))
except Exception as e:
    st.error(f"❌ 無法連線到 Google Sheet，錯誤訊息：{e}")


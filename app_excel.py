import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from data_manager_excel import load_data, save_data, get_kpi_metrics

# Page Config
st.set_page_config(
    page_title="Personal Expense Tracker (Local Excel)",
    page_icon="💰",
    layout="wide"
)

# ----------------- CUSTOM CSS -----------------
# 1. Bigger Metrics (40px)
# 2. Adjust Label size (18px)
st.markdown("""
<style>
    [data-testid="stMetricValue"] {
        font-size: 40px !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    /* Enlarge Data Editor Toolbar Buttons - Multiple Approaches */
    [data-testid="stDataFrame"] button[kind="header"],
    button[data-testid="stBaseButton-headerNoPadding"],
    div[data-testid="stElementToolbar"] button,
    div[data-testid="stElementToolbarButton"] button {
        transform: scale(1.8) !important;
        transform-origin: center !important;
        margin: 0.3rem !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- LOAD DATA (Cached) -----------------
@st.cache_data
def get_data_cached():
    return load_data()

# Use session state for working copy (draft mode)
if 'working_df' not in st.session_state or st.session_state.get('reload_data', False):
    st.session_state.working_df = get_data_cached().copy()
    st.session_state.reload_data = False

df = st.session_state.working_df

# Validation: Check if critical columns exist
required_columns = ['날짜', '구분', '대분류', '금액']
missing_cols = [col for col in required_columns if col not in df.columns]

if missing_cols:
    st.error(f"🚨 엑셀 파일에서 다음 필수 열(Column)을 찾을 수 없습니다: {missing_cols}")
    st.stop()

# ----------------- HEADER -----------------
st.title("💰 Personal Expense Dashboard (Local Excel Mode)")

# ----------------- FILTERS (TOP) -----------------
# Moved Filters to TOP so all charts/metrics reflect the same data
with st.expander("🔍 데이터 검색 및 필터 설정 (Data Search & Filter)", expanded=False):
    col_tools_1, col_tools_2 = st.columns([2, 1])
    
    with col_tools_1:
         search_term = st.text_input(" 통합 검색 (검색어 입력)", placeholder="내용, 메모, 카테고리, 금액 등 검색...", label_visibility="collapsed")
        
    with col_tools_2:
        date_preset = st.radio("기간 선택", ["전체", "이번 달", "지난 달", "월별 선택", "직접 입력"], horizontal=True, label_visibility="collapsed")

    # Advanced Filters Logic
    today = datetime.now()
    d_val = []
    
    if date_preset == "이번 달":
        start_d = today.replace(day=1)
        end_d = (start_d + pd.DateOffset(months=1)) - pd.Timedelta(days=1)
        d_val = [start_d, end_d]
    elif date_preset == "지난 달":
        prev_month = today - pd.DateOffset(months=1)
        start_d = prev_month.replace(day=1)
        end_d = (today.replace(day=1)) - pd.Timedelta(days=1)
        d_val = [start_d, end_d]
    elif date_preset == "월별 선택":
        if not df.empty and '날짜' in df.columns:
            df['YYYYMM'] = df['날짜'].dt.strftime('%Y-%m')
            available_months = sorted(df['YYYYMM'].unique(), reverse=True)
            col_m1, _ = st.columns([1,3])
            with col_m1:
                selected_month = st.selectbox("월 선택", available_months, label_visibility="collapsed")
            if selected_month:
                y, m = map(int, selected_month.split('-'))
                start_d = datetime(y, m, 1)
                if m == 12: end_d = datetime(y+1, 1, 1) - pd.Timedelta(days=1)
                else: end_d = datetime(y, m+1, 1) - pd.Timedelta(days=1)
                d_val = [start_d, end_d]
    elif date_preset == "전체":
        d_val = [] 
    else: # Default or others
        if not df.empty:
            min_date = df['날짜'].min()
            max_date = df['날짜'].max()
            d_val = [min_date, max_date]
            
    if date_preset == "직접 입력":
        date_range = st.date_input("날짜 범위", d_val)
    else:
        date_range = d_val

    col_f2, col_f3 = st.columns(2)
    with col_f2:
        all_types = list(df['구분'].unique())
        selected_types = st.multiselect("구분 (Type)", all_types, default=all_types)
    with col_f3:
        all_cats = list(df['대분류'].unique())
        selected_cats = st.multiselect("대분류 (Category)", all_cats, default=all_cats)

# --- APPLY FILTERS ---
filtered_df = df.copy()

# 1. Search Filter
if search_term:
    mask = (
        filtered_df['내용'].astype(str).str.contains(search_term, case=False, na=False) |
        filtered_df['메모'].astype(str).str.contains(search_term, case=False, na=False) |
        filtered_df['대분류'].astype(str).str.contains(search_term, case=False, na=False) |
        filtered_df['소분류'].astype(str).str.contains(search_term, case=False, na=False)
    )
    filtered_df = filtered_df[mask]

# 2. Date Filter
if len(date_range) == 2:
    filtered_df = filtered_df[
        (filtered_df['날짜'].dt.date >= pd.to_datetime(date_range[0]).date()) & 
        (filtered_df['날짜'].dt.date <= pd.to_datetime(date_range[1]).date())
    ]

# 3. Category/Type Filter
if selected_types:
    filtered_df = filtered_df[filtered_df['구분'].isin(selected_types)]
if selected_cats:
    filtered_df = filtered_df[filtered_df['대분류'].isin(selected_cats)]

# ----------------- DASHBOARD SUMMARY (BIG METRICS) -----------------
st.divider()
st.markdown("### 📊 선택기간 요약 (Dashboard)")

active_df = filtered_df[filtered_df['Is_Active'] == True]
sum_income = active_df[active_df['구분'] == '수입']['금액'].sum()
sum_expense = active_df[active_df['구분'] == '지출']['금액'].sum()
sum_inactive = filtered_df[filtered_df['Is_Active'] == False]['금액'].sum()

m_col1, m_col2, m_col3, m_col4 = st.columns(4)
with m_col1: st.metric("✅ 수입", f"{sum_income:,.0f}원")
with m_col2: st.metric("✅ 지출", f"{sum_expense:,.0f}원")
with m_col3: st.metric("✅ 순수익", f"{(sum_income + sum_expense):,.0f}원") # Expense is negative
with m_col4: st.metric("✅ 그 외 (Filter Flow 0)", f"{sum_inactive:,.0f}원")

st.divider()

# ----------------- CHARTS SECTION -----------------
# Only show charts if we have data
if not active_df.empty:
    col_c1, col_c2 = st.columns(2)
    
    # 1. Monthly Trend
    trend_df = active_df.copy()
    trend_df['날짜'] = pd.to_datetime(trend_df['날짜'])
    monthly_trend = trend_df.groupby([pd.Grouper(key='날짜', freq='MS'), '구분'])['금액'].sum().reset_index()
    
    with col_c1:
        st.subheader("🗓️ 월별 재정 흐름")
        fig_trend = px.bar(
            monthly_trend, x='날짜', y='금액', color='구분',
            title="월별 수입/지출 추이",
            color_discrete_map={'수입': 'blue', '지출': 'red', '이체': 'grey'}
        )
        fig_trend.update_xaxes(tickformat="%Y-%m-%d", dtick="M1")
        fig_trend.update_yaxes(tickformat=",")
        st.plotly_chart(fig_trend, use_container_width=True)
        
    # 2. Category Pie
    expense_data = active_df[active_df['구분'] == '지출'].copy()
    if not expense_data.empty:
        cat_trend = expense_data.groupby('대분류')['금액'].sum().reset_index()
        cat_trend['금액'] = cat_trend['금액'].abs()
        
        with col_c2:
            st.subheader("🍰 지출 카테고리 비중")
            fig_cat = px.pie(
                cat_trend, values='금액', names='대분류',
                title="카테고리별 지출 비율",
                hole=0.4
            )
            fig_cat.update_traces(textposition='inside', textinfo='percent+label')
            fig_cat.update_layout(
                showlegend=True,
                legend=dict(orientation="v", yanchor="top", y=1.0, xanchor="left", x=1.05),
                margin=dict(t=50, b=50, l=0, r=100)
            )
            st.plotly_chart(fig_cat, use_container_width=True)

st.divider()

# ----------------- QUICK ENTRY & DATA EDITOR -----------------
st.header("📝 데이터 입력 및 수정")

# Quick Entry Form
with st.expander("➕ 새 데이터 추가 (Quick Entry)", expanded=True):
    # Function to format amount on change
    if 'qe_amount' not in st.session_state: st.session_state.qe_amount = "0"
    
    # Check for Reset Flag (Safe Clear)
    if st.session_state.get('reset_qe_next_run', False):
        st.session_state.qe_amount = "0"
        st.session_state.reset_qe_next_run = False
    
    def format_amount_callback():
        try:
            val = st.session_state.qe_amount.replace(',', '').strip()
            if val:
                st.session_state.qe_amount = f"{int(val):,}"
        except:
            pass # Keep as is if invalid
            
    col_q1, col_q2, col_q3, col_q4 = st.columns(4)
    with col_q1: new_date = st.date_input("날짜", datetime.now())
    with col_q2: new_time = st.time_input("시간", datetime.now().time())
    with col_q3: new_type = st.selectbox("구분", ['지출', '수입', '이체'])
    with col_q4: 
        # Text Input with Callback for formatting
        st.text_input("금액 (자동 쉼표)", key="qe_amount", on_change=format_amount_callback, help="입력 후 엔터를 치면 쉼표가 자동 적용됩니다.")
        
    col_q5, col_q6, col_q7 = st.columns(3)
    with col_q5: 
        cat_options = list(df['대분류'].unique())
        new_category = st.selectbox("대분류", cat_options + ["직접입력"])
    with col_q6: new_sub_category = st.text_input("소분류", "")
    with col_q7: new_payment = st.text_input("결제수단", "")
        
    new_content = st.text_input("내용", "")
    new_memo = st.text_input("메모", "")
    
    # Save Button (Outside Form)
    if st.button("💾 데이터 추가", type="primary"):
        # Parse Amount from Session State
        try:
            clean_amount = st.session_state.qe_amount.replace(',', '').strip()
            final_amount = int(clean_amount)
        except:
            final_amount = 0
            
        final_cat = new_category if new_category != "직접입력" else "미분류"
        new_row = {
            '날짜': pd.to_datetime(new_date),
            '시간': new_time,
            '구분': new_type,
            '대분류': final_cat,
            '소분류': new_sub_category,
            '내용': new_content,
            '금액': final_amount,
            '결제수단': new_payment,
            '메모': new_memo,
            'Is_Active': True,
            'Flow_Filter': 1
        }
        
        # Add to working DataFrame (no immediate save)
        new_row_df = pd.DataFrame([new_row])
        st.session_state.working_df = pd.concat([st.session_state.working_df, new_row_df], ignore_index=True)
        st.session_state.working_df = st.session_state.working_df.sort_values(by=['날짜', '시간'], ascending=[False, False])
        
        st.toast("✅ 추가됨 (저장 전)", icon="📝")
        st.success("✅ 데이터가 추가되었습니다! '변경사항 저장' 버튼을 눌러 저장하세요.")
        
        # Trigger safe reset on next run
        st.session_state.reset_qe_next_run = True
        st.rerun()

st.caption(f"총 {len(df):,}건 중 **{len(filtered_df):,}건** 표시됨")

# Data Editor Setup
# Comma workaround: Convert Amount to String for viewing
editor_df = filtered_df.copy()
if '금액' in editor_df.columns:
    editor_df['금액'] = editor_df['금액'].apply(lambda x: f"{int(x):,}")

column_config = {
    "날짜": st.column_config.DateColumn("날짜", format="YYYY-MM-DD"),
    "시간": st.column_config.TimeColumn("시간", format="HH:mm:ss"),
    "금액": st.column_config.TextColumn("금액", help="금액 입력 (쉼표 가능)", validate=r"^-?[0-9,]+$"), 
    "Is_Active": st.column_config.CheckboxColumn("활성 상태", help="체크 해제 시 통계 제외"),
    "Flow_Filter": st.column_config.NumberColumn("Flow_Filter (자동관리)", help="이 값은 '활성 상태'에 따라 자동으로 설정됩니다. (수정 불가)", disabled=True),
}

edited_subset = st.data_editor(
    editor_df, 
    num_rows="dynamic",
    use_container_width=True,
    height=500,
    key="expense_editor",
    column_config=column_config
)

if st.button("💾 변경사항 저장 (Save to Excel)", type="primary"):
    # Pre-process: Strip commas from Amount and convert to Int
    if '금액' in edited_subset.columns:
        edited_subset['금액'] = edited_subset['금액'].astype(str).str.replace(',', '').str.strip()
        edited_subset['금액'] = pd.to_numeric(edited_subset['금액'], errors='coerce').fillna(0).astype(int)

    try:
        with st.spinner("💾 엑셀 파일에 저장 중입니다... (잠시만 기다려주세요)"):
            visible_indices = set(filtered_df.index)
            edited_indices = set(edited_subset.index)
            
            deleted_indices = visible_indices - edited_indices
            all_original_indices = set(df.index)
            new_indices = edited_indices - all_original_indices
            common_indices = edited_indices.intersection(all_original_indices)
            
            final_df = df.copy()
            
            if deleted_indices:
                final_df = final_df.drop(index=list(deleted_indices))
            
            if common_indices:
                updates = edited_subset.loc[list(common_indices)]
                final_df.update(updates)
                
            if new_indices:
                new_rows = edited_subset.loc[list(new_indices)]
                final_df = pd.concat([final_df, new_rows])
            
            # FORCE SYNC Flow_Filter
            if 'Is_Active' in final_df.columns:
                final_df['Flow_Filter'] = final_df['Is_Active'].apply(lambda x: 1 if x is True or x==1 else 0)
            
            if '날짜' in final_df.columns:
                 final_df = final_df.sort_values(by=['날짜'], ascending=False)
            
            if save_data(final_df):
                st.toast("✅ 저장 성공!", icon="🎉")
                st.success("✅ 저장이 완료되었습니다!")
                
                # Update session state with saved data
                st.session_state.working_df = final_df.copy()
                
                # Clear cache and trigger reload on next run
                get_data_cached.clear()
                st.session_state.reload_data = True
                
                import time
                time.sleep(1.5)
                st.rerun()
            else:
                st.error("❌ 저장 실패: 파일이 열려있는지 확인하세요.")
            
    except Exception as e:
        st.error(f"Save Error: {e}")

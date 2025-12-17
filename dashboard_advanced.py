import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
import gspread
from google.oauth2.service_account import Credentials

# Page Configuration
st.set_page_config(page_title="재정 상태 통합 대시보드", layout="wide", initial_sidebar_state="collapsed")

# Custom CSS for better styling
st.markdown("""
<style>
    .main {
        padding: 0rem 1rem;
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    }
    h1 {
        color: #1f77b4;
        padding-bottom: 20px;
    }
    .filter-container {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Load Data from Google Sheets
@st.cache_data(ttl=300)  # Cache for 5 minutes
def load_data():
    try:
        # Google Sheets credentials
        SCOPES = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
        
        # Service account JSON path
        SERVICE_ACCOUNT_FILE = r"C:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\service_account.json"
        
        # Authenticate
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # Open spreadsheet
        SPREADSHEET_ID = "1DqpTecTdpRKsXTPImM4iKPT2V-KeJixG85-K6MuLOWY"
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        
        # Get DB_Raw sheet
        sheet = spreadsheet.worksheet("DB_Raw")
        
        # Get all values
        data = sheet.get_all_values()
        
        # Convert to DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])
        
        # Data cleaning
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        df['amount'] = pd.to_numeric(df['amount'], errors='coerce')
        
        # Remove rows with invalid data
        df = df.dropna(subset=['date', 'amount', 'type'])
        
        # Sort by date
        df = df.sort_values('date')
        
        return df
        
    except Exception as e:
        st.error(f"Google Sheets 연결 오류: {e}")
        st.info("로컬 Excel 파일로 전환합니다...")
        
        # Fallback to local Excel file
        file_path = r"c:\Users\JTC7\Desktop\01.Python Project\01.Personal Expense Tracker\01.Document\2024-12-07~2025-12-07_v5_LinkDB.xlsx"
        df = pd.read_excel(file_path, sheet_name="DB_Raw", engine='openpyxl')
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        df['amount'] = pd.to_numeric(df['amount'], errors='coerce')
        df = df.dropna(subset=['date', 'amount', 'type'])
        df = df.sort_values('date')
        return df

try:
    df = load_data()
except Exception as e:
    st.error(f"데이터 로드 실패: {e}")
    st.stop()

# Title
st.title("💎 재정 상태 통합 대시보드")

# ====== [A. 상단 영역: 필터 및 KPI] ======
st.markdown("### 📅 기간 및 필터 설정")

col_preset, col_custom, col_payment = st.columns([2, 3, 2])

with col_preset:
    # Preset date filters
    preset_options = {
        "이번 달": 30,
        "지난 3개월": 90,
        "지난 6개월": 180,
        "전체": None
    }
    selected_preset = st.selectbox("빠른 선택", list(preset_options.keys()), key="preset")
    
    if preset_options[selected_preset] is not None:
        days = preset_options[selected_preset]
        end_date = df['date'].max()
        start_date = end_date - pd.Timedelta(days=days)
    else:
        start_date = df['date'].min()
        end_date = df['date'].max()

with col_custom:
    # Custom date range
    if selected_preset == "전체":
        min_date = df['date'].min().date()
        max_date = df['date'].max().date()
        custom_range = st.date_input(
            "사용자 정의 기간",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date,
            key="custom_date"
        )
        if len(custom_range) == 2:
            start_date = pd.Timestamp(custom_range[0])
            end_date = pd.Timestamp(custom_range[1])
    else:
        # Display current period in a text input to match height
        st.text_input(
            "현재 선택된 기간",
            value=f"{start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}",
            disabled=True,
            key="date_display"
        )
    

with col_payment:
    # Payment method filter (Button-style)
    payment_methods = df['payment_method'].dropna().unique().tolist()
    selected_payments = st.multiselect(
        "💳 결제 수단",
        ["전체"] + payment_methods,
        default=["전체"],
        key="payment_filter"
    )

# Filter data
mask = (df['date'] >= start_date) & (df['date'] <= end_date)
filtered_df = df[mask].copy()

if "전체" not in selected_payments:
    filtered_df = filtered_df[filtered_df['payment_method'].isin(selected_payments)]

# Calculate current period KPIs
total_income = filtered_df[filtered_df['type'] == '수입']['amount'].sum()
total_expense = filtered_df[filtered_df['type'] == '지출']['amount'].sum()
net_income = total_income - total_expense

# Calculate previous period for MoM comparison
period_days = (end_date - start_date).days
prev_start = start_date - pd.Timedelta(days=period_days)
prev_end = start_date

prev_mask = (df['date'] >= prev_start) & (df['date'] < prev_end)
prev_df = df[prev_mask]

prev_income = prev_df[prev_df['type'] == '수입']['amount'].sum()
prev_expense = prev_df[prev_df['type'] == '지출']['amount'].sum()
prev_net = prev_income - prev_expense

# Calculate deltas
income_delta = total_income - prev_income
expense_delta = total_expense - prev_expense
net_delta = net_income - prev_net

# Top expense category
if not filtered_df[filtered_df['type'] == '지출'].empty:
    top_category = filtered_df[filtered_df['type'] == '지출'].groupby('main_category')['amount'].sum().idxmax()
    top_category_amount = filtered_df[filtered_df['type'] == '지출'].groupby('main_category')['amount'].sum().max()
else:
    top_category = "N/A"
    top_category_amount = 0

# KPI Display with enhanced styling
st.markdown("---")
st.markdown("### 💎 핵심 지표")

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    delta_pct = (income_delta / prev_income * 100) if prev_income != 0 else 0
    st.metric(
        "💰 총 수입",
        f"{total_income:,.0f}원",
        delta=f"{income_delta:,.0f}원 ({delta_pct:+.1f}%)",
        delta_color="normal"
    )

with kpi2:
    delta_pct = (expense_delta / prev_expense * 100) if prev_expense != 0 else 0
    st.metric(
        "💸 총 지출",
        f"{total_expense:,.0f}원",
        delta=f"{expense_delta:,.0f}원 ({delta_pct:+.1f}%)",
        delta_color="inverse"
    )

with kpi3:
    delta_pct = (net_delta / prev_net * 100) if prev_net != 0 else 0
    st.metric(
        "✅ 순수익",
        f"{net_income:,.0f}원",
        delta=f"{net_delta:,.0f}원 ({delta_pct:+.1f}%)",
        delta_color="normal"
    )

with kpi4:
    st.metric(
        "🔥 최대 지출",
        top_category,
        f"{top_category_amount:,.0f}원"
    )

st.markdown("---")

# ====== [B. 왼쪽 패널: 재정 흐름 분석] & [C. 오른쪽 패널: 지출 구조] ======
col_left, col_right = st.columns([1, 1])

with col_left:
    st.markdown("### 📈 재정 흐름 분석 (시간)")
    
    # Monthly trend chart - USE FULL DATA (df) to show all 12 months
    monthly_data = df.groupby([pd.Grouper(key='date', freq='ME'), 'type'])['amount'].sum().reset_index()
    monthly_pivot = monthly_data.pivot(index='date', columns='type', values='amount').fillna(0)
    
    if '수입' not in monthly_pivot.columns:
        monthly_pivot['수입'] = 0
    if '지출' not in monthly_pivot.columns:
        monthly_pivot['지출'] = 0
    
    monthly_pivot['순수익'] = monthly_pivot.get('수입', 0) - monthly_pivot.get('지출', 0)
    monthly_pivot = monthly_pivot.reset_index()
    monthly_pivot['월'] = monthly_pivot['date'].dt.strftime('%Y-%m')
    
    
    
    # Monthly Bar Chart Only (Grouped Bars)
    fig_monthly = go.Figure()
    
    fig_monthly.add_trace(go.Bar(
        x=monthly_pivot['월'], 
        y=monthly_pivot['수입'],
        name='수입',
        marker_color='#3498db',  # Blue
        hovertemplate='<b>수입</b><br>%{y:,.0f}원<extra></extra>',
        text=[f'{x:,.0f}' for x in monthly_pivot['수입']],
        textposition='outside'
    ))
    
    fig_monthly.add_trace(go.Bar(
        x=monthly_pivot['월'], 
        y=monthly_pivot['지출'],
        name='지출',
        marker_color='#e74c3c',  # Red
        hovertemplate='<b>지출</b><br>%{y:,.0f}원<extra></extra>',
        text=[f'{x:,.0f}' for x in monthly_pivot['지출']],
        textposition='outside'
    ))
    
    # Net income bars (different colors based on positive/negative)
    colors = ['#e74c3c' if x < 0 else '#27ae60' for x in monthly_pivot['순수익']]
    fig_monthly.add_trace(go.Bar(
        x=monthly_pivot['월'], 
        y=monthly_pivot['순수익'],
        name='순수익',
        marker_color=colors,
        hovertemplate='<b>순수익</b><br>%{y:,.0f}원<extra></extra>',
        text=[f'{x:,.0f}' for x in monthly_pivot['순수익']],
        textposition='outside'
    ))
    
    fig_monthly.update_layout(
        title={
            'text': "월별 재정 추이 (전체 기간)",
            'font': {'size': 18, 'color': '#2c3e50'}
        },
        xaxis_title="월",
        yaxis_title="금액 (원)",
        hovermode='x unified',
        height=500,
        barmode='group',  # Grouped bars
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            bgcolor="rgba(255,255,255,0.8)"
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.05)')
    )
    
    st.plotly_chart(fig_monthly, width='stretch')

with col_right:
    st.markdown("### 💸 지출 구조 분석 (비중)")
    
    # Expense breakdown by main_category
    expense_df = filtered_df[filtered_df['type'] == '지출']
    
    if not expense_df.empty:
        category_expense = expense_df.groupby('main_category')['amount'].sum().reset_index()
        category_expense = category_expense.sort_values('amount', ascending=False)
        
        
        # Donut chart WITHOUT labels (textinfo='none')
        fig_donut = go.Figure(data=[go.Pie(
            labels=category_expense['main_category'],
            values=category_expense['amount'],
            hole=.45,
            marker=dict(colors=['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#1abc9c', '#e67e22']),
            pull=[0.1 if i == 0 else 0 for i in range(len(category_expense))],  # Explode largest
            textposition='none',  # Remove text labels
            textinfo='none',  # Remove all text
            hovertemplate='<b>%{label}</b><br>금액: %{value:,.0f}원<br>비율: %{percent}<extra></extra>'
        )])
        
        fig_donut.update_layout(
            title={
                'text': "카테고리별 지출 비중",
                'font': {'size': 18, 'color': '#2c3e50'}
            },
            height=300,
            showlegend=True,  # Show legend instead of labels
            legend=dict(
                orientation="v",
                yanchor="middle",
                y=0.5,
                xanchor="left",
                x=1.05
            ),
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        st.plotly_chart(fig_donut, width='stretch')
        
        # Top 5 sub-categories
        st.markdown("#### 상위 5개 세부 카테고리")
        
        sub_expense = expense_df.groupby('sub_category')['amount'].sum().reset_index()
        sub_expense = sub_expense.sort_values('amount', ascending=False).head(5)
        
        
        fig_sub = go.Figure(go.Bar(
            x=sub_expense['amount'],
            y=sub_expense['sub_category'],
            orientation='h',
            marker=dict(
                color=sub_expense['amount'],
                colorscale='Reds',
                showscale=False
            ),
            text=[f'{x:,.0f}원' for x in sub_expense['amount']],
            textposition='outside',
            hovertemplate='<b>%{y}</b><br>금액: %{x:,.0f}원<extra></extra>'
        ))
        
        fig_sub.update_layout(
            title={
                'text': "지출액 기준 상위 5개 세부 카테고리",
                'font': {'size': 16, 'color': '#2c3e50'}
            },
            xaxis_title="금액 (원)",
            yaxis_title="",
            height=350,
            yaxis=dict(autorange="reversed"),
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(showgrid=True, gridcolor='rgba(0,0,0,0.05)'),
            margin=dict(l=20, r=20, t=60, b=40)
        )
        
        st.plotly_chart(fig_sub, width='stretch')
    else:
        st.info("선택한 기간에 지출 데이터가 없습니다.")

st.markdown("---")

# ====== [D. 하단 영역: 거래 내역 표] ======
st.markdown("### 📋 거래 내역")

# Show transaction table
if not filtered_df.empty:
    display_df = filtered_df[['date', 'type', 'main_category', 'sub_category', 'amount', 'payment_method', 'merchant', 'memo']].copy()
    display_df.columns = ['날짜', '구분', '대분류', '소분류', '금액', '결제수단', '거래처', '메모']
    display_df['날짜'] = pd.to_datetime(display_df['날짜']).dt.strftime('%Y-%m-%d')
    # Convert all to string to avoid Arrow serialization issues
    display_df = display_df.astype(str)
    
    # Configure columns
    column_config = {
        "금액": st.column_config.NumberColumn(
            "금액",
            help="거래 금액",
            format="₩%d"
        )
    }
    
    st.dataframe(
        display_df,
        width='stretch',
        height=400,
        hide_index=True,
        column_config=column_config
    )
    
    # Summary stats
    col_stat1, col_stat2, col_stat3 = st.columns(3)
    with col_stat1:
        st.metric("총 거래 건수", f"{len(filtered_df):,}건")
    with col_stat2:
        st.metric("평균 지출액", f"{expense_df['amount'].mean():,.0f}원" if not expense_df.empty else "0원")
    with col_stat3:
        st.metric("최대 거래액", f"{filtered_df['amount'].max():,.0f}원")
else:
    st.warning("선택한 기간에 거래 데이터가 없습니다.")

# Footer
st.markdown("---")
st.caption("💡 데이터는 실시간으로 업데이트됩니다. 필터를 조정하여 다양한 분석을 수행하세요.")

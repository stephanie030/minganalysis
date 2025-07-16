#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
庫存分析 Streamlit 動態儀表板
提供互動式的庫存分析和視覺化功能
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# 設定頁面配置
st.set_page_config(
    page_title="庫存分析儀表板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義CSS樣式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 0.5rem;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
# def load_data(file_path="C:\\Users\\ur14068\\Downloads\\分析\\分析\\銘宣爬蟲合併_0421_0708.xlsx"):
def load_data(file_path="銘宣爬蟲合併_0421_0708.xlsx"):
    """載入並處理資料"""
    try:
        df = pd.read_excel(file_path)
        
        # 資料清理
        df['爬取日期'] = pd.to_datetime(df['爬取日期'])
        df['客戶庫存量(現有數量)'] = pd.to_numeric(df['客戶庫存量(現有數量)'], errors='coerce')
        df['包裝_數量'] = pd.to_numeric(df['包裝_數量'], errors='coerce')
        
        return df
    except Exception as e:
        st.error(f"資料載入失敗: {e}")
        return None

def calculate_basic_stats(df):
    """計算基本統計"""
    return {
        '總筆數': len(df),
        '料號數量': df['客戶料號'].nunique(),
        '產品種類': df['客戶品名'].nunique(),
        '倉庫數量': df['客戶庫別'].nunique(),
        '資料天數': df['爬取日期'].nunique(),
        '最新日期': df['爬取日期'].max().strftime('%Y-%m-%d'),
        '總庫存量': df[df['爬取日期'] == df['爬取日期'].max()]['客戶庫存量(現有數量)'].sum()
    }

def create_trend_chart(df):
    """建立趨勢圖表"""
    daily_inventory = df.groupby('爬取日期')['客戶庫存量(現有數量)'].sum().reset_index()
    
    fig = px.line(
        daily_inventory, 
        x='爬取日期', 
        y='客戶庫存量(現有數量)',
        title='每日總庫存量趨勢',
        labels={'客戶庫存量(現有數量)': '庫存量', '爬取日期': '日期'}
    )
    
    fig.update_traces(line=dict(width=3))
    fig.update_layout(
        height=400,
        hovermode='x unified',
        xaxis_title="日期",
        yaxis_title="庫存量"
    )
    
    return fig

def create_warehouse_chart(df):
    """建立倉庫分析圖表"""
    latest_date = df['爬取日期'].max()
    latest_data = df[df['爬取日期'] == latest_date]
    
    warehouse_stats = latest_data.groupby('客戶庫別').agg({
        '客戶庫存量(現有數量)': 'sum',
        '客戶料號': 'nunique'
    }).reset_index()
    
    warehouse_stats.columns = ['倉庫', '總庫存量', '料號數量']
    warehouse_stats = warehouse_stats.sort_values('總庫存量', ascending=True)
    
    fig = px.bar(
        warehouse_stats,
        x='總庫存量',
        y='倉庫',
        orientation='h',
        title='各倉庫庫存分布',
        labels={'總庫存量': '庫存量', '倉庫': '倉庫'}
    )
    
    fig.update_layout(height=400)
    return fig

def create_product_analysis(df, top_n=20):
    """建立產品分析"""
    latest_date = df['爬取日期'].max()
    latest_data = df[df['爬取日期'] == latest_date]
    
    product_inventory = latest_data.groupby(['客戶料號', '客戶品名']).agg({
        '客戶庫存量(現有數量)': 'sum'
    }).reset_index()
    
    product_inventory.columns = ['料號', '品名', '總庫存量']
    product_inventory = product_inventory.sort_values('總庫存量', ascending=False)
    
    return product_inventory.head(top_n)

def detect_anomalies(df, threshold=50):
    """檢測異常變化"""
    anomalies = []
    
    for product in df['客戶料號'].unique():
        product_data = df[df['客戶料號'] == product].groupby('爬取日期')['客戶庫存量(現有數量)'].sum().reset_index()
        
        if len(product_data) > 1:
            product_data['變化率'] = product_data['客戶庫存量(現有數量)'].pct_change() * 100
            
            large_changes = product_data[abs(product_data['變化率']) > threshold]
            
            for _, row in large_changes.iterrows():
                if not pd.isna(row['變化率']):
                    anomalies.append({
                        '料號': product,
                        '日期': row['爬取日期'].strftime('%Y-%m-%d'),
                        '變化率': f"{row['變化率']:+.1f}%",
                        '當前庫存': row['客戶庫存量(現有數量)']
                    })
    
    return pd.DataFrame(anomalies)

def main():
    """主程式"""
    # 標題
    st.markdown('<h1 class="main-header">📊 庫存分析動態儀表板</h1>', unsafe_allow_html=True)
    
    # 側邊欄
    st.sidebar.header("🔧 控制面板")
    
    # 載入資料
    df = load_data()
    if df is None:
        st.stop()
    
    # 基本統計
    stats = calculate_basic_stats(df)
    
    # 側邊欄篩選器
    st.sidebar.subheader("📅 日期範圍")
    date_range = st.sidebar.date_input(
        "選擇分析日期範圍",
        value=(df['爬取日期'].min().date(), df['爬取日期'].max().date()),
        min_value=df['爬取日期'].min().date(),
        max_value=df['爬取日期'].max().date()
    )
    
    st.sidebar.subheader("🏭 倉庫篩選")
    selected_warehouses = st.sidebar.multiselect(
        "選擇要分析的倉庫",
        options=df['客戶庫別'].unique(),
        default=df['客戶庫別'].unique()
    )
    
    # 篩選資料
    if len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = df[
            (df['爬取日期'].dt.date >= start_date) & 
            (df['爬取日期'].dt.date <= end_date) &
            (df['客戶庫別'].isin(selected_warehouses))
        ]
    else:
        filtered_df = df[df['客戶庫別'].isin(selected_warehouses)]
    
    # 重新計算統計
    filtered_stats = calculate_basic_stats(filtered_df)
    
    # 顯示關鍵指標
    st.subheader("📈 關鍵指標")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("總庫存量", f"{filtered_stats['總庫存量']:,.0f}")
    with col2:
        st.metric("料號數量", f"{filtered_stats['料號數量']:,}")
    with col3:
        st.metric("倉庫數量", f"{filtered_stats['倉庫數量']}")
    with col4:
        st.metric("資料天數", f"{filtered_stats['資料天數']}")
    
    # 趨勢分析
    st.subheader("📊 庫存趨勢分析")
    trend_chart = create_trend_chart(filtered_df)
    st.plotly_chart(trend_chart, use_container_width=True)
    
    # 倉庫分析
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🏭 倉庫庫存分布")
        warehouse_chart = create_warehouse_chart(filtered_df)
        st.plotly_chart(warehouse_chart, use_container_width=True)
    
    with col2:
        st.subheader("🔍 庫存量前20名產品")
        top_products = create_product_analysis(filtered_df, 20)
        st.dataframe(top_products, use_container_width=True)
    
    # 異常檢測
    st.subheader("🚨 異常檢測")
    anomaly_threshold = st.slider("異常變化閾值 (%)", 10, 100, 50)
    anomalies = detect_anomalies(filtered_df, anomaly_threshold)
    
    if len(anomalies) > 0:
        st.warning(f"⚠️ 發現 {len(anomalies)} 個異常變化")
        st.dataframe(anomalies, use_container_width=True)
    else:
        st.success("✅ 未發現明顯異常")
    
    # 資料表格
    with st.expander("📋 原始資料"):
        st.dataframe(filtered_df, use_container_width=True)
    
    # 更新時間
    st.sidebar.markdown("---")
    st.sidebar.info(f"📅 最後更新: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    st.sidebar.button("🔄 重新整理資料", key="refresh")

if __name__ == "__main__":
    main()

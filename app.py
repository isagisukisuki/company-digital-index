# app.py - 上市公司数字化转型指数查询系统（百分制）
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.graph_objects as go  # 只保留用到的模块
import numpy as np

# ====================== 核心配置 =======================
DATA_FILE = Path(__file__).parent / "数字化转型指数分析结果.xlsx"
WORD_FREQ_COLS = [
    "人工智能词频数",
    "大数据词频数",
    "云计算词频数",
    "区块链词频数",
    "数字技术运用词频数"
]
# ======================================================

st.set_page_config(
    page_title="数字化转型指数查询系统",
    page_icon="📊",
    layout="wide"
)

# 百分制指数计算
def calculate_percentile_index(df):
    df["年度总词频数"] = df[WORD_FREQ_COLS].sum(axis=1)
    def _calc_year_index(year_df):
        year_max_total = year_df["年度总词频数"].max()
        if year_max_total == 0:
            year_df["数字化转型指数"] = 0.0
        else:
            year_df["数字化转型指数"] = (year_df["年度总词频数"] / year_max_total * 100).round(2)
        year_df["数字化转型指数"] = year_df["数字化转型指数"].clip(lower=0, upper=100)
        year_df.loc[year_df["年度总词频数"] == 0, "数字化转型指数"] = 0.0
        return year_df
    df = df.groupby("年份", group_keys=False).apply(_calc_year_index)
    return df.drop("年度总词频数", axis=1)

# 缓存加载数据
@st.cache_data
def load_data():
    try:
        excel_file = pd.ExcelFile(DATA_FILE, engine="openpyxl")
        sheet_names = [name for name in excel_file.sheet_names if name.isdigit()]
        df_list = []
        for sheet in sheet_names:
            sheet_df = pd.read_excel(DATA_FILE, sheet_name=sheet, engine="openpyxl")
            sheet_df["年份"] = int(sheet)
            df_list.append(sheet_df)
        df = pd.concat(df_list, ignore_index=True).fillna(0)
        if "股票代码" in df.columns:
            df["股票代码"] = df["股票代码"].astype(str).str.zfill(6)
        df = calculate_percentile_index(df)
        unique_stocks = sorted(df['股票代码'].unique())
        unique_companies = sorted(df['企业名称'].unique())
        unique_years = sorted(df['年份'].unique())
        stock_to_company = {k: df[df['股票代码']==k]['企业名称'].iloc[0] for k in unique_stocks}
        return df, unique_stocks, unique_companies, unique_years, stock_to_company
    except Exception as e:
        st.error(f"加载数据失败: {str(e)}")
        return pd.DataFrame(), [], [], [], {}

# 主逻辑
st.title("📊 上市公司数字化转型指数查询系统")
st.markdown("### 查询1999-2023年上市公司的数字化转型指数数据（百分制）")

df, unique_stocks, unique_companies, unique_years, stock_to_company = load_data()

with st.sidebar:
    st.header("🔍 查询条件")
    search_type = st.radio("搜索方式:", ["股票代码", "企业名称"])
    selected_stock, selected_company = None, None
    if search_type == "股票代码":
        selected_stock = st.selectbox(
            "选择股票代码:",
            options=unique_stocks,
            format_func=lambda x: f"{x} - {stock_to_company.get(x, '未知企业')}",
            index=None,
            placeholder="请选择股票代码"
        )
        if selected_stock:
            selected_company = stock_to_company.get(selected_stock, "")
    else:
        selected_company = st.selectbox(
            "选择企业名称:",
            options=unique_companies,
            index=None,
            placeholder="请选择企业名称"
        )
        if selected_company:
            selected_stock = df[df['企业名称'] == selected_company]['股票代码'].iloc[0] if not df[df['企业名称'] == selected_company].empty else None
    selected_year = st.selectbox("选择年份:", options=unique_years, index=None, placeholder="请选择年份(可选)")
    search_button = st.button("📈 执行查询")

if df.empty:
    st.warning("暂无数据可供查询")
else:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📊 数据总量", f"{len(df):,}")
    with col2:
        st.metric("🏢 企业数量", f"{len(unique_companies):,}")
    with col3:
        st.metric("📅 年份跨度", f"{min(unique_years)}-{max(unique_years)}")

    if search_button and selected_stock:
        filtered_data = df[(df['股票代码'] == selected_stock) & (df['年份'] == selected_year)] if selected_year else df[df['股票代码'] == selected_stock]
        if not filtered_data.empty:
            company_name = filtered_data['企业名称'].iloc[0]
            st.subheader(f"📋 {company_name} (股票代码: {selected_stock})")
            company_history = df[df['股票代码'] == selected_stock].sort_values('年份')
            
            # Plotly图表
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=company_history['年份'],
                y=company_history['数字化转型指数'],
                mode='lines+markers',
                name='数字化转型指数（百分制）',
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=8, color='#1f77b4')
            ))
            if selected_year:
                current_value = filtered_data['数字化转型指数'].iloc[0]
                fig.add_trace(go.Scatter(
                    x=[selected_year],
                    y=[current_value],
                    mode='markers',
                    name=f'{selected_year}年',
                    marker=dict(size=12, color='#ff7f0e', symbol='star'),
                    text=f'{selected_year}年: {current_value}分'
                ))
            fig.update_layout(
                title=f'{company_name}历年数字化转型指数趋势 (1999-2023) - 百分制',
                xaxis_title='年份',
                yaxis_title='数字化转型指数（0-100分）',
                template='plotly_white',
                height=500,
                yaxis=dict(range=[0, 100])
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 显示数据
            display_cols = ["股票代码", "企业名称", "年份"] + WORD_FREQ_COLS + ["数字化转型指数"]
            st.subheader("📊 详细数据（含词频）")
            st.dataframe(filtered_data[display_cols] if selected_year else company_history[display_cols], use_container_width=True)
            
            if not selected_year:
                st.subheader("📈 统计分析（百分制）")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("最高指数（分）", f"{company_history['数字化转型指数'].max():.2f}")
                with col2:
                    st.metric("最低指数（分）", f"{company_history['数字化转型指数'].min():.2f}")
                with col3:
                    st.metric("平均指数（分）", f"{company_history['数字化转型指数'].mean():.2f}")
                with col4:
                    st.metric("指数增长（分）", f"{company_history['数字化转型指数'].iloc[-1] - company_history['数字化转型指数'].iloc[0]:+.2f}")
        else:
            st.warning(f"未找到{selected_stock}在{selected_year}年的数据")
    else:
        st.info("请在侧边栏选择股票代码或企业名称，并点击'执行查询'按钮查看数据")
        st.subheader("📊 数据示例（含词频+百分制指数）")
        st.dataframe(df[["股票代码", "企业名称", "年份"] + WORD_FREQ_COLS + ["数字化转型指数"]].head(10), use_container_width=True)
        st.subheader("📝 使用说明")
        st.markdown("""
        1. 选择搜索方式（股票代码/企业名称）
        2. 选择对应标的，点击“执行查询”
        3. 查看趋势图和详细数据（指数为0-100分制）
        """)

st.markdown("""
---
💡 数据来源：数字化转型指数分析结果.xlsx
📌 指数规则：0-100百分制（按年度归一化）
""")

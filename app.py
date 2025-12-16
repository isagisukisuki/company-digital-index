# app.py - 上市公司数字化转型指数查询系统（百分制）
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# ====================== 核心配置（按你的要求修改）======================
# 文件路径（替换为你指定的文件，与app.py同目录）
DATA_FILE = Path(__file__).parent / "数字化转型指数分析结果.xlsx"
# 词频相关列名（按你的要求保留）
WORD_FREQ_COLS = [
    "人工智能词频数",
    "大数据词频数",
    "云计算词频数",
    "区块链词频数",
    "数字技术运用词频数"
]
# =====================================================================

# 设置中文字体支持
st.set_page_config(
    page_title="数字化转型指数查询系统",
    page_icon="📊",
    layout="wide"
)

# 核心函数：计算百分制指数（按年度归一化，确保每年有100分）
def calculate_percentile_index(df):
    # 计算每家企业的年度总词频数
    df["年度总词频数"] = df[WORD_FREQ_COLS].sum(axis=1)
    
    # 按年份分组计算百分制指数
    def _calc_year_index(year_df):
        year_max_total = year_df["年度总词频数"].max()
        if year_max_total == 0:
            year_df["数字化转型指数"] = 0.0
        else:
            # 归一化到0-100分
            year_df["数字化转型指数"] = (year_df["年度总词频数"] / year_max_total * 100).round(2)
        # 强制无负数、词频全零则指数为0
        year_df["数字化转型指数"] = year_df["数字化转型指数"].clip(lower=0, upper=100)
        year_df.loc[year_df["年度总词频数"] == 0, "数字化转型指数"] = 0.0
        return year_df
    
    df = df.groupby("年份", group_keys=False).apply(_calc_year_index)
    df = df.drop("年度总词频数", axis=1)
    return df

# 缓存数据加载函数
@st.cache_data
def load_data():
    try:
        # 读取Excel文件（支持多sheet，sheet名为年份数字）
        excel_file = pd.ExcelFile(DATA_FILE, engine="openpyxl")
        sheet_names = [name for name in excel_file.sheet_names if name.isdigit()]
        
        df_list = []
        for sheet in sheet_names:
            sheet_df = pd.read_excel(DATA_FILE, sheet_name=sheet, engine="openpyxl")
            sheet_df["年份"] = int(sheet)  # 工作表名转为年份数字
            df_list.append(sheet_df)
        
        # 合并所有年份数据
        df = pd.concat(df_list, ignore_index=True)
        df = df.fillna(0)
        
        # 修正股票代码格式（补零到6位）
        if "股票代码" in df.columns:
            df["股票代码"] = df["股票代码"].astype(str).str.zfill(6)
        
        # 计算百分制数字化转型指数（覆盖原始指数）
        df = calculate_percentile_index(df)
        
        # 提取唯一值
        unique_stocks = sorted(df['股票代码'].unique())
        unique_companies = sorted(df['企业名称'].unique())
        unique_years = sorted(df['年份'].unique())
        
        # 创建股票代码到企业名称的映射
        stock_to_company = dict(zip(df['股票代码'], df['企业名称']))
        stock_to_company = {k: stock_to_company[k] for k in unique_stocks}  # 去重
        
        return df, unique_stocks, unique_companies, unique_years, stock_to_company
    except Exception as e:
        st.error(f"加载数据失败: {e}")
        return pd.DataFrame(), [], [], [], {}

# 应用标题
st.title("📊 上市公司数字化转型指数查询系统")
st.markdown("### 查询1999-2023年上市公司的数字化转型指数数据（百分制）")

# 加载数据
with st.spinner("正在加载数据..."):
    df, unique_stocks, unique_companies, unique_years, stock_to_company = load_data()

# 侧边栏 - 查询控件
with st.sidebar:
    st.header("🔍 查询条件")
    
    # 创建股票代码和企业名称的联合选择器
    search_type = st.radio("搜索方式:", ["股票代码", "企业名称"])
    
    selected_stock = None
    selected_company = None
    if search_type == "股票代码":
        selected_stock = st.selectbox(
            "选择股票代码:",
            options=unique_stocks,
            format_func=lambda x: f"{x} - {stock_to_company.get(x, '未知企业')}",
            index=None,
            placeholder="请选择股票代码"
        )
        # 获取对应的企业名称
        if selected_stock:
            selected_company = stock_to_company.get(selected_stock, "")
    else:
        selected_company = st.selectbox(
            "选择企业名称:",
            options=unique_companies,
            index=None,
            placeholder="请选择企业名称"
        )
        # 获取对应的股票代码
        if selected_company:
            # 找到第一个匹配的股票代码
            selected_stock = df[df['企业名称'] == selected_company]['股票代码'].iloc[0] if not df[df['企业名称'] == selected_company].empty else None
    
    # 年份选择器
    selected_year = st.selectbox(
        "选择年份:",
        options=unique_years,
        index=None,
        placeholder="请选择年份(可选)"
    )
    
    # 查询按钮
    search_button = st.button("📈 执行查询")

# 主页面内容
if df.empty:
    st.warning("暂无数据可供查询")
else:
    # 数据概览卡片
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("📊 数据总量", f"{len(df):,}")
    with col2:
        st.metric("🏢 企业数量", f"{len(unique_companies):,}")
    with col3:
        st.metric("📅 年份跨度", f"{min(unique_years)}-{max(unique_years)}")
    
    # 如果用户点击了查询按钮或选择了股票代码
    if search_button and selected_stock:
        # 筛选数据
        if selected_year:
            # 按股票代码和年份筛选
            filtered_data = df[(df['股票代码'] == selected_stock) & (df['年份'] == selected_year)]
        else:
            # 只按股票代码筛选
            filtered_data = df[df['股票代码'] == selected_stock]
        
        if not filtered_data.empty:
            # 获取企业名称
            company_name = filtered_data['企业名称'].iloc[0]
            
            # 显示企业信息
            st.subheader(f"📋 {company_name} (股票代码: {selected_stock})")
            
            # 创建历年数据的折线图
            company_history = df[df['股票代码'] == selected_stock].sort_values('年份')
            
            # 创建折线图（百分制指数）
            fig = go.Figure()
            
            # 添加数字化转型指数折线（百分制）
            fig.add_trace(go.Scatter(
                x=company_history['年份'],
                y=company_history['数字化转型指数'],
                mode='lines+markers',
                name='数字化转型指数（百分制）',
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=8, color='#1f77b4', symbol='circle')
            ))
            
            # 添加当前查询年份的标记点（如果选择了年份）
            if selected_year:
                current_value = filtered_data['数字化转型指数'].iloc[0]
                fig.add_trace(go.Scatter(
                    x=[selected_year],
                    y=[current_value],
                    mode='markers',
                    name=f'{selected_year}年',
                    marker=dict(size=12, color='#ff7f0e', symbol='star'),
                    text=f'{selected_year}年: {current_value}分',
                    hoverinfo='text'
                ))
            
            # 更新布局（Y轴固定0-100，体现百分制）
            fig.update_layout(
                title=f'{company_name}历年数字化转型指数趋势 (1999-2023) - 百分制',
                xaxis_title='年份',
                yaxis_title='数字化转型指数（0-100分）',
                template='plotly_white',
                height=500,
                legend_title='指标',
                hovermode='x unified',
                yaxis=dict(range=[0, 100])  # 强制Y轴0-100
            )
            
            # 显示图表
            st.plotly_chart(fig, use_container_width=True)
            
            # 显示详细数据（包含词频字段）
            st.subheader("📊 详细数据（含词频）")
            # 保留核心列：股票代码、企业名称、年份 + 词频列 + 百分制指数
            display_cols = ["股票代码", "企业名称", "年份"] + WORD_FREQ_COLS + ["数字化转型指数"]
            display_data = filtered_data[display_cols] if selected_year else company_history[display_cols]
            
            st.dataframe(display_data, use_container_width=True)
            
            # 显示统计信息（百分制）
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
                    growth = company_history['数字化转型指数'].iloc[-1] - company_history['数字化转型指数'].iloc[0]
                    st.metric("指数增长（分）", f"{growth:+.2f}")
        else:
            st.warning(f"未找到{selected_stock}在{selected_year}年的数据")
    else:
        # 显示数据示例和使用说明
        st.info("请在侧边栏选择股票代码或企业名称，并点击'执行查询'按钮查看数据")
        
        # 显示一些数据示例（包含词频字段）
        st.subheader("📊 数据示例（含词频+百分制指数）")
        display_cols = ["股票代码", "企业名称", "年份"] + WORD_FREQ_COLS + ["数字化转型指数"]
        st.dataframe(df[display_cols].head(10), use_container_width=True)
        
        # 使用说明
        st.subheader("📝 使用说明")
        st.markdown("""
        1. 在侧边栏选择搜索方式（股票代码或企业名称）
        2. 选择对应的股票代码或企业名称
        3. 可选：选择特定年份进行查询
        4. 点击'执行查询'按钮
        5. 查看企业历年数字化转型指数（百分制）趋势图和详细数据（含词频）
        
        💡 指数说明：
        - 数字化转型指数为0-100分制，每年词频最高的企业为100分
        - 词频全为0的企业，指数为0分
        - 指数 = (企业当年总词频数/当年行业最高总词频数) × 100
        """)

# 页脚信息
st.markdown("""
---
💡 数据来源：数字化转型指数分析结果.xlsx
📅 更新时间：2025年
📌 指数规则：0-100百分制（按年度归一化）
🔧 运行命令：streamlit run app.py
""")

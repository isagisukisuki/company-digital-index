# streamlit run app.py

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from pathlib import Path
import os

# 设置页面配置
st.set_page_config(
    page_title="数字化转型指数查询系统",
    page_icon="📊",
    layout="wide"
)

# 应用标题
st.title("📊 上市公司数字化转型指数查询系统")
st.markdown("### 查询企业数字化转型指数数据")

# 文件路径（相对于脚本所在目录）
DATA_FILE = Path(__file__).parent / "数字化转型指数分析结果.xlsx"

# 必须保留的列名
REQUIRED_COLUMNS = [
    "股票代码", "企业名称", "年份", "数字化转型综合指数",
    "人工智能词频数", "大数据词频数", "云计算词频数",
    "区块链词频数", "数字技术运用词频数"
]

def normalize_index(df):
    """指数归一化到0-100，确保无负数"""
    if "数字化转型综合指数" not in df.columns:
        return df
    
    # 计算最大最小值（避免除以0）
    idx_col = "数字化转型综合指数"
    min_val = df[idx_col].min()
    max_val = df[idx_col].max()
    
    if max_val - min_val == 0:
        df[idx_col] = 0.0
    else:
        # 归一化公式
        df[idx_col] = (df[idx_col] - min_val) / (max_val - min_val) * 100
    
    # 强制边界：0-100
    df[idx_col] = df[idx_col].clip(0, 100).round(2)
    return df

# 缓存数据加载函数
@st.cache_data
def load_data():
    try:
        # 检查文件是否存在
        if not os.path.exists(DATA_FILE):
            st.error(f"❌ 未找到数据文件：{DATA_FILE}")
            st.error(f"当前目录：{os.getcwd()}")
            return pd.DataFrame(), [], [], [], {}
        
        # 读取Excel所有数字命名的sheet
        excel = pd.ExcelFile(DATA_FILE, engine="openpyxl")
        sheet_names = [s for s in excel.sheet_names if s.isdigit()]
        
        if not sheet_names:
            st.error("❌ Excel中无数字年份命名的工作表（如1999、2000）")
            return pd.DataFrame(), [], [], [], {}
        
        # 读取并合并所有sheet
        df_list = []
        for sheet in sheet_names:
            sheet_df = pd.read_excel(excel, sheet_name=sheet)
            sheet_df["年份"] = sheet  # 添加年份列
            # 只保留需要的列
            available_cols = [col for col in REQUIRED_COLUMNS if col in sheet_df.columns]
            sheet_df = sheet_df[available_cols]
            # 修正股票代码格式
            if "股票代码" in sheet_df.columns:
                sheet_df["股票代码"] = sheet_df["股票代码"].astype(str).str.zfill(6)
            df_list.append(sheet_df)
        
        # 合并数据
        df = pd.concat(df_list, ignore_index=True).fillna(0)
        
        # 归一化指数
        df = normalize_index(df)
        
        # 重命名列名，适配原代码
        column_mapping = {
            "数字化转型综合指数": "数字化转型指数"
        }
        df = df.rename(columns=column_mapping)
        
        # 提取唯一的股票代码、企业名称和年份
        unique_stocks = sorted(df['股票代码'].unique())
        unique_companies = sorted(df['企业名称'].unique())
        unique_years = sorted(df['年份'].unique())
        
        # 创建股票代码到企业名称的映射
        stock_to_company = {}
        for stock in unique_stocks:
            company_name = df[df['股票代码'] == stock]['企业名称'].iloc[0]
            stock_to_company[stock] = company_name
        
        return df, unique_stocks, unique_companies, unique_years, stock_to_company
    except Exception as e:
        st.error(f"加载数据失败: {e}")
        return pd.DataFrame(), [], [], [], {}

# 加载数据
with st.spinner("正在加载数据..."):
    df, unique_stocks, unique_companies, unique_years, stock_to_company = load_data()

# 侧边栏 - 查询控件
with st.sidebar:
    st.header("🔍 查询条件")
    
    # 创建股票代码和企业名称的联合选择器
    search_type = st.radio("搜索方式:", ["股票代码", "企业名称"])
    
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
            match = df[df['企业名称'] == selected_company]
            selected_stock = match['股票代码'].iloc[0] if not match.empty else None
    
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
            
            # 创建折线图
            fig = go.Figure()
            
            # 添加数字化转型指数折线
            fig.add_trace(go.Scatter(
                x=company_history['年份'],
                y=company_history['数字化转型指数'],
                mode='lines+markers',
                name='数字化转型指数',
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
                    text=f'{selected_year}年: {current_value:.2f}',
                    hoverinfo='text'
                ))
            
            # 更新布局
            fig.update_layout(
                title=f'{company_name}历年数字化转型指数趋势',
                xaxis_title='年份',
                yaxis_title='数字化转型指数',
                template='plotly_white',
                height=500,
                legend_title='指标',
                hovermode='x unified'
            )
            
            # 显示图表
            st.plotly_chart(fig, use_container_width=True)
            
            # 显示详细数据
            st.subheader("📊 详细数据")
            if selected_year:
                # 显示特定年份的数据
                st.dataframe(filtered_data, use_container_width=True)
            else:
                # 显示所有年份的数据
                st.dataframe(company_history, use_container_width=True)
                
                # 显示统计信息
                st.subheader("📈 统计分析")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("最高指数", f"{company_history['数字化转型指数'].max():.2f}")
                with col2:
                    st.metric("最低指数", f"{company_history['数字化转型指数'].min():.2f}")
                with col3:
                    st.metric("平均指数", f"{company_history['数字化转型指数'].mean():.2f}")
                with col4:
                    if len(company_history) > 1:
                        growth = company_history['数字化转型指数'].iloc[-1] - company_history['数字化转型指数'].iloc[0]
                        st.metric("指数增长", f"{growth:+.2f}")
                    else:
                        st.metric("指数增长", "N/A")
        else:
            if selected_year:
                st.warning(f"未找到{selected_stock}在{selected_year}年的数据")
            else:
                st.warning(f"未找到股票代码{selected_stock}的数据")
    else:
        # 显示数据示例和使用说明
        st.info("请在侧边栏选择股票代码或企业名称，并点击'执行查询'按钮查看数据")
        
        # 显示一些数据示例
        st.subheader("📊 数据示例")
        st.dataframe(df.head(10), use_container_width=True)
        
        # 使用说明
        st.subheader("📝 使用说明")
        st.markdown("""
        1. 在侧边栏选择搜索方式（股票代码或企业名称）
        2. 选择对应的股票代码或企业名称
        3. 可选：选择特定年份进行查询
        4. 点击'执行查询'按钮
        5. 查看企业历年数字化转型指数趋势图和详细数据
        """)

# 页脚信息
st.markdown("""
---
💡 数据来源：数字化转型指数分析结果.xlsx
📅 更新时间：2024年
""")

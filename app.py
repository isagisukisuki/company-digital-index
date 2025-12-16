# app.py - 数字化转型指数查询系统
# 完全使用Plotly + Streamlit，无Altair依赖

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

# 文件路径
DATA_FILE = Path(__file__).parent / "数字化转型指数分析结果.xlsx"

# 缓存数据加载函数
@st.cache_data
def load_data():
    try:
        # 检查文件是否存在
        if not os.path.exists(DATA_FILE):
            st.error(f"❌ 未找到数据文件：{DATA_FILE}")
            return pd.DataFrame(), [], [], [], {}
        
        # 读取Excel所有sheet
        excel = pd.ExcelFile(DATA_FILE, engine="openpyxl")
        
        # 获取所有sheet名，优先使用数字命名的sheet
        sheet_names = excel.sheet_names
        
        # 读取并合并所有sheet
        df_list = []
        for sheet in sheet_names:
            try:
                sheet_df = pd.read_excel(excel, sheet_name=sheet)
                
                # 添加年份列（使用sheet名或从数据中提取）
                if sheet.isdigit():
                    sheet_df["年份"] = sheet
                else:
                    # 尝试从数据中提取年份
                    if "年份" in sheet_df.columns:
                        sheet_df["年份"] = sheet_df["年份"].astype(str)
                    else:
                        sheet_df["年份"] = sheet
                
                # 标准化列名
                column_mapping = {
                    "股票代码": "股票代码",
                    "企业名称": "企业名称",
                    "年份": "年份",
                    "数字化转型综合指数": "数字化转型指数",
                    "人工智能词频数": "人工智能词频数",
                    "大数据词频数": "大数据词频数", 
                    "云计算词频数": "云计算词频数",
                    "区块链词频数": "区块链词频数",
                    "数字技术运用词频数": "数字技术运用词频数"
                }
                
                # 重命名列
                for old_col, new_col in column_mapping.items():
                    if old_col in sheet_df.columns and new_col not in sheet_df.columns:
                        sheet_df = sheet_df.rename(columns={old_col: new_col})
                
                # 确保必要的列存在
                required_cols = ["股票代码", "企业名称", "年份", "数字化转型指数"]
                for col in required_cols:
                    if col not in sheet_df.columns:
                        if col == "数字化转型指数":
                            # 尝试其他可能的列名
                            for possible_col in ["数字化指数", "转型指数", "数字指数"]:
                                if possible_col in sheet_df.columns:
                                    sheet_df = sheet_df.rename(columns={possible_col: col})
                                    break
                
                # 修正股票代码格式
                if "股票代码" in sheet_df.columns:
                    sheet_df["股票代码"] = sheet_df["股票代码"].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(6)
                
                df_list.append(sheet_df)
                
            except Exception as e:
                st.warning(f"读取sheet '{sheet}' 时出错: {e}")
                continue
        
        if not df_list:
            st.error("❌ 没有成功读取任何sheet")
            return pd.DataFrame(), [], [], [], {}
        
        # 合并数据
        df = pd.concat(df_list, ignore_index=True).fillna(0)
        
        # 提取唯一的股票代码、企业名称和年份
        if "股票代码" in df.columns:
            unique_stocks = sorted(df['股票代码'].astype(str).unique())
        else:
            unique_stocks = []
            
        if "企业名称" in df.columns:
            unique_companies = sorted(df['企业名称'].astype(str).unique())
        else:
            unique_companies = []
            
        if "年份" in df.columns:
            unique_years = sorted(df['年份'].astype(str).unique())
        else:
            unique_years = []
        
        # 创建股票代码到企业名称的映射
        stock_to_company = {}
        if "股票代码" in df.columns and "企业名称" in df.columns:
            for stock in unique_stocks:
                company_name = df[df['股票代码'].astype(str) == stock]['企业名称'].iloc[0] if not df[df['股票代码'].astype(str) == stock].empty else str(stock)
                stock_to_company[stock] = company_name
        
        return df, unique_stocks, unique_companies, unique_years, stock_to_company
        
    except Exception as e:
        st.error(f"加载数据失败: {str(e)}")
        return pd.DataFrame(), [], [], [], {}

# 加载数据
with st.spinner("正在加载数据..."):
    df, unique_stocks, unique_companies, unique_years, stock_to_company = load_data()

# 侧边栏 - 查询控件
with st.sidebar:
    st.header("🔍 查询条件")
    
    # 创建股票代码和企业名称的联合选择器
    search_type = st.radio("搜索方式:", ["股票代码", "企业名称"])
    
    if search_type == "股票代码" and unique_stocks:
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
    elif search_type == "企业名称" and unique_companies:
        selected_company = st.selectbox(
            "选择企业名称:",
            options=unique_companies,
            index=None,
            placeholder="请选择企业名称"
        )
        # 获取对应的股票代码
        if selected_company and "股票代码" in df.columns:
            # 找到第一个匹配的股票代码
            match = df[df['企业名称'].astype(str) == selected_company]
            selected_stock = match['股票代码'].iloc[0] if not match.empty else None
    else:
        selected_stock = None
        selected_company = None
    
    # 年份选择器
    if unique_years:
        selected_year = st.selectbox(
            "选择年份:",
            options=unique_years,
            index=None,
            placeholder="请选择年份(可选)"
        )
    else:
        selected_year = None
    
    # 查询按钮
    search_button = st.button("📈 执行查询")
    
    # 数据概览
    st.header("📊 数据概览")
    if not df.empty:
        st.metric("数据总量", f"{len(df):,}")
        st.metric("企业数量", f"{len(unique_companies):,}")
        if unique_years:
            st.metric("年份跨度", f"{min(unique_years)}-{max(unique_years)}")

# 主页面内容
if df.empty:
    st.warning("暂无数据可供查询，请检查数据文件是否存在且格式正确。")
else:
    # 显示数据预览
    st.subheader("📊 数据预览")
    st.dataframe(df.head(10), use_container_width=True)
    
    # 如果用户点击了查询按钮或选择了股票代码
    if search_button and (selected_stock or selected_company):
        # 筛选数据
        if selected_stock:
            # 按股票代码筛选
            filtered_data = df[df['股票代码'].astype(str) == selected_stock]
        elif selected_company:
            # 按企业名称筛选
            filtered_data = df[df['企业名称'].astype(str) == selected_company]
            if not filtered_data.empty:
                selected_stock = filtered_data['股票代码'].iloc[0]
        else:
            filtered_data = pd.DataFrame()
        
        if selected_year:
            # 按年份筛选
            filtered_data = filtered_data[filtered_data['年份'].astype(str) == selected_year]
        
        if not filtered_data.empty:
            # 获取企业名称
            if selected_stock:
                company_name = stock_to_company.get(selected_stock, selected_stock)
            else:
                company_name = selected_company
            
            # 显示企业信息
            st.subheader(f"📋 {company_name} (股票代码: {selected_stock})")
            
            # 创建历年数据的折线图
            if selected_stock:
                company_history = df[df['股票代码'].astype(str) == selected_stock].copy()
            else:
                company_history = df[df['企业名称'].astype(str) == selected_company].copy()
            
            # 按年份排序并转换为数值
            company_history['年份'] = pd.to_numeric(company_history['年份'], errors='coerce')
            company_history = company_history.sort_values('年份')
            
            # 创建折线图
            if '数字化转型指数' in company_history.columns and len(company_history) > 0:
                fig = go.Figure()
                
                # 添加数字化转型指数折线
                fig.add_trace(go.Scatter(
                    x=company_history['年份'],
                    y=company_history['数字化转型指数'],
                    mode='lines+markers',
                    name='数字化转型指数',
                    line=dict(color='#1f77b4', width=3),
                    marker=dict(size=8, color='#1f77b4', symbol='circle'),
                    hovertemplate='年份: %{x}<br>指数: %{y:.2f}<extra></extra>'
                ))
                
                # 添加当前查询年份的标记点（如果选择了年份）
                if selected_year and selected_year in company_history['年份'].astype(str).values:
                    current_data = company_history[company_history['年份'].astype(str) == selected_year]
                    if not current_data.empty:
                        current_value = current_data['数字化转型指数'].iloc[0]
                        fig.add_trace(go.Scatter(
                            x=[float(selected_year)],
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
            else:
                st.warning("数据中没有找到'数字化转型指数'列")
            
            # 显示详细数据
            st.subheader("📊 详细数据")
            st.dataframe(filtered_data, use_container_width=True)
            
            # 显示统计信息
            if not selected_year and len(company_history) > 0:
                st.subheader("📈 统计分析")
                if '数字化转型指数' in company_history.columns:
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        max_val = company_history['数字化转型指数'].max()
                        st.metric("最高指数", f"{max_val:.2f}")
                    with col2:
                        min_val = company_history['数字化转型指数'].min()
                        st.metric("最低指数", f"{min_val:.2f}")
                    with col3:
                        mean_val = company_history['数字化转型指数'].mean()
                        st.metric("平均指数", f"{mean_val:.2f}")
                    with col4:
                        if len(company_history) > 1:
                            growth = company_history['数字化转型指数'].iloc[-1] - company_history['数字化转型指数'].iloc[0]
                            st.metric("指数增长", f"{growth:+.2f}")
                        else:
                            st.metric("指数增长", "N/A")
        else:
            search_term = selected_stock if selected_stock else selected_company
            if selected_year:
                st.warning(f"未找到{search_term}在{selected_year}年的数据")
            else:
                st.warning(f"未找到{search_term}的数据")
    else:
        # 使用说明
        st.subheader("📝 使用说明")
        st.markdown("""
        1. 在侧边栏选择搜索方式（股票代码或企业名称）
        2. 选择对应的股票代码或企业名称
        3. 可选：选择特定年份进行查询
        4. 点击'执行查询'按钮
        5. 查看企业历年数字化转型指数趋势图和详细数据
        """)
        
        # 显示数据统计
        if '数字化转型指数' in df.columns:
            st.subheader("📈 整体数据统计")
            col1, col2, col3 = st.columns(3)
            with col1:
                overall_avg = df['数字化转型指数'].mean()
                st.metric("整体平均指数", f"{overall_avg:.2f}")
            with col2:
                overall_max = df['数字化转型指数'].max()
                st.metric("最高指数", f"{overall_max:.2f}")
            with col3:
                overall_min = df['数字化转型指数'].min()
                st.metric("最低指数", f"{overall_min:.2f}")

# 页脚信息
st.markdown("""
---
💡 数据来源：数字化转型指数分析结果.xlsx
📅 系统版本：1.0.0
""")

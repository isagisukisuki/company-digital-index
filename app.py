# app.py - 上市公司数字化转型指数查询系统（本地运行版）
from pathlib import Path
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import os

# ====================== 本地文件配置 =======================
# 本地数据文件路径（确保和app.py同目录）
DATA_FILE_NAME = "数字化转型指数分析结果.xlsx"
DATA_FILE = Path(__file__).resolve().parent / DATA_FILE_NAME

# 词频字段配置
WORD_FREQ_COLS = [
    "人工智能词频数",
    "大数据词频数",
    "云计算词频数",
    "区块链词频数",
    "数字技术运用词频数"
]
# ==========================================================

# 基础设置
st.set_page_config(
    page_title="数字化转型指数查询系统",
    page_icon="📊",
    layout="wide"
)

# 检查文件是否存在（本地运行关键）
def check_file_exists():
    if not os.path.exists(DATA_FILE):
        st.error(f"❌ 未找到数据文件！请确认 {DATA_FILE_NAME} 放在app.py同一目录下")
        st.info(f"当前脚本路径：{Path(__file__).resolve().parent}")
        st.info(f"期望文件路径：{DATA_FILE}")
        return False
    return True

# 核心：计算百分制指数（按年度归一化）
def calculate_percentile_index(df):
    # 计算总词频数
    df["年度总词频数"] = df[WORD_FREQ_COLS].sum(axis=1)
    
    # 按年份分组计算百分制
    def _yearly_calc(year_df):
        max_total = year_df["年度总词频数"].max()
        if max_total == 0:
            year_df["数字化转型指数"] = 0.0
        else:
            # 归一化到0-100分
            year_df["数字化转型指数"] = (year_df["年度总词频数"] / max_total * 100).round(2)
        # 强制边界：0-100，词频为0则指数为0
        year_df["数字化转型指数"] = year_df["数字化转型指数"].clip(0, 100)
        year_df.loc[year_df["年度总词频数"] == 0, "数字化转型指数"] = 0.0
        return year_df
    
    df = df.groupby("年份", group_keys=False).apply(_yearly_calc)
    return df.drop("年度总词频数", axis=1)

# 本地数据加载（无缓存也可，本地文件读取更快）
def load_local_data():
    try:
        # 读取Excel多sheet（sheet名为年份数字）
        excel = pd.ExcelFile(DATA_FILE, engine="openpyxl")
        sheet_names = [s for s in excel.sheet_names if s.isdigit()]
        
        if not sheet_names:
            st.error("❌ Excel文件中无数字命名的sheet！请确保sheet名为年份（如1999、2000）")
            return pd.DataFrame(), [], [], [], {}
        
        # 合并所有年份数据
        df_list = []
        for sheet in sheet_names:
            df_sheet = pd.read_excel(DATA_FILE, sheet_name=sheet, engine="openpyxl")
            df_sheet["年份"] = int(sheet)  # 转为数字年份
            df_list.append(df_sheet)
        
        df = pd.concat(df_list, ignore_index=True).fillna(0)
        
        # 修正股票代码格式（6位补零）
        if "股票代码" in df.columns:
            df["股票代码"] = df["股票代码"].astype(str).str.zfill(6)
        
        # 计算百分制指数（覆盖原始指数）
        df = calculate_percentile_index(df)
        
        # 提取唯一值
        unique_stocks = sorted(df["股票代码"].unique())
        unique_companies = sorted(df["企业名称"].unique())
        unique_years = sorted(df["年份"].unique())
        
        # 股票代码→企业名称映射
        stock2company = {}
        for stock in unique_stocks:
            company = df[df["股票代码"] == stock]["企业名称"].iloc[0]
            stock2company[stock] = company
        
        return df, unique_stocks, unique_companies, unique_years, stock2company
    
    except Exception as e:
        st.error(f"❌ 加载数据失败：{str(e)}")
        st.error("可能原因：1.Excel格式错误 2.缺少列名（股票代码/企业名称/词频列）")
        return pd.DataFrame(), [], [], [], {}

# ============ 主页面逻辑 ============
st.title("📊 上市公司数字化转型指数查询系统")
st.markdown("### 本地版 | 1999-2023年数据（百分制）")

# 第一步：检查文件
if not check_file_exists():
    st.stop()

# 第二步：加载数据
with st.spinner("📥 正在加载本地数据..."):
    df, unique_stocks, unique_companies, unique_years, stock2company = load_local_data()

# 数据为空则停止
if df.empty:
    st.warning("📭 暂无有效数据，请检查Excel文件内容")
    st.stop()

# ============ 侧边栏查询 ============
with st.sidebar:
    st.header("🔍 查询条件")
    
    # 搜索方式选择
    search_type = st.radio("搜索方式", ["股票代码", "企业名称"], index=0)
    
    selected_stock = None
    selected_company = None
    
    # 股票代码搜索
    if search_type == "股票代码":
        selected_stock = st.selectbox(
            "选择股票代码",
            options=unique_stocks,
            format_func=lambda x: f"{x} - {stock2company.get(x, '未知')}",
            index=None,
            placeholder="输入/选择股票代码"
        )
        if selected_stock:
            selected_company = stock2company.get(selected_stock, "")
    
    # 企业名称搜索
    else:
        selected_company = st.selectbox(
            "选择企业名称",
            options=unique_companies,
            index=None,
            placeholder="输入/选择企业名称"
        )
        if selected_company:
            # 找到对应股票代码
            mask = df["企业名称"] == selected_company
            if mask.any():
                selected_stock = df[mask]["股票代码"].iloc[0]
    
    # 年份选择
    selected_year = st.selectbox(
        "选择年份（可选）",
        options=unique_years,
        index=None,
        placeholder="不选则显示所有年份"
    )
    
    # 查询按钮
    search_btn = st.button("📈 执行查询", type="primary")

# ============ 数据概览 ============
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("📊 总数据量", f"{len(df):,} 条")
with col2:
    st.metric("🏢 企业数量", f"{len(unique_companies):,} 家")
with col3:
    st.metric("📅 年份范围", f"{min(unique_years)}-{max(unique_years)}")

# ============ 查询结果 ============
if search_btn and selected_stock:
    # 筛选数据
    if selected_year:
        filter_df = df[(df["股票代码"] == selected_stock) & (df["年份"] == selected_year)]
    else:
        filter_df = df[df["股票代码"] == selected_stock]
    
    if filter_df.empty:
        st.warning(f"⚠️ 未找到 {selected_stock}（{selected_company}）在 {selected_year if selected_year else '所有年份'} 的数据")
    else:
        # 企业基本信息
        company_name = filter_df["企业名称"].iloc[0]
        st.subheader(f"📋 {company_name}（{selected_stock}）")
        
        # 历年数据趋势图
        history_df = df[df["股票代码"] == selected_stock].sort_values("年份")
        
        # 绘制Plotly折线图（本地显示优化）
        fig = go.Figure()
        # 主趋势线
        fig.add_trace(go.Scatter(
            x=history_df["年份"],
            y=history_df["数字化转型指数"],
            mode="lines+markers",
            name="数字化转型指数",
            line=dict(color="#2E86AB", width=3),
            marker=dict(size=8, color="#2E86AB"),
            hovertemplate="年份：%{x}<br>指数：%{y:.2f}分<extra></extra>"
        ))
        
        # 选中年份标记
        if selected_year:
            current_val = filter_df["数字化转型指数"].iloc[0]
            fig.add_trace(go.Scatter(
                x=[selected_year],
                y=[current_val],
                mode="markers",
                name=f"{selected_year}年",
                marker=dict(size=15, color="#E63946", symbol="star"),
                hovertemplate=f"{selected_year}年：{current_val:.2f}分<extra></extra>"
            ))
        
        # 图表布局（强制0-100分）
        fig.update_layout(
            title=f"{company_name} 历年数字化转型指数趋势",
            xaxis_title="年份",
            yaxis_title="数字化转型指数（0-100分）",
            yaxis_range=[0, 100],
            height=500,
            template="simple_white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # 详细数据展示
        st.subheader("📊 详细数据（含词频）")
        show_cols = ["股票代码", "企业名称", "年份"] + WORD_FREQ_COLS + ["数字化转型指数"]
        st.dataframe(
            filter_df[show_cols].reset_index(drop=True),
            use_container_width=True,
            column_config={
                "数字化转型指数": st.column_config.NumberColumn("数字化转型指数（分）", format="%.2f")
            }
        )
        
        # 统计分析（仅当查询所有年份时）
        if not selected_year:
            st.subheader("📈 统计分析（百分制）")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("最高指数", f"{history_df['数字化转型指数'].max():.2f} 分")
            with col2:
                st.metric("最低指数", f"{history_df['数字化转型指数'].min():.2f} 分")
            with col3:
                st.metric("平均指数", f"{history_df['数字化转型指数'].mean():.2f} 分")
            with col4:
                growth = history_df['数字化转型指数'].iloc[-1] - history_df['数字化转型指数'].iloc[0]
                st.metric("整体增长", f"{growth:+.2f} 分")

# ============ 未查询时显示示例 ============
else:
    st.info("💡 请在左侧边栏选择股票代码/企业名称，点击「执行查询」查看数据")
    
    # 数据示例
    st.subheader("📌 数据示例（前10条）")
    sample_cols = ["股票代码", "企业名称", "年份"] + WORD_FREQ_COLS + ["数字化转型指数"]
    st.dataframe(
        df[sample_cols].head(10).reset_index(drop=True),
        use_container_width=True,
        column_config={
            "数字化转型指数": st.column_config.NumberColumn("数字化转型指数（分）", format="%.2f")
        }
    )
    
    # 本地运行说明
    st.subheader("📝 本地运行说明")
    st.markdown("""
    1. 确保「数字化转型指数分析结果.xlsx」与app.py在同一文件夹
    2. Excel中sheet名必须为数字年份（如1999、2000、2023）
    3. 数据列必须包含：股票代码、企业名称、人工智能词频数、大数据词频数、云计算词频数、区块链词频数、数字技术运用词频数
    4. 指数计算规则：按年度归一化，每年词频最高企业为100分，词频为0则为0分
    """)

# 页脚
st.markdown("---")
st.markdown("✅ 本地运行版 | 指数已统一为0-100百分制 | 无负数、无极低值")

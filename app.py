import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os

# ====================== 全局配置 =======================
# 解决中文显示问题
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
st.set_page_config(page_title="企业数字化转型指数查询", page_icon="📊", layout="wide")

# 数据文件路径
DATA_FILE = "数字化转型指数分析结果.xlsx"

# 必须保留的列名
REQUIRED_COLUMNS = [
    "股票代码", "企业名称", "年份", "数字化转型综合指数",
    "人工智能词频数", "大数据词频数", "云计算词频数",
    "区块链词频数", "数字技术运用词频数"
]

# ====================== 核心函数 =======================
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

def load_data():
    """读取并预处理数据"""
    # 检查文件是否存在
    if not os.path.exists(DATA_FILE):
        st.error(f"❌ 未找到数据文件：{DATA_FILE}")
        st.error(f"当前目录：{os.getcwd()}")
        return pd.DataFrame()
    
    try:
        # 读取Excel所有数字命名的sheet
        excel = pd.ExcelFile(DATA_FILE, engine="openpyxl")
        sheet_names = [s for s in excel.sheet_names if s.isdigit()]
        
        if not sheet_names:
            st.error("❌ Excel中无数字年份命名的工作表（如1999、2000）")
            return pd.DataFrame()
        
        # 读取并合并所有sheet
        df_list = []
        for sheet in sheet_names:
            sheet_df = pd.read_excel(excel, sheet_name=sheet)
            sheet_df["年份"] = sheet  # 添加年份列
            # 只保留需要的列
            sheet_df = sheet_df[[col for col in REQUIRED_COLUMNS if col in sheet_df.columns]]
            # 修正股票代码格式
            if "股票代码" in sheet_df.columns:
                sheet_df["股票代码"] = sheet_df["股票代码"].astype(str).str.zfill(6)
            df_list.append(sheet_df)
        
        # 合并并归一化
        full_df = pd.concat(df_list, ignore_index=True).fillna(0)
        full_df = normalize_index(full_df)
        
        return full_df
    
    except Exception as e:
        st.error(f"❌ 读取数据失败：{str(e)}")
        return pd.DataFrame()

def generate_excel(df):
    """生成Excel下载文件"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="数据")
    return output.getvalue()

def generate_report(company_name, company_data, trend_data):
    """生成企业分析报告"""
    stock_code = company_data["股票代码"].iloc[0] if not company_data.empty else "未知"
    years = sorted(company_data["年份"].unique()) if not company_data.empty else []
    
    # 指数分析
    idx_col = "数字化转型综合指数"
    max_idx = company_data[idx_col].max() if not company_data.empty else 0
    max_year = company_data[company_data[idx_col]==max_idx]["年份"].iloc[0] if not company_data.empty else "无"
    avg_idx = company_data[idx_col].mean() if not company_data.empty else 0
    latest_year = max(years) if years else "无"
    latest_idx = company_data[company_data["年份"]==latest_year][idx_col].iloc[0] if years else 0
    
    # 趋势计算
    trend = "无数据"
    if len(years)>=2:
        first_idx = company_data[company_data["年份"]==min(years)][idx_col].iloc[0]
        if first_idx != 0:
            growth = ((latest_idx - first_idx)/first_idx)*100
            trend = f"上升（{growth:.2f}%）" if growth>0 else f"下降（{growth:.2f}%）" if growth<0 else "平稳"
    
    # 词频分析
    freq_cols = [col for col in REQUIRED_COLUMNS if col.endswith("词频数")]
    freq_data = {col: company_data[col].mean() for col in freq_cols} if not company_data.empty else {}
    
    # 生成报告文本
    report = f"""# {company_name} 数字化转型分析报告
**生成时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
**股票代码**：{stock_code}

## 一、基础信息
- 数据覆盖年份：{years if years else '无'}
- 有效年份数：{len(years)}

## 二、核心指数分析
- 历史最高指数：{max_idx:.2f}（{max_year}年）
- 历年平均指数：{avg_idx:.2f}
- 最新指数（{latest_year}年）：{latest_idx:.2f}
- 整体趋势：{trend}

## 三、技术词频分析（均值）
{chr(10).join([f"- {col}：{freq_data.get(col, 0):.2f}" for col in freq_cols])}

## 四、指数明细
{trend_data.round(2).to_string(index=False)}

## 五、说明
1. 指数取值范围0-100，越高代表转型程度越高
2. 指数已做归一化处理，无负数
"""
    return report

# ====================== 主程序 =======================
def main():
    st.title("📊 企业数字化转型指数查询系统")
    
    # 1. 加载数据
    full_data = load_data()
    if full_data.empty:
        return
    
    # 2. 获取基础信息
    all_years = sorted(full_data["年份"].unique())
    if not all_years:
        st.error("❌ 数据中无有效年份")
        return
    
    # 3. 查询区域
    st.subheader("🔍 企业查询")
    col1, col2, col3 = st.columns(3)
    with col1:
        stock_code = st.text_input("股票代码", placeholder="如：000001")
    with col2:
        company_name = st.text_input("企业名称", placeholder="如：平安银行")
    with col3:
        selected_year = st.selectbox("查询年份", all_years, index=0)
    
    # 4. 筛选当年数据
    year_filter = full_data["年份"] == selected_year
    year_data = full_data[year_filter].copy()
    
    # 5. 筛选企业数据
    company_data = pd.DataFrame()
    if stock_code:
        company_data = full_data[(full_data["股票代码"] == stock_code.strip().zfill(6)) & year_filter].copy()
    elif company_name:
        company_data = full_data[(full_data["企业名称"].str.contains(company_name.strip())) & year_filter].copy()
    
    # 6. 展示当年数据
    st.success(f"✅ 已加载{selected_year}年数据（总计{len(year_data)}家企业）")
    st.subheader("📋 当年企业数据")
    
    # 应用筛选条件
    display_data = year_data.copy()
    if stock_code:
        display_data = display_data[display_data["股票代码"] == stock_code.strip().zfill(6)]
    if company_name:
        display_data = display_data[display_data["企业名称"].str.contains(company_name.strip())]
    
    st.dataframe(display_data, use_container_width=True)
    st.info(f"筛选结果：{len(display_data)}家企业")
    
    # 7. 全行业趋势图
    st.subheader("📈 全行业指数趋势")
    industry_trend = []
    for year in all_years:
        avg_idx = full_data[full_data["年份"]==year]["数字化转型综合指数"].mean()
        industry_trend.append({"年份": year, "平均指数": avg_idx})
    industry_df = pd.DataFrame(industry_trend)
    
    # 使用Streamlit原生折线图
    if not industry_df.empty:
        # 准备图表数据
        chart_data = industry_df.set_index("年份")[["平均指数"]]
        st.line_chart(
            chart_data,
            height=400,
            use_container_width=True
        )
    
    # 8. 企业趋势分析（仅当找到企业时）
    if not company_data.empty:
        # 获取企业名称
        comp_name = company_data["企业名称"].iloc[0] if not company_data.empty else "未知企业"
        comp_code = company_data["股票代码"].iloc[0] if not company_data.empty else "未知代码"
        
        # 准备企业趋势数据
        comp_trend = []
        for year in all_years:
            year_data = full_data[(full_data["股票代码"] == comp_code) & (full_data["年份"] == year)]
            idx_val = year_data["数字化转型综合指数"].iloc[0] if not year_data.empty else 0
            comp_trend.append({"年份": year, "数字化转型综合指数": idx_val})
        comp_trend_df = pd.DataFrame(comp_trend)
        
        # 展示企业趋势图（如果有多年的数据）
        if len(comp_trend_df) > 1:
            st.subheader(f"📈 {comp_name}（{comp_code}）指数趋势")
            chart_data = comp_trend_df.set_index("年份")[["数字化转型综合指数"]]
            st.line_chart(
                chart_data,
                height=400,
                use_container_width=True
            )
        elif len(comp_trend_df) == 1:
            st.subheader(f"📊 {comp_name}（{comp_code}）指数信息")
            st.info(f"当前只有一年的数据：{comp_trend_df['数字化转型综合指数'].iloc[0]:.2f}")
        
        # 展示企业历年数据
        st.subheader(f"📋 {comp_name} 历年完整数据")
        comp_all_data = full_data[full_data["股票代码"] == comp_code].copy()
        st.dataframe(comp_all_data, use_container_width=True)
        
        # 下载功能
        st.subheader("📥 数据下载")
        report_text = generate_report(comp_name, comp_all_data, comp_trend_df)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button(
                label="📄 下载报告(TXT)",
                data=report_text,
                file_name=f"{comp_name}_转型报告_{datetime.now().strftime('%Y%m%d')}.txt",
                mime="text/plain"
            )
        with col2:
            st.download_button(
                label="📊 下载趋势数据(Excel)",
                data=generate_excel(comp_trend_df),
                file_name=f"{comp_name}_趋势数据.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col3:
            st.download_button(
                label="📋 下载历年数据(Excel)",
                data=generate_excel(comp_all_data),
                file_name=f"{comp_name}_历年数据.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # 9. 侧边栏统计信息
    with st.sidebar:
        st.header("📊 数据概览")
        st.metric("数据年份数", len(all_years))
        st.metric("企业总数", len(full_data["股票代码"].unique()))
        st.metric("数据总条数", len(full_data))
        
        st.header("🔧 数据操作")
        if st.button("🔄 重新加载数据"):
            st.rerun()
        
        st.header("📖 使用说明")
        st.info("""
        1. 通过股票代码或企业名称查询
        2. 选择年份查看特定年份数据
        3. 选中企业后可查看趋势分析
        4. 支持数据下载功能
        """)

# ====================== 运行程序 =======================
if __name__ == "__main__":
    main()

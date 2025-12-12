import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os

# 全局设置：解决中文显示 + 优化Pandas性能
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('mode.chained_assignment', None)  # 关闭不必要的警告


# ====================== 路径配置（相对路径，适配云端）======================
DIGITAL_TRANSFORMATION_FILE = "数字化转型指数分析结果.xlsx"
# =====================================================================


# 工具函数：生成Excel下载文件（复用逻辑）
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False, sheet_name='数据')
    return output.getvalue()


# 生成企业综合报告（适配你的列名）
def generate_company_report(company_name, company_data, full_trend_data):
    stock_code = company_data["股票代码"].iloc[0] if ("股票代码" in company_data.columns and not company_data.empty) else "未知"
    available_years = sorted(company_data["年份"].unique()) if ("年份" in company_data.columns and not company_data.empty) else []
    total_years = len(available_years)
    
    index_analysis = {"max_index":0, "max_year":"无", "avg_index":0, "latest_index":0, "trend":"无数据"}
    # 适配列名：数字化转型综合指数
    if "数字化转型综合指数" in company_data.columns and not company_data.empty:
        max_val = company_data["数字化转型综合指数"].max()
        max_year_df = company_data[company_data["数字化转型综合指数"] == max_val]
        index_analysis["max_index"] = round(max_val, 4)
        index_analysis["max_year"] = max_year_df["年份"].iloc[0] if not max_year_df.empty else "无"
        index_analysis["avg_index"] = round(company_data["数字化转型综合指数"].mean(), 4)
        index_analysis["latest_year"] = max(available_years) if available_years else "无"
        latest_df = company_data[company_data["年份"] == index_analysis["latest_year"]]
        index_analysis["latest_index"] = round(latest_df["数字化转型综合指数"].iloc[0], 4) if not latest_df.empty else 0
        
        if len(available_years) >= 2:
            first_df = company_data[company_data["年份"] == min(available_years)]
            first_index = first_df["数字化转型综合指数"].iloc[0] if not first_df.empty else 0
            if first_index != 0:
                growth_rate = round(((index_analysis["latest_index"] - first_index)/first_index)*100, 2)
                index_analysis["trend"] = f"上升（{growth_rate}%）" if growth_rate > 0 else f"下降（{growth_rate}%）" if growth_rate < 0 else "平稳"
            else:
                index_analysis["trend"] = "数据基数为0，无法计算趋势"
    
    # 适配列名：你的词频列名
    word_freq_cols = ["人工智能词频数", "大数据词频数", "云计算词频数", "区块链词频数", "数字技术运用词频数"]
    word_freq_data = {col: 0 for col in word_freq_cols}
    if not company_data.empty:
        for col in word_freq_cols:
            if col in company_data.columns:
                word_freq_data[col] = round(company_data[col].mean(), 2)
    
    report = f"""# {company_name} 数字化转型综合分析报告
**报告生成时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
**股票代码**：{stock_code}

## 一、基础信息
- 数据覆盖年份：{available_years if available_years else '无'}
- 有效数据年份数：{total_years}

## 二、核心转型指数分析
- 历史最高指数：{index_analysis['max_index']}（{index_analysis['max_year']}年）
- 历年平均指数：{index_analysis['avg_index']}
- 最新年份（{index_analysis['latest_year']}）指数：{index_analysis['latest_index']}
- 整体趋势：{index_analysis['trend']}

## 三、技术词频分析（历年均值）
- 人工智能词频数：{word_freq_data['人工智能词频数']}
- 大数据词频数：{word_freq_data['大数据词频数']}
- 云计算词频数：{word_freq_data['云计算词频数']}
- 区块链词频数：{word_freq_data['区块链词频数']}
- 数字技术运用词频数：{word_freq_data['数字技术运用词频数']}

## 四、完整指数明细
{full_trend_data.round(4).to_string(index=False)}

## 五、数据说明
1. 数字化转型综合指数越高代表转型程度越高
2. 词频数据反映对应技术的应用强度
"""
    return report, full_trend_data


# 读取完整数据（优化：读取所有工作表+快速合并+数据类型优化+空值处理）
@st.cache_data(ttl=3600)  # 缓存数据，大幅提升加载速度
def load_full_data(file_path):
    try:
        # 读取所有工作表（sheet_name=None），并快速合并
        all_sheets = pd.read_excel(
            file_path,
            sheet_name=None,
            engine="openpyxl",
            dtype={  # 指定数据类型，减少内存占用+加快读取
                "股票代码": str,
                "年份": str,
                "企业名称": str,
                "人工智能词频数": np.int32,
                "大数据词频数": np.int32,
                "云计算词频数": np.int32,
                "区块链词频数": np.int32,
                "数字技术运用词频数": np.int32,
                "总词频数": np.int32,
                "数字化转型综合指数": np.float32
            }
        )
        
        # 合并所有工作表（添加“年份”列为工作表名）
        merged_df = pd.concat(
            [sheet.assign(年份=sheet_name) for sheet_name, sheet in all_sheets.items()],
            ignore_index=True
        )
        
        # 处理空值：将“总词频数”的空值替换为0
        if "总词频数" in merged_df.columns:
            merged_df["总词频数"] = merged_df["总词频数"].fillna(0).astype(np.int32)
        
        # 轻量清洗（仅去空行）
        merged_df = merged_df.dropna(how="all").reset_index(drop=True)
        return merged_df
    except Exception as e:
        st.error(f"❌ 读取数据失败：{str(e)}")
        return pd.DataFrame()


# 获取所有年份（复用逻辑）
def get_all_years(full_data):
    if "年份" not in full_data.columns:
        st.error("❌ 数据中未找到'年份'列")
        return []
    return sorted(full_data["年份"].unique())


def main():
    st.set_page_config(page_title="数字化转型查询系统", layout="wide")  # 宽布局，提升显示效率
    st.title("企业数字化转型指数查询系统")
    
    # 验证文件存在性
    if not os.path.exists(DIGITAL_TRANSFORMATION_FILE):
        st.error(f"❌ 文件不存在：{DIGITAL_TRANSFORMATION_FILE}")
        st.info("请确认：1. 数据文件与app.py在同一目录；2. 文件名拼写正确")
        return
    
    # 读取数据（缓存后仅加载一次）
    full_data = load_full_data(DIGITAL_TRANSFORMATION_FILE)
    if full_data.empty:
        st.error("❌ 数据为空，请检查Excel文件内容")
        return

    # 获取所有年份
    all_years = get_all_years(full_data)
    if not all_years:
        st.error("❌ 数据中无有效年份")
        return

    # 查询区域（布局紧凑）
    st.subheader("🔍 企业查询")
    col1, col2, col3 = st.columns([1,2,1])
    with col1:
        stock_code = st.text_input("股票代码（如：000001）", placeholder="输入股票代码")
    with col2:
        company_name = st.text_input("企业名称（如：平安银行）", placeholder="输入企业名称")
    with col3:
        selected_year = st.selectbox("查询年份", all_years, index=len(all_years)-1)  # 默认选最新年

    # 筛选企业全量数据（提前筛选，减少后续计算）
    company_all_data = pd.DataFrame()
    if stock_code:
        company_all_data = full_data[full_data["股票代码"] == stock_code.strip()].copy()
    elif company_name:
        company_all_data = full_data[full_data["企业名称"].str.contains(company_name.strip(), na=False)].copy()

    # 筛选当年数据（快速过滤）
    current_year_data = full_data[full_data["年份"] == selected_year].copy()
    current_filtered_data = current_year_data.copy()
    if stock_code:
        current_filtered_data = current_filtered_data[current_filtered_data["股票代码"] == stock_code.strip()]
    if company_name:
        current_filtered_data = current_filtered_data[current_filtered_data["企业名称"].str.contains(company_name.strip(), na=False)]

    # 展示当年数据（适配你的列名）
    st.success(f"✅ 已查询{selected_year}年数据（总计{len(current_year_data)}家企业）")
    st.subheader("📋 企业当年详细数据")
    # 只显示关键列，减少渲染压力
    display_cols = ["股票代码", "企业名称", "年份", "人工智能词频数", "大数据词频数", "云计算词频数", "区块链词频数", "数字技术运用词频数", "总词频数", "数字化转型综合指数"]
    display_cols = [col for col in display_cols if col in current_filtered_data.columns]
    st.dataframe(current_filtered_data[display_cols], use_container_width=True, height=300)  # 限制高度，加快渲染


    # 全行业趋势图（复用逻辑）
    st.subheader("📊 全行业转型指数趋势")
    industry_avg_data = []
    for year in all_years:
        year_data = full_data[full_data["年份"] == year]
        if not year_data.empty and "数字化转型综合指数" in year_data.columns:
            avg_idx = year_data["数字化转型综合指数"].mean()
        else:
            avg_idx = 0
        industry_avg_data.append({"年份": year, "平均指数": round(avg_idx, 4)})
    industry_avg_df = pd.DataFrame(industry_avg_data)
    st.line_chart(industry_avg_df.set_index("年份")["平均指数"], use_container_width=True, height=400)


    # 企业趋势图+下载（有数据才显示）
    if not company_all_data.empty:
        selected_company = company_all_data["企业名称"].iloc[0] if not company_all_data["企业名称"].empty else "未知企业"
        st.subheader(f"📈 {selected_company} 转型指数趋势")
        
        # 企业趋势数据
        full_years_df = pd.DataFrame({"年份": all_years})
        company_trend = pd.merge(
            full_years_df,
            company_all_data[["年份", "数字化转型综合指数"]],
            on="年份",
            how="left"
        ).fillna(0)
        st.line_chart(company_trend.set_index("年份")["数字化转型综合指数"], use_container_width=True, height=400)
        
        # 企业历年数据
        st.subheader(f"📋 {selected_company} 历年完整数据")
        st.dataframe(company_all_data[display_cols].sort_values("年份"), use_container_width=True, height=300)
        
        # 下载功能（紧凑布局）
        st.subheader("📥 报告下载")
        report_text, report_data = generate_company_report(selected_company, company_all_data, company_trend)
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            st.download_button("报告（TXT）", data=report_text, file_name=f"{selected_company}_报告_{datetime.now().strftime('%Y%m%d')}.txt")
        with col_r2:
            st.download_button("趋势数据（Excel）", data=to_excel(report_data), file_name=f"{selected_company}_趋势数据.xlsx")
        with col_r3:
            st.download_button("历年数据（Excel）", data=to_excel(company_all_data[display_cols]), file_name=f"{selected_company}_历年数据.xlsx")
    elif stock_code or company_name:
        st.warning("⚠️ 未找到匹配的企业数据，请检查股票代码/名称")


if __name__ == "__main__":
    main()

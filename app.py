import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

# 全局设置：解决中文显示/对齐问题
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)

# 工具函数：生成Excel下载文件
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='数据')
    writer.close()
    return output.getvalue()

# 生成企业综合报告（修复索引错误）
def generate_company_report(company_name, company_data, full_trend_data):
    stock_code = company_data["股票代码"].iloc[0] if ("股票代码" in company_data.columns and not company_data.empty) else "未知"
    available_years = sorted(company_data["年份"].unique()) if ("年份" in company_data.columns and not company_data.empty) else []
    total_years = len(available_years)
    
    index_analysis = {"max_index":0, "max_year":"无", "avg_index":0, "latest_index":0, "trend":"无数据"}
    if "数字化转型综合指数" in company_data.columns and not company_data.empty:
        max_val = company_data["数字化转型综合指数"].max()
        # 修复索引越界：先筛选非空数据
        max_year_df = company_data[company_data["数字化转型综合指数"] == max_val]
        index_analysis["max_index"] = round(max_val, 2)
        index_analysis["max_year"] = max_year_df["年份"].iloc[0] if not max_year_df.empty else "无"
        index_analysis["avg_index"] = round(company_data["数字化转型综合指数"].mean(), 2)
        index_analysis["latest_year"] = max(available_years) if available_years else "无"
        latest_df = company_data[company_data["年份"] == index_analysis["latest_year"]]
        index_analysis["latest_index"] = round(latest_df["数字化转型综合指数"].iloc[0], 2) if not latest_df.empty else 0
        
        if len(available_years) >= 2:
            first_df = company_data[company_data["年份"] == min(available_years)]
            first_index = first_df["数字化转型综合指数"].iloc[0] if not first_df.empty else 0
            if first_index != 0:
                growth_rate = round(((index_analysis["latest_index"] - first_index)/first_index)*100, 2)
                index_analysis["trend"] = f"上升（{growth_rate}%）" if growth_rate > 0 else f"下降（{growth_rate}%）" if growth_rate < 0 else "平稳"
            else:
                index_analysis["trend"] = "数据基数为0，无法计算趋势"
    
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

## 四、1999-2023年完整指数明细
{full_trend_data.round(2).to_string(index=False)}

## 五、数据说明
1. 数字化转型综合指数越高代表转型程度越高
2. 词频数据反映对应技术的应用强度
3. 无数据年份指数统一填充为0
"""
    return report, full_trend_data

# 读取指定年份工作表的数据（精准匹配）
def load_year_data(file_path, year):
    try:
        df = pd.read_excel(
            file_path,
            sheet_name=str(year),
            engine="openpyxl"
        )
        df["年份"] = year
        # 清洗列名和数据
        if "企业名称" in df.columns:
            df["企业名称"] = df["企业名称"].str.strip()
        if "股票代码" in df.columns:
            df["股票代码"] = df["股票代码"].astype(str).str.strip()
        return df.dropna(how="all").reset_index(drop=True)
    except Exception as e:
        st.warning(f"⚠️ {year}年工作表读取失败：{str(e)}")
        return pd.DataFrame()

def main():
    st.title("企业数字化转型指数查询系统")
    # 核心修改：替换为云端相对路径
    file_path = "./数字化转型指数分析结果.xlsx"
    
    # 1. 获取所有有效年份（工作表名）
    try:
        excel_file = pd.ExcelFile(file_path, engine="openpyxl")
        valid_years = [int(s) for s in excel_file.sheet_names if s.isdigit() and 1999 <= int(s) <= 2023]
        if not valid_years:
            st.error("❌ 未找到1999-2023年的工作表")
            return
        valid_years.sort()
    except Exception as e:
        st.error(f"❌ 读取Excel失败：{str(e)}")
        return

    # 2. 查询区域（合并名称+股票代码查询）
    st.subheader("🔍 企业查询（名称/股票代码）")
    col1, col2, col3 = st.columns(3)
    with col1:
        selected_year = st.selectbox("选择查询年份", valid_years, index=0)
    with col2:
        company_name = st.text_input("输入企业名称（如：ST中绒）")
    with col3:
        stock_code = st.text_input("输入股票代码（如：000514）")

    # 3. 读取当前年份工作表数据
    current_year_data = load_year_data(file_path, selected_year)
    if current_year_data.empty:
        st.info(f"ℹ️ {selected_year}年工作表无数据")
        return

    # 4. 多条件筛选（名称/股票代码）
    filtered_data = current_year_data.copy()
    if company_name:
        filtered_data = filtered_data[filtered_data["企业名称"].str.contains(company_name.strip(), na=False)]
    if stock_code:
        filtered_data = filtered_data[filtered_data["股票代码"] == stock_code.strip()]

    # 5. 展示当年数据
    st.success(f"✅ 已查询{selected_year}年数据")
    st.subheader("📋 企业当年详细数据")
    if not filtered_data.empty:
        st.dataframe(filtered_data, use_container_width=True)
    else:
        st.info(f"ℹ️ {selected_year}年工作表中无匹配数据")

    # 6. 新增：全行业1999-2023平均指数折线图
    st.subheader("📊 全行业1999-2023年平均转型指数趋势")
    industry_avg_data = []
    for year in range(1999, 2024):
        year_df = load_year_data(file_path, year)
        avg_idx = year_df["数字化转型综合指数"].mean() if ("数字化转型综合指数" in year_df.columns and not year_df.empty) else 0
        industry_avg_data.append({"年份": year, "平均指数": avg_idx})
    industry_avg_df = pd.DataFrame(industry_avg_data)
    st.line_chart(
        industry_avg_df.set_index("年份")["平均指数"],
        use_container_width=True,
        color="#2E86AB",
        height=400
    )

    # 7. 企业全量趋势图（跨年份）
    if not filtered_data.empty:
        target_company = filtered_data["企业名称"].iloc[0]
        # 读取该企业所有年份数据
        company_all_data = []
        for year in valid_years:
            year_df = load_year_data(file_path, year)
            if not year_df.empty and "企业名称" in year_df.columns:
                comp_df = year_df[year_df["企业名称"] == target_company]
                if not comp_df.empty:
                    company_all_data.append(comp_df)
        if company_all_data:
            company_all_data = pd.concat(company_all_data, ignore_index=True)
            # 补全1999-2023年份
            full_years = pd.DataFrame({"年份": range(1999, 2024)})
            company_trend = pd.merge(
                full_years,
                company_all_data[["年份", "数字化转型综合指数"]],
                on="年份",
                how="left"
            ).fillna(0)

            st.subheader(f"📈 {target_company} 1999-2023转型指数趋势")
            st.line_chart(
                company_trend.set_index("年份")["数字化转型综合指数"],
                use_container_width=True,
                color="#FF6B6B",
                height=500
            )

            # 报告下载
            st.subheader("📥 综合报告下载")
            report_text, report_data = generate_company_report(target_company, company_all_data, company_trend)
            col_r1, col_r2 = st.columns(2)
            with col_r1:
                st.download_button(
                    label="📄 下载报告（TXT）",
                    data=report_text,
                    file_name=f"{target_company}_报告_{datetime.now().strftime('%Y%m%d')}.txt",
                    mime="text/plain"
                )
            with col_r2:
                st.download_button(
                    label="📊 下载趋势数据（Excel）",
                    data=to_excel(report_data),
                    file_name=f"{target_company}_趋势数据.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os
import altair as alt

# 全局设置：解决中文显示/对齐问题
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)

# ====================== 路径配置 ======================
DIGITAL_TRANSFORMATION_FILE = "数字化转型指数分析结果.xlsx"
# ======================================================

# 保留的列名
RETAIN_COLUMNS = [
    "股票代码",
    "企业名称",
    "年份",
    "数字化转型综合指数",
    "人工智能词频数",
    "大数据词频数",
    "云计算词频数",
    "区块链词频数",
    "数字技术运用词频数"
]

# 工具函数：生成Excel下载文件
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='数据')
    writer.close()
    return output.getvalue()

# 核心：全局指数标准化（词频全零置0 + 0-100缩放 + 强制非负）
def normalize_index_global(df):
    index_col = "数字化转型综合指数"
    if index_col not in df.columns:
        return df
    
    # 步骤1：词频全为0的企业，指数直接置0
    word_freq_cols = [col for col in RETAIN_COLUMNS if col.endswith("词频数")]
    df["词频全零"] = (df[word_freq_cols] == 0).all(axis=1)
    df.loc[df["词频全零"], index_col] = 0
    df = df.drop("词频全零", axis=1)

    # 步骤2：提取所有非零指数的原始数据，做全局Min-Max标准化到0-100
    raw_index = df[index_col].copy()
    non_zero_raw = raw_index[raw_index > 0]
    if len(non_zero_raw) == 0:
        df[index_col] = 0.0
        return df
    
    # 计算全局最大/最小值（基于所有年份的非零原始指数）
    global_min = non_zero_raw.min()
    global_max = non_zero_raw.max()

    # 对所有数据做标准化（包含零值）
    df[index_col] = (raw_index - global_min) / (global_max - global_min) * 100
    # 强制将负数/零值设为0（避免标准化后出现负数）
    df[index_col] = df[index_col].clip(lower=0)
    # 保留2位小数
    df[index_col] = df[index_col].round(2)
    
    return df

# 生成企业综合报告
def generate_company_report(company_name, company_data, full_trend_data):
    stock_code = company_data["股票代码"].iloc[0] if ("股票代码" in company_data.columns and not company_data.empty) else "未知"
    available_years = sorted(company_data["年份"].unique()) if ("年份" in company_data.columns and not company_data.empty) else []
    total_years = len(available_years)
    
    index_analysis = {"max_index":0, "max_year":"无", "avg_index":0, "latest_index":0, "trend":"无数据"}
    if "数字化转型综合指数" in company_data.columns and not company_data.empty:
        max_val = company_data["数字化转型综合指数"].max()
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
    
    word_freq_cols = [col for col in RETAIN_COLUMNS if col.endswith("词频数")]
    word_freq_data = {col: 0 for col in word_freq_cols}
    if not company_data.empty:
        for col in word_freq_cols:
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
{chr(10).join([f"- {col}：{word_freq_data[col]}" for col in word_freq_cols])}

## 四、完整指数明细
{full_trend_data.round(2).to_string(index=False)}

## 五、数据说明
1. 指数越高代表转型程度越高
2. 词频全为0的企业，指数直接置为0
3. 指数已通过全局标准化处理
"""
    return report, full_trend_data

# 读取并全局标准化数据
def load_full_data(file_path):
    try:
        if not os.path.exists(file_path):
            st.error(f"❌ 未找到文件：{file_path}")
            return pd.DataFrame()
        
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_names = [name for name in excel_file.sheet_names if name.isdigit()]
        if not sheet_names:
            st.error("❌ Excel中无纯数字名称的工作表（如1999）")
            return pd.DataFrame()
        
        df_list = []
        for sheet in sheet_names:
            sheet_df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
            sheet_df["年份"] = sheet
            # 保留指定列
            sheet_df = sheet_df[[col for col in RETAIN_COLUMNS if col in sheet_df.columns]]
            # 修正股票代码格式
            if "股票代码" in sheet_df.columns:
                sheet_df["股票代码"] = sheet_df["股票代码"].astype(str).str.zfill(6)
            df_list.append(sheet_df)
        
        full_df = pd.concat(df_list, ignore_index=True)
        full_df = full_df.fillna(0)
        
        # 全局执行指数标准化（核心步骤，覆盖所有数据）
        full_df = normalize_index_global(full_df)
        
        return full_df.reset_index(drop=True)
    except Exception as e:
        st.error(f"❌ 读取数据失败：{str(e)}")
        return pd.DataFrame()

# 获取所有年份
def get_all_years(full_data):
    if "年份" not in full_data.columns:
        st.error("❌ 数据中无有效年份")
        return []
    return sorted(full_data["年份"].unique())

def main():
    st.title("企业数字化转型指数查询系统")
    
    # 读取并全局标准化数据（所有模块共用这一份数据）
    full_data = load_full_data(DIGITAL_TRANSFORMATION_FILE)
    if full_data.empty:
        return

    all_years = get_all_years(full_data)
    if not all_years:
        st.error("❌ 数据中无有效年份")
        return

    # 查询区域
    st.subheader("🔍 企业查询（股票代码/名称）")
    col1, col2, col3 = st.columns(3)
    with col1:
        stock_code = st.text_input("输入股票代码（如：000001）", placeholder="股票代码")
    with col2:
        company_name = st.text_input("输入企业名称（如：平安银行）", placeholder="企业名称")
    with col3:
        selected_year = st.selectbox("选择查询年份", all_years, index=0)

    # 筛选企业数据（基于全局标准化后的数据）
    company_all_data = pd.DataFrame()
    filter_cond = full_data["年份"] == selected_year
    if stock_code and "股票代码" in full_data.columns:
        company_all_data = full_data[(full_data["股票代码"] == stock_code.strip().zfill(6)) & filter_cond].copy()
    elif company_name and "企业名称" in full_data.columns:
        company_all_data = full_data[(full_data["企业名称"].str.contains(company_name.strip(), na=False)) & filter_cond].copy()

    # 展示当年数据
    current_year_data = full_data[filter_cond].copy()
    st.success(f"✅ 已查询{selected_year}年数据（总计{len(current_year_data)}家企业）")
    st.subheader("📋 企业当年详细数据")
    current_filtered_data = current_year_data.copy()
    if stock_code and "股票代码" in current_filtered_data.columns:
        current_filtered_data = current_filtered_data[current_filtered_data["股票代码"] == stock_code.strip().zfill(6)]
    if company_name and "企业名称" in current_filtered_data.columns:
        current_filtered_data = current_filtered_data[current_filtered_data["企业名称"].str.contains(company_name.strip(), na=False)]
    
    if not current_filtered_data.empty:
        st.dataframe(current_filtered_data, use_container_width=True)
        st.info(f"筛选结果：找到{len(current_filtered_data)}家匹配企业")
    else:
        st.info(f"ℹ️ {selected_year}年数据中无匹配企业，请调整查询条件")

    # 全行业趋势图（全局标准化数据）
    if "数字化转型综合指数" in full_data.columns:
        st.subheader("📊 全行业转型指数趋势")
        industry_avg_data = []
        for year in all_years:
            year_data = full_data[full_data["年份"] == year]
            avg_idx = year_data["数字化转型综合指数"].mean() if not year_data.empty else 0
            industry_avg_data.append({"年份": year, "平均指数": round(avg_idx, 2)})
        industry_avg_df = pd.DataFrame(industry_avg_data)
        st.line_chart(
            industry_avg_df.set_index("年份")["平均指数"],
            use_container_width=True,
            color="#2E86AB",
            height=400
        )

    # 企业趋势图（全局标准化数据）
    if not company_all_data.empty and "数字化转型综合指数" in company_all_data.columns:
        selected_company = company_all_data["企业名称"].unique()[0] if ("企业名称" in company_all_data.columns and not company_all_data.empty) else "未知企业"
        stock_code_display = stock_code if stock_code else (company_all_data["股票代码"].iloc[0] if ("股票代码" in company_all_data.columns and not company_all_data.empty) else "未知代码")
        
        # 准备企业趋势数据（全局标准化后的数据）
        company_trend = []
        for year in all_years:
            year_data = full_data[(full_data["年份"] == year) & (full_data["股票代码"] == stock_code_display)]
            idx_val = year_data["数字化转型综合指数"].iloc[0] if not year_data.empty else 0
            company_trend.append({"年份": year, "数字化转型综合指数": idx_val})
        company_trend_df = pd.DataFrame(company_trend)

        # 绘制企业趋势图
        st.subheader(f"📈 {selected_company}（{stock_code_display}）转型指数趋势")
        y_max = company_trend_df["数字化转型综合指数"].max()
        arrow_y = min(y_max * 1.2, 100) if y_max > 0 else 20

        base = alt.Chart(company_trend_df).encode(
            x=alt.X("年份:O", axis=alt.Axis(labelAngle=-45)),
            y=alt.Y("数字化转型综合指数:Q", title="数字化转型综合指数", scale=alt.Scale(domain=[0, 100]))
        )
        normal_line = base.mark_line(color="#FF6B6B", strokeWidth=2)
        normal_points = base.mark_point(size=60, color="#FF6B6B")

        selected_trend_data = company_trend_df[company_trend_df["年份"] == selected_year].copy()
        selected_trend_data["箭头Y"] = arrow_y
        highlight_arrow = alt.Chart(selected_trend_data).mark_point(
            size=300, shape="triangle-down", color="#FF0000", stroke="black", strokeWidth=2
        ).encode(x="年份:O", y="箭头Y:Q")
        highlight_text = highlight_arrow.mark_text(
            align="center", baseline="bottom", dy=-10, color="#FF0000", fontWeight="bold", fontSize=14
        ).encode(text=alt.Text("数字化转型综合指数:Q", format=".2f"))
        line_to_point = alt.Chart(selected_trend_data).mark_line(
            color="#FF0000", strokeDash=[3,3]
        ).encode(x="年份:O", y=alt.Y("数字化转型综合指数:Q"), y2="箭头Y:Q")

        chart = (normal_line + normal_points + line_to_point + highlight_arrow + highlight_text).properties(
            height=500, width="container"
        )
        st.altair_chart(chart, use_container_width=True)
        
        # 展示历年完整数据（全局标准化后的数据）
        st.subheader(f"📋 {selected_company} 历年完整数据")
        company_detail_display = full_data[full_data["股票代码"] == stock_code_display].copy()
        st.dataframe(company_detail_display, use_container_width=True)

        # 下载功能
        st.subheader("📥 综合报告下载")
        report_text, report_data = generate_company_report(selected_company, company_all_data, company_trend_df)
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            st.download_button(label="📄 下载报告（TXT）", data=report_text, file_name=f"{selected_company}_报告_{datetime.now().strftime('%Y%m%d')}.txt", mime="text/plain")
        with col_r2:
            st.download_button(label="📊 下载趋势数据（Excel）", data=to_excel(company_trend_df), file_name=f"{selected_company}_趋势数据.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_r3:
            st.download_button(label="📋 下载历年数据（Excel）", data=to_excel(company_detail_display), file_name=f"{selected_company}_历年数据.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif stock_code or company_name:
        st.warning("⚠️ 未找到匹配的企业数据，请检查股票代码或企业名称是否正确")

if __name__ == "__main__":
    main()

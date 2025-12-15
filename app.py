import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import os
import plotly.express as px
import plotly.graph_objects as go

# 全局设置：解决中文显示/对齐问题
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)

# ====================== 路径配置（已按你的需求修改）======================
DIGITAL_TRANSFORMATION_FILE = r"C:\Users\43474\Desktop\大数据\数字化转型指数分析结果.xlsx"  # 你的数据文件路径
# （注：原代码中WORD_FREQ_FILE等未实际使用，若不需要可保留或删除）
WORD_FREQ_FILE = r"C:\Users\43474\Desktop\大数据\数字化转型指数分析结果.xlsx"  # 若需要可同步修改
PY_FILE_PATH = r"C:\Users\43474\Desktop\大数据\app.py"  # 若你的app.py在这个路径，同步修改
ROOT_FOLDER = r"C:\Users\43474\Desktop\大数据"  # 数据所在文件夹
# =====================================================================

# 工具函数：生成Excel下载文件
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='数据')
    writer.close()
    return output.getvalue()

# 生成企业综合报告
def generate_company_report(company_name, company_data, full_trend_data):
    stock_code = company_data["股票代码"].iloc[0] if ("股票代码" in company_data.columns and not company_data.empty) else "未知"
    available_years = sorted(company_data["年份"].unique()) if ("年份" in company_data.columns and not company_data.empty) else []
    total_years = len(available_years)
    
    index_analysis = {"max_index":0, "max_year":"无", "avg_index":0, "latest_index":0, "trend":"无数据"}
    if "数字化转型综合指数" in company_data.columns and not company_data.empty:  # 字段名改为你的“数字化转型综合指数”
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
{full_trend_data.round(2).to_string(index=False)}

## 五、数据说明
1. 数字化转型综合指数越高代表转型程度越高
2. 词频数据反映对应技术的应用强度
"""
    return report, full_trend_data

# 读取完整数据（匹配你的工作表名称）
def load_full_data(file_path):
    try:
        df = pd.read_excel(
            file_path,
            sheet_name="Sheet1",  # 若你的工作表不是Sheet1，需改为实际名称（比如图中显示的“2023”）
            engine="openpyxl"
        )
        # 清洗数据（兼容文本/数字格式的年份）
        if "年份" in df.columns:
            # 尝试将年份转为整数，避免字符串格式问题
            df["年份"] = pd.to_numeric(df["年份"], errors='coerce').fillna(df["年份"]).astype(str).str.strip()
        if "企业名称" in df.columns:
            df["企业名称"] = df["企业名称"].str.strip()
        if "股票代码" in df.columns:
            df["股票代码"] = df["股票代码"].astype(str).str.strip()
        return df.dropna(how="all").reset_index(drop=True)
    except Exception as e:
        st.error(f"❌ 读取数据失败：{str(e)}")
        return pd.DataFrame()

# 获取数据中所有年份（不限制范围）
def get_all_years(full_data):
    if "年份" not in full_data.columns:
        st.error("❌ 数据中未找到'年份'列")
        return []
    return sorted(full_data["年份"].unique())

def main():
    st.title("企业数字化转型指数查询系统")
    
    # 验证文件是否存在
    if not os.path.exists(DIGITAL_TRANSFORMATION_FILE):
        st.error(f"❌ 文件不存在：{DIGITAL_TRANSFORMATION_FILE}")
        return
    
    # 读取完整数据
    full_data = load_full_data(DIGITAL_TRANSFORMATION_FILE)
    if full_data.empty:
        st.error("❌ 数据为空，请检查Excel文件内容")
        return

    # 获取所有年份
    all_years = get_all_years(full_data)
    if not all_years:
        st.error("❌ 数据中无有效年份")
        return

    # 查询区域：调整顺序，股票代码在前
    st.subheader("🔍 企业查询（股票代码/名称）")
    col1, col2, col3 = st.columns(3)
    with col1:
        stock_code = st.text_input("输入股票代码（如：000001）", placeholder="股票代码")
    with col2:
        company_name = st.text_input("输入企业名称（如：平安银行）", placeholder="企业名称")
    with col3:
        selected_year = st.selectbox("选择查询年份", all_years, index=0)

    # 初始化筛选数据
    filtered_data = full_data.copy()
    company_all_data = pd.DataFrame()
    
    # 优先通过股票代码筛选
    if stock_code:
        company_all_data = full_data[full_data["股票代码"] == stock_code.strip()].copy()
    # 补充企业名称筛选
    elif company_name:
        company_all_data = full_data[full_data["企业名称"].str.contains(company_name.strip(), na=False)].copy()

    # 筛选当前年份数据
    current_year_data = full_data[full_data["年份"] == selected_year].copy()
    
    # 展示当年数据
    st.success(f"✅ 已查询{selected_year}年数据（总计{len(current_year_data)}家企业）")
    st.subheader("📋 企业当年详细数据")
    
    # 筛选当年的匹配数据
    current_filtered_data = current_year_data.copy()
    if stock_code:
        current_filtered_data = current_filtered_data[current_filtered_data["股票代码"] == stock_code.strip()]
    if company_name:
        current_filtered_data = current_filtered_data[current_filtered_data["企业名称"].str.contains(company_name.strip(), na=False)]

    if not current_filtered_data.empty:
        st.dataframe(current_filtered_data, use_container_width=True)
        st.info(f"筛选结果：找到{len(current_filtered_data)}家匹配企业")
    else:
        st.info(f"ℹ️ {selected_year}年数据中无匹配企业，请调整查询条件")

    # 全行业平均指数趋势图
    st.subheader("📊 全行业转型指数趋势")
    industry_avg_data = []
    for year in all_years:
        year_data = full_data[full_data["年份"] == year]
        if not year_data.empty and "数字化转型综合指数" in year_data.columns:  # 字段名改为你的“数字化转型综合指数”
            avg_idx = year_data["数字化转型综合指数"].mean()
        else:
            avg_idx = 0
        industry_avg_data.append({
            "年份": year, 
            "平均指数": round(avg_idx, 4)
        })
    industry_avg_df = pd.DataFrame(industry_avg_data)
    st.line_chart(industry_avg_df.set_index("年份")["平均指数"], use_container_width=True, color="#2E86AB", height=400)

    # 企业全量趋势图：输入股票代码后自动展示
    if not company_all_data.empty:
        # 获取企业名称
        selected_company = company_all_data["企业名称"].unique()[0] if len(company_all_data["企业名称"].unique()) > 0 else "未知企业"
        
        # 补全所有年份的趋势数据
        full_years_df = pd.DataFrame({"年份": all_years})
        company_trend = pd.merge(
            full_years_df,
            company_all_data[["年份", "数字化转型综合指数"]],  # 字段名改为你的“数字化转型综合指数”
            on="年份",
            how="left"
        ).fillna(0)

        # 展示趋势图（修改部分：使用plotly创建带标注的折线图）
        st.subheader(f"📈 {selected_company}（{stock_code if stock_code else '未知代码'}）转型指数趋势")
        
        # 创建plotly图表
        fig = go.Figure()
        
        # 添加主趋势线
        fig.add_trace(go.Scatter(
            x=company_trend["年份"],
            y=company_trend["数字化转型综合指数"],
            mode='lines+markers',
            name='转型指数',
            line=dict(color="#FF6B6B", width=2),
            marker=dict(size=8, color="#FF6B6B")
        ))
        
        # 找到选中年份的数据点并添加标注
        selected_year_data = company_trend[company_trend["年份"] == selected_year]
        if not selected_year_data.empty:
            selected_value = selected_year_data["数字化转型综合指数"].iloc[0]
            
            # 添加选中年份的特殊标记点
            fig.add_trace(go.Scatter(
                x=[selected_year],
                y=[selected_value],
                mode='markers+text',
                name=f'{selected_year}年',
                marker=dict(
                    size=15,
                    color="#FF4757",
                    symbol="star",
                    line=dict(width=2, color="black")
                ),
                text=[f' {selected_year}年: {selected_value:.2f}'],
                textposition="top right",
                textfont=dict(size=12, color="#FF4757", weight="bold")
            ))
            
            # 添加垂直参考线
            fig.add_vline(
                x=selected_year,
                line=dict(color="#70A1FF", dash="dash", width=2),
                annotation_text=f"{selected_year}年",
                annotation_position="top",
                annotation_font=dict(size=10, color="#70A1FF")
            )
        
        # 设置图表样式
        fig.update_layout(
            xaxis_title="年份",
            yaxis_title="数字化转型综合指数",
            height=500,
            showlegend=True,
            hovermode="x unified",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)"
        )
        
        # 在streamlit中显示plotly图表
        st.plotly_chart(fig, use_container_width=True)
        
        # 展示历年完整数据
        st.subheader(f"📋 {selected_company} 历年完整数据")
        display_columns = ["年份", "股票代码", "数字化转型综合指数", "人工智能词频数", "大数据词频数", "云计算词频数", "区块链词频数", "数字技术运用词频数"]  # 字段名同步修改
        # 筛选存在的列
        display_columns = [col for col in display_columns if col in company_all_data.columns]
        company_detail = company_all_data[display_columns].sort_values("年份").reset_index(drop=True)
        st.dataframe(company_detail, use_container_width=True)

        # 下载功能
        st.subheader("📥 综合报告下载")
        report_text, report_data = generate_company_report(selected_company, company_all_data, company_trend)
        col_r1, col_r2, col_r3 = st.columns(3)
        with col_r1:
            st.download_button(label="📄 下载报告（TXT）", data=report_text, file_name=f"{selected_company}_报告_{datetime.now().strftime('%Y%m%d')}.txt", mime="text/plain")
        with col_r2:
            st.download_button(label="📊 下载趋势数据（Excel）", data=to_excel(report_data), file_name=f"{selected_company}_趋势数据.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_r3:
            st.download_button(label="📋 下载历年数据（Excel）", data=to_excel(company_detail), file_name=f"{selected_company}_历年数据.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif stock_code or company_name:
        st.warning("⚠️ 未找到匹配的企业数据，请检查股票代码或企业名称是否正确")

if __name__ == "__main__":
    main()

# app.py
import streamlit as st
import pandas as pd
import os

# 基础设置
st.set_page_config(layout="wide")
st.title("数字化转型指数查询系统")

# 显示当前目录和文件（用于调试）
st.write("### 系统信息")
st.write("当前目录：", os.getcwd())

# 列出所有文件
files = os.listdir(".")
st.write("文件列表：")
for file in files:
    st.write(f"- {file}")

# 查找Excel文件
excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
st.write(f"### 找到 {len(excel_files)} 个Excel文件")

if excel_files:
    # 尝试读取第一个Excel文件
    file_name = excel_files[0]
    try:
        st.write(f"正在读取文件：{file_name}")
        
        # 读取Excel
        df = pd.read_excel(file_name)
        
        st.success(f"✅ 成功读取文件！")
        st.write(f"数据形状：{df.shape} 行 × {df.shape[1]} 列")
        
        # 显示列名
        st.write("### 数据列名")
        st.write(df.columns.tolist())
        
        # 显示前5行
        st.write("### 数据预览")
        st.dataframe(df.head())
        
        # 如果文件有企业名称列，提供查询功能
        if '企业名称' in df.columns or '股票代码' in df.columns:
            st.write("### 查询功能")
            
            # 确定显示哪些列
            display_cols = []
            if '企业名称' in df.columns:
                display_cols.append('企业名称')
            if '股票代码' in df.columns:
                display_cols.append('股票代码')
            if '年份' in df.columns:
                display_cols.append('年份')
            if '数字化转型综合指数' in df.columns:
                display_cols.append('数字化转型综合指数')
            
            # 显示筛选后的数据
            if display_cols:
                st.dataframe(df[display_cols].head(20))
                
            # 简单的搜索功能
            search_col = st.selectbox(
                "选择搜索列",
                [col for col in ['企业名称', '股票代码'] if col in df.columns]
            )
            
            if search_col:
                search_term = st.text_input(f"输入{search_col}进行搜索")
                if search_term:
                    results = df[df[search_col].astype(str).str.contains(search_term, case=False, na=False)]
                    st.write(f"找到 {len(results)} 条记录")
                    st.dataframe(results[display_cols] if display_cols else results)
        
    except Exception as e:
        st.error(f"❌ 读取文件时出错：{str(e)}")
        import traceback
        st.write("详细错误信息：")
        st.code(traceback.format_exc())
else:
    st.error("❌ 没有找到Excel文件！请确保文件已上传。")

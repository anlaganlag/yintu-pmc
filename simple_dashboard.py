#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC简化仪表板 - 无密码版本
避免SessionInfo初始化问题的最简版本
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
from datetime import datetime

# 页面配置
st.set_page_config(
    page_title="银图PMC智能分析平台",
    page_icon="🌟", 
    layout="wide"
)

def safe_init_state(key, default_value):
    """安全初始化session state"""
    try:
        if key not in st.session_state:
            st.session_state[key] = default_value
    except Exception:
        # 如果session_state出问题，使用临时变量
        pass

def load_latest_report():
    """加载最新的分析报告"""
    try:
        # 查找最新的报告文件
        report_files = [
            "银图PMC综合物料分析报告_改进版_20250828_101505.xlsx",
            "过滤WO后无供应商物料报告.xlsx",
            "银图PMC综合物料分析报告_20250828_100022.xlsx"
        ]
        
        for file in report_files:
            try:
                df = pd.read_excel(file, sheet_name='综合物料分析明细')
                return df, file
            except:
                continue
                
        return None, None
    except Exception as e:
        return None, str(e)

def main():
    """主应用"""
    
    # 标题
    st.title("🌟 银图PMC智能分析平台")
    st.markdown("---")
    
    # 加载数据
    with st.spinner("正在加载分析数据..."):
        data, source_file = load_latest_report()
    
    if data is None:
        st.error("❌ 无法加载分析报告数据")
        st.info("请确保分析报告文件存在")
        return
    
    st.success(f"✅ 已加载数据: {source_file}")
    st.info(f"数据记录数: {len(data):,} 条")
    
    # 基本统计
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("总记录数", f"{len(data):,}")
    
    with col2:
        if '供应商' in data.columns:
            no_supplier = data['供应商'].isna().sum()
            st.metric("无供应商记录", f"{no_supplier:,}")
    
    with col3:
        if '物料编码' in data.columns:
            materials = data['物料编码'].nunique()
            st.metric("物料种类", f"{materials:,}")
    
    with col4:
        if '生产单号' in data.columns:
            orders = data['生产单号'].nunique()
            st.metric("生产单数", f"{orders:,}")
    
    # 数据表格
    st.markdown("### 📊 数据明细")
    
    # 显示前1000条记录
    display_data = data.head(1000)
    st.dataframe(display_data, use_container_width=True, height=400)
    
    # 无供应商物料统计
    if '供应商' in data.columns and '物料编码' in data.columns:
        st.markdown("### ❌ 无供应商物料TOP 20")
        
        no_supplier_data = data[data['供应商'].isna() | (data['供应商'] == '')]
        if len(no_supplier_data) > 0:
            material_counts = no_supplier_data['物料编码'].value_counts().head(20)
            
            # 显示图表
            fig = px.bar(
                x=material_counts.values,
                y=material_counts.index,
                orientation='h',
                title="无供应商物料出现频次TOP 20"
            )
            fig.update_layout(height=600)
            st.plotly_chart(fig, use_container_width=True)
            
            # 显示表格
            st.dataframe(material_counts.to_frame('出现次数'), use_container_width=True)
    
    # 下载功能
    st.markdown("### 💾 数据下载")
    
    if st.button("下载CSV数据"):
        try:
            csv = data.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="点击下载CSV文件",
                data=csv,
                file_name=f"PMC分析数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime='text/csv'
            )
        except Exception as e:
            st.error(f"下载失败: {e}")

if __name__ == "__main__":
    main()
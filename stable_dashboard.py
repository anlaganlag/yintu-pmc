#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC稳定版仪表板
解决索引超出范围错误的稳定版本
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import os

# 页面配置 - 必须在最前面
st.set_page_config(
    page_title="银图PMC分析平台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 清理session state中的问题键
def clean_session_state():
    """清理可能导致索引问题的session state"""
    problematic_keys = []
    for key in list(st.session_state.keys()):
        if key.startswith('FormSubmitter:'):
            problematic_keys.append(key)
    
    for key in problematic_keys:
        try:
            del st.session_state[key]
        except:
            pass

# 在应用开始时清理
clean_session_state()

def load_data():
    """加载数据 - 使用缓存避免重复加载"""
    
    # 查找报告文件
    report_files = [
        "银图PMC综合物料分析报告_改进版_20250828_101505.xlsx",
        "过滤WO后无供应商物料报告.xlsx"
    ]
    
    for file_path in report_files:
        if os.path.exists(file_path):
            try:
                data = pd.read_excel(file_path, sheet_name='综合物料分析明细')
                return data, file_path
            except Exception as e:
                continue
    
    return None, "未找到数据文件"

def main():
    """主应用"""
    
    # 固定标题 - 避免动态组件
    st.markdown("# 📊 银图PMC分析平台")
    st.markdown("---")
    
    # 加载数据
    data, source_info = load_data()
    
    if data is None:
        st.error(f"❌ 数据加载失败: {source_info}")
        st.stop()
    
    # 基础信息
    st.success(f"✅ 数据已加载: {source_info}")
    
    # 使用固定的列布局 - 避免动态变化
    metrics_col1, metrics_col2, metrics_col3, metrics_col4 = st.columns(4)
    
    # 计算指标
    total_records = len(data)
    
    no_supplier_count = 0
    if '供应商' in data.columns:
        no_supplier_count = data['供应商'].isna().sum() + (data['供应商'] == '').sum()
    
    material_count = 0
    if '物料编码' in data.columns:
        material_count = data['物料编码'].nunique()
    
    order_count = 0
    if '生产单号' in data.columns:
        order_count = data['生产单号'].nunique()
    
    # 显示指标
    with metrics_col1:
        st.metric("总记录数", f"{total_records:,}")
    
    with metrics_col2:
        st.metric("无供应商记录", f"{no_supplier_count:,}")
    
    with metrics_col3:
        st.metric("物料种类", f"{material_count:,}")
    
    with metrics_col4:
        st.metric("生产单数", f"{order_count:,}")
    
    # 数据预览 - 固定显示1000条
    st.markdown("### 📋 数据预览")
    
    # 使用固定的dataframe - 避免动态变化
    preview_data = data.head(1000)
    st.dataframe(preview_data, height=400, use_container_width=True)
    
    # 无供应商分析
    if '供应商' in data.columns and '物料编码' in data.columns:
        st.markdown("### ❌ 无供应商物料分析")
        
        # 筛选无供应商数据
        no_supplier_mask = data['供应商'].isna() | (data['供应商'] == '') | (data['供应商'] == '无供应商')
        no_supplier_data = data[no_supplier_mask]
        
        if len(no_supplier_data) > 0:
            # 统计TOP 10 - 固定数量
            material_stats = no_supplier_data['物料编码'].value_counts().head(10)
            
            if len(material_stats) > 0:
                # 创建图表
                fig = px.bar(
                    x=material_stats.values,
                    y=material_stats.index,
                    orientation='h',
                    title="无供应商物料TOP 10（按出现频次）",
                    labels={'x': '出现次数', 'y': '物料编码'}
                )
                fig.update_layout(height=500, showlegend=False)
                st.plotly_chart(fig, use_container_width=True, key="no_supplier_chart")
                
                # 显示统计表
                stats_df = material_stats.to_frame('出现次数').reset_index()
                stats_df.columns = ['物料编码', '出现次数']
                st.dataframe(stats_df, use_container_width=True)
    
    # 下载区域
    st.markdown("### 💾 数据下载")
    
    # 创建下载按钮 - 使用固定key
    if st.button("生成CSV下载", key="download_csv_btn"):
        try:
            csv_data = data.to_csv(index=False, encoding='utf-8-sig')
            filename = f"PMC分析数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            
            st.download_button(
                label="📥 点击下载CSV文件",
                data=csv_data,
                file_name=filename,
                mime='text/csv',
                key="csv_download"
            )
            st.success("✅ CSV文件已准备就绪，点击上方按钮下载")
            
        except Exception as e:
            st.error(f"❌ 下载准备失败: {e}")
    
    # 页脚信息
    st.markdown("---")
    st.markdown("🕒 最后更新时间: " + datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"应用运行错误: {e}")
        st.info("正在重新启动...请刷新页面")
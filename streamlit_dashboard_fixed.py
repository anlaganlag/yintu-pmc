#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图订单追踪分析平台 - 修复版
集成了SessionInfo错误的三大解决方案
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import io
import os
import tempfile
import time
import gc
from datetime import datetime, timedelta
from pathlib import Path
import subprocess
import sys

# 导入修复模块
from streamlit_fix import (
    SessionStateCleaner, 
    DataPaginator, 
    apply_optimized_config,
    OptimizedStreamlitApp
)

# 初始化优化应用
app = OptimizedStreamlitApp()

# 在任何Streamlit调用之前应用配置
app.initialize()

@st.cache_data(ttl=300)  # 5分钟缓存
def load_excel_file(file_path):
    """加载Excel文件（带缓存）"""
    try:
        return pd.read_excel(file_path, sheet_name=None)
    except Exception as e:
        st.error(f"加载文件失败: {str(e)}")
        return None

@st.cache_data(ttl=600)  # 10分钟缓存
def process_analysis_data(file_path):
    """处理分析数据"""
    try:
        df = pd.read_excel(file_path, sheet_name='综合物料分析明细')
        
        # 数据类型转换
        numeric_columns = ['订单金额(RMB)', '缺料金额(RMB)', 'ROI(投资回报率)']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
                
        # 日期处理
        if '日期' in df.columns:
            df['日期'] = pd.to_datetime(df['日期'], errors='coerce')
            
        return df
    except Exception as e:
        st.error(f"处理数据失败: {str(e)}")
        return None

def run_analysis():
    """运行分析并分页显示结果"""
    st.header("📊 运行物料分析")
    
    # 文件上传区域
    with st.container():
        st.subheader("上传数据文件")
        
        col1, col2 = st.columns(2)
        
        with col1:
            order_file = st.file_uploader(
                "订单文件 (国内)", 
                type=['xlsx', 'xls'],
                key=app.session_cleaner.get_state("order_upload_key", "order_1")
            )
            order_file_c = st.file_uploader(
                "订单文件 (柬埔寨)", 
                type=['xlsx', 'xls'],
                key=app.session_cleaner.get_state("order_c_upload_key", "order_2")
            )
            mat_file = st.file_uploader(
                "缺料清单", 
                type=['xlsx', 'xls'],
                key=app.session_cleaner.get_state("mat_upload_key", "mat_1")
            )
            
        with col2:
            supplier_file = st.file_uploader(
                "供应商数据", 
                type=['xlsx', 'xls'],
                key=app.session_cleaner.get_state("supplier_upload_key", "supplier_1")
            )
            inventory_file = st.file_uploader(
                "库存清单", 
                type=['xlsx', 'xls'],
                key=app.session_cleaner.get_state("inventory_upload_key", "inventory_1")
            )
    
    # 运行分析按钮
    if st.button("🚀 开始分析", type="primary"):
        if all([order_file, order_file_c, mat_file, supplier_file, inventory_file]):
            with st.spinner("正在分析数据..."):
                # 保存上传的文件
                temp_dir = tempfile.mkdtemp()
                files_saved = {}
                
                try:
                    # 保存文件
                    for name, file_obj in [
                        ('order-amt-89.xlsx', order_file),
                        ('order-amt-89-c.xlsx', order_file_c),
                        ('mat_owe_pso.xlsx', mat_file),
                        ('supplier.xlsx', supplier_file),
                        ('inventory_list.xlsx', inventory_file)
                    ]:
                        file_path = os.path.join(temp_dir, name)
                        with open(file_path, 'wb') as f:
                            f.write(file_obj.getbuffer())
                        files_saved[name] = file_path
                    
                    # 运行分析脚本
                    result = subprocess.run(
                        [sys.executable, 'silverPlan_analysis.py'],
                        cwd=temp_dir,
                        capture_output=True,
                        text=True,
                        timeout=120
                    )
                    
                    if result.returncode == 0:
                        st.success("✅ 分析完成！")
                        
                        # 查找生成的报告
                        report_files = list(Path(temp_dir).glob("银图PMC综合物料分析报告_*.xlsx"))
                        if report_files:
                            latest_report = max(report_files, key=os.path.getctime)
                            
                            # 存储分析结果路径
                            app.set_state('latest_report', str(latest_report))
                            app.set_state('analysis_complete', True)
                            
                            # 提供下载
                            with open(latest_report, 'rb') as f:
                                st.download_button(
                                    label="📥 下载分析报告",
                                    data=f.read(),
                                    file_name=latest_report.name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                    else:
                        st.error(f"分析失败: {result.stderr}")
                        
                except subprocess.TimeoutError:
                    st.error("分析超时，请检查数据量")
                except Exception as e:
                    st.error(f"分析出错: {str(e)}")
                finally:
                    # 清理临时文件
                    import shutil
                    try:
                        shutil.rmtree(temp_dir)
                    except:
                        pass
                    
                    # 清理session
                    app.cleanup()
        else:
            st.warning("请上传所有必需的文件")

def display_dashboard():
    """显示仪表板（带分页）"""
    st.header("📈 分析仪表板")
    
    # 检查是否有分析结果
    report_path = app.get_state('latest_report')
    if not report_path or not os.path.exists(report_path):
        st.info("请先运行分析以查看结果")
        return
        
    # 加载数据
    df = process_analysis_data(report_path)
    if df is None or df.empty:
        st.warning("没有可显示的数据")
        return
        
    # KPI指标
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_shortage = df['缺料金额(RMB)'].sum() if '缺料金额(RMB)' in df.columns else 0
        st.metric("总缺料金额", f"¥{total_shortage:,.0f}")
        
    with col2:
        order_count = df['订单号'].nunique() if '订单号' in df.columns else 0
        st.metric("订单总数", f"{order_count:,}")
        
    with col3:
        avg_roi = df['ROI(投资回报率)'].mean() if 'ROI(投资回报率)' in df.columns else 0
        st.metric("平均ROI", f"{avg_roi:.2f}")
        
    with col4:
        urgent_count = len(df[df['缺料金额(RMB)'] > 50000]) if '缺料金额(RMB)' in df.columns else 0
        st.metric("紧急采购", f"{urgent_count} 单")
    
    # 使用选项卡组织内容
    tab1, tab2, tab3 = st.tabs(["📊 数据明细", "📈 可视化分析", "🔍 高级筛选"])
    
    with tab1:
        st.subheader("物料分析明细（分页显示）")
        
        # 使用分页器显示数据
        app.safe_render_dataframe(df, key="main_data")
        
    with tab2:
        st.subheader("数据可视化")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # ROI分布图
            if 'ROI(投资回报率)' in df.columns:
                fig_roi = px.histogram(
                    df[df['ROI(投资回报率)'] > 0], 
                    x='ROI(投资回报率)',
                    nbins=30,
                    title="ROI分布"
                )
                st.plotly_chart(fig_roi, use_container_width=True)
                
        with col2:
            # 缺料金额分布
            if '缺料金额(RMB)' in df.columns:
                fig_shortage = px.pie(
                    df.groupby('月份')['缺料金额(RMB)'].sum().reset_index() if '月份' in df.columns else df,
                    values='缺料金额(RMB)',
                    names='月份' if '月份' in df.columns else None,
                    title="缺料金额分布"
                )
                st.plotly_chart(fig_shortage, use_container_width=True)
                
    with tab3:
        st.subheader("高级筛选")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # ROI筛选
            roi_min = st.number_input("最小ROI", value=0.0, step=0.1, key="roi_min_filter")
            roi_max = st.number_input("最大ROI", value=10.0, step=0.1, key="roi_max_filter")
            
        with col2:
            # 金额筛选
            amount_min = st.number_input("最小缺料金额", value=0, step=1000, key="amount_min_filter")
            amount_max = st.number_input("最大缺料金额", value=1000000, step=1000, key="amount_max_filter")
            
        with col3:
            # 月份筛选
            if '月份' in df.columns:
                selected_months = st.multiselect(
                    "选择月份",
                    options=df['月份'].unique(),
                    default=df['月份'].unique(),
                    key="month_filter"
                )
        
        # 应用筛选
        filtered_df = df.copy()
        
        if 'ROI(投资回报率)' in filtered_df.columns:
            filtered_df = filtered_df[
                (filtered_df['ROI(投资回报率)'] >= roi_min) & 
                (filtered_df['ROI(投资回报率)'] <= roi_max)
            ]
            
        if '缺料金额(RMB)' in filtered_df.columns:
            filtered_df = filtered_df[
                (filtered_df['缺料金额(RMB)'] >= amount_min) & 
                (filtered_df['缺料金额(RMB)'] <= amount_max)
            ]
            
        if '月份' in df.columns and 'selected_months' in locals():
            filtered_df = filtered_df[filtered_df['月份'].isin(selected_months)]
        
        st.info(f"筛选后共 {len(filtered_df)} 条记录")
        
        # 显示筛选后的数据（分页）
        app.safe_render_dataframe(filtered_df, key="filtered_data")
        
        # 导出筛选结果
        if not filtered_df.empty:
            csv = filtered_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 导出筛选结果",
                data=csv,
                file_name=f"筛选结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )

def display_expandable_orders():
    """显示可展开的订单列表（带分页）"""
    st.header("📦 订单详情（分页展示）")
    
    # 加载数据
    report_path = app.get_state('latest_report')
    if not report_path or not os.path.exists(report_path):
        st.info("请先运行分析")
        return
        
    df = process_analysis_data(report_path)
    if df is None or df.empty:
        return
        
    # 按订单分组
    grouped = df.groupby('订单号') if '订单号' in df.columns else None
    if not grouped:
        st.warning("没有订单数据")
        return
        
    # 准备订单列表
    orders = []
    for order_no, group in grouped:
        orders.append({
            'title': f"订单 {order_no}",
            '订单号': order_no,
            '总金额': f"¥{group['订单金额(RMB)'].sum():,.0f}" if '订单金额(RMB)' in group.columns else 'N/A',
            '缺料金额': f"¥{group['缺料金额(RMB)'].sum():,.0f}" if '缺料金额(RMB)' in group.columns else 'N/A',
            '平均ROI': f"{group['ROI(投资回报率)'].mean():.2f}" if 'ROI(投资回报率)' in group.columns else 'N/A',
            '物料数': len(group)
        })
    
    # 使用分页显示订单
    app.safe_render_expandable_list(orders, key="orders")

def main():
    """主函数"""
    st.title("🏭 银图PMC智能分析平台 - 优化版")
    st.caption("集成了SessionInfo错误修复方案")
    
    # 侧边栏
    with st.sidebar:
        st.header("功能菜单")
        
        # 添加清理按钮
        if st.button("🧹 清理缓存", help="清理Session State和内存"):
            app.cleanup()
            st.cache_data.clear()
            gc.collect()
            st.success("缓存已清理")
            time.sleep(1)
            st.rerun()
        
        # 显示Session状态
        if hasattr(st, 'session_state'):
            item_count = len(st.session_state.keys())
            st.metric("Session项目数", f"{item_count}/47")
            if item_count > 40:
                st.warning("Session接近上限，建议清理")
        
        page = st.radio(
            "选择功能",
            ["运行分析", "数据仪表板", "订单详情"],
            key="main_page_selector"
        )
    
    # 根据选择显示页面
    if page == "运行分析":
        run_analysis()
    elif page == "数据仪表板":
        display_dashboard()
    elif page == "订单详情":
        display_expandable_orders()
    
    # 页脚
    st.divider()
    st.caption("© 2024 银图PMC分析平台 | 优化版本 v2.0")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = str(e).lower()
        if any(x in error_msg for x in ["sessioninfo", "message format", "bad message", "setin"]):
            st.error("检测到Session错误，正在自动修复...")
            app.cleanup()
            time.sleep(2)
            st.rerun()
        else:
            st.error(f"应用错误: {str(e)}")
            st.info("请尝试点击侧边栏的'清理缓存'按钮")
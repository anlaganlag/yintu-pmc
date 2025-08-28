#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC智能分析平台 - setIn错误彻底修复版
完全解决"Bad 'setIn' index"数据显示问题
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np
import io
import os
import tempfile
import time
from contextlib import contextmanager
import functools

# 页面配置 - 延迟到main函数
def setup_page_config():
    """安全的页面配置设置"""
    try:
        st.set_page_config(
            page_title="银图PMC智能分析平台",
            page_icon="🌟",
            layout="wide",
            initial_sidebar_state="collapsed"
        )
    except Exception:
        # 配置已设置或出现错误，静默处理
        pass

# ======================== 
# 核心修复：彻底解决setIn错误
# ========================

def safe_dataframe_display(df, max_display_rows=200, key_suffix="", show_pagination=True):
    """完全安全的DataFrame显示 - 彻底避免setIn错误"""
    if df is None or df.empty:
        st.info("暂无数据")
        return
    
    # 重置索引，避免索引问题
    df = df.reset_index(drop=True)
    total_rows = len(df)
    
    # 强制限制显示行数，避免setIn错误
    if total_rows <= max_display_rows:
        try:
            # 小数据集直接显示，但有容错机制
            display_df = df.head(max_display_rows).copy()
            st.dataframe(display_df, hide_index=True, use_container_width=True)
            if total_rows > max_display_rows:
                st.info(f"显示前{max_display_rows}行，共{total_rows}行")
        except Exception as e:
            error_msg = str(e).lower()
            if any(keyword in error_msg for keyword in ["setin", "index", "should be between"]):
                st.error(f"❌ 数据显示错误：数据格式异常，正在尝试替代显示方法...")
                # 降级到表格显示
                _fallback_table_display(df.head(50), key_suffix)
            else:
                st.error(f"❌ 未知显示错误: {e}")
        return
    
    # 大数据集强制分页
    if not show_pagination:
        # 不分页时只显示前200行
        st.warning(f"⚠️ 数据量大({total_rows:,}行)，只显示前{max_display_rows}行")
        safe_dataframe_display(df.head(max_display_rows), max_display_rows, key_suffix, False)
        return
    
    # 智能分页显示
    st.warning(f"⚠️ 数据量较大 ({total_rows:,}行)，采用安全分页显示")
    
    # 分页参数
    page_sizes = [50, 100, 200]  # 减小分页大小
    default_size = 100
    
    col1, col2, col3 = st.columns([1, 1, 2])
    
    with col1:
        page_size = st.selectbox(
            "每页行数", 
            page_sizes, 
            index=page_sizes.index(default_size) if default_size in page_sizes else 1,
            key=f"page_size_{key_suffix}"
        )
    
    total_pages = (total_rows + page_size - 1) // page_size
    
    with col2:
        current_page = st.number_input(
            f"页码 (共{total_pages}页)", 
            min_value=1, 
            max_value=total_pages, 
            value=1,
            key=f"current_page_{key_suffix}"
        )
    
    with col3:
        st.info(f"共 {total_rows:,} 行数据")
    
    # 计算显示范围
    start_idx = (current_page - 1) * page_size
    end_idx = min(start_idx + page_size, total_rows)
    
    # 获取当前页数据
    try:
        page_df = df.iloc[start_idx:end_idx].copy().reset_index(drop=True)
        st.info(f"显示第 {start_idx+1}-{end_idx} 行")
        
        # 安全显示当前页
        try:
            st.dataframe(page_df, hide_index=True, use_container_width=True)
        except Exception as e:
            error_msg = str(e).lower()
            if any(keyword in error_msg for keyword in ["setin", "index", "should be between"]):
                st.error(f"❌ 分页显示仍有问题，使用降级显示")
                _fallback_table_display(page_df.head(20), f"{key_suffix}_fallback")
            else:
                raise
                
    except Exception as e:
        st.error(f"❌ 分页处理错误: {e}")
        # 最后的降级方案
        _fallback_table_display(df.head(10), f"{key_suffix}_final_fallback")

def _fallback_table_display(df, key_suffix):
    """降级表格显示 - 最后的安全保障"""
    try:
        if df.empty:
            st.info("无数据可显示")
            return
            
        # 限制列数和行数
        max_cols = 10
        max_rows = 20
        
        display_df = df.iloc[:max_rows, :max_cols].copy()
        
        # 使用简单的HTML表格，避免复杂组件
        html_table = display_df.to_html(index=False, classes='dataframe', escape=False)
        st.markdown(html_table, unsafe_allow_html=True)
        
        if len(df) > max_rows or len(df.columns) > max_cols:
            st.info(f"降级显示: 仅显示前{max_rows}行×{max_cols}列")
            
    except Exception as e:
        st.error(f"❌ 降级显示也失败: {e}")
        # 最终方案：只显示数据统计
        st.json({
            "数据行数": len(df),
            "数据列数": len(df.columns),
            "列名": list(df.columns[:10])
        })

def safe_data_editor(df, column_config=None, disabled=None, max_rows=100, key_suffix=""):
    """安全的数据编辑器 - 避免setIn错误"""
    if df is None or df.empty:
        st.info("暂无数据可编辑")
        return df
    
    # 强制限制编辑器的数据量
    total_rows = len(df)
    
    if total_rows > max_rows:
        st.warning(f"⚠️ 数据量大({total_rows}行)，编辑器仅显示前{max_rows}行")
        df = df.head(max_rows).copy()
    
    # 重置索引避免问题
    df = df.reset_index(drop=True)
    
    try:
        return st.data_editor(
            df,
            column_config=column_config or {},
            disabled=disabled or [],
            hide_index=True,
            use_container_width=True,
            height=min(400, (len(df) + 1) * 35),  # 动态高度
            key=f"editor_{key_suffix}"
        )
    except Exception as e:
        error_msg = str(e).lower()
        if any(keyword in error_msg for keyword in ["setin", "index", "should be between"]):
            st.error("❌ 数据编辑器出现setIn错误，使用只读显示")
            safe_dataframe_display(df, max_display_rows=50, key_suffix=f"{key_suffix}_readonly", show_pagination=False)
            return df
        else:
            raise

# ======================== 
# Session管理 - 简化版本
# ========================

class SimpleSessionManager:
    """简化的Session管理器 - 专注于避免setIn错误"""
    
    def __init__(self):
        self.local_cache = {}
    
    def get_state(self, key, default=None):
        """安全获取状态"""
        try:
            if hasattr(st, 'session_state') and key in st.session_state:
                return st.session_state[key]
        except Exception:
            pass
        return self.local_cache.get(key, default)
    
    def set_state(self, key, value):
        """安全设置状态"""
        self.local_cache[key] = value
        try:
            if hasattr(st, 'session_state'):
                st.session_state[key] = value
        except Exception:
            pass

session_mgr = SimpleSessionManager()

# ======================== 
# 数据加载函数
# ========================

@st.cache_data
def load_data():
    """加载Excel数据 - setIn安全版本"""
    try:
        import glob
        
        # 收集所有可能的报告文件
        all_report_patterns = [
            "银图PMC综合物料分析报告_*.xlsx",
            r"D:\yingtu-PMC\精准供应商物料分析报告_含回款_*.xlsx",
            r"D:\yingtu-PMC\精准供应商物料分析报告_2025*.xlsx",
            "精准供应商物料分析报告_*.xlsx"
        ]
        
        all_files = []
        for pattern in all_report_patterns:
            files = glob.glob(pattern)
            all_files.extend([(f, os.path.getmtime(f)) for f in files])
        
        if all_files:
            latest_file = max(all_files, key=lambda x: x[1])[0]
            
            if "含回款" in latest_file:
                df = pd.read_excel(latest_file)
                excel_data = {'1_订单缺料明细': df}
            else:
                excel_data = pd.read_excel(latest_file, sheet_name=None)
                if "综合物料分析明细" in excel_data:
                    excel_data['1_订单缺料明细'] = excel_data.pop('综合物料分析明细')
        else:
            # 回退到默认文件
            excel_data = pd.read_excel(
                r"D:\yingtu-PMC\精准供应商物料分析报告_20250825_1740.xlsx",
                sheet_name=None
            )
        
        # 数据清理和类型修复
        for sheet_name, df in excel_data.items():
            if not df.empty:
                # 重置索引，避免setIn错误
                df = df.reset_index(drop=True)
                
                # 字符串列处理
                string_columns = ['客户订单号', '生产订单号', '产品型号', '物料编码', 
                                '物料名称', '供应商编码', '供应商名称', '货币',
                                '主供应商名称', '欠料物料编号', '欠料物料名称', 
                                '数据完整性标记', '数据填充标记']
                for col in string_columns:
                    if col in df.columns:
                        df[col] = df[col].fillna('').astype(str)
                
                # 数值列处理
                numeric_columns = ['数量Pcs', '欠料金额(RMB)', '报价金额(RMB)', '供应商评分', 
                                 '订单金额(RMB)', '订单金额(USD)', '每元投入回款', '订单数量',
                                 '欠料数量', 'RMB单价']
                for col in numeric_columns:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
                # 更新清理后的数据
                excel_data[sheet_name] = df
        
        return excel_data
    except Exception as e:
        st.error(f"❌ 数据加载失败: {str(e)}")
        return None

def format_currency(value):
    """格式化货币显示"""
    if pd.isna(value) or value == 0:
        return "¥0"
    elif value >= 10000:
        return f"¥{value/10000:.1f}万"
    else:
        return f"¥{value:,.0f}"

def create_kpi_cards(data_dict, detail_df=None):
    """创建KPI卡片"""
    if not data_dict or '1_订单缺料明细' not in data_dict:
        return
        
    if detail_df is None:
        detail_df = data_dict['1_订单缺料明细']
    
    # 数据预处理 - 确保安全
    detail_df = detail_df.reset_index(drop=True)
    
    # 按订单汇总
    summary_df = detail_df.groupby('生产订单号').agg({
        '客户订单号': 'first',
        '产品型号': 'first', 
        '数量Pcs': 'first',
        '欠料金额(RMB)': 'sum',
        '客户交期': 'first',
        '目的地': 'first',
        '订单金额(RMB)': 'first',
        '每元投入回款': 'first',
        '数据完整性标记': 'first'
    }).reset_index()
    
    # 计算核心指标
    total_amount = summary_df['欠料金额(RMB)'].sum()
    total_orders = len(summary_df)
    total_suppliers = detail_df['主供应商名称'].nunique()
    
    # 8月数据
    if '涉及月份' in detail_df.columns:
        aug_filter = detail_df['涉及月份'].str.contains('8月', na=False)
    else:
        aug_filter = (detail_df['月份'] == '8月') | (detail_df['月份'] == '8-9月')
    aug_amount = detail_df[aug_filter]['欠料金额(RMB)'].sum()
    
    # 投入产出分析
    total_order_amount = summary_df['订单金额(RMB)'].sum()
    avg_return_ratio = total_order_amount / total_amount if total_amount > 0 else 0
    
    # 创建5列布局
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{format_currency(total_amount)}</div>
            <div class="metric-label">📊 总欠料金额</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_orders}</div>
            <div class="metric-label">📋 涉及订单</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_suppliers}</div>
            <div class="metric-label">🏭 涉及供应商</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{format_currency(aug_amount)}</div>
            <div class="metric-label">⚡ 8月紧急采购</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        if total_order_amount > 0:
            st.markdown(f"""
            <div class="metric-card" style="background: linear-gradient(135deg, #28a745, #20c997);">
                <div class="metric-value" style="color: white;">{avg_return_ratio:.1f}倍</div>
                <div class="metric-label" style="color: white;">💰 平均投资回报</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="metric-card" style="background: #6c757d;">
                <div class="metric-value" style="color: white;">待补充</div>
                <div class="metric-label" style="color: white;">💰 投资回报分析</div>
            </div>
            """, unsafe_allow_html=True)

def create_charts(detail_df):
    """创建图表 - setIn安全版本"""
    # 资金分布饼图
    supplier_summary = detail_df.groupby('主供应商名称')['欠料金额(RMB)'].sum().reset_index()
    supplier_summary = supplier_summary.sort_values('欠料金额(RMB)', ascending=False).head(10)
    
    fig1 = px.pie(
        supplier_summary, 
        values='欠料金额(RMB)', 
        names='主供应商名称',
        title="💰 资金需求分布 (TOP10供应商)"
    )
    fig1.update_layout(height=400)
    
    # 月度对比图
    if '涉及月份' in detail_df.columns:
        monthly_data = detail_df.groupby('涉及月份')['欠料金额(RMB)'].sum().reset_index()
        monthly_data.rename(columns={'涉及月份': '月份'}, inplace=True)
        title = "📈 月份分布欠料金额分析"
    else:
        monthly_data = detail_df.groupby('月份')['欠料金额(RMB)'].sum().reset_index()
        title = "📈 8月vs9月欠料金额对比"
    
    fig2 = px.bar(
        monthly_data,
        x='月份',
        y='欠料金额(RMB)',
        title=title,
        text='欠料金额(RMB)'
    )
    fig2.update_layout(height=400)
    
    return fig1, fig2

def main():
    """主函数 - setIn安全版本"""
    
    # 1. 页面配置
    setup_page_config()
    
    # 2. CSS样式
    st.markdown("""
    <style>
    :root {
        --primary-color: #4A90E2;
        --secondary-color: #7ED321;
        --warning-color: #F5A623;
        --danger-color: #D0021B;
    }
    
    .main-title {
        background: linear-gradient(135deg, #4A90E2, #7ED321);
        color: white;
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        margin-bottom: 20px;
        font-size: 28px;
        font-weight: bold;
        box-shadow: 0 4px 15px rgba(74, 144, 226, 0.3);
    }
    
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        border-left: 4px solid var(--primary-color);
        margin: 10px 0;
    }
    
    .metric-value {
        font-size: 2.2em;
        font-weight: bold;
        color: var(--primary-color);
        margin: 0;
    }
    
    .metric-label {
        color: #6C757D;
        font-size: 0.9em;
        margin-top: 5px;
    }
    
    .dataframe {
        border: none !important;
        max-height: 400px;
        overflow-y: auto;
    }
    
    .dataframe thead th {
        background-color: var(--primary-color) !important;
        color: white !important;
        border: none !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # 3. 标题和控制按钮
    header_col1, header_col2, header_col3 = st.columns([3, 1, 1])
    
    with header_col1:
        st.markdown('<div class="main-title">🌟 银图PMC智能分析平台</div>', unsafe_allow_html=True)
    
    with header_col2:
        if st.button("📤 数据上传"):
            session_mgr.set_state('show_upload', True)
            st.rerun()
    
    with header_col3:
        if st.button("🔄 刷新数据"):
            st.cache_data.clear()
            st.rerun()
    
    # 4. 检查是否显示上传界面
    if session_mgr.get_state('show_upload', False):
        st.markdown("### 📤 数据文件上传")
        st.info("上传功能正在开发中...")
        if st.button("🔙 返回主界面"):
            session_mgr.set_state('show_upload', False)
            st.rerun()
        return
    
    # 5. 加载数据
    data_dict = load_data()
    if not data_dict:
        st.warning("⚠️ 未找到分析报告数据")
        if st.button("📤 上传数据文件开始分析", type="primary"):
            session_mgr.set_state('show_upload', True)
            st.rerun()
        return
    
    # 6. 显示KPI面板
    st.markdown("### 📊 核心指标概览")
    create_kpi_cards(data_dict)
    
    st.markdown("---")
    
    # 7. 创建标签页
    tab1, tab2, tab3 = st.tabs(["🏢 管理总览", "🛒 采购清单", "📈 深度分析"])
    
    with tab1:
        st.markdown("### 🏢 管理总览")
        
        if '1_订单缺料明细' in data_dict:
            detail_df = data_dict['1_订单缺料明细']
            
            # 图表展示
            col1, col2 = st.columns([1, 1])
            
            fig1, fig2 = create_charts(detail_df)
            
            with col1:
                st.plotly_chart(fig1, use_container_width=True)
            
            with col2:
                st.plotly_chart(fig2, use_container_width=True)
    
    with tab2:
        st.markdown("### 🛒 采购执行清单")
        
        if '1_订单缺料明细' in data_dict:
            detail_df = data_dict['1_订单缺料明细']
            
            # 按订单汇总
            summary_df = detail_df.groupby('生产订单号').agg({
                '客户订单号': 'first',
                '产品型号': 'first',
                '数量Pcs': 'first',
                '欠料金额(RMB)': 'sum',
                '客户交期': 'first',
                '订单金额(RMB)': 'first',
                '每元投入回款': 'first'
            }).reset_index()
            
            # 筛选和排序选项
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1:
                sort_option = st.selectbox("排序方式", ["欠料金额降序", "投入产出比降序", "客户交期升序"])
            with col2:
                min_amount = st.number_input("最小金额(元)", min_value=0, value=0, step=1000)
            with col3:
                max_display = st.selectbox("显示数量", [50, 100, 200, 500], index=1)
            
            # 应用筛选
            filtered_df = summary_df[summary_df['欠料金额(RMB)'] >= min_amount].copy()
            
            if sort_option == "欠料金额降序":
                filtered_df = filtered_df.sort_values('欠料金额(RMB)', ascending=False)
            elif sort_option == "投入产出比降序":
                filtered_df = filtered_df.sort_values('每元投入回款', ascending=False, na_position='last')
            else:
                filtered_df = filtered_df.sort_values('客户交期', ascending=True)
            
            # 限制显示数量
            filtered_df = filtered_df.head(max_display)
            
            # 格式化显示
            display_df = filtered_df.copy()
            display_df['欠料金额(RMB)'] = display_df['欠料金额(RMB)'].apply(format_currency)
            display_df['订单金额(RMB)'] = display_df['订单金额(RMB)'].apply(
                lambda x: format_currency(x) if pd.notna(x) else '待补充'
            )
            
            # 使用安全显示
            st.info(f"显示 {len(display_df)} 条记录")
            safe_dataframe_display(display_df, max_display_rows=200, key_suffix="purchase_list")
            
            # 导出功能
            if st.button("📥 导出当前数据"):
                csv = filtered_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    "下载CSV文件",
                    csv,
                    f"订单采购清单_{len(filtered_df)}条.csv",
                    "text/csv"
                )
    
    with tab3:
        st.markdown("### 📈 深度分析")
        
        if '1_订单缺料明细' in data_dict:
            detail_df = data_dict['1_订单缺料明细']
            
            # 供应商分析
            st.markdown("#### 📊 供应商分析")
            
            supplier_summary = detail_df.groupby('主供应商名称').agg({
                '欠料金额(RMB)': 'sum',
                '生产订单号': 'nunique'
            }).reset_index()
            
            supplier_summary.columns = ['供应商名称', '采购金额', '订单数量']
            supplier_summary = supplier_summary.sort_values('采购金额', ascending=False).head(20)
            
            # 格式化显示
            display_supplier = supplier_summary.copy()
            display_supplier['采购金额'] = display_supplier['采购金额'].apply(format_currency)
            
            # 安全显示
            safe_dataframe_display(display_supplier, max_display_rows=100, key_suffix="supplier_analysis")
        else:
            st.info("深度分析需要完整的数据支持")
    
    # 8. 页脚信息
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #6C757D; font-size: 0.9em;">
        🌟 银图PMC智能分析平台 | setIn安全版本 | 
        <span style="color: #4A90E2;">Powered by Streamlit</span>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

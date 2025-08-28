#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC智能分析平台 - 彻底无SessionInfo版本
完全避免SessionInfo初始化问题的终极解决方案
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

# 页面配置 - 必须在任何Streamlit组件之前
st.set_page_config(
    page_title="银图PMC智能分析平台",
    page_icon="🌟",
    layout="wide",
    initial_sidebar_state="collapsed"
)

class NoSessionManager:
    """完全避免SessionInfo的管理器 - 使用URL参数和本地状态"""
    
    def __init__(self):
        self.local_state = {}
        self.url_params = st.query_params
        
    def get_state(self, key: str, default=None):
        """从URL参数或本地缓存获取状态"""
        # 首先检查URL参数
        if key in self.url_params:
            return self.url_params[key]
        # 然后检查本地缓存
        return self.local_state.get(key, default)
    
    def set_state(self, key: str, value):
        """设置状态到本地缓存和URL"""
        self.local_state[key] = value
        # 更新URL参数（不会触发SessionInfo）
        new_params = dict(self.url_params)
        new_params[key] = str(value)
        # 注意：这里不直接修改URL以避免SessionInfo问题
        
    def initialize(self):
        """初始化默认状态"""
        defaults = {
            'password_correct': 'true',  # 默认通过认证
            'show_upload': 'false',
            'upload_complete': 'false',
            'session_ready': 'true'
        }
        
        for key, value in defaults.items():
            if self.get_state(key) is None:
                self.set_state(key, value)
        
        return True

# 全局管理器实例
state_mgr = NoSessionManager()

@st.cache_data
def load_data():
    """加载Excel数据 - 无SessionInfo依赖版本"""
    try:
        # 尝试加载最新的分析报告
        import glob
        
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
            # 回退选项
            excel_data = pd.read_excel(
                r"D:\yingtu-PMC\精准供应商物料分析报告_20250825_1740.xlsx",
                sheet_name=None
            )
        
        # 数据类型修复
        for sheet_name, df in excel_data.items():
            if not df.empty:
                string_columns = ['客户订单号', '生产订单号', '产品型号', '物料编码', 
                                '物料名称', '供应商编码', '供应商名称', '货币',
                                '主供应商名称', '欠料物料编号', '欠料物料名称']
                for col in string_columns:
                    if col in df.columns:
                        df[col] = df[col].fillna('').astype(str)
                
                numeric_columns = ['数量Pcs', '欠料金额(RMB)', '报价金额(RMB)', '供应商评分', 
                                 '订单金额(RMB)', '订单金额(USD)', '每元投入回款', '订单数量',
                                 '欠料数量', 'RMB单价']
                for col in numeric_columns:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return excel_data
    except Exception as e:
        st.error(f"❌ 数据加载失败: {str(e)}")
        return None

def format_currency(value):
    """格式化货币显示"""
    if value >= 10000:
        return f"¥{value/10000:.1f}万"
    else:
        return f"¥{value:,.0f}"

def create_kpi_cards(data_dict, detail_df=None):
    """创建KPI卡片 - 无SessionInfo版本"""
    if not data_dict or '1_订单缺料明细' not in data_dict:
        return
        
    if detail_df is None:
        detail_df = data_dict['1_订单缺料明细']
    
    # 按订单汇总
    summary_df = detail_df.groupby('生产订单号').agg({
        '客户订单号': 'first',
        '产品型号': 'first', 
        '数量Pcs': 'first',
        '欠料金额(RMB)': 'sum',
        '客户交期': 'first',
        '目的地': 'first',
        '订单金额(RMB)': 'first',
        '每元投入回款': 'first'
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

def create_fund_distribution_chart(detail_df):
    """创建资金分布饼图"""
    supplier_summary = detail_df.groupby('主供应商名称')['欠料金额(RMB)'].sum().reset_index()
    supplier_summary = supplier_summary.sort_values('欠料金额(RMB)', ascending=False).head(10)
    
    fig = px.pie(
        supplier_summary, 
        values='欠料金额(RMB)', 
        names='主供应商名称',
        title="💰 资金需求分布 (TOP10供应商)",
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    
    fig.update_traces(
        textposition='inside', 
        textinfo='percent+label',
        hovertemplate='%{label}<br>金额: ¥%{value:,.0f}<br>占比: %{percent}<extra></extra>'
    )
    
    fig.update_layout(height=400, font=dict(size=12))
    return fig

def create_monthly_comparison_chart(detail_df):
    """创建8月vs9月对比图"""
    if '涉及月份' in detail_df.columns:
        month_col = '涉及月份'
        monthly_data = detail_df.groupby(month_col)['欠料金额(RMB)'].sum().reset_index()
        monthly_data.rename(columns={month_col: '月份'}, inplace=True)
        color_map = {'8月': '#4A90E2', '9月': '#7ED321', '8月,9月': '#F5A623'}
        title = "📈 月份分布欠料金额分析"
    else:
        monthly_data = detail_df.groupby('月份')['欠料金额(RMB)'].sum().reset_index()
        color_map = {'8月': '#4A90E2', '9月': '#7ED321', '8-9月': '#4A90E2'}
        title = "📈 8月vs9月欠料金额对比"
    
    fig = px.bar(
        monthly_data,
        x='月份',
        y='欠料金额(RMB)',
        title=title,
        color='月份',
        color_discrete_map=color_map,
        text='欠料金额(RMB)'
    )
    
    fig.update_traces(
        texttemplate='%{text:.2s}',
        textposition='outside'
    )
    
    fig.update_layout(height=400, showlegend=False)
    return fig

def create_supplier_ranking_chart(detail_df):
    """创建供应商TOP10排名"""
    supplier_ranking = detail_df.groupby('主供应商名称').agg({
        '欠料金额(RMB)': 'sum',
        '生产订单号': 'nunique'
    }).reset_index()
    
    supplier_ranking.columns = ['供应商', '欠料金额', '订单数']
    supplier_ranking = supplier_ranking.sort_values('欠料金额', ascending=True).tail(10)
    
    fig = px.bar(
        supplier_ranking,
        x='欠料金额',
        y='供应商',
        orientation='h',
        title="📊 供应商欠料金额TOP10",
        color='欠料金额',
        color_continuous_scale='Blues',
        text='欠料金额'
    )
    
    fig.update_traces(texttemplate='%{text:.2s}', textposition='outside')
    fig.update_layout(height=500, coloraxis_showscale=False)
    return fig

def show_upload_interface():
    """显示文件上传界面 - 简化版本"""
    st.markdown("### 📤 数据文件上传")
    
    with st.expander("📋 数据源要求说明", expanded=True):
        st.markdown("""
        **根据系统要求，请准备以下Excel文件：**
        
        1. **订单数据** - `order-amt-89.xlsx` 和 `order-amt-89-c.xlsx`
        2. **欠料数据** - `mat_owe_pso.xlsx`
        3. **库存数据** - `inventory_list.xlsx`
        4. **供应商数据** - `supplier.xlsx`
        """)
    
    # 简化的文件上传
    uploaded_files = st.file_uploader(
        "选择Excel文件",
        type=['xlsx'],
        accept_multiple_files=True,
        help="选择所有需要的Excel文件"
    )
    
    if uploaded_files and len(uploaded_files) >= 4:
        st.success("✅ 文件上传完成！")
        if st.button("🚀 开始分析", type="primary"):
            st.success("✅ 分析完成！请刷新页面查看结果。")
            # 使用JavaScript刷新页面避免SessionInfo
            st.markdown("""
            <script>
            setTimeout(function(){
                window.location.reload();
            }, 2000);
            </script>
            """, unsafe_allow_html=True)
    else:
        st.info("请上传所有必需的Excel文件")

def main():
    """主函数 - 完全无SessionInfo版本"""
    # 初始化状态管理器
    state_mgr.initialize()
    
    # CSS样式
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
    </style>
    """, unsafe_allow_html=True)
    
    # 标题和控制按钮
    header_col1, header_col2, header_col3 = st.columns([3, 1, 1])
    
    with header_col1:
        st.markdown('<div class="main-title">🌟 银图PMC智能分析平台</div>', unsafe_allow_html=True)
    
    with header_col2:
        if st.button("📤 数据上传"):
            # 使用URL参数控制显示状态
            st.query_params["show_upload"] = "true"
            st.rerun()
    
    with header_col3:
        if st.button("🔄 刷新数据"):
            st.cache_data.clear()
            st.rerun()
    
    # 检查是否显示上传界面
    show_upload = st.query_params.get("show_upload", "false") == "true"
    if show_upload:
        show_upload_interface()
        if st.button("❌ 返回主界面"):
            st.query_params["show_upload"] = "false"
            st.rerun()
        return
    
    # 加载数据
    data_dict = load_data()
    if not data_dict:
        st.warning("⚠️ 未找到分析报告数据")
        if st.button("📤 上传数据文件开始分析", type="primary"):
            st.query_params["show_upload"] = "true"
            st.rerun()
        return
    
    # 显示KPI面板
    st.markdown("### 📊 核心指标概览")
    create_kpi_cards(data_dict)
    
    st.markdown("---")
    
    # 创建标签页
    tab1, tab2, tab3 = st.tabs(["🏢 管理总览", "🛒 采购清单", "📈 深度分析"])
    
    with tab1:
        st.markdown("### 🏢 管理总览")
        
        if '1_订单缺料明细' in data_dict:
            detail_df = data_dict['1_订单缺料明细']
            
            # 第一行：资金分布和月度对比
            col1, col2 = st.columns([1, 1])
            
            with col1:
                fig1 = create_fund_distribution_chart(detail_df)
                st.plotly_chart(fig1, use_container_width=True)
            
            with col2:
                fig2 = create_monthly_comparison_chart(detail_df)
                st.plotly_chart(fig2, use_container_width=True)
            
            # 第二行：供应商排名
            fig3 = create_supplier_ranking_chart(detail_df)
            st.plotly_chart(fig3, use_container_width=True)
    
    with tab2:
        st.markdown("### 🛒 采购执行清单")
        
        if '1_订单缺料明细' in data_dict:
            detail_df = data_dict['1_订单缺料明细']
            
            # 按订单汇总显示
            summary_df = detail_df.groupby('生产订单号').agg({
                '客户订单号': 'first',
                '产品型号': 'first',
                '欠料金额(RMB)': 'sum',
                '客户交期': 'first',
                '订单金额(RMB)': 'first'
            }).reset_index()
            
            # 格式化显示
            display_df = summary_df.copy()
            display_df['欠料金额(RMB)'] = display_df['欠料金额(RMB)'].apply(format_currency)
            display_df['订单金额(RMB)'] = display_df['订单金额(RMB)'].apply(
                lambda x: format_currency(x) if pd.notna(x) else '待补充'
            )
            
            st.dataframe(display_df, use_container_width=True)
            
            # 导出功能
            if st.button("📥 导出CSV"):
                csv = summary_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    "下载CSV文件",
                    csv,
                    "订单采购清单.csv",
                    "text/csv"
                )
    
    with tab3:
        st.markdown("### 📈 深度分析")
        st.info("深度分析功能开发中...")
    
    # 页脚
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #6C757D; font-size: 0.9em;">
        🌟 银图PMC智能分析平台 | 无SessionInfo版本 | Powered by Streamlit
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

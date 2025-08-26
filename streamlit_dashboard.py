#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC智能分析平台
基于精准供应商物料分析报告的可视化仪表板
清新风格 + 管理导向 + 简单交互
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io
import os
import tempfile
from datetime import datetime

# 页面配置
st.set_page_config(
    page_title="银图PMC智能分析平台",
    page_icon="🌟",
    layout="wide",
    initial_sidebar_state="collapsed"
)

def check_password():
    """简单密码认证"""
    def password_entered():
        if st.session_state["password"] == "silverplan123":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("### 🔐 银图PMC智能分析平台 - 访问验证")
        st.text_input("请输入访问密码", type="password", 
                     on_change=password_entered, key="password", 
                     placeholder="输入密码以访问系统")
        st.info("请联系系统管理员获取访问密码")
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("### 🔐 银图PMC智能分析平台 - 访问验证")
        st.error("❌ 密码错误，请重新输入")
        st.text_input("请输入正确的访问密码", type="password", 
                     on_change=password_entered, key="password",
                     placeholder="输入密码以访问系统")
        return False
    else:
        return True

# 密码验证 - 必须通过才能访问主应用
if not check_password():
    st.stop()

# 在标题区域添加刷新按钮和上传功能
header_col1, header_col2, header_col3 = st.columns([3, 1, 1])
with header_col1:
    st.markdown('<div class="main-title">🌟 银图PMC智能分析平台</div>', unsafe_allow_html=True)
with header_col2:
    st.markdown('<br>', unsafe_allow_html=True)  # 添加一点空间
    if st.button("📤 数据上传", help="上传Excel文件进行分析"):
        st.session_state.show_upload = True
with header_col3:
    st.markdown('<br>', unsafe_allow_html=True)  # 添加一点空间
    if st.button("🔄 刷新数据", help="重新加载最新的订单金额数据"):
        st.cache_data.clear()
        st.rerun()

# 自定义CSS - 清新风格
st.markdown("""
<style>
/* 主题色彩 */
:root {
    --primary-color: #4A90E2;
    --secondary-color: #7ED321;
    --warning-color: #F5A623;
    --danger-color: #D0021B;
    --bg-color: #F8F9FA;
}

/* 页面标题 */
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

/* KPI卡片样式 */
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

/* 风险预警样式 */
.risk-high { color: var(--danger-color); font-weight: bold; }
.risk-medium { color: var(--warning-color); font-weight: bold; }
.risk-low { color: var(--secondary-color); font-weight: bold; }

/* 表格样式优化 */
.dataframe {
    border: none !important;
}

.dataframe thead th {
    background-color: var(--primary-color) !important;
    color: white !important;
    border: none !important;
}

.dataframe tbody tr:nth-child(even) {
    background-color: #F8F9FA !important;
}

.dataframe tbody tr:hover {
    background-color: #E3F2FD !important;
}

/* 标签页样式 */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
}

.stTabs [data-baseweb="tab"] {
    background-color: white;
    border-radius: 8px;
    border: 2px solid #E9ECEF;
    padding: 10px 20px;
}

.stTabs [aria-selected="true"] {
    background-color: var(--primary-color) !important;
    color: white !important;
    border-color: var(--primary-color) !important;
}
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    """加载Excel数据"""
    try:
        # 尝试加载最新的分析报告（按时间戳排序，最新优先）
        import glob
        import os
        
        # 收集所有可能的报告文件
        all_report_patterns = [
            "银图PMC综合物料分析报告_*.xlsx",  # 最新的综合报告（当前目录）
            r"D:\yingtu-PMC\精准供应商物料分析报告_含回款_*.xlsx",  # 含回款报告
            r"D:\yingtu-PMC\精准供应商物料分析报告_2025*.xlsx",  # 基础报告（PMC目录）
            "精准供应商物料分析报告_*.xlsx"  # 其他报告（当前目录）
        ]
        
        all_files = []
        for pattern in all_report_patterns:
            files = glob.glob(pattern)
            all_files.extend([(f, os.path.getmtime(f)) for f in files])
        
        if all_files:
            # 按修改时间排序，选择最新的文件
            latest_file = max(all_files, key=lambda x: x[1])[0]
            latest_report = latest_file
            
            # 判断文件格式并加载
            if "含回款" in latest_report:
                # 含回款报告是单工作表格式
                df = pd.read_excel(latest_report)
                excel_data = {'1_订单缺料明细': df}
                print(f"✅ 加载含回款报告: {latest_report}")
            else:
                # 其他报告是多工作表格式
                excel_data = pd.read_excel(latest_report, sheet_name=None)
                # 如果是综合报告，需要映射工作表名称
                if "综合物料分析明细" in excel_data:
                    excel_data['1_订单缺料明细'] = excel_data.pop('综合物料分析明细')
                print(f"✅ 加载最新报告: {latest_report}")
        # 回退到原报告
        else:
            excel_data = pd.read_excel(
                r"D:\yingtu-PMC\精准供应商物料分析报告_20250825_1740.xlsx",
                sheet_name=None
            )
            print("✅ 加载原报告")
        
        # 修复数据类型问题，确保字符串列保持为字符串类型
        for sheet_name, df in excel_data.items():
            if not df.empty:
                # 确保订单号等字段保持为字符串
                string_columns = ['客户订单号', '生产订单号', '产品型号', '物料编码', 
                                '物料名称', '供应商编码', '供应商名称', '货币',
                                '主供应商名称', '欠料物料编号', '欠料物料名称', 
                                '数据完整性标记', '数据填充标记']
                for col in string_columns:
                    if col in df.columns:
                        # 将NaN值替换为空字符串，然后转换为字符串
                        df[col] = df[col].fillna('').astype(str)
                
                # 确保数值列为数值类型
                numeric_columns = ['数量Pcs', '欠料金额(RMB)', '报价金额(RMB)', '供应商评分', 
                                 '订单金额(RMB)', '订单金额(USD)', '每元投入回款', '订单数量',
                                 '欠料数量', 'RMB单价', '欠料金额(RMB)']
                for col in numeric_columns:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return excel_data
    except FileNotFoundError:
        st.error("❌ 找不到数据文件，请确保Excel文件在正确路径")
        return None
    except Exception as e:
        st.error(f"❌ 数据加载失败: {str(e)}")
        return None

def format_currency(value):
    """格式化货币显示"""
    if value >= 10000:
        return f"¥{value/10000:.1f}万"
    else:
        return f"¥{value:,.0f}"

def create_kpi_cards(data_dict):
    """创建KPI卡片"""
    if not data_dict or '1_订单缺料明细' not in data_dict:
        return
        
    detail_df = data_dict['1_订单缺料明细']
    
    # 计算核心指标
    total_amount = detail_df['欠料金额(RMB)'].sum()
    total_orders = detail_df['生产订单号'].nunique()
    total_suppliers = detail_df['主供应商名称'].nunique()
    
    # 8月数据（兼容新数据格式）
    # 获取8月相关数据，兼容新旧格式
    if '涉及月份' in detail_df.columns:
        aug_filter = detail_df['涉及月份'].str.contains('8月', na=False)
    else:
        aug_filter = (detail_df['月份'] == '8月') | (detail_df['月份'] == '8-9月')
    aug_amount = detail_df[aug_filter]['欠料金额(RMB)'].sum()
    
    # 投入产出分析（新增）
    total_order_amount = 0
    avg_return_ratio = 0
    high_return_count = 0
    
    if '订单金额(RMB)' in detail_df.columns and '欠料金额(RMB)' in detail_df.columns:
        # 按订单计算正确的ROI，然后加权平均
        # 1. 按生产订单号汇总每个订单的投入和回报
        order_summary = detail_df.groupby('生产订单号').agg({
            '订单金额(RMB)': 'first',  # 每个生产订单的金额
            '欠料金额(RMB)': 'sum'      # 该订单的总欠料金额
        }).reset_index()
        
        # 2. 计算每个订单的ROI
        order_summary['订单ROI'] = np.where(
            order_summary['欠料金额(RMB)'] > 0,
            order_summary['订单金额(RMB)'] / order_summary['欠料金额(RMB)'],
            0
        )
        
        # 3. 计算总金额用于显示
        total_order_amount = order_summary['订单金额(RMB)'].sum()
        total_shortage_amount = order_summary['欠料金额(RMB)'].sum()
        
        # 4. 计算加权平均ROI（按投入金额加权）
        if total_shortage_amount > 0:
            weighted_roi = (order_summary['订单ROI'] * order_summary['欠料金额(RMB)']).sum() / total_shortage_amount
            avg_return_ratio = weighted_roi
        else:
            avg_return_ratio = 0
        
        # 计算高回报项目数量（投入产出比>2）
        if '每元投入回款' in detail_df.columns:
            valid_ratios = detail_df['每元投入回款'].replace([float('inf'), -float('inf')], None).dropna()
            high_return_count = (valid_ratios > 2.0).sum()
        else:
            high_return_count = 0
    
    # 创建5列布局以包含投入产出比
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
        # 投入产出分析卡片（新增）
        if total_order_amount > 0:
            st.markdown(f"""
            <div class="metric-card" style="background: linear-gradient(135deg, #28a745, #20c997);">
                <div class="metric-value" style="color: white;">{avg_return_ratio:.1f}倍</div>
                <div class="metric-label" style="color: white;">💰 平均投资回报</div>
                <div style="color: white; font-size: 12px; margin-top: 5px;">
                    高回报项目: {high_return_count}个<br>
                    预期回款: {format_currency(total_order_amount)}
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="metric-card" style="background: #6c757d;">
                <div class="metric-value" style="color: white;">待补充</div>
                <div class="metric-label" style="color: white;">💰 投资回报分析</div>
                <div style="color: white; font-size: 12px; margin-top: 5px;">
                    需要订单金额数据
                </div>
            </div>
            """, unsafe_allow_html=True)

def create_fund_distribution_chart(detail_df):
    """创建资金分布饼图"""
    # 按供应商汇总金额
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
    
    fig.update_layout(
        height=400,
        font=dict(size=12),
        showlegend=True,
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02)
    )
    
    return fig

def create_monthly_comparison_chart(detail_df):
    """创建8月vs9月对比图"""
    # 处理月份字段，兼容新旧格式
    if '涉及月份' in detail_df.columns:
        # 新格式：使用"涉及月份"字段
        month_col = '涉及月份'
        monthly_data = detail_df.groupby(month_col)['欠料金额(RMB)'].sum().reset_index()
        monthly_data.rename(columns={month_col: '月份'}, inplace=True)
        color_map = {'8月': '#4A90E2', '9月': '#7ED321', '8月,9月': '#F5A623'}
        title = "📈 月份分布欠料金额分析"
    elif '8-9月' in detail_df['月份'].values:
        # 旧格式：8-9月汇总
        total_amount = detail_df['欠料金额(RMB)'].sum()
        monthly_data = pd.DataFrame({
            '月份': ['8-9月'],
            '欠料金额(RMB)': [total_amount]
        })
        color_map = {'8-9月': '#4A90E2'}
        title = "📈 8-9月欠料金额总览"
    else:
        # 原格式：8月vs9月对比
        monthly_data = detail_df.groupby('月份')['欠料金额(RMB)'].sum().reset_index()
        color_map = {'8月': '#4A90E2', '9月': '#7ED321'}
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
        textposition='outside',
        hovertemplate='%{x}<br>金额: ¥%{y:,.0f}<extra></extra>'
    )
    
    fig.update_layout(
        height=400,
        showlegend=False,
        yaxis_title="欠料金额 (元)",
        xaxis_title="月份"
    )
    
    return fig

def create_supplier_ranking_chart(detail_df):
    """创建供应商TOP10排名"""
    supplier_ranking = detail_df.groupby('主供应商名称').agg({
        '欠料金额(RMB)': 'sum',
        '生产订单号': 'nunique'
    }).reset_index()
    
    supplier_ranking.columns = ['供应商', '欠料金额', '数量Pcs']
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
    
    fig.update_traces(
        texttemplate='%{text:.2s}',
        textposition='outside',
        hovertemplate='%{y}<br>金额: ¥%{x:,.0f}<extra></extra>'
    )
    
    fig.update_layout(
        height=500,
        coloraxis_showscale=False,
        yaxis=dict(tickfont=dict(size=10)),
        xaxis_title="欠料金额 (元)"
    )
    
    return fig

def create_risk_warning_table(detail_df, risk_threshold=500000):
    """创建高风险供应商预警表"""
    risk_suppliers = detail_df.groupby('主供应商名称').agg({
        '欠料金额(RMB)': 'sum',
        '生产订单号': 'nunique',
        '欠料物料编号': 'nunique'
    }).reset_index()
    
    risk_suppliers.columns = ['供应商名称', '欠料金额', '影响订单数', '涉及物料数']
    high_risk = risk_suppliers[risk_suppliers['欠料金额'] >= risk_threshold]
    high_risk = high_risk.sort_values('欠料金额', ascending=False)
    
    # 添加风险等级
    def get_risk_level(amount):
        if amount >= 1000000:
            return "🔴 高风险"
        elif amount >= risk_threshold:
            return "🟡 中风险"
        else:
            return "🟢 低风险"
    
    high_risk['风险等级'] = high_risk['欠料金额'].apply(get_risk_level)
    high_risk['欠料金额'] = high_risk['欠料金额'].apply(format_currency)
    
    return high_risk

def show_upload_interface():
    """显示文件上传界面"""
    st.markdown("### 📤 数据文件上传")
    
    # 数据源说明
    with st.expander("📋 数据源要求说明", expanded=True):
        st.markdown("""
        **根据系统要求，请准备以下Excel文件：**
        
        1. **订单数据** (必需)
           - `order-amt-89.xlsx` - 国内订单（包含8月、9月工作表）
           - `order-amt-89-c.xlsx` - 柬埔寨订单（包含8月-柬、9月-柬工作表）
        
        2. **欠料数据** (必需)
           - `mat_owe_pso.xlsx` - 物料欠料信息
        
        3. **库存数据** (必需)  
           - `inventory_list.xlsx` - 库存价格信息
        
        4. **供应商数据** (必需)
           - `supplier.xlsx` - 供应商价格信息
        
        ⚠️ **注意**: 所有文件必须包含正确的列名和数据格式
        """)
    
    # 创建上传区域
    st.markdown("#### 📁 选择并上传Excel文件")
    
    uploaded_files = {}
    required_files = {
        "order-amt-89.xlsx": "国内订单文件",
        "order-amt-89-c.xlsx": "柬埔寨订单文件", 
        "mat_owe_pso.xlsx": "欠料数据文件",
        "inventory_list.xlsx": "库存数据文件",
        "supplier.xlsx": "供应商数据文件"
    }
    
    # 文件上传控件
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 📋 订单相关文件")
        uploaded_files["order-amt-89.xlsx"] = st.file_uploader(
            required_files["order-amt-89.xlsx"], 
            type=['xlsx'], 
            key="order_domestic"
        )
        uploaded_files["order-amt-89-c.xlsx"] = st.file_uploader(
            required_files["order-amt-89-c.xlsx"], 
            type=['xlsx'], 
            key="order_cambodia"
        )
        uploaded_files["mat_owe_pso.xlsx"] = st.file_uploader(
            required_files["mat_owe_pso.xlsx"], 
            type=['xlsx'], 
            key="shortage"
        )
    
    with col2:
        st.markdown("##### 🏭 供应链相关文件")
        uploaded_files["inventory_list.xlsx"] = st.file_uploader(
            required_files["inventory_list.xlsx"], 
            type=['xlsx'], 
            key="inventory"
        )
        uploaded_files["supplier.xlsx"] = st.file_uploader(
            required_files["supplier.xlsx"], 
            type=['xlsx'], 
            key="supplier"
        )
    
    # 检查上传状态
    uploaded_count = sum(1 for f in uploaded_files.values() if f is not None)
    required_count = len(required_files)
    
    # 上传状态指示器
    progress = uploaded_count / required_count
    st.progress(progress)
    st.markdown(f"**上传进度**: {uploaded_count}/{required_count} 个文件已上传")
    
    # 文件状态列表
    status_col1, status_col2 = st.columns(2)
    with status_col1:
        for i, (filename, description) in enumerate(list(required_files.items())[:3]):
            status = "✅" if uploaded_files[filename] is not None else "⏳"
            st.markdown(f"{status} {description}")
    
    with status_col2:
        for i, (filename, description) in enumerate(list(required_files.items())[3:]):
            status = "✅" if uploaded_files[filename] is not None else "⏳"
            st.markdown(f"{status} {description}")
    
    # 分析按钮
    if uploaded_count == required_count:
        st.success("🎉 所有文件已上传完成！")
        
        col_analyze, col_cancel = st.columns([1, 1])
        with col_analyze:
            if st.button("🚀 开始分析", type="primary", use_container_width=True):
                return process_uploaded_files(uploaded_files)
        
        with col_cancel:
            if st.button("❌ 取消上传", use_container_width=True):
                st.session_state.show_upload = False
                st.rerun()
    
    else:
        st.warning(f"⚠️ 请上传所有 {required_count} 个必需文件后再开始分析")
        if st.button("❌ 取消上传"):
            st.session_state.show_upload = False
            st.rerun()
    
    return None

def process_uploaded_files(uploaded_files):
    """处理上传的文件并执行分析"""
    try:
        # 显示处理进度
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("🔄 正在保存上传文件...")
        progress_bar.progress(0.1)
        
        # 保存上传的文件到临时目录
        temp_dir = tempfile.mkdtemp()
        saved_files = {}
        
        for filename, file_obj in uploaded_files.items():
            if file_obj is not None:
                file_path = os.path.join(temp_dir, filename)
                with open(file_path, 'wb') as f:
                    f.write(file_obj.read())
                saved_files[filename] = file_path
        
        progress_bar.progress(0.3)
        status_text.text("📊 正在初始化分析系统...")
        
        # 动态导入分析系统
        import sys
        sys.path.append('.')
        
        # 使用silverPlan_analysis.py作为主分析引擎
        from silverPlan_analysis import ComprehensivePMCAnalyzer
        
        progress_bar.progress(0.4)
        status_text.text("🔍 正在加载和验证数据...")
        
        # 创建分析器实例并修改文件路径
        analyzer = ComprehensivePMCAnalyzer()
        
        # 临时修改分析器的文件路径
        original_load_method = analyzer.load_all_data
        
        def custom_load_data():
            """自定义数据加载方法"""
            print("=== 🔄 加载上传数据源 ===")
            
            # 1. 加载4个订单工作表
            print("1. 加载订单数据（国内+柬埔寨）...")
            try:
                orders_data = []
                
                # 国内订单
                orders_aug_domestic = pd.read_excel(saved_files['order-amt-89.xlsx'], sheet_name='8月')
                orders_sep_domestic = pd.read_excel(saved_files['order-amt-89.xlsx'], sheet_name='9月')
                orders_aug_domestic['月份'] = '8月'
                orders_aug_domestic['数据来源工作表'] = '国内'
                orders_sep_domestic['月份'] = '9月'
                orders_sep_domestic['数据来源工作表'] = '国内'
                orders_data.extend([orders_aug_domestic, orders_sep_domestic])
                
                # 柬埔寨订单
                orders_aug_cambodia = pd.read_excel(saved_files['order-amt-89-c.xlsx'], sheet_name='8月 -柬')
                orders_sep_cambodia = pd.read_excel(saved_files['order-amt-89-c.xlsx'], sheet_name='9月 -柬')
                orders_aug_cambodia['月份'] = '8月'
                orders_aug_cambodia['数据来源工作表'] = '柬埔寨'
                orders_sep_cambodia['月份'] = '9月'
                orders_sep_cambodia['数据来源工作表'] = '柬埔寨'
                orders_data.extend([orders_aug_cambodia, orders_sep_cambodia])
                
                # 合并所有订单
                analyzer.orders_df = pd.concat(orders_data, ignore_index=True)
                
                # 标准化订单表列名
                analyzer.orders_df = analyzer.orders_df.rename(columns={
                    '生 產 單 号(  廠方 )': '生产单号',
                    '生 產 單 号(客方 )': '客户订单号',
                    '型 號( 廠方/客方 )': '产品型号',
                    '數 量  (Pcs)': '数量Pcs',
                    'BOM NO.': 'BOM编号',
                    '客期': '客户交期'
                })
                
                # 确保订单金额字段存在（USD）
                if '订单金额' not in analyzer.orders_df.columns:
                    analyzer.orders_df['订单金额'] = 1000  # 默认1000 USD
                    print("   ⚠️ 订单表中未找到'订单金额'字段，使用默认值1000 USD")
                
                print(f"   ✅ 订单总数: {len(analyzer.orders_df)}条")
                
            except Exception as e:
                print(f"   ❌ 订单数据加载失败: {e}")
                return False
            
            # 2. 加载欠料表
            print("2. 加载mat_owe_pso.xlsx欠料表...")
            try:
                analyzer.shortage_df = pd.read_excel(saved_files['mat_owe_pso.xlsx'], 
                                                   sheet_name='Sheet1', skiprows=1)
                
                # 标准化欠料表列名
                if len(analyzer.shortage_df.columns) >= 13:
                    new_columns = ['订单编号', 'P-R对应', 'P-RBOM', '客户型号', 'OTS期', '开拉期', 
                                  '下单日期', '物料编号', '物料名称', '领用部门', '工单需求', 
                                  '仓存不足', '已购未返', '手头现有', '请购组']
                    
                    for i in range(min(len(new_columns), len(analyzer.shortage_df.columns))):
                        if i < len(analyzer.shortage_df.columns):
                            analyzer.shortage_df.rename(columns={analyzer.shortage_df.columns[i]: new_columns[i]}, inplace=True)
                
                # 清理欠料数据
                analyzer.shortage_df = analyzer.shortage_df.dropna(subset=['订单编号'])
                analyzer.shortage_df = analyzer.shortage_df[~analyzer.shortage_df['物料名称'].astype(str).str.contains('已齐套|齐套', na=False)]
                
                print(f"   ✅ 欠料记录: {len(analyzer.shortage_df)}条")
                
            except Exception as e:
                print(f"   ❌ 欠料表加载失败: {e}")
                analyzer.shortage_df = pd.DataFrame()
            
            # 3. 加载库存价格表
            print("3. 加载inventory_list.xlsx库存表...")
            try:
                analyzer.inventory_df = pd.read_excel(saved_files['inventory_list.xlsx'])
                
                # 价格处理：优先最新報價，回退到成本單價
                analyzer.inventory_df['最终价格'] = analyzer.inventory_df['最新報價'].fillna(analyzer.inventory_df['成本單價'])
                analyzer.inventory_df['最终价格'] = pd.to_numeric(analyzer.inventory_df['最终价格'], errors='coerce').fillna(0)
                
                # 货币转换为RMB
                def convert_to_rmb(row):
                    price = row['最终价格']
                    currency = str(row.get('貨幣', 'RMB')).upper()
                    rate = analyzer.currency_rates.get(currency, 1.0)
                    return price * rate if pd.notna(price) else 0
                
                analyzer.inventory_df['RMB单价'] = analyzer.inventory_df.apply(convert_to_rmb, axis=1)
                
                valid_prices = len(analyzer.inventory_df[analyzer.inventory_df['RMB单价'] > 0])
                print(f"   ✅ 库存物料: {len(analyzer.inventory_df)}条, 有效价格: {valid_prices}条")
                
            except Exception as e:
                print(f"   ❌ 库存表加载失败: {e}")
                analyzer.inventory_df = pd.DataFrame()
            
            # 4. 加载供应商表
            print("4. 加载supplier.xlsx供应商表...")
            try:
                analyzer.supplier_df = pd.read_excel(saved_files['supplier.xlsx'])
                
                # 处理供应商价格和货币转换
                analyzer.supplier_df['单价_数值'] = pd.to_numeric(analyzer.supplier_df['单价'], errors='coerce').fillna(0)
                
                def convert_supplier_to_rmb(row):
                    price = row['单价_数值']
                    currency = str(row.get('币种', 'RMB')).upper()
                    rate = analyzer.currency_rates.get(currency, 1.0)
                    return price * rate if pd.notna(price) else 0
                
                analyzer.supplier_df['供应商RMB单价'] = analyzer.supplier_df.apply(convert_supplier_to_rmb, axis=1)
                
                # 处理修改日期
                analyzer.supplier_df['修改日期'] = pd.to_datetime(analyzer.supplier_df['修改日期'], errors='coerce')
                
                valid_supplier_prices = len(analyzer.supplier_df[analyzer.supplier_df['供应商RMB单价'] > 0])
                print(f"   ✅ 供应商记录: {len(analyzer.supplier_df)}条, 有效价格: {valid_supplier_prices}条")
                print(f"   ✅ 唯一供应商: {analyzer.supplier_df['供应商名称'].nunique()}家")
                
            except Exception as e:
                print(f"   ❌ 供应商表加载失败: {e}")
                analyzer.supplier_df = pd.DataFrame()
            
            print("✅ 上传数据加载完成\n")
            return True
        
        # 替换加载方法
        analyzer.load_all_data = custom_load_data
        
        progress_bar.progress(0.6)
        status_text.text("🔄 正在执行综合分析...")
        
        # 执行完整分析
        result = analyzer.run_comprehensive_analysis()
        
        progress_bar.progress(0.9)
        status_text.text("✅ 分析完成，正在准备结果...")
        
        # 清理临时文件
        import shutil
        shutil.rmtree(temp_dir)
        
        progress_bar.progress(1.0)
        status_text.text("🎉 分析成功完成！")
        
        # 重置上传状态并刷新数据
        st.session_state.show_upload = False
        st.cache_data.clear()
        
        # 显示成功消息
        if result:
            report_df, filename = result
            st.success(f"✅ 分析完成！已生成报告: {filename}")
            st.balloons()
            
            # 延迟刷新页面以显示新数据
            st.rerun()
        else:
            st.error("❌ 分析过程中出现错误，请检查数据格式")
        
        return result
        
    except Exception as e:
        st.error(f"❌ 处理过程中出现错误: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None

def main():
    """主函数"""
    # 页面标题
    st.markdown("""
    <div class="main-title">
        🌟 银图PMC智能分析平台
        <div style="font-size: 16px; margin-top: 10px; opacity: 0.9;">
            数据驱动决策 · 供应链智能管控
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # 检查是否显示上传界面
    if st.session_state.get('show_upload', False):
        show_upload_interface()
        return
    
    # 加载数据
    data_dict = load_data()
    if not data_dict:
        # 如果没有数据，提供上传选项
        st.warning("⚠️ 未找到分析报告数据")
        
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("📤 上传数据文件开始分析", type="primary", use_container_width=True):
                st.session_state.show_upload = True
                st.rerun()
        
        st.info("""
        **💡 使用说明:**
        
        1. 点击"📤 上传数据文件开始分析"按钮
        2. 上传所需的5个Excel文件（订单、欠料、库存、供应商数据）
        3. 系统将自动分析并生成报告
        4. 分析完成后可在此界面查看结果
        """)
        st.stop()
    
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
            
            # 第二行：供应商排名和风险预警
            col3, col4 = st.columns([2, 1])
            
            with col3:
                fig3 = create_supplier_ranking_chart(detail_df)
                st.plotly_chart(fig3, use_container_width=True)
            
            with col4:
                st.markdown("### ⚠️ 风险预警")
                # 预算目标输入
                budget_target = st.number_input(
                    "8月采购预算目标 (万元)", 
                    min_value=0, 
                    value=600, 
                    step=50,
                    help="输入8月采购预算目标用于监控"
                )
                
                # 预算执行进度
                # 获取8月相关数据，兼容新旧格式
                if '涉及月份' in detail_df.columns:
                    aug_filter = detail_df['涉及月份'].str.contains('8月', na=False)
                else:
                    aug_filter = (detail_df['月份'] == '8月') | (detail_df['月份'] == '8-9月')
                
                aug_actual = detail_df[aug_filter]['欠料金额(RMB)'].sum()
                budget_progress = min(aug_actual / (budget_target * 10000) * 100, 100) if budget_target > 0 else 0
                
                st.markdown(f"""
                **📊 8月预算执行情况**
                - 实际需求: {format_currency(aug_actual)}
                - 预算目标: ¥{budget_target}万
                - 执行进度: {budget_progress:.1f}%
                """)
                
                st.progress(budget_progress / 100)
                
                # 高风险供应商
                risk_df = create_risk_warning_table(detail_df, 500000)
                if not risk_df.empty:
                    st.markdown("**🚨 高风险供应商 (>¥50万)**")
                    st.dataframe(
                        risk_df[['供应商名称', '欠料金额', '风险等级']].head(5),
                        hide_index=True,
                        use_container_width=True
                    )
    
    with tab2:
        st.markdown("### 🛒 采购执行清单")
        
        # 创建订单维度和供应商维度的子标签页
        order_tab, supplier_tab = st.tabs(["📋 按订单查看", "🏭 按供应商查看"])
        
        with order_tab:
            # 简化条件判断，只需要订单明细数据
            if '1_订单缺料明细' in data_dict:
                # 智能获取数据日期范围
                base_df = data_dict['1_订单缺料明细']
                try:
                    date_series = pd.to_datetime(base_df['客户交期'], errors='coerce')
                    data_min_date = date_series.min().date() if date_series.notna().any() else pd.to_datetime('2025-08-01').date()
                    data_max_date = date_series.max().date() if date_series.notna().any() else pd.to_datetime('2025-09-30').date()
                    
                    # 设置合理的默认筛选范围（最近两个月）
                    default_start = data_min_date
                    default_end = data_max_date
                except:
                    # 容错处理
                    default_start = pd.to_datetime('2025-08-01').date()  
                    default_end = pd.to_datetime('2025-09-30').date()
                
                # 时间区间筛选控件
                col1, col2, col3 = st.columns([1, 1, 1])
                with col1:
                    start_date = st.date_input("开始日期", value=default_start, 
                                             min_value=data_min_date if 'data_min_date' in locals() else None,
                                             max_value=data_max_date if 'data_max_date' in locals() else None)
                with col2:
                    end_date = st.date_input("结束日期", value=default_end,
                                           min_value=data_min_date if 'data_min_date' in locals() else None,
                                           max_value=data_max_date if 'data_max_date' in locals() else None)
                with col3:
                    month_filter = st.selectbox("月份快选", ["全部", "8月", "9月", "8月,9月"], key="order_month")
                
                # 基础数据展示
                detail_df = data_dict['1_订单缺料明细'].copy()
                
                # 时间筛选逻辑
                if month_filter != "全部":
                    # 兼容新旧月份字段
                    month_col = '涉及月份' if '涉及月份' in detail_df.columns else '月份'
                    if month_filter == "8月,9月":
                        # 查找包含8月,9月的记录
                        detail_df = detail_df[detail_df[month_col].str.contains('8月,9月', na=False)]
                    else:
                        # 精确匹配或包含匹配
                        detail_df = detail_df[
                            (detail_df[month_col] == month_filter) | 
                            (detail_df[month_col].str.contains(month_filter, na=False))
                        ]
                else:
                    # 按日期区间筛选
                    detail_df['客户交期_date'] = pd.to_datetime(detail_df['客户交期'], errors='coerce').dt.date
                    detail_df = detail_df[
                        (detail_df['客户交期_date'] >= start_date) & 
                        (detail_df['客户交期_date'] <= end_date)
                    ]
                    detail_df = detail_df.drop('客户交期_date', axis=1)
                
                # 按订单汇总
                summary_df = detail_df.groupby('生产订单号').agg({
                    '客户订单号': 'first',
                    '产品型号': 'first', 
                    '数量Pcs': 'first',  # 添加数量Pcs字段
                    '欠料金额(RMB)': 'sum',
                    '客户交期': 'first',
                    '目的地': 'first',
                    '订单金额(RMB)': 'first',  # 新增：预期回款
                    '每元投入回款': 'first',   # 新增：投入产出比
                    '数据完整性标记': 'first'  # 新增：数据状态
                }).reset_index()
                
                # 过滤大于1000的订单
                summary_df = summary_df[summary_df['欠料金额(RMB)'] >= 1000]
                
                # 初始化filtered_df（修复变量引用错误）
                filtered_df = summary_df.copy()
                
                # 显示筛选状态和导出功能
                col_export1, col_export2, col_export3 = st.columns([3, 0.8, 0.8])
                with col_export1:
                    filter_info = f"时间范围: {start_date} ~ {end_date}" if month_filter == "全部" else f"筛选: {month_filter}"
                    complete_count = len(filtered_df[filtered_df['数据完整性标记'] == '完整']) if '数据完整性标记' in filtered_df.columns else 0
                    # 修复：统计正确的总金额（按PSO去重）
                    total_investment = filtered_df['欠料金额(RMB)'].sum()
                    total_return = filtered_df[filtered_df['数据完整性标记'] == '完整']['订单金额(RMB)'].sum() if complete_count > 0 else 0
                with col_export2:
                    if len(summary_df) > 0:
                        # Windows Excel兼容版本（GBK编码）
                        output_gbk = io.BytesIO()
                        try:
                            csv_string = filtered_df.to_csv(index=False, encoding='gbk')
                            output_gbk.write(csv_string.encode('gbk'))
                        except:
                            # 降级到GB18030
                            csv_string = filtered_df.to_csv(index=False, encoding='gb18030')
                            output_gbk.write(csv_string.encode('gb18030'))
                        
                        output_gbk.seek(0)
                        st.download_button(
                            "📥 Excel版", 
                            data=output_gbk.getvalue(),
                            file_name=f"订单采购清单_{start_date}_{end_date}.csv",
                            mime="text/csv",
                            help="Windows Excel直接打开不乱码"
                        )
                
                with col_export3:
                    if len(summary_df) > 0:
                        # 通用UTF-8版本
                        output_utf8 = io.BytesIO()
                        csv_string = filtered_df.to_csv(index=False, encoding='utf-8-sig')
                        output_utf8.write(csv_string.encode('utf-8-sig'))
                        output_utf8.seek(0)
                        
                        st.download_button(
                            "📥 通用版", 
                            data=output_utf8.getvalue(),
                            file_name=f"订单采购清单_UTF8_{start_date}_{end_date}.csv",
                            mime="text/csv",
                            help="Mac/Linux或其他软件使用"
                        )
                
                # 预计算筛选统计（提升用户体验）
                complete_count = len(summary_df[summary_df['数据完整性标记'] == '完整'])
                high_return_count = len(summary_df[pd.to_numeric(summary_df['每元投入回款'], errors='coerce') > 2.0])
                urgent_date = pd.to_datetime('2025-09-10').date()
                summary_df_temp = summary_df.copy()
                summary_df_temp['客户交期_date'] = pd.to_datetime(summary_df_temp['客户交期'], errors='coerce').dt.date
                urgent_count = len(summary_df_temp[
                    (pd.to_numeric(summary_df_temp['每元投入回款'], errors='coerce') > 2.0) &
                    (summary_df_temp['客户交期_date'] <= urgent_date)
                ])
                
                # 添加投入产出比排序和高回报筛选（带数量提示）
                col_sort1, col_sort2, col_sort3, col_sort4, col_reset = st.columns([1.2, 1.2, 1.2, 1.2, 0.6])
                with col_sort1:
                    sort_option = st.selectbox("排序方式", ["投入产出比降序", "客户交期升序", "欠料金额降序"], key="order_sort")
                with col_sort2:
                    high_return_only = st.checkbox(f"高回报项目 ({high_return_count}条)", help="筛选投入产出比大于2的项目")
                with col_sort3:
                    urgent_projects = st.checkbox(f"紧急高回报 ({urgent_count}条)", help="投入产出比>2且客户交期≤9月10日")
                with col_sort4:
                    show_complete_only = st.checkbox(f"完整数据 ({complete_count}条)", help="只显示有订单金额的项目")
                with col_reset:
                    st.markdown("<br>", unsafe_allow_html=True)  # 对齐按钮位置
                    if st.button("🔄 重置", help="清除所有筛选条件"):
                        st.rerun()
                
                # 应用筛选和排序
                filtered_df = summary_df.copy()
                filter_applied = False
                
                # 数据完整性筛选
                if show_complete_only:
                    filtered_df = filtered_df[filtered_df['数据完整性标记'] == '完整']
                    filter_applied = True
                
                # 高回报筛选
                if high_return_only:
                    filtered_df = filtered_df[(pd.to_numeric(filtered_df['每元投入回款'], errors='coerce') > 2.0)]
                    filter_applied = True
                
                # 紧急高回报筛选
                if urgent_projects:
                    filtered_df['客户交期_date'] = pd.to_datetime(filtered_df['客户交期'], errors='coerce').dt.date
                    urgent_date = pd.to_datetime('2025-09-10').date()
                    filtered_df = filtered_df[
                        (pd.to_numeric(filtered_df['每元投入回款'], errors='coerce') > 2.0) &
                        (filtered_df['客户交期_date'] <= urgent_date)
                    ]
                    filtered_df = filtered_df.drop('客户交期_date', axis=1)
                    filter_applied = True
                
                # 智能回退：如果筛选后无数据，显示警告并回退
                if len(filtered_df) == 0 and filter_applied:
                    st.warning(f"⚠️ 当前筛选条件无匹配数据，显示全部 {len(summary_df)} 条基础数据")
                    filtered_df = summary_df.copy()
                
                # 实时显示筛选结果
                if filter_applied and len(filtered_df) > 0:
                    st.success(f"📊 筛选结果: {len(filtered_df)} 条记录 (从 {len(summary_df)} 条中筛选)")
                else:
                    st.info(f"📊 显示全部: {len(filtered_df)} 条记录")
                
                # 排序逻辑
                if sort_option == "投入产出比降序":
                    filtered_df = filtered_df.sort_values('每元投入回款', ascending=False, na_position='last')
                elif sort_option == "客户交期升序":
                    filtered_df = filtered_df.sort_values('客户交期', ascending=True)
                else:  # 欠料金额降序
                    filtered_df = filtered_df.sort_values('欠料金额(RMB)', ascending=False)
                
                # 初始化选中订单的session state（筛选重置时清空）
                if 'selected_orders' not in st.session_state or filter_applied:
                    st.session_state.selected_orders = set()
                
                # 多选ROI分析功能
                st.markdown("---")
                st.markdown("### 📊 多订单ROI分析")
                
                # 创建主要区域布局：订单表格 + ROI侧边栏
                main_col, sidebar_col = st.columns([3, 1])
                
                with main_col:
                    st.markdown("#### 📋 订单选择")
                    
                    # 全选功能
                    col_select_all, col_info = st.columns([1, 3])
                    with col_select_all:
                        select_all = st.checkbox("全选", key="select_all_orders")
                        if select_all:
                            st.session_state.selected_orders = set(filtered_df['生产订单号'].tolist())
                        elif not select_all and len(st.session_state.selected_orders) == len(filtered_df):
                            st.session_state.selected_orders = set()
                    
                    with col_info:
                        selected_count = len(st.session_state.selected_orders)
                        total_count = len(filtered_df)
                        st.markdown(f"**已选择**: {selected_count}/{total_count} 个订单")
                    
                    # 订单选择表格
                    selection_data = []
                    for _, row in filtered_df.iterrows():
                        order_no = row['生产订单号']
                        is_selected = order_no in st.session_state.selected_orders
                        
                        # 格式化显示数据
                        amount_str = format_currency(row['欠料金额(RMB)'])
                        return_str = format_currency(row.get('订单金额(RMB)', 0)) if pd.notna(row.get('订单金额(RMB)')) else "待补充"
                        roi_value = row.get('每元投入回款', 0)
                        if pd.notna(roi_value) and roi_value != float('inf'):
                            roi_str = f"{roi_value:.2f}倍"
                        else:
                            roi_str = "待补充"
                        
                        selection_data.append({
                            '选择': is_selected,
                            '生产订单号': order_no,
                            '客户订单号': str(row.get('客户订单号', '')),
                            '产品型号': str(row.get('产品型号', '')),
                            '客户交期': str(row.get('客户交期', ''))[:10],
                            '欠料金额': amount_str,
                            '预期回款': return_str,
                            '投入产出比': roi_str,
                            '完整性': row.get('数据完整性标记', '')
                        })
                    
                    # 显示选择表格（使用data_editor实现勾选功能）
                    if len(selection_data) > 0:
                        edited_df = st.data_editor(
                            pd.DataFrame(selection_data),
                            column_config={
                                "选择": st.column_config.CheckboxColumn(
                                    "选择",
                                    help="勾选要分析ROI的订单",
                                    default=False,
                                )
                            },
                            disabled=["生产订单号", "客户订单号", "产品型号", "客户交期", "欠料金额", "预期回款", "投入产出比", "完整性"],
                            hide_index=True,
                            use_container_width=True,
                            height=400
                        )
                        
                        # 更新选中状态
                        new_selected = set()
                        for idx, row in edited_df.iterrows():
                            if row['选择']:
                                new_selected.add(row['生产订单号'])
                        st.session_state.selected_orders = new_selected
                
                with sidebar_col:
                    # ROI分析侧边栏
                    st.markdown("#### 💰 ROI分析结果")
                    
                    selected_count = len(st.session_state.selected_orders)
                    
                    if selected_count == 0:
                        st.info("💡 请选择订单进行ROI分析")
                        st.markdown("""
                        **使用说明:**
                        1. 在左侧表格勾选订单
                        2. 系统自动计算总ROI
                        3. 显示投入回款详情
                        """)
                    else:
                        # 计算选中订单的ROI
                        selected_orders_df = filtered_df[filtered_df['生产订单号'].isin(st.session_state.selected_orders)]
                        
                        # 统计数据 - 按生产订单号计算正确的ROI
                        selected_order_summary = selected_orders_df.groupby('生产订单号').agg({
                            '订单金额(RMB)': 'first',  # 每个生产订单的金额
                            '欠料金额(RMB)': 'sum'      # 该订单的总欠料金额
                        }).reset_index()
                        
                        # 计算每个选中订单的ROI
                        selected_order_summary['订单ROI'] = np.where(
                            selected_order_summary['欠料金额(RMB)'] > 0,
                            selected_order_summary['订单金额(RMB)'] / selected_order_summary['欠料金额(RMB)'],
                            0
                        )
                        
                        total_shortage = selected_order_summary['欠料金额(RMB)'].sum()
                        total_order_amount = selected_order_summary['订单金额(RMB)'].sum() 
                        
                        # 统计有订单金额的订单数
                        orders_with_amount = len(selected_orders_df[
                            (selected_orders_df['订单金额(RMB)'].notna()) & 
                            (selected_orders_df['订单金额(RMB)'] > 0)
                        ])
                        orders_without_amount = selected_count - orders_with_amount
                        
                        # 显示基本统计
                        st.metric("选中订单数", f"{selected_count}个")
                        
                        # ROI计算和显示
                        if total_shortage == 0:
                            st.success("🎉 **立即生产**")
                            st.markdown("选中订单无欠料，可立即安排生产")
                        elif orders_without_amount == selected_count:
                            st.warning("⚠️ **无订单金额**") 
                            st.markdown(f"需要投入：{format_currency(total_shortage)}")
                            st.markdown("缺少订单金额数据，无法计算ROI")
                        elif orders_without_amount > 0:
                            # 部分有订单金额 - 使用加权平均ROI
                            if total_shortage > 0:
                                available_roi = (selected_order_summary['订单ROI'] * selected_order_summary['欠料金额(RMB)']).sum() / total_shortage
                            else:
                                available_roi = 0
                            st.metric("可计算ROI", f"{available_roi:.2f}倍")
                            
                            st.markdown("**💰 资金明细:**")
                            st.markdown(f"- 总投入：{format_currency(total_shortage)}")
                            st.markdown(f"- 可计算回款：{format_currency(total_order_amount)}")
                            st.markdown(f"- 有金额订单：{orders_with_amount}个")
                            st.markdown(f"- 缺少金额：{orders_without_amount}个")
                        else:
                            # 全部有订单金额 - 使用加权平均ROI
                            if total_shortage > 0:
                                total_roi = (selected_order_summary['订单ROI'] * selected_order_summary['欠料金额(RMB)']).sum() / total_shortage
                            else:
                                total_roi = 0
                            st.metric("总体ROI", f"{total_roi:.2f}倍", delta=f"vs目标2.8倍")
                            
                            # ROI颜色指示
                            if total_roi >= 2.8:
                                st.success("🟢 优秀回报项目")
                            elif total_roi >= 2.0:
                                st.warning("🟡 良好回报项目") 
                            else:
                                st.error("🔴 低回报项目")
                            
                            st.markdown("**💰 资金明细:**")
                            st.markdown(f"- 总投入：{format_currency(total_shortage)}")
                            st.markdown(f"- 总回款：{format_currency(total_order_amount)}")
                            st.markdown(f"- 净收益：{format_currency(total_order_amount - total_shortage)}")
                        
                        # 清除选择按钮
                        if st.button("🗑️ 清除选择", use_container_width=True):
                            st.session_state.selected_orders = set()
                            st.rerun()
                
                # 分隔线，分隔多选ROI功能和详细查看功能
                st.markdown("---")
                
                # 格式化显示（为后续展开式详情准备）
                display_df = filtered_df.copy()
                display_df['欠料金额(RMB)'] = display_df['欠料金额(RMB)'].apply(format_currency)
                
                # 格式化订单金额 - 添加USD和RMB显示
                if '订单金额(RMB)' in display_df.columns:
                    display_df['预期回款(RMB)'] = display_df['订单金额(RMB)'].apply(
                        lambda x: format_currency(x) if pd.notna(x) else '待补充'
                    )
                
                if '订单金额(USD)' in display_df.columns:
                    display_df['预期回款(USD)'] = display_df['订单金额(USD)'].apply(
                        lambda x: f"${x:,.2f}" if pd.notna(x) else '待补充'
                    )
                
                # 格式化投入产出比 - 直接重命名显示 
                if '每元投入回款' in display_df.columns:
                    display_df['投入产出比'] = display_df['每元投入回款'].apply(
                        lambda x: f"{x:.2f}" if pd.notna(x) and x != float('inf') else '待补充'
                    )
                
                # 选择显示列，根据可用字段动态调整
                display_columns = ['生产订单号', '客户订单号', '客户交期', '欠料金额(RMB)']
                
                # 添加月份字段
                if '涉及月份' in display_df.columns:
                    display_columns.append('涉及月份')
                elif '月份' in display_df.columns:
                    display_columns.append('月份')
                
                # 添加回款相关字段
                if '预期回款(USD)' in display_df.columns:
                    display_columns.append('预期回款(USD)')
                display_columns.extend(['预期回款(RMB)', '投入产出比', '数据完整性标记', '目的地'])
                
                available_columns = [col for col in display_columns if col in display_df.columns]
                
                # 新增：Expander展开式订单查看
                st.markdown("### 📋 订单明细查看（展开式）")
                
                # 增强的筛选和排序控件
                filter_col1, filter_col2, filter_col3, filter_col4 = st.columns([2, 1, 1, 1])
                with filter_col1:
                    search_order = st.text_input("🔍 快速搜索", placeholder="输入订单号、产品型号或客户订单号...", key="expander_search")
                with filter_col2:
                    expander_sort = st.selectbox("排序方式", ["投入产出比降序", "欠料金额降序", "客户交期升序", "生产订单号"], key="expander_sort")
                with filter_col3:
                    per_page = st.selectbox("每页显示", [20, 50, 100, "全部"], index=1, key="expander_per_page")
                with filter_col4:
                    roi_filter = st.selectbox("ROI筛选", ["全部", "高回报(>2.0)", "中等(1.0-2.0)", "待补充"], key="expander_roi_filter")
                
                # 应用搜索筛选
                expander_orders = filtered_df.copy()
                if search_order:
                    expander_orders = expander_orders[
                        (expander_orders['生产订单号'].astype(str).str.contains(search_order, case=False, na=False)) |
                        (expander_orders['产品型号'].astype(str).str.contains(search_order, case=False, na=False)) |
                        (expander_orders['客户订单号'].astype(str).str.contains(search_order, case=False, na=False))
                    ]
                
                # 应用ROI筛选
                if roi_filter == "高回报(>2.0)":
                    expander_orders = expander_orders[pd.to_numeric(expander_orders['每元投入回款'], errors='coerce') > 2.0]
                elif roi_filter == "中等(1.0-2.0)":
                    roi_values = pd.to_numeric(expander_orders['每元投入回款'], errors='coerce')
                    expander_orders = expander_orders[(roi_values >= 1.0) & (roi_values <= 2.0)]
                elif roi_filter == "待补充":
                    expander_orders = expander_orders[expander_orders['数据完整性标记'] != '完整']
                
                # 应用排序
                if expander_sort == "投入产出比降序":
                    expander_orders = expander_orders.sort_values('每元投入回款', ascending=False, na_position='last')
                elif expander_sort == "欠料金额降序":
                    expander_orders = expander_orders.sort_values('欠料金额(RMB)', ascending=False)
                elif expander_sort == "客户交期升序":
                    expander_orders = expander_orders.sort_values('客户交期', ascending=True)
                else:  # 生产订单号
                    expander_orders = expander_orders.sort_values('生产订单号')
                
                # 分页处理 - 重构版本
                total_orders = len(expander_orders)
                current_page = 1  # 默认值，确保始终被定义
                
                if per_page != "全部":
                    per_page = int(per_page)
                    total_pages = (total_orders + per_page - 1) // per_page
                    
                    if total_pages > 1:
                        # 显示分页控件
                        page_col1, page_col2, page_col3 = st.columns([1, 2, 1])
                        with page_col2:
                            current_page = st.select_slider(
                                f"页数 ({total_orders} 个订单)",
                                options=list(range(1, total_pages + 1)),
                                value=1,
                                key="expander_page"
                            )
                        
                        # 应用分页
                        start_idx = (current_page - 1) * per_page
                        end_idx = start_idx + per_page
                        expander_orders = expander_orders.iloc[start_idx:end_idx]
                    # else: 只有1页，使用默认值 current_page = 1
                # else: 显示全部，使用默认值 current_page = 1
                
                # 显示状态信息 - 改进版本
                if len(expander_orders) > 0:
                    if per_page != "全部":
                        per_page_int = int(per_page)
                        total_pages = (total_orders + per_page_int - 1) // per_page_int
                        if total_pages > 1:
                            status_info = f"📊 显示第 {current_page} 页（{len(expander_orders)} 个订单），总计 {total_orders} 个订单"
                        else:
                            status_info = f"📊 显示 {len(expander_orders)} 个订单（共1页）"
                    else:
                        status_info = f"📊 显示全部 {len(expander_orders)} 个订单"
                    st.info(status_info)
                else:
                    st.warning("🔍 没有找到匹配的订单，请调整筛选条件")
                
                # 使用expander展示每个订单（支持多个同时展开）
                for idx, order_row in expander_orders.iterrows():
                    # 格式化展示数据
                    amount_str = format_currency(order_row['欠料金额(RMB)'])
                    
                    # ROI显示
                    roi_value = order_row.get('每元投入回款', 0)
                    if pd.notna(roi_value) and roi_value != float('inf'):
                        roi_str = f"{roi_value:.1f}倍"
                        if roi_value > 2.0:
                            roi_color = "🟢"
                        elif roi_value >= 1.0:
                            roi_color = "🟡"
                        else:
                            roi_color = "🔴"
                    else:
                        roi_str = "待补充"
                        roi_color = "⚪"
                    
                    # 交期处理
                    delivery_date = order_row['客户交期']
                    if pd.notna(delivery_date):
                        if isinstance(delivery_date, pd.Timestamp):
                            delivery_str = delivery_date.strftime('%m/%d')
                            # 判断是否紧急（30天内）
                            days_left = (delivery_date - pd.Timestamp.now()).days
                            if days_left < 30:
                                delivery_color = "🔴"
                            elif days_left < 60:
                                delivery_color = "🟡"
                            else:
                                delivery_color = "🟢"
                        else:
                            delivery_str = str(delivery_date)[:10]
                            delivery_color = "⚪"
                    else:
                        delivery_str = "待定"
                        delivery_color = "⚪"
                    
                    # 数据完整性状态
                    if order_row.get('数据完整性标记') == '完整':
                        status_emoji = "✅"
                    else:
                        status_emoji = "⚠️"
                    
                    # 创建订单标题
                    order_title = f"{status_emoji} {order_row['生产订单号']} | {order_row['产品型号']} | 💰{amount_str} | {roi_color}ROI:{roi_str} | {delivery_color}交期:{delivery_str}"
                    
                    with st.expander(order_title):
                        # 获取该订单的详细物料信息
                        order_details = detail_df[detail_df['生产订单号'] == order_row['生产订单号']].copy()
                        
                        if len(order_details) > 0:
                            # 订单基本信息（2行布局）
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("客户订单", str(order_row['客户订单号']))
                            with col2:
                                order_qty = order_row.get('数量Pcs', 0)
                                if pd.notna(order_qty) and order_qty != 0:
                                    st.metric("数量Pcs", f"{int(order_qty):,}")
                                else:
                                    st.metric("数量Pcs", "待补充")
                            with col3:
                                full_delivery = order_row['客户交期']
                                if pd.notna(full_delivery) and isinstance(full_delivery, pd.Timestamp):
                                    delivery_full_str = full_delivery.strftime('%Y-%m-%d')
                                else:
                                    delivery_full_str = "待定"
                                st.metric("交期", delivery_full_str)
                            with col4:
                                st.metric("目的地", str(order_row.get('目的地', '未知')))
                            
                            # ROI详细信息（如果有完整数据）
                            if order_row.get('数据完整性标记') == '完整':
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    expected_return = order_row.get('订单金额(RMB)', 0)
                                    st.metric("预期回款", format_currency(expected_return) if pd.notna(expected_return) else "待定")
                                with col2:
                                    st.metric("投入金额", amount_str)
                                with col3:
                                    st.metric("🎯 投资回报率", roi_str)
                            
                            # 物料明细
                            st.markdown("**📦 物料缺料明细:**")
                            
                            # 按供应商分组展示
                            suppliers = order_details['主供应商名称'].unique()
                            # 过滤掉NaN值并转换为字符串
                            suppliers = [str(sup) for sup in suppliers if pd.notna(sup) and str(sup).strip() != '']
                            
                            if len(suppliers) > 1:
                                # 多供应商用标签页
                                supplier_tabs = st.tabs([f"{sup[:8]}..." if len(str(sup)) > 8 else str(sup) for sup in suppliers])
                                
                                for tab, supplier in zip(supplier_tabs, suppliers):
                                    with tab:
                                        supplier_items = order_details[order_details['主供应商名称'] == supplier]
                                        supplier_total = supplier_items['欠料金额(RMB)'].sum()
                                        
                                        st.markdown(f"**供应商: {supplier}** (¥{supplier_total:,.0f})")
                                        
                                        # 物料表格
                                        detail_display = supplier_items[['欠料物料编号', '欠料物料名称', '欠料数量', 'RMB单价', '欠料金额(RMB)']].copy()
                                        detail_display['RMB单价'] = detail_display['RMB单价'].apply(
                                            lambda x: f"¥{x:,.2f}" if pd.notna(x) and x > 0 else "待定"
                                        )
                                        detail_display['欠料金额(RMB)'] = detail_display['欠料金额(RMB)'].apply(
                                            lambda x: f"¥{x:,.0f}" if pd.notna(x) else "¥0"
                                        )
                                        
                                        st.dataframe(detail_display, use_container_width=True, hide_index=True, height=200)
                            else:
                                # 单供应商直接显示
                                detail_display = order_details[['欠料物料编号', '欠料物料名称', '欠料数量', 
                                                               '主供应商名称', 'RMB单价', '欠料金额(RMB)']].copy()
                                detail_display['RMB单价'] = detail_display['RMB单价'].apply(
                                    lambda x: f"¥{x:,.2f}" if pd.notna(x) and x > 0 else "待定"
                                )
                                detail_display['欠料金额(RMB)'] = detail_display['欠料金额(RMB)'].apply(
                                    lambda x: f"¥{x:,.0f}" if pd.notna(x) else "¥0"
                                )
                                
                                st.dataframe(detail_display, use_container_width=True, hide_index=True, height=250)
                            
                            # 汇总信息
                            total_items = len(order_details)
                            unique_suppliers = len(suppliers)
                            total_amount = order_details['欠料金额(RMB)'].sum()
                            
                            summary_text = f"📊 共 {total_items} 项物料，{unique_suppliers} 家供应商，总金额 {format_currency(total_amount)}"
                            if order_row.get('数据完整性标记') == '完整':
                                st.success(summary_text)
                            else:
                                st.info(summary_text + " | ⚠️ 缺少回款数据")
                        else:
                            st.warning(f"未找到订单 {order_row['生产订单号']} 的物料明细")
            else:
                st.warning("⚠️ 未找到订单明细数据，请检查数据文件是否正确加载")
        
        with supplier_tab:
            if '1_订单缺料明细' in data_dict:
                # 智能获取供应商数据日期范围（复用订单维度的数据范围）
                supplier_base_df = data_dict['1_订单缺料明细']
                try:
                    supplier_date_series = pd.to_datetime(supplier_base_df['客户交期'], errors='coerce')
                    supplier_min_date = supplier_date_series.min().date() if supplier_date_series.notna().any() else pd.to_datetime('2025-08-01').date()
                    supplier_max_date = supplier_date_series.max().date() if supplier_date_series.notna().any() else pd.to_datetime('2025-09-30').date()
                    supplier_default_start = supplier_min_date
                    supplier_default_end = supplier_max_date
                except:
                    supplier_default_start = pd.to_datetime('2025-08-01').date()  
                    supplier_default_end = pd.to_datetime('2025-09-30').date()
                
                # 供应商维度筛选控件
                col1, col2, col3, col4 = st.columns([1, 1, 1, 1])
                with col1:
                    supplier_start_date = st.date_input("开始日期", value=supplier_default_start, key="supplier_start")
                with col2:
                    supplier_end_date = st.date_input("结束日期", value=supplier_default_end, key="supplier_end")
                with col3:
                    supplier_month_filter = st.selectbox("月份快选", ["全部", "8月", "9月", "8月,9月"], key="supplier_month")
                with col4:
                    supplier_sort_by = st.selectbox("排序方式", ["采购金额", "数量Pcs", "供应商名称"], key="supplier_sort")
                
                # 获取供应商维度数据
                supplier_detail_df = data_dict['1_订单缺料明细'].copy()
                
                # 时间筛选逻辑
                if supplier_month_filter != "全部":
                    # 兼容新旧月份字段
                    month_col = '涉及月份' if '涉及月份' in supplier_detail_df.columns else '月份'
                    if supplier_month_filter == "8月,9月":
                        # 查找包含8月,9月的记录
                        supplier_detail_df = supplier_detail_df[supplier_detail_df[month_col].str.contains('8月,9月', na=False)]
                    else:
                        # 精确匹配或包含匹配
                        supplier_detail_df = supplier_detail_df[
                            (supplier_detail_df[month_col] == supplier_month_filter) | 
                            (supplier_detail_df[month_col].str.contains(supplier_month_filter, na=False))
                        ]
                else:
                    # 按日期区间筛选
                    supplier_detail_df['客户交期_date'] = pd.to_datetime(supplier_detail_df['客户交期'], errors='coerce').dt.date
                    supplier_detail_df = supplier_detail_df[
                        (supplier_detail_df['客户交期_date'] >= supplier_start_date) & 
                        (supplier_detail_df['客户交期_date'] <= supplier_end_date)
                    ]
                    supplier_detail_df = supplier_detail_df.drop('客户交期_date', axis=1)
                
                # 按供应商汇总
                supplier_summary = supplier_detail_df.groupby('主供应商名称').agg({
                    '生产订单号': lambda x: list(x.unique()),
                    '客户订单号': lambda x: list(x.unique()),
                    '欠料金额(RMB)': 'sum',
                    '月份': lambda x: '; '.join(x.unique()),
                    '客户交期': lambda x: list(x.unique())
                }).reset_index()
                
                # 过滤小额供应商
                supplier_summary = supplier_summary[supplier_summary['欠料金额(RMB)'] >= 1000]
                supplier_summary = supplier_summary.reset_index(drop=True)
                
                # 添加统计列
                supplier_summary['数量Pcs'] = supplier_summary['生产订单号'].apply(len)
                supplier_summary['客户数量'] = supplier_summary['客户订单号'].apply(len)
                
                # 排序
                if supplier_sort_by == "采购金额":
                    supplier_summary = supplier_summary.sort_values('欠料金额(RMB)', ascending=False)
                elif supplier_sort_by == "数量Pcs":
                    supplier_summary = supplier_summary.sort_values('数量Pcs', ascending=False)
                else:
                    supplier_summary = supplier_summary.sort_values('主供应商名称')
                
                # 供应商清单标题和导出
                col_export1, col_export2 = st.columns([3, 1])
                with col_export1:
                    st.markdown(f"**🏭 供应商清单 ({len(supplier_summary)}家供应商)**")
                with col_export2:
                    if len(supplier_summary) > 0:
                        # 创建导出用的简化数据
                        export_df = supplier_summary[['主供应商名称', '欠料金额(RMB)', '数量Pcs', '客户数量', '月份']].copy()
                        export_df['欠料金额(RMB)'] = export_df['欠料金额(RMB)'].apply(lambda x: f"{x:,.2f}")
                        
                        # 使用BytesIO和GBK编码确保Excel兼容性
                        output = io.BytesIO()
                        try:
                            # 优先使用GBK编码（Windows Excel最兼容）
                            csv_string = export_df.to_csv(index=False, encoding='gbk')
                            output.write(csv_string.encode('gbk'))
                        except UnicodeEncodeError:
                            # 如果GBK失败，使用GB18030（支持更多字符）
                            try:
                                csv_string = export_df.to_csv(index=False, encoding='gb18030')
                                output.write(csv_string.encode('gb18030'))
                            except:
                                # 最后回退到UTF-8-SIG
                                csv_string = export_df.to_csv(index=False, encoding='utf-8-sig')
                                output.write(csv_string.encode('utf-8-sig'))
                        
                        output.seek(0)
                        st.download_button(
                            "📥 导出CSV", 
                            data=output.getvalue(),
                            file_name=f"供应商采购清单_{supplier_start_date}_{supplier_end_date}.csv",
                            mime="text/csv"
                        )
                
                # 供应商展开列表
                for idx, supplier_row in supplier_summary.iterrows():
                    formatted_amount = format_currency(supplier_row['欠料金额(RMB)'])
                    supplier_title = f"🏭 {supplier_row['主供应商名称']} | 💰{formatted_amount} | 📋{supplier_row['数量Pcs']}个订单"
                    
                    with st.expander(supplier_title):
                        # 供应商基本信息
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("💰 采购总金额", formatted_amount)
                        with col2:
                            st.metric("📋 涉及订单", f"{supplier_row['数量Pcs']}个")
                        with col3:
                            st.metric("🎯 涉及客户", f"{supplier_row['客户数量']}个")
                        
                        # 该供应商的订单明细
                        st.markdown("**📋 相关订单明细:**")
                        supplier_orders = supplier_detail_df[
                            (supplier_detail_df['主供应商名称'] == supplier_row['主供应商名称']) &
                            (supplier_detail_df['欠料金额(RMB)'] >= 1000)
                        ][['生产订单号', '客户订单号', '欠料物料名称', '欠料金额(RMB)', '客户交期']].copy()
                        
                        supplier_orders['欠料金额(RMB)'] = supplier_orders['欠料金额(RMB)'].apply(format_currency)
                        supplier_orders = supplier_orders.astype(str)
                        st.dataframe(supplier_orders, hide_index=True, use_container_width=True)
                
                # 供应商统计信息
                supplier_total = supplier_summary['欠料金额(RMB)'].sum()
                st.markdown(f"""
                **📊 供应商统计:**
                - 总采购金额: {format_currency(supplier_total)}
                - 供应商数量: {len(supplier_summary)}家
                - 平均采购额: {format_currency(supplier_total/len(supplier_summary) if len(supplier_summary) > 0 else 0)}
                """)
    
    with tab3:
        st.markdown("### 📈 深度分析")
        
        if '3_供应商汇总' in data_dict and '4_多供应商选择表' in data_dict:
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 📊 供应商集中度分析")
                supplier_df = data_dict['3_供应商汇总'].copy()
                supplier_df = supplier_df.sort_values('采购总金额(RMB)', ascending=False).head(15)
                
                # 计算累计占比
                total_amount = supplier_df['采购总金额(RMB)'].sum()
                supplier_df['累计占比'] = supplier_df['采购总金额(RMB)'].cumsum() / total_amount * 100
                
                fig4 = px.bar(
                    supplier_df,
                    x='供应商名称',
                    y='采购总金额(RMB)',
                    title="供应商集中度分析 (帕累托图)",
                    color='采购总金额(RMB)',
                    color_continuous_scale='RdYlBu_r'
                )
                
                fig4.update_layout(
                    height=400, 
                    coloraxis_showscale=False,
                    xaxis=dict(tickangle=45)
                )
                st.plotly_chart(fig4, use_container_width=True)
                
                # 80/20分析
                top_20_percent = len(supplier_df) * 0.2
                top_suppliers_amount = supplier_df.head(int(top_20_percent))['采购总金额(RMB)'].sum()
                concentration_ratio = top_suppliers_amount / total_amount * 100
                
                st.markdown(f"""
                **🎯 集中度分析结果:**
                - 前20%供应商占比: {concentration_ratio:.1f}%
                - 供应链集中度: {'高' if concentration_ratio > 80 else '中等' if concentration_ratio > 60 else '分散'}
                """)
            
            with col2:
                st.markdown("#### 🔄 多供应商选择统计")
                multi_supplier_df = data_dict['4_多供应商选择表']
                
                # 统计多供应商物料
                multi_stats = multi_supplier_df.groupby('物料编号').agg({
                    '供应商名称': 'count',
                    'RMB单价': ['min', 'max', 'mean']
                }).reset_index()
                
                multi_stats.columns = ['物料编号', '供应商数量', '最低价', '最高价', '平均价']
                multi_stats['价格差异率'] = ((multi_stats['最高价'] - multi_stats['最低价']) / multi_stats['最低价'] * 100).fillna(0)
                
                # 选择价格差异最大的物料
                top_diff = multi_stats.nlargest(10, '价格差异率')
                
                fig5 = px.scatter(
                    top_diff,
                    x='供应商数量',
                    y='价格差异率',
                    size='平均价',
                    title="物料供应商选择机会分析",
                    hover_data=['物料编号'],
                    color='价格差异率',
                    color_continuous_scale='Viridis'
                )
                
                fig5.update_layout(height=400)
                st.plotly_chart(fig5, use_container_width=True)
                
                st.markdown(f"""
                **📈 多供应商优化机会:**
                - 可选择物料: {len(multi_stats)}个
                - 平均可选供应商: {multi_stats['供应商数量'].mean():.1f}家
                - 最大节省潜力: {top_diff['价格差异率'].max():.1f}%
                """)
    
    # 页脚信息
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #6C757D; font-size: 0.9em;">
        🌟 银图PMC智能分析平台 | 数据更新时间: 2025-08-25 17:40 | 
        <span style="color: #4A90E2;">Powered by Streamlit</span>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
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

# 页面配置
st.set_page_config(
    page_title="银图PMC智能分析平台",
    page_icon="🌟",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 在标题区域添加刷新按钮
header_col1, header_col2 = st.columns([4, 1])
with header_col1:
    st.markdown('<div class="main-title">🌟 银图PMC智能分析平台</div>', unsafe_allow_html=True)
with header_col2:
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
        # 尝试加载最新的带回款数据报告，如果不存在则使用原报告
        import glob
        report_files = glob.glob(r"D:\yingtu-PMC\精准供应商物料分析报告_含回款_*.xlsx")
        if report_files:
            # 使用最新的含回款报告
            latest_report = max(report_files)
            df = pd.read_excel(latest_report, sheet_name='Sheet1')
            # 转换为字典格式以保持兼容性
            excel_data = {'1_订单缺料明细': df}
        else:
            # 回退到原报告
            excel_data = pd.read_excel(
                r"D:\yingtu-PMC\精准供应商物料分析报告_20250825_1740.xlsx",
                sheet_name=None
            )
        
        # 修复数据类型问题，确保字符串列保持为字符串类型
        for sheet_name, df in excel_data.items():
            if not df.empty:
                # 确保订单号等字段保持为字符串
                string_columns = ['客户订单号', '生产订单号', '产品型号', '物料编码', 
                                '物料名称', '供应商编码', '供应商名称', '货币']
                for col in string_columns:
                    if col in df.columns:
                        df[col] = df[col].astype(str)
                
                # 确保数值列为数值类型
                numeric_columns = ['订单数量', '欠料金额(RMB)', '报价金额(RMB)', '供应商评分', '订单金额(RMB)', '每元投入回款']
                for col in numeric_columns:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                
                # 确保字符串列保持为字符串类型（新增数据完整性标记）
                string_columns.append('数据完整性标记')
        
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
    
    # 8月数据
    aug_amount = detail_df[detail_df['月份'] == '8月']['欠料金额(RMB)'].sum()
    
    # 投入产出分析（新增）
    total_order_amount = 0
    avg_return_ratio = 0
    high_return_count = 0
    
    if '订单金额(RMB)' in detail_df.columns and '每元投入回款' in detail_df.columns:
        # 计算总预期回款
        total_order_amount = detail_df['订单金额(RMB)'].sum()
        
        # 计算平均投入产出比（排除无穷值）
        valid_ratios = detail_df['每元投入回款'].replace([float('inf'), -float('inf')], None).dropna()
        if len(valid_ratios) > 0:
            avg_return_ratio = valid_ratios.mean()
        
        # 计算高回报项目数量（投入产出比>2）
        high_return_count = (valid_ratios > 2.0).sum()
    
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
    monthly_data = detail_df.groupby('月份')['欠料金额(RMB)'].sum().reset_index()
    
    fig = px.bar(
        monthly_data,
        x='月份',
        y='欠料金额(RMB)',
        title="📈 8月vs9月欠料金额对比",
        color='月份',
        color_discrete_map={'8月': '#4A90E2', '9月': '#7ED321'},
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
    
    supplier_ranking.columns = ['供应商', '欠料金额', '订单数量']
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
    
    # 加载数据
    data_dict = load_data()
    if not data_dict:
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
                aug_actual = detail_df[detail_df['月份'] == '8月']['欠料金额(RMB)'].sum()
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
                    month_filter = st.selectbox("月份快选", ["全部", "8月", "9月"], key="order_month")
                
                # 基础数据展示
                detail_df = data_dict['1_订单缺料明细'].copy()
                
                # 时间筛选逻辑
                if month_filter != "全部":
                    detail_df = detail_df[detail_df['月份'] == month_filter]
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
                    st.markdown(f"**📋 订单采购清单 ({len(filtered_df)}条记录)** | {filter_info} | 完整数据: {complete_count}条")
                    if complete_count > 0:
                        st.markdown(f"💰 **投入**: ¥{total_investment/10000:.1f}万 → **回收**: ¥{total_return/10000:.1f}万 | **整体回报**: {total_return/total_investment:.2f}倍")
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
                
                # 格式化显示
                display_df = filtered_df.copy()
                display_df['欠料金额(RMB)'] = display_df['欠料金额(RMB)'].apply(format_currency)
                
                # 格式化订单金额 - 直接重命名显示
                if '订单金额(RMB)' in display_df.columns:
                    display_df['预期回款(RMB)'] = display_df['订单金额(RMB)'].apply(
                        lambda x: format_currency(x) if pd.notna(x) else '待补充'
                    )
                
                # 格式化投入产出比 - 直接重命名显示 
                if '每元投入回款' in display_df.columns:
                    display_df['投入产出比'] = display_df['每元投入回款'].apply(
                        lambda x: f"{x:.2f}" if pd.notna(x) and x != float('inf') else '待补充'
                    )
                
                # 选择显示列（隐藏产品型号，突出关键信息）
                display_columns = ['生产订单号', '客户订单号', '客户交期', '欠料金额(RMB)', 
                                 '预期回款(RMB)', '投入产出比', '数据完整性标记', '目的地']
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
                
                # 分页处理
                total_orders = len(expander_orders)
                if per_page != "全部":
                    per_page = int(per_page)
                    total_pages = (total_orders + per_page - 1) // per_page
                    
                    if total_pages > 1:
                        page_col1, page_col2, page_col3 = st.columns([1, 2, 1])
                        with page_col2:
                            current_page = st.select_slider(
                                f"页数 ({total_orders} 个订单)",
                                options=list(range(1, total_pages + 1)),
                                value=1,
                                key="expander_page"
                            )
                        
                        start_idx = (current_page - 1) * per_page
                        end_idx = start_idx + per_page
                        expander_orders = expander_orders.iloc[start_idx:end_idx]
                else:
                    current_page = 1
                
                # 显示状态信息
                if len(expander_orders) > 0:
                    status_info = f"📊 显示 {len(expander_orders)} 个订单"
                    if per_page != "全部":
                        status_info += f"（第 {current_page} 页，共 {total_orders} 个）"
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
                                st.metric("订单数量", f"{order_row['订单数量']:,}")
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
                            if len(suppliers) > 1:
                                # 多供应商用标签页
                                supplier_tabs = st.tabs([f"{sup[:8]}..." if len(sup) > 8 else sup for sup in suppliers])
                                
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
                    supplier_month_filter = st.selectbox("月份快选", ["全部", "8月", "9月"], key="supplier_month")
                with col4:
                    supplier_sort_by = st.selectbox("排序方式", ["采购金额", "订单数量", "供应商名称"], key="supplier_sort")
                
                # 获取供应商维度数据
                supplier_detail_df = data_dict['1_订单缺料明细'].copy()
                
                # 时间筛选逻辑
                if supplier_month_filter != "全部":
                    supplier_detail_df = supplier_detail_df[supplier_detail_df['月份'] == supplier_month_filter]
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
                supplier_summary['订单数量'] = supplier_summary['生产订单号'].apply(len)
                supplier_summary['客户数量'] = supplier_summary['客户订单号'].apply(len)
                
                # 排序
                if supplier_sort_by == "采购金额":
                    supplier_summary = supplier_summary.sort_values('欠料金额(RMB)', ascending=False)
                elif supplier_sort_by == "订单数量":
                    supplier_summary = supplier_summary.sort_values('订单数量', ascending=False)
                else:
                    supplier_summary = supplier_summary.sort_values('主供应商名称')
                
                # 供应商清单标题和导出
                col_export1, col_export2 = st.columns([3, 1])
                with col_export1:
                    st.markdown(f"**🏭 供应商清单 ({len(supplier_summary)}家供应商)**")
                with col_export2:
                    if len(supplier_summary) > 0:
                        # 创建导出用的简化数据
                        export_df = supplier_summary[['主供应商名称', '欠料金额(RMB)', '订单数量', '客户数量', '月份']].copy()
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
                    supplier_title = f"🏭 {supplier_row['主供应商名称']} | 💰{formatted_amount} | 📋{supplier_row['订单数量']}个订单"
                    
                    with st.expander(supplier_title):
                        # 供应商基本信息
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("💰 采购总金额", formatted_amount)
                        with col2:
                            st.metric("📋 涉及订单", f"{supplier_row['订单数量']}个")
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
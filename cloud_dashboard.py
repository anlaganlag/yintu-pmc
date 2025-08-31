#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC 560订单清单分析 - 云端部署版
Streamlit Cloud优化版本，专为管理层决策设计
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 页面配置
st.set_page_config(
    page_title="银图PMC订单分析",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': "银图PMC智能分析平台 - 生产计划与物料控制系统"
    }
)

# 自定义样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
    }
    .kpi-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin: 0.5rem 0;
    }
    .kpi-card h3 {
        margin: 0;
        font-size: 1rem;
        opacity: 0.9;
    }
    .kpi-card h2 {
        margin: 0.5rem 0;
        font-size: 2rem;
        font-weight: bold;
    }
    .priority-high { color: #28a745; font-weight: bold; }
    .priority-medium { color: #ffc107; font-weight: bold; }
    .priority-low { color: #fd7e14; font-weight: bold; }
    .priority-pause { color: #dc3545; font-weight: bold; }
    .priority-immediate { color: #17a2b8; font-weight: bold; }
    .section-header {
        font-size: 1.8rem;
        font-weight: bold;
        color: #495057;
        margin: 2rem 0 1rem 0;
        border-left: 5px solid #1f77b4;
        padding-left: 1rem;
        background: linear-gradient(90deg, rgba(31, 119, 180, 0.1) 0%, rgba(255, 255, 255, 0) 100%);
        padding: 0.5rem 1rem;
        border-radius: 5px;
    }
    .metric-container {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin: 0.5rem 0;
    }
    .stTab {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px 10px 0 0;
    }
    .upload-section {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        padding: 2rem;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin: 2rem 0;
    }
    .success-message {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# 缓存装饰器用于数据加载
@st.cache_data(ttl=3600)  # 缓存1小时
def load_excel_data(file_content):
    """缓存Excel数据加载"""
    try:
        excel_data = pd.read_excel(file_content, sheet_name=None)
        return excel_data
    except Exception as e:
        st.error(f"数据加载失败: {e}")
        return None

class CloudDashboard:
    def __init__(self):
        self.data = None
        self.detailed_df = None
        self.summary_df = None
        self.priority_df = None
        self.supplier_df = None
        self.material_df = None
        
    def load_data(self, uploaded_file):
        """加载上传的Excel文件"""
        try:
            self.data = load_excel_data(uploaded_file)
            if self.data:
                self.detailed_df = self.data.get('按清单顺序详细分析')
                self.summary_df = self.data.get('按订单汇总(含ROI)')
                self.priority_df = self.data.get('ROI优先级排序')
                self.supplier_df = self.data.get('按供应商汇总')
                self.material_df = self.data.get('采购物料清单')
                return True
        except Exception as e:
            st.error(f"数据加载失败: {e}")
        return False
    
    def display_welcome_screen(self):
        """显示欢迎界面"""
        st.markdown('<div class="main-header">🎯 银图PMC智能分析平台</div>', unsafe_allow_html=True)
        
        st.markdown("""
        <div class="upload-section">
            <h2>🚀 欢迎使用银图PMC订单分析系统</h2>
            <p>专为生产管理层设计的智能决策支持平台</p>
            <p>📊 支持560个订单的全面分析 | 💰 ROI导向的优先级排序 | 🏭 155家供应商管理</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            ### 📋 核心功能
            - ✅ 订单优先级分析
            - ✅ ROI投资回报计算
            - ✅ 供应商采购优化
            - ✅ 物料需求规划
            """)
        
        with col2:
            st.markdown("""
            ### 🎯 管理决策
            - 📈 可立即生产订单识别
            - 💰 投入产出比优化
            - 🏭 供应商风险评估
            - ⏰ 生产排期建议
            """)
        
        with col3:
            st.markdown("""
            ### 📊 数据洞察
            - 🔍 560个订单全覆盖
            - 💎 998种物料分析
            - 🌐 155家供应商管理
            - 📱 移动端友好界面
            """)
    
    def display_kpi_dashboard(self):
        """显示KPI仪表板"""
        if not self.summary_df is not None:
            return
            
        st.markdown('<div class="section-header">📊 关键绩效指标</div>', unsafe_allow_html=True)
        
        # 计算关键指标
        total_orders = len(self.summary_df)
        immediate_orders = len(self.summary_df[self.summary_df['ROI'] >= 999999])
        high_priority = len(self.summary_df[(self.summary_df['ROI'] >= 5.0) & (self.summary_df['ROI'] < 999999)])
        total_shortage = self.summary_df['欠料金额(RMB)'].sum()
        total_order_value = self.summary_df['订单金额(RMB)'].sum()
        overall_roi = total_order_value / total_shortage if total_shortage > 0 else 0
        
        # 显示KPI卡片
        col1, col2, col3, col4, col5 = st.columns(5)
        
        kpi_data = [
            ("📋 订单总数", f"{total_orders:,}", "个生产订单"),
            ("🚀 可立即生产", f"{immediate_orders}", f"{immediate_orders/total_orders*100:.1f}% 无需投入"),
            ("⭐ 高优先级", f"{high_priority}", f"{high_priority/total_orders*100:.1f}% ROI≥5倍"),
            ("💰 总欠料金额", f"¥{total_shortage:,.0f}", "需要投入资金"),
            ("📈 整体ROI", f"{overall_roi:.1f}x", "投入产出比")
        ]
        
        cols = [col1, col2, col3, col4, col5]
        for i, (title, value, subtitle) in enumerate(kpi_data):
            with cols[i]:
                st.markdown(f"""
                <div class="kpi-card">
                    <h3>{title}</h3>
                    <h2>{value}</h2>
                    <small>{subtitle}</small>
                </div>
                """, unsafe_allow_html=True)
    
    def plot_priority_analysis(self):
        """优先级分析图表"""
        st.markdown('<div class="section-header">🎯 生产优先级分析</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # 优先级饼图
            priority_counts = self.summary_df['优先级等级'].value_counts()
            
            colors = ['#17a2b8', '#28a745', '#ffc107', '#fd7e14', '#dc3545']
            
            fig_pie = go.Figure(data=[go.Pie(
                labels=priority_counts.index,
                values=priority_counts.values,
                hole=0.5,
                marker_colors=colors[:len(priority_counts)]
            )])
            
            fig_pie.update_layout(
                title="订单优先级分布",
                font=dict(size=14),
                height=450,
                showlegend=True
            )
            
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            # 欠料金额 vs ROI 散点图
            analysis_data = self.summary_df[self.summary_df['ROI'] < 999999]
            
            color_map = {
                '高优先级': '#28a745',
                '中优先级': '#ffc107', 
                '低优先级': '#fd7e14',
                '暂缓生产': '#dc3545'
            }
            
            fig_scatter = px.scatter(
                analysis_data,
                x='欠料金额(RMB)',
                y='ROI',
                color='优先级等级',
                size='订单金额(RMB)',
                hover_name='生产订单号',
                color_discrete_map=color_map,
                title="投资回报风险分析"
            )
            
            # 添加参考线
            fig_scatter.add_hline(y=1.0, line_dash="dash", line_color="red", annotation_text="盈亏平衡线")
            fig_scatter.add_hline(y=5.0, line_dash="dash", line_color="green", annotation_text="高优先级线")
            
            fig_scatter.update_layout(height=450)
            st.plotly_chart(fig_scatter, use_container_width=True)
    
    def plot_top_analysis(self):
        """Top分析图表"""
        st.markdown('<div class="section-header">🔝 重点关注订单</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # TOP10欠料订单
            top10_shortage = self.summary_df.nlargest(10, '欠料金额(RMB)')
            
            fig_top_shortage = px.bar(
                top10_shortage,
                y='生产订单号',
                x='欠料金额(RMB)',
                color='优先级等级',
                title="欠料金额最高的10个订单",
                color_discrete_map={
                    '高优先级': '#28a745',
                    '中优先级': '#ffc107', 
                    '低优先级': '#fd7e14',
                    '暂缓生产': '#dc3545'
                }
            )
            
            fig_top_shortage.update_layout(
                height=400,
                yaxis={'categoryorder':'total ascending'}
            )
            st.plotly_chart(fig_top_shortage, use_container_width=True)
        
        with col2:
            # TOP10高ROI订单
            high_roi_orders = self.summary_df[self.summary_df['ROI'] < 999999].nlargest(10, 'ROI')
            
            fig_top_roi = px.bar(
                high_roi_orders,
                y='生产订单号',
                x='ROI',
                color='优先级等级',
                title="ROI最高的10个订单（排除无需投入）",
                color_discrete_map={
                    '高优先级': '#28a745',
                    '中优先级': '#ffc107'
                }
            )
            
            fig_top_roi.update_layout(
                height=400,
                yaxis={'categoryorder':'total ascending'}
            )
            st.plotly_chart(fig_top_roi, use_container_width=True)
    
    def display_supplier_analysis(self):
        """供应商分析"""
        st.markdown('<div class="section-header">🏭 供应商采购分析</div>', unsafe_allow_html=True)
        
        if self.supplier_df is not None and not self.supplier_df.empty:
            col1, col2 = st.columns(2)
            
            with col1:
                # TOP10供应商
                top10_suppliers = self.supplier_df.head(10)
                
                fig_suppliers = px.bar(
                    top10_suppliers,
                    y='主供应商名称',
                    x='欠料金额(RMB)',
                    title="TOP10供应商采购需求",
                    color='欠料金额(RMB)',
                    color_continuous_scale='Blues'
                )
                
                fig_suppliers.update_layout(
                    height=500,
                    yaxis={'categoryorder':'total ascending'}
                )
                st.plotly_chart(fig_suppliers, use_container_width=True)
            
            with col2:
                # 供应商分布分析
                supplier_stats = {
                    '总供应商数': len(self.supplier_df),
                    '总采购金额': f"¥{self.supplier_df['欠料金额(RMB)'].sum():,.2f}",
                    '平均采购金额': f"¥{self.supplier_df['欠料金额(RMB)'].mean():,.2f}",
                    '最大单笔采购': f"¥{self.supplier_df['欠料金额(RMB)'].max():,.2f}"
                }
                
                st.markdown("### 📈 供应商统计概览")
                for key, value in supplier_stats.items():
                    st.markdown(f"""
                    <div class="metric-container">
                        <h4>{key}</h4>
                        <h2 style="color: #1f77b4;">{value}</h2>
                    </div>
                    """, unsafe_allow_html=True)
    
    def display_production_recommendations(self):
        """生产建议"""
        st.markdown('<div class="section-header">🎯 生产排期建议</div>', unsafe_allow_html=True)
        
        if self.priority_df is not None:
            # 筛选器
            col1, col2, col3 = st.columns(3)
            
            with col1:
                priority_filter = st.multiselect(
                    "选择优先级",
                    options=self.priority_df['优先级等级'].unique(),
                    default=['无需投入-立即生产', '高优先级']
                )
            
            with col2:
                max_orders = st.slider("显示订单数量", 10, 100, 20)
            
            with col3:
                sort_by = st.selectbox("排序方式", ["建议生产顺序", "ROI", "欠料金额(RMB)"])
            
            # 筛选和排序数据
            if priority_filter:
                filtered_data = self.priority_df[self.priority_df['优先级等级'].isin(priority_filter)]
                
                if sort_by != "建议生产顺序":
                    filtered_data = filtered_data.sort_values(sort_by, ascending=False)
                
                display_data = filtered_data.head(max_orders)
                
                # 格式化显示
                display_columns = ['建议生产顺序', '生产订单号', '产品型号', '数量Pcs', 
                                 '欠料金额(RMB)', 'ROI', '优先级等级']
                
                # 创建显示用DataFrame
                show_df = display_data[display_columns].copy()
                show_df['ROI'] = show_df['ROI'].apply(lambda x: '无需投入' if x >= 999999 else f'{x:.2f}倍')
                show_df['欠料金额(RMB)'] = show_df['欠料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
                
                st.dataframe(
                    show_df,
                    use_container_width=True,
                    height=400
                )
                
                # 显示筛选统计
                st.success(f"📊 当前显示 {len(display_data)} 个订单，筛选自 {len(filtered_data)} 个匹配订单")

def main():
    """主函数"""
    dashboard = CloudDashboard()
    
    # 侧边栏
    with st.sidebar:
        st.image("https://via.placeholder.com/200x80/1f77b4/white?text=YINTU+PMC", width=200)
        
        st.markdown("### 📁 数据上传")
        uploaded_file = st.file_uploader(
            "选择分析报告Excel文件",
            type=['xlsx'],
            help="上传 560订单清单欠料分析报告_xxx.xlsx"
        )
        
        if uploaded_file is not None:
            with st.spinner('正在加载数据...'):
                if dashboard.load_data(uploaded_file):
                    st.markdown("""
                    <div class="success-message">
                        <h4>✅ 数据加载成功！</h4>
                        <p>560个订单数据已就绪</p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.error("❌ 数据加载失败")
        
        # 使用说明
        with st.expander("📖 使用说明"):
            st.markdown("""
            **快速开始：**
            1. 上传Excel分析报告
            2. 查看各个标签页的分析
            3. 使用筛选器定制视图
            
            **主要功能：**
            - 🎯 管理层KPI概览
            - 📊 优先级分析图表
            - 🏭 供应商采购分析
            - 📋 生产排期建议
            """)
    
    # 主内容区域
    if dashboard.data is None:
        dashboard.display_welcome_screen()
    else:
        # 显示主要仪表板
        dashboard.display_kpi_dashboard()
        
        # 标签页
        tab1, tab2, tab3, tab4 = st.tabs([
            "🎯 优先级分析", 
            "🔝 重点订单", 
            "🏭 供应商分析", 
            "📋 生产建议"
        ])
        
        with tab1:
            dashboard.plot_priority_analysis()
        
        with tab2:
            dashboard.plot_top_analysis()
        
        with tab3:
            dashboard.display_supplier_analysis()
        
        with tab4:
            dashboard.display_production_recommendations()

if __name__ == "__main__":
    main()
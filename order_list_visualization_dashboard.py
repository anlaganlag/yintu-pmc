#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
560订单清单可视化仪表板
基于Streamlit的交互式数据可视化平台
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.figure_factory as ff
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 页面配置
st.set_page_config(
    page_title="560订单清单欠料分析仪表板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .kpi-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        border: 2px solid #1f77b4;
        margin: 0.5rem 0;
    }
    .priority-high { color: #28a745; font-weight: bold; }
    .priority-medium { color: #ffc107; font-weight: bold; }
    .priority-low { color: #fd7e14; font-weight: bold; }
    .priority-pause { color: #dc3545; font-weight: bold; }
    .priority-immediate { color: #17a2b8; font-weight: bold; }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #495057;
        margin: 1rem 0;
        border-bottom: 2px solid #1f77b4;
        padding-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

class OrderListVisualizationDashboard:
    def __init__(self):
        self.data_loaded = False
        self.detailed_df = None
        self.summary_df = None
        self.priority_df = None
        self.supplier_df = None
        self.material_df = None
        
    def load_data(self, file_path):
        """加载Excel分析报告数据"""
        try:
            # 加载所有工作表
            excel_data = pd.read_excel(file_path, sheet_name=None)
            
            self.detailed_df = excel_data['按清单顺序详细分析']
            self.summary_df = excel_data['按订单汇总(含ROI)']
            self.priority_df = excel_data['ROI优先级排序']
            self.supplier_df = excel_data['按供应商汇总']
            self.material_df = excel_data['采购物料清单']
            
            self.data_loaded = True
            return True
        except Exception as e:
            st.error(f"数据加载失败: {e}")
            return False
    
    def display_kpi_cards(self):
        """显示关键KPI卡片"""
        if not self.data_loaded:
            return
            
        # 计算关键指标
        total_orders = len(self.summary_df)
        immediate_orders = len(self.summary_df[self.summary_df['ROI'] >= 999999])
        high_priority = len(self.summary_df[(self.summary_df['ROI'] >= 5.0) & (self.summary_df['ROI'] < 999999)])
        total_shortage = self.summary_df['欠料金额(RMB)'].sum()
        total_order_value = self.summary_df['订单金额(RMB)'].sum()
        overall_roi = total_order_value / total_shortage if total_shortage > 0 else 0
        
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.markdown(f"""
            <div class="kpi-card">
                <h3>📋 订单总数</h3>
                <h2>{total_orders:,}</h2>
                <small>个生产订单</small>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            immediate_pct = immediate_orders / total_orders * 100
            st.markdown(f"""
            <div class="kpi-card">
                <h3>🚀 可立即生产</h3>
                <h2 class="priority-immediate">{immediate_orders}</h2>
                <small>{immediate_pct:.1f}% 无需投入</small>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            high_pct = high_priority / total_orders * 100
            st.markdown(f"""
            <div class="kpi-card">
                <h3>⭐ 高优先级</h3>
                <h2 class="priority-high">{high_priority}</h2>
                <small>{high_pct:.1f}% ROI≥5倍</small>
            </div>
            """, unsafe_allow_html=True)
        
        with col4:
            st.markdown(f"""
            <div class="kpi-card">
                <h3>💰 总欠料金额</h3>
                <h2>¥{total_shortage:,.0f}</h2>
                <small>需要投入资金</small>
            </div>
            """, unsafe_allow_html=True)
        
        with col5:
            st.markdown(f"""
            <div class="kpi-card">
                <h3>📈 整体ROI</h3>
                <h2 class="priority-high">{overall_roi:.1f}x</h2>
                <small>投入产出比</small>
            </div>
            """, unsafe_allow_html=True)
    
    def plot_priority_distribution(self):
        """绘制优先级分布图"""
        if not self.data_loaded:
            return
        
        st.markdown('<div class="section-header">📊 订单优先级分布</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            # 优先级饼图
            priority_counts = self.summary_df['优先级等级'].value_counts()
            
            colors = {
                '无需投入-立即生产': '#17a2b8',
                '高优先级': '#28a745', 
                '中优先级': '#ffc107',
                '低优先级': '#fd7e14',
                '暂缓生产': '#dc3545'
            }
            
            fig_pie = go.Figure(data=[go.Pie(
                labels=priority_counts.index,
                values=priority_counts.values,
                hole=0.4,
                marker_colors=[colors.get(label, '#gray') for label in priority_counts.index]
            )])
            
            fig_pie.update_layout(
                title="订单优先级分布",
                font=dict(size=12),
                height=400
            )
            
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            # ROI分布直方图
            roi_data = self.summary_df[self.summary_df['ROI'] < 999999]['ROI']  # 排除无需投入的订单
            
            fig_hist = px.histogram(
                x=roi_data,
                nbins=20,
                title="ROI分布直方图（排除无需投入订单）",
                labels={'x': 'ROI倍数', 'y': '订单数量'}
            )
            
            fig_hist.add_vline(x=1.0, line_dash="dash", line_color="red", 
                              annotation_text="盈亏平衡线")
            fig_hist.add_vline(x=2.0, line_dash="dash", line_color="orange", 
                              annotation_text="中优先级线")
            fig_hist.add_vline(x=5.0, line_dash="dash", line_color="green", 
                              annotation_text="高优先级线")
            
            fig_hist.update_layout(height=400)
            st.plotly_chart(fig_hist, use_container_width=True)
    
    def plot_top_shortage_orders(self):
        """显示欠料金额最高的订单"""
        st.markdown('<div class="section-header">🔝 欠料金额TOP20订单</div>', unsafe_allow_html=True)
        
        top20_shortage = self.summary_df.nlargest(20, '欠料金额(RMB)')
        
        # 创建颜色映射
        color_map = {
            '无需投入-立即生产': '#17a2b8',
            '高优先级': '#28a745',
            '中优先级': '#ffc107', 
            '低优先级': '#fd7e14',
            '暂缓生产': '#dc3545'
        }
        
        fig = px.bar(
            top20_shortage,
            x='欠料金额(RMB)',
            y='生产订单号',
            color='优先级等级',
            color_discrete_map=color_map,
            title="欠料金额最高的20个订单",
            labels={'欠料金额(RMB)': '欠料金额 (RMB)', '生产订单号': '生产订单号'},
            text='产品型号'
        )
        
        fig.update_layout(height=600, yaxis={'categoryorder':'total ascending'})
        fig.update_traces(textposition='inside', textfont_size=10)
        
        st.plotly_chart(fig, use_container_width=True)
    
    def plot_roi_heatmap(self):
        """ROI热力图（按订单清单顺序）"""
        st.markdown('<div class="section-header">🌡️ ROI热力图（按清单顺序）</div>', unsafe_allow_html=True)
        
        # 准备热力图数据
        summary_sorted = self.summary_df.sort_values('清单序号')
        
        # 将ROI转换为热力图格式（10x56的网格）
        roi_values = summary_sorted['ROI'].values
        roi_matrix = []
        
        # 处理无需投入的订单（设为特殊值）
        roi_display = []
        for roi in roi_values:
            if roi >= 999999:
                roi_display.append(100)  # 用100表示无需投入
            else:
                roi_display.append(min(roi, 20))  # 限制最大值为20，便于可视化
        
        # 重塑为矩阵
        rows = 10
        cols = len(roi_display) // rows + (1 if len(roi_display) % rows else 0)
        
        # 填充到完整矩阵
        roi_padded = roi_display + [0] * (rows * cols - len(roi_display))
        roi_matrix = np.array(roi_padded).reshape(rows, cols)
        
        fig = px.imshow(
            roi_matrix,
            aspect="auto",
            color_continuous_scale="RdYlGn",
            title="560个订单ROI热力图（按清单顺序排列）"
        )
        
        fig.update_layout(
            xaxis_title="订单序列（每列代表连续的订单）",
            yaxis_title="行序列",
            height=400
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # 添加说明
        st.info("🔍 热力图说明：绿色=高ROI（优先生产），红色=低ROI（暂缓生产），最亮绿=无需投入")
    
    def plot_supplier_analysis(self):
        """供应商分析图表"""
        st.markdown('<div class="section-header">🏭 供应商采购分析</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            # TOP10供应商采购金额
            top10_suppliers = self.supplier_df.head(10)
            
            fig_supplier = px.bar(
                top10_suppliers,
                x='欠料金额(RMB)',
                y='主供应商名称',
                title="TOP10供应商采购需求",
                labels={'欠料金额(RMB)': '采购金额 (RMB)', '主供应商名称': '供应商'},
                text='涉及物料种类'
            )
            
            fig_supplier.update_layout(
                height=500,
                yaxis={'categoryorder':'total ascending'}
            )
            fig_supplier.update_traces(textposition='inside')
            
            st.plotly_chart(fig_supplier, use_container_width=True)
        
        with col2:
            # 供应商物料种类分布
            fig_scatter = px.scatter(
                self.supplier_df.head(20),
                x='涉及物料种类',
                y='影响订单数',
                size='欠料金额(RMB)',
                hover_name='主供应商名称',
                title="TOP20供应商影响力分析",
                labels={'涉及物料种类': '物料种类数', '影响订单数': '影响订单数'}
            )
            
            fig_scatter.update_layout(height=500)
            st.plotly_chart(fig_scatter, use_container_width=True)
    
    def display_priority_order_table(self):
        """显示优先级排序表"""
        st.markdown('<div class="section-header">📋 生产优先级排序表</div>', unsafe_allow_html=True)
        
        # 添加筛选器
        col1, col2, col3 = st.columns(3)
        
        with col1:
            priority_filter = st.multiselect(
                "选择优先级等级",
                options=self.priority_df['优先级等级'].unique(),
                default=self.priority_df['优先级等级'].unique()
            )
        
        with col2:
            min_roi = st.number_input("最小ROI", min_value=0.0, value=0.0, step=0.1)
        
        with col3:
            max_shortage = st.number_input("最大欠料金额", min_value=0, value=int(self.priority_df['欠料金额(RMB)'].max()), step=10000)
        
        # 筛选数据
        filtered_data = self.priority_df[
            (self.priority_df['优先级等级'].isin(priority_filter)) &
            (self.priority_df['ROI'] >= min_roi) &
            (self.priority_df['欠料金额(RMB)'] <= max_shortage)
        ]
        
        # 显示筛选后的表格
        if not filtered_data.empty:
            # 格式化显示
            display_df = filtered_data[['建议生产顺序', '生产订单号', '产品型号', '数量Pcs', 
                                      '欠料金额(RMB)', 'ROI', '优先级等级', '欠料物料种类']].copy()
            
            # 格式化ROI列
            display_df['ROI'] = display_df['ROI'].apply(lambda x: '无需投入' if x >= 999999 else f'{x:.2f}倍')
            
            # 格式化金额列
            display_df['欠料金额(RMB)'] = display_df['欠料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
            
            st.dataframe(
                display_df,
                use_container_width=True,
                height=400
            )
            
            st.info(f"📊 共筛选出 {len(filtered_data)} 个订单")
        else:
            st.warning("没有符合筛选条件的订单")
    
    def display_material_procurement_list(self):
        """显示物料采购清单"""
        st.markdown('<div class="section-header">📦 物料采购清单</div>', unsafe_allow_html=True)
        
        # 添加搜索功能
        search_term = st.text_input("🔍 搜索物料编号或名称")
        
        # 筛选数据
        if search_term:
            filtered_materials = self.material_df[
                (self.material_df['欠料物料编号'].str.contains(search_term, case=False, na=False)) |
                (self.material_df['欠料物料名称'].str.contains(search_term, case=False, na=False))
            ]
        else:
            filtered_materials = self.material_df.head(50)  # 默认显示前50个
        
        if not filtered_materials.empty:
            # 格式化显示
            display_materials = filtered_materials[['欠料物料编号', '欠料物料名称', '欠料数量',
                                                  '主供应商名称', '供应商单价(原币)', '供应商币种',
                                                  '欠料金额(RMB)', '相关订单(前5个)']].copy()
            
            # 格式化金额
            display_materials['欠料金额(RMB)'] = display_materials['欠料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
            
            st.dataframe(
                display_materials,
                use_container_width=True,
                height=400
            )
            
            # 显示采购汇总
            total_materials = len(filtered_materials)
            total_amount = filtered_materials['欠料金额(RMB)'].sum()
            unique_suppliers = filtered_materials['主供应商名称'].nunique()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("物料种类", f"{total_materials}种")
            with col2:
                st.metric("采购金额", f"¥{total_amount:,.2f}")
            with col3:
                st.metric("涉及供应商", f"{unique_suppliers}家")
        else:
            st.warning("未找到匹配的物料")

def main():
    """主函数"""
    st.markdown('<div class="main-header">🎯 560订单清单欠料分析仪表板</div>', unsafe_allow_html=True)
    
    # 创建仪表板实例
    dashboard = OrderListVisualizationDashboard()
    
    # 侧边栏文件上传
    st.sidebar.header("📁 数据加载")
    uploaded_file = st.sidebar.file_uploader(
        "上传分析报告Excel文件",
        type=['xlsx'],
        help="请上传 '560订单清单欠料分析报告_xxx.xlsx' 文件"
    )
    
    if uploaded_file is not None:
        if dashboard.load_data(uploaded_file):
            st.sidebar.success("✅ 数据加载成功！")
            
            # 主要内容区域
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "📊 管理层概览", 
                "🎯 优先级分析", 
                "🏭 供应商分析", 
                "📋 生产排序表", 
                "📦 采购清单"
            ])
            
            with tab1:
                st.subheader("管理层关键指标概览")
                dashboard.display_kpi_cards()
                st.markdown("---")
                dashboard.plot_priority_distribution()
                st.markdown("---")
                dashboard.plot_roi_heatmap()
            
            with tab2:
                st.subheader("订单优先级分析")
                dashboard.plot_top_shortage_orders()
                st.markdown("---")
                
                # ROI vs 欠料金额散点图
                st.markdown('<div class="section-header">💰 ROI vs 欠料金额分析</div>', unsafe_allow_html=True)
                
                color_map = {
                    '无需投入-立即生产': '#17a2b8',
                    '高优先级': '#28a745',
                    '中优先级': '#ffc107', 
                    '低优先级': '#fd7e14',
                    '暂缓生产': '#dc3545'
                }
                
                # 过滤掉无需投入的订单进行散点图分析
                analysis_data = dashboard.summary_df[dashboard.summary_df['ROI'] < 999999]
                
                fig_scatter = px.scatter(
                    analysis_data,
                    x='欠料金额(RMB)',
                    y='ROI',
                    color='优先级等级',
                    size='订单金额(RMB)',
                    hover_name='生产订单号',
                    hover_data=['产品型号', '欠料物料种类'],
                    color_discrete_map=color_map,
                    title="订单投资回报分析（气泡大小=订单金额）",
                    labels={'欠料金额(RMB)': '欠料金额 (RMB)', 'ROI': 'ROI (倍数)'}
                )
                
                # 添加ROI参考线
                fig_scatter.add_hline(y=1.0, line_dash="dash", line_color="red", 
                                     annotation_text="盈亏平衡线")
                fig_scatter.add_hline(y=2.0, line_dash="dash", line_color="orange", 
                                     annotation_text="中优先级线")
                fig_scatter.add_hline(y=5.0, line_dash="dash", line_color="green", 
                                     annotation_text="高优先级线")
                
                fig_scatter.update_layout(height=600)
                st.plotly_chart(fig_scatter, use_container_width=True)
            
            with tab3:
                st.subheader("供应商采购需求分析")
                dashboard.plot_supplier_analysis()
            
            with tab4:
                st.subheader("生产优先级排序")
                dashboard.display_priority_order_table()
            
            with tab5:
                st.subheader("物料采购清单")
                dashboard.display_material_procurement_list()
            
        else:
            st.sidebar.error("❌ 数据加载失败，请检查文件格式")
    else:
        st.info("👆 请在侧边栏上传560订单清单分析报告Excel文件开始分析")
        
        # 显示使用说明
        with st.expander("📖 使用说明"):
            st.markdown("""
            ### 如何使用此仪表板：
            
            1. **上传数据文件**：在左侧侧边栏上传 `560订单清单欠料分析报告_xxx.xlsx` 文件
            
            2. **管理层概览**：
               - 查看关键KPI指标
               - 订单优先级分布分析
               - ROI热力图（按订单清单顺序）
            
            3. **优先级分析**：
               - 欠料金额最高的订单识别
               - ROI vs 欠料金额散点图分析
               - 投资回报风险评估
            
            4. **供应商分析**：
               - TOP10供应商采购需求
               - 供应商影响力分析（物料种类vs影响订单数）
            
            5. **生产排序表**：
               - 按ROI排序的生产建议
               - 支持多维度筛选
               - 实时筛选结果统计
            
            6. **采购清单**：
               - 详细的998种物料采购清单
               - 支持物料搜索功能
               - 采购汇总统计
            """)

if __name__ == "__main__":
    main()
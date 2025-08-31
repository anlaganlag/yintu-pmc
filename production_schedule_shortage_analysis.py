#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生产排期欠料分析系统
基于生产排期表分析对应的欠料情况、供应商信息和金额
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class ProductionScheduleShortageAnalyzer:
    def __init__(self):
        self.schedule_df = None          # 生产排期数据
        self.shortage_analysis_df = None # 综合分析报表数据
        self.production_orders = []      # 提取的生产订单列表
        self.final_result = None         # 最终分析结果
        
    def load_production_schedule(self, schedule_file='250-08-28沙井成品生產排期.xlsx'):
        """加载生产排期表"""
        print("=== 🔄 加载生产排期表 ===")
        
        try:
            # 读取生产排期Excel文件
            df = pd.read_excel(schedule_file)
            print(f"   ✅ 成功读取文件，共{len(df)}行数据")
            
            # 提取生产订单信息
            # 根据前面的分析，PSO订单号在Unnamed: 1列中，从第4行开始
            production_orders = []
            
            for idx, row in df.iterrows():
                pso_value = str(row.get('Unnamed: 1', '')).strip()
                if pso_value.startswith('PSO') and len(pso_value) > 5:
                    # 提取订单相关信息
                    order_info = {
                        '生产单号': pso_value,
                        '本厂型号': str(row.get('Unnamed: 2', '')).strip(),
                        '客户型号': str(row.get('Unnamed: 3', '')).strip(),
                        '订单数量': pd.to_numeric(row.get('Unnamed: 4', 0), errors='coerce'),
                        'OTS期': row.get('Unnamed: 5', ''),
                        '生产线': str(row.get('TO:塑膠部、采購部、五金部、PIE/EN、QA/IQC、成品一、二部、三部、江門、貨倉部、絲印部、船務部、工模部。', '')).strip(),
                        '排期行号': idx + 1  # Excel行号（从1开始）
                    }
                    
                    # 提取备注信息（通常在后面的列中）
                    remarks = []
                    for col in df.columns:
                        cell_value = str(row.get(col, '')).strip()
                        if cell_value and col.startswith('Unnamed:') and len(cell_value) > 10:
                            if any(keyword in cell_value for keyword in ['已批GB', '齐料', '货好', '验货', '预计']):
                                remarks.append(cell_value)
                    
                    order_info['备注'] = '; '.join(remarks[:2])  # 最多取前2个重要备注
                    production_orders.append(order_info)
            
            self.schedule_df = pd.DataFrame(production_orders)
            self.production_orders = [order['生产单号'] for order in production_orders]
            
            print(f"   ✅ 提取生产订单: {len(self.production_orders)}个")
            print(f"   📊 订单号范围: {self.production_orders[0]} 至 {self.production_orders[-1]}")
            
            # 显示前几个订单信息
            print("   🔍 前5个生产订单:")
            for i, order in enumerate(production_orders[:5]):
                print(f"      {i+1}. {order['生产单号']}: {order['客户型号']} ({order['订单数量']}pcs)")
            
            return True
            
        except Exception as e:
            print(f"   ❌ 生产排期表加载失败: {e}")
            return False
    
    def load_shortage_analysis(self, analysis_file='银图PMC综合物料分析报告_20250831_170129.xlsx'):
        """加载已生成的综合分析报表"""
        print("=== 🔄 加载综合分析报表 ===")
        
        try:
            # 读取综合分析报表的明细sheet
            self.shortage_analysis_df = pd.read_excel(analysis_file, sheet_name='综合物料分析明细')
            print(f"   ✅ 成功读取综合分析报表，共{len(self.shortage_analysis_df)}条记录")
            
            # 统计报表中的订单数量
            total_orders_in_analysis = self.shortage_analysis_df['生产订单号'].nunique()
            print(f"   📊 分析报表包含{total_orders_in_analysis}个唯一订单")
            
            return True
            
        except Exception as e:
            print(f"   ❌ 综合分析报表加载失败: {e}")
            return False
    
    def analyze_schedule_shortage(self):
        """分析生产排期订单的欠料情况"""
        print("=== 🎯 分析生产排期订单欠料情况 ===")
        
        if self.schedule_df is None or self.shortage_analysis_df is None:
            print("❌ 数据未准备完成")
            return False
        
        # 匹配生产排期订单与欠料分析
        matched_data = []
        unmatched_orders = []
        
        print(f"1. 匹配{len(self.production_orders)}个生产排期订单...")
        
        for pso in self.production_orders:
            # 在分析报表中查找对应订单
            order_records = self.shortage_analysis_df[
                self.shortage_analysis_df['生产订单号'] == pso
            ].copy()
            
            if not order_records.empty:
                # 添加排期信息
                schedule_info = self.schedule_df[self.schedule_df['生产单号'] == pso].iloc[0]
                for _, record in order_records.iterrows():
                    enhanced_record = record.to_dict()
                    enhanced_record.update({
                        '排期_本厂型号': schedule_info['本厂型号'],
                        '排期_客户型号': schedule_info['客户型号'], 
                        '排期_订单数量': schedule_info['订单数量'],
                        '排期_OTS期': schedule_info['OTS期'],
                        '排期_生产线': schedule_info['生产线'],
                        '排期_备注': schedule_info['备注'],
                        '匹配状态': '已匹配'
                    })
                    matched_data.append(enhanced_record)
            else:
                # 未匹配的订单
                schedule_info = self.schedule_df[self.schedule_df['生产单号'] == pso].iloc[0]
                unmatched_record = {
                    '生产订单号': pso,
                    '客户订单号': '',
                    '产品型号': schedule_info['客户型号'],
                    '数量Pcs': schedule_info['订单数量'],
                    '排期_本厂型号': schedule_info['本厂型号'],
                    '排期_客户型号': schedule_info['客户型号'],
                    '排期_订单数量': schedule_info['订单数量'],
                    '排期_OTS期': schedule_info['OTS期'],
                    '排期_生产线': schedule_info['生产线'],
                    '排期_备注': schedule_info['备注'],
                    '匹配状态': '未找到欠料记录',
                    '欠料物料编号': '',
                    '欠料物料名称': '',
                    '欠料数量': 0,
                    '主供应商名称': '',
                    '供应商单价(原币)': 0,
                    '币种': '',
                    'RMB单价': 0,
                    '欠料金额(RMB)': 0,
                    '订单金额(RMB)': 0,
                    '每元投入回款': '无需投入',
                    '数据完整性标记': '不缺料订单'
                }
                matched_data.append(unmatched_record)
                unmatched_orders.append(pso)
        
        self.final_result = pd.DataFrame(matched_data)
        
        print(f"   ✅ 匹配完成:")
        print(f"      - 有欠料记录订单: {len(self.production_orders) - len(unmatched_orders)}个")
        print(f"      - 无欠料记录订单: {len(unmatched_orders)}个")
        print(f"      - 总分析记录数: {len(self.final_result)}条")
        
        if unmatched_orders:
            print(f"   🔍 无欠料记录的订单（可能不缺料）:")
            for order in unmatched_orders[:5]:  # 显示前5个
                schedule_info = self.schedule_df[self.schedule_df['生产单号'] == order].iloc[0]
                print(f"      - {order}: {schedule_info['客户型号']} ({schedule_info['订单数量']}pcs)")
            if len(unmatched_orders) > 5:
                print(f"      ... 还有{len(unmatched_orders)-5}个")
        
        return True
    
    def generate_shortage_summary(self):
        """生成欠料汇总分析"""
        print("=== 💰 生成欠料汇总分析 ===")
        
        if self.final_result is None:
            print("❌ 没有分析结果数据")
            return None
        
        # 1. 按订单汇总
        print("1. 按订单汇总欠料情况...")
        order_summary = self.final_result.groupby('生产订单号').agg({
            '排期_客户型号': 'first',
            '排期_订单数量': 'first',
            '排期_OTS期': 'first',
            '匹配状态': 'first',
            '欠料金额(RMB)': 'sum',
            '订单金额(RMB)': 'first',  # 订单金额每个订单只有一个值
            '欠料物料编号': 'count'  # 计算欠料物料种类数
        }).reset_index()
        
        order_summary.rename(columns={'欠料物料编号': '欠料物料种类'}, inplace=True)
        order_summary['欠料物料种类'] = order_summary['欠料物料种类'].replace(0, 0)  # 将NaN替换为0
        
        # 计算每个订单的投入产出比
        order_summary['投入产出比'] = order_summary.apply(lambda row:
            row['订单金额(RMB)'] / row['欠料金额(RMB)'] if row['欠料金额(RMB)'] > 0
            else '无需投入' if row['订单金额(RMB)'] > 0
            else 0, axis=1
        )
        
        # 2. 按供应商汇总
        print("2. 按供应商汇总采购需求...")
        shortage_records = self.final_result[
            (self.final_result['欠料物料编号'].notna()) & 
            (self.final_result['欠料物料编号'] != '')
        ]
        
        if not shortage_records.empty:
            supplier_summary = shortage_records.groupby('主供应商名称').agg({
                '欠料金额(RMB)': 'sum',
                '欠料物料编号': 'nunique',  # 唯一物料数
                '生产订单号': 'nunique'     # 影响的订单数
            }).reset_index()
            
            supplier_summary.rename(columns={
                '欠料物料编号': '需采购物料种类',
                '生产订单号': '影响订单数'
            }, inplace=True)
            
            supplier_summary = supplier_summary.sort_values('欠料金额(RMB)', ascending=False)
        else:
            supplier_summary = pd.DataFrame()
        
        # 3. 按物料汇总
        print("3. 按物料汇总采购清单...")
        if not shortage_records.empty:
            material_summary = shortage_records.groupby(['欠料物料编号', '欠料物料名称']).agg({
                '欠料数量': 'sum',
                '主供应商名称': 'first',
                '供应商单价(原币)': 'first',
                '币种': 'first',
                'RMB单价': 'first',
                '欠料金额(RMB)': 'sum',
                '生产订单号': lambda x: ', '.join(x.unique()[:3])  # 显示前3个相关订单
            }).reset_index()
            
            material_summary = material_summary.sort_values('欠料金额(RMB)', ascending=False)
            material_summary.rename(columns={'生产订单号': '相关生产订单'}, inplace=True)
        else:
            material_summary = pd.DataFrame()
        
        return {
            'order_summary': order_summary,
            'supplier_summary': supplier_summary, 
            'material_summary': material_summary
        }
    
    def print_summary_statistics(self, summaries):
        """打印汇总统计信息"""
        print("=== 📊 生产排期欠料统计 ===")
        
        order_summary = summaries['order_summary']
        supplier_summary = summaries['supplier_summary']
        material_summary = summaries['material_summary']
        
        # 基础统计
        total_orders = len(order_summary)
        shortage_orders = len(order_summary[order_summary['欠料金额(RMB)'] > 0])
        ready_orders = total_orders - shortage_orders
        
        total_shortage_amount = order_summary['欠料金额(RMB)'].sum()
        total_order_amount = order_summary['订单金额(RMB)'].sum()
        
        print(f"📋 订单统计:")
        print(f"   - 生产排期订单总数: {total_orders}个")
        print(f"   - 需要投入采购的订单: {shortage_orders}个 ({shortage_orders/total_orders*100:.1f}%)")
        print(f"   - 可立即生产的订单: {ready_orders}个 ({ready_orders/total_orders*100:.1f}%)")
        
        print(f"💰 金额统计:")
        print(f"   - 总欠料金额: ¥{total_shortage_amount:,.2f}")
        print(f"   - 总订单金额: ¥{total_order_amount:,.2f}")
        if total_shortage_amount > 0:
            overall_roi = total_order_amount / total_shortage_amount
            print(f"   - 整体投入产出比: {overall_roi:.2f}倍")
        
        if not supplier_summary.empty:
            print(f"🏭 供应商统计:")
            print(f"   - 涉及供应商: {len(supplier_summary)}家")
            print(f"   - 需采购物料种类: {supplier_summary['需采购物料种类'].sum()}种")
            
            print("   📈 主要供应商采购需求:")
            for _, supplier in supplier_summary.head(5).iterrows():
                print(f"      {supplier['主供应商名称']}: ¥{supplier['欠料金额(RMB)']:,.2f} " +
                      f"({supplier['需采购物料种类']}种物料, 影响{supplier['影响订单数']}个订单)")
        
        # 显示高价值订单
        print(f"🔝 欠料金额最高的5个订单:")
        top_shortage_orders = order_summary.nlargest(5, '欠料金额(RMB)')
        for _, order in top_shortage_orders.iterrows():
            roi_text = f"{order['投入产出比']:.2f}倍" if isinstance(order['投入产出比'], (int, float)) else order['投入产出比']
            print(f"   {order['生产订单号']}: {order['排期_客户型号']} " +
                  f"(欠料¥{order['欠料金额(RMB)']:,.2f}, ROI {roi_text})")
        
        # 显示无需投入的订单
        if ready_orders > 0:
            print(f"🚀 可立即生产的订单(前5个):")
            ready_order_list = order_summary[order_summary['欠料金额(RMB)'] == 0].head(5)
            for _, order in ready_order_list.iterrows():
                print(f"   {order['生产订单号']}: {order['排期_客户型号']} " +
                      f"({order['排期_订单数量']}pcs, 订单额¥{order['订单金额(RMB)']:,.2f})")
    
    def save_analysis_report(self, summaries):
        """保存分析报告到Excel"""
        print("=== 💾 保存分析报告 ===")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'生产排期欠料分析报告_{timestamp}.xlsx'
        
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # 主报表：详细分析明细
                self.final_result.to_excel(writer, sheet_name='生产排期欠料明细', index=False)
                
                # 汇总表
                if summaries:
                    summaries['order_summary'].to_excel(writer, sheet_name='按订单汇总', index=False)
                    
                    if not summaries['supplier_summary'].empty:
                        summaries['supplier_summary'].to_excel(writer, sheet_name='按供应商汇总', index=False)
                    
                    if not summaries['material_summary'].empty:
                        summaries['material_summary'].to_excel(writer, sheet_name='采购物料清单', index=False)
                
                # 统计汇总页
                order_summary = summaries['order_summary']
                stats_data = {
                    '统计项目': [
                        '生产排期订单总数',
                        '需要采购投入的订单数',
                        '可立即生产的订单数',
                        '总欠料金额(RMB)',
                        '总订单金额(RMB)',
                        '整体投入产出比',
                        '涉及供应商数量',
                        '需采购物料种类'
                    ],
                    '数值': [
                        len(order_summary),
                        len(order_summary[order_summary['欠料金额(RMB)'] > 0]),
                        len(order_summary[order_summary['欠料金额(RMB)'] == 0]),
                        f"¥{order_summary['欠料金额(RMB)'].sum():,.2f}",
                        f"¥{order_summary['订单金额(RMB)'].sum():,.2f}",
                        f"{order_summary['订单金额(RMB)'].sum() / order_summary['欠料金额(RMB)'].sum():.2f}倍" 
                        if order_summary['欠料金额(RMB)'].sum() > 0 else "无需投入",
                        len(summaries['supplier_summary']) if not summaries['supplier_summary'].empty else 0,
                        summaries['supplier_summary']['需采购物料种类'].sum() if not summaries['supplier_summary'].empty else 0
                    ]
                }
                
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='统计汇总', index=False)
            
            print(f"✅ 分析报告已保存: {filename}")
            print(f"📋 包含工作表: 生产排期欠料明细, 按订单汇总, 按供应商汇总, 采购物料清单, 统计汇总")
            
            return filename
            
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return None
    
    def run_analysis(self, schedule_file='250-08-28沙井成品生產排期.xlsx', 
                     analysis_file='银图PMC综合物料分析报告_20250831_170129.xlsx'):
        """运行完整分析流程"""
        print("🚀 开始生产排期欠料分析")
        print("="*60)
        
        try:
            # 1. 加载生产排期表
            if not self.load_production_schedule(schedule_file):
                return None
            
            # 2. 加载综合分析报表
            if not self.load_shortage_analysis(analysis_file):
                return None
            
            # 3. 分析匹配
            if not self.analyze_schedule_shortage():
                return None
            
            # 4. 生成汇总
            summaries = self.generate_shortage_summary()
            if summaries is None:
                return None
            
            # 5. 打印统计信息
            self.print_summary_statistics(summaries)
            
            # 6. 保存报告
            filename = self.save_analysis_report(summaries)
            
            print("\n" + "="*60)
            print(" "*15 + "🎉 分析完成！")
            print("="*60)
            
            return filename
            
        except Exception as e:
            print(f"❌ 分析过程出错: {e}")
            import traceback
            traceback.print_exc()
            return None

if __name__ == "__main__":
    analyzer = ProductionScheduleShortageAnalyzer()
    result = analyzer.run_analysis()
    
    if result:
        print(f"\n🎊 分析成功完成！报告已保存为: {result}")
    else:
        print("\n❌ 分析失败，请检查数据和错误信息")
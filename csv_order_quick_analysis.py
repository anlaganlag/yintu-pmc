#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基于ALL_PMC_ORDER.csv的快速订单欠料分析系统
594个订单的高效欠料分析 + ROI优先级排序
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class QuickCSVOrderAnalyzer:
    def __init__(self):
        self.order_csv_df = None
        self.shortage_df = None
        self.inventory_df = None
        self.supplier_df = None
        self.final_result = None
        
        # 汇率设置
        self.currency_rates = {'RMB': 1.0, 'USD': 7.30, 'HKD': 0.93, 'EUR': 7.85}
    
    def load_data(self):
        """快速加载所有必需数据"""
        print("=== 🔄 快速加载数据 ===")
        
        # 1. 加载CSV订单清单
        try:
            self.order_csv_df = pd.read_csv('ALL_PMC_ORERDER.csv', encoding='utf-8-sig')
            print(f"   ✅ CSV订单: {len(self.order_csv_df)}个")
        except Exception as e:
            print(f"   ❌ CSV加载失败: {e}")
            return False
        
        # 2. 加载欠料数据
        try:
            self.shortage_df = pd.read_excel('input/mat_owe_pso.xlsx', sheet_name='Sheet1', skiprows=1)
            if len(self.shortage_df.columns) >= 13:
                new_columns = ['订单编号', 'P-R对应', 'P-RBOM', '客户型号', 'OTS期', '开拉期', 
                              '下单日期', '物料编号', '物料名称', '领用部门', '工单需求', 
                              '仓存不足', '已购未返', '手头现有', '请购组']
                for i in range(min(len(new_columns), len(self.shortage_df.columns))):
                    self.shortage_df.rename(columns={self.shortage_df.columns[i]: new_columns[i]}, inplace=True)
            
            self.shortage_df = self.shortage_df.dropna(subset=['订单编号'])
            self.shortage_df = self.shortage_df[~self.shortage_df['物料名称'].astype(str).str.contains('已齐套|齐套', na=False)]
            print(f"   ✅ 欠料数据: {len(self.shortage_df)}条")
        except Exception as e:
            print(f"   ❌ 欠料数据加载失败: {e}")
            self.shortage_df = pd.DataFrame()
        
        # 3. 加载库存价格（简化版）
        try:
            self.inventory_df = pd.read_excel('input/inventory_list.xlsx')
            self.inventory_df['最终价格'] = self.inventory_df['最新報價'].fillna(self.inventory_df['成本單價'])
            self.inventory_df['最终价格'] = pd.to_numeric(self.inventory_df['最终价格'], errors='coerce').fillna(0)
            
            # 简化货币转换
            self.inventory_df['RMB单价'] = self.inventory_df['最终价格'] * self.inventory_df['貨幣'].map(
                {'USD': 7.30, 'HKD': 0.93, 'EUR': 7.85}
            ).fillna(1.0)
            
            print(f"   ✅ 库存数据: {len(self.inventory_df)}条")
        except Exception as e:
            print(f"   ❌ 库存数据加载失败: {e}")
            self.inventory_df = pd.DataFrame()
        
        # 4. 加载供应商数据（简化版）
        try:
            self.supplier_df = pd.read_excel('input/supplier.xlsx')
            self.supplier_df['单价_数值'] = pd.to_numeric(self.supplier_df['单价'], errors='coerce').fillna(0)
            self.supplier_df['供应商RMB单价'] = self.supplier_df['单价_数值'] * self.supplier_df['币种'].map(
                {'USD': 7.30, 'HKD': 0.93, 'EUR': 7.85}
            ).fillna(1.0)
            print(f"   ✅ 供应商数据: {len(self.supplier_df)}条")
        except Exception as e:
            print(f"   ❌ 供应商数据加载失败: {e}")
            self.supplier_df = pd.DataFrame()
        
        print("✅ 数据加载完成\\n")
        return True
    
    def quick_analysis(self):
        """快速分析方法"""
        print("=== 🎯 快速分析订单欠料情况 ===")
        
        results = []
        matched_orders = set()
        
        print(f"分析{len(self.order_csv_df)}个订单...")
        
        # 预处理：创建查找字典以提高效率
        shortage_dict = {}
        if not self.shortage_df.empty:
            for _, row in self.shortage_df.iterrows():
                order_no = str(row['订单编号']).strip()
                if order_no not in shortage_dict:
                    shortage_dict[order_no] = []
                shortage_dict[order_no].append(row)
        
        inventory_dict = {}
        if not self.inventory_df.empty:
            for _, row in self.inventory_df.iterrows():
                material_code = row['物項編號']
                inventory_dict[material_code] = row
        
        supplier_dict = {}
        if not self.supplier_df.empty:
            for _, row in self.supplier_df.iterrows():
                material_code = row['物项编号']
                if material_code not in supplier_dict:
                    supplier_dict[material_code] = []
                supplier_dict[material_code].append(row)
        
        # 批量处理订单
        for idx, csv_row in self.order_csv_df.iterrows():
            if idx % 100 == 0 and idx > 0:
                print(f"   处理进度: {idx}/{len(self.order_csv_df)}")
            
            order_no = str(csv_row.get('订单号', '')).strip()
            pr_order = str(csv_row.get('P-R订单号转换', '')).strip()
            
            # 查找欠料信息
            shortage_records = shortage_dict.get(order_no, [])
            if not shortage_records and pr_order and pr_order != 'nan':
                shortage_records = shortage_dict.get(pr_order, [])
            
            default_order_amount = 7300  # 默认1000USD * 7.3汇率
            
            if shortage_records:
                matched_orders.add(order_no)
                
                for shortage_row in shortage_records:
                    material_code = shortage_row.get('物料编号', '')
                    material_name = shortage_row.get('物料名称', '')
                    shortage_qty = pd.to_numeric(shortage_row.get('仓存不足', 0), errors='coerce')
                    
                    # 查找库存价格
                    inventory_info = inventory_dict.get(material_code, {})
                    rmb_price = inventory_info.get('RMB单价', 0)
                    
                    # 查找最低价供应商
                    supplier_records = supplier_dict.get(material_code, [])
                    if supplier_records:
                        # 选择最低价供应商
                        best_supplier = min(supplier_records, 
                                          key=lambda x: x.get('供应商RMB单价', float('inf')))
                        supplier_name = best_supplier.get('供应商名称', '')
                        supplier_price = best_supplier.get('单价', 0)
                        supplier_currency = best_supplier.get('币种', 'RMB')
                    else:
                        supplier_name = ''
                        supplier_price = 0
                        supplier_currency = 'RMB'
                    
                    # 计算欠料金额
                    shortage_amount = shortage_qty * rmb_price
                    
                    result = {
                        'CSV序号': csv_row.get('序号', idx + 1),
                        'Excel行号': csv_row.get('Excel行号', ''),
                        '生产线': csv_row.get('生产线', ''),
                        '订单号': order_no,
                        'P-R转换': pr_order if pr_order != 'nan' else '',
                        '本厂型号': csv_row.get('本廠型號', ''),
                        '客户型号': csv_row.get('客戶型號', ''),
                        '订单数量': pd.to_numeric(csv_row.get('訂單數量', 0), errors='coerce'),
                        'OTS期': csv_row.get('OTS期', ''),
                        
                        '欠料物料编号': material_code,
                        '欠料物料名称': material_name,
                        '欠料数量': shortage_qty,
                        '库存单价(RMB)': rmb_price,
                        '欠料金额(RMB)': shortage_amount,
                        
                        '主供应商': supplier_name,
                        '供应商单价': supplier_price,
                        '供应商币种': supplier_currency,
                        
                        '订单金额(RMB)': default_order_amount,
                        '匹配状态': '有欠料'
                    }
                    results.append(result)
            else:
                # 无欠料订单
                result = {
                    'CSV序号': csv_row.get('序号', idx + 1),
                    'Excel行号': csv_row.get('Excel行号', ''),
                    '生产线': csv_row.get('生产线', ''),
                    '订单号': order_no,
                    'P-R转换': pr_order if pr_order != 'nan' else '',
                    '本厂型号': csv_row.get('本廠型號', ''),
                    '客户型号': csv_row.get('客戶型號', ''),
                    '订单数量': pd.to_numeric(csv_row.get('訂單數量', 0), errors='coerce'),
                    'OTS期': csv_row.get('OTS期', ''),
                    
                    '欠料物料编号': '',
                    '欠料物料名称': '',
                    '欠料数量': 0,
                    '库存单价(RMB)': 0,
                    '欠料金额(RMB)': 0,
                    
                    '主供应商': '',
                    '供应商单价': 0,
                    '供应商币种': 'RMB',
                    
                    '订单金额(RMB)': default_order_amount,
                    '匹配状态': '不缺料'
                }
                results.append(result)
        
        self.final_result = pd.DataFrame(results)
        
        print(f"   ✅ 分析完成:")
        print(f"      - 有欠料订单: {len(matched_orders)}个")
        print(f"      - 总分析记录: {len(results)}条")
        
        return True
    
    def calculate_priority(self):
        """计算优先级"""
        print("=== 💰 计算ROI优先级 ===")
        
        # 按订单汇总
        order_summary = self.final_result.groupby('订单号').agg({
            'CSV序号': 'first',
            '生产线': 'first',
            '客户型号': 'first',
            '订单数量': 'first',
            'OTS期': 'first',
            '欠料金额(RMB)': 'sum',
            '订单金额(RMB)': 'first',
            '欠料物料编号': 'count'
        }).reset_index()
        
        order_summary.rename(columns={'欠料物料编号': '欠料物料种类'}, inplace=True)
        
        # 修正欠料物料种类
        for idx, row in order_summary.iterrows():
            actual_count = len(self.final_result[
                (self.final_result['订单号'] == row['订单号']) & 
                (self.final_result['欠料物料编号'] != '')
            ])
            order_summary.at[idx, '欠料物料种类'] = actual_count
        
        # 计算ROI
        def calc_roi(row):
            if row['欠料金额(RMB)'] > 0:
                return row['订单金额(RMB)'] / row['欠料金额(RMB)']
            elif row['订单金额(RMB)'] > 0:
                return 999999  # 无需投入
            else:
                return 0
        
        order_summary['ROI'] = order_summary.apply(calc_roi, axis=1)
        
        # 优先级分类
        def get_priority(roi):
            if roi >= 999999:
                return "无需投入"
            elif roi >= 5.0:
                return "高优先级"
            elif roi >= 2.0:
                return "中优先级"
            elif roi >= 1.0:
                return "低优先级"
            else:
                return "暂缓生产"
        
        order_summary['优先级'] = order_summary['ROI'].apply(get_priority)
        
        # 排序
        priority_sorted = order_summary.sort_values('ROI', ascending=False)
        priority_sorted['生产顺序'] = range(1, len(priority_sorted) + 1)
        
        # 高优先级订单（ROI>=2.0或无需投入）
        high_priority = priority_sorted[
            (priority_sorted['ROI'] >= 2.0) | (priority_sorted['ROI'] >= 999999)
        ]
        
        self.order_summary = order_summary
        self.priority_sorted = priority_sorted
        self.high_priority = high_priority
        
        print(f"   📊 优先级统计:")
        priority_stats = order_summary['优先级'].value_counts()
        for priority, count in priority_stats.items():
            pct = count / len(order_summary) * 100
            print(f"      {priority}: {count}个 ({pct:.1f}%)")
        
        print(f"   💡 建议优先生产: {len(high_priority)}个订单")
        
        return True
    
    def save_report(self):
        """保存报告"""
        print("=== 💾 保存报告 ===")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'594订单CSV快速分析_{timestamp}.xlsx'
        
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # 1. 详细分析
                self.final_result.to_excel(writer, sheet_name='详细分析', index=False)
                
                # 2. 订单汇总
                self.order_summary.to_excel(writer, sheet_name='订单汇总', index=False)
                
                # 3. 优先级排序
                self.priority_sorted.to_excel(writer, sheet_name='优先级排序', index=False)
                
                # 4. 建议优先生产
                self.high_priority.to_excel(writer, sheet_name='建议优先生产', index=False)
                
                # 5. 供应商汇总
                if len(self.final_result[self.final_result['主供应商'] != '']) > 0:
                    supplier_summary = self.final_result[self.final_result['主供应商'] != ''].groupby('主供应商').agg({
                        '欠料金额(RMB)': 'sum',
                        '欠料物料编号': 'nunique',
                        '订单号': 'nunique'
                    }).reset_index().sort_values('欠料金额(RMB)', ascending=False)
                    supplier_summary.to_excel(writer, sheet_name='供应商汇总', index=False)
            
            print(f"✅ 报告已保存: {filename}")
            return filename
            
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return None
    
    def run(self):
        """运行分析"""
        print("🚀 开始594个CSV订单快速分析")
        print("="*50)
        
        start_time = datetime.now()
        
        try:
            if not self.load_data():
                return None
            
            if not self.quick_analysis():
                return None
            
            if not self.calculate_priority():
                return None
            
            filename = self.save_report()
            
            # 显示结果
            duration = datetime.now() - start_time
            
            print("\\n" + "="*50)
            print("🎉 分析完成！")
            print("="*50)
            
            total_orders = len(self.order_summary)
            no_shortage = len(self.order_summary[self.order_summary['欠料金额(RMB)'] == 0])
            high_priority_count = len(self.high_priority)
            total_shortage = self.final_result['欠料金额(RMB)'].sum()
            total_order_value = self.order_summary['订单金额(RMB)'].sum()
            
            print(f"📊 关键统计:")
            print(f"   - 总订单数: {total_orders}个")
            print(f"   - 不缺料订单: {no_shortage}个 ({no_shortage/total_orders*100:.1f}%)")
            print(f"   - 建议优先生产: {high_priority_count}个 ({high_priority_count/total_orders*100:.1f}%)")
            print(f"   - 总欠料金额: ¥{total_shortage:,.2f}")
            print(f"   - 整体ROI: {total_order_value/total_shortage:.2f}倍" if total_shortage > 0 else "无需投入")
            print(f"   - 处理时间: {duration}")
            
            # Top5欠料订单
            top5 = self.order_summary.nlargest(5, '欠料金额(RMB)')
            print(f"\\n🔝 欠料金额最高5个订单:")
            for _, order in top5.iterrows():
                if order['欠料金额(RMB)'] > 0:
                    roi = order['ROI']
                    roi_text = f"{roi:.2f}倍" if roi < 999999 else "无需投入"
                    print(f"   {order['订单号']}: {order['客户型号']} (¥{order['欠料金额(RMB)']:,.2f}, ROI {roi_text})")
            
            return filename
            
        except Exception as e:
            print(f"❌ 分析失败: {e}")
            return None

if __name__ == "__main__":
    analyzer = QuickCSVOrderAnalyzer()
    result = analyzer.run()
    
    if result:
        print(f"\\n🎊 分析成功！文件: {result}")
    else:
        print("\\n❌ 分析失败")
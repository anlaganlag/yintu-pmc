#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图工厂订单物料汇总分析系统
基于国内+柬埔寨订单 与 汇总表 进行数据匹配分析
生成三个汇总报表：订单缺料明细、8月采购汇总、供应商汇总
"""

import pandas as pd
import numpy as np
from datetime import datetime
import re

class OrderMaterialAnalyzer:
    def __init__(self):
        self.domestic_orders = None
        self.cambodia_orders = None 
        self.summary_orders = None
        self.inventory = None
        self.merged_data = None
        
    def load_data(self):
        """加载所有数据源"""
        print("=== 加载数据源 ===")
        
        # 1. 加载国内订单 (8月+9月)
        print("1. 加载国内订单...")
        domestic_aug = pd.read_excel(r'D:\yingtu-PMC\(国内)8月9月订单.xlsx', sheet_name='8月')
        domestic_sep = pd.read_excel(r'D:\yingtu-PMC\(国内)8月9月订单.xlsx', sheet_name='9月')
        domestic_aug['月份'] = '8月'
        domestic_sep['月份'] = '9月'
        domestic_aug['工厂'] = '国内'
        domestic_sep['工厂'] = '国内'
        self.domestic_orders = pd.concat([domestic_aug, domestic_sep], ignore_index=True)
        print(f"   国内订单总数: {len(self.domestic_orders)}条")
        
        # 2. 加载柬埔寨订单 (8月+9月) 
        print("2. 加载柬埔寨订单...")
        cambodia_aug = pd.read_excel(r'D:\yingtu-PMC\(柬埔寨)8月9月订单.xlsx', sheet_name='8月 -柬')
        cambodia_sep = pd.read_excel(r'D:\yingtu-PMC\(柬埔寨)8月9月订单.xlsx', sheet_name='9月 -柬')
        cambodia_aug['月份'] = '8月'
        cambodia_sep['月份'] = '9月'  
        cambodia_aug['工厂'] = '柬埔寨'
        cambodia_sep['工厂'] = '柬埔寨'
        self.cambodia_orders = pd.concat([cambodia_aug, cambodia_sep], ignore_index=True)
        print(f"   柬埔寨订单总数: {len(self.cambodia_orders)}条")
        
        # 3. 加载汇总表 (8月+9月)
        print("3. 加载订单汇总表...")
        summary_aug = pd.read_excel(r'D:\yingtu-PMC\8月9月订单汇总表.xlsx', sheet_name='8月份', skiprows=1)
        summary_sep = pd.read_excel(r'D:\yingtu-PMC\8月9月订单汇总表.xlsx', sheet_name='9月份 ', skiprows=1)
        summary_aug['月份'] = '8月'
        summary_sep['月份'] = '9月'
        self.summary_orders = pd.concat([summary_aug, summary_sep], ignore_index=True)
        # 清理无效行
        self.summary_orders = self.summary_orders.dropna(subset=['本廠P/O'])
        print(f"   汇总表记录数: {len(self.summary_orders)}条")
        
        # 4. 加载库存清单
        print("4. 加载库存清单...")
        self.inventory = pd.read_excel(r'D:\yingtu-PMC\银图工厂库存清单-20250822.xlsx', 
                                     sheet_name='银图库存总表')
        print(f"   库存记录数: {len(self.inventory)}条")
        
        print("✅ 数据加载完成\\n")
        
    def merge_orders(self):
        """合并国内和柬埔寨订单数据"""
        print("=== 合并订单数据 ===")
        
        # 合并所有订单
        all_orders = pd.concat([self.domestic_orders, self.cambodia_orders], ignore_index=True)
        
        # 标准化列名
        all_orders = all_orders.rename(columns={
            '生 產 單 号(  廠方 )': '生产订单号',
            '生 產 單 号(客方 )': '客户订单号', 
            '型 號( 廠方/客方 )': '产品型号',
            '數 量  (Pcs)': '订单数量',
            '目的地': '目的地',
            '客期': '客户要求交期',
            'BOM NO.': 'BOM编号'
        })
        
        # 与汇总表进行LEFT JOIN匹配
        print("匹配汇总表数据...")
        merged = all_orders.merge(
            self.summary_orders,
            left_on='生产订单号',
            right_on='本廠P/O',
            how='left',
            suffixes=('', '_汇总表')
        )
        
        self.merged_data = merged
        print(f"   合并后总记录数: {len(merged)}条")
        print(f"   成功匹配汇总表: {len(merged[merged['本廠P/O'].notna()])}条")
        print(f"   未匹配(主要是TSO): {len(merged[merged['本廠P/O'].isna()])}条")
        print("✅ 订单合并完成\\n")
        
    def parse_shortage_info(self, row):
        """解析缺料信息"""
        shortage_list = []
        suppliers = []
        
        # 解析物料状况
        material_status = str(row['物料狀況']) if pd.notna(row['物料狀況']) else ''
        if material_status and material_status != 'OK':
            shortage_list.append(f"物料状况:{material_status}")
            
        # 解析PU-五金
        pu_metal = str(row['PU-五金']) if pd.notna(row['PU-五金']) else ''
        if pu_metal == 'NO':
            shortage_list.append("五金物料")
        elif pu_metal not in ['OK', 'NO', ''] and pu_metal != 'nan':
            # 提取供应商信息，如 "412-03021001C-威達,25D (8/15測)"
            supplier_match = re.search(r'-([^-,\\(]+)', pu_metal)
            if supplier_match:
                supplier = supplier_match.group(1)
                suppliers.append(supplier)
                shortage_list.append(f"五金物料(供应商:{supplier})")
            else:
                shortage_list.append(f"五金物料:{pu_metal}")
                
        # 解析PU-電子料  
        pu_electronic = str(row['PU-電子料']) if pd.notna(row['PU-電子料']) else ''
        if pu_electronic == 'NO':
            shortage_list.append("电子料")
        elif pu_electronic not in ['OK', 'NO', ''] and pu_electronic != 'nan':
            supplier_match = re.search(r'-([^-,\\(]+)', pu_electronic)
            if supplier_match:
                supplier = supplier_match.group(1)
                suppliers.append(supplier)
                shortage_list.append(f"电子料(供应商:{supplier})")
            else:
                shortage_list.append(f"电子料:{pu_electronic}")
                
        # 解析PU-包裝
        pu_package = str(row['PU-包裝']) if pd.notna(row['PU-包裝']) else ''
        if pu_package == 'NO':
            shortage_list.append("包装材料")
        elif pu_package not in ['OK', 'NO', ''] and pu_package != 'nan':
            shortage_list.append(f"包装材料:{pu_package}")
            
        return {
            'shortage_items': '; '.join(shortage_list) if shortage_list else '无缺料',
            'suppliers': '; '.join(set(suppliers)) if suppliers else '',
            'shortage_status': '有缺料' if shortage_list else '齐料'
        }
        
    def generate_report1_order_shortage(self):
        """生成表1: 订单缺料明细"""
        print("=== 生成表1: 订单缺料明细 ===")
        
        # 基于合并数据生成报表
        report_data = []
        
        for idx, row in self.merged_data.iterrows():
            # 解析缺料信息
            shortage_info = self.parse_shortage_info(row)
            
            # 获取采购价格信息 (从库存清单匹配)
            # 这里需要通过BOM或型号匹配，暂时使用平均价格
            avg_price = "待查询"  # 后续可以通过BOM匹配具体价格
            
            # 安全获取字段值，处理可能的列名冲突
            month = row['月份'] if '月份' in row.index else (row['月份_汇总表'] if '月份_汇总表' in row.index else '未知')
            
            report_data.append({
                '客户订单号': row['客户订单号'],
                '生产订单号': row['生产订单号'], 
                '产品型号': row['产品型号'],
                '订单数量': row['订单数量'],
                '月份': month,
                '工厂': row['工厂'],
                '目的地': row['目的地'],
                '客户交期': row['客户要求交期'],
                '缺料状态': shortage_info['shortage_status'],
                '缺料物料': shortage_info['shortage_items'],
                '涉及供应商': shortage_info['suppliers'],
                '预估采购成本': avg_price,
                'BOM编号': row['BOM编号']
            })
            
        report1_df = pd.DataFrame(report_data)
        
        print(f"   生成订单缺料明细: {len(report1_df)}条记录")
        print(f"   有缺料订单: {len(report1_df[report1_df['缺料状态']=='有缺料'])}条")
        print(f"   齐料订单: {len(report1_df[report1_df['缺料状态']=='齐料'])}条")
        
        return report1_df
        
    def generate_report2_august_purchase(self):
        """生成表2: 8月订单所需采购汇总"""
        print("=== 生成表2: 8月订单采购汇总 ===")
        
        # 筛选8月订单 - 安全处理列名
        month_col = '月份' if '月份' in self.merged_data.columns else '月份_汇总表'
        aug_orders = self.merged_data[self.merged_data[month_col] == '8月'].copy() if month_col in self.merged_data.columns else self.merged_data.copy()
        
        purchase_summary = []
        
        for idx, row in aug_orders.iterrows():
            if pd.isna(row['本廠P/O']):  # 跳过未匹配的TSO订单
                continue
                
            shortage_info = self.parse_shortage_info(row)
            
            if shortage_info['shortage_status'] == '有缺料':
                # 估算采购金额 (这里需要更详细的BOM成本计算)
                estimated_cost = row['订单数量'] * 10 if pd.notna(row['订单数量']) else 0  # 临时估算
                
                purchase_summary.append({
                    '生产订单号': row['生产订单号'],
                    '产品型号': row['产品型号'], 
                    '订单数量': row['订单数量'],
                    '需采购物料': shortage_info['shortage_items'],
                    '涉及供应商': shortage_info['suppliers'],
                    '预估采购金额': estimated_cost,
                    '目的地': row['目的地'],
                    '客户交期': row['客户要求交期']
                })
                
        report2_df = pd.DataFrame(purchase_summary)
        
        print(f"   8月需采购订单: {len(report2_df)}条")
        if len(report2_df) > 0:
            total_cost = report2_df['预估采购金额'].sum()
            print(f"   预估采购总金额: {total_cost:,.0f}元")
        
        return report2_df
        
    def generate_report3_supplier_summary(self):
        """生成表3: 按供应商汇总的订单清单"""
        print("=== 生成表3: 供应商汇总 ===")
        
        supplier_summary = {}
        
        for idx, row in self.merged_data.iterrows():
            if pd.isna(row['本廠P/O']):  # 跳过未匹配的TSO订单
                continue
                
            shortage_info = self.parse_shortage_info(row)
            
            if shortage_info['suppliers']:
                suppliers = shortage_info['suppliers'].split('; ')
                for supplier in suppliers:
                    if supplier not in supplier_summary:
                        supplier_summary[supplier] = {
                            'orders': [],
                            'materials': set(),
                            'total_quantity': 0
                        }
                    
                    month = row['月份'] if '月份' in row.index else (row['月份_汇总表'] if '月份_汇总表' in row.index else '未知')
                    supplier_summary[supplier]['orders'].append({
                        '生产订单号': row['生产订单号'],
                        '产品型号': row['产品型号'],
                        '数量': row['订单数量'],
                        '月份': month
                    })
                    
                    # 添加物料类型
                    if '五金' in shortage_info['shortage_items']:
                        supplier_summary[supplier]['materials'].add('五金物料')
                    if '电子料' in shortage_info['shortage_items']:
                        supplier_summary[supplier]['materials'].add('电子料')
                    if '包装' in shortage_info['shortage_items']:
                        supplier_summary[supplier]['materials'].add('包装材料')
                        
                    supplier_summary[supplier]['total_quantity'] += row['订单数量'] if pd.notna(row['订单数量']) else 0
        
        # 转换为DataFrame
        report3_data = []
        for supplier, info in supplier_summary.items():
            order_list = [f"{order['生产订单号']}({order['产品型号']})" for order in info['orders']]
            
            report3_data.append({
                '供应商': supplier,
                '涉及物料类型': '; '.join(info['materials']),
                '相关订单数量': len(info['orders']),
                '相关订单列表': '; '.join(order_list),
                '总订单数量': info['total_quantity'],
                '预估采购金额': info['total_quantity'] * 15  # 临时估算
            })
            
        report3_df = pd.DataFrame(report3_data)
        report3_df = report3_df.sort_values('总订单数量', ascending=False)
        
        print(f"   涉及供应商: {len(report3_df)}家")
        
        return report3_df
        
    def save_reports(self, report1_df, report2_df, report3_df):
        """保存所有报表"""
        print("=== 保存报表文件 ===")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f'订单物料汇总分析报告_{timestamp}.xlsx'
        
        with pd.ExcelWriter(f'D:\\yingtu-PMC\\{filename}', engine='openpyxl') as writer:
            report1_df.to_excel(writer, sheet_name='1_订单缺料明细', index=False)
            report2_df.to_excel(writer, sheet_name='2_8月采购汇总', index=False) 
            report3_df.to_excel(writer, sheet_name='3_供应商汇总', index=False)
            
        print(f"✅ 报表已保存: {filename}")
        return filename
        
    def run_analysis(self):
        """执行完整分析流程"""
        try:
            self.load_data()
            self.merge_orders()
            
            report1 = self.generate_report1_order_shortage()
            report2 = self.generate_report2_august_purchase()
            report3 = self.generate_report3_supplier_summary()
            
            filename = self.save_reports(report1, report2, report3)
            
            print("\\n" + "="*80)
            print(" "*25 + "📊 分析完成汇总")
            print("="*80)
            print(f"总订单数: {len(self.merged_data)}")
            print(f"成功匹配: {len(self.merged_data[self.merged_data['本廠P/O'].notna()])}")
            print(f"缺料订单: {len(report1[report1['缺料状态']=='有缺料'])}")
            print(f"8月需采购: {len(report2)}")
            print(f"涉及供应商: {len(report3)}")
            print(f"\\n报表文件: {filename}")
            
            return report1, report2, report3
            
        except Exception as e:
            print(f"❌ 分析过程出错: {e}")
            import traceback
            traceback.print_exc()
            return None, None, None

if __name__ == "__main__":
    analyzer = OrderMaterialAnalyzer()
    analyzer.run_analysis()
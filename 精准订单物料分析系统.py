#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图工厂精准订单物料分析系统 (混合方案)
基于订单表 + 欠料表 + 库存表的精确匹配分析
实现汇总表需求.md的三个要求
"""

import pandas as pd
import numpy as np
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class PreciseOrderMaterialAnalyzer:
    def __init__(self):
        self.orders_df = None           # 订单数据
        self.shortage_df = None         # 欠料数据  
        self.inventory_df = None        # 库存价格数据
        self.merged_precise = None      # 精确匹配结果
        self.merged_estimated = None    # 估算补全结果
        self.currency_rates = {         # 汇率 (转换为RMB)
            'RMB': 1.0,
            'HKD': 0.93,  # 1 HKD = 0.93 RMB
            'USD': 7.20,  # 1 USD = 7.20 RMB  
            'EUR': 7.85   # 1 EUR = 7.85 RMB
        }
        
    def load_all_data(self):
        """加载所有数据源"""
        print("=== 🔄 加载数据源 ===")
        
        # 1. 加载全部订单数据
        print("1. 加载(全部)8月9月订单...")
        orders_aug = pd.read_excel(r'D:\yingtu-PMC\(全部)8月9月订单.xlsx', sheet_name='8月')
        orders_sep = pd.read_excel(r'D:\yingtu-PMC\(全部)8月9月订单.xlsx', sheet_name='9月')
        orders_aug['月份'] = '8月'
        orders_sep['月份'] = '9月'
        self.orders_df = pd.concat([orders_aug, orders_sep], ignore_index=True)
        
        # 标准化列名
        self.orders_df = self.orders_df.rename(columns={
            '生 產 單 号(  廠方 )': '生产单号',
            '生 產 單 号(客方 )': '客户订单号',
            '型 號( 廠方/客方 )': '产品型号',
            '數 量  (Pcs)': '订单数量',
            'BOM NO.': 'BOM编号',
            '客期': '交期'
        })
        print(f"   ✅ 订单总数: {len(self.orders_df)}条 (8月:{len(orders_aug)}, 9月:{len(orders_sep)})")
        
        # 2. 加载欠料表
        print("2. 加载owe-mat.xls欠料表...")
        try:
            # 读取第一个sheet，跳过表头
            self.shortage_df = pd.read_excel(r'D:\yingtu-PMC\owe-mat.xls', 
                                           sheet_name='Sheet1', engine='xlrd', skiprows=1)
            
            # 标准化列名 (基于分析结果)
            if len(self.shortage_df.columns) >= 11:
                self.shortage_df = self.shortage_df.rename(columns={
                    self.shortage_df.columns[0]: '订单编号',
                    self.shortage_df.columns[1]: '客户型号', 
                    self.shortage_df.columns[2]: 'OTS期',
                    self.shortage_df.columns[3]: '开拉期',
                    self.shortage_df.columns[4]: '下单日期',
                    self.shortage_df.columns[5]: '物料编号',
                    self.shortage_df.columns[6]: '物料名称',
                    self.shortage_df.columns[7]: '领用部门',
                    self.shortage_df.columns[8]: '工单需求',
                    self.shortage_df.columns[9]: '仓存不足',
                    self.shortage_df.columns[10]: '已购未返'
                })
            
            # 清理数据
            self.shortage_df = self.shortage_df.dropna(subset=['订单编号'])
            # 移除包含"已齐套"的记录 (这些不是欠料)
            self.shortage_df = self.shortage_df[~self.shortage_df['物料名称'].astype(str).str.contains('已齐套|齐套', na=False)]
            print(f"   ✅ 欠料记录: {len(self.shortage_df)}条")
            
        except Exception as e:
            print(f"   ❌ 欠料表加载失败: {e}")
            self.shortage_df = pd.DataFrame()
        
        # 3. 加载库存价格表
        print("3. 加载inventory_list.xlsx库存表...")
        self.inventory_df = pd.read_excel(r'D:\yingtu-PMC\inventory_list.xlsx')
        
        # 价格处理：优先最新报价，回退到成本单价
        self.inventory_df['最终价格'] = self.inventory_df['最新報價'].fillna(self.inventory_df['成本單價'])
        self.inventory_df['最终价格'] = pd.to_numeric(self.inventory_df['最终价格'], errors='coerce').fillna(0)
        
        # 货币转换为RMB
        def convert_to_rmb(row):
            price = row['最终价格']
            currency = str(row.get('貨幣', 'RMB')).upper()
            rate = self.currency_rates.get(currency, 1.0)
            return price * rate
        
        self.inventory_df['RMB价格'] = self.inventory_df.apply(convert_to_rmb, axis=1)
        
        valid_prices = len(self.inventory_df[self.inventory_df['RMB价格'] > 0])
        print(f"   ✅ 库存物料: {len(self.inventory_df)}条, 有效价格: {valid_prices}条")
        
        print("✅ 数据加载完成\n")
        
    def precise_matching(self):
        """方案1: 精确匹配 - 订单→欠料→库存→价格"""
        print("=== 🎯 精确匹配分析 (方案1) ===")
        
        if self.shortage_df.empty:
            print("❌ 欠料表为空，跳过精确匹配")
            self.merged_precise = pd.DataFrame()
            return
        
        # 1. 订单与欠料匹配
        print("1. 订单与欠料表匹配...")
        # 清理订单编号格式
        self.orders_df['生产单号_清理'] = self.orders_df['生产单号'].astype(str).str.strip()
        self.shortage_df['订单编号_清理'] = self.shortage_df['订单编号'].astype(str).str.strip()
        
        # LEFT JOIN: 订单 ← 欠料
        order_shortage = self.orders_df.merge(
            self.shortage_df,
            left_on='生产单号_清理',
            right_on='订单编号_清理',
            how='left'
        )
        
        # 只保留有欠料的记录
        precise_records = order_shortage[order_shortage['物料编号'].notna()].copy()
        print(f"   ✅ 匹配到欠料的订单: {precise_records['生产单号'].nunique()}个")
        print(f"   ✅ 欠料明细记录: {len(precise_records)}条")
        
        if len(precise_records) == 0:
            print("❌ 无匹配的欠料记录")
            self.merged_precise = pd.DataFrame()
            return
        
        # 2. 欠料与库存价格匹配
        print("2. 欠料与库存价格匹配...")
        precise_with_price = precise_records.merge(
            self.inventory_df[['物項編號', '物項名稱', 'RMB价格', '貨幣', '最终价格']],
            left_on='物料编号',
            right_on='物項編號',
            how='left'
        )
        
        # 3. 计算欠料金额
        print("3. 计算精确欠料金额...")
        precise_with_price['仓存不足_数值'] = pd.to_numeric(precise_with_price['仓存不足'], errors='coerce').fillna(0)
        precise_with_price['欠料金额'] = precise_with_price['仓存不足_数值'] * precise_with_price['RMB价格']
        precise_with_price['计算方式'] = '精确计算'
        
        # 统计匹配情况
        matched_price = len(precise_with_price[precise_with_price['RMB价格'] > 0])
        total_amount = precise_with_price['欠料金额'].sum()
        
        print(f"   ✅ 匹配到价格: {matched_price}/{len(precise_with_price)}条")
        print(f"   💰 精确计算欠料总金额: ¥{total_amount:,.2f}")
        
        self.merged_precise = precise_with_price
        print("✅ 精确匹配完成\n")
        
    def estimated_completion(self):
        """方案2: 估算补全 - 对未在欠料表的订单进行估算"""
        print("=== 📊 估算补全分析 (方案2) ===")
        
        # 1. 找出未在精确匹配中的订单
        if not self.merged_precise.empty:
            precise_orders = set(self.merged_precise['生产单号'].unique())
            remaining_orders = self.orders_df[~self.orders_df['生产单号'].isin(precise_orders)].copy()
        else:
            remaining_orders = self.orders_df.copy()
            
        print(f"1. 需要估算的订单: {len(remaining_orders)}条")
        
        if len(remaining_orders) == 0:
            print("✅ 无需估算，所有订单都已精确匹配")
            self.merged_estimated = pd.DataFrame()
            return
        
        # 2. 基于产品复杂度的估算模型
        print("2. 建立估算模型...")
        
        # 产品复杂度评估 (基于型号特征)
        def estimate_complexity(product_model):
            model_str = str(product_model).upper()
            complexity = 1.0  # 基础复杂度
            
            # 根据产品特征调整复杂度
            if any(feature in model_str for feature in ['SP', 'AS']):
                complexity *= 1.2  # 标准产品
            if any(feature in model_str for feature in ['BT', 'BAB']):  
                complexity *= 1.5  # 复杂产品
            if '/' in model_str:
                complexity *= 1.1  # 组合型号
                
            return complexity
        
        # 3. 估算物料成本
        print("3. 估算物料成本...")
        
        # 从库存表获取不同类型物料的平均价格
        avg_prices = {}
        material_types = ['BOM', '物項']
        
        for mat_type in material_types:
            type_materials = self.inventory_df[
                (self.inventory_df['物項類型'] == mat_type) & 
                (self.inventory_df['RMB价格'] > 0)
            ]
            if len(type_materials) > 0:
                avg_prices[mat_type] = type_materials['RMB价格'].mean()
            else:
                avg_prices[mat_type] = 10.0  # 默认价格
                
        print(f"   📊 BOM类物料均价: ¥{avg_prices.get('BOM', 0):.2f}")
        print(f"   📊 物項类物料均价: ¥{avg_prices.get('物項', 0):.2f}")
        
        # 4. 为每个订单估算欠料金额
        estimated_records = []
        
        for _, order in remaining_orders.iterrows():
            complexity = estimate_complexity(order['产品型号'])
            order_qty = pd.to_numeric(order['订单数量'], errors='coerce')
            
            if pd.isna(order_qty) or order_qty <= 0:
                order_qty = 1000  # 默认数量
            
            # 估算不同物料类型的需求
            bom_cost = avg_prices.get('BOM', 10) * complexity * (order_qty / 1000) * 20  # BOM物料
            item_cost = avg_prices.get('物項', 10) * complexity * (order_qty / 1000) * 15  # 物項类
            
            total_estimated_cost = bom_cost + item_cost
            
            estimated_records.append({
                'ITEM NO.': order.get('ITEM NO.', ''),
                '生产单号': order['生产单号'],
                '客户订单号': order['客户订单号'],
                '产品型号': order['产品型号'],
                '订单数量': order['订单数量'],
                '月份': order['月份'],
                '目的地': order.get('目的地', ''),
                '交期': order.get('交期', ''),
                'BOM编号': order.get('BOM编号', ''),
                
                '物料编号': f"估算-{order['产品型号']}",
                '物料名称': f"估算物料成本-{order['产品型号']}",
                '仓存不足': 1,  # 估算数量
                '欠料金额': total_estimated_cost,
                '计算方式': '估算',
                'RMB价格': total_estimated_cost,
                '复杂度系数': complexity
            })
        
        self.merged_estimated = pd.DataFrame(estimated_records)
        
        total_estimated = self.merged_estimated['欠料金额'].sum()
        print(f"   💰 估算欠料总金额: ¥{total_estimated:,.2f}")
        print("✅ 估算补全完成\n")
        
    def generate_report1_order_shortage_detail(self):
        """生成表1: 订单缺料明细"""
        print("=== 📋 生成表1: 订单缺料明细 ===")
        
        # 合并精确匹配和估算结果
        all_results = []
        
        if not self.merged_precise.empty:
            all_results.append(self.merged_precise)
        if not self.merged_estimated.empty:
            all_results.append(self.merged_estimated)
            
        if not all_results:
            print("❌ 无数据生成报表1")
            return pd.DataFrame()
        
        combined_df = pd.concat(all_results, ignore_index=True, sort=False)
        
        # 统一字段，生成最终报表
        report1_data = []
        
        for _, row in combined_df.iterrows():
            report1_data.append({
                '客户订单号': row['客户订单号'],
                '生产订单号': row['生产单号'],
                '产品型号': row['产品型号'], 
                '订单数量': row['订单数量'],
                '月份': row['月份'],
                '目的地': row.get('目的地', ''),
                '客户交期': row.get('交期', ''),
                '欠缺物料编号': row['物料编号'],
                '欠缺物料名称': row['物料名称'],
                '欠料数量': row.get('仓存不足', 0),
                '采购单价(RMB)': row.get('RMB价格', 0),
                '欠料金额(RMB)': row.get('欠料金额', 0),
                '计算方式': row.get('计算方式', ''),
                'BOM编号': row.get('BOM编号', '')
            })
        
        report1_df = pd.DataFrame(report1_data)
        
        print(f"   📊 订单缺料明细: {len(report1_df)}条记录")
        print(f"   📊 涉及订单: {report1_df['生产订单号'].nunique()}个")
        print(f"   💰 总欠料金额: ¥{report1_df['欠料金额(RMB)'].sum():,.2f}")
        
        return report1_df
        
    def generate_report2_august_purchase_summary(self):
        """生成表2: 8月订单所需采购汇总"""
        print("=== 💰 生成表2: 8月订单采购汇总 ===")
        
        # 合并数据并筛选8月
        all_results = []
        if not self.merged_precise.empty:
            all_results.append(self.merged_precise)
        if not self.merged_estimated.empty:
            all_results.append(self.merged_estimated)
            
        if not all_results:
            print("❌ 无数据生成报表2")
            return pd.DataFrame()
        
        combined_df = pd.concat(all_results, ignore_index=True, sort=False)
        august_df = combined_df[combined_df['月份'] == '8月'].copy()
        
        if len(august_df) == 0:
            print("❌ 无8月数据")
            return pd.DataFrame()
        
        # 按订单汇总采购需求
        august_summary = august_df.groupby('生产单号').agg({
            '客户订单号': 'first',
            '产品型号': 'first',
            '订单数量': 'first',
            '目的地': 'first',
            '交期': 'first',
            '物料编号': lambda x: '; '.join(x.astype(str)),
            '物料名称': lambda x: '; '.join(x.astype(str)),
            '欠料金额': 'sum',
            '计算方式': 'first'
        }).reset_index()
        
        # 重命名和整理
        report2_data = []
        for _, row in august_summary.iterrows():
            report2_data.append({
                '生产订单号': row['生产单号'],
                '客户订单号': row['客户订单号'],
                '产品型号': row['产品型号'],
                '订单数量': row['订单数量'],
                '目的地': row['目的地'],
                '客户交期': row['交期'],
                '需采购物料': row['物料名称'],
                '采购总金额(RMB)': row['欠料金额'],
                '计算方式': row['计算方式']
            })
        
        report2_df = pd.DataFrame(report2_data)
        report2_df = report2_df.sort_values('采购总金额(RMB)', ascending=False)
        
        print(f"   📊 8月需采购订单: {len(report2_df)}个")
        print(f"   💰 8月采购总金额: ¥{report2_df['采购总金额(RMB)'].sum():,.2f}")
        
        return report2_df
        
    def generate_report3_supplier_summary(self):
        """生成表3: 按供应商汇总订单清单"""
        print("=== 🏭 生成表3: 供应商汇总 ===")
        
        # 从物料名称中提取供应商信息
        supplier_data = {}
        
        # 处理精确匹配的记录
        if not self.merged_precise.empty:
            for _, row in self.merged_precise.iterrows():
                material_name = str(row.get('物料名称', ''))
                
                # 从物料名称中提取供应商 (基于常见模式)
                suppliers = self.extract_supplier_from_material(material_name)
                
                for supplier in suppliers:
                    if supplier not in supplier_data:
                        supplier_data[supplier] = {
                            'orders': [],
                            'total_amount': 0,
                            'material_types': set()
                        }
                    
                    supplier_data[supplier]['orders'].append({
                        '生产订单号': row['生产单号'],
                        '产品型号': row['产品型号'],
                        '欠料金额': row.get('欠料金额', 0)
                    })
                    supplier_data[supplier]['total_amount'] += row.get('欠料金额', 0)
                    supplier_data[supplier]['material_types'].add(self.classify_material_type(material_name))
        
        # 处理估算记录 (按产品类型分类供应商)
        if not self.merged_estimated.empty:
            for _, row in self.merged_estimated.iterrows():
                supplier = f"估算供应商-{row['产品型号'][:6]}"
                
                if supplier not in supplier_data:
                    supplier_data[supplier] = {
                        'orders': [],
                        'total_amount': 0,
                        'material_types': set()
                    }
                
                supplier_data[supplier]['orders'].append({
                    '生产订单号': row['生产单号'],
                    '产品型号': row['产品型号'],
                    '欠料金额': row.get('欠料金额', 0)
                })
                supplier_data[supplier]['total_amount'] += row.get('欠料金额', 0)
                supplier_data[supplier]['material_types'].add('估算物料')
        
        # 生成供应商汇总报表
        report3_data = []
        for supplier, info in supplier_data.items():
            order_list = [f"{order['生产订单号']}({order['产品型号']})" for order in info['orders']]
            
            report3_data.append({
                '供应商': supplier,
                '物料类型': '; '.join(info['material_types']),
                '相关订单数': len(info['orders']),
                '相关订单列表': '; '.join(order_list),
                '采购总金额(RMB)': info['total_amount'],
                '平均单价': info['total_amount'] / len(info['orders']) if len(info['orders']) > 0 else 0
            })
        
        report3_df = pd.DataFrame(report3_data)
        report3_df = report3_df.sort_values('采购总金额(RMB)', ascending=False)
        
        print(f"   📊 涉及供应商: {len(report3_df)}家")
        if len(report3_df) > 0:
            print(f"   💰 总采购金额: ¥{report3_df['采购总金额(RMB)'].sum():,.2f}")
            print(f"   🔝 最大供应商: {report3_df.iloc[0]['供应商']} (¥{report3_df.iloc[0]['采购总金额(RMB)']:,.2f})")
        
        return report3_df
    
    def extract_supplier_from_material(self, material_name):
        """从物料名称中提取供应商"""
        suppliers = []
        material_str = str(material_name)
        
        # 常见供应商模式
        supplier_patterns = [
            r'友興邦', r'威達', r'萬至達', r'朗特', r'銘富通', 
            r'樂拓', r'科立', r'超盛', r'森和谷', r'瑞格',
            r'大昆輪', r'天赋利', r'轩泉', r'锋哲', r'華輝'
        ]
        
        for pattern in supplier_patterns:
            if re.search(pattern, material_str):
                suppliers.append(pattern)
        
        # 如果没找到，使用通用标识
        if not suppliers:
            suppliers = ['未知供应商']
            
        return suppliers
    
    def classify_material_type(self, material_name):
        """分类物料类型"""
        material_str = str(material_name).upper()
        
        if any(keyword in material_str for keyword in ['PCB', 'PC', '电路']):
            return 'PCB/电路板'
        elif any(keyword in material_str for keyword in ['PP', 'ABS', '塑料']):
            return '塑料件'
        elif any(keyword in material_str for keyword in ['SECC', 'SUS', '金属']):
            return '金属件'
        elif any(keyword in material_str for keyword in ['包装', '彩盒', '纸箱']):
            return '包装材料'
        else:
            return '其他'
    
    def save_final_reports(self, report1, report2, report3):
        """保存最终报表"""
        print("=== 💾 保存报表文件 ===")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f'精准订单物料分析报告_{timestamp}.xlsx'
        filepath = f'D:\\yingtu-PMC\\{filename}'
        
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                if not report1.empty:
                    report1.to_excel(writer, sheet_name='1_订单缺料明细', index=False)
                if not report2.empty:
                    report2.to_excel(writer, sheet_name='2_8月采购汇总', index=False)
                if not report3.empty:
                    report3.to_excel(writer, sheet_name='3_供应商汇总', index=False)
                    
                # 添加汇总统计sheet
                summary_data = {
                    '项目': ['总订单数', '有欠料订单数', '精确匹配订单', '估算订单', '8月需采购订单', '涉及供应商数', '总欠料金额(RMB)'],
                    '数量': [
                        len(self.orders_df),
                        len(report1) if not report1.empty else 0,
                        len(self.merged_precise) if not self.merged_precise.empty else 0,
                        len(self.merged_estimated) if not self.merged_estimated.empty else 0,
                        len(report2) if not report2.empty else 0,
                        len(report3) if not report3.empty else 0,
                        f"¥{report1['欠料金额(RMB)'].sum():,.2f}" if not report1.empty else "¥0.00"
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='0_汇总统计', index=False)
                
            print(f"✅ 报表已保存: {filename}")
            return filename
            
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return None
    
    def run_complete_analysis(self):
        """运行完整的混合方案分析"""
        print("🚀 开始精准订单物料分析 (混合方案)")
        print("="*80)
        
        try:
            # 1. 加载数据
            self.load_all_data()
            
            # 2. 精确匹配 (方案1)
            self.precise_matching()
            
            # 3. 估算补全 (方案2)  
            self.estimated_completion()
            
            # 4. 生成三个报表
            report1 = self.generate_report1_order_shortage_detail()
            report2 = self.generate_report2_august_purchase_summary()
            report3 = self.generate_report3_supplier_summary()
            
            # 5. 保存报表
            filename = self.save_final_reports(report1, report2, report3)
            
            # 6. 输出最终汇总
            print("\n" + "="*80)
            print(" "*25 + "🎉 分析完成！")
            print("="*80)
            
            total_orders = len(self.orders_df)
            precise_orders = len(self.merged_precise) if not self.merged_precise.empty else 0
            estimated_orders = len(self.merged_estimated) if not self.merged_estimated.empty else 0
            total_amount = report1['欠料金额(RMB)'].sum() if not report1.empty else 0
            
            print(f"📊 数据处理汇总:")
            print(f"   - 总订单数: {total_orders}个")
            print(f"   - 精确匹配: {precise_orders}条 ({precise_orders/total_orders*100:.1f}%)")
            print(f"   - 估算补全: {estimated_orders}条 ({estimated_orders/total_orders*100:.1f}%)")
            print(f"   - 总欠料金额: ¥{total_amount:,.2f}")
            
            if filename:
                print(f"📄 报表文件: {filename}")
                
            return report1, report2, report3
            
        except Exception as e:
            print(f"❌ 分析过程出错: {e}")
            import traceback
            traceback.print_exc()
            return None, None, None

if __name__ == "__main__":
    analyzer = PreciseOrderMaterialAnalyzer()
    analyzer.run_complete_analysis()
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基于订单清单的精准欠料分析系统
按照1.txt中560个订单的顺序进行欠料、供应商、金额分析
支持ROI优先级排序
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class OrderListShortageAnalyzer:
    def __init__(self):
        self.order_list = []             # 来自1.txt的订单列表
        self.orders_df = None            # 订单数据（4个sheet合并）
        self.shortage_df = None          # 欠料数据
        self.inventory_df = None         # 库存价格数据
        self.supplier_df = None          # 供应商数据
        self.final_result = None         # 最终分析结果
        
        # 汇率设置（转换为RMB）
        self.currency_rates = {
            'RMB': 1.0,
            'USD': 7.30,
            'HKD': 0.93,
            'EUR': 7.85
        }
    
    def load_order_list(self, order_file='1.txt'):
        """加载订单清单"""
        print("=== 🔄 加载订单清单 ===")
        
        try:
            with open(order_file, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            # 提取订单号（去除空行和编号）
            orders = []
            for line in lines:
                order = line.strip()
                if order and (order.startswith('PSO') or order.startswith('RSO') or 
                             order.startswith('MSO') or order.startswith('TSO') or 
                             order.startswith('P-R')):
                    orders.append(order)
            
            self.order_list = orders
            print(f"   ✅ 成功加载订单清单: {len(self.order_list)}个订单")
            
            # 统计订单类型
            pso_count = len([o for o in orders if o.startswith('PSO')])
            rso_count = len([o for o in orders if o.startswith('RSO')])
            mso_count = len([o for o in orders if o.startswith('MSO')])
            tso_count = len([o for o in orders if o.startswith('TSO')])
            pr_count = len([o for o in orders if o.startswith('P-R')])
            
            print(f"   📊 订单类型分布: PSO:{pso_count}, RSO:{rso_count}, MSO:{mso_count}, TSO:{tso_count}, P-R:{pr_count}")
            
            # 显示前5个和后5个订单
            print(f"   🔍 订单范围: {orders[0]} 至 {orders[-1]}")
            print("   📋 前5个订单:", orders[:5])
            print("   📋 后5个订单:", orders[-5:])
            
            return True
            
        except Exception as e:
            print(f"   ❌ 订单清单加载失败: {e}")
            return False
    
    def load_all_data(self):
        """加载所有数据源（基于silverPlan_analysis.py的逻辑）"""
        print("=== 🔄 加载所有数据源 ===")
        
        # 1. 加载订单数据（4个工作表）
        print("1. 加载订单数据...")
        try:
            orders_data = []
            
            # 国内订单
            orders_aug_domestic = pd.read_excel('input/order-amt-89.xlsx', sheet_name='8月')
            orders_sep_domestic = pd.read_excel('input/order-amt-89.xlsx', sheet_name='9月')
            orders_aug_domestic['月份'] = '8月'
            orders_aug_domestic['数据来源工作表'] = '国内'
            orders_sep_domestic['月份'] = '9月'
            orders_sep_domestic['数据来源工作表'] = '国内'
            orders_data.extend([orders_aug_domestic, orders_sep_domestic])
            
            # 柬埔寨订单
            orders_aug_cambodia = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name='8月 -柬')
            orders_sep_cambodia = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name='9月 -柬')
            orders_aug_cambodia['月份'] = '8月'
            orders_aug_cambodia['数据来源工作表'] = '柬埔寨'
            orders_sep_cambodia['月份'] = '9月'
            orders_sep_cambodia['数据来源工作表'] = '柬埔寨'
            orders_data.extend([orders_aug_cambodia, orders_sep_cambodia])
            
            # 合并所有订单
            self.orders_df = pd.concat(orders_data, ignore_index=True)
            
            # 标准化列名
            self.orders_df = self.orders_df.rename(columns={
                '生 產 單 号(  廠方 )': '生产单号',
                '生 產 單 号(客方 )': '客户订单号',
                '型 號( 廠方/客方 )': '产品型号',
                '數 量  (Pcs)': '数量Pcs',
                'BOM NO.': 'BOM编号',
                '客期': '客户交期'
            })
            
            # 确保订单金额字段存在
            if '订单金额' not in self.orders_df.columns:
                self.orders_df['订单金额'] = 1000  # 默认值
            
            print(f"   ✅ 订单数据: {len(self.orders_df)}条")
            
        except Exception as e:
            print(f"   ❌ 订单数据加载失败: {e}")
            return False
        
        # 2. 加载欠料数据
        print("2. 加载欠料数据...")
        try:
            self.shortage_df = pd.read_excel('input/mat_owe_pso.xlsx', sheet_name='Sheet1', skiprows=1)
            
            # 标准化列名
            if len(self.shortage_df.columns) >= 13:
                new_columns = ['订单编号', 'P-R对应', 'P-RBOM', '客户型号', 'OTS期', '开拉期', 
                              '下单日期', '物料编号', '物料名称', '领用部门', '工单需求', 
                              '仓存不足', '已购未返', '手头现有', '请购组']
                
                for i in range(min(len(new_columns), len(self.shortage_df.columns))):
                    if i < len(self.shortage_df.columns):
                        self.shortage_df.rename(columns={self.shortage_df.columns[i]: new_columns[i]}, inplace=True)
            
            # 清理欠料数据
            self.shortage_df = self.shortage_df.dropna(subset=['订单编号'])
            self.shortage_df = self.shortage_df[~self.shortage_df['物料名称'].astype(str).str.contains('已齐套|齐套', na=False)]
            
            print(f"   ✅ 欠料数据: {len(self.shortage_df)}条")
            
        except Exception as e:
            print(f"   ❌ 欠料数据加载失败: {e}")
            self.shortage_df = pd.DataFrame()
        
        # 3. 加载库存价格数据
        print("3. 加载库存价格数据...")
        try:
            self.inventory_df = pd.read_excel('input/inventory_list.xlsx')
            
            # 价格处理
            self.inventory_df['最终价格'] = self.inventory_df['最新報價'].fillna(self.inventory_df['成本單價'])
            self.inventory_df['最终价格'] = pd.to_numeric(self.inventory_df['最终价格'], errors='coerce').fillna(0)
            
            # 货币转换
            def convert_to_rmb(row):
                price = row['最终价格']
                currency = str(row.get('貨幣', 'RMB')).upper()
                rate = self.currency_rates.get(currency, 1.0)
                return price * rate if pd.notna(price) else 0
            
            self.inventory_df['RMB单价'] = self.inventory_df.apply(convert_to_rmb, axis=1)
            
            print(f"   ✅ 库存数据: {len(self.inventory_df)}条")
            
        except Exception as e:
            print(f"   ❌ 库存数据加载失败: {e}")
            self.inventory_df = pd.DataFrame()
        
        # 4. 加载供应商数据
        print("4. 加载供应商数据...")
        try:
            self.supplier_df = pd.read_excel('input/supplier.xlsx')
            
            # 处理供应商价格
            self.supplier_df['单价_数值'] = pd.to_numeric(self.supplier_df['单价'], errors='coerce').fillna(0)
            
            def convert_supplier_to_rmb(row):
                price = row['单价_数值']
                currency = str(row.get('币种', 'RMB')).upper()
                rate = self.currency_rates.get(currency, 1.0)
                return price * rate if pd.notna(price) else 0
            
            self.supplier_df['供应商RMB单价'] = self.supplier_df.apply(convert_supplier_to_rmb, axis=1)
            self.supplier_df['修改日期'] = pd.to_datetime(self.supplier_df['修改日期'], errors='coerce')
            
            print(f"   ✅ 供应商数据: {len(self.supplier_df)}条")
            
        except Exception as e:
            print(f"   ❌ 供应商数据加载失败: {e}")
            self.supplier_df = pd.DataFrame()
        
        print("✅ 所有数据源加载完成\n")
        return True
    
    def select_best_supplier(self, material_suppliers):
        """选择最优供应商（基于最低价格）"""
        if len(material_suppliers) == 0:
            return None
        if len(material_suppliers) == 1:
            return material_suppliers.iloc[0]
        
        # 筛选有有效价格的供应商
        valid_suppliers = material_suppliers[material_suppliers['供应商RMB单价'] > 0]
        
        if len(valid_suppliers) == 0:
            return material_suppliers.iloc[0]
        
        # 选择最低价供应商
        lowest_price_idx = valid_suppliers['供应商RMB单价'].idxmin()
        return valid_suppliers.loc[lowest_price_idx]
    
    def analyze_order_list_shortage(self):
        """按照1.txt订单顺序分析欠料情况"""
        print("=== 🎯 按订单清单分析欠料情况 ===")
        
        if not self.order_list or not hasattr(self, 'shortage_df'):
            print("❌ 数据未准备完成")
            return False
        
        analysis_results = []
        matched_count = 0
        unmatched_count = 0
        pr_orders = []
        
        print(f"开始分析{len(self.order_list)}个订单...")
        
        for i, order_no in enumerate(self.order_list):
            if (i + 1) % 50 == 0:
                print(f"   处理进度: {i + 1}/{len(self.order_list)}")
            
            # 记录P-R订单
            if order_no.startswith('P-R'):
                pr_orders.append(order_no)
            
            # 1. 查找订单基本信息
            order_info = self.orders_df[self.orders_df['生产单号'] == order_no]
            
            if not order_info.empty:
                order_base = order_info.iloc[0]
                product_model = order_base.get('产品型号', '')
                quantity = order_base.get('数量Pcs', 0)
                customer_order = order_base.get('客户订单号', '')
                order_amount_usd = pd.to_numeric(order_base.get('订单金额', 0), errors='coerce')
                order_amount_rmb = order_amount_usd * self.currency_rates['USD']
                month = order_base.get('月份', '')
                source = order_base.get('数据来源工作表', '')
                bom_no = order_base.get('BOM编号', '')
                delivery_date = order_base.get('客户交期', '')
            else:
                # 订单基本信息不存在的情况
                product_model = ''
                quantity = 0
                customer_order = ''
                order_amount_usd = 0
                order_amount_rmb = 0
                month = ''
                source = ''
                bom_no = ''
                delivery_date = ''
            
            # 2. 查找欠料信息
            shortage_records = self.shortage_df[self.shortage_df['订单编号'].astype(str).str.strip() == order_no]
            
            if not shortage_records.empty:
                matched_count += 1
                
                # 处理每个欠料物料
                for _, shortage_row in shortage_records.iterrows():
                    material_code = shortage_row.get('物料编号', '')
                    material_name = shortage_row.get('物料名称', '')
                    shortage_qty = pd.to_numeric(shortage_row.get('仓存不足', 0), errors='coerce')
                    
                    # 3. 查找库存价格
                    inventory_info = self.inventory_df[self.inventory_df['物項編號'] == material_code]
                    if not inventory_info.empty:
                        rmb_price = inventory_info.iloc[0].get('RMB单价', 0)
                        currency = inventory_info.iloc[0].get('貨幣', 'RMB')
                        original_price = inventory_info.iloc[0].get('最终价格', 0)
                    else:
                        rmb_price = 0
                        currency = 'RMB'
                        original_price = 0
                    
                    # 4. 查找最优供应商
                    supplier_records = self.supplier_df[self.supplier_df['物项编号'] == material_code]
                    if not supplier_records.empty:
                        best_supplier = self.select_best_supplier(supplier_records)
                        supplier_name = best_supplier.get('供应商名称', '')
                        supplier_code = best_supplier.get('供应商号', '')
                        supplier_price = best_supplier.get('单价', 0)
                        supplier_currency = best_supplier.get('币种', 'RMB')
                        min_order_qty = best_supplier.get('起订数量', 0)
                        modify_date = best_supplier.get('修改日期', '')
                    else:
                        supplier_name = ''
                        supplier_code = ''
                        supplier_price = 0
                        supplier_currency = 'RMB'
                        min_order_qty = 0
                        modify_date = ''
                    
                    # 5. 计算欠料金额
                    shortage_amount_rmb = shortage_qty * rmb_price
                    
                    # 6. 创建记录
                    record = {
                        '清单序号': i + 1,
                        '生产订单号': order_no,
                        '客户订单号': customer_order,
                        '产品型号': product_model,
                        '数量Pcs': quantity,
                        '月份': month,
                        '数据来源': source,
                        'BOM编号': bom_no,
                        '客户交期': delivery_date,
                        
                        '欠料物料编号': material_code,
                        '欠料物料名称': material_name,
                        '欠料数量': shortage_qty,
                        
                        '库存RMB单价': rmb_price,
                        '库存原始价格': original_price,
                        '库存币种': currency,
                        
                        '主供应商名称': supplier_name,
                        '主供应商号': supplier_code,
                        '供应商单价(原币)': supplier_price,
                        '供应商币种': supplier_currency,
                        '起订数量': min_order_qty,
                        '供应商修改日期': modify_date,
                        
                        '欠料金额(RMB)': shortage_amount_rmb,
                        '订单金额(USD)': order_amount_usd,
                        '订单金额(RMB)': order_amount_rmb,
                        
                        '工单需求': shortage_row.get('工单需求', ''),
                        '已购未返': shortage_row.get('已购未返', ''),
                        '手头现有': shortage_row.get('手头现有', ''),
                        '请购组': shortage_row.get('请购组', ''),
                        
                        '匹配状态': '有欠料'
                    }
                    
                    analysis_results.append(record)
            
            else:
                # 无欠料记录的订单
                unmatched_count += 1
                record = {
                    '清单序号': i + 1,
                    '生产订单号': order_no,
                    '客户订单号': customer_order,
                    '产品型号': product_model,
                    '数量Pcs': quantity,
                    '月份': month,
                    '数据来源': source,
                    'BOM编号': bom_no,
                    '客户交期': delivery_date,
                    
                    '欠料物料编号': '',
                    '欠料物料名称': '',
                    '欠料数量': 0,
                    
                    '库存RMB单价': 0,
                    '库存原始价格': 0,
                    '库存币种': 'RMB',
                    
                    '主供应商名称': '',
                    '主供应商号': '',
                    '供应商单价(原币)': 0,
                    '供应商币种': 'RMB',
                    '起订数量': 0,
                    '供应商修改日期': '',
                    
                    '欠料金额(RMB)': 0,
                    '订单金额(USD)': order_amount_usd,
                    '订单金额(RMB)': order_amount_rmb,
                    
                    '工单需求': '',
                    '已购未返': '',
                    '手头现有': '',
                    '请购组': '',
                    
                    '匹配状态': '不缺料' if not order_info.empty else '未找到订单信息'
                }
                
                analysis_results.append(record)
        
        self.final_result = pd.DataFrame(analysis_results)
        
        print(f"   ✅ 分析完成:")
        print(f"      - 有欠料记录订单: {matched_count}个")
        print(f"      - 无欠料记录订单: {unmatched_count}个")
        print(f"      - P-R订单数量: {len(pr_orders)}个")
        print(f"      - 总分析记录: {len(analysis_results)}条")
        
        if pr_orders:
            print(f"   🔍 P-R订单列表: {pr_orders}")
        
        return True
    
    def calculate_roi_and_priority(self):
        """计算ROI和优先级排序"""
        print("=== 💰 计算ROI和优先级 ===")
        
        if self.final_result is None:
            print("❌ 没有分析结果数据")
            return False
        
        # 1. 按订单汇总计算ROI
        print("1. 按订单汇总计算ROI...")
        
        order_summary = self.final_result.groupby('生产订单号').agg({
            '清单序号': 'first',
            '客户订单号': 'first',
            '产品型号': 'first',
            '数量Pcs': 'first',
            '月份': 'first',
            '数据来源': 'first',
            '订单金额(USD)': 'first',
            '订单金额(RMB)': 'first',
            '欠料金额(RMB)': 'sum',  # 汇总所有物料的欠料金额
            '欠料物料编号': 'count',  # 计算欠料物料种类数
            '匹配状态': 'first'
        }).reset_index()
        
        order_summary.rename(columns={'欠料物料编号': '欠料物料种类'}, inplace=True)
        
        # 修正欠料物料种类计算（排除空值）
        for idx, row in order_summary.iterrows():
            order_records = self.final_result[self.final_result['生产订单号'] == row['生产订单号']]
            actual_material_count = len(order_records[order_records['欠料物料编号'] != ''])
            order_summary.at[idx, '欠料物料种类'] = actual_material_count
        
        # 2. 计算ROI（订单金额/欠料金额）
        def calculate_roi(row):
            order_amount = row['订单金额(RMB)']
            shortage_amount = row['欠料金额(RMB)']
            
            if shortage_amount > 0:
                return order_amount / shortage_amount
            elif order_amount > 0:
                return 999999  # 无需投入的订单设为极高ROI
            else:
                return 0
        
        order_summary['ROI'] = order_summary.apply(calculate_roi, axis=1)
        
        # 3. 确定生产优先级
        def get_priority_level(roi):
            if roi >= 999999:
                return "无需投入-立即生产"
            elif roi >= 5.0:
                return "高优先级"
            elif roi >= 2.0:
                return "中优先级"
            elif roi >= 1.0:
                return "低优先级"
            else:
                return "暂缓生产"
        
        order_summary['优先级等级'] = order_summary['ROI'].apply(get_priority_level)
        
        # 4. 按ROI排序（降序）
        priority_sorted = order_summary.sort_values('ROI', ascending=False).reset_index(drop=True)
        priority_sorted['建议生产顺序'] = range(1, len(priority_sorted) + 1)
        
        # 统计优先级分布
        priority_stats = order_summary['优先级等级'].value_counts()
        print(f"   📊 优先级分布: {dict(priority_stats)}")
        
        total_orders = len(order_summary)
        immediate_production = len(order_summary[order_summary['ROI'] >= 999999])
        high_priority = len(order_summary[(order_summary['ROI'] >= 5.0) & (order_summary['ROI'] < 999999)])
        
        print(f"   🚀 可立即生产: {immediate_production}个 ({immediate_production/total_orders*100:.1f}%)")
        print(f"   ⭐ 高优先级: {high_priority}个 ({high_priority/total_orders*100:.1f}%)")
        
        # 保存汇总结果
        self.order_summary = order_summary
        self.priority_sorted = priority_sorted
        
        return True
    
    def generate_summary_reports(self):
        """生成汇总报告"""
        print("=== 📊 生成汇总报告 ===")
        
        # 1. 供应商汇总
        shortage_records = self.final_result[
            (self.final_result['欠料物料编号'] != '') & 
            (self.final_result['主供应商名称'] != '')
        ]
        
        if not shortage_records.empty:
            supplier_summary = shortage_records.groupby('主供应商名称').agg({
                '欠料金额(RMB)': 'sum',
                '欠料物料编号': 'nunique',
                '生产订单号': 'nunique'
            }).reset_index()
            
            supplier_summary.rename(columns={
                '欠料物料编号': '涉及物料种类',
                '生产订单号': '影响订单数'
            }, inplace=True)
            
            supplier_summary = supplier_summary.sort_values('欠料金额(RMB)', ascending=False)
        else:
            supplier_summary = pd.DataFrame()
        
        # 2. 物料汇总（采购清单）
        if not shortage_records.empty:
            material_summary = shortage_records.groupby(['欠料物料编号', '欠料物料名称']).agg({
                '欠料数量': 'sum',
                '主供应商名称': 'first',
                '供应商单价(原币)': 'first',
                '供应商币种': 'first',
                '库存RMB单价': 'first',
                '欠料金额(RMB)': 'sum',
                '生产订单号': lambda x: ', '.join(sorted(x.unique())[:5])  # 显示前5个相关订单
            }).reset_index()
            
            material_summary = material_summary.sort_values('欠料金额(RMB)', ascending=False)
            material_summary.rename(columns={'生产订单号': '相关订单(前5个)'}, inplace=True)
        else:
            material_summary = pd.DataFrame()
        
        self.supplier_summary = supplier_summary
        self.material_summary = material_summary
        
        # 打印统计信息
        total_shortage_amount = self.final_result['欠料金额(RMB)'].sum()
        total_order_amount = self.order_summary['订单金额(RMB)'].sum()
        
        print(f"   💰 总欠料金额: ¥{total_shortage_amount:,.2f}")
        print(f"   💰 总订单金额: ¥{total_order_amount:,.2f}")
        if total_shortage_amount > 0:
            overall_roi = total_order_amount / total_shortage_amount
            print(f"   📈 整体ROI: {overall_roi:.2f}倍")
        
        if not supplier_summary.empty:
            print(f"   🏭 涉及供应商: {len(supplier_summary)}家")
            print(f"   📦 需采购物料: {supplier_summary['涉及物料种类'].sum()}种")
        
        return True
    
    def save_comprehensive_report(self):
        """保存综合分析报告"""
        print("=== 💾 保存综合分析报告 ===")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'560订单清单欠料分析报告_{timestamp}.xlsx'
        
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # 1. 主报表：按1.txt顺序的详细分析
                self.final_result.to_excel(writer, sheet_name='按清单顺序详细分析', index=False)
                
                # 2. 按订单汇总（包含ROI）
                self.order_summary.to_excel(writer, sheet_name='按订单汇总(含ROI)', index=False)
                
                # 3. ROI优先级排序
                self.priority_sorted.to_excel(writer, sheet_name='ROI优先级排序', index=False)
                
                # 4. 供应商汇总
                if not self.supplier_summary.empty:
                    self.supplier_summary.to_excel(writer, sheet_name='按供应商汇总', index=False)
                
                # 5. 采购物料清单
                if not self.material_summary.empty:
                    self.material_summary.to_excel(writer, sheet_name='采购物料清单', index=False)
                
                # 6. 统计汇总
                stats_data = {
                    '统计项目': [
                        '订单清单总数',
                        '有欠料订单数',
                        '无需投入订单数',
                        'P-R订单数',
                        '总欠料金额(RMB)',
                        '总订单金额(RMB)',
                        '整体ROI',
                        '涉及供应商数',
                        '需采购物料种类',
                        '高优先级订单数(ROI≥5)',
                        '中优先级订单数(2≤ROI<5)',
                        '低优先级订单数(1≤ROI<2)'
                    ],
                    '数值': [
                        len(self.order_list),
                        len(self.order_summary[self.order_summary['欠料金额(RMB)'] > 0]),
                        len(self.order_summary[self.order_summary['ROI'] >= 999999]),
                        len([o for o in self.order_list if o.startswith('P-R')]),
                        f"¥{self.final_result['欠料金额(RMB)'].sum():,.2f}",
                        f"¥{self.order_summary['订单金额(RMB)'].sum():,.2f}",
                        f"{self.order_summary['订单金额(RMB)'].sum() / self.final_result['欠料金额(RMB)'].sum():.2f}倍" 
                        if self.final_result['欠料金额(RMB)'].sum() > 0 else "无需投入",
                        len(self.supplier_summary) if not self.supplier_summary.empty else 0,
                        self.supplier_summary['涉及物料种类'].sum() if not self.supplier_summary.empty else 0,
                        len(self.order_summary[(self.order_summary['ROI'] >= 5.0) & (self.order_summary['ROI'] < 999999)]),
                        len(self.order_summary[(self.order_summary['ROI'] >= 2.0) & (self.order_summary['ROI'] < 5.0)]),
                        len(self.order_summary[(self.order_summary['ROI'] >= 1.0) & (self.order_summary['ROI'] < 2.0)])
                    ]
                }
                
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='统计汇总', index=False)
            
            print(f"✅ 综合报告已保存: {filename}")
            print(f"📋 包含工作表:")
            print(f"   1️⃣ 按清单顺序详细分析 ({len(self.final_result)}条记录)")
            print(f"   2️⃣ 按订单汇总(含ROI) ({len(self.order_summary)}个订单)")
            print(f"   3️⃣ ROI优先级排序 (生产建议顺序)")
            print(f"   4️⃣ 按供应商汇总 (采购部门使用)")
            print(f"   5️⃣ 采购物料清单 (具体采购清单)")
            print(f"   6️⃣ 统计汇总 (管理层概览)")
            
            return filename
            
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return None
    
    def run_comprehensive_analysis(self):
        """运行完整的综合分析"""
        print("🚀 开始560订单清单欠料分析")
        print("="*70)
        
        try:
            # 1. 加载订单清单
            if not self.load_order_list():
                return None
            
            # 2. 加载所有数据源
            if not self.load_all_data():
                return None
            
            # 3. 按订单清单分析欠料
            if not self.analyze_order_list_shortage():
                return None
            
            # 4. 计算ROI和优先级
            if not self.calculate_roi_and_priority():
                return None
            
            # 5. 生成汇总报告
            if not self.generate_summary_reports():
                return None
            
            # 6. 保存报告
            filename = self.save_comprehensive_report()
            if not filename:
                return None
            
            # 7. 输出最终汇总
            print("\n" + "="*70)
            print(" "*20 + "🎉 分析完成！")
            print("="*70)
            
            # 显示关键统计
            immediate_orders = len(self.order_summary[self.order_summary['ROI'] >= 999999])
            high_priority = len(self.order_summary[(self.order_summary['ROI'] >= 5.0) & (self.order_summary['ROI'] < 999999)])
            total_shortage = self.final_result['欠料金额(RMB)'].sum()
            total_order_value = self.order_summary['订单金额(RMB)'].sum()
            
            print(f"📊 关键结果:")
            print(f"   - 分析订单总数: {len(self.order_list)}个")
            print(f"   - 可立即生产: {immediate_orders}个 ({immediate_orders/len(self.order_list)*100:.1f}%)")
            print(f"   - 高优先级(ROI≥5): {high_priority}个")
            print(f"   - 总欠料金额: ¥{total_shortage:,.2f}")
            print(f"   - 总订单价值: ¥{total_order_value:,.2f}")
            print(f"   - 整体ROI: {total_order_value/total_shortage:.2f}倍" if total_shortage > 0 else "   - 整体ROI: 无需投入")
            
            # 显示top5高价值欠料订单
            top5_shortage = self.order_summary.nlargest(5, '欠料金额(RMB)')
            print(f"\n🔝 欠料金额最高的5个订单:")
            for _, order in top5_shortage.iterrows():
                if order['欠料金额(RMB)'] > 0:
                    roi_text = f"{order['ROI']:.2f}倍" if order['ROI'] < 999999 else "无需投入"
                    print(f"   {order['生产订单号']}: {order['产品型号']} (欠料¥{order['欠料金额(RMB)']:,.2f}, ROI {roi_text})")
            
            print(f"\n📄 已生成报告: {filename}")
            
            return filename
            
        except Exception as e:
            print(f"❌ 分析过程出错: {e}")
            import traceback
            traceback.print_exc()
            return None

if __name__ == "__main__":
    analyzer = OrderListShortageAnalyzer()
    result = analyzer.run_comprehensive_analysis()
    
    if result:
        print(f"\n🎊 分析成功完成！报告文件: {result}")
    else:
        print("\n❌ 分析失败，请检查数据和错误信息")
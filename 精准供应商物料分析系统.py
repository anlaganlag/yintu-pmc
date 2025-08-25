#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图工厂精准供应商物料分析系统 (最终版本)
基于订单表 + 欠料表 + 供应商表的完整匹配分析
实现混合供应商选择方案 (A+B)
"""

import pandas as pd
import numpy as np
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class FinalSupplierMaterialAnalyzer:
    def __init__(self):
        self.orders_df = None           # 订单数据
        self.shortage_df = None         # 欠料数据  
        self.supplier_df = None         # 供应商数据
        self.merged_precise = None      # 精确匹配结果
        self.merged_estimated = None    # 估算补全结果
        self.multi_supplier_df = None   # 多供应商选择表
        
        self.currency_rates = {         # 汇率 (转换为RMB)
            'RMB': 1.0,
            'USD': 7.20,  # 1 USD = 7.20 RMB  
            'HKD': 0.93,  # 1 HKD = 0.93 RMB
            'EUR': 7.85   # 1 EUR = 7.85 RMB
        }
        
    def load_all_data(self):
        """加载所有数据源"""
        print("=== 🔄 加载数据源 ===")
        
        # 1. 加载全部订单数据
        print("1. 加载(全部)8月9月订单...")
        orders_aug = pd.read_excel(r'D:\yingtu-PMC\(全部)8月9月订单.xlsx', sheet_name='8月')
        orders_sep = pd.read_excel(r'D:\yingtu-PMC\(全部)8月9月订单.xlsx', sheet_name='9月')
        
        # 方案1: 基于订单工作表名称确定月份
        orders_aug['月份'] = '8月'
        orders_aug['数据来源工作表'] = '8月'
        orders_sep['月份'] = '9月'  
        orders_sep['数据来源工作表'] = '9月'
        
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
        
        # 验证月份字段的正确性
        month_distribution = self.orders_df['月份'].value_counts()
        print(f"   📊 月份分布验证: {dict(month_distribution)}")
        
        # 验证数据来源工作表
        source_distribution = self.orders_df['数据来源工作表'].value_counts()
        print(f"   📋 数据来源验证: {dict(source_distribution)}")
        
        # 2. 加载新的欠料表
        print("2. 加载mat_owe_pso.xlsx欠料表...")
        try:
            # 读取第一个sheet，跳过表头
            self.shortage_df = pd.read_excel(r'D:\yingtu-PMC\mat_owe_pso.xlsx', 
                                           sheet_name='Sheet1', skiprows=1)
            
            # 标准化列名 (基于分析结果)
            if len(self.shortage_df.columns) >= 13:
                new_columns = ['订单编号', 'P-R对应', 'P-RBOM', '客户型号', 'OTS期', '开拉期', 
                              '下单日期', '物料编号', '物料名称', '领用部门', '工单需求', 
                              '仓存不足', '已购未返', '手头现有', '请购组']
                
                # 确保列数匹配
                for i in range(min(len(new_columns), len(self.shortage_df.columns))):
                    if i < len(self.shortage_df.columns):
                        self.shortage_df.rename(columns={self.shortage_df.columns[i]: new_columns[i]}, inplace=True)
            
            # 清理数据
            self.shortage_df = self.shortage_df.dropna(subset=['订单编号'])
            # 移除包含"已齐套"的记录
            self.shortage_df = self.shortage_df[~self.shortage_df['物料名称'].astype(str).str.contains('已齐套|齐套', na=False)]
            
            print(f"   ✅ 欠料记录: {len(self.shortage_df)}条")
            
        except Exception as e:
            print(f"   ❌ 欠料表加载失败: {e}")
            self.shortage_df = pd.DataFrame()
        
        # 3. 加载供应商表
        print("3. 加载supplier.xlsx供应商表...")
        try:
            self.supplier_df = pd.read_excel(r'D:\yingtu-PMC\supplier.xlsx')
            
            # 处理价格和货币转换
            self.supplier_df['单价_数值'] = pd.to_numeric(self.supplier_df['单价'], errors='coerce').fillna(0)
            
            # 货币转换为RMB
            def convert_to_rmb(row):
                price = row['单价_数值']
                currency = str(row.get('币种', 'RMB')).upper()
                rate = self.currency_rates.get(currency, 1.0)
                return price * rate
            
            self.supplier_df['RMB单价'] = self.supplier_df.apply(convert_to_rmb, axis=1)
            
            # 处理修改日期
            self.supplier_df['修改日期'] = pd.to_datetime(self.supplier_df['修改日期'], errors='coerce')
            
            valid_prices = len(self.supplier_df[self.supplier_df['RMB单价'] > 0])
            print(f"   ✅ 供应商记录: {len(self.supplier_df)}条, 有效价格: {valid_prices}条")
            print(f"   ✅ 唯一供应商: {self.supplier_df['供应商名称'].nunique()}家")
            print(f"   ✅ 唯一物料: {self.supplier_df['物项编号'].nunique()}个")
            
        except Exception as e:
            print(f"   ❌ 供应商表加载失败: {e}")
            self.supplier_df = pd.DataFrame()
        
        print("✅ 数据加载完成\n")
        
    def select_primary_supplier(self, material_suppliers):
        """为物料选择主供应商 - 方案A算法"""
        if len(material_suppliers) == 0:
            return None
        
        if len(material_suppliers) == 1:
            return material_suppliers.iloc[0]
        
        # 多供应商选择逻辑
        scored_suppliers = material_suppliers.copy()
        scored_suppliers['综合得分'] = 0
        
        # 1. 最新修改日期得分 (40分)
        if scored_suppliers['修改日期'].notna().sum() > 0:
            latest_date = scored_suppliers['修改日期'].max()
            scored_suppliers.loc[scored_suppliers['修改日期'] == latest_date, '综合得分'] += 40
            
            # 其他日期按比例给分
            date_scores = scored_suppliers['修改日期'].apply(
                lambda x: 30 if pd.notna(x) and (latest_date - x).days <= 365 else 20 if pd.notna(x) else 0
            )
            scored_suppliers['综合得分'] += date_scores
        
        # 2. 最低价格得分 (35分)
        if (scored_suppliers['RMB单价'] > 0).sum() > 0:
            min_price = scored_suppliers[scored_suppliers['RMB单价'] > 0]['RMB单价'].min()
            scored_suppliers.loc[scored_suppliers['RMB单价'] == min_price, '综合得分'] += 35
            
            # 其他价格按比例给分
            price_scores = scored_suppliers['RMB单价'].apply(
                lambda x: 25 if x > 0 and x <= min_price * 1.1 else 15 if x > 0 and x <= min_price * 1.2 else 5 if x > 0 else 0
            )
            scored_suppliers['综合得分'] += price_scores
        
        # 3. 供应商稳定性得分 (25分)
        # 供应商号越小越稳定 (简单假设)
        scored_suppliers['供应商号_数值'] = pd.to_numeric(scored_suppliers['供应商号'], errors='coerce')
        if scored_suppliers['供应商号_数值'].notna().sum() > 0:
            min_supplier_no = scored_suppliers['供应商号_数值'].min()
            scored_suppliers.loc[scored_suppliers['供应商号_数值'] == min_supplier_no, '综合得分'] += 25
        
        # 选择得分最高的供应商
        best_supplier = scored_suppliers.loc[scored_suppliers['综合得分'].idxmax()]
        return best_supplier
        
    def precise_matching_with_supplier(self):
        """精确匹配 - 订单→欠料→供应商"""
        print("=== 🎯 精确匹配分析 (含供应商选择) ===")
        
        if self.shortage_df.empty or self.supplier_df.empty:
            print("❌ 欠料表或供应商表为空，跳过精确匹配")
            self.merged_precise = pd.DataFrame()
            return
        
        # 1. 订单与欠料匹配
        print("1. 订单与欠料表匹配...")
        self.orders_df['生产单号_清理'] = self.orders_df['生产单号'].astype(str).str.strip()
        self.shortage_df['订单编号_清理'] = self.shortage_df['订单编号'].astype(str).str.strip()
        
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
        
        # 2. 为每个物料选择主供应商
        print("2. 为每个欠料物料选择主供应商...")
        precise_with_supplier = []
        multi_supplier_records = []  # 记录多供应商物料
        
        unique_materials = precise_records['物料编号'].unique()
        processed_materials = 0
        
        for material_code in unique_materials:
            # 获取该物料的所有记录
            material_records = precise_records[precise_records['物料编号'] == material_code]
            
            # 获取该物料的所有供应商
            material_suppliers = self.supplier_df[self.supplier_df['物项编号'] == material_code]
            
            if len(material_suppliers) == 0:
                # 无供应商信息，标记为待查找
                for _, record in material_records.iterrows():
                    record_dict = record.to_dict()
                    record_dict.update({
                        '主供应商名称': '未找到供应商',
                        '主供应商号': '',
                        '主供应商单价': 0,
                        '主供应商币种': '',
                        'RMB单价': 0,
                        '修改日期': pd.NaT,
                        '起订数量': 0,
                        '计算方式': '无供应商信息'
                    })
                    precise_with_supplier.append(record_dict)
            else:
                # 有供应商信息
                # 记录多供应商情况
                if len(material_suppliers) > 1:
                    for _, supplier in material_suppliers.iterrows():
                        multi_supplier_records.append({
                            '物料编号': material_code,
                            '物料名称': material_records.iloc[0]['物料名称'],
                            '供应商名称': supplier['供应商名称'],
                            '供应商号': supplier['供应商号'],
                            'RMB单价': supplier['RMB单价'],
                            '原币单价': supplier['单价'],
                            '币种': supplier['币种'],
                            '起订数量': supplier['起订数量'],
                            '修改日期': supplier['修改日期'],
                            '是否主选': False
                        })
                
                # 选择主供应商
                primary_supplier = self.select_primary_supplier(material_suppliers)
                
                # 标记主选供应商
                if len(material_suppliers) > 1:
                    for record in multi_supplier_records:
                        if (record['物料编号'] == material_code and 
                            record['供应商号'] == primary_supplier['供应商号']):
                            record['是否主选'] = True
                
                # 为该物料的所有订单记录添加主供应商信息
                for _, record in material_records.iterrows():
                    record_dict = record.to_dict()
                    record_dict.update({
                        '主供应商名称': primary_supplier['供应商名称'],
                        '主供应商号': primary_supplier['供应商号'],
                        '主供应商单价': primary_supplier['单价'],
                        '主供应商币种': primary_supplier['币种'],
                        'RMB单价': primary_supplier['RMB单价'],
                        '修改日期': primary_supplier['修改日期'],
                        '起订数量': primary_supplier['起订数量'],
                        '计算方式': '精确匹配'
                    })
                    precise_with_supplier.append(record_dict)
            
            processed_materials += 1
            if processed_materials % 100 == 0:
                print(f"   处理进度: {processed_materials}/{len(unique_materials)} 物料")
        
        # 3. 计算欠料金额
        print("3. 计算精确欠料金额...")
        self.merged_precise = pd.DataFrame(precise_with_supplier)
        
        if not self.merged_precise.empty:
            self.merged_precise['仓存不足_数值'] = pd.to_numeric(self.merged_precise['仓存不足'], errors='coerce').fillna(0)
            self.merged_precise['欠料金额_RMB'] = self.merged_precise['仓存不足_数值'] * self.merged_precise['RMB单价']
        
        # 保存多供应商选择表
        self.multi_supplier_df = pd.DataFrame(multi_supplier_records)
        
        # 统计结果
        if not self.merged_precise.empty:
            matched_supplier = len(self.merged_precise[self.merged_precise['主供应商名称'] != '未找到供应商'])
            total_amount = self.merged_precise['欠料金额_RMB'].sum()
            multi_supplier_materials = len(self.multi_supplier_df['物料编号'].unique()) if not self.multi_supplier_df.empty else 0
            
            print(f"   ✅ 匹配到供应商: {matched_supplier}/{len(self.merged_precise)}条")
            print(f"   💰 精确计算欠料总金额: ¥{total_amount:,.2f}")
            print(f"   🔄 多供应商物料: {multi_supplier_materials}个")
        
        print("✅ 精确匹配完成\n")
        
    def estimated_completion(self):
        """估算补全 - 对未在欠料表的订单进行估算"""
        print("=== 📊 估算补全分析 ===")
        
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
        
        # 2. 基于供应商数据的估算模型
        print("2. 基于供应商数据建立估算模型...")
        
        # 从供应商表获取不同价格区间的统计
        if not self.supplier_df.empty:
            valid_suppliers = self.supplier_df[self.supplier_df['RMB单价'] > 0]
            
            price_stats = {
                '低价物料': valid_suppliers[valid_suppliers['RMB单价'] <= 1]['RMB单价'].mean(),
                '中价物料': valid_suppliers[(valid_suppliers['RMB单价'] > 1) & (valid_suppliers['RMB单价'] <= 10)]['RMB单价'].mean(),
                '高价物料': valid_suppliers[valid_suppliers['RMB单价'] > 10]['RMB单价'].mean()
            }
            
            # 填充NaN值
            for key, value in price_stats.items():
                if pd.isna(value):
                    if key == '低价物料':
                        price_stats[key] = 0.5
                    elif key == '中价物料':
                        price_stats[key] = 5.0
                    else:
                        price_stats[key] = 50.0
        else:
            price_stats = {'低价物料': 0.5, '中价物料': 5.0, '高价物料': 50.0}
        
        print(f"   📊 估算价格基准: 低价¥{price_stats['低价物料']:.2f}, 中价¥{price_stats['中价物料']:.2f}, 高价¥{price_stats['高价物料']:.2f}")
        
        # 3. 产品复杂度评估
        def estimate_product_cost(product_model, order_qty):
            model_str = str(product_model).upper()
            
            # 基础复杂度
            if any(feature in model_str for feature in ['SP', 'AS']):
                base_cost = price_stats['中价物料'] * 20  # 标准产品约20个物料
                complexity = 1.0
            elif any(feature in model_str for feature in ['BT', 'BAB', 'TOB']):
                base_cost = price_stats['高价物料'] * 15  # 复杂产品约15个高价物料
                complexity = 1.5
            else:
                base_cost = price_stats['低价物料'] * 25  # 简单产品约25个低价物料
                complexity = 0.8
            
            # 数量影响 (规模效应)
            qty_factor = min(1.0, max(0.3, 1000 / max(order_qty, 100)))
            
            return base_cost * complexity * qty_factor
        
        # 4. 生成估算记录
        estimated_records = []
        
        for _, order in remaining_orders.iterrows():
            order_qty = pd.to_numeric(order['订单数量'], errors='coerce')
            if pd.isna(order_qty) or order_qty <= 0:
                order_qty = 1000  # 默认数量
            
            estimated_cost = estimate_product_cost(order['产品型号'], order_qty)
            
            estimated_records.append({
                'ITEM NO.': order.get('ITEM NO.', ''),
                '生产单号': order['生产单号'],
                '客户订单号': order['客户订单号'],
                '产品型号': order['产品型号'],
                '订单数量': order['订单数量'],
                '月份': order['月份'],
                '数据来源工作表': order.get('数据来源工作表', ''),
                '目的地': order.get('目的地', ''),
                '交期': order.get('交期', ''),
                'BOM编号': order.get('BOM编号', ''),
                
                # 欠料估算信息
                '物料编号': f"估算-{order['产品型号']}",
                '物料名称': f"估算物料成本组合-{order['产品型号']}",
                '仓存不足': 1,  # 估算数量
                '主供应商名称': '估算供应商',
                '主供应商号': 'EST001',
                'RMB单价': estimated_cost,
                '欠料金额_RMB': estimated_cost,
                '计算方式': '估算',
                '修改日期': datetime.now()
            })
        
        self.merged_estimated = pd.DataFrame(estimated_records)
        
        total_estimated = self.merged_estimated['欠料金额_RMB'].sum()
        print(f"   💰 估算欠料总金额: ¥{total_estimated:,.2f}")
        print("✅ 估算补全完成\n")
        
    def generate_report1_order_shortage_detail(self):
        """生成表1: 订单缺料明细 (主供应商版本)"""
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
        
        # 生成最终报表
        report1_data = []
        
        for _, row in combined_df.iterrows():
            report1_data.append({
                '客户订单号': row['客户订单号'],
                '生产订单号': row['生产单号'],
                '产品型号': row['产品型号'], 
                '订单数量': row['订单数量'],
                '月份': row['月份'],
                '数据来源工作表': row.get('数据来源工作表', ''),
                '目的地': row.get('目的地', ''),
                '客户交期': row.get('交期', ''),
                'BOM编号': row.get('BOM编号', ''),
                
                '欠料物料编号': row['物料编号'],
                '欠料物料名称': row['物料名称'],
                '欠料数量': row.get('仓存不足', 0),
                
                '主供应商名称': row.get('主供应商名称', ''),
                '主供应商号': row.get('主供应商号', ''),
                '供应商单价(原币)': row.get('主供应商单价', 0),
                '币种': row.get('主供应商币种', 'RMB'),
                'RMB单价': row.get('RMB单价', 0),
                '起订数量': row.get('起订数量', 0),
                '供应商修改日期': row.get('修改日期', ''),
                
                '欠料金额(RMB)': row.get('欠料金额_RMB', 0),
                '计算方式': row.get('计算方式', ''),
                
                # 额外信息
                '工单需求': row.get('工单需求', ''),
                '已购未返': row.get('已购未返', ''),
                '手头现有': row.get('手头现有', ''),
                '请购组': row.get('请购组', '')
            })
        
        report1_df = pd.DataFrame(report1_data)
        
        print(f"   📊 订单缺料明细: {len(report1_df)}条记录")
        print(f"   📊 涉及订单: {report1_df['生产订单号'].nunique()}个")
        print(f"   📊 涉及供应商: {report1_df['主供应商名称'].nunique()}家")
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
            '物料编号': lambda x: '; '.join(x.astype(str).unique()),
            '物料名称': lambda x: '; '.join(x.astype(str).unique()),
            '主供应商名称': lambda x: '; '.join([str(s) for s in x.unique() if str(s) != 'nan']),
            '欠料金额_RMB': 'sum',
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
                '需采购物料清单': row['物料名称'],
                '涉及供应商': row['主供应商名称'],
                '采购总金额(RMB)': row['欠料金额_RMB'],
                '平均物料单价': row['欠料金额_RMB'] / max(len(row['物料编号'].split('; ')), 1),
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
        
        # 处理精确匹配的记录
        supplier_data = {}
        
        if not self.merged_precise.empty:
            for _, row in self.merged_precise.iterrows():
                supplier_name = row.get('主供应商名称', '未知供应商')
                
                if supplier_name not in supplier_data:
                    supplier_data[supplier_name] = {
                        '供应商号': row.get('主供应商号', ''),
                        'orders': [],
                        'total_amount': 0,
                        'material_count': 0,
                        'order_count': set()
                    }
                
                supplier_data[supplier_name]['orders'].append({
                    '生产订单号': row['生产单号'],
                    '产品型号': row['产品型号'],
                    '物料编号': row['物料编号'],
                    '欠料金额': row.get('欠料金额_RMB', 0)
                })
                supplier_data[supplier_name]['total_amount'] += row.get('欠料金额_RMB', 0)
                supplier_data[supplier_name]['material_count'] += 1
                supplier_data[supplier_name]['order_count'].add(row['生产单号'])
        
        # 处理估算记录
        if not self.merged_estimated.empty:
            for _, row in self.merged_estimated.iterrows():
                supplier_name = row.get('主供应商名称', '估算供应商')
                
                if supplier_name not in supplier_data:
                    supplier_data[supplier_name] = {
                        '供应商号': row.get('主供应商号', ''),
                        'orders': [],
                        'total_amount': 0,
                        'material_count': 0,
                        'order_count': set()
                    }
                
                supplier_data[supplier_name]['orders'].append({
                    '生产订单号': row['生产单号'],
                    '产品型号': row['产品型号'],
                    '物料编号': row['物料编号'],
                    '欠料金额': row.get('欠料金额_RMB', 0)
                })
                supplier_data[supplier_name]['total_amount'] += row.get('欠料金额_RMB', 0)
                supplier_data[supplier_name]['material_count'] += 1
                supplier_data[supplier_name]['order_count'].add(row['生产单号'])
        
        # 生成供应商汇总报表
        report3_data = []
        for supplier, info in supplier_data.items():
            order_list = list(set([f"{order['生产订单号']}({order['产品型号']})" for order in info['orders']]))
            
            report3_data.append({
                '供应商名称': supplier,
                '供应商号': info['供应商号'],
                '相关订单数': len(info['order_count']),
                '相关物料数': info['material_count'],
                '相关订单列表': '; '.join(order_list[:10]) + ('...' if len(order_list) > 10 else ''),  # 限制长度
                '采购总金额(RMB)': info['total_amount'],
                '平均物料单价': info['total_amount'] / info['material_count'] if info['material_count'] > 0 else 0,
                '平均订单金额': info['total_amount'] / len(info['order_count']) if len(info['order_count']) > 0 else 0
            })
        
        report3_df = pd.DataFrame(report3_data)
        report3_df = report3_df.sort_values('采购总金额(RMB)', ascending=False)
        
        print(f"   📊 涉及供应商: {len(report3_df)}家")
        if len(report3_df) > 0:
            print(f"   💰 总采购金额: ¥{report3_df['采购总金额(RMB)'].sum():,.2f}")
            print(f"   🔝 最大供应商: {report3_df.iloc[0]['供应商名称']} (¥{report3_df.iloc[0]['采购总金额(RMB)']:,.2f})")
        
        return report3_df
    
    def generate_report4_multi_supplier_choice(self):
        """生成表4: 多供应商选择表 (方案B)"""
        print("=== 🔄 生成表4: 多供应商选择表 ===")
        
        if self.multi_supplier_df.empty:
            print("❌ 无多供应商数据")
            return pd.DataFrame()
        
        # 按物料分组展示
        report4_data = []
        
        for material_code in self.multi_supplier_df['物料编号'].unique():
            material_suppliers = self.multi_supplier_df[
                self.multi_supplier_df['物料编号'] == material_code
            ].sort_values(['是否主选', 'RMB单价'], ascending=[False, True])
            
            for rank, (_, supplier) in enumerate(material_suppliers.iterrows(), 1):
                report4_data.append({
                    '物料编号': supplier['物料编号'],
                    '物料名称': supplier['物料名称'],
                    '供应商排序': f"选项{rank}",
                    '供应商名称': supplier['供应商名称'],
                    '供应商号': supplier['供应商号'],
                    '是否主选': '✅主选' if supplier['是否主选'] else '备选',
                    'RMB单价': supplier['RMB单价'],
                    '原币单价': supplier['原币单价'],
                    '币种': supplier['币种'],
                    '起订数量': supplier['起订数量'],
                    '修改日期': supplier['修改日期'],
                    '价格排名': material_suppliers['RMB单价'].rank().loc[supplier.name],
                    '日期排名': material_suppliers['修改日期'].rank(ascending=False, na_option='bottom').loc[supplier.name] if pd.notna(supplier['修改日期']) else '无日期'
                })
        
        report4_df = pd.DataFrame(report4_data)
        
        print(f"   📊 多供应商物料: {report4_df['物料编号'].nunique()}个")
        print(f"   📊 供应商选择记录: {len(report4_df)}条")
        
        return report4_df
    
    def save_final_reports(self, report1, report2, report3, report4):
        """保存最终报表"""
        print("=== 💾 保存报表文件 ===")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f'精准供应商物料分析报告_{timestamp}.xlsx'
        filepath = f'D:\\yingtu-PMC\\{filename}'
        
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                # 主要报表
                if not report1.empty:
                    report1.to_excel(writer, sheet_name='1_订单缺料明细', index=False)
                if not report2.empty:
                    report2.to_excel(writer, sheet_name='2_8月采购汇总', index=False)
                if not report3.empty:
                    report3.to_excel(writer, sheet_name='3_供应商汇总', index=False)
                if not report4.empty:
                    report4.to_excel(writer, sheet_name='4_多供应商选择表', index=False)
                    
                # 汇总统计sheet
                summary_data = {
                    '项目': [
                        '总订单数', '有欠料订单数', '精确匹配订单', '估算订单', 
                        '8月需采购订单', '涉及供应商数', '多供应商物料数',
                        '总欠料金额(RMB)', '8月采购金额(RMB)'
                    ],
                    '数量': [
                        len(self.orders_df),
                        len(report1) if not report1.empty else 0,
                        len(self.merged_precise) if not self.merged_precise.empty else 0,
                        len(self.merged_estimated) if not self.merged_estimated.empty else 0,
                        len(report2) if not report2.empty else 0,
                        len(report3) if not report3.empty else 0,
                        self.multi_supplier_df['物料编号'].nunique() if not self.multi_supplier_df.empty else 0,
                        f"¥{report1['欠料金额(RMB)'].sum():,.2f}" if not report1.empty else "¥0.00",
                        f"¥{report2['采购总金额(RMB)'].sum():,.2f}" if not report2.empty else "¥0.00"
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
        """运行完整的供应商物料分析"""
        print("🚀 开始精准供应商物料分析 (混合方案A+B)")
        print("="*80)
        
        try:
            # 1. 加载数据
            self.load_all_data()
            
            # 2. 精确匹配 (含供应商选择)
            self.precise_matching_with_supplier()
            
            # 3. 估算补全  
            self.estimated_completion()
            
            # 4. 生成四个报表
            report1 = self.generate_report1_order_shortage_detail()
            report2 = self.generate_report2_august_purchase_summary()
            report3 = self.generate_report3_supplier_summary()
            report4 = self.generate_report4_multi_supplier_choice()
            
            # 5. 保存报表
            filename = self.save_final_reports(report1, report2, report3, report4)
            
            # 6. 输出最终汇总
            print("\n" + "="*80)
            print(" "*20 + "🎉 分析完成！完美实现混合方案！")
            print("="*80)
            
            total_orders = len(self.orders_df)
            precise_records = len(self.merged_precise) if not self.merged_precise.empty else 0
            estimated_orders = len(self.merged_estimated) if not self.merged_estimated.empty else 0
            total_amount = report1['欠料金额(RMB)'].sum() if not report1.empty else 0
            suppliers_count = len(report3) if not report3.empty else 0
            multi_supplier_materials = self.multi_supplier_df['物料编号'].nunique() if not self.multi_supplier_df.empty else 0
            
            print(f"📊 完整数据分析汇总:")
            print(f"   - 总订单数: {total_orders}个")
            print(f"   - 精确匹配记录: {precise_records}条")
            print(f"   - 估算补全订单: {estimated_orders}条")
            print(f"   - 涉及供应商: {suppliers_count}家")
            print(f"   - 多供应商物料: {multi_supplier_materials}个")
            print(f"   - 总欠料金额: ¥{total_amount:,.2f}")
            
            if filename:
                print(f"📄 完整报表: {filename}")
                print("📋 包含表格:")
                print("   1️⃣ 订单缺料明细 (主供应商)")
                print("   2️⃣ 8月采购汇总")  
                print("   3️⃣ 供应商汇总")
                print("   4️⃣ 多供应商选择表 ← 新增")
                print("   0️⃣ 汇总统计")
                
            return report1, report2, report3, report4
            
        except Exception as e:
            print(f"❌ 分析过程出错: {e}")
            import traceback
            traceback.print_exc()
            return None, None, None, None

if __name__ == "__main__":
    analyzer = FinalSupplierMaterialAnalyzer()
    analyzer.run_complete_analysis()
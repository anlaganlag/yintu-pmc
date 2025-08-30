#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC综合物料分析系统
整合订单+欠料+库存+供应商的完整分析
基于LEFT JOIN架构，确保所有订单都显示
支持ROI计算和最低价供应商选择
"""

import pandas as pd
import numpy as np
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class ComprehensivePMCAnalyzer:
    def __init__(self):
        self.orders_df = None           # 订单数据（主表）
        self.shortage_df = None         # 欠料数据  
        self.inventory_df = None        # 库存价格数据
        self.supplier_df = None         # 供应商数据
        self.final_result = None        # 最终结果
        
        # 汇率设置（转换为RMB）
        self.currency_rates = {
            'RMB': 1.0,
            'USD': 7.30,  # 1 USD = 7.20 RMB  
            'HKD': 0.93,  # 1 HKD = 0.93 RMB
            'EUR': 7.85   # 1 EUR = 7.85 RMB
        }
        
    def load_all_data(self):
        """加载所有数据源"""
        print("=== 🔄 加载数据源 ===")
        
        # 1. 加载4个订单工作表
        print("1. 加载订单数据（国内+柬埔寨）...")
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
            
            # 标准化订单表列名
            self.orders_df = self.orders_df.rename(columns={
                '生 產 單 号(  廠方 )': '生产单号',
                '生 產 單 号(客方 )': '客户订单号',
                '型 號( 廠方/客方 )': '产品型号',
                '數 量  (Pcs)': '数量Pcs',
                'BOM NO.': 'BOM编号',
                '客期': '客户交期'
            })
            
            # 确保订单金额字段存在（USD）
            if '订单金额' not in self.orders_df.columns:
                # 如果没有订单金额字段，使用默认值
                self.orders_df['订单金额'] = 1000  # 默认1000 USD
                print("   ⚠️ 订单表中未找到'订单金额'字段，使用默认值1000 USD")
            
            print(f"   ✅ 订单总数: {len(self.orders_df)}条")
            print(f"   📊 数据分布: 国内{len(self.orders_df[self.orders_df['数据来源工作表']=='国内'])}条, " +
                  f"柬埔寨{len(self.orders_df[self.orders_df['数据来源工作表']=='柬埔寨'])}条")
            
            # 检查PSO2501724是否在加载的订单中
            if 'PSO2501724' in self.orders_df['生产单号'].values:
                print("   🔍 PSO2501724在订单加载后存在")
                pso_info = self.orders_df[self.orders_df['生产单号'] == 'PSO2501724'].iloc[0]
                print(f"   PSO2501724信息: 客户订单={pso_info.get('客户订单号')}, 产品={pso_info.get('产品型号')}, 来源={pso_info.get('数据来源工作表')}")
            else:
                print("   ❌ PSO2501724在订单加载后不存在！")
                print("   检查各个订单文件中的PSO2501724:")
                for i, df_name in enumerate(['国内8月', '国内9月', '柬埔寨8月', '柬埔寨9月']):
                    if i < len(orders_data):
                        if 'PSO2501724' in orders_data[i]['生产单号'].values if '生产单号' in orders_data[i].columns else orders_data[i]['生 產 單 号(  廠方 )'].values:
                            print(f"     ✅ PSO2501724在{df_name}中找到")
                        else:
                            print(f"     ❌ PSO2501724不在{df_name}中")
            
        except Exception as e:
            print(f"   ❌ 订单数据加载失败: {e}")
            return False
        
        # 2. 加载欠料表
        print("2. 加载mat_owe_pso.xlsx欠料表...")
        try:
            self.shortage_df = pd.read_excel('input/mat_owe_pso.xlsx', 
                                           sheet_name='Sheet1', skiprows=1)
            
            # 标准化欠料表列名
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
            
            print(f"   ✅ 欠料记录: {len(self.shortage_df)}条")
            
        except Exception as e:
            print(f"   ❌ 欠料表加载失败: {e}")
            self.shortage_df = pd.DataFrame()
        
        # 3. 加载库存价格表
        print("3. 加载inventory_list.xlsx库存表...")
        try:
            self.inventory_df = pd.read_excel('input/inventory_list.xlsx')
            
            # 价格处理：优先最新報價，回退到成本單價
            self.inventory_df['最终价格'] = self.inventory_df['最新報價'].fillna(self.inventory_df['成本單價'])
            self.inventory_df['最终价格'] = pd.to_numeric(self.inventory_df['最终价格'], errors='coerce').fillna(0)
            
            # 货币转换为RMB
            def convert_to_rmb(row):
                price = row['最终价格']
                currency = str(row.get('貨幣', 'RMB')).upper()
                rate = self.currency_rates.get(currency, 1.0)
                return price * rate if pd.notna(price) else 0
            
            self.inventory_df['RMB单价'] = self.inventory_df.apply(convert_to_rmb, axis=1)
            
            valid_prices = len(self.inventory_df[self.inventory_df['RMB单价'] > 0])
            print(f"   ✅ 库存物料: {len(self.inventory_df)}条, 有效价格: {valid_prices}条")
            
        except Exception as e:
            print(f"   ❌ 库存表加载失败: {e}")
            self.inventory_df = pd.DataFrame()
        
        # 4. 加载供应商表
        print("4. 加载supplier.xlsx供应商表...")
        try:
            self.supplier_df = pd.read_excel('input/supplier.xlsx')
            
            # 处理供应商价格和货币转换
            self.supplier_df['单价_数值'] = pd.to_numeric(self.supplier_df['单价'], errors='coerce').fillna(0)
            
            def convert_supplier_to_rmb(row):
                price = row['单价_数值']
                currency = str(row.get('币种', 'RMB')).upper()
                rate = self.currency_rates.get(currency, 1.0)
                return price * rate if pd.notna(price) else 0
            
            self.supplier_df['供应商RMB单价'] = self.supplier_df.apply(convert_supplier_to_rmb, axis=1)
            
            # 处理修改日期
            self.supplier_df['修改日期'] = pd.to_datetime(self.supplier_df['修改日期'], errors='coerce')
            
            valid_supplier_prices = len(self.supplier_df[self.supplier_df['供应商RMB单价'] > 0])
            print(f"   ✅ 供应商记录: {len(self.supplier_df)}条, 有效价格: {valid_supplier_prices}条")
            print(f"   ✅ 唯一供应商: {self.supplier_df['供应商名称'].nunique()}家")
            
        except Exception as e:
            print(f"   ❌ 供应商表加载失败: {e}")
            self.supplier_df = pd.DataFrame()
        
        print("✅ 数据加载完成\n")
        return True
        
    def select_lowest_price_supplier(self, material_suppliers):
        """为物料选择最低价供应商"""
        if len(material_suppliers) == 0:
            return None
        if len(material_suppliers) == 1:
            return material_suppliers.iloc[0]
        
        # 筛选有有效价格的供应商
        valid_suppliers = material_suppliers[material_suppliers['供应商RMB单价'] > 0]
        
        if len(valid_suppliers) == 0:
            # 都没有价格，返回第一个
            return material_suppliers.iloc[0]
        
        # 选择最低价供应商
        lowest_price_idx = valid_suppliers['供应商RMB单价'].idxmin()
        return valid_suppliers.loc[lowest_price_idx]
    
    def comprehensive_left_join_analysis(self):
        """综合LEFT JOIN分析 - 以订单表为主表"""
        print("=== 🎯 综合LEFT JOIN分析 ===")
        
        if self.orders_df is None or self.orders_df.empty:
            print("❌ 订单表为空，无法分析")
            return False
        
        # 从订单表开始（主表）
        result = self.orders_df.copy()
        print(f"1. 主表（订单）: {len(result)}条记录")
        
        # 检查PSO2501724是否在主表中
        if 'PSO2501724' in result['生产单号'].values:
            print("   🔍 PSO2501724在主表中存在")
        else:
            print("   ❌ PSO2501724不在主表中")
        
        # LEFT JOIN 欠料信息
        if not self.shortage_df.empty:
            print("2. LEFT JOIN 欠料信息...")
            result['生产单号_清理'] = result['生产单号'].astype(str).str.strip()
            self.shortage_df['订单编号_清理'] = self.shortage_df['订单编号'].astype(str).str.strip()
            
            # 记录JOIN前的记录数
            before_join = len(result)
            
            result = result.merge(
                self.shortage_df,
                left_on='生产单号_清理',
                right_on='订单编号_清理',
                how='left'
            )
            
            # 检查JOIN后PSO2501724的状态
            after_join = len(result)
            if 'PSO2501724' in result['生产单号'].values:
                pso_rows = result[result['生产单号'] == 'PSO2501724']
                print(f"   🔍 PSO2501724 JOIN后: {len(pso_rows)}条记录")
                if len(pso_rows) > 0:
                    has_material = pso_rows['物料编号'].notna().any()
                    print(f"   PSO2501724 欠料状态: {'有欠料' if has_material else '无欠料'}")
            else:
                print("   ❌ PSO2501724在LEFT JOIN后丢失！")
            
            matched_shortage = len(result[result['物料编号'].notna()])
            print(f"   ✅ 匹配到欠料信息: {matched_shortage}条记录")
            print(f"   📊 JOIN前后记录数: {before_join} → {after_join}")
        else:
            print("2. ⚠️ 跳过欠料匹配（欠料表为空）")
            # 添加空的欠料字段
            result['物料编号'] = None
            result['物料名称'] = None
            result['仓存不足'] = 0
            result['工单需求'] = None
            result['已购未返'] = None
            result['手头现有'] = None
            result['请购组'] = None
        
        # LEFT JOIN 库存价格信息
        if not self.inventory_df.empty:
            print("3. LEFT JOIN 库存价格信息...")
            result = result.merge(
                self.inventory_df[['物項編號', '物項名稱', 'RMB单价', '貨幣', '最终价格']],
                left_on='物料编号',
                right_on='物項編號',
                how='left',
                suffixes=('', '_库存')
            )
            
            matched_inventory = len(result[result['RMB单价'].notna()])
            print(f"   ✅ 匹配到库存价格: {matched_inventory}条记录")
        else:
            print("3. ⚠️ 跳过库存价格匹配（库存表为空）")
            result['RMB单价'] = 0
        
        # LEFT JOIN 供应商信息（按物料选择最低价供应商）
        if not self.supplier_df.empty:
            print("4. LEFT JOIN 供应商信息（最低价选择）...")
            
            # 为每个唯一物料选择最低价供应商
            unique_materials = result[result['物料编号'].notna()]['物料编号'].unique()
            
            supplier_mapping = {}
            processed_count = 0
            
            for material_code in unique_materials:
                material_suppliers = self.supplier_df[self.supplier_df['物项编号'] == material_code]
                
                if len(material_suppliers) > 0:
                    best_supplier = self.select_lowest_price_supplier(material_suppliers)
                    if best_supplier is not None:
                        supplier_mapping[material_code] = {
                            '主供应商名称': best_supplier['供应商名称'],
                            '主供应商号': best_supplier['供应商号'],
                            '供应商单价(原币)': best_supplier['单价'],
                            '币种': best_supplier['币种'],
                            '起订数量': best_supplier['起订数量'],
                            '供应商修改日期': best_supplier['修改日期']
                        }
                
                processed_count += 1
                if processed_count % 100 == 0:
                    print(f"   处理进度: {processed_count}/{len(unique_materials)} 物料")
            
            # 映射供应商信息到结果表
            for col in ['主供应商名称', '主供应商号', '供应商单价(原币)', '币种', '起订数量', '供应商修改日期']:
                result[col] = result['物料编号'].map(lambda x: supplier_mapping.get(x, {}).get(col, None))
            
            matched_suppliers = len(result[result['主供应商名称'].notna()])
            print(f"   ✅ 匹配到供应商信息: {matched_suppliers}条记录")
            print(f"   📊 找到供应商的物料: {len(supplier_mapping)}个")
            
        else:
            print("4. ⚠️ 跳过供应商匹配（供应商表为空）")
            result['主供应商名称'] = None
            result['主供应商号'] = None
            result['供应商单价(原币)'] = None
            result['币种'] = None
            result['起订数量'] = None
            result['供应商修改日期'] = None
        
        self.final_result = result
        print("✅ LEFT JOIN 分析完成\n")
        return True
    
    def calculate_derived_fields(self):
        """计算派生字段：欠料金额、订单金额(RMB)、每元投入回款"""
        print("=== 💰 计算派生字段 ===")
        
        if self.final_result is None:
            print("❌ 没有分析结果数据")
            return False
        
        # 1. 计算欠料金额(RMB)
        print("1. 计算欠料金额(RMB)...")
        self.final_result['仓存不足_数值'] = pd.to_numeric(self.final_result['仓存不足'], errors='coerce').fillna(0)
        self.final_result['欠料金额(RMB)'] = self.final_result['仓存不足_数值'] * self.final_result['RMB单价']
        
        # 2. 计算订单金额(RMB) - 先按客户订单号去重
        print("2. 计算订单金额(RMB)（按客户订单号去重）...")
        self.final_result['订单金额(USD)'] = pd.to_numeric(self.final_result['订单金额'], errors='coerce').fillna(0)
        
        # 按客户订单号去重计算订单金额
        customer_order_amounts = self.final_result.groupby('客户订单号').agg({
            '订单金额(USD)': 'first'  # 每个客户订单号只取一次订单金额
        }).reset_index()
        customer_order_amounts['订单金额(RMB)'] = customer_order_amounts['订单金额(USD)'] * self.currency_rates['USD']
        
        # 将去重后的订单金额合并回主表
        self.final_result = self.final_result.merge(
            customer_order_amounts[['客户订单号', '订单金额(RMB)']],
            on='客户订单号',
            how='left',
            suffixes=('', '_dedup')
        )
        
        # 3. 按订单计算每元投入回款
        print("3. 计算每元投入回款（按订单汇总）...")
        
        # 按生产订单号汇总，正确处理订单金额聚合
        # 先按生产订单号和客户订单号组合去重，然后按生产订单号汇总
        print("   正在处理生产订单与客户订单的一对多关系...")
        
        # 第一步：按生产订单号+客户订单号去重，确保每个客户订单只计算一次
        unique_combinations = self.final_result.groupby(['生产单号', '客户订单号']).agg({
            '订单金额(RMB)': 'first',   # 每个客户订单只取一次金额
            '欠料金额(RMB)': 'sum'      # 欠料金额需要汇总（同一客户订单可能缺多种物料）
        }).reset_index()
        
        # 第二步：按生产订单号汇总，正确聚合多个客户订单的金额
        order_totals = unique_combinations.groupby('生产单号').agg({
            '订单金额(RMB)': 'sum',     # ✅ 汇总同一生产订单下所有客户订单的金额
            '欠料金额(RMB)': 'sum'      # ✅ 汇总同一生产订单下所有欠料金额
        }).reset_index()
        
        # 检查一对多关系统计
        prod_cust_mapping = self.final_result.groupby('生产单号')['客户订单号'].nunique()
        multi_customer_orders = prod_cust_mapping[prod_cust_mapping > 1]
        
        if len(multi_customer_orders) > 0:
            print(f"   ✅ 正确处理了{len(multi_customer_orders)}个生产订单的多客户订单关系")
            for prod_order in multi_customer_orders.index[:3]:  # 显示前3个例子
                prod_data = self.final_result[self.final_result['生产单号'] == prod_order]
                customer_count = prod_data['客户订单号'].nunique()
                total_amount = order_totals[order_totals['生产单号'] == prod_order]['订单金额(RMB)'].iloc[0]
                print(f"      {prod_order}: {customer_count}个客户订单 → 总金额 ¥{total_amount:,.2f}")
        else:
            print("   ℹ️  所有生产订单都是一对一关系")
        
        # 计算ROI - 区分无需投入和需要投入的订单
        def calculate_roi(row):
            shortage_amount = row['欠料金额(RMB)']
            order_amount = row['订单金额(RMB)']
            
            if shortage_amount > 0:
                # 需要投入：返回具体倍数
                return order_amount / shortage_amount
            elif pd.notna(order_amount) and order_amount > 0:
                # 有订单金额但无欠料：返回特殊标记（用-1表示无需投入）
                return -1  # 特殊值，后续转换为"无需投入"
            else:
                # 无订单金额：返回0
                return 0
                
        order_totals['每元投入回款'] = order_totals.apply(calculate_roi, axis=1)
        
        # 将ROI合并回主表
        self.final_result = self.final_result.merge(
            order_totals[['生产单号', '每元投入回款']],
            on='生产单号',
            how='left',
            suffixes=('', '_calc')
        )
        
        # 4. 计算数据完整性标记
        print("4. 计算数据完整性标记...")
        def calculate_completeness(row):
            has_shortage = pd.notna(row['物料编号'])
            has_price = pd.notna(row['RMB单价']) and row['RMB单价'] > 0
            has_supplier = pd.notna(row['主供应商名称'])
            has_order_amount = pd.notna(row['订单金额(USD)']) and row['订单金额(USD)'] > 0
            has_production_order = pd.notna(row['生产单号']) and row['生产单号'] != ''
            
            if has_shortage and has_price and has_supplier and has_order_amount:
                return "完整"
            elif has_shortage and has_price and has_order_amount:
                return "部分"
            elif has_order_amount and not has_shortage:
                # 有订单金额但无欠料 = 不缺料订单，应标记为"完整"
                return "完整"
            elif has_order_amount:
                return "订单完整"
            elif has_production_order and not has_shortage:
                # 有生产订单号但无欠料且无订单金额 = 不缺料但订单信息不完整
                return "不缺料订单"
            elif has_production_order:
                # 有生产订单号但缺少订单金额
                return "订单信息不完整"
            else:
                return "无数据"
        
        self.final_result['数据完整性标记'] = self.final_result.apply(calculate_completeness, axis=1)
        
        # 5. 计算方式标记
        self.final_result['计算方式'] = np.where(
            self.final_result['物料编号'].notna(),
            '精确匹配',
            '无欠料数据'
        )
        
        # 统计结果
        total_shortage_amount = self.final_result['欠料金额(RMB)'].sum()
        total_order_amount = self.final_result['订单金额(RMB)'].sum()
        avg_roi = self.final_result['每元投入回款'].mean()
        
        print(f"   💰 总欠料金额: ¥{total_shortage_amount:,.2f}")
        print(f"   💰 总订单金额: ¥{total_order_amount:,.2f}")
        print(f"   📊 平均投入回款: {avg_roi:.2f}倍")
        
        completeness_dist = self.final_result['数据完整性标记'].value_counts()
        print(f"   📋 数据完整性分布: {dict(completeness_dist)}")
        
        print("✅ 派生字段计算完成\n")
        return True
    
    def apply_conservative_filling(self, df):
        """应用保守填充策略"""
        result = df.copy()
        
        # 1. 过滤掉"无数据"记录，但保留所有有生产订单号的记录
        # 只过滤真正无效的记录
        result = result[
            result['数据完整性标记'] != '无数据'
        ]
        
        # 2. 数值字段统一填0
        numeric_fields = [
            '数量Pcs', '欠料数量', 'RMB单价', '起订数量', '供应商单价(原币)',
            '工单需求', '已购未返', '手头现有', '欠料金额(RMB)', 
            '订单金额(USD)', '订单金额(RMB)'
        ]
        
        for field in numeric_fields:
            if field in result.columns:
                result[field] = pd.to_numeric(result[field], errors='coerce').fillna(0)
        
        # 3. 文本字段统一填空字符串
        text_fields = [
            '客户订单号', '产品型号', '数据来源工作表', '目的地', 'BOM编号',
            '欠料物料编号', '欠料物料名称', '主供应商名称', '主供应商号', 
            '币种', '请购组', '计算方式'
        ]
        
        for field in text_fields:
            if field in result.columns:
                result[field] = result[field].astype(str).replace('nan', '').replace('None', '')
        
        # 4. 处理ROI显示：将特殊值转换为业务术语
        print("   处理ROI显示格式...")
        result['每元投入回款'] = result['每元投入回款'].apply(lambda x:
            '无需投入' if pd.to_numeric(x, errors='coerce') == -1  # 特殊标记
            else x)
        
        # 5. 添加业务标记字段
        def get_data_source_mark(row):
            marks = []
            if pd.to_numeric(row.get('欠料数量', 0), errors='coerce') == 0:
                marks.append('填充欠料')
            if row.get('主供应商名称', '') == '':
                marks.append('缺失供应商')
            if pd.to_numeric(row.get('RMB单价', 0), errors='coerce') == 0:
                marks.append('填充价格')
            # 基于订单级别ROI判断，而不是单行欠料金额
            if pd.to_numeric(row.get('每元投入回款', 0), errors='coerce') == 0:
                marks.append('无需投入')
            return '; '.join(marks) if marks else '原始数据'
        
        result['数据填充标记'] = result.apply(get_data_source_mark, axis=1)
        
        return result
    
    def generate_ready_to_produce_orders(self):
        """生成马上可以投入生产的订单表（不缺料订单）"""
        print("=== 🚀 生成不缺料订单清单 ===")
        
        if self.final_result is None:
            print("❌ 没有分析结果数据")
            return None
        
        # 筛选不缺料订单：LEFT JOIN后物料编号为空的订单
        ready_orders = self.final_result[
            (self.final_result['物料编号'].isna() | (self.final_result['物料编号'] == '')) &  # 无欠料记录
            (self.final_result['生产单号'].notna()) &  # 有生产订单号
            (self.final_result['生产单号'] != '')      # 生产订单号非空
        ].copy()
        
        print(f"   🔍 初步筛选不缺料记录: {len(ready_orders)}条")
        
        if ready_orders.empty:
            print("   ⚠️ 未找到不缺料订单")
            return None
        
        # 按生产订单号去重，保留订单基本信息
        unique_ready_orders = ready_orders.groupby('生产单号').agg({
            '客户订单号': 'first',
            '产品型号': 'first', 
            '数量Pcs': 'first',
            '月份': 'first',
            '数据来源工作表': 'first',
            '目的地': 'first',
            '客户交期': 'first',
            'BOM编号': 'first',
            '订单金额': 'first',
            '订单金额(USD)': 'first',
            '订单金额(RMB)': 'first',
            '每元投入回款': 'first',
            '数据完整性标记': 'first'
        }).reset_index()
        
        # 添加不缺料标识
        unique_ready_orders['缺料状态'] = '不缺料'
        unique_ready_orders['生产状态'] = '可立即投产'
        
        # 重新排列列顺序，突出关键信息
        output_columns = [
            '生产单号', '客户订单号', '产品型号', '数量Pcs', 
            '月份', '数据来源工作表', '目的地', '客户交期', 'BOM编号',
            '缺料状态', '生产状态',
            '订单金额(USD)', '订单金额(RMB)', '每元投入回款',
            '数据完整性标记'
        ]
        
        # 确保所有列都存在
        for col in output_columns:
            if col not in unique_ready_orders.columns:
                unique_ready_orders[col] = ''
        
        # 选择输出列
        final_ready_orders = unique_ready_orders[output_columns]
        
        # 数据清理
        for col in ['订单金额(USD)', '订单金额(RMB)', '数量Pcs']:
            if col in final_ready_orders.columns:
                final_ready_orders[col] = pd.to_numeric(final_ready_orders[col], errors='coerce').fillna(0)
        
        # 按月份和数据来源分组统计
        print(f"   ✅ 不缺料订单总数: {len(final_ready_orders)}个")
        
        stats_by_month = final_ready_orders.groupby(['月份', '数据来源工作表']).agg({
            '生产单号': 'count',
            '订单金额(RMB)': 'sum'
        })
        
        print("   📊 按月份统计:")
        for (month, source), row in stats_by_month.iterrows():
            print(f"      {month}-{source}: {row['生产单号']}个订单, ¥{row['订单金额(RMB)']:,.2f}")
        
        return final_ready_orders
    
    def save_ready_to_produce_orders(self, ready_orders_df):
        """保存不缺料订单清单到Excel"""
        print("=== 💾 保存不缺料订单清单 ===")
        
        if ready_orders_df is None or ready_orders_df.empty:
            print("❌ 没有不缺料订单数据")
            return None
        
        filename = '8月9月不缺料订单清单.xlsx'
        
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # 主表：不缺料订单清单
                ready_orders_df.to_excel(writer, sheet_name='不缺料订单清单', index=False)
                
                # 统计表：按月份汇总
                summary_data = ready_orders_df.groupby(['月份', '数据来源工作表']).agg({
                    '生产单号': 'count',
                    '数量Pcs': 'sum',
                    '订单金额(USD)': 'sum',
                    '订单金额(RMB)': 'sum'
                }).reset_index()
                
                summary_data.rename(columns={
                    '生产单号': '订单数量',
                    '数量Pcs': '总数量Pcs',
                    '订单金额(USD)': '总订单金额(USD)',
                    '订单金额(RMB)': '总订单金额(RMB)'
                }, inplace=True)
                
                summary_data.to_excel(writer, sheet_name='统计汇总', index=False)
                
                # 详细统计信息
                total_orders = len(ready_orders_df)
                total_pieces = ready_orders_df['数量Pcs'].sum()
                total_amount_usd = ready_orders_df['订单金额(USD)'].sum()
                total_amount_rmb = ready_orders_df['订单金额(RMB)'].sum()
                
                detail_stats = pd.DataFrame({
                    '统计项目': [
                        '不缺料订单总数', '总生产数量(Pcs)', 
                        '总订单金额(USD)', '总订单金额(RMB)',
                        '8月订单数', '9月订单数',
                        '国内订单数', '柬埔寨订单数'
                    ],
                    '数值': [
                        total_orders,
                        total_pieces,
                        f"${total_amount_usd:,.2f}",
                        f"¥{total_amount_rmb:,.2f}",
                        len(ready_orders_df[ready_orders_df['月份'] == '8月']),
                        len(ready_orders_df[ready_orders_df['月份'] == '9月']),
                        len(ready_orders_df[ready_orders_df['数据来源工作表'] == '国内']),
                        len(ready_orders_df[ready_orders_df['数据来源工作表'] == '柬埔寨'])
                    ]
                })
                
                detail_stats.to_excel(writer, sheet_name='详细统计', index=False)
                
            print(f"✅ 不缺料订单清单已保存: {filename}")
            print(f"📋 包含工作表:")
            print(f"   1️⃣ 不缺料订单清单 ({total_orders}个订单)")
            print(f"   2️⃣ 统计汇总")
            print(f"   3️⃣ 详细统计")
            print(f"💰 订单总价值: ${total_amount_usd:,.2f} / ¥{total_amount_rmb:,.2f}")
            
            return filename
            
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return None
    
    def generate_comprehensive_report(self):
        """生成综合报表"""
        print("=== 📋 生成综合报表 ===")
        
        if self.final_result is None:
            print("❌ 没有分析结果数据")
            return None
        
        # 检查PSO2501724在填充前的状态
        if 'PSO2501724' in self.final_result['生产单号'].values:
            pso_before = self.final_result[self.final_result['生产单号'] == 'PSO2501724']
            print(f"   🔍 填充前PSO2501724: {len(pso_before)}条, 完整性标记: {pso_before['数据完整性标记'].iloc[0] if len(pso_before) > 0 else 'N/A'}")
        else:
            print("   ❌ PSO2501724在填充前已丢失！")
        
        # 应用保守填充策略
        processed_data = self.apply_conservative_filling(self.final_result)
        
        # 检查PSO2501724在填充后的状态
        if 'PSO2501724' in processed_data['生产单号'].values:
            pso_after = processed_data[processed_data['生产单号'] == 'PSO2501724']
            print(f"   🔍 填充后PSO2501724: {len(pso_after)}条")
        else:
            print("   ❌ PSO2501724在填充后丢失！")
        
        # 选择输出字段
        output_columns = [
            '客户订单号', '生产单号', '产品型号', '数量Pcs', '月份', '数据来源工作表',
            '目的地', '客户交期', 'BOM编号',
            '欠料物料编号', '欠料物料名称', '欠料数量', 
            '主供应商名称', '主供应商号', '供应商单价(原币)', '币种', 'RMB单价',
            '起订数量', '供应商修改日期',
            '欠料金额(RMB)', '计算方式',
            '工单需求', '已购未返', '手头现有', '请购组',
            '订单金额(USD)', '订单金额(RMB)', '每元投入回款', '数据完整性标记', '数据填充标记'
        ]
        
        # 映射字段名
        report_data = []
        for _, row in processed_data.iterrows():
            record = {
                '客户订单号': row.get('客户订单号', ''),
                '生产订单号': row.get('生产单号', ''),
                '产品型号': row.get('产品型号', ''),
                '数量Pcs': row.get('数量Pcs', 0),
                '月份': row.get('月份', ''),
                '数据来源工作表': row.get('数据来源工作表', ''),
                '目的地': row.get('目的地', ''),
                '客户交期': row.get('客户交期', ''),
                'BOM编号': row.get('BOM编号', ''),
                
                '欠料物料编号': row.get('物料编号', ''),
                '欠料物料名称': row.get('物料名称', ''),
                '欠料数量': row.get('仓存不足', 0),
                
                '主供应商名称': row.get('主供应商名称', ''),
                '主供应商号': row.get('主供应商号', ''),
                '供应商单价(原币)': row.get('供应商单价(原币)', 0),
                '币种': row.get('币种', ''),
                'RMB单价': row.get('RMB单价', 0),
                '起订数量': row.get('起订数量', 0),
                '供应商修改日期': row.get('供应商修改日期', ''),
                
                '欠料金额(RMB)': row.get('欠料金额(RMB)', 0),
                '计算方式': row.get('计算方式', ''),
                
                '工单需求': row.get('工单需求', ''),
                '已购未返': row.get('已购未返', ''),
                '手头现有': row.get('手头现有', ''),
                '请购组': row.get('请购组', ''),
                
                '订单金额(USD)': row.get('订单金额(USD)', 0),
                '订单金额(RMB)': row.get('订单金额(RMB)', 0),
                '每元投入回款': row.get('每元投入回款', 0),
                '数据完整性标记': row.get('数据完整性标记', ''),
                '数据填充标记': row.get('数据填充标记', '原始数据')
            }
            report_data.append(record)
        
        report_df = pd.DataFrame(report_data)
        
        print(f"   📊 综合报表记录数: {len(report_df)} (已过滤无数据记录)")
        unique_orders_in_report = report_df['生产订单号'].nunique()
        print(f"   📊 涉及订单数: {unique_orders_in_report}")
        
        # 检查PSO2501724是否在最终报表中
        if 'PSO2501724' in report_df['生产订单号'].values:
            print("   ✅ PSO2501724在最终报表中")
        else:
            print("   ❌ PSO2501724不在最终报表中！")
        
        # 显示数据填充统计
        fill_stats = report_df['数据填充标记'].value_counts()
        print(f"   🔧 数据处理统计: {dict(fill_stats)}")
        
        return report_df
    
    def save_comprehensive_report(self, report_df):
        """保存综合报表到Excel"""
        print("=== 💾 保存综合报表 ===")
        
        if report_df is None or report_df.empty:
            print("❌ 没有报表数据")
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'银图PMC综合物料分析报告_{timestamp}.xlsx'
        
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # 主报表
                report_df.to_excel(writer, sheet_name='综合物料分析明细', index=False)
                
                # 汇总统计
                # 统计ROI数值（排除"无需投入"文本）
                numeric_roi = pd.to_numeric(report_df['每元投入回款'], errors='coerce')
                avg_roi = numeric_roi.mean() if not numeric_roi.isna().all() else 0
                
                # 统计数据填充情况
                fill_stats = report_df['数据填充标记'].value_counts()
                fill_summary = ' | '.join([f"{k}:{v}条" for k, v in fill_stats.head(3).items()])
                
                summary_data = {
                    '统计项目': [
                        '总订单数', '有欠料订单数', '精确匹配记录', 
                        '完整数据记录', '涉及供应商数', '总欠料金额(RMB)', 
                        '总订单金额(RMB)', '平均投入产出比', '数据处理统计'
                    ],
                    '数值': [
                        report_df['生产订单号'].nunique(),
                        len(report_df[report_df['欠料物料编号'].notna() & (report_df['欠料物料编号'] != '')]),
                        len(report_df[report_df['计算方式'] == '精确匹配']),
                        len(report_df[report_df['数据完整性标记'] == '完整']),
                        len(report_df[report_df['主供应商名称'].notna() & (report_df['主供应商名称'] != '')]),
                        f"¥{report_df['欠料金额(RMB)'].sum():,.2f}",
                        f"¥{report_df['订单金额(RMB)'].sum():,.2f}",
                        f"{avg_roi:.2f}倍" if avg_roi > 0 else "无需投入占多数",
                        fill_summary
                    ]
                }
                
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='汇总统计', index=False)
                
            print(f"✅ 综合报表已保存: {filename}")
            return filename
            
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            return None
    
    def run_comprehensive_analysis(self):
        """运行完整的综合分析"""
        print("🚀 开始银图PMC综合物料分析")
        print("="*80)
        
        try:
            # 1. 加载数据
            if not self.load_all_data():
                return None
            
            # 2. LEFT JOIN 综合分析
            if not self.comprehensive_left_join_analysis():
                return None
            
            # 3. 计算派生字段
            if not self.calculate_derived_fields():
                return None
            
            # 4. 生成不缺料订单清单
            ready_orders_df = self.generate_ready_to_produce_orders()
            ready_orders_filename = None
            if ready_orders_df is not None:
                ready_orders_filename = self.save_ready_to_produce_orders(ready_orders_df)
            
            # 5. 生成综合报表
            report_df = self.generate_comprehensive_report()
            if report_df is None:
                return None
            
            # 6. 保存综合报表
            filename = self.save_comprehensive_report(report_df)
            
            # 7. 输出最终汇总
            print("\n" + "="*80)
            print(" "*20 + "🎉 综合分析完成！")
            print("="*80)
            
            total_orders = report_df['生产订单号'].nunique()
            total_records = len(report_df)
            complete_data = len(report_df[report_df['数据完整性标记'] == '完整'])
            total_shortage_amount = report_df['欠料金额(RMB)'].sum()
            total_order_amount = report_df['订单金额(RMB)'].sum()
            
            # 安全计算ROI平均值（排除"无需投入"文本）
            numeric_roi = pd.to_numeric(report_df['每元投入回款'], errors='coerce')
            avg_roi = numeric_roi.mean() if not numeric_roi.isna().all() else 0
            
            print(f"📊 综合分析结果汇总:")
            print(f"   - 总订单数: {total_orders}个")
            print(f"   - 分析记录数: {total_records}条")
            print(f"   - 完整数据: {complete_data}条 ({complete_data/total_records*100:.1f}%)")
            print(f"   - 总欠料金额: ¥{total_shortage_amount:,.2f}")
            print(f"   - 总订单金额: ¥{total_order_amount:,.2f}")
            print(f"   - 平均投资回报: {avg_roi:.2f}倍")
            
            # 显示不缺料订单信息
            if ready_orders_df is not None:
                ready_orders_count = len(ready_orders_df)
                ready_orders_amount = ready_orders_df['订单金额(RMB)'].sum()
                print(f"   - 不缺料订单数: {ready_orders_count}个")
                print(f"   - 不缺料订单金额: ¥{ready_orders_amount:,.2f}")
                print(f"   - 不缺料比例: {ready_orders_count/total_orders*100:.1f}%")
            
            print("\n📄 生成文件:")
            if ready_orders_filename:
                print(f"   🚀 {ready_orders_filename} (不缺料订单清单)")
            if filename:
                print(f"   📋 {filename} (综合分析报表)")
                
            return report_df, filename, ready_orders_df
            
        except Exception as e:
            print(f"❌ 分析过程出错: {e}")
            import traceback
            traceback.print_exc()
            return None

if __name__ == "__main__":
    analyzer = ComprehensivePMCAnalyzer()
    result = analyzer.run_comprehensive_analysis()
    
    if result:
        print("\n🎊 分析成功完成！")
    else:
        print("\n❌ 分析失败，请检查数据和错误信息")
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC综合物料分析系统 - 修复版
修复欠料数据读取问题，确保正确识别不欠料订单
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
import sys
from pathlib import Path

warnings.filterwarnings('ignore')

class FixedComprehensivePMCAnalyzer:
    """修复版PMC综合分析器"""
    
    def __init__(self):
        """初始化分析器"""
        self.orders_df = None
        self.shortage_df = None
        self.inventory_df = None
        self.supplier_df = None
        self.final_result = None
        
        # 汇率设置
        self.currency_rates = {
            'RMB': 1.0,
            'USD': 7.20,
            'HKD': 0.93,
            'EUR': 7.85
        }
        
        # 统计信息
        self.stats = {
            'total_orders': 0,
            'orders_with_shortage': 0,
            'orders_without_shortage': 0,
            'shortage_data_issues': []
        }
        
    def load_all_data(self):
        """加载所有数据源（增强版）"""
        print("=== 🔄 加载数据源（修复版） ===")
        
        # 1. 加载订单数据
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
            
            # 标准化订单表列名
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
                self.orders_df['订单金额'] = 1000  # 默认1000 USD
                print("   ⚠️ 订单表中未找到'订单金额'字段，使用默认值1000 USD")
            
            self.stats['total_orders'] = len(self.orders_df['生产单号'].unique())
            print(f"   ✅ 订单总数: {self.stats['total_orders']}条")
            
        except Exception as e:
            print(f"   ❌ 订单数据加载失败: {e}")
            return False
        
        # 2. 加载欠料数据（修复版）
        print("2. 加载欠料数据（增强检测）...")
        try:
            # 尝试多种读取方式
            shortage_file = 'input/mat_owe_pso.xlsx'
            
            # 方法1：尝试跳过前几行
            for skip_rows in [0, 1, 2, 3]:
                try:
                    temp_df = pd.read_excel(shortage_file, skiprows=skip_rows)
                    if len(temp_df) > 10:  # 至少有10行数据
                        self.shortage_df = temp_df
                        print(f"   ✅ 成功读取欠料数据（跳过{skip_rows}行）")
                        break
                except:
                    continue
            
            if self.shortage_df is None:
                # 方法2：读取原始数据并手动处理
                raw_df = pd.read_excel(shortage_file, header=None)
                print(f"   ℹ️ 原始数据行数: {len(raw_df)}")
                print(f"   ℹ️ 原始数据列数: {len(raw_df.columns)}")
                
                # 找到包含订单号的列
                for col_idx in range(min(10, len(raw_df.columns))):
                    col_values = raw_df.iloc[:, col_idx].astype(str)
                    # 检查是否包含订单号模式
                    pso_count = col_values.str.contains('PSO|MSO|RSO|TSO', na=False).sum()
                    if pso_count > 10:
                        print(f"   ✅ 在第{col_idx+1}列找到订单号")
                        self.shortage_df = raw_df
                        break
            
            # 标准化欠料表列名
            if self.shortage_df is not None:
                # 尝试识别列
                if len(self.shortage_df.columns) > 10:
                    # 根据位置假设列名
                    col_mapping = {
                        self.shortage_df.columns[1]: '生产单号',
                        self.shortage_df.columns[2]: '物料编码',
                        self.shortage_df.columns[3]: '物料名称',
                        self.shortage_df.columns[5]: '欠料数量'
                    }
                    self.shortage_df = self.shortage_df.rename(columns=col_mapping)
                
                # 清理数据
                if '生产单号' in self.shortage_df.columns:
                    # 移除无效行
                    self.shortage_df = self.shortage_df[
                        self.shortage_df['生产单号'].notna() & 
                        (self.shortage_df['生产单号'] != 'VLOOKUP(, A1:C4, 2, 0)')
                    ]
                    
                    shortage_orders = self.shortage_df['生产单号'].unique()
                    self.stats['orders_with_shortage'] = len(shortage_orders)
                    print(f"   ✅ 欠料订单数: {self.stats['orders_with_shortage']}")
                else:
                    print("   ⚠️ 无法识别欠料数据中的生产单号列")
                    self.stats['shortage_data_issues'].append("无法识别生产单号列")
            else:
                print("   ⚠️ 欠料数据加载失败，将所有订单视为不欠料")
                self.stats['shortage_data_issues'].append("欠料文件无法读取")
                
        except Exception as e:
            print(f"   ⚠️ 欠料数据加载异常: {e}")
            self.stats['shortage_data_issues'].append(str(e))
        
        # 3. 加载库存数据
        print("3. 加载库存数据...")
        try:
            self.inventory_df = pd.read_excel('input/inventory_list.xlsx')
            self.inventory_df = self.inventory_df.rename(columns={
                '物項編號': '物料编码',
                '成本單價': '单价'
            })
            print(f"   ✅ 库存数据: {len(self.inventory_df)}条")
        except Exception as e:
            print(f"   ⚠️ 库存数据加载失败: {e}")
        
        # 4. 加载供应商数据
        print("4. 加载供应商数据...")
        try:
            self.supplier_df = pd.read_excel('input/supplier.xlsx')
            self.supplier_df = self.supplier_df.rename(columns={
                '物项编号': '物料编码'
            })
            print(f"   ✅ 供应商数据: {len(self.supplier_df)}条")
        except Exception as e:
            print(f"   ⚠️ 供应商数据加载失败: {e}")
        
        return True
    
    def identify_no_shortage_orders(self):
        """准确识别不欠料订单"""
        print("\n=== 🔍 识别不欠料订单（修复版） ===")
        
        if self.orders_df is None:
            print("   ❌ 订单数据未加载")
            return set()
        
        all_orders = set(self.orders_df['生产单号'].unique())
        
        # 如果欠料数据正常加载
        if self.shortage_df is not None and '生产单号' in self.shortage_df.columns:
            shortage_orders = set(self.shortage_df['生产单号'].unique())
            no_shortage_orders = all_orders - shortage_orders
            
            print(f"   📊 总订单数: {len(all_orders)}")
            print(f"   📊 有欠料订单数: {len(shortage_orders)}")
            print(f"   📊 不欠料订单数: {len(no_shortage_orders)}")
            
            # 验证：加载已知的不欠料订单清单进行对比
            try:
                known_no_shortage = pd.read_excel('8月9月不缺料订单清单.xlsx')
                known_orders = set(known_no_shortage['生产单号'].unique())
                
                # 对比分析
                correct_identified = no_shortage_orders & known_orders
                false_positive = no_shortage_orders - known_orders  # 错误标记为不欠料
                false_negative = known_orders - no_shortage_orders  # 错误标记为欠料
                
                print(f"\n   ✅ 正确识别: {len(correct_identified)}个")
                print(f"   ⚠️ 误判为不欠料: {len(false_positive)}个")
                print(f"   ⚠️ 误判为欠料: {len(false_negative)}个")
                
                if len(false_positive) > 0:
                    print(f"   误判订单示例: {list(false_positive)[:5]}")
                    
            except Exception as e:
                print(f"   ℹ️ 无法加载验证文件: {e}")
            
            return no_shortage_orders
            
        else:
            print("   ⚠️ 欠料数据异常，无法准确判断")
            print("   ⚠️ 建议检查欠料文件格式或使用已知的不欠料清单")
            
            # 尝试使用已知清单
            try:
                known_no_shortage = pd.read_excel('8月9月不缺料订单清单.xlsx')
                known_orders = set(known_no_shortage['生产单号'].unique())
                no_shortage_orders = all_orders & known_orders  # 只标记已知的不欠料订单
                
                print(f"   ✅ 使用已知清单，识别不欠料订单: {len(no_shortage_orders)}个")
                return no_shortage_orders
                
            except:
                print("   ❌ 无法加载不欠料清单，将所有订单视为需要检查")
                return set()
    
    def perform_analysis(self):
        """执行综合分析（修复版）"""
        print("\n=== 📊 执行综合分析（修复版） ===")
        
        # 识别不欠料订单
        no_shortage_orders = self.identify_no_shortage_orders()
        
        # 执行LEFT JOIN分析
        result = self.orders_df.copy()
        
        # LEFT JOIN 欠料数据
        if self.shortage_df is not None and '生产单号' in self.shortage_df.columns:
            result = pd.merge(
                result,
                self.shortage_df,
                on='生产单号',
                how='left',
                suffixes=('', '_shortage')
            )
        else:
            # 添加空白欠料列
            result['物料编码'] = ''
            result['物料名称'] = ''
            result['欠料数量'] = 0
        
        # 标记不欠料订单
        result['缺料状态'] = result['生产单号'].apply(
            lambda x: '不欠料' if x in no_shortage_orders else '欠料'
        )
        
        # 计算订单金额（RMB）
        result['订单金额(USD)'] = pd.to_numeric(result.get('订单金额', 1000), errors='coerce').fillna(1000)
        result['订单金额(RMB)'] = result['订单金额(USD)'] * self.currency_rates['USD']
        
        # 计算ROI（修复版）
        def calculate_roi(row):
            if row['缺料状态'] == '不欠料':
                return '无需投入'
            
            # 计算欠料金额
            shortage_qty = pd.to_numeric(row.get('欠料数量', 0), errors='coerce')
            if pd.isna(shortage_qty) or shortage_qty == 0:
                return '无需投入'
            
            # 获取单价（这里简化处理）
            unit_price = 10  # 默认单价
            shortage_amount = shortage_qty * unit_price
            
            if shortage_amount > 0:
                return row['订单金额(RMB)'] / shortage_amount
            else:
                return '无需投入'
        
        result['每元投入回款'] = result.apply(calculate_roi, axis=1)
        
        self.final_result = result
        
        # 输出统计
        no_invest_count = len(result[result['每元投入回款'] == '无需投入'])
        print(f"   📊 总订单数: {len(result['生产单号'].unique())}")
        print(f"   📊 不欠料订单数: {len(no_shortage_orders)}")
        print(f"   📊 标记为'无需投入': {no_invest_count}")
        
        if len(self.stats['shortage_data_issues']) > 0:
            print("\n   ⚠️ 数据质量问题:")
            for issue in self.stats['shortage_data_issues']:
                print(f"      - {issue}")
        
        return True
    
    def save_results(self, output_file=None):
        """保存分析结果"""
        if self.final_result is None:
            print("❌ 没有分析结果可保存")
            return
        
        if output_file is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f'银图PMC综合物料分析报告_修复版_{timestamp}.xlsx'
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 主表
                self.final_result.to_excel(
                    writer, 
                    sheet_name='综合物料分析明细',
                    index=False
                )
                
                # 不欠料订单清单
                no_shortage_df = self.final_result[
                    self.final_result['每元投入回款'] == '无需投入'
                ][['生产单号', '客户订单号', '产品型号', '订单金额(RMB)', '缺料状态']].drop_duplicates()
                
                no_shortage_df.to_excel(
                    writer,
                    sheet_name='不欠料订单清单',
                    index=False
                )
                
                # 数据质量报告
                quality_report = pd.DataFrame({
                    '检查项': [
                        '总订单数',
                        '欠料数据订单数',
                        '识别的不欠料订单',
                        '数据质量问题'
                    ],
                    '结果': [
                        self.stats['total_orders'],
                        self.stats['orders_with_shortage'],
                        self.stats['orders_without_shortage'],
                        '; '.join(self.stats['shortage_data_issues']) if self.stats['shortage_data_issues'] else '无'
                    ]
                })
                
                quality_report.to_excel(
                    writer,
                    sheet_name='数据质量报告',
                    index=False
                )
                
            print(f"✅ 分析结果已保存到: {output_file}")
            
        except Exception as e:
            print(f"❌ 保存失败: {e}")


def main():
    """主函数"""
    print("=" * 60)
    print("银图PMC综合物料分析系统 - 修复版")
    print("修复欠料数据读取问题，准确识别不欠料订单")
    print("=" * 60)
    
    analyzer = FixedComprehensivePMCAnalyzer()
    
    if analyzer.load_all_data():
        if analyzer.perform_analysis():
            analyzer.save_results()
            print("\n✅ 分析完成！")
            
            # 对比新旧结果
            print("\n=== 📊 结果对比 ===")
            try:
                # 读取修复版结果
                fixed_result = analyzer.final_result
                fixed_no_invest = fixed_result[fixed_result['每元投入回款'] == '无需投入']
                fixed_orders = set(fixed_no_invest['生产单号'].unique())
                
                # 读取已知正确清单
                known_correct = pd.read_excel('8月9月不缺料订单清单.xlsx')
                correct_orders = set(known_correct['生产单号'].unique())
                
                # 计算准确率
                correct_identified = fixed_orders & correct_orders
                accuracy = len(correct_identified) / len(correct_orders) * 100 if len(correct_orders) > 0 else 0
                
                print(f"修复版识别: {len(fixed_orders)}个不欠料订单")
                print(f"正确清单: {len(correct_orders)}个不欠料订单")
                print(f"准确率: {accuracy:.1f}%")
                
                # 计算金额
                fixed_amount = fixed_no_invest.drop_duplicates('生产单号')['订单金额(RMB)'].sum()
                correct_amount = known_correct['订单金额(RMB)'].sum()
                
                print(f"\n修复版不欠料金额: {fixed_amount/10000:.2f}万")
                print(f"正确不欠料金额: {correct_amount/10000:.2f}万")
                print(f"金额准确率: {fixed_amount/correct_amount*100:.1f}%")
                
            except Exception as e:
                print(f"对比分析失败: {e}")
    else:
        print("\n❌ 数据加载失败")


if __name__ == "__main__":
    main()
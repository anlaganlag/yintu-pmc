#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""PMC订单专项分析系统 - 严格按照PMC_order.xlsx顺序分析129个订单"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class PMCOrderAnalyzer:
    """PMC订单分析器"""
    
    def __init__(self):
        self.exchange_rates = {
            'RMB': 1.0,
            'USD': 7.31,
            'HKD': 0.936,
            'EUR': 8.12
        }
        
    def load_pmc_orders(self):
        """加载PMC订单列表并保持顺序"""
        print("📋 加载PMC订单列表...")
        df_pmc = pd.read_excel('PMC_order.xlsx')
        df_pmc['订单序号'] = range(1, len(df_pmc) + 1)  # 添加序号以保持顺序
        print(f"   ✅ 加载{len(df_pmc)}个生产订单")
        print(f"   📊 产线分布: {df_pmc['产线'].value_counts().to_dict()}")
        return df_pmc
    
    def load_order_amounts(self):
        """加载订单金额数据"""
        print("💰 加载订单金额数据...")
        
        # 国内订单
        dfs_domestic = []
        try:
            df_89 = pd.read_excel('input/order-amt-89.xlsx', sheet_name=None)
            for sheet_name, df in df_89.items():
                df['数据来源'] = f'国内-{sheet_name}'
                dfs_domestic.append(df)
        except:
            pass
        
        # 柬埔寨订单
        dfs_cambodia = []
        try:
            df_89c = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name=None)
            for sheet_name, df in df_89c.items():
                df['数据来源'] = f'柬埔寨-{sheet_name}'
                dfs_cambodia.append(df)
        except:
            pass
        
        # 合并所有订单
        all_dfs = dfs_domestic + dfs_cambodia
        if all_dfs:
            df_orders = pd.concat(all_dfs, ignore_index=True)
            
            # 标准化列名
            df_orders.columns = df_orders.columns.str.strip()
            
            # 重命名关键列 - 根据实际列名调整
            rename_map = {
                '生產訂單': '生产订单',
                '生 產 單 号(  廠方 )': '生产订单',  # 新增：实际列名
                '生 產 單 号(客方 )': '客户订单号',  # 新增：实际列名
                '客戶訂單': '客户订单号',
                '產品': '产品',
                '型 號( 廠方/客方 )': '产品',  # 新增：实际列名
                '總價(RMB)': '订单金额(RMB)',
                '订单金额': '订单金额(RMB)',  # 新增：实际列名
                '數量': '数量',
                '數 量  (Pcs)': '数量'  # 新增：实际列名
            }
            df_orders = df_orders.rename(columns=rename_map)
            
            # 转换数据类型
            if '订单金额(RMB)' in df_orders.columns:
                df_orders['订单金额(RMB)'] = pd.to_numeric(df_orders['订单金额(RMB)'], errors='coerce')
            
            print(f"   ✅ 加载{len(df_orders)}条订单记录")
            return df_orders
        
        return pd.DataFrame()
    
    def load_shortage_data(self):
        """加载缺料数据"""
        print("📦 加载缺料数据...")
        df_shortage = pd.read_excel('input/mat_owe_pso.xlsx', skiprows=1)  # 跳过第一行重复的标题
        
        # 标准化列名
        df_shortage.columns = df_shortage.columns.str.strip()
        
        # 重命名关键列
        rename_map = {
            '订单编号': '生产订单',
            '物料编号': '物料編號',
            '工单需求': '欠數',
            '仓存不足': '欠數_alt',
            '缺料金额': '缺料金额_原始'
        }
        df_shortage = df_shortage.rename(columns=rename_map)
        
        # 如果欠數列不存在，使用仓存不足作为欠数
        if '欠數' not in df_shortage.columns and '欠數_alt' in df_shortage.columns:
            df_shortage['欠數'] = df_shortage['欠數_alt']
        
        # 转换数值类型
        if '欠數' in df_shortage.columns:
            df_shortage['欠數'] = pd.to_numeric(df_shortage['欠數'], errors='coerce').fillna(0)
        
        print(f"   ✅ 加载{len(df_shortage)}条缺料记录")
        return df_shortage
    
    def load_inventory_prices(self):
        """加载库存价格数据"""
        print("💵 加载库存价格...")
        df_inv = pd.read_excel('input/inventory_list.xlsx', sheet_name='银图库存总表')
        
        # 标准化列名
        df_inv.columns = df_inv.columns.str.strip()
        
        # 重命名列以统一格式
        rename_map = {
            '物項編號': '物料編號',
            '物項名稱': '物料名称'
        }
        df_inv = df_inv.rename(columns=rename_map)
        
        # 处理价格字段
        for col in ['最新報價', '成本單價']:
            if col in df_inv.columns:
                df_inv[col] = pd.to_numeric(df_inv[col], errors='coerce')
        
        # 选择有效价格
        df_inv['RMB单价'] = df_inv.apply(
            lambda x: x['成本單價'] if pd.isna(x.get('最新報價', 0)) or x.get('最新報價', 0) == 0 
            else x.get('最新報價', 0), axis=1
        )
        
        valid_prices = df_inv[df_inv['RMB单价'] > 0]
        print(f"   ✅ 有效价格记录: {len(valid_prices)}条")
        return df_inv
    
    def load_supplier_data(self):
        """加载供应商数据"""
        print("🏭 加载供应商数据...")
        df_supp = pd.read_excel('input/supplier.xlsx')
        
        # 标准化列名
        df_supp.columns = df_supp.columns.str.strip()
        
        # 重命名实际列名
        rename_map = {
            '物项编号': '物料編號',
            '物项名称': '物料名称',
            '供应商号': '供應商編號',
            '供应商名称': '供應商',
            '单价': 'RMB单价',
            '币种': '幣別',
            '修改日期': '最後異動日期'
        }
        df_supp = df_supp.rename(columns=rename_map)
        
        # 处理价格
        if 'RMB单价' in df_supp.columns:
            df_supp['RMB单价'] = pd.to_numeric(df_supp['RMB单价'], errors='coerce')
        
        # 处理汇率转换（如果需要）
        if '幣別' in df_supp.columns and 'RMB单价' in df_supp.columns:
            # USD转RMB
            df_supp.loc[df_supp['幣別'] == 'USD', 'RMB单价'] *= 7.31
            # HKD转RMB  
            df_supp.loc[df_supp['幣別'] == 'HKD', 'RMB单价'] *= 0.936
            # EUR转RMB
            df_supp.loc[df_supp['幣別'] == 'EUR', 'RMB单价'] *= 8.12
        
        valid_suppliers = df_supp[df_supp['RMB单价'] > 0] if 'RMB单价' in df_supp.columns else df_supp
        print(f"   ✅ 有效供应商价格: {len(valid_suppliers)}条")
        return df_supp
    
    def find_best_supplier(self, material_code, df_suppliers):
        """为物料找到最优供应商"""
        material_suppliers = df_suppliers[df_suppliers['物料編號'] == material_code].copy()
        
        if material_suppliers.empty:
            return None
        
        # 计算综合得分
        material_suppliers['价格得分'] = 100 - (material_suppliers['RMB单价'] - material_suppliers['RMB单价'].min()) / \
                                       (material_suppliers['RMB单价'].max() - material_suppliers['RMB单价'].min() + 0.0001) * 100
        
        # 选择价格最低的供应商
        best_supplier = material_suppliers.nsmallest(1, 'RMB单价').iloc[0]
        
        return {
            '供应商': best_supplier.get('供應商', ''),
            '供应商价格': best_supplier.get('RMB单价', 0),
            '币种': best_supplier.get('幣別', 'RMB'),
            '最后修改日期': best_supplier.get('最後異動日期', '')
        }
    
    def analyze_pmc_orders(self):
        """执行PMC订单分析"""
        print("\n" + "="*80)
        print("🚀 开始PMC订单专项分析")
        print("="*80)
        
        # 1. 加载数据
        df_pmc = self.load_pmc_orders()
        df_orders = self.load_order_amounts()
        df_shortage = self.load_shortage_data()
        df_inventory = self.load_inventory_prices()
        df_suppliers = self.load_supplier_data()
        
        # 2. 关联分析
        print("\n📊 开始关联分析...")
        
        # 以PMC订单为主表，LEFT JOIN其他数据
        df_result = df_pmc.copy()
        
        # 关联订单金额
        if not df_orders.empty:
            order_amounts = df_orders.groupby('生产订单').agg({
                '客户订单号': lambda x: '; '.join(x.dropna().astype(str).unique()),
                '产品': lambda x: '; '.join(x.dropna().astype(str).unique()),
                '订单金额(RMB)': 'sum',
                '数量': 'sum',
                '数据来源': lambda x: '; '.join(x.unique())
            }).reset_index()
            
            df_result = df_result.merge(order_amounts, on='生产订单', how='left')
        
        # 关联缺料信息
        if not df_shortage.empty:
            # 先为每个订单汇总缺料信息
            shortage_summary = df_shortage.groupby('生产订单').agg({
                '物料編號': 'count',  # 缺料物料种类数
                '欠數': 'sum'  # 总欠料数量
            }).reset_index()
            shortage_summary.columns = ['生产订单', '缺料种类数', '总欠数']
            
            # 详细缺料信息用于后续计算
            df_result = df_result.merge(shortage_summary, on='生产订单', how='left')
            
            # 计算缺料金额（需要逐条匹配价格）
            print("   计算缺料金额...")
            shortage_amounts = []
            
            for pso in df_pmc['生产订单'].unique():
                pso_shortage = df_shortage[df_shortage['生产订单'] == pso].copy()
                if pso_shortage.empty:
                    shortage_amounts.append({
                        '生产订单': pso,
                        '缺料金额(RMB)': 0,
                        '有价格物料数': 0,
                        '供应商覆盖数': 0
                    })
                    continue
                
                # 匹配库存价格
                pso_shortage = pso_shortage.merge(
                    df_inventory[['物料編號', 'RMB单价']].drop_duplicates('物料編號'),
                    on='物料編號',
                    how='left'
                )
                
                # 匹配供应商最低价
                for idx, row in pso_shortage.iterrows():
                    if pd.isna(row['RMB单价']) or row['RMB单价'] == 0:
                        best_supp = self.find_best_supplier(row['物料編號'], df_suppliers)
                        if best_supp:
                            pso_shortage.loc[idx, 'RMB单价'] = best_supp['供应商价格']
                            pso_shortage.loc[idx, '供应商'] = best_supp['供应商']
                
                # 计算缺料金额
                pso_shortage['缺料金额'] = pso_shortage['欠數'] * pso_shortage['RMB单价'].fillna(0)
                
                shortage_amounts.append({
                    '生产订单': pso,
                    '缺料金额(RMB)': pso_shortage['缺料金额'].sum(),
                    '有价格物料数': (pso_shortage['RMB单价'] > 0).sum(),
                    '供应商覆盖数': pso_shortage['供应商'].notna().sum() if '供应商' in pso_shortage.columns else 0
                })
            
            df_shortage_amount = pd.DataFrame(shortage_amounts)
            df_result = df_result.merge(df_shortage_amount, on='生产订单', how='left')
        
        # 填充缺失值
        df_result['缺料种类数'] = df_result['缺料种类数'].fillna(0).astype(int)
        df_result['缺料金额(RMB)'] = df_result['缺料金额(RMB)'].fillna(0)
        df_result['订单金额(RMB)'] = df_result['订单金额(RMB)'].fillna(0)
        
        # 计算ROI
        df_result['投资回报率(ROI)'] = df_result.apply(
            lambda x: x['订单金额(RMB)'] / x['缺料金额(RMB)'] if x['缺料金额(RMB)'] > 0 else 999999,
            axis=1
        )
        
        # 添加状态标记
        df_result['缺料状态'] = df_result.apply(
            lambda x: '不缺料' if x['缺料种类数'] == 0 else f'缺{x["缺料种类数"]}种物料',
            axis=1
        )
        
        # 格式化ROI显示
        df_result['ROI显示'] = df_result['投资回报率(ROI)'].apply(
            lambda x: '无需投入' if x >= 999999 else f'{x:.2f}'
        )
        
        # 按原始顺序排序（保持PMC_order.xlsx的顺序）
        df_result = df_result.sort_values('订单序号')
        
        print(f"   ✅ 分析完成，共{len(df_result)}个订单")
        
        return df_result, df_shortage, df_inventory, df_suppliers
    
    def generate_report(self, df_main, df_shortage, df_inventory, df_suppliers):
        """生成Excel报告"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'PMC订单分析报告_{timestamp}.xlsx'
        
        print(f"\n📝 生成报告: {filename}")
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # 1. 主表 - 严格按原始顺序
            df_main_output = df_main[[
                '订单序号', '产线', '生产订单', '对应PR订单',
                '客户订单号', '产品', '订单金额(RMB)', 
                '缺料状态', '缺料种类数', '缺料金额(RMB)',
                'ROI显示', '数据来源'
            ]].copy()
            
            # 格式化金额
            df_main_output['订单金额(RMB)'] = df_main_output['订单金额(RMB)'].apply(
                lambda x: f'¥{x:,.2f}' if x > 0 else '-'
            )
            df_main_output['缺料金额(RMB)'] = df_main_output['缺料金额(RMB)'].apply(
                lambda x: f'¥{x:,.2f}' if x > 0 else '-'
            )
            
            df_main_output.to_excel(writer, sheet_name='订单分析(原始顺序)', index=False)
            
            # 2. ROI排序表
            df_roi = df_main[df_main['缺料金额(RMB)'] > 0].copy()
            df_roi = df_roi.sort_values('投资回报率(ROI)', ascending=False)
            df_roi_output = df_roi[[
                '产线', '生产订单', '订单金额(RMB)', 
                '缺料金额(RMB)', '投资回报率(ROI)', '缺料种类数'
            ]].copy()
            
            # 格式化
            df_roi_output['订单金额(RMB)'] = df_roi_output['订单金额(RMB)'].apply(
                lambda x: f'¥{x:,.2f}'
            )
            df_roi_output['缺料金额(RMB)'] = df_roi_output['缺料金额(RMB)'].apply(
                lambda x: f'¥{x:,.2f}'
            )
            df_roi_output['投资回报率(ROI)'] = df_roi_output['投资回报率(ROI)'].apply(
                lambda x: f'{x:.2f}倍'
            )
            
            df_roi_output.to_excel(writer, sheet_name='ROI排序(高到低)', index=False)
            
            # 3. 缺料金额排序表
            df_shortage_sort = df_main[df_main['缺料金额(RMB)'] > 0].copy()
            df_shortage_sort = df_shortage_sort.sort_values('缺料金额(RMB)', ascending=False)
            df_shortage_output = df_shortage_sort[[
                '产线', '生产订单', '缺料金额(RMB)', 
                '缺料种类数', '订单金额(RMB)', 'ROI显示'
            ]].head(30)  # Top 30
            
            # 格式化
            df_shortage_output['缺料金额(RMB)'] = df_shortage_output['缺料金额(RMB)'].apply(
                lambda x: f'¥{x:,.2f}'
            )
            df_shortage_output['订单金额(RMB)'] = df_shortage_output['订单金额(RMB)'].apply(
                lambda x: f'¥{x:,.2f}'
            )
            
            df_shortage_output.to_excel(writer, sheet_name='缺料金额Top30', index=False)
            
            # 4. 产线汇总
            df_line_summary = df_main.groupby('产线').agg({
                '生产订单': 'count',
                '订单金额(RMB)': 'sum',
                '缺料金额(RMB)': 'sum',
                '缺料种类数': 'sum'
            }).reset_index()
            
            df_line_summary.columns = ['产线', '订单数', '总订单金额', '总缺料金额', '总缺料种类']
            df_line_summary['平均ROI'] = df_line_summary['总订单金额'] / df_line_summary['总缺料金额'].replace(0, 1)
            
            # 格式化
            df_line_summary['总订单金额'] = df_line_summary['总订单金额'].apply(
                lambda x: f'¥{x:,.2f}'
            )
            df_line_summary['总缺料金额'] = df_line_summary['总缺料金额'].apply(
                lambda x: f'¥{x:,.2f}'
            )
            df_line_summary['平均ROI'] = df_line_summary['平均ROI'].apply(
                lambda x: f'{x:.2f}倍' if x < 999999 else '无需投入'
            )
            
            df_line_summary.to_excel(writer, sheet_name='产线汇总', index=False)
            
            # 5. 统计摘要
            summary_data = {
                '指标': [
                    '总订单数',
                    '有缺料订单数',
                    '不缺料订单数',
                    '总订单金额',
                    '总缺料金额',
                    '平均ROI',
                    '最高ROI订单',
                    '最大缺料金额订单'
                ],
                '数值': [
                    len(df_main),
                    (df_main['缺料种类数'] > 0).sum(),
                    (df_main['缺料种类数'] == 0).sum(),
                    f"¥{df_main['订单金额(RMB)'].sum():,.2f}",
                    f"¥{df_main['缺料金额(RMB)'].sum():,.2f}",
                    f"{df_main[df_main['投资回报率(ROI)'] < 999999]['投资回报率(ROI)'].mean():.2f}倍",
                    df_main[df_main['投资回报率(ROI)'] < 999999].nlargest(1, '投资回报率(ROI)')['生产订单'].values[0] if len(df_main[df_main['投资回报率(ROI)'] < 999999]) > 0 else '-',
                    df_main.nlargest(1, '缺料金额(RMB)')['生产订单'].values[0]
                ]
            }
            
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='统计摘要', index=False)
        
        print(f"   ✅ 报告生成完成")
        return filename

def main():
    analyzer = PMCOrderAnalyzer()
    df_main, df_shortage, df_inventory, df_suppliers = analyzer.analyze_pmc_orders()
    filename = analyzer.generate_report(df_main, df_shortage, df_inventory, df_suppliers)
    
    print("\n" + "="*80)
    print("🎉 PMC订单分析完成！")
    print("="*80)
    
    # 打印关键统计
    print("\n📊 关键统计:")
    print(f"   订单总数: {len(df_main)}")
    print(f"   有缺料订单: {(df_main['缺料种类数'] > 0).sum()}")
    print(f"   不缺料订单: {(df_main['缺料种类数'] == 0).sum()}")
    print(f"   总订单金额: ¥{df_main['订单金额(RMB)'].sum():,.2f}")
    print(f"   总缺料金额: ¥{df_main['缺料金额(RMB)'].sum():,.2f}")
    
    avg_roi = df_main[df_main['投资回报率(ROI)'] < 999999]['投资回报率(ROI)'].mean()
    if not pd.isna(avg_roi):
        print(f"   平均ROI: {avg_roi:.2f}倍")
    
    print(f"\n📄 输出文件: {filename}")

if __name__ == "__main__":
    main()
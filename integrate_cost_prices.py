#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""整合补充的成本价信息到PMC报告"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def integrate_cost_prices():
    """整合新的成本价信息"""
    print("=" * 80)
    print("🔄 开始整合补充的成本价信息")
    print("=" * 80)
    
    # 1. 读取各数据源
    print("\n📋 加载数据文件...")
    
    # 读取PMC订单
    df_pmc = pd.read_excel('PMC_order.xlsx')
    pmc_orders = set(df_pmc['生产订单'].unique())
    print(f"   ✅ PMC订单: {len(pmc_orders)}个")
    
    # 读取补充成本价
    df_cost_new = pd.read_excel('qw訂單有欠料無成本價250901v3.xlsx')
    print(f"   ✅ 补充成本价记录: {len(df_cost_new)}条")
    
    # 读取原始缺料数据
    df_shortage = pd.read_excel('input/mat_owe_pso.xlsx', skiprows=1)
    df_shortage.columns = df_shortage.columns.str.strip()
    print(f"   ✅ 原始缺料记录: {len(df_shortage)}条")
    
    # 读取原始库存价格
    df_inventory = pd.read_excel('input/inventory_list.xlsx', sheet_name='银图库存总表')
    df_inventory.columns = df_inventory.columns.str.strip()
    
    # 读取原始供应商价格
    df_supplier = pd.read_excel('input/supplier.xlsx')
    df_supplier.columns = df_supplier.columns.str.strip()
    
    # 2. 分析补充成本价的适用性
    print("\n📊 分析补充成本价适用性...")
    
    # 找出PMC订单中可以使用补充成本价的订单
    cost_orders = set(df_cost_new['production_order_no'].unique())
    matched_orders = pmc_orders.intersection(cost_orders)
    
    print(f"   匹配订单数: {len(matched_orders)}个 (占PMC订单{len(matched_orders)/len(pmc_orders)*100:.1f}%)")
    
    # 筛选适用的补充成本价记录
    df_cost_matched = df_cost_new[df_cost_new['production_order_no'].isin(matched_orders)].copy()
    df_cost_matched = df_cost_matched[df_cost_matched['Item_price'] > 0]  # 只使用有效价格
    
    print(f"   有效补充记录: {len(df_cost_matched)}条")
    
    # 3. 创建价格查找表（优先级：补充成本价 > 库存价 > 供应商价）
    print("\n💰 构建综合价格表...")
    
    # 准备补充成本价表
    price_supplement = {}
    for _, row in df_cost_matched.iterrows():
        key = (row['production_order_no'], row['mat'])
        price_supplement[key] = {
            '单价': row['Item_price'],
            '来源': '补充成本价',
            '说明': row.get('說明', '')
        }
    
    print(f"   补充价格条目: {len(price_supplement)}个")
    
    # 4. 重新计算缺料金额
    print("\n🔧 重新计算缺料金额...")
    
    # 重命名缺料表列名
    df_shortage = df_shortage.rename(columns={
        '订单编号': '生产订单',
        '物料编号': '物料編號',
        '物料名称': '物料名称',
        '工单需求': '欠数',
        '仓存不足': '欠数_alt',
        '供应商名称': '供应商',
        'OTS期': '客户交期'
    })
    
    # 如果欠数不存在，使用仓存不足
    if '欠数' not in df_shortage.columns and '欠数_alt' in df_shortage.columns:
        df_shortage['欠数'] = df_shortage['欠数_alt']
    
    # 只处理PMC订单
    df_shortage_pmc = df_shortage[df_shortage['生产订单'].isin(pmc_orders)].copy()
    
    # 计算缺料金额
    enhanced_records = []
    price_source_stats = {'补充成本价': 0, '库存价': 0, '供应商价': 0, '无价格': 0}
    
    for _, row in df_shortage_pmc.iterrows():
        pso = row['生产订单']
        material = row['物料編號']
        shortage_qty = pd.to_numeric(row.get('欠数', 0), errors='coerce')
        
        if pd.isna(shortage_qty) or shortage_qty <= 0:
            continue
        
        # 查找价格（优先级顺序）
        price = 0
        price_source = '无价格'
        
        # 1. 尝试使用补充成本价
        if (pso, material) in price_supplement:
            price = price_supplement[(pso, material)]['单价']
            price_source = '补充成本价'
            price_source_stats['补充成本价'] += 1
        
        # 2. 尝试使用库存价格
        elif material in df_inventory['物項編號'].values:
            inv_row = df_inventory[df_inventory['物項編號'] == material].iloc[0]
            # 优先使用最新报价，否则使用成本单价
            if pd.notna(inv_row.get('最新報價')) and inv_row.get('最新報價') > 0:
                price = inv_row['最新報價']
            elif pd.notna(inv_row.get('成本單價')) and inv_row.get('成本單價') > 0:
                price = inv_row['成本單價']
            if price > 0:
                price_source = '库存价'
                price_source_stats['库存价'] += 1
        
        # 3. 尝试使用供应商价格
        if price == 0:
            # 重命名供应商表列名
            supplier_cols = {
                '物项编号': '物料編號',
                '单价': '报价',
                '币种': '币别'
            }
            df_supplier_renamed = df_supplier.rename(columns=supplier_cols)
            
            if material in df_supplier_renamed['物料編號'].values:
                supp_rows = df_supplier_renamed[df_supplier_renamed['物料編號'] == material]
                supp_rows = supp_rows[pd.to_numeric(supp_rows['报价'], errors='coerce') > 0]
                if len(supp_rows) > 0:
                    # 选择最低价
                    min_price_row = supp_rows.loc[pd.to_numeric(supp_rows['报价'], errors='coerce').idxmin()]
                    price = pd.to_numeric(min_price_row['报价'], errors='coerce')
                    # 货币转换
                    if min_price_row.get('币别') == 'USD':
                        price *= 7.31
                    elif min_price_row.get('币别') == 'HKD':
                        price *= 0.936
                    elif min_price_row.get('币别') == 'EUR':
                        price *= 8.12
                    price_source = '供应商价'
                    price_source_stats['供应商价'] += 1
        
        if price == 0:
            price_source_stats['无价格'] += 1
        
        # 计算缺料金额
        shortage_amount = shortage_qty * price
        
        enhanced_records.append({
            '生产订单': pso,
            '物料編號': material,
            '物料名称': row.get('物料名称', ''),
            '欠数': shortage_qty,
            '单价': price,
            '价格来源': price_source,
            '缺料金额': shortage_amount,
            '供应商': row.get('供应商', ''),
            '客户交期': row.get('客户交期', '')
        })
    
    df_enhanced = pd.DataFrame(enhanced_records)
    
    print(f"   价格来源统计:")
    for source, count in price_source_stats.items():
        print(f"      {source}: {count}条")
    
    # 5. 汇总到订单级别
    print("\n📈 汇总订单级别数据...")
    
    order_summary = []
    for pso in df_pmc['生产订单'].unique():
        pso_data = df_enhanced[df_enhanced['生产订单'] == pso]
        pmc_info = df_pmc[df_pmc['生产订单'] == pso].iloc[0]
        
        if len(pso_data) > 0:
            # 统计使用补充成本价的物料
            supplemented_items = pso_data[pso_data['价格来源'] == '补充成本价']
            
            order_summary.append({
                '产线': pmc_info['产线'],
                '生产订单': pso,
                '对应PR订单': pmc_info.get('对应PR订单', ''),
                '缺料种类数': len(pso_data),
                '原缺料金额': pso_data[pso_data['价格来源'] != '补充成本价']['缺料金额'].sum(),
                '补充价格物料数': len(supplemented_items),
                '补充金额': supplemented_items['缺料金额'].sum(),
                '新缺料金额': pso_data['缺料金额'].sum(),
                '金额变化': pso_data['缺料金额'].sum() - pso_data[pso_data['价格来源'] != '补充成本价']['缺料金额'].sum()
            })
        else:
            order_summary.append({
                '产线': pmc_info['产线'],
                '生产订单': pso,
                '对应PR订单': pmc_info.get('对应PR订单', ''),
                '缺料种类数': 0,
                '原缺料金额': 0,
                '补充价格物料数': 0,
                '补充金额': 0,
                '新缺料金额': 0,
                '金额变化': 0
            })
    
    df_order_summary = pd.DataFrame(order_summary)
    
    # 添加订单序号以保持顺序
    df_pmc_with_seq = df_pmc.copy()
    df_pmc_with_seq['订单序号'] = range(1, len(df_pmc_with_seq) + 1)
    
    df_order_summary = df_order_summary.merge(df_pmc_with_seq[['生产订单', '订单序号']], on='生产订单', how='left')
    df_order_summary = df_order_summary.sort_values('订单序号')
    
    # 6. 生成报告
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'PMC订单分析_整合成本价_{timestamp}.xlsx'
    
    print(f"\n💾 生成报告: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 订单汇总表
        df_order_output = df_order_summary.drop('订单序号', axis=1).copy()
        df_order_output['原缺料金额'] = df_order_output['原缺料金额'].apply(lambda x: f'¥{x:,.2f}')
        df_order_output['补充金额'] = df_order_output['补充金额'].apply(lambda x: f'¥{x:,.2f}')
        df_order_output['新缺料金额'] = df_order_output['新缺料金额'].apply(lambda x: f'¥{x:,.2f}')
        df_order_output['金额变化'] = df_order_output['金额变化'].apply(lambda x: f'+¥{x:,.2f}' if x > 0 else f'¥{x:,.2f}')
        
        df_order_output.to_excel(writer, sheet_name='订单汇总(整合成本价)', index=False)
        
        # 缺料明细表
        df_detail = df_enhanced.copy()
        df_detail['单价'] = df_detail['单价'].apply(lambda x: f'{x:.4f}' if x > 0 else '-')
        df_detail['缺料金额'] = df_detail['缺料金额'].apply(lambda x: f'¥{x:,.2f}')
        df_detail.to_excel(writer, sheet_name='缺料明细(含补充价格)', index=False)
        
        # 补充成本价使用情况
        df_supplement_used = df_enhanced[df_enhanced['价格来源'] == '补充成本价'].copy()
        if len(df_supplement_used) > 0:
            df_supplement_used['单价'] = df_supplement_used['单价'].apply(lambda x: f'{x:.4f}')
            df_supplement_used['缺料金额'] = df_supplement_used['缺料金额'].apply(lambda x: f'¥{x:,.2f}')
            df_supplement_used.to_excel(writer, sheet_name='使用补充成本价明细', index=False)
        
        # 统计汇总
        total_original = df_enhanced[df_enhanced['价格来源'] != '补充成本价']['缺料金额'].sum()
        total_supplement = df_enhanced[df_enhanced['价格来源'] == '补充成本价']['缺料金额'].sum()
        total_new = df_enhanced['缺料金额'].sum()
        
        summary_data = {
            '统计项': [
                'PMC订单总数',
                '有缺料订单数',
                '使用补充成本价订单数',
                '补充成本价物料条目',
                '原缺料金额合计',
                '补充金额合计',
                '新缺料金额合计',
                '金额增加',
                '价格完整率提升'
            ],
            '数值': [
                len(df_pmc),
                (df_order_summary['缺料种类数'] > 0).sum(),
                (df_order_summary['补充价格物料数'] > 0).sum(),
                len(df_supplement_used),
                f'¥{total_original:,.2f}',
                f'¥{total_supplement:,.2f}',
                f'¥{total_new:,.2f}',
                f'¥{total_new - total_original:,.2f}',
                f'{price_source_stats["补充成本价"] / sum(price_source_stats.values()) * 100:.1f}%'
            ]
        }
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='统计汇总', index=False)
    
    print(f"   ✅ 报告生成完成")
    
    # 打印统计结果
    print("\n" + "=" * 80)
    print("📊 整合结果统计")
    print("=" * 80)
    print(f"   使用补充成本价订单: {(df_order_summary['补充价格物料数'] > 0).sum()}个")
    print(f"   补充价格物料总数: {len(df_supplement_used)}条")
    print(f"   补充缺料金额: ¥{total_supplement:,.2f}")
    print(f"   总缺料金额变化: ¥{total_original:,.2f} → ¥{total_new:,.2f}")
    print(f"   增加金额: ¥{total_new - total_original:,.2f}")
    
    return output_file

if __name__ == "__main__":
    output_file = integrate_cost_prices()
    print(f"\n✅ 完成！输出文件: {output_file}")
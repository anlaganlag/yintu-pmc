#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""修正缺料计算逻辑，使用仓存不足字段重新生成报告"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def correct_shortage_calculation():
    """修正缺料计算逻辑，重新生成完整报告"""
    print("=" * 80)
    print("🔧 修正缺料计算逻辑，使用【仓存不足】字段")
    print("=" * 80)
    
    # 汇率设定
    USD_TO_RMB = 7.3
    
    # 1. 读取PMC订单
    print("\n📋 加载PMC订单...")
    df_pmc = pd.read_excel('PMC_order.xlsx')
    df_pmc['订单序号'] = range(1, len(df_pmc) + 1)
    print(f"   ✅ PMC订单: {len(df_pmc)}个")
    
    # 2. 读取和处理订单金额（美元→人民币）
    print("\n💱 处理订单金额（美元→人民币）...")
    dfs_orders = []
    
    # 国内订单
    try:
        df_89 = pd.read_excel('input/order-amt-89.xlsx', sheet_name=None)
        for sheet, df in df_89.items():
            df['数据来源'] = f'国内-{sheet}'
            if '订单金额' in df.columns:
                df['订单金额(USD)'] = df['订单金额']
                df['订单金额(RMB)'] = df['订单金额'] * USD_TO_RMB
            dfs_orders.append(df)
    except Exception as e:
        print(f"   ⚠️ 读取国内订单失败: {e}")
    
    # 柬埔寨订单
    try:
        df_89c = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name=None)
        for sheet, df in df_89c.items():
            df['数据来源'] = f'柬埔寨-{sheet}'
            if '订单金额' in df.columns:
                df['订单金额(USD)'] = df['订单金额']
                df['订单金额(RMB)'] = df['订单金额'] * USD_TO_RMB
            dfs_orders.append(df)
    except Exception as e:
        print(f"   ⚠️ 读取柬埔寨订单失败: {e}")
    
    # 处理订单数据
    if dfs_orders:
        df_all_orders = pd.concat(dfs_orders, ignore_index=True)
        df_all_orders.columns = df_all_orders.columns.str.strip()
        df_all_orders = df_all_orders.rename(columns={
            '生 產 單 号(  廠方 )': '生产订单',
            '生 產 單 号(客方 )': '客户订单号',
            '型 號( 廠方/客方 )': '产品',
            '數 量  (Pcs)': '数量',
            '客期': '客户交期_订单'
        })
        
        # 汇总订单金额
        order_amounts = df_all_orders.groupby('生产订单').agg({
            '客户订单号': lambda x: '; '.join(x.dropna().astype(str).unique()) if x.dropna().any() else '',
            '产品': lambda x: '; '.join(x.dropna().astype(str).unique()) if x.dropna().any() else '',
            '订单金额(USD)': 'sum',
            '订单金额(RMB)': 'sum',
            '数量': 'sum',
            '数据来源': lambda x: '; '.join(x.unique()),
            '客户交期_订单': lambda x: x.dropna().iloc[0] if len(x.dropna()) > 0 else ''
        }).reset_index()
        
        print(f"   ✅ 订单金额转换: ${order_amounts['订单金额(USD)'].sum():,.2f} → ¥{order_amounts['订单金额(RMB)'].sum():,.2f}")
    else:
        order_amounts = pd.DataFrame()
    
    # 3. 读取缺料数据（修正：使用仓存不足）
    print("\n📦 读取缺料数据（使用仓存不足字段）...")
    df_shortage = pd.read_excel('input/mat_owe_pso.xlsx', skiprows=1)
    df_shortage.columns = df_shortage.columns.str.strip()
    df_shortage = df_shortage.rename(columns={
        '订单编号': '生产订单',
        '物料编号': '物料編號',
        '物料名称': '物料名称',
        '仓存不足': '欠数',  # 关键修正：使用仓存不足作为欠数
        'OTS期': '客户交期',
        '供应商名称': '供应商'
    })
    
    # 确保欠数为数值
    df_shortage['欠数'] = pd.to_numeric(df_shortage['欠数'], errors='coerce').fillna(0)
    df_shortage = df_shortage[df_shortage['欠数'] > 0]  # 只保留实际缺料的记录
    
    print(f"   ✅ 缺料记录: {len(df_shortage)}条（欠数>0）")
    
    # 4. 读取价格数据
    print("\n💰 读取价格数据...")
    
    # 库存价格
    df_inventory = pd.read_excel('input/inventory_list.xlsx', sheet_name='银图库存总表')
    df_inventory.columns = df_inventory.columns.str.strip()
    df_inventory = df_inventory.rename(columns={'物項編號': '物料編號'})
    
    # 处理库存价格
    for col in ['最新報價', '成本單價']:
        if col in df_inventory.columns:
            df_inventory[col] = pd.to_numeric(df_inventory[col], errors='coerce')
    
    df_inventory['RMB单价'] = df_inventory.apply(
        lambda x: x['成本單價'] if pd.isna(x.get('最新報價', 0)) or x.get('最新報價', 0) == 0 
        else x.get('最新報價', 0), axis=1
    )
    
    # 供应商价格
    df_supplier = pd.read_excel('input/supplier.xlsx')
    df_supplier.columns = df_supplier.columns.str.strip()
    df_supplier = df_supplier.rename(columns={
        '物项编号': '物料編號',
        '单价': '报价',
        '币种': '币别'
    })
    df_supplier['报价'] = pd.to_numeric(df_supplier['报价'], errors='coerce')
    
    # 补充成本价
    df_cost_supplement = pd.read_excel('qw訂單有欠料無成本價250901v3.xlsx')
    cost_supplement_dict = {}
    for _, row in df_cost_supplement.iterrows():
        if row['Item_price'] > 0:
            key = (row['production_order_no'], row['mat'])
            cost_supplement_dict[key] = row['Item_price']
    
    print(f"   ✅ 库存价格: {len(df_inventory)}条")
    print(f"   ✅ 供应商价格: {len(df_supplier)}条") 
    print(f"   ✅ 补充成本价: {len(cost_supplement_dict)}条")
    
    # 5. 计算缺料金额（修正后）
    print("\n🔢 重新计算缺料金额...")
    
    shortage_details = []
    price_source_stats = {'补充成本价': 0, '库存价': 0, '供应商价': 0, '无价格': 0}
    
    for _, row in df_shortage.iterrows():
        pso = row['生产订单']
        material = row['物料編號']
        shortage_qty = row['欠数']
        
        if shortage_qty <= 0:
            continue
        
        # 价格查找（优先级顺序）
        price = 0
        price_source = '无价格'
        
        # 1. 补充成本价
        if (pso, material) in cost_supplement_dict:
            price = cost_supplement_dict[(pso, material)]
            price_source = '补充成本价'
            price_source_stats['补充成本价'] += 1
            
        # 2. 库存价格
        elif material in df_inventory['物料編號'].values:
            inv_row = df_inventory[df_inventory['物料編號'] == material].iloc[0]
            if pd.notna(inv_row.get('RMB单价')) and inv_row.get('RMB单价') > 0:
                price = inv_row['RMB单价']
                price_source = '库存价'
                price_source_stats['库存价'] += 1
                
        # 3. 供应商价格
        if price == 0 and material in df_supplier['物料編號'].values:
            supp_rows = df_supplier[df_supplier['物料編號'] == material]
            supp_rows = supp_rows[supp_rows['报价'] > 0]
            if len(supp_rows) > 0:
                min_price_row = supp_rows.loc[supp_rows['报价'].idxmin()]
                price = min_price_row['报价']
                # 汇率转换
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
        
        shortage_details.append({
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
    
    df_shortage_details = pd.DataFrame(shortage_details)
    
    print(f"   价格来源统计:")
    for source, count in price_source_stats.items():
        print(f"      {source}: {count}条")
    
    # 6. 汇总到订单级别
    print("\n📊 汇总订单级别数据...")
    
    # 缺料汇总
    shortage_summary = df_shortage_details.groupby('生产订单').agg({
        '物料編號': lambda x: '; '.join(x.unique()[:10]),
        '物料名称': lambda x: '; '.join(x.dropna().unique()[:10]),
        '供应商': lambda x: '; '.join(x.dropna().unique()[:5]) if x.dropna().any() else '',
        '客户交期': lambda x: x.dropna().iloc[0] if len(x.dropna()) > 0 else '',
        '欠数': 'count',  # 缺料种类数
        '缺料金额': 'sum'
    }).reset_index()
    
    shortage_summary.columns = ['生产订单', '欠料物料编号汇总', '欠料物料名称汇总', 
                               '供应商汇总', '客户交期_缺料', '缺料种类数', '缺料金额(RMB)']
    
    # 补充价格统计
    supplement_summary = df_shortage_details[df_shortage_details['价格来源'] == '补充成本价'].groupby('生产订单').agg({
        '欠数': 'count',
        '缺料金额': 'sum'
    }).reset_index()
    supplement_summary.columns = ['生产订单', '补充价格物料数', '补充价格金额']
    
    # 7. 构建主表
    print("\n📋 构建主表数据...")
    
    main_data = []
    for _, pmc_row in df_pmc.iterrows():
        pso = pmc_row['生产订单']
        
        # 基本信息
        row_data = {
            '订单序号': pmc_row['订单序号'],
            '产线': pmc_row['产线'],
            '生产订单': pso,
            '对应PR订单': pmc_row.get('对应PR订单', '')
        }
        
        # 订单金额信息
        order_info = order_amounts[order_amounts['生产订单'] == pso] if not order_amounts.empty else pd.DataFrame()
        if not order_info.empty:
            row_data['客户交期'] = order_info.iloc[0].get('客户交期_订单', '')
            row_data['客户订单号'] = order_info.iloc[0].get('客户订单号', '')
            row_data['产品'] = order_info.iloc[0].get('产品', '')
            row_data['订单金额(USD)'] = order_info.iloc[0].get('订单金额(USD)', 0)
            row_data['订单金额(RMB)'] = order_info.iloc[0].get('订单金额(RMB)', 0)
            row_data['数据来源'] = order_info.iloc[0].get('数据来源', '')
        else:
            row_data.update({
                '客户交期': '',
                '客户订单号': '',
                '产品': '',
                '订单金额(USD)': 0,
                '订单金额(RMB)': 0,
                '数据来源': ''
            })
        
        # 缺料信息
        shortage_info = shortage_summary[shortage_summary['生产订单'] == pso]
        if not shortage_info.empty:
            shortage_row = shortage_info.iloc[0]
            row_data['欠料物料编号汇总'] = shortage_row['欠料物料编号汇总']
            row_data['欠料物料名称汇总'] = shortage_row['欠料物料名称汇总']
            row_data['供应商汇总'] = shortage_row['供应商汇总']
            if not row_data['客户交期']:
                row_data['客户交期'] = shortage_row['客户交期_缺料']
            row_data['缺料种类数'] = shortage_row['缺料种类数']
            row_data['缺料金额(RMB)'] = shortage_row['缺料金额(RMB)']
        else:
            row_data.update({
                '欠料物料编号汇总': '无欠料',
                '欠料物料名称汇总': '无欠料',
                '供应商汇总': '无欠料',
                '缺料种类数': 0,
                '缺料金额(RMB)': 0
            })
        
        # 补充价格信息
        supplement_info = supplement_summary[supplement_summary['生产订单'] == pso]
        if not supplement_info.empty:
            supplement_row = supplement_info.iloc[0]
            row_data['补充价格物料数'] = supplement_row['补充价格物料数']
            row_data['补充价格金额'] = supplement_row['补充价格金额']
        else:
            row_data['补充价格物料数'] = 0
            row_data['补充价格金额'] = 0
        
        # 缺料状态和ROI
        if row_data['缺料种类数'] == 0:
            row_data['缺料状态'] = '不缺料'
            row_data['ROI显示'] = '无需投入'
            row_data['投资回报率'] = 999999
        else:
            row_data['缺料状态'] = f"缺{row_data['缺料种类数']}种物料"
            if row_data['缺料金额(RMB)'] > 0:
                roi = row_data['订单金额(RMB)'] / row_data['缺料金额(RMB)']
                row_data['投资回报率'] = roi
                row_data['ROI显示'] = f'{roi:.2f}'
            else:
                row_data['投资回报率'] = 999999
                row_data['ROI显示'] = '无需投入'
        
        main_data.append(row_data)
    
    df_main = pd.DataFrame(main_data)
    
    # 8. 生成各种分析表
    print("\n📈 生成分析表...")
    
    # ROI排序
    df_roi = df_main[df_main['缺料金额(RMB)'] > 0].copy()
    df_roi = df_roi.sort_values('投资回报率', ascending=False)
    
    # 缺料金额排序
    df_shortage_top = df_main[df_main['缺料金额(RMB)'] > 0].copy()
    df_shortage_top = df_shortage_top.sort_values('缺料金额(RMB)', ascending=False).head(30)
    
    # 产线汇总
    df_line_summary = df_main.groupby('产线').agg({
        '生产订单': 'count',
        '订单金额(RMB)': 'sum',
        '缺料金额(RMB)': 'sum',
        '缺料种类数': 'sum',
        '补充价格物料数': 'sum',
        '补充价格金额': 'sum'
    }).reset_index()
    df_line_summary.columns = ['产线', '订单数', '总订单金额', '总缺料金额', '总缺料种类', '补充价格物料总数', '补充价格金额总计']
    df_line_summary['平均ROI'] = df_line_summary['总订单金额'] / df_line_summary['总缺料金额'].replace(0, 1)
    
    # 9. 生成修正后的报告
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'PMC订单分析_修正版_{timestamp}.xlsx'
    
    print(f"\n💾 生成修正报告: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 1. 主表
        df_main_output = df_main.copy()
        
        # 格式化显示
        df_main_output['订单金额(USD)'] = df_main_output['订单金额(USD)'].apply(
            lambda x: f'${x:,.2f}' if x > 0 else '-'
        )
        df_main_output['订单金额(RMB)'] = df_main_output['订单金额(RMB)'].apply(
            lambda x: f'¥{x:,.2f}' if x > 0 else '-'
        )
        df_main_output['缺料金额(RMB)'] = df_main_output['缺料金额(RMB)'].apply(
            lambda x: f'¥{x:,.2f}' if x > 0 else '-'
        )
        df_main_output['补充价格金额'] = df_main_output['补充价格金额'].apply(
            lambda x: f'¥{x:,.2f}' if x > 0 else '-'
        )
        
        final_columns = [
            '订单序号', '产线', '生产订单', '对应PR订单', '客户交期',
            '欠料物料编号汇总', '欠料物料名称汇总', '供应商汇总',
            '客户订单号', '产品', '订单金额(USD)', '订单金额(RMB)', 
            '缺料状态', '缺料种类数', '缺料金额(RMB)',
            '补充价格物料数', '补充价格金额',
            'ROI显示', '数据来源'
        ]
        
        df_main_final = df_main_output[final_columns]
        df_main_final.to_excel(writer, sheet_name='订单分析(修正版)', index=False)
        
        # 2. ROI排序
        df_roi_output = df_roi[['产线', '生产订单', '订单金额(RMB)', '缺料金额(RMB)', '投资回报率', '缺料种类数']].copy()
        df_roi_output['订单金额(RMB)'] = df_roi_output['订单金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_roi_output['缺料金额(RMB)'] = df_roi_output['缺料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_roi_output['投资回报率'] = df_roi_output['投资回报率'].apply(lambda x: f'{x:.2f}倍')
        df_roi_output.to_excel(writer, sheet_name='ROI排序(高到低)', index=False)
        
        # 3. 缺料金额Top30
        df_shortage_output = df_shortage_top[['产线', '生产订单', '缺料金额(RMB)', '缺料种类数', '订单金额(RMB)', 'ROI显示']].copy()
        df_shortage_output['缺料金额(RMB)'] = df_shortage_output['缺料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_shortage_output['订单金额(RMB)'] = df_shortage_output['订单金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_shortage_output.to_excel(writer, sheet_name='缺料金额Top30', index=False)
        
        # 4. 产线汇总
        df_line_output = df_line_summary.copy()
        df_line_output['总订单金额'] = df_line_output['总订单金额'].apply(lambda x: f'¥{x:,.2f}')
        df_line_output['总缺料金额'] = df_line_output['总缺料金额'].apply(lambda x: f'¥{x:,.2f}')
        df_line_output['补充价格金额总计'] = df_line_output['补充价格金额总计'].apply(lambda x: f'¥{x:,.2f}')
        df_line_output['平均ROI'] = df_line_output['平均ROI'].apply(
            lambda x: f'{x:.2f}倍' if x < 999999 else '无需投入'
        )
        df_line_output.to_excel(writer, sheet_name='产线汇总', index=False)
        
        # 5. 欠料物料明细
        df_detail_output = df_shortage_details.copy()
        df_detail_output['单价'] = df_detail_output['单价'].apply(lambda x: f'{x:.4f}' if x > 0 else '-')
        df_detail_output['缺料金额'] = df_detail_output['缺料金额'].apply(lambda x: f'¥{x:,.2f}')
        df_detail_output.to_excel(writer, sheet_name='欠料物料明细', index=False)
        
        # 6. 修正说明和统计
        total_orders = len(df_main)
        shortage_orders = (df_main['缺料种类数'] > 0).sum()
        no_shortage_orders = (df_main['缺料种类数'] == 0).sum()
        supplement_orders = (df_main['补充价格物料数'] > 0).sum()
        
        total_order_amount_rmb = df_main['订单金额(RMB)'].sum()
        total_shortage_amount = df_main['缺料金额(RMB)'].sum()
        total_supplement_amount = df_main['补充价格金额'].sum()
        
        avg_roi = total_order_amount_rmb / total_shortage_amount if total_shortage_amount > 0 else 999999
        
        # 验证特定订单
        pso1_verify = df_main[df_main['生产订单'] == 'PSO2501032']['缺料金额(RMB)'].iloc[0] if len(df_main[df_main['生产订单'] == 'PSO2501032']) > 0 else 0
        pso2_verify = df_main[df_main['生产订单'] == 'PSO2501213']['缺料金额(RMB)'].iloc[0] if len(df_main[df_main['生产订单'] == 'PSO2501213']) > 0 else 0
        
        correction_data = {
            '修正项目': [
                '修正内容',
                '总订单数量',
                '有缺料订单',
                '不缺料订单',
                '使用补充成本价订单',
                '订单总金额(人民币)',
                '缺料总金额(修正后)',
                '补充成本价总金额',
                '整体ROI(修正后)',
                'PSO2501032缺料金额',
                'PSO2501213缺料金额',
                '计算字段',
                '价格来源统计'
            ],
            '数值/说明': [
                '使用【仓存不足】字段替代【工单需求】',
                f'{total_orders}个',
                f'{shortage_orders}个 ({shortage_orders/total_orders*100:.1f}%)',
                f'{no_shortage_orders}个 ({no_shortage_orders/total_orders*100:.1f}%)',
                f'{supplement_orders}个 ({supplement_orders/total_orders*100:.1f}%)',
                f'¥{total_order_amount_rmb:,.2f}',
                f'¥{total_shortage_amount:,.2f}',
                f'¥{total_supplement_amount:,.2f}',
                f'{avg_roi:.2f}倍',
                f'¥{pso1_verify:,.2f}',
                f'¥{pso2_verify:,.2f}',
                '仓存不足 × 单价',
                f'补充:{price_source_stats["补充成本价"]}条, 库存:{price_source_stats["库存价"]}条, 供应商:{price_source_stats["供应商价"]}条'
            ]
        }
        
        df_correction = pd.DataFrame(correction_data)
        df_correction.to_excel(writer, sheet_name='修正说明', index=False)
    
    print(f"   ✅ 修正报告生成完成")
    
    # 打印修正结果
    print("\n" + "=" * 80)
    print("🎉 缺料计算修正完成")
    print("=" * 80)
    print(f"   📋 总订单: {total_orders}个")
    print(f"   💰 订单总额: ¥{total_order_amount_rmb:,.2f}")
    print(f"   💸 缺料总额: ¥{total_shortage_amount:,.2f} (大幅降低)")
    print(f"   📈 修正ROI: {avg_roi:.2f}倍")
    print(f"   🔍 PSO2501032: ¥{pso1_verify:,.2f}")
    print(f"   🔍 PSO2501213: ¥{pso2_verify:,.2f}")
    print(f"   ⚡ 关键修正: 使用【仓存不足】替代【工单需求】")
    
    return output_file

if __name__ == "__main__":
    output_file = correct_shortage_calculation()
    print(f"\n✅ 修正完成！输出文件: {output_file}")
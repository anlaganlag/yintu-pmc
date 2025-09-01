#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""快速生成修正后的完整报告"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def generate_corrected_report():
    """快速生成修正后的报告"""
    print("🚀 快速生成修正后的PMC报告")
    
    USD_TO_RMB = 7.3
    
    # 1. 读取基础数据
    print("📋 读取基础数据...")
    
    # PMC订单
    df_pmc = pd.read_excel('PMC_order.xlsx')
    df_pmc['订单序号'] = range(1, len(df_pmc) + 1)
    
    # 订单金额
    dfs_orders = []
    for file_name in ['input/order-amt-89.xlsx', 'input/order-amt-89-c.xlsx']:
        try:
            df_sheets = pd.read_excel(file_name, sheet_name=None)
            for sheet, df in df_sheets.items():
                df['数据来源'] = f'{file_name.split("/")[1]}-{sheet}'
                if '订单金额' in df.columns:
                    df['订单金额(USD)'] = df['订单金额']
                    df['订单金额(RMB)'] = df['订单金额'] * USD_TO_RMB
                dfs_orders.append(df)
        except:
            continue
    
    if dfs_orders:
        df_all_orders = pd.concat(dfs_orders, ignore_index=True)
        df_all_orders.columns = df_all_orders.columns.str.strip()
        df_all_orders = df_all_orders.rename(columns={
            '生 產 單 号(  廠方 )': '生产订单',
            '生 產 單 号(客方 )': '客户订单号',
            '型 號( 廠方/客方 )': '产品',
            '客期': '客户交期'
        })
        
        order_amounts = df_all_orders.groupby('生产订单').agg({
            '客户订单号': lambda x: '; '.join(x.dropna().astype(str).unique()) if x.dropna().any() else '',
            '产品': lambda x: '; '.join(x.dropna().astype(str).unique()) if x.dropna().any() else '',
            '订单金额(USD)': 'sum',
            '订单金额(RMB)': 'sum',
            '数据来源': lambda x: '; '.join(x.unique()),
            '客户交期': lambda x: x.dropna().iloc[0] if len(x.dropna()) > 0 else ''
        }).reset_index()
    else:
        order_amounts = pd.DataFrame()
    
    # 2. 缺料数据（使用仓存不足）
    print("📦 处理缺料数据...")
    df_shortage = pd.read_excel('input/mat_owe_pso.xlsx', skiprows=1)
    df_shortage.columns = df_shortage.columns.str.strip()
    df_shortage['欠数'] = pd.to_numeric(df_shortage['仓存不足'], errors='coerce').fillna(0)
    df_shortage = df_shortage[df_shortage['欠数'] > 0]
    df_shortage = df_shortage.rename(columns={
        '订单编号': '生产订单',
        '物料编号': '物料編號',
        '物料名称': '物料名称',
        'OTS期': '客户交期_缺料',
        '供应商名称': '供应商'
    })
    
    # 3. 价格数据
    print("💰 处理价格数据...")
    
    # 库存价格
    df_inventory = pd.read_excel('input/inventory_list.xlsx', sheet_name='银图库存总表')
    df_inventory.columns = df_inventory.columns.str.strip()
    df_inventory = df_inventory.rename(columns={'物項編號': '物料編號'})
    for col in ['最新報價', '成本單價']:
        if col in df_inventory.columns:
            df_inventory[col] = pd.to_numeric(df_inventory[col], errors='coerce')
    df_inventory['RMB单价'] = df_inventory.apply(
        lambda x: x['成本單價'] if pd.isna(x.get('最新報價', 0)) or x.get('最新報價', 0) == 0 
        else x.get('最新報價', 0), axis=1
    )
    
    # 4. 计算缺料金额
    print("🔢 计算缺料金额...")
    
    # 匹配库存价格
    df_shortage_with_price = df_shortage.merge(
        df_inventory[['物料編號', 'RMB单价']].drop_duplicates(),
        on='物料編號',
        how='left'
    )
    
    # 补充成本价
    try:
        df_cost = pd.read_excel('qw訂單有欠料無成本價250901v3.xlsx')
        for _, row in df_cost.iterrows():
            if row['Item_price'] > 0:
                mask = (df_shortage_with_price['生产订单'] == row['production_order_no']) & \
                       (df_shortage_with_price['物料編號'] == row['mat'])
                df_shortage_with_price.loc[mask, 'RMB单价'] = row['Item_price']
                df_shortage_with_price.loc[mask, '价格来源'] = '补充成本价'
    except:
        print("   ⚠️ 未找到补充成本价文件")
    
    # 填充价格来源
    df_shortage_with_price['价格来源'] = df_shortage_with_price.get('价格来源', '库存价')
    df_shortage_with_price.loc[df_shortage_with_price['RMB单价'].isna(), '价格来源'] = '无价格'
    
    # 计算缺料金额
    df_shortage_with_price['RMB单价'] = df_shortage_with_price['RMB单价'].fillna(0)
    df_shortage_with_price['缺料金额'] = df_shortage_with_price['欠数'] * df_shortage_with_price['RMB单价']
    
    # 5. 汇总数据
    print("📊 汇总订单数据...")
    
    # 缺料汇总
    shortage_summary = df_shortage_with_price.groupby('生产订单').agg({
        '物料編號': lambda x: '; '.join(x.unique()[:10]),
        '物料名称': lambda x: '; '.join(x.dropna().unique()[:10]) if x.dropna().any() else '',
        '供应商': lambda x: '; '.join(x.dropna().unique()[:5]) if x.dropna().any() else '',
        '客户交期_缺料': lambda x: x.dropna().iloc[0] if len(x.dropna()) > 0 else '',
        '欠数': 'count',
        '缺料金额': 'sum'
    }).reset_index()
    shortage_summary.columns = ['生产订单', '欠料物料编号汇总', '欠料物料名称汇总', 
                               '供应商汇总', '客户交期_缺料', '缺料种类数', '缺料金额(RMB)']
    
    # 补充价格统计
    supplement_summary = df_shortage_with_price[df_shortage_with_price['价格来源'] == '补充成本价'].groupby('生产订单').agg({
        '欠数': 'count',
        '缺料金额': 'sum'
    }).reset_index()
    supplement_summary.columns = ['生产订单', '补充价格物料数', '补充价格金额']
    
    # 6. 构建主表
    print("📋 构建主表...")
    
    main_data = []
    for _, pmc_row in df_pmc.iterrows():
        pso = pmc_row['生产订单']
        
        row_data = {
            '订单序号': pmc_row['订单序号'],
            '产线': pmc_row['产线'],
            '生产订单': pso,
            '对应PR订单': pmc_row.get('对应PR订单', '')
        }
        
        # 订单金额
        order_info = order_amounts[order_amounts['生产订单'] == pso] if not order_amounts.empty else pd.DataFrame()
        if not order_info.empty:
            order_row = order_info.iloc[0]
            row_data.update({
                '客户交期': order_row.get('客户交期', ''),
                '客户订单号': order_row.get('客户订单号', ''),
                '产品': order_row.get('产品', ''),
                '订单金额(USD)': order_row.get('订单金额(USD)', 0),
                '订单金额(RMB)': order_row.get('订单金额(RMB)', 0),
                '数据来源': order_row.get('数据来源', '')
            })
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
            row_data.update({
                '欠料物料编号汇总': shortage_row['欠料物料编号汇总'],
                '欠料物料名称汇总': shortage_row['欠料物料名称汇总'],
                '供应商汇总': shortage_row['供应商汇总'],
                '缺料种类数': shortage_row['缺料种类数'],
                '缺料金额(RMB)': shortage_row['缺料金额(RMB)']
            })
            if not row_data['客户交期']:
                row_data['客户交期'] = shortage_row['客户交期_缺料']
        else:
            row_data.update({
                '欠料物料编号汇总': '无欠料',
                '欠料物料名称汇总': '无欠料',
                '供应商汇总': '无欠料',
                '缺料种类数': 0,
                '缺料金额(RMB)': 0
            })
        
        # 补充价格
        supplement_info = supplement_summary[supplement_summary['生产订单'] == pso]
        if not supplement_info.empty:
            supplement_row = supplement_info.iloc[0]
            row_data.update({
                '补充价格物料数': supplement_row['补充价格物料数'],
                '补充价格金额': supplement_row['补充价格金额']
            })
        else:
            row_data.update({
                '补充价格物料数': 0,
                '补充价格金额': 0
            })
        
        # 状态和ROI
        if row_data['缺料种类数'] == 0:
            row_data.update({
                '缺料状态': '不缺料',
                'ROI显示': '无需投入',
                '投资回报率': 999999
            })
        else:
            row_data['缺料状态'] = f"缺{row_data['缺料种类数']}种物料"
            if row_data['缺料金额(RMB)'] > 0:
                roi = row_data['订单金额(RMB)'] / row_data['缺料金额(RMB)']
                row_data.update({
                    'ROI显示': f'{roi:.2f}',
                    '投资回报率': roi
                })
            else:
                row_data.update({
                    'ROI显示': '无需投入',
                    '投资回报率': 999999
                })
        
        main_data.append(row_data)
    
    df_main = pd.DataFrame(main_data)
    
    # 7. 生成报告
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'PMC订单分析_最终修正版_{timestamp}.xlsx'
    
    print(f"💾 生成报告: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 主表
        df_output = df_main.copy()
        
        # 格式化
        df_output['订单金额(USD)'] = df_output['订单金额(USD)'].apply(lambda x: f'${x:,.2f}' if x > 0 else '-')
        df_output['订单金额(RMB)'] = df_output['订单金额(RMB)'].apply(lambda x: f'¥{x:,.2f}' if x > 0 else '-')
        df_output['缺料金额(RMB)'] = df_output['缺料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}' if x > 0 else '-')
        df_output['补充价格金额'] = df_output['补充价格金额'].apply(lambda x: f'¥{x:,.2f}' if x > 0 else '-')
        
        # 列顺序
        columns = [
            '订单序号', '产线', '生产订单', '对应PR订单', '客户交期',
            '欠料物料编号汇总', '欠料物料名称汇总', '供应商汇总',
            '客户订单号', '产品', '订单金额(USD)', '订单金额(RMB)', 
            '缺料状态', '缺料种类数', '缺料金额(RMB)',
            '补充价格物料数', '补充价格金额',
            'ROI显示', '数据来源'
        ]
        
        df_output[columns].to_excel(writer, sheet_name='订单分析(最终修正版)', index=False)
        
        # ROI排序
        df_roi = df_main[df_main['缺料金额(RMB)'] > 0].sort_values('投资回报率', ascending=False)
        df_roi_output = df_roi[['产线', '生产订单', '订单金额(RMB)', '缺料金额(RMB)', '投资回报率', '缺料种类数']].copy()
        df_roi_output['订单金额(RMB)'] = df_roi_output['订单金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_roi_output['缺料金额(RMB)'] = df_roi_output['缺料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_roi_output['投资回报率'] = df_roi_output['投资回报率'].apply(lambda x: f'{x:.2f}倍')
        df_roi_output.to_excel(writer, sheet_name='ROI排序(高到低)', index=False)
        
        # 缺料Top30
        df_top30 = df_main[df_main['缺料金额(RMB)'] > 0].nlargest(30, '缺料金额(RMB)')
        df_top30_output = df_top30[['产线', '生产订单', '缺料金额(RMB)', '缺料种类数', '订单金额(RMB)', 'ROI显示']].copy()
        df_top30_output['缺料金额(RMB)'] = df_top30_output['缺料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_top30_output['订单金额(RMB)'] = df_top30_output['订单金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_top30_output.to_excel(writer, sheet_name='缺料金额Top30', index=False)
        
        # 统计摘要
        total_orders = len(df_main)
        shortage_orders = (df_main['缺料种类数'] > 0).sum()
        total_order_amount = df_main['订单金额(RMB)'].sum()
        total_shortage_amount = df_main['缺料金额(RMB)'].sum()
        avg_roi = total_order_amount / total_shortage_amount if total_shortage_amount > 0 else 999999
        
        # 验证结果
        pso1_amount = df_main[df_main['生产订单'] == 'PSO2501032']['缺料金额(RMB)'].iloc[0] if len(df_main[df_main['生产订单'] == 'PSO2501032']) > 0 else 0
        pso2_amount = df_main[df_main['生产订单'] == 'PSO2501213']['缺料金额(RMB)'].iloc[0] if len(df_main[df_main['生产订单'] == 'PSO2501213']) > 0 else 0
        
        summary_data = {
            '指标': [
                '总订单数',
                '有缺料订单',
                '不缺料订单',
                '订单总金额(RMB)',
                '缺料总金额(修正后)',
                '整体ROI(修正后)',
                'PSO2501032缺料金额',
                'PSO2501213缺料金额',
                '修正说明'
            ],
            '数值': [
                f'{total_orders}个',
                f'{shortage_orders}个 ({shortage_orders/total_orders*100:.1f}%)',
                f'{total_orders-shortage_orders}个 ({(total_orders-shortage_orders)/total_orders*100:.1f}%)',
                f'¥{total_order_amount:,.2f}',
                f'¥{total_shortage_amount:,.2f}',
                f'{avg_roi:.2f}倍',
                f'¥{pso1_amount:,.2f}',
                f'¥{pso2_amount:,.2f}',
                '使用【仓存不足】字段替代【工单需求】'
            ]
        }
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='修正统计摘要', index=False)
        
        # 明细数据
        df_detail = df_shortage_with_price.copy()
        df_detail['单价'] = df_detail['RMB单价'].apply(lambda x: f'{x:.4f}' if x > 0 else '-')
        df_detail['缺料金额'] = df_detail['缺料金额'].apply(lambda x: f'¥{x:,.2f}')
        df_detail = df_detail[['生产订单', '物料編號', '物料名称', '欠数', '单价', '缺料金额', '价格来源', '供应商']]
        df_detail.to_excel(writer, sheet_name='缺料明细(修正)', index=False)
    
    # 打印结果
    print(f"✅ 报告生成完成")
    print(f"📊 验证结果:")
    print(f"   PSO2501032: ¥{pso1_amount:,.2f}")
    print(f"   PSO2501213: ¥{pso2_amount:,.2f}")
    print(f"   总缺料金额: ¥{total_shortage_amount:,.2f}")
    print(f"   修正后ROI: {avg_roi:.2f}倍")
    
    return output_file

if __name__ == "__main__":
    output_file = generate_corrected_report()
    print(f"🎉 修正报告完成: {output_file}")
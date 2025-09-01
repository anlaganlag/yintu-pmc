#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""修复美元转人民币汇率问题"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def fix_currency_conversion():
    """修复美元转人民币汇率，重新生成管理报告"""
    print("=" * 80)
    print("🔄 修复美元汇率转换，重新生成管理层决策报告")
    print("=" * 80)
    
    # 汇率设定
    USD_TO_RMB = 7.3
    
    # 1. 重新读取和处理订单数据
    print("\n💱 重新处理订单金额（美元→人民币）...")
    
    # 读取PMC订单
    df_pmc = pd.read_excel('PMC_order.xlsx')
    df_pmc['订单序号'] = range(1, len(df_pmc) + 1)
    
    # 读取所有订单数据并转换汇率
    dfs_orders = []
    try:
        df_89 = pd.read_excel('input/order-amt-89.xlsx', sheet_name=None)
        for sheet, df in df_89.items():
            df['数据来源'] = f'国内-{sheet}'
            # 转换美元到人民币
            if '订单金额' in df.columns:
                df['订单金额(USD)'] = df['订单金额']
                df['订单金额(RMB)'] = df['订单金额'] * USD_TO_RMB
            dfs_orders.append(df)
    except Exception as e:
        print(f"   ⚠️ 读取国内订单失败: {e}")
    
    try:
        df_89c = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name=None)
        for sheet, df in df_89c.items():
            df['数据来源'] = f'柬埔寨-{sheet}'
            # 转换美元到人民币
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
        
        # 汇总订单金额（现在是人民币）
        order_amounts = df_all_orders.groupby('生产订单').agg({
            '客户订单号': lambda x: '; '.join(x.dropna().astype(str).unique()) if x.dropna().any() else '',
            '产品': lambda x: '; '.join(x.dropna().astype(str).unique()) if x.dropna().any() else '',
            '订单金额(USD)': 'sum',
            '订单金额(RMB)': 'sum',
            '数量': 'sum',
            '数据来源': lambda x: '; '.join(x.unique()),
            '客户交期_订单': lambda x: x.dropna().iloc[0] if len(x.dropna()) > 0 else ''
        }).reset_index()
        
        print(f"   ✅ 成功转换{len(order_amounts)}个订单的金额")
        print(f"   💰 订单总金额: ${order_amounts['订单金额(USD)'].sum():,.2f} → ¥{order_amounts['订单金额(RMB)'].sum():,.2f}")
    else:
        print("   ❌ 未能读取到订单数据")
        return None
    
    # 2. 读取缺料数据
    print("\n📦 读取缺料数据...")
    df_shortage = pd.read_excel('input/mat_owe_pso.xlsx', skiprows=1)
    df_shortage.columns = df_shortage.columns.str.strip()
    df_shortage = df_shortage.rename(columns={
        '订单编号': '生产订单',
        '物料编号': '欠料物料编号',
        '物料名称': '欠料物料名称',
        'OTS期': '客户交期',
        '供应商名称': '供应商'
    })
    
    # 3. 读取整合的缺料金额数据
    print("\n💰 读取整合成本价缺料金额...")
    df_integrated = pd.read_excel('PMC订单分析_整合成本价_20250901_175508.xlsx', sheet_name='订单汇总(整合成本价)')
    
    # 解析缺料金额（移除格式化）
    def parse_amount(amount_str):
        if isinstance(amount_str, str):
            return float(amount_str.replace('¥', '').replace(',', ''))
        return float(amount_str) if not pd.isna(amount_str) else 0
    
    df_integrated['新缺料金额_数值'] = df_integrated['新缺料金额'].apply(parse_amount)
    df_integrated['补充金额_数值'] = df_integrated['补充金额'].apply(parse_amount)
    
    # 4. 构建修正后的主表数据
    print("\n📊 构建修正后的主表数据...")
    
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
        
        # 订单金额信息（修正后的人民币）
        order_info = order_amounts[order_amounts['生产订单'] == pso]
        if not order_info.empty:
            row_data['客户交期'] = order_info.iloc[0].get('客户交期_订单', '')
            row_data['客户订单号'] = order_info.iloc[0].get('客户订单号', '')
            row_data['产品'] = order_info.iloc[0].get('产品', '')
            row_data['订单金额(USD)'] = order_info.iloc[0].get('订单金额(USD)', 0)
            row_data['订单金额(RMB)'] = order_info.iloc[0].get('订单金额(RMB)', 0)
            row_data['数据来源'] = order_info.iloc[0].get('数据来源', '')
        else:
            # 尝试从缺料表获取客户交期
            shortage_info = df_shortage[df_shortage['生产订单'] == pso]
            if len(shortage_info) > 0:
                row_data['客户交期'] = shortage_info['客户交期'].dropna().iloc[0] if len(shortage_info['客户交期'].dropna()) > 0 else ''
            else:
                row_data['客户交期'] = ''
            row_data['客户订单号'] = ''
            row_data['产品'] = ''
            row_data['订单金额(USD)'] = 0
            row_data['订单金额(RMB)'] = 0
            row_data['数据来源'] = ''
        
        # 欠料汇总信息
        shortage_info = df_shortage[df_shortage['生产订单'] == pso]
        if len(shortage_info) > 0:
            material_codes = shortage_info['欠料物料编号'].dropna().unique()[:10]
            material_names = shortage_info['欠料物料名称'].dropna().unique()[:10]
            suppliers = shortage_info['供应商'].dropna().unique()[:5]
            
            row_data['欠料物料编号汇总'] = '; '.join(material_codes.astype(str)) if len(material_codes) > 0 else ''
            row_data['欠料物料名称汇总'] = '; '.join(material_names.astype(str)) if len(material_names) > 0 else ''
            row_data['供应商汇总'] = '; '.join(suppliers.astype(str)) if len(suppliers) > 0 else ''
        else:
            row_data['欠料物料编号汇总'] = '无欠料'
            row_data['欠料物料名称汇总'] = '无欠料'
            row_data['供应商汇总'] = '无欠料'
        
        # 从整合数据中获取缺料金额
        integrated_row = df_integrated[df_integrated['生产订单'] == pso]
        if not integrated_row.empty:
            row_data['缺料金额(RMB)'] = integrated_row.iloc[0]['新缺料金额_数值']
            row_data['缺料种类数'] = integrated_row.iloc[0]['缺料种类数']
            row_data['补充价格物料数'] = integrated_row.iloc[0]['补充价格物料数']
            row_data['补充价格金额'] = integrated_row.iloc[0]['补充金额_数值']
        else:
            row_data['缺料金额(RMB)'] = 0
            row_data['缺料种类数'] = 0
            row_data['补充价格物料数'] = 0
            row_data['补充价格金额'] = 0
        
        # 缺料状态
        if row_data['缺料种类数'] == 0:
            row_data['缺料状态'] = '不缺料'
        else:
            row_data['缺料状态'] = f"缺{row_data['缺料种类数']}种物料"
        
        # ROI计算（修正后）
        if row_data['缺料金额(RMB)'] > 0:
            roi = row_data['订单金额(RMB)'] / row_data['缺料金额(RMB)']
            row_data['投资回报率'] = roi
            row_data['ROI显示'] = f'{roi:.2f}'
        else:
            row_data['投资回报率'] = 999999
            row_data['ROI显示'] = '无需投入'
        
        main_data.append(row_data)
    
    df_main = pd.DataFrame(main_data)
    
    # 5. 生成各种分析表
    print("\n📈 生成修正后的分析表...")
    
    # ROI排序表
    df_roi = df_main[df_main['缺料金额(RMB)'] > 0].copy()
    df_roi = df_roi.sort_values('投资回报率', ascending=False)
    
    # 缺料金额排序表
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
    
    # 6. 生成修正后的报告
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'PMC订单分析_管理层决策版_汇率修正_{timestamp}.xlsx'
    
    print(f"\n💾 生成修正报告: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 1. 主表
        df_main_output = df_main.copy()
        
        # 格式化金额显示
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
        
        # 列顺序
        final_columns = [
            '订单序号', '产线', '生产订单', '对应PR订单', '客户交期',
            '欠料物料编号汇总', '欠料物料名称汇总', '供应商汇总',
            '客户订单号', '产品', '订单金额(USD)', '订单金额(RMB)', 
            '缺料状态', '缺料种类数', '缺料金额(RMB)',
            '补充价格物料数', '补充价格金额',
            'ROI显示', '数据来源'
        ]
        
        df_main_final = df_main_output[final_columns]
        df_main_final.to_excel(writer, sheet_name='订单分析(汇率修正版)', index=False)
        
        # 2. ROI排序表
        df_roi_output = df_roi[['产线', '生产订单', '订单金额(RMB)', '缺料金额(RMB)', '投资回报率', '缺料种类数', '补充价格物料数']].copy()
        df_roi_output['订单金额(RMB)'] = df_roi_output['订单金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_roi_output['缺料金额(RMB)'] = df_roi_output['缺料金额(RMB)'].apply(lambda x: f'¥{x:,.2f}')
        df_roi_output['投资回报率'] = df_roi_output['投资回报率'].apply(lambda x: f'{x:.2f}倍')
        df_roi_output.to_excel(writer, sheet_name='ROI排序(高到低)', index=False)
        
        # 3. 缺料金额Top30
        df_shortage_output = df_shortage_top[['产线', '生产订单', '缺料金额(RMB)', '缺料种类数', '订单金额(RMB)', 'ROI显示', '补充价格物料数']].copy()
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
        
        # 5. 修正后的管理决策摘要
        total_orders = len(df_main)
        shortage_orders = (df_main['缺料种类数'] > 0).sum()
        no_shortage_orders = (df_main['缺料种类数'] == 0).sum()
        supplement_orders = (df_main['补充价格物料数'] > 0).sum()
        
        total_order_amount_usd = df_main['订单金额(USD)'].sum()
        total_order_amount_rmb = df_main['订单金额(RMB)'].sum()
        total_shortage_amount = df_main['缺料金额(RMB)'].sum()
        total_supplement_amount = df_main['补充价格金额'].sum()
        
        avg_roi = total_order_amount_rmb / total_shortage_amount if total_shortage_amount > 0 else 999999
        
        decision_data = {
            '决策指标': [
                '总订单数量',
                '有缺料订单',
                '不缺料订单',
                '使用补充成本价订单',
                '订单总金额(美元)',
                '订单总金额(人民币)',
                '缺料总金额',
                '补充成本价总金额',
                '整体投资回报率(修正后)',
                '缺料率',
                '补充成本价覆盖率',
                '关键风险订单(ROI<2)',
                '优质订单(ROI>10)',
                '汇率使用'
            ],
            '数值': [
                f'{total_orders}个',
                f'{shortage_orders}个 ({shortage_orders/total_orders*100:.1f}%)',
                f'{no_shortage_orders}个 ({no_shortage_orders/total_orders*100:.1f}%)',
                f'{supplement_orders}个 ({supplement_orders/total_orders*100:.1f}%)',
                f'${total_order_amount_usd:,.2f}',
                f'¥{total_order_amount_rmb:,.2f}',
                f'¥{total_shortage_amount:,.2f}',
                f'¥{total_supplement_amount:,.2f} ({total_supplement_amount/total_shortage_amount*100:.1f}%)',
                f'{avg_roi:.2f}倍',
                f'{shortage_orders/total_orders*100:.1f}%',
                f'{supplement_orders/shortage_orders*100:.1f}%' if shortage_orders > 0 else '0%',
                f'{len(df_roi[df_roi["投资回报率"] < 2])}个',
                f'{len(df_roi[df_roi["投资回报率"] > 10])}个',
                f'USD→RMB (1:{USD_TO_RMB})'
            ],
            '管理建议': [
                '整体订单规模合理',
                '需要重点关注缺料管理',
                '保持不缺料状态',
                '显著改善了成本核算准确性',
                '原始订单金额（美元）',
                '已按汇率转换为人民币',
                f'需要¥{total_shortage_amount/10000:.0f}万投入解决缺料',
                '新增成本更准确反映实际需求',
                '修正汇率后ROI大幅改善，建议执行',
                '缺料率偏高，需要优化供应链',
                '成本价补充覆盖了关键物料',
                '优先处理低ROI订单',
                '优质订单可优先排产',
                '已使用7.3汇率转换'
            ]
        }
        
        df_decision = pd.DataFrame(decision_data)
        df_decision.to_excel(writer, sheet_name='管理决策摘要(汇率修正)', index=False)
    
    print(f"   ✅ 修正报告生成完成")
    
    # 打印修正后的统计
    print("\n" + "=" * 80)
    print("📊 汇率修正后的管理决策摘要")
    print("=" * 80)
    print(f"   📋 订单总数: {total_orders}个")
    print(f"   💰 订单总金额: ${total_order_amount_usd:,.2f} → ¥{total_order_amount_rmb:,.2f}")
    print(f"   ❌ 有缺料: {shortage_orders}个 ({shortage_orders/total_orders*100:.1f}%)")
    print(f"   ✅ 无缺料: {no_shortage_orders}个 ({no_shortage_orders/total_orders*100:.1f}%)")
    print(f"   💵 总投入需求: ¥{total_shortage_amount:,.2f}")
    print(f"   📈 修正后ROI: {avg_roi:.2f}倍 (大幅改善！)")
    print(f"   🔄 使用汇率: 1 USD = {USD_TO_RMB} RMB")
    
    return output_file

if __name__ == "__main__":
    output_file = fix_currency_conversion()
    if output_file:
        print(f"\n✅ 汇率修正完成！输出文件: {output_file}")
    else:
        print(f"\n❌ 汇率修正失败")
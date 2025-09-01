#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""补充PMC订单分析报告信息"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def enhance_report():
    """补充报告信息"""
    print("📋 开始补充报告信息...")
    
    # 1. 读取现有报告
    report_file = 'PMC订单分析报告_20250901_172726.xlsx'
    df_main = pd.read_excel(report_file, sheet_name='订单分析(原始顺序)')
    print(f"   ✅ 读取主表{len(df_main)}条记录")
    
    # 2. 读取PMC订单
    df_pmc = pd.read_excel('PMC_order.xlsx')
    
    # 3. 读取缺料明细数据
    df_shortage = pd.read_excel('input/mat_owe_pso.xlsx', skiprows=1)
    df_shortage.columns = df_shortage.columns.str.strip()
    
    # 重命名列
    df_shortage = df_shortage.rename(columns={
        '订单编号': '生产订单',
        '物料编号': '欠料物料编号',
        '物料名称': '欠料物料名称',
        'OTS期': '客户交期',
        '供应商名称': '供应商'
    })
    
    # 4. 读取订单交期信息
    dfs_orders = []
    try:
        # 国内订单
        df_89 = pd.read_excel('input/order-amt-89.xlsx', sheet_name=None)
        for sheet, df in df_89.items():
            dfs_orders.append(df)
    except:
        pass
    
    try:
        # 柬埔寨订单
        df_89c = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name=None)
        for sheet, df in df_89c.items():
            dfs_orders.append(df)
    except:
        pass
    
    if dfs_orders:
        df_all_orders = pd.concat(dfs_orders, ignore_index=True)
        df_all_orders.columns = df_all_orders.columns.str.strip()
        
        # 获取交期信息
        df_all_orders = df_all_orders.rename(columns={
            '生 產 單 号(  廠方 )': '生产订单',
            '客期': '客户交期_订单'
        })
        
        # 合并交期到主表
        if '客户交期_订单' in df_all_orders.columns:
            order_dates = df_all_orders.groupby('生产订单')['客户交期_订单'].first().reset_index()
            df_main = df_main.merge(order_dates, on='生产订单', how='left')
    
    # 5. 为每个订单汇总欠料物料信息
    print("   📦 汇总欠料物料信息...")
    shortage_summary = []
    
    for pso in df_main['生产订单'].unique():
        # 获取该订单的所有欠料
        pso_shortage = df_shortage[df_shortage['生产订单'] == pso].copy()
        
        if len(pso_shortage) > 0:
            # 汇总物料编号（前10个）
            material_codes = pso_shortage['欠料物料编号'].dropna().unique()[:10]
            material_codes_str = '; '.join(material_codes.astype(str))
            
            # 汇总物料名称（前10个）
            material_names = pso_shortage['欠料物料名称'].dropna().unique()[:10]
            material_names_str = '; '.join(material_names.astype(str))
            
            # 汇总供应商（去重）
            suppliers = pso_shortage['供应商'].dropna().unique()[:5]
            suppliers_str = '; '.join(suppliers.astype(str))
            
            # 获取客户交期（从缺料表）
            customer_date = pso_shortage['客户交期'].dropna().iloc[0] if '客户交期' in pso_shortage.columns and len(pso_shortage['客户交期'].dropna()) > 0 else ''
            
            shortage_summary.append({
                '生产订单': pso,
                '欠料物料编号汇总': material_codes_str if material_codes_str else '-',
                '欠料物料名称汇总': material_names_str if material_names_str else '-',
                '供应商汇总': suppliers_str if suppliers_str else '-',
                '客户交期_缺料': str(customer_date) if customer_date else '-',
                '欠料种类总数': len(material_codes)
            })
        else:
            shortage_summary.append({
                '生产订单': pso,
                '欠料物料编号汇总': '无欠料',
                '欠料物料名称汇总': '无欠料',
                '供应商汇总': '无欠料',
                '客户交期_缺料': '-',
                '欠料种类总数': 0
            })
    
    df_shortage_summary = pd.DataFrame(shortage_summary)
    
    # 6. 合并到主表
    df_enhanced = df_main.merge(df_shortage_summary, on='生产订单', how='left')
    
    # 整合客户交期（优先使用订单表的，其次使用缺料表的）
    if '客户交期_订单' in df_enhanced.columns:
        df_enhanced['客户交期'] = df_enhanced.apply(
            lambda x: x['客户交期_订单'] if pd.notna(x.get('客户交期_订单')) else x.get('客户交期_缺料', '-'),
            axis=1
        )
        # 删除临时列
        df_enhanced = df_enhanced.drop(['客户交期_订单', '客户交期_缺料'], axis=1)
    else:
        df_enhanced['客户交期'] = df_enhanced.get('客户交期_缺料', '-')
        if '客户交期_缺料' in df_enhanced.columns:
            df_enhanced = df_enhanced.drop(['客户交期_缺料'], axis=1)
    
    # 7. 重新排列列顺序
    base_cols = ['订单序号', '产线', '生产订单', '对应PR订单', '客户交期']
    material_cols = ['欠料物料编号汇总', '欠料物料名称汇总', '供应商汇总']
    other_cols = [col for col in df_enhanced.columns if col not in base_cols + material_cols]
    
    # 构建最终列顺序
    final_cols = []
    for col in base_cols:
        if col in df_enhanced.columns:
            final_cols.append(col)
    for col in material_cols:
        if col in df_enhanced.columns:
            final_cols.append(col)
    for col in other_cols:
        final_cols.append(col)
    
    df_enhanced = df_enhanced[final_cols]
    
    # 8. 创建欠料明细表（详细版本）
    print("   📝 生成欠料明细表...")
    df_detail = []
    
    for pso in df_pmc['生产订单'].unique():
        pso_shortage = df_shortage[df_shortage['生产订单'] == pso].copy()
        if len(pso_shortage) > 0:
            # 保留每个物料的详细信息
            pso_info = df_pmc[df_pmc['生产订单'] == pso].iloc[0]
            for _, row in pso_shortage.iterrows():
                df_detail.append({
                    '产线': pso_info['产线'],
                    '生产订单': pso,
                    '对应PR订单': pso_info.get('对应PR订单', ''),
                    '客户交期': row.get('客户交期', ''),
                    '欠料物料编号': row.get('欠料物料编号', ''),
                    '欠料物料名称': row.get('欠料物料名称', ''),
                    '欠数': row.get('工单需求', 0),
                    '供应商': row.get('供应商', ''),
                    '缺料金额': row.get('缺料金额', 0)
                })
    
    df_detail = pd.DataFrame(df_detail)
    
    # 9. 保存增强后的报告
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'PMC订单分析报告_增强版_{timestamp}.xlsx'
    
    print(f"\n💾 保存增强报告: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 主表（增强版）
        df_enhanced.to_excel(writer, sheet_name='订单分析(增强版)', index=False)
        
        # 欠料明细表
        if len(df_detail) > 0:
            df_detail.to_excel(writer, sheet_name='欠料物料明细', index=False)
        
        # 复制原报告的其他表
        try:
            xls = pd.ExcelFile(report_file)
            for sheet in ['ROI排序(高到低)', '缺料金额Top30', '产线汇总', '统计摘要']:
                if sheet in xls.sheet_names:
                    df_sheet = pd.read_excel(report_file, sheet_name=sheet)
                    df_sheet.to_excel(writer, sheet_name=sheet, index=False)
        except:
            pass
    
    print(f"   ✅ 报告生成完成")
    
    # 打印统计
    print("\n📊 补充信息统计:")
    print(f"   有客户交期信息: {(df_enhanced['客户交期'] != '-').sum()}个订单")
    print(f"   有欠料物料信息: {(df_enhanced['欠料物料编号汇总'] != '无欠料').sum()}个订单")
    print(f"   有供应商信息: {(df_enhanced['供应商汇总'] != '无欠料').sum()}个订单")
    if len(df_detail) > 0:
        print(f"   欠料明细记录: {len(df_detail)}条")
    
    return output_file

if __name__ == "__main__":
    output_file = enhance_report()
    print(f"\n✅ 增强报告已生成: {output_file}")
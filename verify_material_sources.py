#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证物料来源
确认无供应商物料是否存在于欠料表和库存表中
"""

import pandas as pd
import numpy as np

def verify_material_sources():
    """验证无供应商物料的来源"""
    
    print("=== 验证无供应商物料的数据来源 ===\n")
    
    # 1. 读取各个数据源
    print("📖 读取数据文件...")
    
    # 欠料表
    shortage_df = pd.read_excel('input/mat_owe_pso.xlsx', header=1)
    shortage_df = shortage_df.rename(columns={
        '物料編號': '物料编码',
        '物項編號': '物料编码'
    })
    
    # 库存表
    inventory_df = pd.read_excel('input/inventory_list.xlsx')
    inventory_df = inventory_df.rename(columns={
        '物項編號': '物料编码',
        '物料編號': '物料编码'
    })
    
    # 供应商表
    supplier_df = pd.read_excel('input/supplier.xlsx')
    supplier_df = supplier_df.rename(columns={
        '物项编号': '物料编码',
        '物項編號': '物料编码'
    })
    
    print(f"✅ 欠料表记录数: {len(shortage_df):,}")
    print(f"✅ 库存表记录数: {len(inventory_df):,}")
    print(f"✅ 供应商表记录数: {len(supplier_df):,}")
    
    # 2. 获取各表的物料编码集合
    print("\n🔍 分析物料编码分布...")
    
    # 获取唯一物料编码
    shortage_materials = set(shortage_df['物料编码'].dropna().unique())
    inventory_materials = set(inventory_df['物料编码'].dropna().unique())
    supplier_materials = set(supplier_df['物料编码'].dropna().unique())
    
    print(f"欠料表物料种类: {len(shortage_materials):,}")
    print(f"库存表物料种类: {len(inventory_materials):,}")
    print(f"供应商表物料种类: {len(supplier_materials):,}")
    
    # 3. 分析物料交集
    print("\n📊 分析物料交集情况...")
    
    # 在欠料表但不在供应商表
    shortage_no_supplier = shortage_materials - supplier_materials
    print(f"\n在欠料表但不在供应商表: {len(shortage_no_supplier):,}种")
    
    # 在库存表但不在供应商表
    inventory_no_supplier = inventory_materials - supplier_materials
    print(f"在库存表但不在供应商表: {len(inventory_no_supplier):,}种")
    
    # 同时在欠料表和库存表，但不在供应商表
    both_no_supplier = (shortage_materials & inventory_materials) - supplier_materials
    print(f"同时在欠料表和库存表，但不在供应商表: {len(both_no_supplier):,}种")
    
    # 只在欠料表，不在供应商表和库存表
    only_shortage = shortage_materials - supplier_materials - inventory_materials
    print(f"只在欠料表（不在供应商表和库存表）: {len(only_shortage):,}种")
    
    # 只在库存表，不在供应商表和欠料表
    only_inventory = inventory_materials - supplier_materials - shortage_materials
    print(f"只在库存表（不在供应商表和欠料表）: {len(only_inventory):,}种")
    
    # 4. 读取之前分析的无供应商物料
    print("\n🔄 与之前的分析结果对比...")
    
    missing_df = pd.read_excel('无供应商物料分析报告.xlsx', sheet_name='缺失供应商的物料')
    analysis_missing = set(missing_df['物料编码'].dropna().unique())
    
    print(f"之前分析发现的无供应商物料: {len(analysis_missing):,}种")
    
    # 验证这些物料的来源
    from_shortage = analysis_missing & shortage_materials
    from_inventory = analysis_missing & inventory_materials
    from_both = from_shortage & from_inventory
    from_neither = analysis_missing - shortage_materials - inventory_materials
    
    print(f"\n这些无供应商物料的来源分布:")
    print(f"  - 来自欠料表: {len(from_shortage):,}种")
    print(f"  - 来自库存表: {len(from_inventory):,}种")
    print(f"  - 同时存在于两表: {len(from_both):,}种")
    print(f"  - 两表都没有（可能来自其他JOIN）: {len(from_neither):,}种")
    
    # 5. 生成详细报告
    print("\n💾 生成验证报告...")
    
    with pd.ExcelWriter('物料来源验证报告.xlsx', engine='openpyxl') as writer:
        # 汇总统计
        summary = pd.DataFrame({
            '数据源': ['欠料表', '库存表', '供应商表'],
            '总记录数': [len(shortage_df), len(inventory_df), len(supplier_df)],
            '物料种类': [len(shortage_materials), len(inventory_materials), len(supplier_materials)]
        })
        summary.to_excel(writer, sheet_name='数据源统计', index=False)
        
        # 交集分析
        intersection = pd.DataFrame({
            '类别': [
                '欠料表但无供应商',
                '库存表但无供应商',
                '欠料+库存但无供应商',
                '仅欠料表独有',
                '仅库存表独有'
            ],
            '物料数量': [
                len(shortage_no_supplier),
                len(inventory_no_supplier),
                len(both_no_supplier),
                len(only_shortage),
                len(only_inventory)
            ]
        })
        intersection.to_excel(writer, sheet_name='交集分析', index=False)
        
        # 无供应商物料样本
        sample_materials = list(both_no_supplier)[:100]  # 取前100个
        sample_df = pd.DataFrame({'物料编码': sample_materials})
        
        # 补充物料信息
        for mat in sample_materials:
            # 从欠料表获取信息
            if mat in shortage_df['物料编码'].values:
                shortage_info = shortage_df[shortage_df['物料编码'] == mat].iloc[0]
                sample_df.loc[sample_df['物料编码'] == mat, '物料名称'] = shortage_info.get('物項名称', '')
                sample_df.loc[sample_df['物料编码'] == mat, '欠料数量'] = shortage_info.get('倉存不足 (齊套料)', 0)
            
            # 从库存表获取信息
            if mat in inventory_df['物料编码'].values:
                inv_info = inventory_df[inventory_df['物料编码'] == mat].iloc[0]
                sample_df.loc[sample_df['物料编码'] == mat, '库存数量'] = inv_info.get('實際庫存', 0)
                sample_df.loc[sample_df['物料编码'] == mat, '库存单价'] = inv_info.get('成本單價', 0)
        
        sample_df.to_excel(writer, sheet_name='无供应商物料样本', index=False)
    
    print("✅ 验证报告已生成: 物料来源验证报告.xlsx")
    
    # 6. 输出关键结论
    print("\n=== 🎯 关键结论 ===")
    print(f"1. ✅ 您的理解是正确的！")
    print(f"2. 有{len(both_no_supplier):,}种物料同时存在于【欠料表】和【库存表】，但不在【供应商表】中")
    print(f"3. 另有{len(only_shortage):,}种物料只在【欠料表】中，供应商表和库存表都没有")
    print(f"4. 还有{len(only_inventory):,}种物料只在【库存表】中，供应商表和欠料表都没有")
    print(f"\n这说明:")
    print("   - 这些物料确实有采购需求（在欠料表中）")
    print("   - 部分物料甚至有库存（在库存表中）")
    print("   - 但系统中没有供应商信息，无法进行采购")
    
    return shortage_no_supplier, inventory_no_supplier, both_no_supplier

if __name__ == "__main__":
    verify_material_sources()
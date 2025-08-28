#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析无法匹配供应商的物料
找出所有没有供应商信息的物料并分析原因
"""

import pandas as pd
import numpy as np
from datetime import datetime

def analyze_no_supplier_materials():
    """分析无法匹配供应商的物料"""
    
    print("=== 分析无法匹配供应商的物料 ===\n")
    
    # 1. 读取最新的分析报告
    report_file = "银图PMC综合物料分析报告_改进版_20250828_101505.xlsx"
    print(f"📖 读取分析报告: {report_file}")
    
    df = pd.read_excel(report_file, sheet_name='综合物料分析明细')
    print(f"总记录数: {len(df):,}")
    
    # 2. 筛选无供应商记录
    print("\n🔍 筛选无供应商的物料...")
    
    # 创建无供应商标记
    df['无供应商'] = df['供应商'].isna() | (df['供应商'] == '') | (df['供应商'] == '无供应商')
    
    no_supplier_df = df[df['无供应商']].copy()
    has_supplier_df = df[~df['无供应商']].copy()
    
    print(f"✅ 有供应商记录: {len(has_supplier_df):,} ({len(has_supplier_df)/len(df)*100:.1f}%)")
    print(f"❌ 无供应商记录: {len(no_supplier_df):,} ({len(no_supplier_df)/len(df)*100:.1f}%)")
    
    # 3. 分析无供应商物料的特征
    print("\n📊 分析无供应商物料特征...")
    
    # 按物料编码汇总
    if '物料编码' in no_supplier_df.columns:
        material_summary = no_supplier_df.groupby('物料编码').agg({
            '生产单号': 'count',
            '欠料数量': 'sum' if '欠料数量' in no_supplier_df.columns else 'count',
            '物料名称': 'first' if '物料名称' in no_supplier_df.columns else 'count'
        }).rename(columns={
            '生产单号': '出现次数',
            '欠料数量': '总欠料数量',
            '物料名称': '物料描述'
        })
        
        # 排序
        material_summary = material_summary.sort_values('出现次数', ascending=False)
        
        # 过滤掉空值和特殊值
        material_summary = material_summary[
            ~material_summary.index.isna() & 
            (material_summary.index != '已齊套')
        ]
        
        print(f"无供应商的物料种类: {len(material_summary):,}")
    
    # 4. 读取供应商数据进行对比
    print("\n🔄 读取供应商数据库进行对比...")
    supplier_file = "input/supplier.xlsx"
    supplier_df = pd.read_excel(supplier_file)
    
    # 标准化供应商数据的列名
    supplier_df = supplier_df.rename(columns={
        '物项编号': '物料编码',
        '物項編號': '物料编码',
        '物项名称': '物料名称',
        '物項名称': '物料名称',
        '供应商名称': '供应商'
    })
    
    print(f"供应商数据库记录数: {len(supplier_df):,}")
    print(f"供应商数据库物料种类: {supplier_df['物料编码'].nunique():,}")
    
    # 5. 查找缺失的物料
    print("\n🔎 查找供应商数据库中缺失的物料...")
    
    # 获取所有无供应商的物料编码
    no_supplier_materials = set(material_summary.index)
    supplier_materials = set(supplier_df['物料编码'].dropna().unique())
    
    # 找出真正缺失的物料（在供应商数据库中不存在）
    missing_materials = no_supplier_materials - supplier_materials
    print(f"供应商数据库中完全缺失的物料: {len(missing_materials):,}")
    
    # 6. 分类无供应商物料
    print("\n📝 物料分类...")
    
    categories = {
        '电子元件': [],
        '机械零件': [],
        '包装材料': [],
        '原材料': [],
        '其他': []
    }
    
    for material in missing_materials:
        if pd.notna(material) and isinstance(material, str):
            # 根据编码前缀分类
            if material.startswith('9-'):
                categories['电子元件'].append(material)
            elif any(material.startswith(prefix) for prefix in ['131-', '303-', '302-', '710-', '720-', '731-']):
                categories['机械零件'].append(material)
            elif material.startswith('8-') or material.startswith('7-'):
                categories['包装材料'].append(material)
            elif material.startswith('1-') or material.startswith('2-'):
                categories['原材料'].append(material)
            else:
                categories['其他'].append(material)
    
    for category, materials in categories.items():
        if materials:
            print(f"\n{category}: {len(materials)}个")
            for i, mat in enumerate(materials[:5], 1):
                desc = material_summary.loc[mat, '物料描述'] if mat in material_summary.index else ''
                count = material_summary.loc[mat, '出现次数'] if mat in material_summary.index else 0
                print(f"  {i}. {mat} - {desc} (出现{count}次)")
            if len(materials) > 5:
                print(f"  ... 还有{len(materials)-5}个")
    
    # 7. 生成报告
    print("\n💾 生成无供应商物料报告...")
    
    output_file = "无供应商物料分析报告.xlsx"
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet1: 汇总统计
        summary_data = {
            '指标': [
                '总记录数',
                '有供应商记录数',
                '无供应商记录数', 
                '无供应商比例',
                '无供应商物料种类',
                '供应商数据库缺失物料'
            ],
            '数值': [
                len(df),
                len(has_supplier_df),
                len(no_supplier_df),
                f"{len(no_supplier_df)/len(df)*100:.2f}%",
                len(material_summary),
                len(missing_materials)
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='汇总统计', index=False)
        
        # Sheet2: 无供应商物料明细
        material_summary.to_excel(writer, sheet_name='无供应商物料清单')
        
        # Sheet3: 缺失物料清单
        missing_df = pd.DataFrame({
            '物料编码': list(missing_materials),
            '状态': '供应商数据库缺失'
        })
        missing_df = missing_df.merge(
            material_summary[['物料描述', '出现次数']], 
            left_on='物料编码', 
            right_index=True,
            how='left'
        )
        missing_df = missing_df.sort_values('出现次数', ascending=False, na_position='last')
        missing_df.to_excel(writer, sheet_name='缺失供应商的物料', index=False)
        
        # Sheet4: 完整无供应商记录
        no_supplier_export = no_supplier_df[
            ['生产单号', '产品型号', '物料编码', '物料名称', '欠料数量'] 
            if all(col in no_supplier_df.columns for col in ['生产单号', '产品型号', '物料编码', '物料名称', '欠料数量'])
            else no_supplier_df.columns[:10]
        ].head(10000)  # 限制导出10000条
        no_supplier_export.to_excel(writer, sheet_name='无供应商记录明细', index=False)
    
    print(f"✅ 报告已生成: {output_file}")
    
    # 8. 输出关键发现
    print("\n=== 🎯 关键发现 ===")
    print(f"1. 约{len(no_supplier_df)/len(df)*100:.1f}%的记录没有匹配到供应商")
    print(f"2. 共{len(missing_materials)}种物料在供应商数据库中完全缺失")
    print(f"3. 缺失物料主要集中在: {max(categories.items(), key=lambda x: len(x[1]))[0]}")
    
    # 找出影响最大的物料（出现次数最多）
    top_missing = material_summary.head(10)
    print(f"\n=== 📈 影响最大的10个无供应商物料 ===")
    for idx, (material, row) in enumerate(top_missing.iterrows(), 1):
        if pd.notna(material) and material != '已齊套':
            print(f"{idx}. {material}: 出现{int(row['出现次数'])}次")
    
    return no_supplier_df, missing_materials, material_summary

if __name__ == "__main__":
    analyze_no_supplier_materials()
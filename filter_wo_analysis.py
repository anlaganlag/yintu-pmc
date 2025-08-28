#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析和过滤WO开头的数据
WO开头的生产单号可以过滤，重新统计无供应商物料情况
"""

import pandas as pd
import numpy as np
from datetime import datetime

def analyze_wo_data():
    """分析WO开头的数据分布情况"""
    
    print("=== 分析WO开头数据的分布情况 ===\n")
    
    # 1. 读取最新的分析报告
    report_file = "银图PMC综合物料分析报告_改进版_20250828_101505.xlsx"
    print(f"📖 读取分析报告: {report_file}")
    
    df = pd.read_excel(report_file, sheet_name='综合物料分析明细')
    print(f"总记录数: {len(df):,}")
    
    # 2. 分析生产单号分布
    print("\n🔍 分析生产单号前缀分布...")
    
    if '生产单号' in df.columns:
        # 统计不同前缀的分布
        df['生产单号'] = df['生产单号'].astype(str)
        df['前缀'] = df['生产单号'].str[:2]
        
        prefix_counts = df['前缀'].value_counts()
        print("生产单号前缀分布:")
        for prefix, count in prefix_counts.head(10).items():
            print(f"  {prefix}: {count:,}条 ({count/len(df)*100:.1f}%)")
        
        # 专门分析WO开头的数据
        wo_records = df[df['生产单号'].str.startswith('WO', na=False)]
        non_wo_records = df[~df['生产单号'].str.startswith('WO', na=False)]
        
        print(f"\n📊 WO开头数据统计:")
        print(f"  WO开头记录: {len(wo_records):,}条 ({len(wo_records)/len(df)*100:.1f}%)")
        print(f"  非WO记录: {len(non_wo_records):,}条 ({len(non_wo_records)/len(df)*100:.1f}%)")
        
        return wo_records, non_wo_records, df
    else:
        print("❌ 未找到生产单号列")
        return None, None, df

def filter_wo_and_reanalyze():
    """过滤WO开头的记录并重新分析无供应商物料"""
    
    print("\n=== 过滤WO数据并重新分析 ===\n")
    
    # 1. 获取数据
    wo_records, non_wo_records, original_df = analyze_wo_data()
    
    if non_wo_records is None:
        print("❌ 无法获取数据，分析终止")
        return
    
    # 2. 分析过滤前后的无供应商情况
    print("\n📊 对比过滤前后的无供应商情况...")
    
    # 过滤前
    original_df['无供应商'] = (
        original_df['供应商'].isna() | 
        (original_df['供应商'] == '') | 
        (original_df['供应商'] == '无供应商')
    )
    
    original_no_supplier = original_df[original_df['无供应商']]
    
    # 过滤后（去掉WO开头）
    non_wo_records['无供应商'] = (
        non_wo_records['供应商'].isna() | 
        (non_wo_records['供应商'] == '') | 
        (non_wo_records['供应商'] == '无供应商')
    )
    
    filtered_no_supplier = non_wo_records[non_wo_records['无供应商']]
    
    print("对比结果:")
    print(f"  过滤前总记录: {len(original_df):,}")
    print(f"  过滤前无供应商: {len(original_no_supplier):,} ({len(original_no_supplier)/len(original_df)*100:.1f}%)")
    print(f"  过滤后总记录: {len(non_wo_records):,}")
    print(f"  过滤后无供应商: {len(filtered_no_supplier):,} ({len(filtered_no_supplier)/len(non_wo_records)*100:.1f}%)")
    
    # 3. 分析WO记录的无供应商情况
    if len(wo_records) > 0:
        wo_records['无供应商'] = (
            wo_records['供应商'].isna() | 
            (wo_records['供应商'] == '') | 
            (wo_records['供应商'] == '无供应商')
        )
        wo_no_supplier = wo_records[wo_records['无供应商']]
        
        print(f"\n🎯 WO记录中的无供应商情况:")
        print(f"  WO总记录: {len(wo_records):,}")
        print(f"  WO无供应商: {len(wo_no_supplier):,} ({len(wo_no_supplier)/len(wo_records)*100:.1f}%)")
    
    # 4. 重新统计过滤后的无供应商物料
    print("\n📝 重新统计过滤后的无供应商物料...")
    
    if '物料编码' in filtered_no_supplier.columns:
        # 按物料编码汇总
        filtered_material_summary = filtered_no_supplier.groupby('物料编码').agg({
            '生产单号': 'count',
            '欠料数量': 'sum' if '欠料数量' in filtered_no_supplier.columns else 'count',
            '物料名称': 'first' if '物料名称' in filtered_no_supplier.columns else 'count'
        }).rename(columns={
            '生产单号': '出现次数',
            '欠料数量': '总欠料数量',
            '物料名称': '物料描述'
        })
        
        # 过滤掉空值
        filtered_material_summary = filtered_material_summary[
            ~filtered_material_summary.index.isna() & 
            (filtered_material_summary.index != '已齊套') &
            (filtered_material_summary.index != 'nan')
        ]
        
        filtered_material_summary = filtered_material_summary.sort_values('出现次数', ascending=False)
        
        print(f"过滤后无供应商物料种类: {len(filtered_material_summary):,}")
        
        # 5. 对比TOP物料变化
        print(f"\n📈 过滤后TOP 10无供应商物料:")
        top_filtered = filtered_material_summary.head(10)
        for idx, (material, row) in enumerate(top_filtered.iterrows(), 1):
            if pd.notna(material):
                print(f"  {idx}. {material}: {int(row['出现次数'])}次")
    
    # 6. 生成过滤后的报告
    print("\n💾 生成过滤后的分析报告...")
    
    output_file = "过滤WO后无供应商物料报告.xlsx"
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Sheet1: 过滤对比统计
        comparison_data = {
            '统计项': [
                '总记录数',
                '无供应商记录数',
                '无供应商比例',
                '无供应商物料种类',
                'WO记录数',
                'WO无供应商记录数'
            ],
            '过滤前': [
                len(original_df),
                len(original_no_supplier),
                f"{len(original_no_supplier)/len(original_df)*100:.2f}%",
                len(original_no_supplier['物料编码'].unique()) if '物料编码' in original_no_supplier.columns else 'N/A',
                len(wo_records) if wo_records is not None else 0,
                len(wo_no_supplier) if 'wo_no_supplier' in locals() else 'N/A'
            ],
            '过滤后（去掉WO）': [
                len(non_wo_records),
                len(filtered_no_supplier),
                f"{len(filtered_no_supplier)/len(non_wo_records)*100:.2f}%",
                len(filtered_material_summary) if 'filtered_material_summary' in locals() else 'N/A',
                0,
                0
            ]
        }
        comparison_df = pd.DataFrame(comparison_data)
        comparison_df.to_excel(writer, sheet_name='过滤前后对比', index=False)
        
        # Sheet2: 过滤后无供应商物料清单
        if 'filtered_material_summary' in locals():
            filtered_material_summary.to_excel(writer, sheet_name='过滤后无供应商物料')
        
        # Sheet3: WO记录分析
        if len(wo_records) > 0:
            wo_sample = wo_records.head(1000)[['生产单号', '产品型号', '物料编码', '物料名称', '供应商']].copy()
            wo_sample.to_excel(writer, sheet_name='WO记录样本', index=False)
        
        # Sheet4: 过滤后的详细记录样本
        filtered_sample = filtered_no_supplier.head(5000)[
            ['生产单号', '产品型号', '物料编码', '物料名称', '欠料数量']
            if all(col in filtered_no_supplier.columns for col in ['生产单号', '产品型号', '物料编码', '物料名称', '欠料数量'])
            else filtered_no_supplier.columns[:8]
        ].copy()
        filtered_sample.to_excel(writer, sheet_name='过滤后无供应商记录', index=False)
    
    print(f"✅ 过滤报告已生成: {output_file}")
    
    # 7. 记录过滤规则和影响
    print(f"\n=== 📋 过滤规则记录 ===")
    print(f"过滤规则: 去除生产单号以'WO'开头的记录")
    print(f"过滤理由: WO开头的数据可以忽略（用户指定）")
    print(f"过滤影响:")
    print(f"  - 减少记录: {len(wo_records):,}条")
    print(f"  - 减少无供应商记录: {len(wo_no_supplier) if 'wo_no_supplier' in locals() else 'N/A'}条")
    print(f"  - 无供应商比例变化: {len(original_no_supplier)/len(original_df)*100:.1f}% → {len(filtered_no_supplier)/len(non_wo_records)*100:.1f}%")
    
    # 8. 生成过滤规则文档
    filter_doc = f"""# 数据过滤规则记录

## 过滤规则
- **规则**: 过滤生产单号以'WO'开头的记录
- **实施日期**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- **实施原因**: 用户指定WO开头的数据可以过滤

## 过滤影响统计
- **过滤前总记录**: {len(original_df):,}条
- **过滤后总记录**: {len(non_wo_records):,}条
- **被过滤记录**: {len(wo_records) if wo_records is not None else 0:,}条
- **过滤比例**: {len(wo_records)/len(original_df)*100:.2f}%

## 无供应商物料影响
- **过滤前无供应商记录**: {len(original_no_supplier):,}条
- **过滤后无供应商记录**: {len(filtered_no_supplier):,}条
- **减少无供应商记录**: {len(original_no_supplier) - len(filtered_no_supplier):,}条

## 生成文件
- `过滤WO后无供应商物料报告.xlsx` - 详细分析报告
- `WO过滤规则记录.md` - 本文档

## 注意事项
- WO开头的记录已被永久过滤
- 后续分析应基于过滤后的数据
- 如需恢复WO数据，请重新运行原始分析
"""
    
    with open('WO过滤规则记录.md', 'w', encoding='utf-8') as f:
        f.write(filter_doc)
    
    print(f"✅ 过滤规则文档已生成: WO过滤规则记录.md")
    
    return filtered_no_supplier, filtered_material_summary

if __name__ == "__main__":
    filter_wo_and_reanalyze()
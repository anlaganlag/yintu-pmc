#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从订单表提取所有成品信息
提取生产单号、产品型号、BOM编号等成品相关信息
"""

import pandas as pd
import numpy as np
from pathlib import Path

def extract_finished_products():
    """提取所有成品信息"""
    
    print("=== 从订单表提取成品信息 ===\n")
    
    finished_products = []
    
    # 1. 处理国内订单
    print("📋 处理国内订单...")
    domestic_file = "input/order-amt-89.xlsx"
    
    for sheet_name in ['8月', '9月']:
        print(f"  - 处理 {sheet_name} 工作表")
        df = pd.read_excel(domestic_file, sheet_name=sheet_name)
        
        # 标准化列名
        df = df.rename(columns={
            '生 產 單 号(  廠方 )': '生产单号',
            '生 產 單 号( 廠方 )': '生产单号',  # 不同spacing
            '型 號( 廠方/客方 )': '产品型号',
            'BOM NO.': 'BOM编号',
            '數 量  (Pcs)': '数量',
            '數 量 (Pcs)': '数量',  # 不同spacing
            'Unite Price': '单价',
            '订单金额': '金额'
        })
        
        # 提取成品信息
        products = df[['生产单号', '产品型号', 'BOM编号', '数量', '单价', '金额']].copy()
        products['数据来源'] = f'国内-{sheet_name}'
        products['月份'] = sheet_name
        
        # 清理数据
        products = products.dropna(subset=['生产单号', '产品型号'])
        products['生产单号'] = products['生产单号'].astype(str)
        products['产品型号'] = products['产品型号'].astype(str)
        
        finished_products.append(products)
        print(f"    找到 {len(products)} 个成品记录")
    
    # 2. 处理柬埔寨订单
    print("\n📋 处理柬埔寨订单...")
    cambodia_file = "input/order-amt-89-c.xlsx"
    
    for sheet_name in ['8月 -柬', '9月 -柬']:
        month = sheet_name.replace(' -柬', '')
        print(f"  - 处理 {sheet_name} 工作表")
        df = pd.read_excel(cambodia_file, sheet_name=sheet_name)
        
        # 标准化列名
        df = df.rename(columns={
            '生 產 單 号(  廠方 )': '生产单号',
            '生 產 單 号( 廠方 )': '生产单号',  # 不同spacing
            '型 號( 廠方/客方 )': '产品型号',
            'BOM NO.': 'BOM编号',
            '數 量  (Pcs)': '数量',
            '數 量 (Pcs)': '数量',  # 不同spacing
            'Unite Price': '单价',
            '订单金额': '金额'
        })
        
        # 提取成品信息
        products = df[['生产单号', '产品型号', 'BOM编号', '数量', '单价', '金额']].copy()
        products['数据来源'] = f'柬埔寨-{month}'
        products['月份'] = month
        
        # 清理数据
        products = products.dropna(subset=['生产单号', '产品型号'])
        products['生产单号'] = products['生产单号'].astype(str)
        products['产品型号'] = products['产品型号'].astype(str)
        
        finished_products.append(products)
        print(f"    找到 {len(products)} 个成品记录")
    
    # 3. 合并所有数据
    print("\n🔗 合并所有成品数据...")
    all_products = pd.concat(finished_products, ignore_index=True)
    print(f"合并后总记录数: {len(all_products)}")
    
    # 4. 生成成品清单（去重）
    print("\n📝 生成成品清单...")
    
    # 按生产单号去重，保留最新记录
    product_list = all_products.drop_duplicates(subset=['生产单号'], keep='last').copy()
    print(f"去重后成品种类: {len(product_list)}")
    
    # 按产品型号统计
    product_summary = all_products.groupby('产品型号').agg({
        '生产单号': 'count',
        '数量': 'sum',
        '金额': 'sum',
        '数据来源': lambda x: ', '.join(x.unique())
    }).rename(columns={
        '生产单号': '订单数量',
        '数量': '总生产数量',
        '金额': '总订单金额'
    }).reset_index()
    
    product_summary = product_summary.sort_values('总订单金额', ascending=False)
    print(f"不同产品型号数量: {len(product_summary)}")
    
    # 5. 保存结果
    print("\n💾 保存结果...")
    
    output_file = "成品信息提取结果.xlsx"
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 详细成品列表
        all_products.to_excel(writer, sheet_name='所有成品记录', index=False)
        
        # 去重成品清单
        product_list.to_excel(writer, sheet_name='成品清单', index=False)
        
        # 产品型号汇总
        product_summary.to_excel(writer, sheet_name='产品型号汇总', index=False)
        
        # 生成成品物料编码清单（基于生产单号）
        material_codes = product_list[['生产单号', '产品型号', 'BOM编号']].copy()
        material_codes['物料编码'] = material_codes['生产单号']  # 生产单号作为成品物料编码
        material_codes['物料名称'] = material_codes['产品型号']  # 产品型号作为物料名称
        material_codes = material_codes[['物料编码', '物料名称', 'BOM编号', '生产单号', '产品型号']]
        material_codes.to_excel(writer, sheet_name='成品物料编码清单', index=False)
    
    print(f"✅ 结果已保存到: {output_file}")
    
    # 6. 显示汇总信息
    print(f"\n=== 📈 提取汇总 ===")
    print(f"总订单记录数: {len(all_products):,}")
    print(f"不重复成品数量: {len(product_list):,}")
    print(f"不同产品型号: {len(product_summary):,}")
    print(f"总生产数量: {all_products['数量'].sum():,.0f} pcs")
    print(f"总订单金额: ¥{all_products['金额'].sum():,.2f}")
    
    print(f"\n=== 🏆 Top 10 产品型号（按订单金额） ===")
    print(product_summary.head(10)[['产品型号', '订单数量', '总生产数量', '总订单金额']])
    
    return all_products, product_list, product_summary

if __name__ == "__main__":
    extract_finished_products()
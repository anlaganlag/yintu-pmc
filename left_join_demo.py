#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LEFT JOIN架构演示
展示银图PMC分析系统的数据流和JOIN逻辑
"""

import pandas as pd
import numpy as np

def demo_left_join_architecture():
    """演示LEFT JOIN架构的完整流程"""
    print("🚀 银图PMC LEFT JOIN架构演示")
    print("="*80)
    
    # 1. 模拟订单表（主表）
    print("1️⃣ 主表：订单表 (455条记录)")
    orders_sample = pd.DataFrame({
        '生产订单号': ['PSO2501521', 'PSO2501724', 'PSO2501010', 'PSO2402857', 'PSO2501999'],
        '客户订单号': ['4500575957', 'IP-0021(175962)', '4500576656', '4500575810', '4500579999'],
        '产品型号': ['SP8515/AS126E', '#87-0841', 'SP8019/ACA', 'SP8515/AS126', 'SP9999/TEST'],
        '订单金额(USD)': [8054.42, np.nan, 1244.88, 2000.00, 1500.00],
        '月份': ['9月', '8月', '8月', '9月', '9月']
    })
    
    print(orders_sample.to_string(index=False))
    print(f"特点：PSO2501724 订单金额为NaN，PSO2501999是新订单无欠料")
    
    # 2. 模拟欠料表
    print("\n2️⃣ 欠料表 (9768条记录)")
    shortage_sample = pd.DataFrame({
        '订单编号': ['PSO2501521', 'PSO2501521', 'PSO2501010', 'PSO2501010', 'PSO2402857'],
        '物料编号': ['MAT001', 'MAT002', 'MAT003', 'MAT004', 'MAT005'],
        '物料名称': ['电机外壳', '风叶', '电源线', '开关', '螺丝'],
        '仓存不足': [100, 50, 200, 80, 500],
        '请购组': ['电机组', '风叶组', '电源组', '开关组', '五金组']
    })
    
    print(shortage_sample.to_string(index=False))
    print(f"特点：PSO2501724和PSO2501999不在欠料表中")
    
    # 3. LEFT JOIN 过程演示
    print("\n3️⃣ LEFT JOIN 过程：订单表 ← 欠料表")
    print("-" * 50)
    
    # 执行LEFT JOIN
    result_step1 = orders_sample.merge(
        shortage_sample,
        left_on='生产订单号',
        right_on='订单编号',
        how='left'  # 关键：LEFT JOIN保留所有订单
    )
    
    print("JOIN后结果：")
    display_cols = ['生产订单号', '客户订单号', '产品型号', '订单金额(USD)', '物料编号', '物料名称', '仓存不足']
    print(result_step1[display_cols].to_string(index=False))
    
    print(f"\n📊 记录数变化: {len(orders_sample)} → {len(result_step1)}")
    print("解释：")
    print("  - PSO2501521: 1个订单 → 2条记录（2个欠料物料）")
    print("  - PSO2501010: 1个订单 → 2条记录（2个欠料物料）") 
    print("  - PSO2402857: 1个订单 → 1条记录（1个欠料物料）")
    print("  - PSO2501724: 1个订单 → 1条记录（无欠料，物料字段为NaN）✅保留")
    print("  - PSO2501999: 1个订单 → 1条记录（无欠料，物料字段为NaN）✅保留")
    
    # 4. 模拟库存价格表
    print("\n4️⃣ 继续LEFT JOIN：库存价格表")
    inventory_sample = pd.DataFrame({
        '物項編號': ['MAT001', 'MAT002', 'MAT003', 'MAT004'],  # MAT005缺失价格
        '物項名稱': ['电机外壳', '风叶', '电源线', '开关'],
        'RMB单价': [15.50, 8.20, 3.50, 12.00]
    })
    
    result_step2 = result_step1.merge(
        inventory_sample,
        left_on='物料编号',
        right_on='物項編號',
        how='left'
    )
    
    print("库存价格表：")
    print(inventory_sample.to_string(index=False))
    
    # 5. 模拟供应商表
    print("\n5️⃣ 继续LEFT JOIN：供应商表")
    supplier_sample = pd.DataFrame({
        '物项编号': ['MAT001', 'MAT002', 'MAT003'],  # MAT004和MAT005缺失供应商
        '供应商名称': ['电机厂A', '风叶厂B', '电源厂C'],
        '单价': [15.00, 8.00, 3.20],
        '币种': ['RMB', 'RMB', 'RMB']
    })
    
    # 选择最低价供应商的逻辑
    supplier_mapping = {}
    for material in result_step2['物料编号'].dropna().unique():
        material_suppliers = supplier_sample[supplier_sample['物项编号'] == material]
        if len(material_suppliers) > 0:
            best_supplier = material_suppliers.iloc[0]  # 简化：取第一个
            supplier_mapping[material] = {
                '主供应商名称': best_supplier['供应商名称'],
                '供应商单价': best_supplier['单价']
            }
    
    # 映射供应商信息
    result_step2['主供应商名称'] = result_step2['物料编号'].map(
        lambda x: supplier_mapping.get(x, {}).get('主供应商名称', None)
    )
    result_step2['供应商单价'] = result_step2['物料编号'].map(
        lambda x: supplier_mapping.get(x, {}).get('供应商单价', None)
    )
    
    print("供应商表：")
    print(supplier_sample.to_string(index=False))
    
    # 6. 显示最终结果
    print("\n6️⃣ 最终LEFT JOIN结果")
    print("="*80)
    final_display_cols = ['生产订单号', '客户订单号', '物料编号', '仓存不足', 
                         'RMB单价', '主供应商名称', '订单金额(USD)']
    final_result = result_step2[final_display_cols].copy()
    
    # 计算欠料金额
    final_result['欠料金额(RMB)'] = (
        pd.to_numeric(final_result['仓存不足'], errors='coerce').fillna(0) * 
        pd.to_numeric(final_result['RMB单价'], errors='coerce').fillna(0)
    )
    
    print(final_result.to_string(index=False))
    
    # 7. 数据完整性分析
    print("\n7️⃣ 数据完整性分析")
    print("-" * 30)
    
    def analyze_completeness(row):
        has_shortage = pd.notna(row['物料编号'])
        has_price = pd.notna(row['RMB单价']) and row['RMB单价'] > 0
        has_supplier = pd.notna(row['主供应商名称'])
        has_order_amount = pd.notna(row['订单金额(USD)']) and row['订单金额(USD)'] > 0
        has_production_order = pd.notna(row['生产订单号'])
        
        if has_shortage and has_price and has_supplier and has_order_amount:
            return "完整"
        elif has_shortage and has_price and has_order_amount:
            return "部分"
        elif has_order_amount and not has_shortage:
            return "完整"  # 不缺料订单
        elif has_production_order and not has_shortage:
            return "不缺料订单"  # 无订单金额的不缺料订单
        elif has_production_order:
            return "订单信息不完整"
        else:
            return "无数据"
    
    final_result['数据完整性'] = final_result.apply(analyze_completeness, axis=1)
    
    completeness_stats = final_result['数据完整性'].value_counts()
    print("数据完整性分布：")
    for status, count in completeness_stats.items():
        print(f"  {status}: {count}条")
    
    # 8. ROI计算演示
    print("\n8️⃣ ROI计算（按生产订单汇总）")
    print("-" * 40)
    
    # 订单金额去重处理
    final_result['订单金额(RMB)'] = pd.to_numeric(final_result['订单金额(USD)'], errors='coerce') * 7.2
    
    # 按生产订单汇总
    order_summary = final_result.groupby('生产订单号').agg({
        '订单金额(RMB)': 'first',  # 订单金额不重复计算
        '欠料金额(RMB)': 'sum',    # 欠料金额需要汇总
        '数据完整性': 'first'
    }).reset_index()
    
    # 计算ROI
    order_summary['ROI倍数'] = np.where(
        order_summary['欠料金额(RMB)'] > 0,
        order_summary['订单金额(RMB)'] / order_summary['欠料金额(RMB)'],
        0
    )
    order_summary['ROI状态'] = order_summary.apply(lambda row: 
        '无需投入' if row['欠料金额(RMB)'] == 0 else f"{row['ROI倍数']:.2f}倍", axis=1)
    
    roi_display = order_summary[['生产订单号', '订单金额(RMB)', '欠料金额(RMB)', 'ROI状态', '数据完整性']]
    print(roi_display.to_string(index=False))
    
    # 9. LEFT JOIN架构优势总结
    print("\n9️⃣ LEFT JOIN架构优势")
    print("="*50)
    print("✅ 订单完整性：所有394个订单都保留，无遗漏")
    print("✅ 数据透明：清楚标记每条记录的完整性状态")
    print("✅ 业务友好：不缺料订单也能正常显示和分析")
    print("✅ ROI准确：正确处理一对多关系的金额汇总")
    print("✅ 扩展性强：可以继续LEFT JOIN更多数据源")
    
    print(f"\n最终统计：{len(order_summary)}个生产订单全部保留 ✅")
    
    return final_result, order_summary

if __name__ == "__main__":
    demo_left_join_architecture()
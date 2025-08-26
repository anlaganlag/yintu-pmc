#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
JOIN方式对比演示：INNER JOIN vs LEFT JOIN
展示不同JOIN方式对数据完整性的影响
"""

import pandas as pd
import numpy as np

def compare_join_methods():
    """对比不同JOIN方式的结果"""
    print("⚔️ JOIN方式对比：INNER JOIN vs LEFT JOIN")
    print("="*80)
    
    # 原始数据
    orders = pd.DataFrame({
        '生产订单号': ['PSO2501521', 'PSO2501724', 'PSO2501010', 'PSO2402857', 'PSO2501999'],
        '客户订单号': ['4500575957', 'IP-0021(175962)', '4500576656', '4500575810', '4500579999'],
        '产品型号': ['SP8515/AS126E', '#87-0841', 'SP8019/ACA', 'SP8515/AS126', 'SP9999/TEST'],
        '订单金额(USD)': [8054.42, np.nan, 1244.88, 2000.00, 1500.00]
    })
    
    shortage = pd.DataFrame({
        '订单编号': ['PSO2501521', 'PSO2501521', 'PSO2501010', 'PSO2501010', 'PSO2402857'],
        '物料编号': ['MAT001', 'MAT002', 'MAT003', 'MAT004', 'MAT005'],
        '仓存不足': [100, 50, 200, 80, 500]
    })
    
    print("📊 原始数据统计:")
    print(f"  - 订单表: {len(orders)}个订单")
    print(f"  - 欠料表: {len(shortage)}条欠料记录")
    print(f"  - 有欠料的订单: {shortage['订单编号'].nunique()}个")
    print(f"  - 无欠料的订单: {len(orders) - shortage['订单编号'].nunique()}个")
    
    # 方法1: INNER JOIN（传统方法）
    print("\n🚫 方法1: INNER JOIN（只保留有匹配的记录）")
    print("-" * 50)
    
    inner_result = orders.merge(
        shortage,
        left_on='生产订单号',
        right_on='订单编号',
        how='inner'  # 内连接：只保留两表都有的记录
    )
    
    inner_orders = inner_result['生产订单号'].nunique()
    print(f"结果: {len(inner_result)}条记录，涉及{inner_orders}个订单")
    print("保留的订单:", inner_result['生产订单号'].unique().tolist())
    print("❌ 丢失的订单:", [x for x in orders['生产订单号'] if x not in inner_result['生产订单号'].values])
    
    print("\nINNER JOIN结果预览:")
    display_cols = ['生产订单号', '客户订单号', '产品型号', '物料编号', '仓存不足']
    print(inner_result[display_cols].head().to_string(index=False))
    
    # 方法2: LEFT JOIN（银图PMC方法）
    print("\n✅ 方法2: LEFT JOIN（保留所有订单）")
    print("-" * 50)
    
    left_result = orders.merge(
        shortage,
        left_on='生产订单号', 
        right_on='订单编号',
        how='left'  # 左连接：保留左表所有记录
    )
    
    left_orders = left_result['生产订单号'].nunique()
    print(f"结果: {len(left_result)}条记录，涉及{left_orders}个订单")
    print("保留的订单:", left_result['生产订单号'].unique().tolist())
    print("✅ 丢失的订单:", [x for x in orders['生产订单号'] if x not in left_result['生产订单号'].values])
    
    print("\nLEFT JOIN结果预览:")
    print(left_result[display_cols].to_string(index=False))
    
    # 业务影响分析
    print("\n💼 业务影响分析")
    print("="*60)
    
    # 分析无欠料订单的情况
    no_shortage_orders = left_result[left_result['物料编号'].isna()]
    
    print("📈 无欠料订单分析:")
    for _, order in no_shortage_orders.iterrows():
        order_amount = order['订单金额(USD)']
        status = "有订单金额" if pd.notna(order_amount) else "无订单金额"
        amount_str = f"${order_amount:,.2f}" if pd.notna(order_amount) else "N/A"
        print(f"  - {order['生产订单号']}: {order['产品型号']} | {status} ({amount_str})")
    
    # 财务影响分析
    print("\n💰 财务影响分析:")
    
    # INNER JOIN的财务统计
    inner_order_amount = 0
    for order_no in inner_result['生产订单号'].unique():
        order_data = orders[orders['生产订单号'] == order_no].iloc[0]
        if pd.notna(order_data['订单金额(USD)']):
            inner_order_amount += order_data['订单金额(USD)']
    
    # LEFT JOIN的财务统计  
    left_order_amount = 0
    for order_no in left_result['生产订单号'].unique():
        order_data = orders[orders['生产订单号'] == order_no].iloc[0]
        if pd.notna(order_data['订单金额(USD)']):
            left_order_amount += order_data['订单金额(USD)']
    
    missing_amount = left_order_amount - inner_order_amount
    missing_ratio = (missing_amount / left_order_amount * 100) if left_order_amount > 0 else 0
    
    print(f"  INNER JOIN统计金额: ${inner_order_amount:,.2f}")
    print(f"  LEFT JOIN统计金额:  ${left_order_amount:,.2f}")
    print(f"  丢失金额: ${missing_amount:,.2f} ({missing_ratio:.1f}%)")
    
    # ROI计算对比
    print("\n📊 ROI计算准确性对比:")
    
    def calculate_roi_summary(result_df, method_name):
        """计算ROI汇总"""
        # 模拟价格数据
        result_df = result_df.copy()
        result_df['RMB单价'] = result_df['物料编号'].map({
            'MAT001': 15.5, 'MAT002': 8.2, 'MAT003': 3.5, 
            'MAT004': 12.0, 'MAT005': 10.0
        })
        result_df['欠料金额(RMB)'] = (
            pd.to_numeric(result_df['仓存不足'], errors='coerce').fillna(0) * 
            pd.to_numeric(result_df['RMB单价'], errors='coerce').fillna(0)
        )
        result_df['订单金额(RMB)'] = pd.to_numeric(result_df['订单金额(USD)'], errors='coerce') * 7.2
        
        # 按订单汇总
        order_summary = result_df.groupby('生产订单号').agg({
            '订单金额(RMB)': 'first',
            '欠料金额(RMB)': 'sum'
        }).reset_index()
        
        total_investment = order_summary['欠料金额(RMB)'].sum()
        total_return = order_summary[order_summary['订单金额(RMB)'].notna()]['订单金额(RMB)'].sum()
        avg_roi = (total_return / total_investment) if total_investment > 0 else 0
        
        print(f"  {method_name}:")
        print(f"    - 分析订单数: {len(order_summary)}")
        print(f"    - 总投入: ¥{total_investment:,.2f}")
        print(f"    - 总回款: ¥{total_return:,.2f}")
        print(f"    - 平均ROI: {avg_roi:.2f}倍")
        
        return order_summary
    
    inner_roi = calculate_roi_summary(inner_result, "INNER JOIN")
    left_roi = calculate_roi_summary(left_result, "LEFT JOIN ")
    
    # 决策支持对比
    print("\n🎯 决策支持能力对比:")
    print("  INNER JOIN方式:")
    print("    ❌ 只能分析有缺料的订单")
    print("    ❌ 无法识别不缺料的盈利订单") 
    print("    ❌ 管理层看不到完整的生产计划")
    print("    ❌ ROI分析不完整，可能误导决策")
    
    print("  LEFT JOIN方式:")
    print("    ✅ 分析所有订单，无遗漏")
    print("    ✅ 清楚识别不缺料订单（立即可生产）")
    print("    ✅ 管理层获得完整的生产全貌")
    print("    ✅ ROI分析准确，支持正确决策")
    
    print("\n🏆 总结:")
    print(f"  LEFT JOIN保证了{left_orders}/{len(orders)}个订单的完整分析")
    print(f"  避免了${missing_amount:,.2f}的订单金额被忽略")
    print("  为管理层提供了完整、准确的决策数据")
    
    return inner_result, left_result

if __name__ == "__main__":
    compare_join_methods()
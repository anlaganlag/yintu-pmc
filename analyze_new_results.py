#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析使用欠料清单3.xlsx后的结果
对比用户验证的1967.31万数据
"""

import pandas as pd
import numpy as np

def analyze_new_results():
    """分析新的分析结果"""
    print("=" * 60)
    print("分析使用欠料清单3.xlsx后的结果")
    print("=" * 60)
    
    # 1. 读取新生成的不缺料清单
    print("\n=== 1. 新生成的不缺料清单分析 ===")
    
    try:
        new_list = pd.read_excel('8月9月不缺料订单清单.xlsx', sheet_name='不缺料订单清单')
        print(f"新生成不缺料清单:")
        print(f"  订单数: {len(new_list)}")
        
        # 计算总金额
        total_amount_rmb = new_list['订单金额(RMB)'].sum()
        total_amount_usd = new_list['订单金额(USD)'].sum()
        
        print(f"  总金额(RMB): {total_amount_rmb:,.2f} ({total_amount_rmb/10000:.2f}万)")
        print(f"  总金额(USD): {total_amount_usd:,.2f}")
        
    except Exception as e:
        print(f"❌ 读取新清单失败: {e}")
        return
    
    # 2. 对比用户验证的清单
    print("\n=== 2. 对比用户验证的清单 ===")
    
    try:
        user_verified = pd.read_excel('8月9月不缺料订单清单.xlsx', sheet_name='统计汇总')  
        print("从统计汇总中查看...")
        print(user_verified)
    except:
        # 如果没有统计汇总，尝试其他方式
        pass
    
    try:
        # 读取用户的原始验证清单
        original_verified = pd.read_excel('8月9月不缺料订单清单.xlsx')
        print("原始验证清单信息:")
        print(f"工作表名: {pd.ExcelFile('8月9月不缺料订单清单.xlsx').sheet_names}")
        
    except Exception as e:
        print(f"读取统计信息失败: {e}")
    
    # 3. 与脚本输出的12215.94万对比
    print("\n=== 3. 结果对比分析 ===")
    
    script_output = 12215.94  # 从脚本输出的12215.94万
    user_verified_amount = 1967.31  # 用户验证的金额
    
    print(f"脚本输出(新): {script_output:.2f}万")
    print(f"用户验证: {user_verified_amount:.2f}万")
    
    difference = script_output - user_verified_amount
    print(f"差异: {difference:.2f}万")
    
    # 4. 分析可能的原因
    print("\n=== 4. 差异分析 ===")
    
    # 检查订单数量
    print(f"新清单订单数: {len(new_list)}")
    
    # 按月份统计
    if '月份' in new_list.columns:
        monthly_stats = new_list.groupby('月份').agg({
            '订单金额(RMB)': ['sum', 'count']
        }).round(2)
        print("按月份统计:")
        print(monthly_stats)
    
    # 5. 检查是否还有数据质量问题
    print("\n=== 5. 数据质量检查 ===")
    
    # 检查异常大额订单
    high_value_orders = new_list[new_list['订单金额(RMB)'] > 1000000]  # 超过100万的订单
    if len(high_value_orders) > 0:
        print(f"发现 {len(high_value_orders)} 个大额订单(>100万):")
        for _, order in high_value_orders.iterrows():
            print(f"  {order['生产单号']}: {order['订单金额(RMB)']:,.2f}")
    
    # 检查订单类型分布
    if '生产单号' in new_list.columns:
        order_types = {}
        for order in new_list['生产单号']:
            prefix = str(order)[:3]
            order_types[prefix] = order_types.get(prefix, 0) + 1
        
        print("订单类型分布:")
        for prefix, count in sorted(order_types.items()):
            print(f"  {prefix}: {count}个")
    
    # 6. 结论
    print("\n=== 6. 结论 ===")
    
    if difference > 10000:  # 差异超过1亿
        print("❌ 结果差异仍然很大，可能存在以下问题:")
        print("  1. 欠料清单3.xlsx中仍有数据质量问题")
        print("  2. 订单覆盖范围可能过广")
        print("  3. 需要进一步检查数据源的准确性")
    elif difference > 5000:  # 差异5000-10000万
        print("⚠️ 结果有较大差异，需要进一步调查:")
        print("  1. 检查订单范围是否正确")
        print("  2. 验证欠料数据的准确性") 
    elif difference > 1000:  # 差异1000-5000万
        print("⚠️ 结果有一定差异，但可能在合理范围内")
        print("  1. 可能是数据更新导致的差异")
        print("  2. 建议进行抽样验证")
    else:
        print("✅ 结果差异较小，修复效果良好")
    
    return {
        'new_amount': script_output,
        'verified_amount': user_verified_amount,
        'difference': difference,
        'new_order_count': len(new_list)
    }

if __name__ == "__main__":
    results = analyze_new_results()
    
    print("\n" + "="*60)
    if results:
        if abs(results['difference']) < 1000:
            print("🎉 修复成功！结果接近用户验证数据")
        elif abs(results['difference']) < 5000:
            print("✅ 修复部分成功，但仍有差异需要调查")
        else:
            print("❌ 修复效果不理想，需要进一步分析数据源")
    print("="*60)
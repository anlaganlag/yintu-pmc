#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成带投入产出比分析的精准供应商物料报告
基于现有report + order-amt数据，新增回款分析字段
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os

def load_order_amount_data():
    """加载并合并订单金额数据"""
    try:
        # 读取两个订单金额文件
        df_amt1 = pd.read_excel('order-amt-89.xlsx')
        df_amt2 = pd.read_excel('order-amt-89-c.xlsx')
        
        # 提取关键字段并清洗
        amt1_clean = df_amt1[['生 產 單 号(  廠方 )', '订单金额']].dropna()
        amt1_clean.columns = ['生产订单号', '订单金额']
        
        amt2_clean = df_amt2[['生 產 單 号(  廠方 )', '订单金额']].dropna()
        amt2_clean.columns = ['生产订单号', '订单金额']
        
        # 合并并按生产订单号汇总
        all_amt = pd.concat([amt1_clean, amt2_clean], ignore_index=True)
        all_amt['订单金额'] = pd.to_numeric(all_amt['订单金额'], errors='coerce')
        
        # 按生产订单号汇总订单金额
        order_amt_summary = all_amt.groupby('生产订单号')['订单金额'].sum().reset_index()
        
        print(f"成功加载订单金额数据: {len(order_amt_summary)} 个PSO, 总金额: ¥{order_amt_summary['订单金额'].sum():,.2f}")
        return order_amt_summary
        
    except Exception as e:
        print(f"加载订单金额数据失败: {e}")
        return pd.DataFrame()

def generate_enhanced_report():
    """生成增强版报告"""
    try:
        # 1. 加载现有报告
        print("正在加载现有报告...")
        df_report = pd.read_excel('report精准供应商物料分析报告_20250825_1740.xlsx')
        print(f"现有报告: {len(df_report)} 行, {len(df_report.columns)} 列")
        
        # 2. 加载订单金额数据
        print("正在加载订单金额数据...")
        order_amt_data = load_order_amount_data()
        
        if order_amt_data.empty:
            print("⚠️ 订单金额数据为空，将继续生成报告但无投入产出分析")
        
        # 3. 关联数据
        print("正在关联订单金额数据...")
        # 左连接，保留所有原始记录
        enhanced_df = df_report.merge(order_amt_data, on='生产订单号', how='left')
        
        # 4. 计算新字段 - 修复按订单维度计算
        print("正在计算投入产出比...")
        
        # 新增字段1: 订单金额(RMB) - 已通过merge获得，重命名
        enhanced_df.rename(columns={'订单金额': '订单金额(RMB)'}, inplace=True)
        
        # 关键修复：按生产订单号计算正确的投入产出比
        print("修复投入产出比计算逻辑...")
        
        # 按PSO汇总欠料金额，计算正确的投入产出比
        pso_shortage_summary = enhanced_df.groupby('生产订单号')['欠料金额(RMB)'].sum().reset_index()
        pso_shortage_summary.columns = ['生产订单号', 'PSO欠料金额汇总']
        
        # 关联PSO欠料金额汇总
        enhanced_df = enhanced_df.merge(pso_shortage_summary, on='生产订单号', how='left')
        
        # 新增字段2: 每元投入回款（正确计算：整个PSO订单金额 / 整个PSO欠料金额汇总）
        enhanced_df['每元投入回款'] = enhanced_df.apply(
            lambda row: row['订单金额(RMB)'] / row['PSO欠料金额汇总'] 
            if pd.notna(row['订单金额(RMB)']) and row['PSO欠料金额汇总'] > 0 
            else None, axis=1
        )
        
        # 新增字段3: 数据完整性标记  
        enhanced_df['数据完整性标记'] = enhanced_df['订单金额(RMB)'].apply(
            lambda x: '完整' if pd.notna(x) and x > 0 else '待补充订单金额'
        )
        
        # 删除临时字段
        enhanced_df = enhanced_df.drop('PSO欠料金额汇总', axis=1)
        
        # 5. 数据质量检查
        total_records = len(enhanced_df)
        complete_records = (enhanced_df['数据完整性标记'] == '完整').sum()
        coverage_rate = complete_records / total_records * 100
        
        print(f"✅ 数据关联完成:")
        print(f"   总记录数: {total_records}")
        print(f"   完整记录数: {complete_records}")
        print(f"   数据覆盖率: {coverage_rate:.1f}%")
        
        if complete_records > 0:
            avg_return_ratio = enhanced_df[enhanced_df['数据完整性标记'] == '完整']['每元投入回款'].mean()
            print(f"   平均投入产出比: {avg_return_ratio:.2f}")
        
        # 6. 生成新文件
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        output_file = f'精准供应商物料分析报告_含回款_{timestamp}.xlsx'
        
        print(f"正在生成报告: {output_file}")
        enhanced_df.to_excel(output_file, index=False)
        
        print(f"🎉 报告生成成功!")
        print(f"📁 文件路径: {os.path.abspath(output_file)}")
        print(f"📊 字段数量: {len(enhanced_df.columns)} (新增3个分析字段)")
        
        return output_file, enhanced_df
        
    except Exception as e:
        print(f"❌ 报告生成失败: {e}")
        return None, None

if __name__ == "__main__":
    print("=== 银图PMC投入产出分析报告生成器 ===")
    print("正在整合订单金额数据，计算投入产出比...")
    print()
    
    # 检查必要文件
    required_files = [
        'report精准供应商物料分析报告_20250825_1740.xlsx',
        'order-amt-89.xlsx',
        'order-amt-89-c.xlsx'
    ]
    
    missing_files = [f for f in required_files if not os.path.exists(f)]
    if missing_files:
        print(f"❌ 缺少必要文件: {missing_files}")
        exit(1)
    
    # 生成报告
    output_file, enhanced_data = generate_enhanced_report()
    
    if output_file:
        print(f"\n✅ 任务完成！新报告已生成: {output_file}")
        print("可以在Streamlit Dashboard中使用此报告进行投入产出分析")
    else:
        print("\n❌ 任务失败，请检查错误信息")
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
深度分析欠料清单3.xlsx的数据结构
找出为什么所有订单都被认为是不缺料
"""

import pandas as pd
import numpy as np

def debug_shortage_data():
    """调试欠料数据"""
    print("=" * 60)
    print("深度分析欠料清单3.xlsx数据结构")
    print("=" * 60)
    
    # 1. 检查当前使用的欠料文件
    print("\n=== 1. 检查当前mat_owe_pso.xlsx ===")
    
    try:
        current_df = pd.read_excel('input/mat_owe_pso.xlsx', sheet_name='Sheet1', skiprows=1)
        print(f"当前欠料文件:")
        print(f"  记录数: {len(current_df)}")
        print(f"  列数: {len(current_df.columns)}")
        print(f"  列名: {current_df.columns.tolist()}")
        
        # 检查缺料情况
        if len(current_df.columns) >= 12:
            shortage_col = current_df.columns[11]  # 第12列通常是缺料数量
            shortage_data = pd.to_numeric(current_df[shortage_col], errors='coerce').fillna(0)
            has_shortage = (shortage_data > 0).sum()
            no_shortage = (shortage_data == 0).sum()
            
            print(f"  缺料列名: {shortage_col}")
            print(f"  有缺料记录: {has_shortage}")
            print(f"  无缺料记录: {no_shortage}")
            print(f"  缺料记录示例: {shortage_data[shortage_data > 0].head()}")
        
        # 检查订单覆盖
        if len(current_df.columns) >= 1:
            order_col = current_df.columns[0]
            unique_orders = current_df[order_col].nunique()
            print(f"  订单列名: {order_col}")
            print(f"  唯一订单数: {unique_orders}")
            
            # 显示前10个订单
            sample_orders = current_df[order_col].unique()[:10]
            print(f"  订单示例: {sample_orders}")
        
    except Exception as e:
        print(f"❌ 读取当前文件失败: {e}")
    
    # 2. 检查原始欠料清单3.xlsx
    print("\n=== 2. 检查原始欠料清单3.xlsx ===")
    
    try:
        # 检查工作表
        excel_file = pd.ExcelFile('欠料比较/欠料清单3.xlsx')
        print(f"工作表列表: {excel_file.sheet_names}")
        
        # 读取主要工作表
        for sheet_name in excel_file.sheet_names:
            print(f"\n--- 工作表: {sheet_name} ---")
            df = pd.read_excel('欠料比较/欠料清单3.xlsx', sheet_name=sheet_name)
            print(f"  记录数: {len(df)}")
            print(f"  列数: {len(df.columns)}")
            print(f"  列名: {df.columns.tolist()}")
            
            # 如果是主要数据表，分析缺料情况
            if '工作表2' in sheet_name or len(df) > 1000:
                # 寻找缺料相关列
                shortage_cols = [col for col in df.columns if '缺' in str(col) or '不足' in str(col) or '仓存' in str(col)]
                print(f"  疑似缺料列: {shortage_cols}")
                
                if shortage_cols:
                    shortage_col = shortage_cols[0]
                    shortage_data = pd.to_numeric(df[shortage_col], errors='coerce').fillna(0)
                    has_shortage = (shortage_data > 0).sum()
                    no_shortage = (shortage_data == 0).sum()
                    
                    print(f"  有缺料记录: {has_shortage}")
                    print(f"  无缺料记录: {no_shortage}")
                    
                    # 显示有缺料的记录样例
                    shortage_records = df[shortage_data > 0]
                    if len(shortage_records) > 0:
                        print(f"  缺料记录样例:")
                        for i, (_, row) in enumerate(shortage_records.head(3).iterrows()):
                            order_col = df.columns[0] if len(df.columns) > 0 else None
                            order_id = row[order_col] if order_col else 'Unknown'
                            shortage_amount = row[shortage_col]
                            print(f"    {order_id}: 缺料 {shortage_amount}")
                    
                # 检查订单列
                if len(df.columns) > 0:
                    order_col = df.columns[0]
                    unique_orders = df[order_col].nunique()
                    print(f"  唯一订单数: {unique_orders}")
    
    except Exception as e:
        print(f"❌ 读取原始文件失败: {e}")
    
    # 3. 对比用户验证的清单，找出应该有缺料的订单
    print("\n=== 3. 对比分析 ===")
    
    try:
        # 读取所有订单
        orders_data = []
        df1 = pd.read_excel('input/order-amt-89.xlsx', sheet_name='8月')
        df2 = pd.read_excel('input/order-amt-89.xlsx', sheet_name='9月')
        df1['生产单号'] = df1['生 產 單 号(  廠方 )']
        df2['生产单号'] = df2['生 產 單 号(  廠方 )']
        orders_data.extend([df1, df2])
        
        df3 = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name='8月 -柬')
        df4 = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name='9月 -柬')
        df3['生产单号'] = df3['生 產 單 号(  廠方 )']
        df4['生产单号'] = df4['生 產 單 号(  廠方 )']
        orders_data.extend([df3, df4])
        
        all_orders_df = pd.concat(orders_data)
        all_orders = set(all_orders_df['生产单号'].astype(str).unique())
        print(f"完整订单列表: {len(all_orders)}个")
        
        # 根据您的验证，应该有 len(all_orders) - 85 = 310个订单有缺料
        should_have_shortage = len(all_orders) - 85
        print(f"按您的验证，应该有缺料的订单: {should_have_shortage}个")
        print(f"当前分析显示有缺料的订单: 0个")
        print(f"❌ 这说明欠料数据有严重问题！")
        
    except Exception as e:
        print(f"对比分析失败: {e}")
    
    return True

if __name__ == "__main__":
    debug_shortage_data()
    
    print("\n" + "="*60)
    print("🔍 问题总结:")
    print("1. 欠料清单3.xlsx可能数据结构不正确")
    print("2. 所有订单都被错误归类为'不缺料'")
    print("3. 需要检查数据转换过程是否出错")
    print("4. 可能需要使用您提供的准确清单作为基准")
    print("="*60)
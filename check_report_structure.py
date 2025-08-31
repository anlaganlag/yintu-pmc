#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查最新分析报告文件的数据结构
"""

import pandas as pd
import glob
import os

def check_latest_report():
    """检查最新报告的结构"""
    print("=" * 60)
    print("检查最新分析报告文件结构")
    print("=" * 60)
    
    # 找到最新的报告文件
    files = glob.glob("银图PMC综合物料分析报告_*.xlsx")
    if not files:
        print("❌ 未找到分析报告文件")
        return
    
    latest_file = max(files, key=lambda x: os.path.getmtime(x))
    print(f"最新报告: {latest_file}")
    
    try:
        # 读取所有工作表
        excel_data = pd.read_excel(latest_file, sheet_name=None)
        print(f"工作表列表: {list(excel_data.keys())}")
        
        for sheet_name, df in excel_data.items():
            print(f"\n=== 工作表: {sheet_name} ===")
            print(f"行数: {len(df)}")
            print(f"列数: {len(df.columns)}")
            print("列名:")
            for i, col in enumerate(df.columns, 1):
                print(f"  {i:2d}. {col}")
            
            # 如果是主要数据表，检查关键数据
            if "明细" in sheet_name or len(df) > 1000:
                print(f"\n关键数据检查:")
                
                # 检查欠料相关列
                shortage_cols = [col for col in df.columns if '欠料' in col or '缺料' in col]
                print(f"欠料相关列: {shortage_cols}")
                
                # 检查订单相关列
                order_cols = [col for col in df.columns if '订单' in col]
                print(f"订单相关列: {order_cols}")
                
                if shortage_cols and order_cols:
                    # 分析不缺料数据
                    shortage_col = shortage_cols[0]
                    order_col = order_cols[0] if '客户订单' in order_cols[0] else order_cols[-1]
                    
                    shortage_data = pd.to_numeric(df[shortage_col], errors='coerce').fillna(0)
                    no_shortage_mask = shortage_data == 0
                    no_shortage_count = no_shortage_mask.sum()
                    
                    print(f"使用欠料列: {shortage_col}")
                    print(f"不缺料记录数: {no_shortage_count}")
                    
                    if no_shortage_count > 0:
                        no_shortage_df = df[no_shortage_mask]
                        
                        # 计算金额
                        amount_cols = [col for col in df.columns if '金额' in col and 'RMB' in col]
                        if amount_cols:
                            amount_col = [col for col in amount_cols if '订单' in col][0]
                            print(f"使用金额列: {amount_col}")
                            
                            total_amount = pd.to_numeric(no_shortage_df[amount_col], errors='coerce').fillna(0).sum()
                            print(f"不缺料订单总金额: {total_amount:,.2f} RMB ({total_amount/10000:.2f}万)")
                
                # 显示前几行数据
                print(f"\n前3行数据样例:")
                print(df.head(3))
                
    except Exception as e:
        print(f"❌ 读取失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_latest_report()
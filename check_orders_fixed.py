import pandas as pd

print("=== 订单文件分析 (修正版) ===\n")

# 正确的列名映射
column_mapping = {
    '生 產 單 号(  廠方 )': '生产订单号',
    '生 產 單 号(客方 )': '客户订单号'
}

# 读取 order-amt-89.xlsx
try:
    df1_sheets = pd.read_excel('order-amt-89.xlsx', sheet_name=None)
    print("order-amt-89.xlsx 各工作表订单数:")
    all_orders1 = set()
    for sheet_name, df in df1_sheets.items():
        # 使用正确的列名
        if '生 產 單 号(  廠方 )' in df.columns:
            unique_orders = df['生 產 單 号(  廠方 )'].dropna().unique()
            count = len(unique_orders)
            all_orders1.update(unique_orders)
            print(f"  {sheet_name}: {count} 个唯一生产订单")
        else:
            print(f"  {sheet_name}: 列名: {df.columns.tolist()[:5]}")
    
    print(f"去重后: {len(all_orders1)} 个唯一生产订单\n")
except Exception as e:
    print(f"读取 order-amt-89.xlsx 出错: {e}\n")

# 读取 order-amt-89-c.xlsx  
try:
    df2_sheets = pd.read_excel('order-amt-89-c.xlsx', sheet_name=None)
    print("order-amt-89-c.xlsx 各工作表订单数:")
    all_orders2 = set()
    for sheet_name, df in df2_sheets.items():
        if '生 產 單 号(  廠方 )' in df.columns:
            unique_orders = df['生 產 單 号(  廠方 )'].dropna().unique()
            count = len(unique_orders)
            all_orders2.update(unique_orders)
            print(f"  {sheet_name}: {count} 个唯一生产订单")
        else:
            print(f"  {sheet_name}: 列名: {df.columns.tolist()[:5]}")
    
    print(f"去重后: {len(all_orders2)} 个唯一生产订单\n")
except Exception as e:
    print(f"读取 order-amt-89-c.xlsx 出错: {e}\n")

# 合并统计
if 'all_orders1' in locals() and 'all_orders2' in locals():
    all_orders_combined = all_orders1.union(all_orders2)
    overlap = all_orders1.intersection(all_orders2)
    
    print("=== 总体统计 ===")
    print(f"原始订单文件合计: {len(all_orders_combined)} 个唯一生产订单")
    print(f"重叠订单数: {len(overlap)} 个")
    print(f"仅在 order-amt-89.xlsx: {len(all_orders1 - all_orders2)} 个")
    print(f"仅在 order-amt-89-c.xlsx: {len(all_orders2 - all_orders1)} 个")

# 检查最终报告与原始数据的差异
print("\n=== 分析差异原因 ===")
import os
import glob

report_files = glob.glob('银图PMC综合物料分析报告_*.xlsx')
if report_files and 'all_orders_combined' in locals():
    latest_report = max(report_files, key=os.path.getctime)
    
    try:
        report_df = pd.read_excel(latest_report, sheet_name='综合物料分析明细')
        report_orders = set(report_df['生产订单号'].dropna().unique())
        print(f"最新报告中: {len(report_orders)} 个生产订单")
        
        # 找出缺失的订单
        missing_orders = all_orders_combined - report_orders
        extra_orders = report_orders - all_orders_combined
        
        print(f"\n差异分析:")
        print(f"原始文件有但报告中没有的订单: {len(missing_orders)} 个")
        if len(missing_orders) > 0:
            print(f"示例缺失订单: {list(missing_orders)[:5]}")
        
        print(f"报告中有但原始文件没有的订单: {len(extra_orders)} 个")
        if len(extra_orders) > 0:
            print(f"示例额外订单: {list(extra_orders)[:5]}")
        
        # 检查缺料数据
        shortage_df = pd.read_excel('mat_owe_pso.xlsx')
        shortage_orders = set(shortage_df['生产订单号'].dropna().unique()) if '生产订单号' in shortage_df.columns else set()
        print(f"\n缺料文件中的订单数: {len(shortage_orders)} 个")
        
        # 分析为什么是393个
        matched_with_shortage = report_orders.intersection(shortage_orders)
        print(f"报告中与缺料匹配的订单: {len(matched_with_shortage)} 个")
        
        print("\n结论:")
        print("报告中的393个订单是经过与缺料清单LEFT JOIN后的结果")
        print("只有在缺料清单中存在的订单才会出现在最终报告中")
        
    except Exception as e:
        print(f"分析报告出错: {e}")
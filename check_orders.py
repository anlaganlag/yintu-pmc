import pandas as pd

# 读取两个订单文件
print("=== 订单文件分析 ===\n")

# 读取 order-amt-89.xlsx
try:
    df1_sheets = pd.read_excel('order-amt-89.xlsx', sheet_name=None)
    print("order-amt-89.xlsx 各工作表订单数:")
    total1 = 0
    all_orders1 = set()
    for sheet_name, df in df1_sheets.items():
        if '生产订单号' in df.columns:
            unique_orders = df['生产订单号'].dropna().unique()
            count = len(unique_orders)
            all_orders1.update(unique_orders)
            print(f"  {sheet_name}: {count} 个唯一订单")
            total1 += count
        else:
            print(f"  {sheet_name}: 未找到'生产订单号'列")
    
    print(f"\n小计: {total1} 个订单（重复计算）")
    print(f"去重后: {len(all_orders1)} 个唯一订单\n")
except Exception as e:
    print(f"读取 order-amt-89.xlsx 出错: {e}\n")

# 读取 order-amt-89-c.xlsx  
try:
    df2_sheets = pd.read_excel('order-amt-89-c.xlsx', sheet_name=None)
    print("order-amt-89-c.xlsx 各工作表订单数:")
    total2 = 0
    all_orders2 = set()
    for sheet_name, df in df2_sheets.items():
        if '生产订单号' in df.columns:
            unique_orders = df['生产订单号'].dropna().unique()
            count = len(unique_orders)
            all_orders2.update(unique_orders)
            print(f"  {sheet_name}: {count} 个唯一订单")
            total2 += count
        else:
            print(f"  {sheet_name}: 未找到'生产订单号'列")
    
    print(f"\n小计: {total2} 个订单（重复计算）")
    print(f"去重后: {len(all_orders2)} 个唯一订单\n")
except Exception as e:
    print(f"读取 order-amt-89-c.xlsx 出错: {e}\n")

# 合并统计
if 'all_orders1' in locals() and 'all_orders2' in locals():
    all_orders_combined = all_orders1.union(all_orders2)
    overlap = all_orders1.intersection(all_orders2)
    
    print("=== 总体统计 ===")
    print(f"两个文件合计: {len(all_orders_combined)} 个唯一订单")
    print(f"重叠订单数: {len(overlap)} 个")
    print(f"仅在 order-amt-89.xlsx: {len(all_orders1 - all_orders2)} 个")
    print(f"仅在 order-amt-89-c.xlsx: {len(all_orders2 - all_orders1)} 个")

# 检查最终报告
print("\n=== 检查生成的报告 ===")
import os
import glob

# 查找最新的分析报告
report_files = glob.glob('银图PMC综合物料分析报告_*.xlsx')
if report_files:
    latest_report = max(report_files, key=os.path.getctime)
    print(f"最新报告: {latest_report}")
    
    try:
        report_df = pd.read_excel(latest_report, sheet_name='综合物料分析明细')
        report_orders = report_df['生产订单号'].nunique()
        print(f"报告中的订单数: {report_orders} 个")
        
        # 分析差异原因
        print("\n可能的原因:")
        print("1. 部分订单可能因为数据缺失被过滤")
        print("2. 部分订单可能因为金额限制被过滤") 
        print("3. 部分订单可能因为时间范围被过滤")
        
    except Exception as e:
        print(f"读取报告出错: {e}")
else:
    print("未找到分析报告文件")
import pandas as pd
import numpy as np

def simple_comparison():
    """
    简化的对比分析，修复格式化问题
    """
    target_orders = ['PSO2500829', 'PSO2501369', 'PSO2501602', 'PSO2501060', 'PSO2501332']
    
    print("=== 5个订单的处理金额对比分析 ===\n")
    
    # 1. 从当前分析结果中提取数据
    try:
        current_file = "PMC排产Sep01-Sep07订单ROI分析_含供应商汇总.xlsx"
        df_current = pd.read_excel(current_file, sheet_name=0)
        
        print("1. 当前分析方法结果:")
        print("   文件:", current_file)
        
        current_orders = df_current[df_current['生产订单'].isin(target_orders)].copy()
        
        if len(current_orders) > 0:
            for _, row in current_orders.iterrows():
                order_no = row['生产订单']
                order_amt_usd = row.get('订单金额(USD)', 0)
                order_amt_rmb = row.get('订单金额(RMB)', 0) 
                shortage_rmb = row.get('缺料金额(RMB)', 0)
                roi = row.get('ROI显示', 'N/A')
                
                # 安全的格式化处理
                try:
                    usd_str = f"${order_amt_usd:,.2f}" if pd.notna(order_amt_usd) else "N/A"
                    rmb_str = f"¥{order_amt_rmb:,.2f}" if pd.notna(order_amt_rmb) else "N/A" 
                    shortage_str = f"¥{shortage_rmb:,.2f}" if pd.notna(shortage_rmb) else "N/A"
                except:
                    usd_str = str(order_amt_usd)
                    rmb_str = str(order_amt_rmb)
                    shortage_str = str(shortage_rmb)
                
                print(f"\n   {order_no}:")
                print(f"     订单金额(USD): {usd_str}")
                print(f"     订单金额(RMB): {rmb_str}")
                print(f"     缺料金额(RMB): {shortage_str}")
                print(f"     ROI: {roi}")
        else:
            print("   未找到目标订单")
            
    except Exception as e:
        print(f"   读取当前分析文件失败: {str(e)}")
    
    print("\n" + "="*50)
    
    # 2. 从silverPlan分析结果中提取数据
    try:
        # 查找最新的silverPlan报告
        import glob
        import os
        
        silver_files = glob.glob("银图PMC综合物料分析报告_*.xlsx")
        if silver_files:
            latest_silver = max(silver_files, key=os.path.getctime)
            print(f"\n2. silverPlan_analysis.py结果:")
            print(f"   文件: {latest_silver}")
            
            df_silver = pd.read_excel(latest_silver, sheet_name=0)
            
            # 按生产订单汇总
            silver_summary = df_silver.groupby('生产订单号').agg({
                '订单金额(USD)': 'first',
                '订单金额(RMB)': 'first', 
                '欠料金额(RMB)': 'sum',
                '每元投入回款': 'first'
            }).reset_index()
            
            silver_orders = silver_summary[silver_summary['生产订单号'].isin(target_orders)]
            
            if len(silver_orders) > 0:
                for _, row in silver_orders.iterrows():
                    order_no = row['生产订单号']
                    order_amt_usd = row.get('订单金额(USD)', 0)
                    order_amt_rmb = row.get('订单金额(RMB)', 0)
                    shortage_rmb = row.get('欠料金额(RMB)', 0)
                    roi = row.get('每元投入回款', 'N/A')
                    
                    try:
                        usd_str = f"${order_amt_usd:,.2f}" if pd.notna(order_amt_usd) else "N/A"
                        rmb_str = f"¥{order_amt_rmb:,.2f}" if pd.notna(order_amt_rmb) else "N/A"
                        shortage_str = f"¥{shortage_rmb:,.2f}" if pd.notna(shortage_rmb) else "N/A"
                    except:
                        usd_str = str(order_amt_usd)
                        rmb_str = str(order_amt_rmb)
                        shortage_str = str(shortage_rmb)
                    
                    print(f"\n   {order_no}:")
                    print(f"     订单金额(USD): {usd_str}")
                    print(f"     订单金额(RMB): {rmb_str}")
                    print(f"     缺料金额(RMB): {shortage_str}")
                    print(f"     ROI: {roi}")
            else:
                print("   未找到目标订单")
        else:
            print("\n2. silverPlan_analysis.py结果:")
            print("   未找到分析报告文件")
            
    except Exception as e:
        print(f"\n2. silverPlan_analysis.py结果:")
        print(f"   读取分析文件失败: {str(e)}")
    
    print("\n" + "="*50)
    print("\n3. 主要差异分析:")
    print("   - 当前方法: 基于合并后的汇总数据，订单级别汇总")
    print("   - silverPlan方法: 基于详细物料级别数据，然后汇总到订单")
    print("   - 金额差异可能来源:")
    print("     * 数据源范围不同（PMC vs 全量数据）") 
    print("     * 汇总层级不同（订单汇总 vs 物料汇总）")
    print("     * 缺料计算方式差异")
    print("     * 供应商选择逻辑差异")

if __name__ == "__main__":
    simple_comparison()
import pandas as pd
import numpy as np

def extract_order_data():
    """
    提取指定订单的数据进行对比分析
    """
    target_orders = ['PSO2500829', 'PSO2501369', 'PSO2501602', 'PSO2501060', 'PSO2501332']
    
    # 从当前分析结果中提取数据
    current_file = "PMC排产Sep01-Sep07订单ROI分析_含供应商汇总.xlsx"
    
    try:
        df_current = pd.read_excel(current_file, sheet_name=0)
        print("当前分析文件列名:", df_current.columns.tolist())
        
        # 提取目标订单数据
        current_orders = df_current[df_current['生产订单'].isin(target_orders)].copy()
        
        print(f"\n=== 当前分析结果 ===")
        print(f"找到订单数: {len(current_orders)}")
        
        for _, row in current_orders.iterrows():
            order_no = row['生产订单']
            order_amt_usd = row.get('订单金额(USD)', 0)
            order_amt_rmb = row.get('订单金额(RMB)', 0) 
            shortage_rmb = row.get('缺料金额(RMB)', 0)
            roi = row.get('ROI显示', 'N/A')
            supplier = row.get('供应商汇总', 'N/A')
            
            print(f"\n订单: {order_no}")
            print(f"  订单金额(USD): ${order_amt_usd:,.2f}")
            print(f"  订单金额(RMB): ¥{order_amt_rmb:,.2f}")
            print(f"  缺料金额(RMB): ¥{shortage_rmb:,.2f}")
            print(f"  ROI: {roi}")
            print(f"  供应商: {supplier}")
        
        return current_orders
        
    except Exception as e:
        print(f"读取当前分析文件时出错: {str(e)}")
        return None

def run_silverplan_comparison():
    """
    运行silverPlan_analysis.py并提取对比数据
    """
    print(f"\n=== silverPlan_analysis.py 分析结果 ===")
    
    try:
        # 执行silverPlan_analysis.py
        import subprocess
        result = subprocess.run(['python', 'silverPlan_analysis.py'], 
                              capture_output=True, text=True, encoding='utf-8')
        
        if result.returncode != 0:
            print(f"silverPlan_analysis.py 执行失败:")
            print(f"错误输出: {result.stderr}")
            return None
            
        print("silverPlan_analysis.py 执行成功")
        
        # 查找最新的分析报告
        import glob
        import os
        
        pattern = "银图PMC综合物料分析报告_*.xlsx"
        files = glob.glob(pattern)
        
        if not files:
            print(f"未找到分析报告文件，模式: {pattern}")
            return None
            
        # 选择最新文件
        latest_file = max(files, key=os.path.getctime)
        print(f"找到最新报告: {latest_file}")
        
        # 读取分析结果
        df_silver = pd.read_excel(latest_file, sheet_name=0)
        print("silverPlan分析文件列名:", df_silver.columns.tolist())
        
        target_orders = ['PSO2500829', 'PSO2501369', 'PSO2501602', 'PSO2501060', 'PSO2501332']
        
        # 查找订单号列
        order_col = None
        for col in df_silver.columns:
            if '生产订单' in col or 'PSO' in str(df_silver[col].iloc[0]):
                order_col = col
                break
                
        if order_col is None:
            print("未找到订单号列")
            return None
            
        print(f"使用订单号列: {order_col}")
        
        # 提取目标订单
        silver_orders = df_silver[df_silver[order_col].isin(target_orders)].copy()
        
        print(f"silverPlan找到订单数: {len(silver_orders)}")
        
        for _, row in silver_orders.iterrows():
            order_no = row[order_col]
            
            # 查找各种金额列
            order_amt_cols = [col for col in df_silver.columns if '订单金额' in col]
            shortage_cols = [col for col in df_silver.columns if '缺料' in col and ('金额' in col or 'RMB' in col)]
            roi_cols = [col for col in df_silver.columns if 'ROI' in col]
            supplier_cols = [col for col in df_silver.columns if '供应商' in col]
            
            print(f"\n订单: {order_no}")
            
            for col in order_amt_cols:
                print(f"  {col}: {row.get(col, 'N/A')}")
                
            for col in shortage_cols:
                value = row.get(col, 0)
                if pd.notna(value) and value != 0:
                    print(f"  {col}: ¥{value:,.2f}")
                    
            for col in roi_cols:
                print(f"  {col}: {row.get(col, 'N/A')}")
                
            for col in supplier_cols:
                print(f"  {col}: {row.get(col, 'N/A')}")
        
        return silver_orders
        
    except Exception as e:
        print(f"运行silverPlan分析时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    print("开始对比分析...")
    
    # 提取当前分析数据
    current_data = extract_order_data()
    
    # 运行silverPlan分析
    silver_data = run_silverplan_comparison()
    
    print(f"\n=== 对比总结 ===")
    if current_data is not None and silver_data is not None:
        print("两个分析结果都已获取，可以进行详细对比")
    else:
        print("无法完成完整对比，请检查文件和数据")
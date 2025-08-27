import pandas as pd
import glob

# 查找最新的报告文件
files = glob.glob('银图PMC综合物料分析报告_*.xlsx')
print(f'Found files: {files}')

if files:
    # 读取最新文件
    df = pd.read_excel(files[0], sheet_name='综合物料分析明细')
    print('\nAll columns in the Excel file:')
    for i, col in enumerate(df.columns):
        print(f"{i+1}. {col}")
    
    # 检查是否有特定的列
    required_cols = ['产品型号', '数量Pcs', '目的地']
    print('\nChecking for required columns:')
    for col in required_cols:
        if col in df.columns:
            print(f"✓ '{col}' exists")
        else:
            print(f"✗ '{col}' NOT FOUND")
else:
    print("No report files found")
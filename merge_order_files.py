import pandas as pd
import numpy as np
from datetime import datetime

# Read both Excel files
df1 = pd.read_excel('order-amt-89.xlsx')
df2 = pd.read_excel('order-amt-89-c.xlsx')

print(f"File 1: {len(df1)} rows")
print(f"File 2: {len(df2)} rows")

# Identify common columns
common_cols = list(set(df1.columns) & set(df2.columns))
print(f"\nCommon columns: {common_cols}")

# Columns only in df1
df1_only = list(set(df1.columns) - set(df2.columns))
print(f"Columns only in file 1: {df1_only}")

# Add missing columns to df2 with NaN values
for col in df1_only:
    df2[col] = np.nan

# Ensure both dataframes have the same column order
df2 = df2[df1.columns]

# Merge the dataframes
merged_df = pd.concat([df1, df2], ignore_index=True)

print(f"\nMerged result: {len(merged_df)} rows")
print(f"Columns: {merged_df.columns.tolist()}")

# Check for duplicates based on key columns
key_cols = ['生 產 單 号(  廠方 )', '型 號( 廠方/客方 )', '數 量  (Pcs)']
duplicates = merged_df[merged_df.duplicated(subset=key_cols, keep=False)]
print(f"\nDuplicate rows found: {len(duplicates)}")

# Save merged file
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
output_file = f'merged_orders_{timestamp}.xlsx'
merged_df.to_excel(output_file, index=False)
print(f"\nMerged file saved as: {output_file}")

# Generate summary statistics
print("\n=== Summary Statistics ===")
print(f"Total orders: {len(merged_df)}")
print(f"Total order amount: {merged_df['订单金额'].sum():,.2f}")
print(f"Average order amount: {merged_df['订单金额'].mean():,.2f}")
print(f"Total quantity: {merged_df['數 量  (Pcs)'].sum():,.0f}")

# Group by production order type
pso_orders = merged_df[merged_df['生 產 單 号(  廠方 )'].str.startswith('PSO', na=False)]
tso_orders = merged_df[merged_df['生 產 單 号(  廠方 )'].str.startswith('TSO', na=False)]

print(f"\nPSO orders: {len(pso_orders)} (Amount: {pso_orders['订单金额'].sum():,.2f})")
print(f"TSO orders: {len(tso_orders)} (Amount: {tso_orders['订单金额'].sum():,.2f})")
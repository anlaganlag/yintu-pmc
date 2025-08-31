#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修复欠料数据文件的VLOOKUP问题
清理并标准化mat_owe_pso.xlsx文件
"""

import pandas as pd
import numpy as np
from datetime import datetime

def fix_shortage_data():
    """修复欠料数据文件"""
    print("=== 修复欠料数据文件 ===")
    
    input_file = 'input/mat_owe_pso.xlsx'
    output_file = 'input/mat_owe_pso_fixed.xlsx'
    
    try:
        # 读取原始数据，使用正确的标题行
        print("1. 读取原始欠料数据...")
        df = pd.read_excel(input_file, skiprows=1)  # 跳过空行，使用第2行作为标题
        
        print(f"   原始数据: {len(df)}行 × {len(df.columns)}列")
        print(f"   列名: {df.columns.tolist()}")
        
        # 检查并修复VLOOKUP问题
        print("2. 检查VLOOKUP公式...")
        
        # 找到包含VLOOKUP的列
        vlookup_columns = []
        for col in df.columns:
            if df[col].astype(str).str.contains('VLOOKUP', na=False).any():
                vlookup_count = df[col].astype(str).str.contains('VLOOKUP', na=False).sum()
                vlookup_columns.append((col, vlookup_count))
                print(f"   列 [{col}] 包含 {vlookup_count} 个VLOOKUP公式")
        
        # 修复VLOOKUP公式
        if vlookup_columns:
            print("3. 修复VLOOKUP公式...")
            
            for col_name, count in vlookup_columns:
                # 将VLOOKUP公式替换为NaN（空值）
                df.loc[df[col_name].astype(str).str.contains('VLOOKUP', na=False), col_name] = np.nan
                print(f"   已将列 [{col_name}] 中的 {count} 个VLOOKUP公式替换为空值")
        
        # 验证数据质量
        print("4. 验证数据质量...")
        
        # 检查订单编号列
        if '订单编号' in df.columns:
            unique_orders = df['订单编号'].nunique()
            total_rows = len(df)
            print(f"   唯一订单数: {unique_orders}")
            print(f"   总记录数: {total_rows}")
            print(f"   平均每订单记录数: {total_rows/unique_orders:.1f}")
            
            # 统计订单类型
            order_types = {}
            for order in df['订单编号'].dropna():
                prefix = str(order)[:3] if len(str(order)) >= 3 else str(order)
                order_types[prefix] = order_types.get(prefix, 0) + 1
            
            print("   订单类型分布:")
            for prefix, count in sorted(order_types.items()):
                print(f"     {prefix}: {count}条记录")
        
        # 检查缺料数量列
        if '仓存不足' in df.columns:
            # 计算实际有缺料的记录
            shortage_records = df[pd.to_numeric(df['仓存不足'], errors='coerce') > 0]
            print(f"   有缺料记录数: {len(shortage_records)}")
            
            # 按订单统计缺料情况
            if len(shortage_records) > 0:
                shortage_orders = shortage_records['订单编号'].nunique()
                print(f"   有缺料的订单数: {shortage_orders}")
        
        # 保存修复后的文件
        print("5. 保存修复后的文件...")
        df.to_excel(output_file, index=False)
        print(f"   ✅ 修复后的文件已保存到: {output_file}")
        
        # 生成修复报告
        print("6. 生成修复报告...")
        
        report_data = {
            '检查项': [
                '原始记录数',
                '唯一订单数',
                '修复的VLOOKUP公式数',
                '有缺料记录数',
                '数据列数'
            ],
            '结果': [
                len(df),
                df['订单编号'].nunique() if '订单编号' in df.columns else 'N/A',
                sum([count for col, count in vlookup_columns]),
                len(df[pd.to_numeric(df.get('仓存不足', 0), errors='coerce') > 0]),
                len(df.columns)
            ]
        }
        
        report_df = pd.DataFrame(report_data)
        report_file = f'欠料数据修复报告_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
            report_df.to_excel(writer, sheet_name='修复报告', index=False)
            
            # 添加示例数据
            if len(df) > 0:
                sample_df = df.head(100)  # 前100行作为示例
                sample_df.to_excel(writer, sheet_name='数据示例', index=False)
        
        print(f"   ✅ 修复报告已保存到: {report_file}")
        
        return True
        
    except Exception as e:
        print(f"   ❌ 修复失败: {e}")
        return False

def validate_fixed_data():
    """验证修复后的数据"""
    print("\n=== 验证修复后的数据 ===")
    
    try:
        fixed_file = 'input/mat_owe_pso_fixed.xlsx'
        df = pd.read_excel(fixed_file)
        
        print(f"修复后数据: {len(df)}行 × {len(df.columns)}列")
        
        # 检查是否还有VLOOKUP
        vlookup_found = False
        for col in df.columns:
            if df[col].astype(str).str.contains('VLOOKUP', na=False).any():
                vlookup_count = df[col].astype(str).str.contains('VLOOKUP', na=False).sum()
                print(f"⚠️ 列 [{col}] 仍包含 {vlookup_count} 个VLOOKUP")
                vlookup_found = True
        
        if not vlookup_found:
            print("✅ 确认：文件中不再包含VLOOKUP公式")
        
        # 验证订单数据
        if '订单编号' in df.columns:
            orders = df['订单编号'].dropna().unique()
            print(f"✅ 验证：包含 {len(orders)} 个唯一订单")
            
            # 展示订单示例
            sample_orders = list(orders)[:10]
            print(f"订单示例: {sample_orders}")
        
        return True
        
    except Exception as e:
        print(f"❌ 验证失败: {e}")
        return False

if __name__ == "__main__":
    print("欠料数据修复工具")
    print("=" * 50)
    
    if fix_shortage_data():
        validate_fixed_data()
        print("\n🎉 欠料数据修复完成！")
        print("\n建议:")
        print("1. 使用修复后的文件 'input/mat_owe_pso_fixed.xlsx' 进行分析")
        print("2. 检查修复报告了解详细信息")
        print("3. 重新运行PMC分析脚本验证结果")
    else:
        print("\n❌ 修复失败，请检查错误信息")
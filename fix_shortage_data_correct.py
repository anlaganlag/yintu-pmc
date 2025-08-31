#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正确转换欠料清单3.xlsx，保持列结构完整
修复列名错乱和数据错位问题
"""

import pandas as pd
import numpy as np
from datetime import datetime

def fix_shortage_data_correct():
    """正确修复欠料数据"""
    print("=" * 60)
    print("正确转换欠料清单3.xlsx")
    print("=" * 60)
    
    try:
        # 1. 读取原始欠料清单3.xlsx的工作表2
        print("\n=== 1. 读取原始欠料清单3.xlsx ===")
        df = pd.read_excel('欠料比较/欠料清单3.xlsx', sheet_name='工作表2')
        
        print(f"原始数据: {len(df)}行 × {len(df.columns)}列")
        print(f"原始列名:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1:2d}. {col}")
        
        # 2. 标准化列名 - 移除换行符并简化
        print("\n=== 2. 标准化列名 ===")
        
        # 创建标准列名映射
        standard_columns = [
            '序号',                    # 序
            '订单编号',                # 訂單編號(已勾測料)
            '客户型号',                # 客型號(生產排程)  
            'OTS期',                   # OTS期(生產排程)
            '开拉期',                  # 開拉期(生產排程)
            '下单日期',                # 下單日期(WO創建日)
            '物料编号',                # 物料編號
            '物料名称',                # 物項名称
            '领用部门',                # 領用部門
            '工单需求',                # 工單需求(已估計)
            '仓存不足',                # 倉存不足(齊套料) - 关键的缺料列
            '已购未返',                # 已購未返(庫存表)
            '手头现有',                # 手頭現有（多倉）
            '请购组',                  # 請購組po采購
            '代用品',                  # 代用品
            '最新最低单价',            # 最新最低單價
            '供应商号',                # 供應商號
            '供应商名称',              # 供應商名稱
            '缺料金额',                # 缺料需求金額
            '币种'                     # 幣種
        ]
        
        # 应用新列名
        df.columns = standard_columns[:len(df.columns)]
        
        print("标准化后的列名:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1:2d}. {col}")
        
        # 3. 数据清理
        print("\n=== 3. 数据清理 ===")
        
        # 清理订单编号
        df['订单编号'] = df['订单编号'].astype(str).str.strip()
        
        # 过滤有效订单
        valid_orders_mask = df['订单编号'].str.contains('PSO|MSO|RSO|TSO|FOR|SP', na=False)
        original_count = len(df)
        df = df[valid_orders_mask]
        filtered_count = len(df)
        
        print(f"过滤无效订单: {original_count} → {filtered_count} 条记录")
        print(f"唯一订单数: {df['订单编号'].nunique()}")
        
        # 处理缺料数量 - 这是关键！
        df['仓存不足'] = pd.to_numeric(df['仓存不足'], errors='coerce').fillna(0)
        
        has_shortage_count = (df['仓存不足'] > 0).sum()
        no_shortage_count = (df['仓存不足'] == 0).sum()
        
        print(f"有缺料记录: {has_shortage_count}条")
        print(f"无缺料记录: {no_shortage_count}条")
        
        # 显示缺料情况样例
        shortage_samples = df[df['仓存不足'] > 0][['订单编号', '物料编号', '仓存不足']].head()
        print("缺料记录样例:")
        print(shortage_samples)
        
        # 4. 补充silverPlan_analysis.py需要的标准列
        print("\n=== 4. 补充标准列 ===")
        
        # 添加P-R相关列（silverPlan_analysis.py需要）
        df.insert(1, 'P-R对应', '')
        df.insert(2, 'P-RBOM', '')
        
        # 重新排序列，确保与silverPlan_analysis.py期望的顺序一致
        expected_columns = [
            '订单编号', 'P-R对应', 'P-RBOM', '客户型号', 'OTS期', '开拉期',
            '下单日期', '物料编号', '物料名称', '领用部门', '工单需求',
            '仓存不足', '已购未返', '手头现有', '请购组'
        ]
        
        # 保持这些列，其他列追加到后面
        available_expected = [col for col in expected_columns if col in df.columns]
        remaining_cols = [col for col in df.columns if col not in expected_columns]
        final_columns = available_expected + remaining_cols
        
        df = df[final_columns]
        print(f"最终列结构: {df.columns.tolist()}")
        
        # 5. 验证数据质量
        print("\n=== 5. 验证数据质量 ===")
        
        # 统计订单类型
        order_types = {}
        for order in df['订单编号'].unique():
            prefix = order[:3] if len(order) >= 3 else order
            order_types[prefix] = order_types.get(prefix, 0) + 1
        
        print("订单类型分布:")
        for prefix, count in sorted(order_types.items()):
            print(f"  {prefix}: {count}个订单")
        
        # 统计缺料情况
        shortage_orders = df[df['仓存不足'] > 0]['订单编号'].nunique()
        no_shortage_orders = df[df['仓存不足'] == 0]['订单编号'].nunique()
        
        print(f"有缺料的订单数: {shortage_orders}")
        print(f"无缺料的订单数: {no_shortage_orders}")
        print(f"总订单数: {df['订单编号'].nunique()}")
        
        # 6. 备份原文件并保存新文件
        print("\n=== 6. 保存修复后的文件 ===")
        
        # 备份当前文件
        try:
            backup_name = f"input/mat_owe_pso_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            import shutil
            shutil.copy('input/mat_owe_pso.xlsx', backup_name)
            print(f"已备份原文件: {backup_name}")
        except:
            print("备份原文件失败（可能不存在）")
        
        # 保存新文件，格式要与silverPlan_analysis.py兼容
        output_file = 'input/mat_owe_pso.xlsx'
        
        # 创建Excel文件，第一行作为表头
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 写入主数据，第一行留空作为标题行
            df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1)
            
            # 在第一行写入列标题（silverPlan_analysis.py会跳过第一行）
            worksheet = writer.sheets['Sheet1']
            for col_idx, column in enumerate(df.columns, 1):
                worksheet.cell(row=1, column=col_idx).value = column
        
        print(f"✅ 修复完成！文件已保存到: {output_file}")
        
        # 7. 验证修复结果
        print("\n=== 7. 验证修复结果 ===")
        
        # 模拟silverPlan_analysis.py的读取方式
        test_df = pd.read_excel(output_file, sheet_name='Sheet1', skiprows=1)
        print(f"验证读取: {len(test_df)}条记录")
        print(f"验证列名: {test_df.columns.tolist()[:5]}...")
        
        # 检查缺料数据
        if len(test_df.columns) >= 12:
            shortage_col = test_df.columns[11]  # 第12列应该是缺料列
            shortage_data = pd.to_numeric(test_df[shortage_col], errors='coerce').fillna(0)
            has_shortage = (shortage_data > 0).sum()
            
            print(f"验证缺料列: {shortage_col}")
            print(f"验证有缺料记录数: {has_shortage}")
            
            if has_shortage > 0:
                print("✅ 缺料数据正确保留")
                return True
            else:
                print("❌ 缺料数据丢失")
                return False
        
        return True
        
    except Exception as e:
        print(f"❌ 修复失败: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = fix_shortage_data_correct()
    
    print("\n" + "="*60)
    if success:
        print("🎉 数据修复完成！")
        print("现在重新运行分析应该能得到正确结果")
        print("\n建议操作:")
        print("python silverPlan_analysis.py")
    else:
        print("❌ 数据修复失败")
    print("="*60)
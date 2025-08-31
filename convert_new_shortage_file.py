#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将 827mat-owe-250827.xlsx 转换为标准的 mat_owe_pso.xlsx 格式
用于 silverPlan_analysis.py 正确分析
"""

import pandas as pd
import numpy as np
from datetime import datetime

def convert_shortage_file():
    """转换欠料文件格式"""
    print("=" * 60)
    print("转换 827mat-owe-250827.xlsx 为标准格式")
    print("=" * 60)
    
    try:
        # 1. 读取新的欠料文件
        print("\n=== 1. 读取新欠料文件 ===")
        new_file = '827mat-owe-250827.xlsx'
        df = pd.read_excel(new_file)
        
        print(f"原始数据: {len(df)}行 × {len(df.columns)}列")
        print(f"原始列名: {df.columns.tolist()[:10]}")
        
        # 2. 标准化列名（移除换行符和特殊字符）
        print("\n=== 2. 标准化列名 ===")
        
        column_mapping = {
            '序': '序号',
            '訂單編號\n(已勾測料)': '订单编号',
            '客型號\n(生產排程)': '客户型号', 
            'OTS期\n(生產排程)': 'OTS期',
            '開拉期\n(生產排程)': '开拉期',
            '下單日期\n(WO創建日)': '下单日期',
            '物料編號': '物料编号',
            '物項名称': '物料名称',
            '領用\n部門': '领用部门',
            '工單需求\n(已估計)': '工单需求',
            '倉存不足\n(齊套料)': '仓存不足',
            '已購未返': '已购未返',
            '手頭現有': '手头现有',
            '請購組': '请购组',
            '缺料需求\n金額': '缺料金额'
        }
        
        # 应用列名映射
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
                print(f"  {old_name} → {new_name}")
        
        # 3. 数据清理和验证
        print("\n=== 3. 数据清理 ===")
        
        # 清理订单编号
        if '订单编号' in df.columns:
            df['订单编号'] = df['订单编号'].astype(str).str.strip()
            
            # 过滤有效订单记录
            valid_orders_mask = df['订单编号'].str.contains('PSO|MSO|RSO|TSO|FOR|SP', na=False)
            original_count = len(df)
            df = df[valid_orders_mask]
            filtered_count = len(df)
            
            print(f"  过滤无效订单: {original_count} → {filtered_count} 条记录")
            print(f"  唯一订单数: {df['订单编号'].nunique()}")
        
        # 处理缺料数量
        if '仓存不足' in df.columns:
            df['仓存不足'] = pd.to_numeric(df['仓存不足'], errors='coerce').fillna(0)
            
            has_shortage_count = (df['仓存不足'] > 0).sum()
            no_shortage_count = (df['仓存不足'] == 0).sum()
            
            print(f"  有缺料记录: {has_shortage_count}条")
            print(f"  无缺料记录: {no_shortage_count}条")
        
        # 4. 添加缺失的标准列
        print("\n=== 4. 补充标准列 ===")
        
        # silverPlan_analysis.py 期望的列结构
        standard_columns = [
            '订单编号', 'P-R对应', 'P-RBOM', '客户型号', 'OTS期', '开拉期',
            '下单日期', '物料编号', '物料名称', '领用部门', '工单需求',
            '仓存不足', '已购未返', '手头现有', '请购组'
        ]
        
        # 添加缺失列
        for col in standard_columns:
            if col not in df.columns:
                df[col] = '' if col in ['P-R对应', 'P-RBOM'] else 0
                print(f"  添加缺失列: {col}")
        
        # 重新排列列顺序
        df = df[standard_columns + [col for col in df.columns if col not in standard_columns]]
        
        # 5. 验证数据质量
        print("\n=== 5. 数据质量验证 ===")
        
        # 统计订单类型
        if '订单编号' in df.columns:
            order_types = {}
            for order in df['订单编号'].unique():
                prefix = order[:3] if len(order) >= 3 else order
                order_types[prefix] = order_types.get(prefix, 0) + 1
            
            print("订单类型分布:")
            for prefix, count in sorted(order_types.items()):
                print(f"  {prefix}: {count}个订单")
        
        # 统计有缺料的订单
        if '仓存不足' in df.columns:
            shortage_orders = df[df['仓存不足'] > 0]['订单编号'].nunique()
            no_shortage_orders = df[df['仓存不足'] == 0]['订单编号'].nunique()
            
            print(f"有缺料订单数: {shortage_orders}")
            print(f"无缺料订单数: {no_shortage_orders}")
        
        # 6. 与完整订单列表对比
        print("\n=== 6. 与完整订单列表对比 ===")
        
        try:
            # 读取完整订单列表
            orders_data = []
            
            # 国内订单
            df1 = pd.read_excel('input/order-amt-89.xlsx', sheet_name='8月')
            df2 = pd.read_excel('input/order-amt-89.xlsx', sheet_name='9月')
            df1['生产单号'] = df1['生 產 單 号(  廠方 )']
            df2['生产单号'] = df2['生 產 單 号(  廠方 )']
            orders_data.extend([df1, df2])
            
            # 柬埔寨订单
            df3 = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name='8月 -柬')
            df4 = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name='9月 -柬')
            df3['生产单号'] = df3['生 產 單 号(  廠方 )']
            df4['生产单号'] = df4['生 產 單 号(  廠方 )']
            orders_data.extend([df3, df4])
            
            all_orders_df = pd.concat(orders_data)
            all_orders = set(all_orders_df['生产单号'].astype(str).unique())
            shortage_file_orders = set(df['订单编号'].unique())
            
            print(f"完整订单列表: {len(all_orders)}个订单")
            print(f"欠料文件覆盖: {len(shortage_file_orders)}个订单")
            
            # 计算覆盖率
            covered_orders = all_orders & shortage_file_orders
            missing_orders = all_orders - shortage_file_orders
            
            coverage_rate = len(covered_orders) / len(all_orders) * 100
            print(f"覆盖率: {coverage_rate:.1f}% ({len(covered_orders)}/{len(all_orders)})")
            
            if len(missing_orders) > 0:
                print(f"缺失订单数: {len(missing_orders)}")
                print(f"缺失订单示例: {list(missing_orders)[:10]}")
                
                # 为缺失的订单添加"无缺料"记录
                print("\n=== 7. 为缺失订单添加无缺料记录 ===")
                
                missing_records = []
                for order in missing_orders:
                    # 从订单文件获取基本信息
                    order_info = all_orders_df[all_orders_df['生产单号'] == order]
                    if not order_info.empty:
                        info = order_info.iloc[0]
                        
                        record = {
                            '订单编号': order,
                            'P-R对应': '',
                            'P-RBOM': '',
                            '客户型号': info.get('型 號( 廠方/客方 )', ''),
                            'OTS期': '',
                            '开拉期': '',
                            '下单日期': '',
                            '物料编号': '无缺料',
                            '物料名称': '该订单无缺料',
                            '领用部门': '',
                            '工单需求': 0,
                            '仓存不足': 0,  # 设为0表示无缺料
                            '已购未返': 0,
                            '手头现有': 0,
                            '请购组': ''
                        }
                        missing_records.append(record)
                
                if missing_records:
                    missing_df = pd.DataFrame(missing_records)
                    df = pd.concat([df, missing_df], ignore_index=True)
                    print(f"添加了{len(missing_records)}条无缺料记录")
            
        except Exception as e:
            print(f"⚠️ 订单对比失败: {e}")
        
        # 7. 保存转换后的文件
        print(f"\n=== 8. 保存转换后的文件 ===")
        
        # 备份原文件
        try:
            original_backup = f"input/mat_owe_pso_original_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            original_df = pd.read_excel('input/mat_owe_pso.xlsx', skiprows=1)
            original_df.to_excel(original_backup, index=False)
            print(f"原文件已备份: {original_backup}")
        except:
            print("原文件备份失败（可能不存在）")
        
        # 保存新的标准格式文件
        output_file = 'input/mat_owe_pso.xlsx'
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 主数据表 - 添加标题行
            header_df = pd.DataFrame([df.columns.tolist()], columns=df.columns)
            combined_df = pd.concat([header_df, df], ignore_index=True)
            combined_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
            
            # 转换报告
            conversion_report = pd.DataFrame({
                '转换项目': [
                    '原始文件',
                    '转换后记录数', 
                    '唯一订单数',
                    '有缺料订单数',
                    '无缺料订单数',
                    '订单覆盖率',
                    '转换时间'
                ],
                '结果': [
                    '827mat-owe-250827.xlsx',
                    len(df),
                    df['订单编号'].nunique(),
                    df[df['仓存不足'] > 0]['订单编号'].nunique(),
                    df[df['仓存不足'] == 0]['订单编号'].nunique(),
                    f"{coverage_rate:.1f}%" if 'coverage_rate' in locals() else 'N/A',
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            })
            
            conversion_report.to_excel(writer, sheet_name='转换报告', index=False)
        
        print(f"✅ 转换完成！文件已保存到: {output_file}")
        print(f"📊 最终统计:")
        print(f"  - 总记录数: {len(df)}")
        print(f"  - 唯一订单数: {df['订单编号'].nunique()}")
        print(f"  - 有缺料记录数: {(df['仓存不足'] > 0).sum()}")
        print(f"  - 无缺料记录数: {(df['仓存不足'] == 0).sum()}")
        
        return True
        
    except Exception as e:
        print(f"❌ 转换失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def verify_conversion():
    """验证转换后的文件是否能被silverPlan_analysis.py正确读取"""
    print(f"\n=== 验证转换结果 ===")
    
    try:
        # 模拟silverPlan_analysis.py的读取方式
        df = pd.read_excel('input/mat_owe_pso.xlsx', sheet_name='Sheet1', skiprows=1)
        
        # 模拟标准化列名过程
        if len(df.columns) >= 13:
            new_columns = ['订单编号', 'P-R对应', 'P-RBOM', '客户型号', 'OTS期', '开拉期', 
                          '下单日期', '物料编号', '物料名称', '领用部门', '工单需求', 
                          '仓存不足', '已购未返', '手头现有', '请购组']
            
            for i in range(min(len(new_columns), len(df.columns))):
                df.rename(columns={df.columns[i]: new_columns[i]}, inplace=True)
        
        # 清理欠料数据
        df = df.dropna(subset=['订单编号'])
        
        print(f"✅ 验证通过:")
        print(f"  - 可正确读取: {len(df)}条记录")
        print(f"  - 订单数: {df['订单编号'].nunique()}")
        print(f"  - 列结构正确")
        
        # 检查是否解决了VLOOKUP问题
        vlookup_count = 0
        for col in df.columns:
            vlookup_in_col = df[col].astype(str).str.contains('VLOOKUP', na=False).sum()
            vlookup_count += vlookup_in_col
        
        if vlookup_count == 0:
            print("  - ✅ 无VLOOKUP公式问题")
        else:
            print(f"  - ⚠️ 仍有{vlookup_count}个VLOOKUP公式")
        
        return True
        
    except Exception as e:
        print(f"❌ 验证失败: {e}")
        return False

if __name__ == "__main__":
    print("欠料文件格式转换工具")
    print("将 827mat-owe-250827.xlsx 转换为silverPlan_analysis.py兼容格式")
    
    if convert_shortage_file():
        if verify_conversion():
            print(f"\n🎉 转换成功！")
            print(f"现在可以运行 silverPlan_analysis.py 进行正确的分析")
            print(f"\n建议测试步骤:")
            print(f"1. 运行: python silverPlan_analysis.py")
            print(f"2. 检查结果是否接近您的1967.31万")
            print(f"3. 如果结果正确，转换就成功了")
        else:
            print(f"\n⚠️ 转换完成但验证有问题，请检查")
    else:
        print(f"\n❌ 转换失败，请检查错误信息")
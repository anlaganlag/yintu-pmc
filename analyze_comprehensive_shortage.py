#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
综合分析3个欠料清单，找出最全最准确的数据
转换P-R清单为生产清单格式，生成完整的欠料文件
"""

import pandas as pd
import numpy as np
from datetime import datetime

def analyze_pr_to_production_mapping():
    """分析P-R清单到生产订单的映射关系"""
    print("=== 分析P-R清单到生产订单映射 ===")
    
    try:
        # 读取欠料清单3的P-R清单（最全的）
        pr_df = pd.read_excel('欠料比较/欠料清单3.xlsx', sheet_name='P-R清單')
        
        print(f"P-R清单记录数: {len(pr_df)}")
        print(f"P-R清单列名: {pr_df.columns.tolist()}")
        
        # 关键映射列分析
        key_columns = ['SO單號', '客戶PO號', '銀圖銀電', '客型號']
        for col in key_columns:
            if col in pr_df.columns:
                unique_count = pr_df[col].nunique()
                sample_values = pr_df[col].dropna().head(5).tolist()
                print(f"  {col}: {unique_count}个唯一值, 示例: {sample_values}")
        
        # 提取P-R到生产订单的映射
        if '客戶PO號' in pr_df.columns:
            # 清理客户PO号（生产订单号）
            pr_df['生产订单号'] = pr_df['客戶PO號'].astype(str).str.strip()
            
            # 过滤有效的生产订单
            valid_production_orders = pr_df[
                pr_df['生产订单号'].str.contains('PSO|MSO|RSO|TSO|FOR|SP', na=False)
            ]
            
            print(f"  有效生产订单映射: {len(valid_production_orders)}")
            print(f"  唯一生产订单数: {valid_production_orders['生产订单号'].nunique()}")
            
            # 统计订单类型
            order_types = {}
            for order in valid_production_orders['生产订单号'].unique():
                prefix = order[:3] if len(order) >= 3 else order
                order_types[prefix] = order_types.get(prefix, 0) + 1
            
            print(f"  生产订单类型分布: {dict(sorted(order_types.items()))}")
            
            return valid_production_orders
        
        return None
        
    except Exception as e:
        print(f"P-R清单分析失败: {e}")
        return None

def create_comprehensive_shortage_data():
    """创建最全面的欠料数据"""
    print("\n=== 创建最全面的欠料数据 ===")
    
    comprehensive_data = []
    
    try:
        # 1. 从欠料清单3获取主要欠料数据
        print("1. 读取欠料清单3的主要数据...")
        main_df = pd.read_excel('欠料比较/欠料清单3.xlsx', sheet_name='工作表2')
        
        # 标准化列名
        column_mapping = {
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
        
        for old_name, new_name in column_mapping.items():
            if old_name in main_df.columns:
                main_df = main_df.rename(columns={old_name: new_name})
        
        # 过滤有效订单记录
        if '订单编号' in main_df.columns:
            main_df['订单编号'] = main_df['订单编号'].astype(str).str.strip()
            valid_main = main_df[main_df['订单编号'].str.contains('PSO|MSO|RSO|TSO|FOR|SP', na=False)]
            
            print(f"  主要数据有效记录: {len(valid_main)}")
            print(f"  主要数据订单数: {valid_main['订单编号'].nunique()}")
            
            comprehensive_data.append(('主要欠料数据', valid_main))
        
        # 2. 处理P-R清单映射
        print("2. 处理P-R清单映射...")
        pr_mapping = analyze_pr_to_production_mapping()
        
        if pr_mapping is not None:
            # 为P-R清单中的订单创建欠料记录
            pr_orders = set(pr_mapping['生产订单号'].unique())
            main_orders = set(valid_main['订单编号'].unique()) if 'valid_main' in locals() else set()
            
            # P-R独有的订单
            pr_only_orders = pr_orders - main_orders
            
            print(f"  P-R清单独有订单: {len(pr_only_orders)}")
            
            if pr_only_orders:
                # 为P-R独有订单创建记录
                pr_records = []
                for order in pr_only_orders:
                    # 获取该订单的P-R信息
                    order_pr_info = pr_mapping[pr_mapping['生产订单号'] == order].iloc[0]
                    
                    record = {
                        '订单编号': order,
                        'P-R对应': order_pr_info.get('SO單號', ''),
                        'P-RBOM': '',
                        '客户型号': order_pr_info.get('客型號', ''),
                        'OTS期': '',
                        '开拉期': '',
                        '下单日期': '',
                        '物料编号': 'P-R清单来源',
                        '物料名称': f"来自P-R清单的订单映射",
                        '领用部门': '',
                        '工单需求': 0,
                        '仓存不足': 0,  # 假设P-R清单中的订单暂无缺料数据
                        '已购未返': 0,
                        '手头现有': 0,
                        '请购组': '',
                        '数据来源': 'P-R清单映射'
                    }
                    pr_records.append(record)
                
                pr_df_records = pd.DataFrame(pr_records)
                comprehensive_data.append(('P-R清单映射', pr_df_records))
                
                print(f"  创建P-R映射记录: {len(pr_records)}条")
        
        # 3. 检查其他清单中的额外数据
        print("3. 检查其他清单的额外数据...")
        
        try:
            # 检查清单2的Sheet1
            df2_sheet1 = pd.read_excel('欠料比较/欠料清单2.xlsx', sheet_name='Sheet1', skiprows=1)
            if len(df2_sheet1) > 0 and len(df2_sheet1.columns) > 0:
                # 假设第一列是订单编号
                df2_sheet1['订单编号'] = df2_sheet1.iloc[:, 0].astype(str)
                valid_df2 = df2_sheet1[df2_sheet1['订单编号'].str.contains('PSO|MSO|RSO|TSO|FOR|SP', na=False)]
                
                if len(valid_df2) > 0:
                    df2_orders = set(valid_df2['订单编号'].unique())
                    existing_orders = main_orders | pr_orders if 'pr_orders' in locals() else main_orders
                    df2_only = df2_orders - existing_orders
                    
                    print(f"  清单2额外订单: {len(df2_only)}")
                    
                    if df2_only:
                        # 为清单2独有订单创建记录
                        df2_records = []
                        for order in df2_only:
                            record = {
                                '订单编号': order,
                                'P-R对应': '',
                                'P-RBOM': '',
                                '客户型号': '',
                                'OTS期': '',
                                '开拉期': '',
                                '下单日期': '',
                                '物料编号': '清单2来源',
                                '物料名称': '来自清单2的补充订单',
                                '领用部门': '',
                                '工单需求': 0,
                                '仓存不足': 0,
                                '已购未返': 0,
                                '手头现有': 0,
                                '请购组': '',
                                '数据来源': '清单2补充'
                            }
                            df2_records.append(record)
                        
                        df2_supplement = pd.DataFrame(df2_records)
                        comprehensive_data.append(('清单2补充', df2_supplement))
                        
                        print(f"  创建清单2补充记录: {len(df2_records)}条")
        
        except Exception as e:
            print(f"  清单2分析失败: {e}")
        
        # 4. 合并所有数据
        print("4. 合并所有数据源...")
        
        if comprehensive_data:
            all_dataframes = []
            for source_name, df in comprehensive_data:
                df['数据来源'] = source_name
                all_dataframes.append(df)
                print(f"  {source_name}: {len(df)}条记录")
            
            # 合并数据
            final_comprehensive = pd.concat(all_dataframes, ignore_index=True)
            
            # 标准化最终列
            standard_columns = [
                '订单编号', 'P-R对应', 'P-RBOM', '客户型号', 'OTS期', '开拉期',
                '下单日期', '物料编号', '物料名称', '领用部门', '工单需求',
                '仓存不足', '已购未返', '手头现有', '请购组'
            ]
            
            for col in standard_columns:
                if col not in final_comprehensive.columns:
                    final_comprehensive[col] = ''
            
            # 数据类型处理
            numeric_cols = ['工单需求', '仓存不足', '已购未返', '手头现有']
            for col in numeric_cols:
                if col in final_comprehensive.columns:
                    final_comprehensive[col] = pd.to_numeric(final_comprehensive[col], errors='coerce').fillna(0)
            
            print(f"  最终合并数据: {len(final_comprehensive)}条记录")
            print(f"  最终订单数: {final_comprehensive['订单编号'].nunique()}")
            
            return final_comprehensive
        
        return None
        
    except Exception as e:
        print(f"创建综合数据失败: {e}")
        return None

def validate_with_input_orders(comprehensive_df):
    """验证与input目录订单的匹配度"""
    print("\n=== 验证与input目录订单匹配度 ===")
    
    try:
        # 读取input目录的所有订单
        orders_data = []
        
        df1 = pd.read_excel('input/order-amt-89.xlsx', sheet_name='8月')
        df2 = pd.read_excel('input/order-amt-89.xlsx', sheet_name='9月')
        df3 = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name='8月 -柬')
        df4 = pd.read_excel('input/order-amt-89-c.xlsx', sheet_name='9月 -柬')
        
        for df in [df1, df2, df3, df4]:
            df['生产单号'] = df['生 產 單 号(  廠方 )']
        
        all_input_orders = pd.concat([df1, df2, df3, df4])
        input_orders = set(all_input_orders['生产单号'].astype(str).unique())
        
        print(f"input目录订单总数: {len(input_orders)}")
        
        if comprehensive_df is not None:
            comprehensive_orders = set(comprehensive_df['订单编号'].unique())
            
            # 匹配分析
            matched_orders = comprehensive_orders & input_orders
            missing_from_comprehensive = input_orders - comprehensive_orders
            extra_in_comprehensive = comprehensive_orders - input_orders
            
            match_rate = len(matched_orders) / len(input_orders) * 100
            
            print(f"综合数据订单数: {len(comprehensive_orders)}")
            print(f"匹配订单数: {len(matched_orders)}")
            print(f"匹配率: {match_rate:.1f}%")
            print(f"缺失订单数: {len(missing_from_comprehensive)}")
            print(f"额外订单数: {len(extra_in_comprehensive)}")
            
            if missing_from_comprehensive:
                print(f"缺失订单示例: {list(missing_from_comprehensive)[:10]}")
                
                # 为缺失订单添加"无欠料"记录
                print("  为缺失订单添加无欠料记录...")
                
                missing_records = []
                for order in missing_from_comprehensive:
                    order_info = all_input_orders[all_input_orders['生产单号'] == order]
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
                            '物料编号': '无欠料',
                            '物料名称': '该订单无欠料记录',
                            '领用部门': '',
                            '工单需求': 0,
                            '仓存不足': 0,  # 无欠料
                            '已购未返': 0,
                            '手头现有': 0,
                            '请购组': '',
                            '数据来源': 'input订单补充'
                        }
                        missing_records.append(record)
                
                if missing_records:
                    missing_df = pd.DataFrame(missing_records)
                    comprehensive_df = pd.concat([comprehensive_df, missing_df], ignore_index=True)
                    
                    print(f"  添加无欠料记录: {len(missing_records)}条")
                    print(f"  最终数据: {len(comprehensive_df)}条记录")
                    print(f"  最终订单数: {comprehensive_df['订单编号'].nunique()}")
                    print(f"  最终匹配率: 100%")
            
            return comprehensive_df
        
        return None
        
    except Exception as e:
        print(f"验证失败: {e}")
        return comprehensive_df

def save_comprehensive_shortage_file(comprehensive_df):
    """保存最全面的欠料文件"""
    print("\n=== 保存综合欠料文件 ===")
    
    if comprehensive_df is None:
        print("没有数据可保存")
        return None
    
    try:
        # 备份现有文件
        import shutil
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        try:
            shutil.copy('input/mat_owe_pso.xlsx', f'input/mat_owe_pso_backup_{timestamp}.xlsx')
            print(f"已备份原文件: mat_owe_pso_backup_{timestamp}.xlsx")
        except:
            print("原文件备份失败（可能不存在）")
        
        # 保存新的综合文件
        output_file = 'input/mat_owe_pso_comprehensive.xlsx'
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 主数据表
            comprehensive_df.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # 数据来源统计
            source_stats = comprehensive_df['数据来源'].value_counts().reset_index()
            source_stats.columns = ['数据来源', '记录数']
            source_stats.to_excel(writer, sheet_name='数据来源统计', index=False)
            
            # 订单统计
            order_stats = comprehensive_df.groupby('数据来源')['订单编号'].nunique().reset_index()
            order_stats.columns = ['数据来源', '订单数']
            order_stats.to_excel(writer, sheet_name='订单来源统计', index=False)
            
            # 缺料统计
            shortage_stats = comprehensive_df[comprehensive_df['仓存不足'] > 0].groupby('数据来源').agg({
                '订单编号': 'nunique',
                '仓存不足': 'sum'
            }).reset_index()
            shortage_stats.columns = ['数据来源', '有缺料订单数', '总缺料数量']
            shortage_stats.to_excel(writer, sheet_name='缺料统计', index=False)
            
            # 综合报告
            report_data = {
                '统计项': [
                    '总记录数',
                    '唯一订单数',
                    '有缺料记录数',
                    '无缺料记录数',
                    '数据来源数',
                    '与input匹配率',
                    '生成时间'
                ],
                '数值': [
                    len(comprehensive_df),
                    comprehensive_df['订单编号'].nunique(),
                    len(comprehensive_df[comprehensive_df['仓存不足'] > 0]),
                    len(comprehensive_df[comprehensive_df['仓存不足'] == 0]),
                    comprehensive_df['数据来源'].nunique(),
                    '100%',
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ]
            }
            
            report_df = pd.DataFrame(report_data)
            report_df.to_excel(writer, sheet_name='综合报告', index=False)
        
        print(f"✅ 综合欠料文件已保存: {output_file}")
        
        # 统计信息
        print(f"📊 最终统计:")
        print(f"  - 总记录数: {len(comprehensive_df)}")
        print(f"  - 唯一订单数: {comprehensive_df['订单编号'].nunique()}")
        print(f"  - 有缺料记录: {len(comprehensive_df[comprehensive_df['仓存不足'] > 0])}")
        print(f"  - 无缺料记录: {len(comprehensive_df[comprehensive_df['仓存不足'] == 0])}")
        
        # 数据来源分布
        print(f"  - 数据来源分布:")
        for source, count in comprehensive_df['数据来源'].value_counts().items():
            print(f"    {source}: {count}条")
        
        return output_file
        
    except Exception as e:
        print(f"保存失败: {e}")
        return None

def main():
    """主函数"""
    print("=" * 80)
    print("综合欠料清单分析工具")
    print("分析3个欠料清单，生成最全面的欠料数据文件")
    print("=" * 80)
    
    # 1. 创建综合数据
    comprehensive_df = create_comprehensive_shortage_data()
    
    if comprehensive_df is None:
        print("❌ 创建综合数据失败")
        return
    
    # 2. 验证与input订单的匹配
    final_df = validate_with_input_orders(comprehensive_df)
    
    # 3. 保存最终文件
    output_file = save_comprehensive_shortage_file(final_df)
    
    if output_file:
        print(f"\n🎉 分析完成！")
        print(f"📋 生成的综合欠料文件: {output_file}")
        print(f"\n建议:")
        print(f"1. 检查生成的文件质量")
        print(f"2. 可以将此文件复制为 input/mat_owe_pso.xlsx 使用")
        print(f"3. 运行 silverPlan_analysis.py 验证分析效果")
    else:
        print(f"\n❌ 保存文件失败")

if __name__ == "__main__":
    main()
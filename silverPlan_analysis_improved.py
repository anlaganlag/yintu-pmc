#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银图PMC综合物料分析系统 - 改进版
使用健壮的文件加载机制
支持配置化管理和错误恢复
"""

import pandas as pd
import numpy as np
from datetime import datetime
import warnings
import sys
from pathlib import Path
from typing import Optional, Dict, List

# 导入健壮的文件加载器
from robust_file_loader import RobustExcelLoader
from file_config import FileConfig

warnings.filterwarnings('ignore')

class ImprovedComprehensivePMCAnalyzer:
    """改进的PMC综合分析器"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        初始化分析器
        
        Args:
            config_file: 配置文件路径
        """
        self.config = FileConfig(config_file)
        self.loader = RobustExcelLoader(self.config)
        
        # 数据容器
        self.orders_df = None
        self.shortage_df = None
        self.inventory_df = None
        self.supplier_df = None
        self.final_result = None
        
        # 汇率设置
        self.currency_rates = {
            'RMB': 1.0,
            'USD': 7.20,
            'HKD': 0.93,
            'EUR': 7.85
        }
        
        # 错误和警告收集
        self.errors = []
        self.warnings = []
        
    def load_all_data(self) -> bool:
        """
        加载所有数据源（使用健壮的加载器）
        
        Returns:
            成功返回True，失败返回False
        """
        print("=== 🔄 使用健壮加载器加载数据 ===")
        
        try:
            # 使用健壮加载器加载所有文件
            data = self.loader.load_all_files()
            
            # 分配数据到相应的属性
            if 'orders' in data:
                self.orders_df = self._standardize_order_columns(data['orders'])
                print(f"✅ 订单数据加载成功: {len(self.orders_df)} 条记录")
            else:
                self.errors.append("订单数据加载失败")
                print("❌ 订单数据加载失败")
                return False
                
            if 'shortage' in data:
                self.shortage_df = self._standardize_shortage_columns(data['shortage'])
                print(f"✅ 欠料数据加载成功: {len(self.shortage_df)} 条记录")
            else:
                self.warnings.append("欠料数据未加载，将使用默认值")
                print("⚠️ 欠料数据未加载，将使用默认值")
                
            if 'inventory' in data:
                self.inventory_df = self._standardize_inventory_columns(data['inventory'])
                print(f"✅ 库存数据加载成功: {len(self.inventory_df)} 条记录")
            else:
                self.warnings.append("库存数据未加载，价格信息可能不完整")
                print("⚠️ 库存数据未加载，价格信息可能不完整")
                
            if 'supplier' in data:
                self.supplier_df = self._standardize_supplier_columns(data['supplier'])
                print(f"✅ 供应商数据加载成功: {len(self.supplier_df)} 条记录")
            else:
                self.warnings.append("供应商数据未加载，供应商信息可能不完整")
                print("⚠️ 供应商数据未加载，供应商信息可能不完整")
                
            # 打印加载摘要
            print("\n" + self.loader.get_load_summary())
            
            return True
            
        except Exception as e:
            self.errors.append(f"数据加载失败: {str(e)}")
            print(f"❌ 数据加载失败: {e}")
            return False
            
    def _standardize_order_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """标准化订单表列名"""
        column_mapping = {
            '生 產 單 号(  廠方 )': '生产单号',
            '生 產 單 号( 廠方 )': '生产单号',  # Different spacing
            '生産単号': '生产单号',
            '生产单': '生产单号',
            '生 產 單 号(客方 )': '客户订单号',
            '客户订单': '客户订单号',
            '型 號( 廠方/客方 )': '产品型号',
            '型号': '产品型号',
            '產品型號': '产品型号',
            '數 量  (Pcs)': '数量Pcs',
            '数量': '数量Pcs',
            'BOM NO.': 'BOM编号',
            'BOM编号': 'BOM编号',
            '客期': '客户交期',
            '交期': '客户交期'
        }
        
        # 智能列名映射
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
                
        # 确保关键列存在
        if '订单金额' not in df.columns:
            df['订单金额'] = 1000  # 默认值
            self.warnings.append("订单金额字段缺失，使用默认值1000 USD")
            
        # 确保生产单号为字符串类型
        if '生产单号' in df.columns:
            df['生产单号'] = df['生产单号'].astype(str)
            
        return df
        
    def _standardize_shortage_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """标准化欠料表列名"""
        column_mapping = {
            '生産単号': '生产单号',
            '生产单': '生产单号',
            '訂單編號\n(已勾測料)': '生产单号',  # With newline
            '訂單編號 (已勾測料)': '生产单号',  # With space
            'PSO': '生产单号',
            '物料编号': '物料编码',
            '物料代码': '物料编码',
            '物料編號': '物料编码',  # Actual column name
            '物料': '物料编码',
            '欠料数': '欠料数量',
            '欠数': '欠料数量',
            '欠料': '欠料数量',
            '倉存不足\n(齊套料)': '欠料数量',  # With newline
            '倉存不足 (齊套料)': '欠料数量',  # With space
            '物料描述': '物料名称',
            '描述': '物料名称',
            '物項名称': '物料名称'  # Actual column name
        }
        
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
                
        # 数据类型转换
        if '欠料数量' in df.columns:
            df['欠料数量'] = pd.to_numeric(df['欠料数量'], errors='coerce').fillna(0)
            
        return df
        
    def _standardize_inventory_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """标准化库存表列名"""
        column_mapping = {
            '物料编号': '物料编码',
            '物料代码': '物料编码',
            '物項編號': '物料编码',  # Traditional Chinese
            '价格': '单价',
            '库存单价': '单价',
            '成本': '单价',
            '成本單價': '单价',  # Traditional Chinese
            '單價': '单价'  # Traditional Chinese
        }
        
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
                
        # 数据类型转换
        if '单价' in df.columns:
            df['单价'] = pd.to_numeric(df['单价'], errors='coerce').fillna(0)
            
        return df
        
    def _standardize_supplier_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """标准化供应商表列名"""
        column_mapping = {
            '物料编号': '物料编码',
            '物料代码': '物料编码',
            '物项编号': '物料编码',  # Actual column name in supplier file
            '物項編號': '物料编码',  # Traditional Chinese
            '供应商名称': '供应商',
            '厂商': '供应商',
            '价格': '单价',
            '采购单价': '单价',
            '修改日期': '最后修改日期',
            '更新日期': '最后修改日期'
        }
        
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
                
        # 数据类型转换
        if '单价' in df.columns:
            df['单价'] = pd.to_numeric(df['单价'], errors='coerce').fillna(0)
            
        if '最后修改日期' in df.columns:
            df['最后修改日期'] = pd.to_datetime(df['最后修改日期'], errors='coerce')
            
        return df
        
    def analyze(self) -> bool:
        """
        执行完整的PMC分析
        
        Returns:
            成功返回True，失败返回False
        """
        print("\n=== 📊 开始PMC综合分析 ===")
        
        try:
            # 1. 加载数据
            if not self.load_all_data():
                return False
                
            # 2. 执行LEFT JOIN分析
            print("\n=== 🔗 执行LEFT JOIN连接 ===")
            self.final_result = self._perform_left_join_analysis()
            
            # 3. 计算ROI
            print("\n=== 💹 计算ROI指标 ===")
            self._calculate_roi()
            
            # 4. 选择最优供应商
            print("\n=== 🏆 选择最优供应商 ===")
            self._select_best_suppliers()
            
            # 5. 生成统计摘要
            print("\n=== 📈 生成统计摘要 ===")
            self._generate_summary()
            
            return True
            
        except Exception as e:
            self.errors.append(f"分析过程失败: {str(e)}")
            print(f"❌ 分析失败: {e}")
            return False
            
    def _perform_left_join_analysis(self) -> pd.DataFrame:
        """执行LEFT JOIN分析"""
        # 以订单为主表
        result = self.orders_df.copy()
        
        # 调试信息
        print(f"订单数据列: {result.columns.tolist()[:10]}")
        print(f"订单数据有'生产单号'列: {'生产单号' in result.columns}")
        
        # LEFT JOIN 欠料数据
        if self.shortage_df is not None:
            print(f"欠料数据列: {self.shortage_df.columns.tolist()[:10]}")
            print(f"欠料数据有'生产单号'列: {'生产单号' in self.shortage_df.columns}")
            
            # 检查两个DataFrame都有'生产单号'列
            if '生产单号' not in result.columns:
                print("❌ 订单数据缺少'生产单号'列")
                raise KeyError("订单数据缺少'生产单号'列")
            if '生产单号' not in self.shortage_df.columns:
                print("❌ 欠料数据缺少'生产单号'列")
                raise KeyError("欠料数据缺少'生产单号'列")
                
            result = pd.merge(
                result,
                self.shortage_df,
                on='生产单号',
                how='left',
                suffixes=('', '_shortage')
            )
        else:
            # 如果没有欠料数据，添加默认列
            result['物料编码'] = ''
            result['物料名称'] = ''
            result['欠料数量'] = 0
            
        # LEFT JOIN 库存价格
        if self.inventory_df is not None:
            result = pd.merge(
                result,
                self.inventory_df[['物料编码', '单价']].rename(columns={'单价': '库存单价'}),
                on='物料编码',
                how='left'
            )
        else:
            result['库存单价'] = 0
            
        # LEFT JOIN 供应商数据
        if self.supplier_df is not None:
            result = pd.merge(
                result,
                self.supplier_df,
                on='物料编码',
                how='left',
                suffixes=('', '_supplier')
            )
        else:
            result['供应商'] = ''
            result['单价'] = 0
            
        return result
        
    def _calculate_roi(self):
        """计算ROI"""
        if self.final_result is None:
            return
            
        # 计算欠料金额
        self.final_result['欠料金额'] = (
            self.final_result['欠料数量'] * 
            self.final_result['单价'].fillna(self.final_result['库存单价']).fillna(0)
        )
        
        # 计算ROI
        def calc_roi(row):
            if pd.isna(row['欠料金额']) or row['欠料金额'] == 0:
                return '无需投入'
            else:
                roi_value = row['订单金额'] / row['欠料金额']
                return f"{roi_value:.2f}"
                
        self.final_result['ROI'] = self.final_result.apply(calc_roi, axis=1)
        
    def _select_best_suppliers(self):
        """选择最优供应商"""
        if self.supplier_df is None or self.final_result is None:
            return
            
        # 实现供应商选择逻辑
        # 这里可以根据多个因素选择最优供应商
        pass
        
    def _generate_summary(self):
        """生成统计摘要"""
        if self.final_result is None:
            return
            
        print(f"订单总数: {len(self.final_result)}")
        print(f"涉及物料种类: {self.final_result['物料编码'].nunique()}")
        print(f"总欠料金额: {self.final_result['欠料金额'].sum():,.2f}")
        
        # 打印错误和警告
        if self.errors:
            print("\n❌ 错误信息:")
            for error in self.errors:
                print(f"  - {error}")
                
        if self.warnings:
            print("\n⚠️ 警告信息:")
            for warning in self.warnings:
                print(f"  - {warning}")
                
    def save_results(self, output_file: Optional[str] = None):
        """
        保存分析结果
        
        Args:
            output_file: 输出文件名
        """
        if self.final_result is None:
            print("❌ 没有分析结果可保存")
            return
            
        if output_file is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = f'银图PMC综合物料分析报告_改进版_{timestamp}.xlsx'
            
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 主表
                self.final_result.to_excel(
                    writer, 
                    sheet_name='综合物料分析明细',
                    index=False
                )
                
                # 汇总统计
                summary_df = self._create_summary_dataframe()
                if summary_df is not None:
                    summary_df.to_excel(
                        writer,
                        sheet_name='汇总统计',
                        index=False
                    )
                    
            print(f"✅ 分析结果已保存到: {output_file}")
            
        except Exception as e:
            print(f"❌ 保存失败: {e}")
            
    def _create_summary_dataframe(self) -> Optional[pd.DataFrame]:
        """创建汇总统计DataFrame"""
        if self.final_result is None:
            return None
            
        summary_data = {
            '统计项': [
                '订单总数',
                '物料种类',
                '总欠料金额',
                '平均ROI',
                '数据完整性'
            ],
            '数值': [
                len(self.final_result),
                self.final_result['物料编码'].nunique(),
                f"{self.final_result['欠料金额'].sum():,.2f}",
                self._calculate_average_roi(),
                f"{self._calculate_data_completeness():.1%}"
            ]
        }
        
        return pd.DataFrame(summary_data)
        
    def _calculate_average_roi(self) -> str:
        """计算平均ROI"""
        if self.final_result is None:
            return 'N/A'
            
        numeric_rois = []
        for roi in self.final_result['ROI']:
            if roi != '无需投入':
                try:
                    numeric_rois.append(float(roi))
                except:
                    pass
                    
        if numeric_rois:
            return f"{np.mean(numeric_rois):.2f}"
        return 'N/A'
        
    def _calculate_data_completeness(self) -> float:
        """计算数据完整性"""
        if self.final_result is None:
            return 0.0
            
        total_fields = len(self.final_result.columns) * len(self.final_result)
        non_null_fields = self.final_result.count().sum()
        
        return non_null_fields / total_fields if total_fields > 0 else 0.0


def main():
    """主函数"""
    print("=" * 60)
    print("银图PMC综合物料分析系统 - 改进版")
    print("支持健壮的文件加载和错误恢复")
    print("=" * 60)
    
    analyzer = ImprovedComprehensivePMCAnalyzer()
    
    if analyzer.analyze():
        analyzer.save_results()
        print("\n✅ 分析完成！")
    else:
        print("\n❌ 分析失败，请检查错误信息")
        

if __name__ == "__main__":
    main()
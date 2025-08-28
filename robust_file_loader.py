#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
健壮的Excel文件加载器
支持多种编码、错误恢复、数据验证
"""

import pandas as pd
import numpy as np
from pathlib import Path
from typing import Optional, List, Dict, Any, Union
import logging
import chardet
import warnings
from file_config import FileConfig

warnings.filterwarnings('ignore')

class RobustExcelLoader:
    """健壮的Excel文件加载器"""
    
    def __init__(self, config: Optional[FileConfig] = None):
        """
        初始化加载器
        
        Args:
            config: 文件配置管理器实例
        """
        self.config = config or FileConfig()
        self.logger = logging.getLogger(__name__)
        self.loaded_data = {}
        self.load_errors = []
        
    def load_excel_with_fallback(
        self, 
        file_path: Union[str, Path],
        sheet_names: Optional[List[str]] = None,
        required_columns: Optional[List[str]] = None,
        dtype_mapping: Optional[Dict[str, type]] = None
    ) -> Optional[pd.DataFrame]:
        """
        使用多种策略加载Excel文件
        
        Args:
            file_path: 文件路径
            sheet_names: 工作表名称列表（按优先级排序）
            required_columns: 必需的列名列表
            dtype_mapping: 列数据类型映射
            
        Returns:
            成功加载的DataFrame，或None
        """
        file_path = Path(file_path)
        
        if not file_path.exists():
            self.logger.error(f"文件不存在: {file_path}")
            self.load_errors.append(f"文件不存在: {file_path}")
            return None
            
        # 尝试不同的加载策略
        strategies = [
            self._load_with_openpyxl,
            self._load_with_xlrd,
            self._load_with_default,
            self._load_as_csv
        ]
        
        for strategy in strategies:
            try:
                df = strategy(file_path, sheet_names)
                if df is not None:
                    # 验证和处理数据
                    df = self._validate_and_clean_data(
                        df, 
                        required_columns, 
                        dtype_mapping
                    )
                    if df is not None:
                        self.logger.info(f"成功加载文件: {file_path}")
                        return df
            except Exception as e:
                self.logger.warning(f"策略 {strategy.__name__} 失败: {e}")
                continue
                
        self.logger.error(f"所有加载策略都失败: {file_path}")
        self.load_errors.append(f"无法加载文件: {file_path}")
        return None
        
    def _load_with_openpyxl(
        self, 
        file_path: Path, 
        sheet_names: Optional[List[str]] = None
    ) -> Optional[pd.DataFrame]:
        """使用openpyxl引擎加载"""
        try:
            # 尝试不同的header行
            header_rows = [None, 0, 1, 2]
            
            for header_row in header_rows:
                try:
                    if sheet_names:
                        for sheet in sheet_names:
                            try:
                                if header_row is None:
                                    df = pd.read_excel(
                                        file_path, 
                                        sheet_name=sheet,
                                        engine='openpyxl'
                                    )
                                else:
                                    df = pd.read_excel(
                                        file_path, 
                                        sheet_name=sheet,
                                        engine='openpyxl',
                                        header=header_row
                                    )
                                    
                                # 检查DataFrame是否有效
                                if isinstance(df, pd.DataFrame) and not df.empty and len(df.columns) > 1:
                                    # 跳过列名都是Unnamed的情况
                                    unnamed_cols = [col for col in df.columns if str(col).startswith('Unnamed')]
                                    if len(unnamed_cols) < len(df.columns) / 2:
                                        self.logger.info(f"使用openpyxl成功加载工作表: {sheet}, header={header_row}")
                                        return df
                            except:
                                continue
                    else:
                        if header_row is None:
                            df = pd.read_excel(file_path, engine='openpyxl')
                        else:
                            df = pd.read_excel(file_path, engine='openpyxl', header=header_row)
                            
                        # 检查DataFrame是否有效
                        if isinstance(df, pd.DataFrame) and not df.empty and len(df.columns) > 1:
                            # 跳过列名都是Unnamed的情况
                            unnamed_cols = [col for col in df.columns if str(col).startswith('Unnamed')]
                            if len(unnamed_cols) < len(df.columns) / 2:
                                self.logger.info(f"使用openpyxl成功加载，header={header_row}")
                                return df
                except:
                    continue
                    
        except Exception as e:
            self.logger.warning(f"openpyxl加载失败: {e}")
            return None
            
    def _load_with_xlrd(
        self, 
        file_path: Path, 
        sheet_names: Optional[List[str]] = None
    ) -> Optional[pd.DataFrame]:
        """使用xlrd引擎加载（用于旧版Excel）"""
        try:
            if sheet_names:
                for sheet in sheet_names:
                    try:
                        df = pd.read_excel(
                            file_path, 
                            sheet_name=sheet,
                            engine='xlrd'
                        )
                        self.logger.info(f"使用xlrd成功加载工作表: {sheet}")
                        return df
                    except:
                        continue
            else:
                return pd.read_excel(file_path, engine='xlrd')
        except Exception as e:
            self.logger.warning(f"xlrd加载失败: {e}")
            return None
            
    def _load_with_default(
        self, 
        file_path: Path, 
        sheet_names: Optional[List[str]] = None
    ) -> Optional[pd.DataFrame]:
        """使用默认引擎加载"""
        try:
            if sheet_names:
                for sheet in sheet_names:
                    try:
                        # 尝试不同的工作表名称变体
                        sheet_variants = [
                            sheet,
                            sheet.strip(),
                            sheet.replace(' ', ''),
                            sheet.replace('-', ''),
                            sheet.replace('_', ''),
                        ]
                        
                        for variant in sheet_variants:
                            try:
                                df = pd.read_excel(file_path, sheet_name=variant)
                                self.logger.info(f"成功加载工作表: {variant}")
                                return df
                            except:
                                continue
                    except:
                        continue
                        
            # 尝试加载第一个工作表
            return pd.read_excel(file_path, sheet_name=0)
            
        except Exception as e:
            self.logger.warning(f"默认加载失败: {e}")
            return None
            
    def _load_as_csv(
        self, 
        file_path: Path, 
        sheet_names: Optional[List[str]] = None
    ) -> Optional[pd.DataFrame]:
        """尝试作为CSV加载（某些Excel文件可能是CSV格式）"""
        try:
            # 检测文件编码
            with open(file_path, 'rb') as f:
                raw_data = f.read(10000)
                result = chardet.detect(raw_data)
                encoding = result['encoding'] or 'utf-8'
                
            # 尝试不同的分隔符
            separators = [',', '\t', ';', '|']
            for sep in separators:
                try:
                    df = pd.read_csv(
                        file_path, 
                        encoding=encoding,
                        sep=sep,
                        error_bad_lines=False,
                        warn_bad_lines=False
                    )
                    if len(df.columns) > 1:  # 确保有多列
                        self.logger.info(f"作为CSV成功加载，编码: {encoding}, 分隔符: {sep}")
                        return df
                except:
                    continue
                    
        except Exception as e:
            self.logger.warning(f"CSV加载失败: {e}")
            return None
            
    def _validate_and_clean_data(
        self,
        df: pd.DataFrame,
        required_columns: Optional[List[str]] = None,
        dtype_mapping: Optional[Dict[str, type]] = None
    ) -> Optional[pd.DataFrame]:
        """
        验证和清理数据
        
        Args:
            df: 原始DataFrame
            required_columns: 必需的列名列表
            dtype_mapping: 列数据类型映射
            
        Returns:
            清理后的DataFrame，或None
        """
        if df is None or df.empty:
            self.logger.warning("DataFrame为空")
            return None
            
        # 删除完全空的行和列
        df = df.dropna(how='all', axis=0)
        df = df.dropna(how='all', axis=1)
        
        # 标准化列名（去除空格和特殊字符）
        df.columns = df.columns.str.strip()
        df.columns = df.columns.str.replace('\s+', ' ', regex=True)
        
        # 检查必需列
        if required_columns:
            missing_cols = set(required_columns) - set(df.columns)
            if missing_cols:
                # 尝试模糊匹配列名
                for missing_col in missing_cols:
                    found = False
                    for col in df.columns:
                        if self._fuzzy_match_column(missing_col, col):
                            df = df.rename(columns={col: missing_col})
                            found = True
                            break
                    if not found:
                        self.logger.warning(f"缺少必需列: {missing_col}")
                        
        # 转换数据类型
        if dtype_mapping:
            for col, dtype in dtype_mapping.items():
                if col in df.columns:
                    try:
                        if dtype == str:
                            df[col] = df[col].astype(str)
                        elif dtype in [int, float]:
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                        elif dtype == pd.Timestamp:
                            df[col] = pd.to_datetime(df[col], errors='coerce')
                    except Exception as e:
                        self.logger.warning(f"转换列 {col} 到 {dtype} 失败: {e}")
                        
        # 处理重复行
        duplicate_count = df.duplicated().sum()
        if duplicate_count > 0:
            self.logger.warning(f"发现 {duplicate_count} 个重复行，已删除")
            df = df.drop_duplicates()
            
        return df
        
    def _fuzzy_match_column(self, target: str, candidate: str) -> bool:
        """
        模糊匹配列名
        
        Args:
            target: 目标列名
            candidate: 候选列名
            
        Returns:
            是否匹配
        """
        # 转换为小写并去除空格
        target = target.lower().replace(' ', '').replace('_', '')
        candidate = candidate.lower().replace(' ', '').replace('_', '')
        
        # 完全匹配
        if target == candidate:
            return True
            
        # 包含关系
        if target in candidate or candidate in target:
            return True
            
        # 常见的列名映射
        mappings = {
            '生产单号': ['生產單号', '生产单', 'productionorder', 'pso'],
            '客户订单号': ['客户订单', '客方订单', 'customerorder'],
            '产品型号': ['型号', '产品', 'model', 'productmodel'],
            '数量': ['数量pcs', '數量', 'quantity', 'qty'],
            '物料编码': ['物料编号', '物料代码', 'materialcode'],
            '物料名称': ['物料描述', '物料', 'materialname'],
            '供应商': ['供应商名称', 'supplier', 'vendor']
        }
        
        for key, values in mappings.items():
            if target in key.lower() or key.lower() in target:
                for value in values:
                    if value.lower() in candidate or candidate in value.lower():
                        return True
                        
        return False
        
    def load_all_files(self) -> Dict[str, pd.DataFrame]:
        """
        加载所有配置的文件
        
        Returns:
            加载的数据字典
        """
        results = {}
        
        # 加载订单数据
        self._load_order_files(results)
        
        # 加载其他文件
        file_configs = {
            'shortage': {
                'key': 'shortage',
                'required_columns': None  # Skip validation, will be handled by standardization
            },
            'inventory': {
                'key': 'inventory',
                'required_columns': None  # Skip validation, will be handled by standardization
            },
            'supplier': {
                'key': 'supplier',
                'required_columns': None  # Skip validation, will be handled by standardization
            }
        }
        
        for name, config in file_configs.items():
            file_path = self.config.find_file(config['key'])
            if file_path:
                df = self.load_excel_with_fallback(
                    file_path,
                    sheet_names=self.config.config['file_paths'][config['key']].get('sheets'),
                    required_columns=config.get('required_columns')
                )
                if df is not None:
                    results[name] = df
                else:
                    self.logger.warning(f"无法加载 {name} 文件")
                    
        return results
        
    def _load_order_files(self, results: Dict[str, pd.DataFrame]):
        """加载订单文件"""
        order_dfs = []
        
        # 国内订单
        domestic_path = self.config.find_file('orders_domestic')
        if domestic_path:
            for sheet, month in [('8月', '8月'), ('9月', '9月')]:
                df = self.load_excel_with_fallback(
                    domestic_path,
                    sheet_names=[sheet],
                    required_columns=None  # Skip validation, columns will be standardized
                )
                if df is not None:
                    df['月份'] = month
                    df['数据来源工作表'] = '国内'
                    order_dfs.append(df)
                    
        # 柬埔寨订单
        cambodia_path = self.config.find_file('orders_cambodia')
        if cambodia_path:
            for sheet, month in [('8月 -柬', '8月'), ('9月 -柬', '9月')]:
                df = self.load_excel_with_fallback(
                    cambodia_path,
                    sheet_names=[sheet, sheet.replace(' ', ''), f'{month}-柬'],
                    required_columns=None  # Skip validation, columns will be standardized
                )
                if df is not None:
                    df['月份'] = month
                    df['数据来源工作表'] = '柬埔寨'
                    order_dfs.append(df)
                    
        if order_dfs:
            results['orders'] = pd.concat(order_dfs, ignore_index=True)
            self.logger.info(f"成功加载 {len(results['orders'])} 条订单数据")
        else:
            self.logger.error("未能加载任何订单数据")
            
    def get_load_summary(self) -> str:
        """
        获取加载摘要
        
        Returns:
            加载摘要字符串
        """
        summary = []
        summary.append("=== 文件加载摘要 ===")
        
        if self.loaded_data:
            summary.append("\n✅ 成功加载的文件:")
            for name, df in self.loaded_data.items():
                summary.append(f"  - {name}: {len(df)} 行, {len(df.columns)} 列")
                
        if self.load_errors:
            summary.append("\n❌ 加载错误:")
            for error in self.load_errors:
                summary.append(f"  - {error}")
                
        return "\n".join(summary)
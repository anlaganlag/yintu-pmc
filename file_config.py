#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件配置管理模块
提供健壮的文件路径管理和加载功能
"""

import os
import json
from pathlib import Path
from typing import Dict, List, Optional, Union
import logging

class FileConfig:
    """文件配置管理器"""
    
    def __init__(self, config_file: str = "config.json"):
        """
        初始化文件配置管理器
        
        Args:
            config_file: 配置文件路径
        """
        self.config_file = config_file
        self.base_dir = Path.cwd()
        self._setup_logging()
        self.config = self._load_config()
        
    def _setup_logging(self):
        """设置日志系统"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('file_loading.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
    def _load_config(self) -> Dict:
        """加载配置文件"""
        default_config = {
            "file_paths": {
                "orders_domestic": {
                    "primary": "order-amt-89.xlsx",
                    "alternatives": [
                        "订单-国内-89.xlsx",
                        "order_domestic_89.xlsx",
                        "data/order-amt-89.xlsx"
                    ],
                    "sheets": ["8月", "9月"]
                },
                "orders_cambodia": {
                    "primary": "order-amt-89-c.xlsx",
                    "alternatives": [
                        "订单-柬埔寨-89.xlsx",
                        "order_cambodia_89.xlsx",
                        "data/order-amt-89-c.xlsx"
                    ],
                    "sheets": ["8月 -柬", "9月 -柬", "8月-柬", "9月-柬"]
                },
                "shortage": {
                    "primary": "mat_owe_pso.xlsx",
                    "alternatives": [
                        "D:/yingtu-PMC/mat_owe_pso.xlsx",
                        "D:/827-yintu/yintu-pmc/mat_owe_pso.xlsx",
                        "data/mat_owe_pso.xlsx",
                        "欠料清单.xlsx",
                        "material_shortage.xlsx"
                    ],
                    "sheets": [None, "Sheet1", "欠料表"]
                },
                "inventory": {
                    "primary": "inventory_list.xlsx",
                    "alternatives": [
                        "D:/yingtu-PMC/inventory_list.xlsx",
                        "D:/827-yintu/yintu-pmc/inventory_list.xlsx",
                        "data/inventory_list.xlsx",
                        "库存清单.xlsx",
                        "stock_list.xlsx"
                    ],
                    "sheets": [None, "Sheet1", "库存表"]
                },
                "supplier": {
                    "primary": "supplier.xlsx",
                    "alternatives": [
                        "D:/yingtu-PMC/supplier.xlsx",
                        "D:/827-yintu/yintu-pmc/supplier.xlsx",
                        "data/supplier.xlsx",
                        "供应商清单.xlsx",
                        "supplier_list.xlsx"
                    ],
                    "sheets": [None, "Sheet1", "供应商表"]
                }
            },
            "search_paths": [
                ".",
                "./data",
                "./input",
                "../",
                "D:/yingtu-PMC",
                "D:/827-yintu/yintu-pmc",
                os.path.expanduser("~/Desktop"),
                os.path.expanduser("~/Downloads")
            ],
            "encoding_options": ["utf-8", "gbk", "gb2312", "gb18030"],
            "retry_times": 3,
            "enable_user_input": True
        }
        
        # 尝试加载已存在的配置文件
        config_path = Path(self.config_file)
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # 合并配置，保留用户自定义的设置
                    default_config.update(loaded_config)
                    self.logger.info(f"成功加载配置文件: {self.config_file}")
            except Exception as e:
                self.logger.warning(f"加载配置文件失败，使用默认配置: {e}")
        else:
            # 保存默认配置
            self._save_config(default_config)
            
        return default_config
    
    def _save_config(self, config: Dict):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.logger.info(f"配置已保存到: {self.config_file}")
        except Exception as e:
            self.logger.error(f"保存配置失败: {e}")
            
    def find_file(self, file_key: str, required: bool = True) -> Optional[Path]:
        """
        查找文件的实际路径
        
        Args:
            file_key: 文件配置键名
            required: 是否必需文件
            
        Returns:
            找到的文件路径，或None
        """
        if file_key not in self.config["file_paths"]:
            self.logger.error(f"未知的文件配置键: {file_key}")
            return None
            
        file_config = self.config["file_paths"][file_key]
        primary = file_config["primary"]
        alternatives = file_config.get("alternatives", [])
        
        # 所有可能的文件名
        all_names = [primary] + alternatives
        
        # 在所有搜索路径中查找
        for search_path in self.config["search_paths"]:
            search_dir = Path(search_path).expanduser().resolve()
            if not search_dir.exists():
                continue
                
            for filename in all_names:
                file_path = search_dir / filename
                
                # 处理绝对路径
                if os.path.isabs(filename):
                    file_path = Path(filename)
                    
                if file_path.exists():
                    self.logger.info(f"找到文件 {file_key}: {file_path}")
                    # 更新配置，下次优先使用找到的路径
                    self.config["file_paths"][file_key]["primary"] = str(file_path)
                    self._save_config(self.config)
                    return file_path
                    
        # 文件未找到
        if required:
            self.logger.error(f"必需文件 {file_key} 未找到: {primary}")
            if self.config.get("enable_user_input", True):
                return self._request_user_input(file_key)
        else:
            self.logger.warning(f"可选文件 {file_key} 未找到: {primary}")
            
        return None
    
    def _request_user_input(self, file_key: str) -> Optional[Path]:
        """
        请求用户输入文件路径
        
        Args:
            file_key: 文件配置键名
            
        Returns:
            用户指定的文件路径，或None
        """
        try:
            print(f"\n⚠️ 未找到文件: {self.config['file_paths'][file_key]['primary']}")
            print("请输入文件的完整路径（或输入 'skip' 跳过）：")
            user_input = input().strip()
            
            if user_input.lower() == 'skip':
                return None
                
            file_path = Path(user_input)
            if file_path.exists():
                # 更新配置
                self.config["file_paths"][file_key]["primary"] = str(file_path)
                self._save_config(self.config)
                self.logger.info(f"用户指定文件路径: {file_path}")
                return file_path
            else:
                print(f"❌ 文件不存在: {user_input}")
                return None
                
        except Exception as e:
            self.logger.error(f"获取用户输入失败: {e}")
            return None
            
    def get_all_file_paths(self) -> Dict[str, Optional[Path]]:
        """
        获取所有配置的文件路径
        
        Returns:
            文件路径字典
        """
        paths = {}
        for key in self.config["file_paths"].keys():
            paths[key] = self.find_file(key, required=False)
        return paths
        
    def validate_files(self) -> bool:
        """
        验证所有必需文件是否存在
        
        Returns:
            所有必需文件都存在返回True
        """
        required_files = ["orders_domestic", "orders_cambodia", "shortage"]
        all_found = True
        
        for file_key in required_files:
            if not self.find_file(file_key, required=True):
                all_found = False
                
        return all_found
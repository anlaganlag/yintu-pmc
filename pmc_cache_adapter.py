#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PMC系统专用缓存适配器
为PMC分析系统提供定制化的缓存功能
"""

import pandas as pd
import hashlib
import time
from datetime import datetime, timedelta
from typing import Optional, Dict, Any
from cache_manager import DatabaseCacheManager, cached

class PMCCacheAdapter:
    """PMC系统缓存适配器"""
    
    def __init__(self):
        """初始化PMC缓存适配器"""
        self.cache = DatabaseCacheManager("cache/pmc_analysis_cache.db")
        self._cache_config = {
            'analysis_report': {'ttl': 3600, 'tags': 'analysis,report'},      # 分析报告: 1小时
            'filtered_data': {'ttl': 1800, 'tags': 'filter,data'},            # 筛选数据: 30分钟
            'roi_calculation': {'ttl': 2700, 'tags': 'calculation,roi'},       # ROI计算: 45分钟
            'summary_stats': {'ttl': 900, 'tags': 'stats,summary'},            # 统计信息: 15分钟
            'supplier_data': {'ttl': 7200, 'tags': 'supplier,reference'},      # 供应商数据: 2小时
            'order_data': {'ttl': 7200, 'tags': 'order,reference'}             # 订单数据: 2小时
        }
        print("✅ PMC缓存适配器已初始化")
    
    def get_analysis_report(self, report_file: str = None) -> Optional[Dict[str, pd.DataFrame]]:
        """获取缓存的分析报告"""
        if report_file:
            key = f"analysis_report:{hashlib.md5(report_file.encode()).hexdigest()}"
        else:
            key = "analysis_report:latest"
        
        return self.cache.get(key)
    
    def set_analysis_report(self, data: Dict[str, pd.DataFrame], report_file: str = None) -> bool:
        """缓存分析报告"""
        if report_file:
            key = f"analysis_report:{hashlib.md5(report_file.encode()).hexdigest()}"
        else:
            key = "analysis_report:latest"
        
        config = self._cache_config['analysis_report']
        return self.cache.set(key, data, ttl=config['ttl'], tags=config['tags'])
    
    def get_filtered_data(self, filters: Dict[str, Any]) -> Optional[pd.DataFrame]:
        """获取缓存的筛选数据"""
        filter_hash = hashlib.md5(str(sorted(filters.items())).encode()).hexdigest()
        key = f"filtered_data:{filter_hash}"
        return self.cache.get(key)
    
    def set_filtered_data(self, data: pd.DataFrame, filters: Dict[str, Any]) -> bool:
        """缓存筛选数据"""
        filter_hash = hashlib.md5(str(sorted(filters.items())).encode()).hexdigest()
        key = f"filtered_data:{filter_hash}"
        
        config = self._cache_config['filtered_data']
        return self.cache.set(key, data, ttl=config['ttl'], tags=config['tags'])
    
    def get_roi_calculation(self, order_ids: list) -> Optional[Dict[str, Any]]:
        """获取缓存的ROI计算结果"""
        order_hash = hashlib.md5(str(sorted(order_ids)).encode()).hexdigest()
        key = f"roi_calculation:{order_hash}"
        return self.cache.get(key)
    
    def set_roi_calculation(self, result: Dict[str, Any], order_ids: list) -> bool:
        """缓存ROI计算结果"""
        order_hash = hashlib.md5(str(sorted(order_ids)).encode()).hexdigest()
        key = f"roi_calculation:{order_hash}"
        
        config = self._cache_config['roi_calculation']
        return self.cache.set(key, result, ttl=config['ttl'], tags=config['tags'])
    
    def get_summary_stats(self, data_version: str) -> Optional[Dict[str, Any]]:
        """获取缓存的统计信息"""
        key = f"summary_stats:{data_version}"
        return self.cache.get(key)
    
    def set_summary_stats(self, stats: Dict[str, Any], data_version: str) -> bool:
        """缓存统计信息"""
        key = f"summary_stats:{data_version}"
        
        config = self._cache_config['summary_stats']
        return self.cache.set(key, stats, ttl=config['ttl'], tags=config['tags'])
    
    def clear_analysis_cache(self) -> int:
        """清除所有分析相关缓存"""
        cleared = 0
        cleared += self.cache.clear_by_tags('analysis')
        cleared += self.cache.clear_by_tags('filter')
        cleared += self.cache.clear_by_tags('calculation')
        return cleared
    
    def clear_reference_cache(self) -> int:
        """清除参考数据缓存（订单、供应商等）"""
        cleared = 0
        cleared += self.cache.clear_by_tags('reference')
        return cleared
    
    def preload_reference_data(self, orders_df: pd.DataFrame, suppliers_df: pd.DataFrame) -> bool:
        """预加载参考数据"""
        try:
            # 缓存订单数据
            order_config = self._cache_config['order_data']
            self.cache.set('reference:orders', orders_df, 
                         ttl=order_config['ttl'], tags=order_config['tags'])
            
            # 缓存供应商数据  
            supplier_config = self._cache_config['supplier_data']
            self.cache.set('reference:suppliers', suppliers_df,
                         ttl=supplier_config['ttl'], tags=supplier_config['tags'])
            
            print(f"✅ 预加载参考数据: 订单{len(orders_df)}行, 供应商{len(suppliers_df)}行")
            return True
            
        except Exception as e:
            print(f"❌ 预加载参考数据失败: {e}")
            return False
    
    def get_cache_health(self) -> Dict[str, Any]:
        """获取缓存健康状态"""
        stats = self.cache.get_stats()
        
        # 计算各类型缓存数量
        all_keys = self.cache.get_keys()
        type_stats = {}
        for key in all_keys:
            cache_type = key.split(':')[0]
            type_stats[cache_type] = type_stats.get(cache_type, 0) + 1
        
        return {
            'overall_stats': stats,
            'cache_types': type_stats,
            'health_score': self._calculate_health_score(stats),
            'recommendations': self._get_recommendations(stats, type_stats)
        }
    
    def _calculate_health_score(self, stats: Dict[str, Any]) -> str:
        """计算缓存健康评分"""
        try:
            hit_rate = float(stats.get('hit_rate', '0%').rstrip('%'))
            total_size_mb = stats.get('total_size_mb', 0)
            expired_keys = stats.get('expired_keys', 0)
            
            score = 100
            
            # 命中率评分 (60分)
            if hit_rate >= 80:
                hit_score = 60
            elif hit_rate >= 60:
                hit_score = 45
            elif hit_rate >= 40:
                hit_score = 30
            else:
                hit_score = 0
            
            # 存储大小评分 (25分)
            if total_size_mb <= 50:  # 50MB以下
                size_score = 25
            elif total_size_mb <= 100:
                size_score = 20
            elif total_size_mb <= 200:
                size_score = 15
            else:
                size_score = 0
            
            # 过期键评分 (15分)
            if expired_keys == 0:
                expired_score = 15
            elif expired_keys <= 5:
                expired_score = 10
            else:
                expired_score = 0
            
            total_score = hit_score + size_score + expired_score
            
            if total_score >= 90:
                return f"优秀 ({total_score}/100)"
            elif total_score >= 70:
                return f"良好 ({total_score}/100)"
            elif total_score >= 50:
                return f"一般 ({total_score}/100)"
            else:
                return f"需要优化 ({total_score}/100)"
                
        except:
            return "无法计算"
    
    def _get_recommendations(self, stats: Dict[str, Any], type_stats: Dict[str, int]) -> list:
        """获取优化建议"""
        recommendations = []
        
        try:
            hit_rate = float(stats.get('hit_rate', '0%').rstrip('%'))
            total_size_mb = stats.get('total_size_mb', 0)
            expired_keys = stats.get('expired_keys', 0)
            
            if hit_rate < 60:
                recommendations.append("⚠️ 命中率较低，建议增加缓存时间或预加载热点数据")
            
            if total_size_mb > 100:
                recommendations.append("💾 缓存占用过大，建议清理过期数据或减少TTL")
            
            if expired_keys > 10:
                recommendations.append("🧹 发现较多过期键，建议执行清理操作")
            
            if len(type_stats.get('filtered_data', 0)) > 20:
                recommendations.append("🔍 筛选缓存过多，可能存在内存泄漏")
            
            if not recommendations:
                recommendations.append("✅ 缓存系统运行良好")
                
        except:
            recommendations.append("❌ 无法生成建议")
        
        return recommendations

# 创建全局PMC缓存适配器实例
pmc_cache = PMCCacheAdapter()

# PMC专用缓存装饰器
def pmc_cached(cache_type: str = 'analysis_report'):
    """PMC专用缓存装饰器"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            # 生成缓存键
            func_name = func.__name__
            args_hash = hashlib.md5(str(args).encode()).hexdigest()[:8]
            kwargs_hash = hashlib.md5(str(sorted(kwargs.items())).encode()).hexdigest()[:8]
            cache_key = f"{cache_type}:{func_name}:{args_hash}:{kwargs_hash}"
            
            # 尝试从缓存获取
            cached_result = pmc_cache.cache.get(cache_key)
            if cached_result is not None:
                return cached_result
            
            # 执行函数并缓存结果
            result = func(*args, **kwargs)
            
            if cache_type in pmc_cache._cache_config:
                config = pmc_cache._cache_config[cache_type]
                pmc_cache.cache.set(cache_key, result, ttl=config['ttl'], tags=config['tags'])
            
            return result
        return wrapper
    return decorator

if __name__ == "__main__":
    # 测试PMC缓存适配器
    print("🧪 测试PMC缓存适配器")
    
    # 测试分析报告缓存
    test_report = {
        '综合物料分析明细': pd.DataFrame({'订单号': ['PSO001', 'PSO002'], '金额': [1000, 2000]}),
        '汇总统计': pd.DataFrame({'项目': ['总金额'], '数值': [3000]})
    }
    
    pmc_cache.set_analysis_report(test_report)
    cached_report = pmc_cache.get_analysis_report()
    print(f"✅ 分析报告缓存测试: {len(cached_report)}个工作表")
    
    # 测试筛选数据缓存
    test_filters = {'月份': '8月', 'ROI': '>2.0'}
    test_filtered = pd.DataFrame({'订单': ['A', 'B'], 'ROI': [2.5, 3.0]})
    
    pmc_cache.set_filtered_data(test_filtered, test_filters)
    cached_filtered = pmc_cache.get_filtered_data(test_filters)
    print(f"✅ 筛选数据缓存测试: {len(cached_filtered)}行")
    
    # 显示缓存健康状态
    health = pmc_cache.get_cache_health()
    print(f"📊 缓存健康状态: {health['health_score']}")
    print(f"💡 建议: {health['recommendations'][0]}")
    
    print("✅ PMC缓存适配器测试完成")
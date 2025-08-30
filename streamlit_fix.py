#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit SessionInfo错误终极修复方案
集成三大解决策略：Session清理、分页加载、配置优化
"""

import streamlit as st
import pandas as pd
import numpy as np
import time
import gc
import sys
from typing import Any, Optional, Dict, List
import functools
import threading
from collections import deque
import weakref

# ========== 方案一：Session State 清理机制 ==========

class SessionStateCleaner:
    """智能Session State清理器"""
    
    def __init__(self, max_items: int = 40, cleanup_interval: int = 100):
        """
        Args:
            max_items: Session中最大保留项目数（Streamlit限制为47）
            cleanup_interval: 每多少次操作后自动清理
        """
        self.max_items = max_items
        self.cleanup_interval = cleanup_interval
        self.operation_count = 0
        self.protected_keys = {'_analyzed_data', '_file_upload_status', '_current_page'}
        self.usage_tracker = {}  # 追踪每个key的使用频率
        
    def track_usage(self, key: str):
        """追踪key使用频率"""
        if key not in self.usage_tracker:
            self.usage_tracker[key] = 0
        self.usage_tracker[key] += 1
        
    def get_state(self, key: str, default=None):
        """安全获取状态值"""
        self.track_usage(key)
        self.operation_count += 1
        
        # 定期清理
        if self.operation_count >= self.cleanup_interval:
            self.cleanup()
            
        try:
            if hasattr(st, 'session_state') and key in st.session_state:
                return st.session_state[key]
        except Exception:
            pass
        return default
        
    def set_state(self, key: str, value: Any) -> bool:
        """安全设置状态值"""
        self.track_usage(key)
        self.operation_count += 1
        
        # 预防性清理
        if self._needs_cleanup():
            self.cleanup()
            
        try:
            if hasattr(st, 'session_state'):
                # 检查是否接近限制
                current_count = len(st.session_state.keys())
                if current_count >= self.max_items and key not in st.session_state:
                    self._force_cleanup()
                    
                st.session_state[key] = value
                return True
        except Exception as e:
            if "index" in str(e).lower() or "setin" in str(e).lower():
                # 遇到索引错误，强制清理
                self._force_cleanup()
                # 重试一次
                try:
                    st.session_state[key] = value
                    return True
                except:
                    pass
        return False
        
    def _needs_cleanup(self) -> bool:
        """检查是否需要清理"""
        try:
            if hasattr(st, 'session_state'):
                return len(st.session_state.keys()) > self.max_items * 0.8
        except:
            pass
        return False
        
    def cleanup(self):
        """智能清理Session State"""
        self.operation_count = 0
        
        try:
            if not hasattr(st, 'session_state'):
                return
                
            all_keys = list(st.session_state.keys())
            
            # 保留受保护的key
            keys_to_remove = []
            for key in all_keys:
                if key not in self.protected_keys:
                    # 基于使用频率决定是否删除
                    usage = self.usage_tracker.get(key, 0)
                    if usage < 5:  # 使用次数少于5次的可以删除
                        keys_to_remove.append(key)
                        
            # 删除最少使用的keys
            keys_to_remove.sort(key=lambda k: self.usage_tracker.get(k, 0))
            remove_count = min(len(keys_to_remove), len(all_keys) - self.max_items + 5)
            
            for key in keys_to_remove[:remove_count]:
                try:
                    del st.session_state[key]
                    if key in self.usage_tracker:
                        del self.usage_tracker[key]
                except:
                    pass
                    
        except Exception:
            pass
            
    def _force_cleanup(self):
        """强制清理（紧急情况）"""
        try:
            if not hasattr(st, 'session_state'):
                return
                
            # 只保留核心keys
            all_keys = list(st.session_state.keys())
            for key in all_keys:
                if key not in self.protected_keys:
                    try:
                        del st.session_state[key]
                    except:
                        pass
                        
            # 清理使用追踪
            self.usage_tracker = {k: v for k, v in self.usage_tracker.items() 
                                 if k in self.protected_keys}
            
            # 强制垃圾回收
            gc.collect()
            
        except Exception:
            pass

# ========== 方案三：分页加载功能 ==========

class DataPaginator:
    """数据分页加载器"""
    
    def __init__(self, page_size: int = 50):
        """
        Args:
            page_size: 每页显示的记录数
        """
        self.page_size = page_size
        self.current_page = 0
        self.total_pages = 0
        self.data_cache = None
        
    def paginate_dataframe(self, df: pd.DataFrame, page_key: str = "page") -> pd.DataFrame:
        """分页显示DataFrame"""
        if df is None or df.empty:
            return df
            
        total_rows = len(df)
        self.total_pages = (total_rows + self.page_size - 1) // self.page_size
        
        # 获取当前页码
        if f"{page_key}_current" in st.session_state:
            self.current_page = st.session_state[f"{page_key}_current"]
        else:
            self.current_page = 0
            
        # 确保页码在有效范围内
        self.current_page = max(0, min(self.current_page, self.total_pages - 1))
        
        # 计算起止索引
        start_idx = self.current_page * self.page_size
        end_idx = min(start_idx + self.page_size, total_rows)
        
        # 创建分页控制器
        col1, col2, col3, col4, col5 = st.columns([1, 1, 2, 1, 1])
        
        with col1:
            if st.button("⏮ 首页", key=f"{page_key}_first", disabled=self.current_page == 0):
                st.session_state[f"{page_key}_current"] = 0
                st.rerun()
                
        with col2:
            if st.button("◀ 上一页", key=f"{page_key}_prev", disabled=self.current_page == 0):
                st.session_state[f"{page_key}_current"] = self.current_page - 1
                st.rerun()
                
        with col3:
            st.write(f"第 {self.current_page + 1} / {self.total_pages} 页 (共 {total_rows} 条)")
            
        with col4:
            if st.button("下一页 ▶", key=f"{page_key}_next", disabled=self.current_page >= self.total_pages - 1):
                st.session_state[f"{page_key}_current"] = self.current_page + 1
                st.rerun()
                
        with col5:
            if st.button("末页 ⏭", key=f"{page_key}_last", disabled=self.current_page >= self.total_pages - 1):
                st.session_state[f"{page_key}_current"] = self.total_pages - 1
                st.rerun()
                
        # 返回当前页数据
        return df.iloc[start_idx:end_idx]
        
    def paginate_expandable_items(self, items: List, page_key: str = "expand_page") -> List:
        """分页显示可展开项目（如订单）"""
        if not items:
            return items
            
        total_items = len(items)
        self.total_pages = (total_items + self.page_size - 1) // self.page_size
        
        # 获取当前页码
        if f"{page_key}_current" in st.session_state:
            self.current_page = st.session_state[f"{page_key}_current"]
        else:
            self.current_page = 0
            
        # 确保页码在有效范围内
        self.current_page = max(0, min(self.current_page, self.total_pages - 1))
        
        # 计算起止索引
        start_idx = self.current_page * self.page_size
        end_idx = min(start_idx + self.page_size, total_items)
        
        # 显示分页信息
        st.info(f"显示第 {start_idx + 1} - {end_idx} 个项目，共 {total_items} 个")
        
        # 分页控制按钮组
        cols = st.columns(5)
        with cols[1]:
            if st.button("上一页", key=f"{page_key}_prev", disabled=self.current_page == 0):
                st.session_state[f"{page_key}_current"] = self.current_page - 1
                st.rerun()
        with cols[2]:
            st.write(f"第 {self.current_page + 1}/{self.total_pages} 页")
        with cols[3]:
            if st.button("下一页", key=f"{page_key}_next", disabled=self.current_page >= self.total_pages - 1):
                st.session_state[f"{page_key}_current"] = self.current_page + 1
                st.rerun()
                
        # 返回当前页项目
        return items[start_idx:end_idx]

# ========== 方案五：Streamlit配置优化 ==========

def apply_optimized_config():
    """应用优化的Streamlit配置"""
    
    # 1. 设置页面配置（必须在任何其他st调用之前）
    try:
        st.set_page_config(
            page_title="银图PMC分析平台",
            page_icon="📊",
            layout="wide",
            initial_sidebar_state="collapsed",  # 初始折叠侧边栏减少组件
            menu_items={
                'Get Help': None,
                'Report a bug': None,
                'About': None  # 移除菜单项减少开销
            }
        )
    except:
        pass  # 可能已经设置过
        
    # 2. 注入自定义CSS优化渲染
    st.markdown("""
    <style>
    /* 减少不必要的动画 */
    .stApp {
        animation: none !important;
    }
    
    /* 优化表格渲染 */
    .dataframe {
        font-size: 12px !important;
    }
    
    /* 限制展开器动画 */
    .streamlit-expanderHeader {
        transition: none !important;
    }
    
    /* 减少按钮状态变化 */
    .stButton button {
        transition: none !important;
    }
    
    /* 隐藏不必要的元素 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 优化滚动性能 */
    .main {
        overflow-y: auto;
        -webkit-overflow-scrolling: touch;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # 3. 设置运行时配置
    if hasattr(st, 'runtime'):
        try:
            # 禁用文件监视（减少开销）
            st.runtime.exists._file_watchers = []
        except:
            pass
            
    # 4. 限制缓存大小
    if hasattr(st, 'cache_data'):
        st.cache_data.clear()  # 清理旧缓存
        
    # 5. 强制垃圾回收
    gc.collect()
    
    return True

# ========== 集成使用示例 ==========

class OptimizedStreamlitApp:
    """优化后的Streamlit应用框架"""
    
    def __init__(self):
        self.session_cleaner = SessionStateCleaner(max_items=35)
        self.paginator = DataPaginator(page_size=30)
        self.initialized = False
        
    def initialize(self):
        """初始化应用"""
        if not self.initialized:
            apply_optimized_config()
            self.initialized = True
            
    def safe_render_dataframe(self, df: pd.DataFrame, key: str = "df"):
        """安全渲染DataFrame（带分页）"""
        if df is None or df.empty:
            st.warning("没有数据可显示")
            return
            
        # 使用分页显示
        paged_df = self.paginator.paginate_dataframe(df, page_key=key)
        
        # 渲染数据
        st.dataframe(paged_df, use_container_width=True)
        
    def safe_render_expandable_list(self, items: List[Dict], key: str = "items"):
        """安全渲染可展开列表（带分页）"""
        if not items:
            st.info("没有项目可显示")
            return
            
        # 分页显示项目
        paged_items = self.paginator.paginate_expandable_items(items, page_key=key)
        
        # 渲染每个项目
        for idx, item in enumerate(paged_items):
            # 使用安全的key生成
            safe_key = f"{key}_{self.paginator.current_page}_{idx}"
            
            with st.expander(f"{item.get('title', f'项目 {idx+1}')}"):
                # 显示项目内容
                for k, v in item.items():
                    if k != 'title':
                        st.write(f"**{k}**: {v}")
                        
    def get_state(self, key: str, default=None):
        """获取状态值"""
        return self.session_cleaner.get_state(key, default)
        
    def set_state(self, key: str, value: Any):
        """设置状态值"""
        return self.session_cleaner.set_state(key, value)
        
    def cleanup(self):
        """手动触发清理"""
        self.session_cleaner.cleanup()
        gc.collect()

# ========== 导出给主应用使用 ==========

def create_optimized_app():
    """创建优化后的应用实例"""
    return OptimizedStreamlitApp()

# 使用示例
if __name__ == "__main__":
    # 创建优化应用
    app = create_optimized_app()
    app.initialize()
    
    st.title("Streamlit优化测试")
    
    # 测试数据
    test_df = pd.DataFrame({
        '订单号': [f'ORD{i:04d}' for i in range(200)],
        '金额': np.random.randint(1000, 10000, 200),
        '状态': np.random.choice(['处理中', '已完成', '待确认'], 200)
    })
    
    # 使用优化的渲染
    st.header("分页数据表")
    app.safe_render_dataframe(test_df, key="test_table")
    
    # 测试可展开项目
    test_items = [
        {'title': f'订单{i}', '金额': i*1000, '状态': '处理中'}
        for i in range(100)
    ]
    
    st.header("分页展开列表")
    app.safe_render_expandable_list(test_items, key="test_list")
    
    # 显示清理按钮
    if st.button("手动清理缓存"):
        app.cleanup()
        st.success("缓存已清理")
        st.rerun()
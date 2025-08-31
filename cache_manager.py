#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PMC系统数据库缓存层
提供高性能数据缓存和管理功能
"""

import sqlite3
import json
import pickle
import time
import hashlib
import threading
import pandas as pd
from datetime import datetime, timedelta
from typing import Any, Optional, Dict, List, Union
from contextlib import contextmanager
import os

class DatabaseCacheManager:
    """数据库缓存管理器"""
    
    def __init__(self, db_path: str = "cache/pmc_cache.db"):
        """
        初始化缓存管理器
        
        Args:
            db_path: SQLite数据库文件路径
        """
        self.db_path = db_path
        self.lock = threading.RLock()
        self._stats = {
            'hits': 0,
            'misses': 0,
            'sets': 0,
            'deletes': 0,
            'errors': 0
        }
        
        # 确保目录存在
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        
        # 初始化数据库
        self._init_database()
        
        print(f"✅ 缓存管理器已初始化: {db_path}")
    
    def _init_database(self):
        """初始化数据库表结构"""
        with self._get_connection() as conn:
            conn.executescript('''
                -- 缓存数据表
                CREATE TABLE IF NOT EXISTS cache_data (
                    cache_key TEXT PRIMARY KEY,
                    cache_value BLOB NOT NULL,
                    data_type TEXT NOT NULL DEFAULT 'pickle',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    expires_at TIMESTAMP,
                    access_count INTEGER DEFAULT 0,
                    last_accessed TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    data_size INTEGER DEFAULT 0,
                    tags TEXT DEFAULT ''
                );
                
                -- 缓存统计表
                CREATE TABLE IF NOT EXISTS cache_stats (
                    stat_date DATE PRIMARY KEY,
                    hits INTEGER DEFAULT 0,
                    misses INTEGER DEFAULT 0,
                    sets INTEGER DEFAULT 0,
                    deletes INTEGER DEFAULT 0,
                    errors INTEGER DEFAULT 0,
                    total_size INTEGER DEFAULT 0,
                    active_keys INTEGER DEFAULT 0
                );
                
                -- 索引优化
                CREATE INDEX IF NOT EXISTS idx_expires_at ON cache_data(expires_at);
                CREATE INDEX IF NOT EXISTS idx_tags ON cache_data(tags);
                CREATE INDEX IF NOT EXISTS idx_created_at ON cache_data(created_at);
                CREATE INDEX IF NOT EXISTS idx_last_accessed ON cache_data(last_accessed);
            ''')
    
    @contextmanager
    def _get_connection(self):
        """获取数据库连接（线程安全）"""
        with self.lock:
            conn = sqlite3.connect(self.db_path, timeout=30.0)
            conn.row_factory = sqlite3.Row
            try:
                yield conn
                conn.commit()
            except Exception as e:
                conn.rollback()
                self._stats['errors'] += 1
                raise
            finally:
                conn.close()
    
    def set(self, key: str, value: Any, ttl: int = 3600, tags: str = "") -> bool:
        """
        设置缓存
        
        Args:
            key: 缓存键
            value: 缓存值（支持任意Python对象）
            ttl: 过期时间（秒），0表示永不过期
            tags: 标签（用于分类管理）
        
        Returns:
            bool: 是否设置成功
        """
        try:
            # 序列化数据 - 统一使用pickle处理复杂对象
            if isinstance(value, pd.DataFrame):
                import io
                buffer = io.BytesIO()
                value.to_pickle(buffer)
                serialized_value = buffer.getvalue()
                data_type = 'dataframe'
            elif isinstance(value, dict) and any(isinstance(v, pd.DataFrame) for v in value.values()):
                # 字典包含DataFrame，使用pickle
                serialized_value = pickle.dumps(value)
                data_type = 'pickle'
            elif isinstance(value, (dict, list)) and not any(isinstance(item, pd.DataFrame) for item in (value.values() if isinstance(value, dict) else value)):
                # 简单字典/列表，使用JSON
                serialized_value = json.dumps(value, ensure_ascii=False).encode('utf-8')
                data_type = 'json'
            else:
                # 其他复杂对象，使用pickle
                serialized_value = pickle.dumps(value)
                data_type = 'pickle'
            
            # 计算过期时间
            expires_at = None
            if ttl > 0:
                expires_at = datetime.now() + timedelta(seconds=ttl)
            
            # 存储到数据库
            with self._get_connection() as conn:
                conn.execute('''
                    INSERT OR REPLACE INTO cache_data 
                    (cache_key, cache_value, data_type, expires_at, data_size, tags)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (key, serialized_value, data_type, expires_at, len(serialized_value), tags))
            
            self._stats['sets'] += 1
            return True
            
        except Exception as e:
            print(f"❌ 缓存设置失败 {key}: {e}")
            self._stats['errors'] += 1
            return False
    
    def get(self, key: str) -> Optional[Any]:
        """
        获取缓存
        
        Args:
            key: 缓存键
        
        Returns:
            缓存值或None
        """
        try:
            with self._get_connection() as conn:
                row = conn.execute('''
                    SELECT cache_value, data_type, expires_at 
                    FROM cache_data 
                    WHERE cache_key = ?
                ''', (key,)).fetchone()
                
                if not row:
                    self._stats['misses'] += 1
                    return None
                
                # 检查是否过期
                if row['expires_at']:
                    expires_at = datetime.fromisoformat(row['expires_at'])
                    if datetime.now() > expires_at:
                        # 删除过期数据
                        conn.execute('DELETE FROM cache_data WHERE cache_key = ?', (key,))
                        self._stats['misses'] += 1
                        return None
                
                # 更新访问统计
                conn.execute('''
                    UPDATE cache_data 
                    SET access_count = access_count + 1, last_accessed = CURRENT_TIMESTAMP
                    WHERE cache_key = ?
                ''', (key,))
                
                # 反序列化数据
                data_type = row['data_type']
                serialized_value = row['cache_value']
                
                if data_type == 'json':
                    value = json.loads(serialized_value.decode('utf-8'))
                elif data_type == 'dataframe':
                    import io
                    buffer = io.BytesIO(serialized_value)
                    value = pd.read_pickle(buffer)
                else:
                    value = pickle.loads(serialized_value)
                
                self._stats['hits'] += 1
                return value
                
        except Exception as e:
            print(f"❌ 缓存获取失败 {key}: {e}")
            self._stats['errors'] += 1
            return None
    
    def delete(self, key: str) -> bool:
        """删除缓存"""
        try:
            with self._get_connection() as conn:
                cursor = conn.execute('DELETE FROM cache_data WHERE cache_key = ?', (key,))
                deleted = cursor.rowcount > 0
                if deleted:
                    self._stats['deletes'] += 1
                return deleted
        except Exception as e:
            print(f"❌ 缓存删除失败 {key}: {e}")
            self._stats['errors'] += 1
            return False
    
    def clear_by_tags(self, tags: str) -> int:
        """按标签清除缓存"""
        try:
            with self._get_connection() as conn:
                cursor = conn.execute(
                    'DELETE FROM cache_data WHERE tags LIKE ?', 
                    (f'%{tags}%',)
                )
                deleted_count = cursor.rowcount
                self._stats['deletes'] += deleted_count
                return deleted_count
        except Exception as e:
            print(f"❌ 按标签清除缓存失败 {tags}: {e}")
            self._stats['errors'] += 1
            return 0
    
    def clear_expired(self) -> int:
        """清理过期缓存"""
        try:
            with self._get_connection() as conn:
                cursor = conn.execute('''
                    DELETE FROM cache_data 
                    WHERE expires_at IS NOT NULL AND expires_at < CURRENT_TIMESTAMP
                ''')
                deleted_count = cursor.rowcount
                self._stats['deletes'] += deleted_count
                return deleted_count
        except Exception as e:
            print(f"❌ 清理过期缓存失败: {e}")
            self._stats['errors'] += 1
            return 0
    
    def get_stats(self) -> Dict[str, Any]:
        """获取缓存统计信息"""
        try:
            with self._get_connection() as conn:
                # 实时统计
                cache_stats = conn.execute('''
                    SELECT 
                        COUNT(*) as total_keys,
                        SUM(data_size) as total_size,
                        SUM(access_count) as total_accesses,
                        COUNT(CASE WHEN expires_at IS NULL THEN 1 END) as permanent_keys,
                        COUNT(CASE WHEN expires_at > CURRENT_TIMESTAMP THEN 1 END) as active_keys,
                        COUNT(CASE WHEN expires_at <= CURRENT_TIMESTAMP THEN 1 END) as expired_keys
                    FROM cache_data
                ''').fetchone()
                
                hit_rate = 0
                if self._stats['hits'] + self._stats['misses'] > 0:
                    hit_rate = self._stats['hits'] / (self._stats['hits'] + self._stats['misses'])
                
                return {
                    'runtime_stats': dict(self._stats),
                    'hit_rate': f"{hit_rate:.1%}",
                    'total_keys': cache_stats['total_keys'] or 0,
                    'total_size': cache_stats['total_size'] or 0,
                    'total_size_mb': (cache_stats['total_size'] or 0) / (1024 * 1024),
                    'total_accesses': cache_stats['total_accesses'] or 0,
                    'permanent_keys': cache_stats['permanent_keys'] or 0,
                    'active_keys': cache_stats['active_keys'] or 0,
                    'expired_keys': cache_stats['expired_keys'] or 0
                }
        except Exception as e:
            print(f"❌ 获取统计信息失败: {e}")
            return {'error': str(e)}
    
    def get_keys(self, pattern: str = "%") -> List[str]:
        """获取缓存键列表"""
        try:
            with self._get_connection() as conn:
                rows = conn.execute(
                    'SELECT cache_key FROM cache_data WHERE cache_key LIKE ? ORDER BY last_accessed DESC',
                    (pattern,)
                ).fetchall()
                return [row['cache_key'] for row in rows]
        except Exception as e:
            print(f"❌ 获取键列表失败: {e}")
            return []
    
    def get_or_set(self, key: str, value_func, ttl: int = 3600, tags: str = "") -> Any:
        """获取缓存，如果不存在则设置"""
        cached_value = self.get(key)
        if cached_value is not None:
            return cached_value
        
        # 计算新值
        new_value = value_func()
        self.set(key, new_value, ttl=ttl, tags=tags)
        return new_value

# 全局缓存管理器实例
cache_manager = DatabaseCacheManager()

# 便捷装饰器
def cached(ttl: int = 3600, tags: str = ""):
    """缓存装饰器"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            # 生成缓存键
            key_data = f"{func.__name__}:{str(args)}:{str(sorted(kwargs.items()))}"
            cache_key = hashlib.md5(key_data.encode()).hexdigest()
            
            return cache_manager.get_or_set(
                cache_key, 
                lambda: func(*args, **kwargs),
                ttl=ttl,
                tags=tags
            )
        return wrapper
    return decorator

if __name__ == "__main__":
    # 测试缓存管理器
    print("🧪 测试数据库缓存管理器")
    
    # 测试基本功能
    cache_manager.set("test_key", {"data": "测试数据", "timestamp": time.time()}, ttl=60)
    print(f"✅ 获取测试数据: {cache_manager.get('test_key')}")
    
    # 测试DataFrame缓存
    test_df = pd.DataFrame({"A": [1, 2, 3], "B": ["a", "b", "c"]})
    cache_manager.set("test_df", test_df, ttl=300, tags="dataframe,test")
    cached_df = cache_manager.get("test_df")
    print(f"✅ DataFrame缓存: {len(cached_df)}行")
    
    # 显示统计信息
    stats = cache_manager.get_stats()
    print(f"📊 缓存统计: {stats}")
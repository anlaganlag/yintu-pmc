#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit Dashboard 缓存集成
将PMC缓存系统集成到现有的Dashboard中
"""

import streamlit as st
import pandas as pd
import hashlib
from datetime import datetime
from pmc_cache_adapter import pmc_cache, pmc_cached

def create_streamlit_cache_integration():
    """创建Streamlit缓存集成代码"""
    
    integration_code = '''
# =============================================================================
# 添加到 streamlit_dashboard.py 顶部 - PMC缓存系统集成
# =============================================================================

# 导入缓存模块
from pmc_cache_adapter import pmc_cache, pmc_cached
import hashlib
from datetime import datetime

# =============================================================================
# 1. 替换原有的 load_latest_report 函数
# =============================================================================

@pmc_cached(cache_type='analysis_report')
def load_cached_latest_report():
    """加载最新分析报告（带缓存）"""
    try:
        # 原有的文件发现逻辑
        import glob
        import os
        
        all_report_patterns = [
            "银图PMC综合物料分析报告_*.xlsx",
            r"D:\\yingtu-PMC\\精准供应商物料分析报告_含回款_*.xlsx",
            r"D:\\yingtu-PMC\\精准供应商物料分析报告_2025*.xlsx",
            "精准供应商物料分析报告_*.xlsx"
        ]
        
        all_files = []
        for pattern in all_report_patterns:
            files = glob.glob(pattern)
            all_files.extend([(f, os.path.getmtime(f)) for f in files])
        
        if all_files:
            latest_file = max(all_files, key=lambda x: x[1])[0]
            
            # 使用缓存键包含文件修改时间
            file_mtime = os.path.getmtime(latest_file)
            cache_key = f"report:{hashlib.md5(latest_file.encode()).hexdigest()}:{int(file_mtime)}"
            
            # 尝试从缓存获取
            cached_data = pmc_cache.get_analysis_report(cache_key)
            if cached_data:
                print(f"✅ 从缓存加载报告: {latest_file}")
                return cached_data
            
            # 加载新数据
            if "含回款" in latest_file:
                df = pd.read_excel(latest_file)
                excel_data = {'1_订单缺料明细': df}
            else:
                excel_data = pd.read_excel(latest_file, sheet_name=None)
                if "综合物料分析明细" in excel_data:
                    excel_data['1_订单缺料明细'] = excel_data.pop('综合物料分析明细')
            
            # 缓存数据
            pmc_cache.set_analysis_report(excel_data, cache_key)
            print(f"✅ 加载并缓存报告: {latest_file}")
            return excel_data
        
        return None
        
    except Exception as e:
        st.error(f"❌ 加载报告失败: {e}")
        return None

# =============================================================================
# 2. 带缓存的数据筛选函数
# =============================================================================

def get_cached_filtered_data(original_df, filters):
    """获取缓存的筛选数据"""
    # 生成筛选器哈希
    filter_str = f"{len(original_df)}:{sorted(filters.items())}"
    filter_hash = hashlib.md5(filter_str.encode()).hexdigest()
    
    # 尝试从缓存获取
    cached_result = pmc_cache.get_filtered_data(filters)
    if cached_result is not None:
        return cached_result
    
    # 执行筛选逻辑（这里需要根据实际筛选逻辑调整）
    filtered_df = original_df.copy()
    
    # 示例筛选逻辑
    if filters.get('month_filter') and filters['month_filter'] != '全部':
        filtered_df = filtered_df[filtered_df['月份'] == filters['month_filter']]
    
    if filters.get('roi_threshold'):
        filtered_df = filtered_df[filtered_df['每元投入回款'] >= filters['roi_threshold']]
    
    # 缓存结果
    pmc_cache.set_filtered_data(filtered_df, filters)
    return filtered_df

# =============================================================================
# 3. ROI计算缓存
# =============================================================================

def get_cached_roi_calculation(selected_orders, analysis_df):
    """获取缓存的ROI计算结果"""
    
    if not selected_orders:
        return None
    
    # 尝试从缓存获取
    cached_roi = pmc_cache.get_roi_calculation(list(selected_orders))
    if cached_roi:
        return cached_roi
    
    # 计算ROI（根据实际逻辑调整）
    selected_data = analysis_df[analysis_df['生产订单号'].isin(selected_orders)]
    
    roi_summary = selected_data.groupby('生产订单号').agg({
        '订单金额(RMB)': 'first',
        '欠料金额(RMB)': 'sum'
    }).reset_index()
    
    roi_summary['订单ROI'] = np.where(
        roi_summary['欠料金额(RMB)'] > 0,
        roi_summary['订单金额(RMB)'] / roi_summary['欠料金额(RMB)'],
        999
    )
    
    total_investment = roi_summary['欠料金额(RMB)'].sum()
    total_return = roi_summary['订单金额(RMB)'].sum()
    avg_roi = total_return / total_investment if total_investment > 0 else 0
    
    result = {
        'summary': roi_summary,
        'total_investment': total_investment,
        'total_return': total_return,
        'avg_roi': avg_roi,
        'order_count': len(selected_orders)
    }
    
    # 缓存结果
    pmc_cache.set_roi_calculation(result, list(selected_orders))
    return result

# =============================================================================
# 4. 安全的数据编辑器（集成缓存）
# =============================================================================

def safe_cached_data_editor(df, key_suffix="", **kwargs):
    """安全的数据编辑器，集成缓存功能"""
    
    # 生成数据版本号
    data_version = f"{len(df)}:{hashlib.md5(str(df.columns.tolist()).encode()).hexdigest()[:8]}"
    
    # 检查缓存的稳定数据
    cache_key = f"stable_data:{key_suffix}:{data_version}"
    stable_df = pmc_cache.cache.get(cache_key)
    
    if stable_df is None:
        stable_df = df.copy().reset_index(drop=True)
        # 缓存稳定数据（短时间缓存）
        pmc_cache.cache.set(cache_key, stable_df, ttl=300, tags='stable,ui')
    
    # 数据一致性检查
    hash_key = f'data_hash_{key_suffix}'
    current_hash = hashlib.md5(data_version.encode()).hexdigest()[:8]
    
    if hash_key not in st.session_state:
        st.session_state[hash_key] = current_hash
    
    if st.session_state[hash_key] != current_hash:
        # 数据发生变化，重置状态
        if hasattr(st.session_state, 'selected_orders'):
            st.session_state.selected_orders = set()
        st.session_state[hash_key] = current_hash
        st.warning(f"⚠️ 数据已更新，已重置选择状态", icon="🔄")
    
    # 使用稳定数据
    unique_key = f"data_editor_{key_suffix}_{current_hash}"
    
    try:
        return st.data_editor(stable_df, key=unique_key, **kwargs)
    except Exception as e:
        st.error(f"❌ 数据编辑器错误: {str(e)}")
        # 清除相关缓存
        pmc_cache.cache.delete(cache_key)
        return stable_df

# =============================================================================
# 5. 使用方法 - 在原有代码中替换以下函数调用
# =============================================================================

# 替换1: 数据加载
# 原来: excel_data = load_latest_report()
# 替换为: excel_data = load_cached_latest_report()

# 替换2: 数据编辑器  
# 原来: edited_df = st.data_editor(df, ...)
# 替换为: edited_df = safe_cached_data_editor(df, key_suffix='main', ...)

# 替换3: 筛选数据
# 在筛选逻辑中使用: filtered_df = get_cached_filtered_data(df, filters)

# 替换4: ROI计算
# 在ROI分析中使用: roi_result = get_cached_roi_calculation(selected_orders, df)

'''
    
    return integration_code

def create_cache_management_ui():
    """创建缓存管理界面代码"""
    
    ui_code = '''
# =============================================================================
# 缓存管理界面 - 添加到 streamlit_dashboard.py 的侧边栏或新页面
# =============================================================================

def show_cache_management():
    """显示缓存管理界面"""
    st.markdown("### 🗄️ 缓存管理")
    
    # 获取缓存健康状态
    health = pmc_cache.get_cache_health()
    
    # 显示健康评分
    col1, col2 = st.columns(2)
    with col1:
        st.metric("缓存健康评分", health['health_score'])
    with col2:
        hit_rate = health['overall_stats'].get('hit_rate', '0%')
        st.metric("缓存命中率", hit_rate)
    
    # 显示统计信息
    stats = health['overall_stats']
    
    st.markdown("#### 📊 详细统计")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("总缓存数", f"{stats.get('total_keys', 0)}个")
        st.metric("缓存大小", f"{stats.get('total_size_mb', 0):.1f}MB")
    
    with col2:
        st.metric("缓存命中", f"{stats.get('runtime_stats', {}).get('hits', 0)}次")
        st.metric("缓存未命中", f"{stats.get('runtime_stats', {}).get('misses', 0)}次")
    
    with col3:
        st.metric("活跃键", f"{stats.get('active_keys', 0)}个")
        st.metric("过期键", f"{stats.get('expired_keys', 0)}个")
    
    # 显示建议
    st.markdown("#### 💡 优化建议")
    for recommendation in health['recommendations']:
        st.info(recommendation)
    
    # 缓存操作
    st.markdown("#### 🛠️ 缓存操作")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("🧹 清理过期缓存"):
            cleared = pmc_cache.cache.clear_expired()
            st.success(f"✅ 已清理 {cleared} 个过期缓存")
    
    with col2:
        if st.button("🔄 清除分析缓存"):
            cleared = pmc_cache.clear_analysis_cache()
            st.success(f"✅ 已清除 {cleared} 个分析缓存")
    
    with col3:
        if st.button("📊 刷新统计"):
            st.experimental_rerun()
    
    # 显示缓存类型分布
    if health['cache_types']:
        st.markdown("#### 📂 缓存分类")
        cache_types_df = pd.DataFrame([
            {'类型': k, '数量': v} for k, v in health['cache_types'].items()
        ])
        st.dataframe(cache_types_df, use_container_width=True)

# 在主界面添加缓存管理标签页
def add_cache_tab_to_main():
    """在主界面添加缓存管理标签"""
    # 在现有的标签页中添加
    tab1, tab2, tab3, tab_cache = st.tabs(["📊 管理概览", "📋 采购清单", "🏭 供应商视图", "🗄️ 缓存管理"])
    
    with tab_cache:
        show_cache_management()
'''
    
    return ui_code

if __name__ == "__main__":
    print("🔧 Streamlit Dashboard 缓存集成")
    print("=" * 60)
    
    # 生成集成代码
    integration = create_streamlit_cache_integration()
    with open('streamlit_cache_integration_code.py', 'w', encoding='utf-8') as f:
        f.write(integration)
    
    # 生成界面代码
    ui_code = create_cache_management_ui()
    with open('cache_management_ui_code.py', 'w', encoding='utf-8') as f:
        f.write(ui_code)
    
    print("✅ 已生成集成代码文件:")
    print("  📄 streamlit_cache_integration_code.py - 核心缓存集成")
    print("  📄 cache_management_ui_code.py - 缓存管理界面")
    
    print(f"\n📋 集成步骤:")
    print("1. 将 streamlit_cache_integration_code.py 的内容添加到 streamlit_dashboard.py 顶部")
    print("2. 替换相关函数调用（load_latest_report → load_cached_latest_report）")
    print("3. 替换 st.data_editor → safe_cached_data_editor")
    print("4. 添加缓存管理界面")
    print("5. 重启 streamlit 服务")
    
    print(f"\n🎯 预期效果:")
    print("✅ 解决 'Bad setIn index' 错误")
    print("✅ 大幅提升页面响应速度")
    print("✅ 减少重复数据计算")
    print("✅ 提供缓存管理和监控")
    print("=" * 60)
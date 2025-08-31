
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
            r"D:\yingtu-PMC\精准供应商物料分析报告_含回款_*.xlsx",
            r"D:\yingtu-PMC\精准供应商物料分析报告_2025*.xlsx",
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


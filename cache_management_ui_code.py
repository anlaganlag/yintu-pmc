
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

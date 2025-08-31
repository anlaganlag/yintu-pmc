
# =============================================================================
# 紧急修复：添加到 streamlit_dashboard.py 顶部
# =============================================================================

import hashlib
import pandas as pd
import streamlit as st

def safe_data_editor(df, key_suffix="", **kwargs):
    """
    安全的数据编辑器，防止索引不一致错误
    
    Args:
        df: DataFrame数据
        key_suffix: 唯一标识符
        **kwargs: data_editor的其他参数
    """
    # 1. 创建数据快照，确保稳定性
    stable_df = df.copy().reset_index(drop=True)
    
    # 2. 生成数据哈希值用于检测变化
    try:
        data_hash = hashlib.md5(
            str(len(stable_df)).encode() + 
            str(stable_df.columns.tolist()).encode()
        ).hexdigest()[:8]
    except:
        data_hash = f"len_{len(stable_df)}"
    
    # 3. 检查数据一致性
    hash_key = f'data_hash_{key_suffix}'
    if hash_key not in st.session_state:
        st.session_state[hash_key] = data_hash
    
    # 4. 如果数据发生变化，重置相关状态
    if st.session_state[hash_key] != data_hash:
        # 重置选择状态
        if hasattr(st.session_state, 'selected_orders'):
            st.session_state.selected_orders = set()
        
        # 更新哈希值
        st.session_state[hash_key] = data_hash
        
        # 提示用户
        st.warning(f"⚠️ 数据已更新（{len(stable_df)}行），已重置选择状态", icon="🔄")
    
    # 5. 添加长度验证
    if len(stable_df) == 0:
        st.info("📋 暂无数据显示")
        return pd.DataFrame()
    
    # 6. 使用稳定的唯一键
    unique_key = f"data_editor_{key_suffix}_{data_hash}"
    
    try:
        # 7. 调用原始data_editor，使用稳定数据
        return st.data_editor(
            stable_df,
            key=unique_key,
            **kwargs
        )
    except Exception as e:
        st.error(f"❌ 数据编辑器错误: {str(e)}")
        st.info("🔧 已自动重置，请重试")
        
        # 强制清除相关状态
        keys_to_clear = [k for k in st.session_state.keys() if key_suffix in k]
        for key in keys_to_clear:
            del st.session_state[key]
        
        return stable_df

@st.cache_data(ttl=1800, show_spinner="🔄 加载分析数据...")
def get_stable_analysis_data():
    """获取稳定的分析数据（30分钟缓存）"""
    try:
        from streamlit_dashboard import load_latest_report
        return load_latest_report()
    except:
        st.error("❌ 无法加载分析数据")
        return None

# =============================================================================
# 使用方法：替换原来的 st.data_editor 调用
# =============================================================================

# 原来的代码：
# edited_df = st.data_editor(editor_df, ...)

# 替换为：
# edited_df = safe_data_editor(
#     editor_df, 
#     key_suffix="main_editor",
#     column_config={...},
#     use_container_width=True
# )

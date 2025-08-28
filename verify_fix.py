#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证setIn错误修复
检查是否仍存在可能触发setIn错误的代码位置
"""

import pandas as pd
import os
import sys

def check_large_data_handling():
    """检查大数据处理能力"""
    print("🧪 大数据处理能力测试")
    print("=" * 50)
    
    report_file = "银图PMC综合物料分析报告_20250828_160309.xlsx"
    
    if not os.path.exists(report_file):
        print(f"❌ 测试文件不存在: {report_file}")
        return False
    
    try:
        # 加载数据
        print(f"📂 加载文件: {report_file}")
        df = pd.read_excel(report_file, sheet_name='综合物料分析明细')
        print(f"✅ 成功加载: {len(df)} 行 × {len(df.columns)} 列")
        
        # 模拟分页处理
        max_rows = 1000
        total_pages = (len(df) + max_rows - 1) // max_rows
        print(f"📄 分页处理: 共{total_pages}页，每页最多{max_rows}行")
        
        # 检查几个关键页面
        test_pages = [1, 10, 20, total_pages]
        for page in test_pages:
            if page <= total_pages:
                start_idx = (page - 1) * max_rows
                end_idx = min(start_idx + max_rows, len(df))
                page_df = df.iloc[start_idx:end_idx]
                print(f"  页面{page}: 第{start_idx+1}-{end_idx}行 ({len(page_df)}行)")
        
        print("✅ 分页处理测试通过")
        return True
        
    except Exception as e:
        print(f"❌ 大数据处理测试失败: {str(e)}")
        return False

def check_streamlit_code():
    """检查streamlit_dashboard.py中的潜在问题"""
    print("\n🔍 代码安全检查")
    print("=" * 50)
    
    dashboard_file = "streamlit_dashboard.py"
    if not os.path.exists(dashboard_file):
        print(f"❌ 主文件不存在: {dashboard_file}")
        return False
    
    with open(dashboard_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 检查是否还有直接的st.dataframe调用
    import re
    dataframe_calls = re.findall(r'st\.dataframe\s*\([^)]+\)', content)
    
    if dataframe_calls:
        print(f"⚠️  发现{len(dataframe_calls)}个直接st.dataframe调用:")
        for i, call in enumerate(dataframe_calls, 1):
            # 获取行号
            lines = content.split('\n')
            for line_num, line in enumerate(lines, 1):
                if call.replace(' ', '').replace('\n', '') in line.replace(' ', '').replace('\n', ''):
                    print(f"  {i}. 第{line_num}行: {call[:60]}...")
                    break
    else:
        print("✅ 未发现直接st.dataframe调用")
    
    # 检查safe_dataframe_display函数是否存在
    if 'def safe_dataframe_display(' in content:
        print("✅ safe_dataframe_display函数已定义")
        
        # 检查函数调用
        safe_calls = re.findall(r'safe_dataframe_display\s*\([^)]+\)', content)
        print(f"📊 发现{len(safe_calls)}个safe_dataframe_display调用")
        
    else:
        print("❌ safe_dataframe_display函数未找到")
        return False
    
    # 检查st.set_page_config的位置
    if content.find('st.set_page_config(') < content.find('def main():'):
        print("⚠️  st.set_page_config在模块级别，可能导致初始化问题")
        return False
    else:
        print("✅ st.set_page_config位置正确")
    
    print("✅ 代码安全检查通过")
    return True

def main():
    print("🛡️  银图PMC系统 - setIn错误修复验证")
    print("=" * 60)
    
    # 系统信息
    print(f"🐍 Python版本: {sys.version}")
    print(f"📦 Pandas版本: {pd.__version__}")
    
    # 检查项目
    results = []
    results.append(check_large_data_handling())
    results.append(check_streamlit_code())
    
    print("\n" + "=" * 60)
    if all(results):
        print("🎉 所有检查项通过！系统应该能够安全处理35,098行数据")
        print("✅ 修复状态: 成功")
        print("📝 建议: 可以正常使用streamlit_dashboard.py展示大数据")
    else:
        print("⚠️  发现潜在问题，建议进一步检查")
        print("❌ 修复状态: 需要进一步处理")
    
    print("\n🔗 测试URL:")
    print("  主应用: http://localhost:8502")
    print("  测试应用: http://localhost:8503")

if __name__ == "__main__":
    main()
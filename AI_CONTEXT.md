# AI编程工具快速理解指南

## 🎯 一句话说明
这是一个制造业PMC(生产计划与物料控制)分析系统，通过整合订单、欠料、库存、供应商数据，计算ROI并推荐最优采购方案。

## 📋 快速事实
- **语言**: Python 3.10+
- **框架**: Streamlit (Web) + Pandas (数据处理)
- **规模**: ~7000行代码，处理6844条订单记录
- **核心算法**: LEFT JOIN + ROI计算 + 供应商选择
- **输入**: 5个Excel文件
- **输出**: Web仪表板 + Excel分析报告

## 🏗️ 项目结构（按重要性排序）
```
yintu-pmc/
├── silverPlan_analysis.py      # ⭐ 核心分析引擎(必看)
├── streamlit_dashboard.py      # ⭐ Web界面(必看)
├── CLAUDE.md                   # 📚 项目说明文档
├── workflow.md                 # 📊 数据处理流程
├── requirements.txt            # 📦 依赖包列表
└── [其他辅助分析脚本]         # 可选了解
```

## 💡 核心概念速查

### 1. 业务术语
- **PMC**: Production & Material Control (生产与物料控制)
- **ROI**: 订单金额 ÷ 欠料金额 (投资回报率)
- **欠料**: 生产所需但库存不足的物料
- **生产单号**: 内部生产订单编号(主键)
- **客户订单号**: 客户的订单编号

### 2. 技术架构
```python
# 核心数据流
订单表(主表) → LEFT JOIN 欠料表 → LEFT JOIN 库存表 → LEFT JOIN 供应商表
     ↓
ROI计算 + 供应商选择
     ↓
Excel报告 + Web展示
```

### 3. 关键类/函数
```python
# silverPlan_analysis.py
class ComprehensivePMCAnalyzer:
    load_all_data()              # 加载5个Excel文件
    comprehensive_left_join_analysis()  # LEFT JOIN整合
    calculate_derived_fields()    # 计算ROI等指标
    generate_comprehensive_report()  # 生成报告

# streamlit_dashboard.py
load_data()                      # 加载最新报告
create_kpi_cards()              # 显示关键指标
main()                          # Streamlit主入口
```

## 🔧 常见任务指令

### 修复Bug时
1. 检查列名是否存在: `if '列名' in df.columns:`
2. 数值转换要安全: `pd.to_numeric(df['列'], errors='coerce')`
3. 聚合前检查列: 使用动态agg_dict而非硬编码

### 添加新功能时
1. 数据在`silverPlan_analysis.py`的`ComprehensivePMCAnalyzer`类处理
2. UI在`streamlit_dashboard.py`的`main()`函数添加
3. 保持LEFT JOIN架构，确保订单不丢失

### 优化性能时
1. 使用`@st.cache_data`缓存数据加载
2. DataFrame操作用向量化而非循环
3. 大数据集分批处理并显示进度

## ⚠️ 关键注意事项

### 数据完整性
- **必须**使用LEFT JOIN，不能用INNER JOIN（会丢失订单）
- 订单表是主表，其他都是补充信息
- 空值处理：数值填0，文本填空字符串

### ROI计算特殊逻辑
```python
if 欠料金额 > 0:
    ROI = 订单金额 / 欠料金额
elif 订单金额 > 0:
    ROI = "无需投入"  # 不缺料可直接生产
else:
    ROI = 0
```

### 货币转换
- USD → RMB: ×7.20
- HKD → RMB: ×0.93
- EUR → RMB: ×7.85

## 📊 数据规模参考
- 订单数: 394个
- 总记录: 6844条（一个订单可能缺多种物料）
- 供应商: 150+家
- 处理时间: <30秒

## 🚀 快速启动
```bash
# 安装依赖
pip install -r requirements.txt

# 运行分析（生成Excel报告）
python silverPlan_analysis.py

# 启动Web界面
streamlit run streamlit_dashboard.py
```

## 📝 调试技巧
1. 查看生成的Excel了解数据结构
2. 检查`workflow.md`了解处理流程
3. 关注`数据完整性标记`字段判断数据质量
4. PSO2501724是测试用的订单号

## 🎯 优化目标
- 将ROI从2.0-2.5提升到2.8-3.0（+20%）
- 减少30-40%的缺料率
- 5分钟内响应异常

---
**提示**: 如需深入了解，按顺序阅读：
1. `CLAUDE.md` - 业务背景
2. `workflow.md` - 技术流程
3. `silverPlan_analysis.py` - 核心实现
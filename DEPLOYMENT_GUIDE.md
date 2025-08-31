# 🚀 银图PMC订单分析系统 - Streamlit Cloud部署指南

## 📋 部署准备清单

### 1. 必需文件
确保以下文件已准备完毕：
```
yintu-pmc/
├── cloud_dashboard.py          # 云端优化版仪表板
├── requirements.txt            # Python依赖包
├── .streamlit/config.toml     # Streamlit配置
├── DEPLOYMENT_GUIDE.md       # 本部署指南
└── README.md                  # 项目说明（可选）
```

### 2. GitHub仓库准备
- ✅ 创建GitHub公开仓库
- ✅ 上传必要文件
- ✅ 确保仓库结构清晰

## 🛠️ 详细部署步骤

### 步骤 1: 创建GitHub仓库

1. 访问 [GitHub](https://github.com) 并登录
2. 点击 "New Repository" 创建新仓库
3. 仓库名称建议：`yintu-pmc-dashboard`
4. 设置为 **Public** （Streamlit Cloud免费版需要公开仓库）
5. 勾选 "Add a README file"

### 步骤 2: 上传项目文件

**方式A: 通过GitHub网页界面**
1. 在仓库页面点击 "Add file" → "Upload files"
2. 拖拽以下文件到页面：
   - `cloud_dashboard.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
3. 填写提交信息: "Initial deployment setup"
4. 点击 "Commit changes"

**方式B: 通过Git命令行**
```bash
# 克隆仓库
git clone https://github.com/your-username/yintu-pmc-dashboard.git
cd yintu-pmc-dashboard

# 复制文件
cp /path/to/cloud_dashboard.py .
cp /path/to/requirements.txt .
mkdir -p .streamlit
cp /path/to/.streamlit/config.toml .streamlit/

# 提交更改
git add .
git commit -m "Initial deployment setup"
git push origin main
```

### 步骤 3: 部署到Streamlit Cloud

1. **访问Streamlit Cloud**
   - 前往 [share.streamlit.io](https://share.streamlit.io)
   - 使用GitHub账号登录

2. **创建新应用**
   - 点击 "New app"
   - 选择您的GitHub仓库：`yintu-pmc-dashboard`
   - Branch: `main`
   - Main file path: `cloud_dashboard.py`
   - App URL: 自定义应用URL（如：`yintu-pmc-analysis`）

3. **高级设置**（可选）
   - Python version: 3.9
   - 添加环境变量（如需要）

4. **部署**
   - 点击 "Deploy!" 按钮
   - 等待部署完成（通常3-5分钟）

## 🎯 部署后配置

### 访问地址
部署成功后，您的应用将可通过以下地址访问：
```
https://your-app-name.streamlit.app
```

### 应用特性
- 🌐 **24/7在线访问** - 管理层随时查看
- 📱 **移动端适配** - 手机平板友好
- 🔒 **安全可靠** - HTTPS加密传输
- ⚡ **快速响应** - 云端CDN加速

## 📊 使用指南

### 1. 数据上传
- 管理层访问应用后，在左侧边栏上传Excel分析报告
- 支持文件：`560订单清单欠料分析报告_xxx.xlsx`
- 最大文件大小：200MB

### 2. 核心功能页面
- **🎯 优先级分析**: 订单ROI分布和投资风险评估
- **🔝 重点订单**: 欠料最高和ROI最高的订单
- **🏭 供应商分析**: 155家供应商的采购需求分析
- **📋 生产建议**: 基于ROI的生产排期建议

### 3. 交互功能
- 多维度筛选器
- 实时数据更新
- 可视化图表交互
- 数据导出功能

## 🔧 维护和更新

### 更新应用代码
1. 修改本地文件
2. 提交到GitHub仓库
3. Streamlit Cloud会自动重新部署

### 监控应用状态
- 访问 [share.streamlit.io](https://share.streamlit.io) 查看应用状态
- 查看部署日志和错误信息
- 监控应用使用情况

### 故障排除
常见问题及解决方案：

**问题1: 部署失败**
- 检查requirements.txt依赖包版本
- 确认Python代码无语法错误
- 查看部署日志错误信息

**问题2: 文件上传失败**
- 确认Excel文件格式正确
- 文件大小不超过200MB
- 工作表名称与代码匹配

**问题3: 图表显示异常**
- 清理浏览器缓存
- 检查数据格式完整性
- 确认plotly版本兼容性

## 📞 技术支持

### 联系方式
- 技术问题：[创建GitHub Issue](https://github.com/your-username/yintu-pmc-dashboard/issues)
- 功能建议：通过GitHub提交Pull Request

### 有用链接
- [Streamlit官方文档](https://docs.streamlit.io)
- [Streamlit Cloud帮助](https://docs.streamlit.io/streamlit-cloud)
- [Plotly图表库](https://plotly.com/python/)

## 🎉 部署成功检查清单

- ✅ 应用成功部署并可访问
- ✅ 文件上传功能正常
- ✅ 所有页面和图表正常显示
- ✅ 筛选器和交互功能工作正常
- ✅ 移动端访问测试通过
- ✅ 管理层用户培训完成

---

## 📈 应用特色功能预览

### KPI仪表板
- 560个订单总览
- 94个可立即生产订单（16.8%）
- 161个高优先级订单（28.8%）
- ¥13,086,682总欠料金额
- 7.71倍整体投资回报率

### 可视化图表
- 优先级分布饼图
- ROI vs 欠料金额散点图
- TOP10供应商采购需求
- 生产排期建议表

### 管理决策支持
- 基于ROI的生产优先级排序
- 供应商风险评估和优化建议
- 物料采购计划和预算控制
- 实时生产状态监控

**🎯 立即开始部署，让您的管理层随时随地做出数据驱动的决策！**
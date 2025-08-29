# 自动化数据分析报告生成工具

## 项目简介

这是一个专门为解决Excel数据分析报告繁琐流程而设计的Python自动化工具。它可以：

- 📊 **自动计算环比和同比增长率**：无需手动制作透视表
- 🔄 **多时间点数据对比**：自动处理当前期、上月、去年同期数据
- 📈 **多维度透视分析**：支持按地区、产品、渠道等多个维度进行交叉分析
- 📋 **一键生成报告**：自动输出格式化的Excel分析报告
- ⚡ **操作简单**：仅需几行代码即可完成复杂的数据分析

## 功能特点

### 1. 智能时间处理
- 自动识别数据中的时间点
- 智能匹配上月和去年同期数据
- 支持自定义分析基准日期

### 2. 灵活的维度分析
- 支持单维度分析（如按地区）
- 支持多维度交叉分析（如按地区+产品）
- 支持多种聚合方式（求和、平均、计数等）

### 3. 完整的增长率计算
- 环比增长率（与上月对比）
- 同比增长率（与去年同期对比）
- 自动处理缺失数据和除零情况

### 4. 丰富的报告输出
- 详细的分维度分析表
- 汇总统计信息
- 增长趋势概览
- 格式化的Excel报告

## 安装说明

### 环境要求
- Python 3.7+
- 支持Windows、macOS、Linux

### 安装依赖
```bash
pip install -r requirements.txt
```

或手动安装：
```bash
pip install pandas numpy openpyxl xlrd
```

## 快速开始

### 1. 准备数据
确保你的数据文件（Excel或CSV）包含：
- **日期列**：用于时间对比分析
- **分析维度列**：如地区、产品、渠道等
- **数值指标列**：如销售额、数量等

### 2. 运行示例
```bash
python example_usage.py
```

这将：
- 创建示例数据文件
- 执行完整的分析流程
- 生成分析报告

### 3. 查看结果
运行后会生成：
- `sample_sales_data.xlsx`：示例数据
- `数据分析报告_YYYYMMDD.xlsx`：分析报告

## 使用方法

### 基础用法

```python
from data_analysis_report import DataAnalysisReport

# 创建分析器
analyzer = DataAnalysisReport()

# 加载数据
df = analyzer.load_data('your_data.xlsx')

# 准备时间段数据
analyzer.prepare_data_by_periods(df, date_column='date')

# 按维度分析
result = analyzer.analyze_by_dimensions(
    dimensions=['region'],          # 分析维度
    value_column='sales_amount',    # 分析指标
    aggregation='sum'              # 聚合方式
)

# 查看结果
print(result)
```

### 完整分析流程

```python
from data_analysis_report import DataAnalysisReport

# 1. 初始化
analyzer = DataAnalysisReport()

# 2. 加载数据
df = analyzer.load_data('sales_data.xlsx')

# 3. 准备时间段数据（当前期、上月、去年同期）
analyzer.prepare_data_by_periods(df, date_column='date')

# 4. 定义分析配置
analyses = [
    {
        'name': '按地区分析',
        'dimensions': ['region'],
        'value_column': 'sales_amount',
        'aggregation': 'sum'
    },
    {
        'name': '按产品分析',
        'dimensions': ['product'],
        'value_column': 'sales_amount',
        'aggregation': 'sum'
    },
    {
        'name': '地区×产品交叉分析',
        'dimensions': ['region', 'product'],
        'value_column': 'sales_amount',
        'aggregation': 'sum'
    }
]

# 5. 执行分析
results = {}
for config in analyses:
    results[config['name']] = analyzer.analyze_by_dimensions(
        config['dimensions'], 
        config['value_column'], 
        config['aggregation']
    )

# 6. 生成汇总
summary = analyzer.generate_summary_report(results, analyses)

# 7. 导出报告
analyzer.export_report(results, summary, 'analysis_report.xlsx')
```

## 配置说明

### 数据文件要求

| 字段类型 | 说明 | 示例 |
|---------|------|------|
| 日期列 | 用于时间对比，必须为日期格式 | 2024-01-31 |
| 维度列 | 分析分组依据 | 地区、产品、渠道 |
| 指标列 | 数值型分析指标 | 销售额、数量、利润 |

### 聚合方式选择

| 聚合方式 | 适用场景 | 说明 |
|---------|---------|------|
| sum | 销售额、数量汇总 | 求和 |
| mean | 平均价格、评分 | 平均值 |
| count | 订单数、客户数 | 计数 |
| max | 最高价格、最大值 | 最大值 |
| min | 最低价格、最小值 | 最小值 |

### 分析维度设置

```python
# 单维度分析
dimensions = ['region']                    # 按地区

# 多维度交叉分析
dimensions = ['region', 'product']         # 按地区和产品
dimensions = ['region', 'product', 'channel']  # 三维交叉分析
```

## 输出报告说明

生成的Excel报告包含以下工作表：

### 1. 汇总分析
- 各分析维度的总体统计
- 增长率的平均值、最大值、最小值
- 正增长和负增长项目统计

### 2. 详细分析表（每个维度一个）
包含列：
- 分析维度字段
- 当前期数值
- 上月数值  
- 去年同期数值
- 环比增长率(%)
- 同比增长率(%)

## 常见问题

### Q1: 数据中某些时间点缺失怎么办？
A: 工具会自动处理缺失数据，缺失的增长率会显示为空值（NaN）。

### Q2: 如何处理除零错误？
A: 工具自动处理分母为零的情况，这种情况下增长率显示为空值。

### Q3: 可以分析多个指标吗？
A: 可以，为每个指标单独运行分析即可：

```python
# 分析销售额
result1 = analyzer.analyze_by_dimensions(['region'], 'sales_amount', 'sum')

# 分析销售数量
result2 = analyzer.analyze_by_dimensions(['region'], 'quantity', 'sum')
```

### Q4: 支持什么格式的数据文件？
A: 支持Excel（.xlsx, .xls）和CSV文件。

### Q5: 如何自定义分析时间点？
A: 在prepare_data_by_periods中指定target_date参数：

```python
from datetime import datetime
target_date = datetime(2024, 1, 31)
analyzer.prepare_data_by_periods(df, 'date', target_date)
```

## 技术支持

如有问题或建议，请查看代码注释或修改源码以适应特定需求。

## 版本历史

- v1.0.0: 初始版本，支持基础的多维度分析和增长率计算
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据分析报告工具使用示例
演示如何使用DataAnalysisReport类进行自动化分析
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from data_analysis_report import DataAnalysisReport

def create_sample_data():
    """
    创建示例数据用于演示
    """
    print("正在创建示例数据...")
    
    # 设置随机种子保证结果可复现
    np.random.seed(42)
    
    # 定义时间范围（过去13个月的数据）
    end_date = datetime(2024, 1, 31)
    dates = []
    for i in range(13):
        date = end_date - pd.DateOffset(months=i)
        dates.append(date)
    dates.reverse()
    
    # 定义维度数据
    regions = ['北京', '上海', '广州', '深圳', '杭州']
    categories = ['电子产品', '服装', '食品', '书籍', '家居用品']
    channels = ['线上', '线下']
    
    # 生成数据
    data = []
    for date in dates:
        for region in regions:
            for category in categories:
                for channel in channels:
                    # 基础销售额
                    base_amount = np.random.normal(50000, 15000)
                    
                    # 添加季节性因素
                    month = date.month
                    seasonal_factor = 1.0
                    if month in [11, 12]:  # 双十一、双十二
                        seasonal_factor = 1.5
                    elif month in [6, 7]:  # 暑期促销
                        seasonal_factor = 1.2
                    
                    # 添加增长趋势
                    months_from_start = (date.year - dates[0].year) * 12 + (date.month - dates[0].month)
                    growth_factor = 1 + (months_from_start * 0.02)  # 每月2%增长
                    
                    # 添加随机波动
                    random_factor = np.random.uniform(0.8, 1.2)
                    
                    # 计算最终金额
                    amount = max(0, base_amount * seasonal_factor * growth_factor * random_factor)
                    
                    # 计算销售数量
                    avg_price = np.random.uniform(100, 500)
                    quantity = int(amount / avg_price)
                    
                    data.append({
                        'date': date,
                        'region': region,
                        'category': category,
                        'channel': channel,
                        'sales_amount': round(amount, 2),
                        'quantity': quantity,
                        'avg_price': round(avg_price, 2)
                    })
    
    df = pd.DataFrame(data)
    
    # 保存为Excel文件
    df.to_excel('/workspace/sample_sales_data.xlsx', index=False)
    print(f"示例数据已生成：{len(df)}条记录")
    print("数据已保存为：sample_sales_data.xlsx")
    
    return df

def run_analysis_example():
    """
    运行完整的分析示例
    """
    print("\n" + "="*60)
    print("开始数据分析报告生成")
    print("="*60)
    
    # 创建示例数据
    sample_data = create_sample_data()
    
    # 创建分析器实例
    analyzer = DataAnalysisReport()
    
    # 加载数据
    print("\n1. 加载数据...")
    df = analyzer.load_data('/workspace/sample_sales_data.xlsx')
    
    if df is None:
        print("数据加载失败")
        return
    
    print(f"数据概览：")
    print(f"  - 时间范围：{df['date'].min()} 到 {df['date'].max()}")
    print(f"  - 数据维度：{df.columns.tolist()}")
    
    # 准备时间段数据
    print("\n2. 准备时间段数据...")
    target_date = df['date'].max()  # 使用最新日期作为分析基准
    analyzer.prepare_data_by_periods(df, date_column='date', target_date=target_date)
    
    print(f"分析基准日期：{target_date.strftime('%Y-%m-%d')}")
    print(f"上月对比日期：{analyzer.last_month_date.strftime('%Y-%m-%d')}")
    print(f"去年同期日期：{analyzer.last_year_date.strftime('%Y-%m-%d')}")
    
    # 定义分析维度和指标
    analysis_configs = [
        {
            'name': '按地区分析',
            'dimensions': ['region'],
            'value_column': 'sales_amount',
            'aggregation': 'sum'
        },
        {
            'name': '按品类分析',
            'dimensions': ['category'],
            'value_column': 'sales_amount',
            'aggregation': 'sum'
        },
        {
            'name': '按渠道分析',
            'dimensions': ['channel'],
            'value_column': 'sales_amount',
            'aggregation': 'sum'
        },
        {
            'name': '按地区品类交叉分析',
            'dimensions': ['region', 'category'],
            'value_column': 'sales_amount',
            'aggregation': 'sum'
        },
        {
            'name': '按地区销售数量分析',
            'dimensions': ['region'],
            'value_column': 'quantity',
            'aggregation': 'sum'
        }
    ]
    
    # 执行分析
    print("\n3. 执行多维度分析...")
    analysis_results = {}
    
    for config in analysis_configs:
        print(f"  正在分析：{config['name']}")
        result = analyzer.analyze_by_dimensions(
            dimensions=config['dimensions'],
            value_column=config['value_column'],
            aggregation=config['aggregation']
        )
        analysis_results[config['name']] = result
        
        if len(result) > 0:
            print(f"    - 发现 {len(result)} 个分析项目")
            
            # 显示前几行结果
            print("    - 前3项结果：")
            display_cols = config['dimensions'] + [f"当前期_{config['value_column']}", '环比增长率(%)', '同比增长率(%)']
            available_cols = [col for col in display_cols if col in result.columns]
            print(result[available_cols].head(3).to_string(index=False))
        else:
            print("    - 无数据")
        print()
    
    # 生成汇总报告
    print("4. 生成汇总分析...")
    summary = analyzer.generate_summary_report(analysis_results, analysis_configs)
    
    # 显示汇总结果
    for dim_name, stats in summary.items():
        print(f"\n【{dim_name}】汇总统计：")
        print(f"  总记录数：{stats['总记录数']}")
        print(f"  环比增长率 - 平均：{stats['环比增长率统计']['平均值']:.2f}%，"
              f"最大：{stats['环比增长率统计']['最大值']:.2f}%，"
              f"正增长项目：{stats['环比增长率统计']['正增长项目数']}个")
        print(f"  同比增长率 - 平均：{stats['同比增长率统计']['平均值']:.2f}%，"
              f"最大：{stats['同比增长率统计']['最大值']:.2f}%，"
              f"正增长项目：{stats['同比增长率统计']['正增长项目数']}个")
    
    # 导出报告
    print("\n5. 导出分析报告...")
    output_file = f"/workspace/数据分析报告_{target_date.strftime('%Y%m%d')}.xlsx"
    analyzer.export_report(analysis_results, summary, output_file)
    
    print("\n" + "="*60)
    print("分析完成！")
    print("="*60)
    print(f"1. 原始数据：sample_sales_data.xlsx")
    print(f"2. 分析报告：数据分析报告_{target_date.strftime('%Y%m%d')}.xlsx")
    print("\n报告包含以下工作表：")
    print("  - 汇总分析：各维度的整体统计")
    for config in analysis_configs:
        print(f"  - {config['name']}_详细分析：详细的增长率分析")

def quick_analysis_template():
    """
    快速分析模板 - 适用于标准数据格式
    """
    print("\n" + "="*50)
    print("快速分析模板")
    print("="*50)
    
    # 这是一个模板，用户可以修改以下配置来适应自己的数据
    config_template = """
# 快速分析配置模板
# 请根据你的数据修改以下配置

data_config = {
    'file_path': 'your_data.xlsx',      # 你的数据文件路径
    'date_column': 'date',              # 日期列名
    'target_date': None,                # 分析目标日期，None表示使用最新日期
}

analysis_config = [
    {
        'name': '按区域分析销售额',
        'dimensions': ['region'],       # 改为你的地区字段名
        'value_column': 'amount',       # 改为你的金额字段名
        'aggregation': 'sum'            # 聚合方式：sum/mean/count/max/min
    },
    {
        'name': '按产品分析销售额',
        'dimensions': ['product'],      # 改为你的产品字段名
        'value_column': 'amount',
        'aggregation': 'sum'
    },
    # 可以添加更多分析维度...
]

# 使用方法：
from data_analysis_report import DataAnalysisReport

analyzer = DataAnalysisReport()
df = analyzer.load_data(data_config['file_path'])
analyzer.prepare_data_by_periods(df, data_config['date_column'], data_config['target_date'])

results = {}
for config in analysis_config:
    results[config['name']] = analyzer.analyze_by_dimensions(
        config['dimensions'], config['value_column'], config['aggregation']
    )

summary = analyzer.generate_summary_report(results, analysis_config)
analyzer.export_report(results, summary, 'analysis_report.xlsx')
"""
    
    print("将以下代码复制并修改配置即可快速使用：")
    print(config_template)

if __name__ == "__main__":
    # 运行完整示例
    run_analysis_example()
    
    # 显示快速使用模板
    quick_analysis_template()
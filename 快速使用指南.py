#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据分析报告工具 - 快速使用指南

这个脚本演示了如何使用自动化数据分析工具来替代Excel的复杂透视表操作
一键生成包含环比、同比增长率的多维度分析报告
"""

from data_analysis_report import DataAnalysisReport
import pandas as pd

def 快速使用示例():
    """
    最简单的使用方法 - 仅需要修改几个参数即可
    """
    print("🚀 快速使用示例")
    print("="*50)
    
    # ⭐ 步骤1：修改这里的配置以匹配你的数据
    config = {
        'data_file': 'sample_sales_data.xlsx',  # 👈 改为你的数据文件名
        'date_column': 'date',                  # 👈 改为你的日期列名
        'dimensions': [                         # 👈 改为你的分析维度
            'region',      # 地区
            'category',    # 品类
            'channel'      # 渠道
        ],
        'value_column': 'sales_amount',         # 👈 改为你的分析指标列名
        'output_file': '我的分析报告.xlsx'      # 👈 改为你想要的输出文件名
    }
    
    try:
        # 步骤2：创建分析器并加载数据
        analyzer = DataAnalysisReport()
        df = analyzer.load_data(config['data_file'])
        
        if df is None:
            print("❌ 数据加载失败，请检查文件路径和格式")
            return
        
        # 步骤3：准备时间对比数据
        analyzer.prepare_data_by_periods(df, config['date_column'])
        
        # 步骤4：执行分析
        results = {}
        for dim in config['dimensions']:
            name = f"按{dim}分析"
            print(f"📊 正在执行：{name}")
            results[name] = analyzer.analyze_by_dimensions(
                dimensions=[dim],
                value_column=config['value_column'],
                aggregation='sum'
            )
        
        # 步骤5：生成报告
        summary = analyzer.generate_summary_report(results, [])
        analyzer.export_report(results, summary, config['output_file'])
        
        print(f"✅ 分析完成！报告已保存为：{config['output_file']}")
        
    except Exception as e:
        print(f"❌ 分析过程中出现错误：{str(e)}")

def 高级使用示例():
    """
    高级使用方法 - 支持多种分析配置
    """
    print("\n🔧 高级使用示例")
    print("="*50)
    
    # 创建分析器
    analyzer = DataAnalysisReport()
    
    # 加载数据
    df = analyzer.load_data('sample_sales_data.xlsx')
    analyzer.prepare_data_by_periods(df, 'date')
    
    # 定义多种分析配置
    analysis_configs = [
        {
            'name': '销售额按地区分析',
            'dimensions': ['region'],
            'value_column': 'sales_amount',
            'aggregation': 'sum'
        },
        {
            'name': '销售额按品类分析',
            'dimensions': ['category'],
            'value_column': 'sales_amount',
            'aggregation': 'sum'
        },
        {
            'name': '销售数量按地区分析',
            'dimensions': ['region'],
            'value_column': 'quantity',
            'aggregation': 'sum'
        },
        {
            'name': '平均单价按品类分析',
            'dimensions': ['category'],
            'value_column': 'avg_price',
            'aggregation': 'mean'
        },
        {
            'name': '地区×品类交叉分析',
            'dimensions': ['region', 'category'],
            'value_column': 'sales_amount',
            'aggregation': 'sum'
        }
    ]
    
    # 执行所有分析
    results = {}
    for config in analysis_configs:
        print(f"📈 正在执行：{config['name']}")
        results[config['name']] = analyzer.analyze_by_dimensions(
            config['dimensions'],
            config['value_column'],
            config['aggregation']
        )
    
    # 生成并导出报告
    summary = analyzer.generate_summary_report(results, analysis_configs)
    analyzer.export_report(results, summary, '高级分析报告.xlsx')
    print("✅ 高级分析完成！报告已保存为：高级分析报告.xlsx")

def 自定义分析模板():
    """
    提供一个空白模板，用户可以直接修改使用
    """
    template = '''
# 🎯 自定义分析模板 - 复制以下代码并修改配置

from data_analysis_report import DataAnalysisReport

def 我的分析():
    # 1. 配置你的数据信息
    data_config = {
        'file_path': 'your_data.xlsx',      # 👈 修改：你的数据文件路径
        'date_column': 'date',              # 👈 修改：日期列名
        'target_date': None,                # 👈 可选：指定分析日期（None=最新）
    }
    
    # 2. 配置分析维度和指标
    analysis_configs = [
        {
            'name': '按区域分析销售额',         # 👈 修改：分析名称
            'dimensions': ['region'],        # 👈 修改：分析维度列名
            'value_column': 'amount',        # 👈 修改：分析指标列名
            'aggregation': 'sum'             # 👈 修改：聚合方式(sum/mean/count/max/min)
        },
        {
            'name': '按产品分析销售额',
            'dimensions': ['product'],
            'value_column': 'amount',
            'aggregation': 'sum'
        },
        # 👆 可以添加更多分析配置...
    ]
    
    # 3. 执行分析（无需修改）
    analyzer = DataAnalysisReport()
    df = analyzer.load_data(data_config['file_path'])
    analyzer.prepare_data_by_periods(
        df, 
        data_config['date_column'], 
        data_config['target_date']
    )
    
    results = {}
    for config in analysis_configs:
        results[config['name']] = analyzer.analyze_by_dimensions(
            config['dimensions'], 
            config['value_column'], 
            config['aggregation']
        )
    
    summary = analyzer.generate_summary_report(results, analysis_configs)
    analyzer.export_report(results, summary, 'my_analysis_report.xlsx')
    print("✅ 分析完成！")

# 运行分析
if __name__ == "__main__":
    我的分析()
'''
    
    print("\n📝 自定义分析模板")
    print("="*50)
    print("复制以下代码到新文件中，修改配置后运行：")
    print(template)

def 常见问题解答():
    """
    常见问题和解决方案
    """
    qa = [
        {
            'Q': '我的数据没有某个时间点的数据怎么办？',
            'A': '工具会自动处理缺失数据，对应的增长率会显示为空值，不影响其他数据的分析。'
        },
        {
            'Q': '如何分析多个指标？',
            'A': '为每个指标创建单独的分析配置，或者运行多次analysis_by_dimensions。'
        },
        {
            'Q': '支持哪些聚合方式？',
            'A': 'sum(求和)、mean(平均)、count(计数)、max(最大)、min(最小)。'
        },
        {
            'Q': '可以自定义对比时间点吗？',
            'A': '可以，在prepare_data_by_periods中指定target_date参数。'
        },
        {
            'Q': '如何处理非数值型数据？',
            'A': '可以使用count聚合方式来统计数量，或者先预处理数据转换为数值。'
        }
    ]
    
    print("\n❓ 常见问题解答")
    print("="*50)
    for i, item in enumerate(qa, 1):
        print(f"{i}. {item['Q']}")
        print(f"   答：{item['A']}\n")

def main():
    """
    主函数 - 运行所有示例
    """
    print("🎉 欢迎使用自动化数据分析报告工具")
    print("这个工具可以帮你告别Excel复杂的透视表操作，一键生成专业分析报告！")
    print()
    
    # 运行快速示例
    快速使用示例()
    
    # 运行高级示例
    高级使用示例()
    
    # 显示自定义模板
    自定义分析模板()
    
    # 显示常见问题
    常见问题解答()
    
    print("🎯 使用建议：")
    print("1. 首先运行快速示例，了解基本功能")
    print("2. 查看生成的Excel报告，理解输出格式")
    print("3. 复制自定义模板，修改配置适应你的数据")
    print("4. 遇到问题时参考常见问题解答")

if __name__ == "__main__":
    main()
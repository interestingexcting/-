#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化数据分析报告生成工具
支持多时间点数据对比分析，生成包含环比和同比增长的报告
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import warnings
warnings.filterwarnings('ignore')

class DataAnalysisReport:
    def __init__(self):
        self.data = {}
        self.report_date = None
        self.last_month_date = None
        self.last_year_date = None
        
    def load_data(self, file_path, date_column='date', **kwargs):
        """
        加载Excel或CSV数据
        
        Parameters:
        file_path: 数据文件路径
        date_column: 日期列名称
        **kwargs: pandas读取参数
        """
        try:
            if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                df = pd.read_excel(file_path, **kwargs)
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path, **kwargs)
            else:
                raise ValueError("仅支持Excel(.xlsx, .xls)和CSV文件")
            
            # 确保日期列为datetime类型
            if date_column in df.columns:
                df[date_column] = pd.to_datetime(df[date_column])
            
            print(f"成功加载数据：{len(df)}行，{len(df.columns)}列")
            return df
            
        except Exception as e:
            print(f"数据加载失败：{str(e)}")
            return None
    
    def prepare_data_by_periods(self, df, date_column='date', target_date=None):
        """
        按时间段准备数据（当前期、上月、去年同期）
        
        Parameters:
        df: 数据DataFrame
        date_column: 日期列名
        target_date: 目标分析日期，默认为数据中最新日期
        """
        if target_date is None:
            target_date = df[date_column].max()
        
        self.report_date = target_date
        
        # 计算对比时间点
        self.last_month_date = target_date - pd.DateOffset(months=1)
        self.last_year_date = target_date - pd.DateOffset(years=1)
        
        # 按时间段筛选数据
        current_data = df[df[date_column] == target_date].copy()
        last_month_data = df[df[date_column] == self.last_month_date].copy()
        last_year_data = df[df[date_column] == self.last_year_date].copy()
        
        self.data = {
            'current': current_data,
            'last_month': last_month_data,
            'last_year': last_year_data
        }
        
        print(f"当前期数据：{len(current_data)}条")
        print(f"上月数据：{len(last_month_data)}条")
        print(f"去年同期数据：{len(last_year_data)}条")
        
        return self.data
    
    def calculate_growth_rates(self, current_value, last_month_value, last_year_value):
        """
        计算环比和同比增长率
        """
        # 环比增长率
        if pd.isna(last_month_value) or last_month_value == 0:
            mom_growth = np.nan
        else:
            mom_growth = (current_value - last_month_value) / last_month_value * 100
        
        # 同比增长率
        if pd.isna(last_year_value) or last_year_value == 0:
            yoy_growth = np.nan
        else:
            yoy_growth = (current_value - last_year_value) / last_year_value * 100
        
        return mom_growth, yoy_growth
    
    def analyze_by_dimensions(self, dimensions, value_column, aggregation='sum'):
        """
        按指定维度进行分析
        
        Parameters:
        dimensions: 分析维度列表，如['region', 'category']
        value_column: 分析指标列名
        aggregation: 聚合方式，支持 'sum', 'mean', 'count', 'max', 'min'
        """
        if not self.data:
            raise ValueError("请先使用prepare_data_by_periods准备数据")
        
        results = []
        
        # 对每个时间段进行聚合
        agg_data = {}
        for period, df in self.data.items():
            if len(df) > 0:
                if aggregation == 'sum':
                    agg_df = df.groupby(dimensions)[value_column].sum().reset_index()
                elif aggregation == 'mean':
                    agg_df = df.groupby(dimensions)[value_column].mean().reset_index()
                elif aggregation == 'count':
                    agg_df = df.groupby(dimensions)[value_column].count().reset_index()
                elif aggregation == 'max':
                    agg_df = df.groupby(dimensions)[value_column].max().reset_index()
                elif aggregation == 'min':
                    agg_df = df.groupby(dimensions)[value_column].min().reset_index()
                else:
                    raise ValueError("聚合方式必须是: sum, mean, count, max, min")
                
                agg_data[period] = agg_df
            else:
                # 创建空的DataFrame保持结构一致
                agg_data[period] = pd.DataFrame(columns=dimensions + [value_column])
        
        # 合并数据并计算增长率
        if len(agg_data['current']) > 0:
            # 以当前期数据为主表
            result_df = agg_data['current'].copy()
            result_df = result_df.rename(columns={value_column: f'当前期_{value_column}'})
            
            # 合并上月数据
            if len(agg_data['last_month']) > 0:
                last_month_df = agg_data['last_month'].rename(columns={value_column: f'上月_{value_column}'})
                result_df = result_df.merge(last_month_df, on=dimensions, how='left')
            else:
                result_df[f'上月_{value_column}'] = np.nan
            
            # 合并去年同期数据
            if len(agg_data['last_year']) > 0:
                last_year_df = agg_data['last_year'].rename(columns={value_column: f'去年同期_{value_column}'})
                result_df = result_df.merge(last_year_df, on=dimensions, how='left')
            else:
                result_df[f'去年同期_{value_column}'] = np.nan
            
            # 计算增长率
            result_df['环比增长率(%)'] = result_df.apply(
                lambda row: self.calculate_growth_rates(
                    row[f'当前期_{value_column}'], 
                    row[f'上月_{value_column}'], 
                    row[f'去年同期_{value_column}']
                )[0], axis=1
            )
            
            result_df['同比增长率(%)'] = result_df.apply(
                lambda row: self.calculate_growth_rates(
                    row[f'当前期_{value_column}'], 
                    row[f'上月_{value_column}'], 
                    row[f'去年同期_{value_column}']
                )[1], axis=1
            )
            
            # 格式化数值
            for col in result_df.columns:
                if '增长率' in col:
                    result_df[col] = result_df[col].round(2)
            
            return result_df
        else:
            print("当前期无数据")
            return pd.DataFrame()
    
    def generate_summary_report(self, analysis_results, dimensions):
        """
        生成汇总分析报告
        """
        summary = {}
        
        for dim_name, result_df in analysis_results.items():
            if len(result_df) > 0:
                summary[dim_name] = {
                    '总记录数': len(result_df),
                    '环比增长率统计': {
                        '平均值': result_df['环比增长率(%)'].mean(),
                        '最大值': result_df['环比增长率(%)'].max(),
                        '最小值': result_df['环比增长率(%)'].min(),
                        '正增长项目数': (result_df['环比增长率(%)'] > 0).sum(),
                        '负增长项目数': (result_df['环比增长率(%)'] < 0).sum()
                    },
                    '同比增长率统计': {
                        '平均值': result_df['同比增长率(%)'].mean(),
                        '最大值': result_df['同比增长率(%)'].max(),
                        '最小值': result_df['同比增长率(%)'].min(),
                        '正增长项目数': (result_df['同比增长率(%)'] > 0).sum(),
                        '负增长项目数': (result_df['同比增长率(%)'] < 0).sum()
                    }
                }
        
        return summary
    
    def export_report(self, analysis_results, summary, output_path):
        """
        导出分析报告到Excel文件
        """
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 写入汇总页
            summary_rows = []
            for dim_name, stats in summary.items():
                summary_rows.append({
                    '分析维度': dim_name,
                    '总记录数': stats['总记录数'],
                    '环比平均增长率(%)': round(stats['环比增长率统计']['平均值'], 2),
                    '环比最大增长率(%)': round(stats['环比增长率统计']['最大值'], 2),
                    '环比最小增长率(%)': round(stats['环比增长率统计']['最小值'], 2),
                    '环比正增长项目数': stats['环比增长率统计']['正增长项目数'],
                    '同比平均增长率(%)': round(stats['同比增长率统计']['平均值'], 2),
                    '同比最大增长率(%)': round(stats['同比增长率统计']['最大值'], 2),
                    '同比最小增长率(%)': round(stats['同比增长率统计']['最小值'], 2),
                    '同比正增长项目数': stats['同比增长率统计']['正增长项目数']
                })
            
            summary_df = pd.DataFrame(summary_rows)
            summary_df.to_excel(writer, sheet_name='汇总分析', index=False)
            
            # 写入详细分析结果
            for dim_name, result_df in analysis_results.items():
                sheet_name = f'{dim_name}_详细分析'[:31]  # Excel工作表名称限制31字符
                result_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"报告已导出至：{output_path}")

def main():
    """
    主函数 - 演示如何使用数据分析报告工具
    """
    print("=" * 60)
    print("自动化数据分析报告生成工具")
    print("=" * 60)
    
    # 创建分析器实例
    analyzer = DataAnalysisReport()
    
    # 这里需要替换为你的实际数据文件路径
    # data_file = "your_data.xlsx"  # 请替换为实际文件路径
    
    print("\n使用说明：")
    print("1. 请将你的数据文件放在工作目录中")
    print("2. 数据文件应包含日期列和分析指标列")
    print("3. 修改下面的配置参数以匹配你的数据结构")
    print("4. 运行脚本生成自动化报告")
    
    # 配置参数（请根据实际数据修改）
    config = {
        'date_column': 'date',        # 日期列名
        'value_column': 'amount',     # 分析指标列名
        'dimensions': [               # 分析维度
            ['region'],               # 按地区分析
            ['category'],             # 按类别分析
            ['region', 'category']    # 按地区和类别交叉分析
        ],
        'aggregation': 'sum'         # 聚合方式
    }
    
    print(f"\n当前配置：")
    for key, value in config.items():
        print(f"  {key}: {value}")

if __name__ == "__main__":
    main()
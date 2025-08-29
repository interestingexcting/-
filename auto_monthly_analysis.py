#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
全自动月度数据分析脚本

功能描述：
1. 自动识别Excel文件中的维度列（文本类型）和指标列（数字类型）
2. 基于识别的列自动进行月度对比分析
3. 计算环比增幅并生成分析报告
4. 全程无需手动指定列名，实现完全自动化

作者：AI Assistant
创建时间：2024年
"""

import pandas as pd
import numpy as np
import warnings
from pathlib import Path
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 忽略警告信息
warnings.filterwarnings('ignore')


class AutoMonthlyAnalyzer:
    """全自动月度数据分析器"""
    
    def __init__(self, exclude_columns=None):
        """
        初始化分析器
        
        Args:
            exclude_columns (list): 需要排除分析的列名列表
                                  例如：['客户ID', '订单号', '产品编码', '日期']
        """
        # 默认排除的列（通常这些列不适合做分析维度或指标）
        self.default_exclude = [
            '客户ID', '客户id', 'customer_id', 'cust_id',
            '订单号', '订单编号', 'order_id', 'order_no',
            '产品编码', '产品ID', 'product_id', 'prod_id',
            '日期', '时间', 'date', 'time', 'datetime',
            'ID', 'id', 'Id', 'iD'
        ]
        
        # 用户自定义排除列
        self.exclude_columns = exclude_columns or []
        
        # 合并排除列表（转换为小写便于比较）
        all_exclude = self.default_exclude + self.exclude_columns
        self.exclude_set = set([col.lower() for col in all_exclude])
        
        # 存储识别结果
        self.dimension_columns = []  # 维度列
        self.metric_columns = []     # 指标列
        
        logger.info(f"分析器初始化完成，排除列: {all_exclude}")
    
    def load_data(self, current_month_file, previous_month_file):
        """
        加载本月和上月的Excel数据
        
        Args:
            current_month_file (str): 本月数据文件路径
            previous_month_file (str): 上月数据文件路径
            
        Returns:
            tuple: (本月DataFrame, 上月DataFrame)
        """
        try:
            logger.info("开始读取Excel文件...")
            
            # 读取本月数据
            current_df = pd.read_excel(current_month_file)
            logger.info(f"本月数据读取成功: {current_month_file} ({current_df.shape[0]}行 x {current_df.shape[1]}列)")
            
            # 读取上月数据
            previous_df = pd.read_excel(previous_month_file)
            logger.info(f"上月数据读取成功: {previous_month_file} ({previous_df.shape[0]}行 x {previous_df.shape[1]}列)")
            
            return current_df, previous_df
            
        except Exception as e:
            logger.error(f"读取Excel文件失败: {e}")
            raise
    
    def identify_column_types(self, df):
        """
        自动识别DataFrame中的维度列和指标列
        
        Args:
            df (pd.DataFrame): 需要分析的DataFrame
            
        Returns:
            tuple: (维度列列表, 指标列列表)
        """
        logger.info("开始自动识别列类型...")
        
        dimension_cols = []
        metric_cols = []
        
        for col in df.columns:
            # 检查是否在排除列表中
            if col.lower() in self.exclude_set:
                logger.info(f"跳过排除列: {col}")
                continue
            
            # 获取列的数据类型
            dtype = df[col].dtype
            
            # 判断列类型
            if dtype in ['object', 'category'] or pd.api.types.is_string_dtype(dtype):
                # 文本类型 -> 维度列
                dimension_cols.append(col)
                logger.info(f"识别为维度列: {col} (类型: {dtype})")
                
            elif pd.api.types.is_numeric_dtype(dtype):
                # 数字类型 -> 指标列
                metric_cols.append(col)
                logger.info(f"识别为指标列: {col} (类型: {dtype})")
                
            else:
                logger.warning(f"未知列类型，跳过: {col} (类型: {dtype})")
        
        # 存储识别结果
        self.dimension_columns = dimension_cols
        self.metric_columns = metric_cols
        
        logger.info(f"列类型识别完成 - 维度列: {len(dimension_cols)}个, 指标列: {len(metric_cols)}个")
        
        return dimension_cols, metric_cols
    
    def aggregate_data(self, df, dimension_cols, metric_cols):
        """
        根据维度列分组，对指标列进行聚合
        
        Args:
            df (pd.DataFrame): 原始数据
            dimension_cols (list): 维度列列表
            metric_cols (list): 指标列列表
            
        Returns:
            pd.DataFrame: 聚合后的数据
        """
        if not dimension_cols:
            logger.warning("没有维度列，无法进行分组聚合")
            # 如果没有维度列，直接对指标列求和
            result = df[metric_cols].sum().to_frame().T
            return result
        
        if not metric_cols:
            logger.warning("没有指标列，无法进行聚合")
            return pd.DataFrame()
        
        logger.info(f"开始分组聚合 - 分组维度: {dimension_cols}, 聚合指标: {metric_cols}")
        
        # 按维度列分组，对指标列求和
        aggregated = df.groupby(dimension_cols)[metric_cols].sum().reset_index()
        
        logger.info(f"聚合完成，结果: {aggregated.shape[0]}行 x {aggregated.shape[1]}列")
        
        return aggregated
    
    def merge_monthly_data(self, current_agg, previous_agg, dimension_cols):
        """
        合并本月和上月的聚合数据，生成宽表
        
        Args:
            current_agg (pd.DataFrame): 本月聚合数据
            previous_agg (pd.DataFrame): 上月聚合数据
            dimension_cols (list): 维度列列表
            
        Returns:
            pd.DataFrame: 合并后的宽表
        """
        logger.info("开始合并本月和上月数据...")
        
        if dimension_cols:
            # 有维度列的情况下进行外连接
            merged = pd.merge(
                current_agg, 
                previous_agg, 
                on=dimension_cols, 
                how='outer', 
                suffixes=('_本月', '_上月')
            )
        else:
            # 没有维度列的情况下直接横向拼接
            current_agg.columns = [f"{col}_本月" for col in current_agg.columns]
            previous_agg.columns = [f"{col}_上月" for col in previous_agg.columns]
            merged = pd.concat([current_agg, previous_agg], axis=1)
        
        # 用0填充缺失值（表示该维度组合在某月没有数据）
        merged = merged.fillna(0)
        
        logger.info(f"数据合并完成，最终结果: {merged.shape[0]}行 x {merged.shape[1]}列")
        
        return merged
    
    def calculate_growth_rate(self, merged_df, metric_cols):
        """
        计算所有指标的环比增幅
        
        Args:
            merged_df (pd.DataFrame): 合并后的数据
            metric_cols (list): 指标列列表
            
        Returns:
            pd.DataFrame: 包含环比增幅的数据
        """
        logger.info("开始计算环比增幅...")
        
        result_df = merged_df.copy()
        
        # 为每个指标计算环比增幅
        for metric in metric_cols:
            current_col = f"{metric}_本月"
            previous_col = f"{metric}_上月"
            growth_col = f"{metric}_环比增幅"
            
            if current_col in result_df.columns and previous_col in result_df.columns:
                # 计算环比增幅: (本月 - 上月) / 上月
                # 使用numpy的divide函数处理除零情况
                growth_rate = np.divide(
                    result_df[current_col] - result_df[previous_col],
                    result_df[previous_col],
                    out=np.full_like(result_df[current_col], np.inf, dtype=float),
                    where=(result_df[previous_col] != 0)
                )
                
                # 将无穷大的值标记为特殊情况
                growth_rate = np.where(
                    (result_df[previous_col] == 0) & (result_df[current_col] > 0),
                    np.inf,  # 上月为0，本月大于0 -> 无穷大增长
                    growth_rate
                )
                
                growth_rate = np.where(
                    (result_df[previous_col] == 0) & (result_df[current_col] == 0),
                    0,  # 上月为0，本月也为0 -> 无变化
                    growth_rate
                )
                
                result_df[growth_col] = growth_rate
                
                logger.info(f"完成指标 {metric} 的环比增幅计算")
        
        logger.info("环比增幅计算完成")
        
        return result_df
    
    def format_results(self, result_df, metric_cols):
        """
        格式化最终结果，将增幅转换为百分比格式
        
        Args:
            result_df (pd.DataFrame): 计算结果数据
            metric_cols (list): 指标列列表
            
        Returns:
            pd.DataFrame: 格式化后的结果
        """
        logger.info("开始格式化结果...")
        
        formatted_df = result_df.copy()
        
        # 格式化环比增幅列为百分比
        for metric in metric_cols:
            growth_col = f"{metric}_环比增幅"
            if growth_col in formatted_df.columns:
                # 创建格式化的百分比列
                formatted_col = f"{metric}_环比增幅_百分比"
                
                # 处理特殊值
                def format_percentage(value):
                    if pd.isna(value):
                        return "N/A"
                    elif np.isinf(value):
                        return "∞" if value > 0 else "-∞"
                    else:
                        return f"{value:.2%}"
                
                formatted_df[formatted_col] = formatted_df[growth_col].apply(format_percentage)
        
        logger.info("结果格式化完成")
        
        return formatted_df
    
    def save_report(self, formatted_df, output_file="月度分析报告_自动版.xlsx"):
        """
        保存分析报告到Excel文件
        
        Args:
            formatted_df (pd.DataFrame): 格式化后的分析结果
            output_file (str): 输出文件路径
        """
        try:
            logger.info(f"开始保存分析报告到: {output_file}")
            
            # 创建Excel写入器
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # 保存主要分析结果
                formatted_df.to_excel(writer, sheet_name='月度分析报告', index=False)
                
                # 创建汇总信息表
                summary_data = {
                    '分析项目': ['总维度组合数', '分析维度列数', '分析指标列数', '排除列数'],
                    '数值': [
                        len(formatted_df),
                        len(self.dimension_columns),
                        len(self.metric_columns),
                        len(self.exclude_columns) + len(self.default_exclude)
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='分析汇总', index=False)
                
                # 创建列信息表
                column_info = []
                for col in self.dimension_columns:
                    column_info.append({'列名': col, '类型': '维度列', '说明': '用于分组的维度'})
                for col in self.metric_columns:
                    column_info.append({'列名': col, '类型': '指标列', '说明': '用于聚合的数值指标'})
                
                if column_info:
                    column_df = pd.DataFrame(column_info)
                    column_df.to_excel(writer, sheet_name='列信息', index=False)
            
            logger.info(f"分析报告保存成功: {output_file}")
            
        except Exception as e:
            logger.error(f"保存报告失败: {e}")
            raise
    
    def run_analysis(self, current_month_file="本月数据.xlsx", previous_month_file="上月数据.xlsx"):
        """
        执行完整的自动化分析流程
        
        Args:
            current_month_file (str): 本月数据文件路径
            previous_month_file (str): 上月数据文件路径
            
        Returns:
            pd.DataFrame: 最终的分析结果
        """
        try:
            logger.info("="*60)
            logger.info("开始执行全自动月度数据分析")
            logger.info("="*60)
            
            # 1. 加载数据
            current_df, previous_df = self.load_data(current_month_file, previous_month_file)
            
            # 2. 自动识别列类型（基于本月数据的结构）
            dimension_cols, metric_cols = self.identify_column_types(current_df)
            
            if not metric_cols:
                raise ValueError("未识别到任何指标列（数字类型列），无法进行分析")
            
            # 3. 分别聚合本月和上月数据
            logger.info("-" * 40)
            logger.info("聚合本月数据...")
            current_agg = self.aggregate_data(current_df, dimension_cols, metric_cols)
            
            logger.info("聚合上月数据...")
            previous_agg = self.aggregate_data(previous_df, dimension_cols, metric_cols)
            
            # 4. 合并数据
            logger.info("-" * 40)
            merged_df = self.merge_monthly_data(current_agg, previous_agg, dimension_cols)
            
            # 5. 计算环比增幅
            logger.info("-" * 40)
            result_df = self.calculate_growth_rate(merged_df, metric_cols)
            
            # 6. 格式化结果
            logger.info("-" * 40)
            formatted_df = self.format_results(result_df, metric_cols)
            
            # 7. 保存报告
            logger.info("-" * 40)
            self.save_report(formatted_df)
            
            # 8. 显示分析汇总
            self.print_analysis_summary(formatted_df)
            
            logger.info("="*60)
            logger.info("全自动月度数据分析完成！")
            logger.info("="*60)
            
            return formatted_df
            
        except Exception as e:
            logger.error(f"分析过程中发生错误: {e}")
            raise
    
    def print_analysis_summary(self, result_df):
        """打印分析汇总信息"""
        print("\n" + "="*50)
        print("📊 分析汇总报告")
        print("="*50)
        print(f"🔍 识别到的维度列 ({len(self.dimension_columns)}个): {', '.join(self.dimension_columns)}")
        print(f"📈 识别到的指标列 ({len(self.metric_columns)}个): {', '.join(self.metric_columns)}")
        print(f"📋 分析的维度组合数: {len(result_df)}个")
        print(f"💾 报告已保存到: 月度分析报告_自动版.xlsx")
        
        # 显示前几行结果预览
        if len(result_df) > 0:
            print(f"\n📋 结果预览 (前5行):")
            print(result_df.head().to_string())
        
        print("="*50)


def main():
    """主函数 - 执行自动化分析"""
    
    # ===== 配置区域 =====
    # 您可以在这里自定义排除的列名
    EXCLUDE_COLUMNS = [
        # 添加您想要排除的列名，例如：
        # '客户编号',
        # '流水号', 
        # '备注',
        # '创建时间'
    ]
    
    # Excel文件路径
    CURRENT_MONTH_FILE = "本月数据.xlsx"
    PREVIOUS_MONTH_FILE = "上月数据.xlsx"
    
    # ===== 执行分析 =====
    try:
        # 创建分析器实例
        analyzer = AutoMonthlyAnalyzer(exclude_columns=EXCLUDE_COLUMNS)
        
        # 检查文件是否存在
        current_path = Path(CURRENT_MONTH_FILE)
        previous_path = Path(PREVIOUS_MONTH_FILE)
        
        if not current_path.exists():
            print(f"❌ 错误: 找不到本月数据文件 '{CURRENT_MONTH_FILE}'")
            print("请确保文件存在于当前目录中")
            return
            
        if not previous_path.exists():
            print(f"❌ 错误: 找不到上月数据文件 '{PREVIOUS_MONTH_FILE}'")
            print("请确保文件存在于当前目录中")
            return
        
        # 执行自动化分析
        result = analyzer.run_analysis(CURRENT_MONTH_FILE, PREVIOUS_MONTH_FILE)
        
        print("\n✅ 分析完成！请查看生成的Excel报告文件。")
        
    except Exception as e:
        print(f"❌ 分析失败: {e}")
        print("请检查Excel文件格式和数据内容是否正确")


if __name__ == "__main__":
    main()
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交互式Excel数据分析工具
功能：智能分析Excel文件，生成数据对比汇总报告

作者：Python数据自动化专家
版本：1.0
"""

import pandas as pd
import os
import sys
from datetime import datetime
from typing import List, Dict, Tuple, Optional


class ExcelDataAnalyzer:
    """Excel数据分析器类"""
    
    def __init__(self):
        """初始化分析器"""
        self.current_month_data = None
        self.previous_month_data = None
        self.dimension_columns = []
        self.metric_columns = []
        
    def validate_date_format(self, date_str: str) -> bool:
        """
        验证日期格式是否为YYYY-MM-DD
        
        Args:
            date_str: 日期字符串
            
        Returns:
            bool: 格式是否正确
        """
        try:
            datetime.strptime(date_str.strip(), '%Y-%m-%d')
            return True
        except ValueError:
            return False
    
    def generate_filename(self, date_str: str) -> str:
        """
        根据日期生成文件名
        
        Args:
            date_str: 日期字符串（YYYY-MM-DD）
            
        Returns:
            str: 生成的文件名
        """
        return f"数据_{date_str.strip()}.xlsx"
    
    def check_file_exists(self, filename: str) -> bool:
        """
        检查文件是否存在
        
        Args:
            filename: 文件名
            
        Returns:
            bool: 文件是否存在
        """
        return os.path.exists(filename)
    
    def load_excel_data(self, filename: str) -> Optional[pd.DataFrame]:
        """
        加载Excel文件数据
        
        Args:
            filename: Excel文件路径
            
        Returns:
            Optional[pd.DataFrame]: 数据DataFrame，如果失败返回None
        """
        try:
            # 尝试读取Excel文件的第一个工作表
            df = pd.read_excel(filename, sheet_name=0)
            print(f"✓ 成功读取文件: {filename}")
            print(f"  数据形状: {df.shape}")
            return df
        except Exception as e:
            print(f"✗ 读取文件失败: {filename}")
            print(f"  错误信息: {str(e)}")
            return None
    
    def analyze_columns(self, df: pd.DataFrame) -> Tuple[List[str], List[str]]:
        """
        智能分析列名，区分维度列和指标列
        
        Args:
            df: DataFrame数据
            
        Returns:
            Tuple[List[str], List[str]]: (维度列列表, 指标列列表)
        """
        dimension_cols = []
        metric_cols = []
        
        for col in df.columns:
            # 获取列的数据类型
            dtype = df[col].dtype
            
            # 检查列中非空值的类型
            non_null_values = df[col].dropna()
            
            if len(non_null_values) == 0:
                # 如果列全为空，跳过
                continue
            
            # 根据数据类型和内容判断列的性质
            if dtype in ['object', 'string'] or str(dtype).startswith('string'):
                # 文本类型列，通常为维度
                dimension_cols.append(col)
            elif dtype in ['int64', 'int32', 'float64', 'float32'] or 'int' in str(dtype) or 'float' in str(dtype):
                # 数值类型列，通常为指标
                metric_cols.append(col)
            else:
                # 其他类型（如日期），尝试判断内容
                try:
                    # 尝试将前几个值转换为数字
                    sample_values = non_null_values.head(10)
                    pd.to_numeric(sample_values, errors='raise')
                    metric_cols.append(col)
                except:
                    # 无法转换为数字，归类为维度
                    dimension_cols.append(col)
        
        return dimension_cols, metric_cols
    
    def display_dimension_options(self, dimensions: List[str]) -> None:
        """
        显示可选的维度列
        
        Args:
            dimensions: 维度列列表
        """
        print("\n" + "="*60)
        print("📊 可用的分析维度：")
        print("="*60)
        
        for i, dim in enumerate(dimensions, 1):
            print(f"{i:2d}. {dim}")
        
        print("\n💡 提示：")
        print("   - 输入数字选择维度（如：1 或 1,2,3）")
        print("   - 多个维度用英文逗号分隔")
        print("   - 输入 'all' 选择所有维度")
        print("="*60)
    
    def get_user_dimension_selection(self, dimensions: List[str]) -> List[str]:
        """
        获取用户选择的维度
        
        Args:
            dimensions: 可选维度列表
            
        Returns:
            List[str]: 用户选择的维度列表
        """
        while True:
            try:
                selection = input("\n请选择分析维度（输入数字）: ").strip()
                
                if selection.lower() == 'all':
                    return dimensions
                
                # 解析用户输入
                if ',' in selection:
                    # 多个选择
                    indices = [int(x.strip()) for x in selection.split(',')]
                else:
                    # 单个选择
                    indices = [int(selection)]
                
                # 验证索引范围
                selected_dims = []
                for idx in indices:
                    if 1 <= idx <= len(dimensions):
                        selected_dims.append(dimensions[idx - 1])
                    else:
                        print(f"✗ 无效的选择: {idx}，请输入1-{len(dimensions)}之间的数字")
                        break
                else:
                    # 所有索引都有效
                    if selected_dims:
                        print(f"✓ 已选择维度: {selected_dims}")
                        return selected_dims
                
            except ValueError:
                print("✗ 输入格式错误，请输入数字（如：1 或 1,2,3）")
            except KeyboardInterrupt:
                print("\n\n程序已退出")
                sys.exit(0)
    
    def group_and_summarize(self, df: pd.DataFrame, group_by_cols: List[str], 
                          metric_cols: List[str]) -> pd.DataFrame:
        """
        按指定维度分组并汇总指标
        
        Args:
            df: 数据DataFrame
            group_by_cols: 分组维度列
            metric_cols: 指标列
            
        Returns:
            pd.DataFrame: 汇总后的数据
        """
        try:
            # 确保所有指标列都是数值类型
            df_copy = df.copy()
            for col in metric_cols:
                df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce')
            
            # 按维度分组并对指标列求和
            grouped = df_copy.groupby(group_by_cols)[metric_cols].sum().reset_index()
            
            print(f"✓ 数据分组汇总完成，共 {len(grouped)} 个分组")
            return grouped
            
        except Exception as e:
            print(f"✗ 分组汇总失败: {str(e)}")
            return pd.DataFrame()
    
    def calculate_comparison(self, current_df: pd.DataFrame, previous_df: pd.DataFrame,
                           group_by_cols: List[str], metric_cols: List[str]) -> pd.DataFrame:
        """
        计算两个月数据的对比和环比
        
        Args:
            current_df: 本月数据
            previous_df: 上月数据
            group_by_cols: 分组维度列
            metric_cols: 指标列
            
        Returns:
            pd.DataFrame: 包含对比和环比的结果
        """
        try:
            # 合并两个月的数据
            merged = pd.merge(current_df, previous_df, on=group_by_cols, 
                            how='outer', suffixes=('_本月', '_上月'))
            
            # 填充缺失值为0
            for col in metric_cols:
                merged[f'{col}_本月'] = merged[f'{col}_本月'].fillna(0)
                merged[f'{col}_上月'] = merged[f'{col}_上月'].fillna(0)
            
            # 计算环比增长率和绝对变化
            for col in metric_cols:
                current_col = f'{col}_本月'
                previous_col = f'{col}_上月'
                change_col = f'{col}_变化'
                growth_col = f'{col}_环比(%)'
                
                # 计算绝对变化
                merged[change_col] = merged[current_col] - merged[previous_col]
                
                # 计算环比增长率（处理分母为0的情况）
                merged[growth_col] = merged.apply(
                    lambda row: (
                        round((row[current_col] - row[previous_col]) / row[previous_col] * 100, 2)
                        if row[previous_col] != 0
                        else (100.0 if row[current_col] > 0 else 0.0)
                    ), axis=1
                )
            
            # 重新排列列的顺序，使结果更易读
            result_columns = group_by_cols.copy()
            for col in metric_cols:
                result_columns.extend([
                    f'{col}_上月', f'{col}_本月', f'{col}_变化', f'{col}_环比(%)'
                ])
            
            result = merged[result_columns]
            
            print(f"✓ 环比分析完成，共 {len(result)} 个维度组合")
            return result
            
        except Exception as e:
            print(f"✗ 对比分析失败: {str(e)}")
            return pd.DataFrame()
    
    def format_and_display_results(self, result_df: pd.DataFrame) -> None:
        """
        格式化并显示分析结果
        
        Args:
            result_df: 结果DataFrame
        """
        if result_df.empty:
            print("✗ 没有结果可以显示")
            return
        
        print("\n" + "="*100)
        print("📈 数据对比分析结果")
        print("="*100)
        
        # 设置pandas显示选项
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 20)
        
        # 显示结果
        print(result_df.to_string(index=False))
        
        print("\n" + "="*100)
        print("📊 数据汇总信息:")
        print(f"   总维度组合数: {len(result_df)}")
        
        # 显示环比变化统计
        growth_columns = [col for col in result_df.columns if '环比(%)' in col]
        if growth_columns:
            print(f"   环比增长指标数: {len(growth_columns)}")
            for col in growth_columns:
                positive_count = (result_df[col] > 0).sum()
                negative_count = (result_df[col] < 0).sum()
                zero_count = (result_df[col] == 0).sum()
                print(f"   {col}: 增长({positive_count}) | 下降({negative_count}) | 持平({zero_count})")
        
        print("="*100)
    
    def run(self) -> None:
        """运行主程序"""
        print("="*80)
        print("🎯 欢迎使用交互式Excel数据分析工具")
        print("="*80)
        print("功能：智能分析Excel文件，生成环比对比报告")
        print("支持：自动列识别、交互式维度选择、环比计算")
        print()
        
        try:
            # 1. 获取用户输入的日期
            print("📅 第一步：输入分析日期")
            print("-" * 40)
            
            while True:
                current_date = input("请输入本月末日期 (YYYY-MM-DD 格式，如 2023-10-31): ").strip()
                if self.validate_date_format(current_date):
                    break
                print("✗ 日期格式错误，请使用 YYYY-MM-DD 格式")
            
            while True:
                previous_date = input("请输入上月末日期 (YYYY-MM-DD 格式，如 2023-09-30): ").strip()
                if self.validate_date_format(previous_date):
                    break
                print("✗ 日期格式错误，请使用 YYYY-MM-DD 格式")
            
            # 2. 生成文件名并检查文件存在性
            print(f"\n📁 第二步：查找Excel文件")
            print("-" * 40)
            
            current_filename = self.generate_filename(current_date)
            previous_filename = self.generate_filename(previous_date)
            
            print(f"查找本月文件: {current_filename}")
            print(f"查找上月文件: {previous_filename}")
            
            if not self.check_file_exists(current_filename):
                print(f"✗ 本月文件不存在: {current_filename}")
                return
            
            if not self.check_file_exists(previous_filename):
                print(f"✗ 上月文件不存在: {previous_filename}")
                return
            
            # 3. 加载数据
            print(f"\n📊 第三步：加载数据文件")
            print("-" * 40)
            
            self.current_month_data = self.load_excel_data(current_filename)
            self.previous_month_data = self.load_excel_data(previous_filename)
            
            if self.current_month_data is None or self.previous_month_data is None:
                print("✗ 数据加载失败，程序终止")
                return
            
            # 4. 智能分析列结构
            print(f"\n🔍 第四步：智能分析数据结构")
            print("-" * 40)
            
            self.dimension_columns, self.metric_columns = self.analyze_columns(self.current_month_data)
            
            print(f"✓ 识别到 {len(self.dimension_columns)} 个维度列: {self.dimension_columns}")
            print(f"✓ 识别到 {len(self.metric_columns)} 个指标列: {self.metric_columns}")
            
            if not self.dimension_columns:
                print("✗ 未找到维度列，无法进行分析")
                return
            
            if not self.metric_columns:
                print("✗ 未找到指标列，无法进行分析")
                return
            
            # 5. 交互式选择分析维度
            print(f"\n🎯 第五步：选择分析维度")
            self.display_dimension_options(self.dimension_columns)
            selected_dimensions = self.get_user_dimension_selection(self.dimension_columns)
            
            # 6. 执行数据分析
            print(f"\n⚙️  第六步：执行数据分析")
            print("-" * 40)
            
            print("正在处理本月数据...")
            current_summary = self.group_and_summarize(
                self.current_month_data, selected_dimensions, self.metric_columns
            )
            
            print("正在处理上月数据...")
            previous_summary = self.group_and_summarize(
                self.previous_month_data, selected_dimensions, self.metric_columns
            )
            
            if current_summary.empty or previous_summary.empty:
                print("✗ 数据汇总失败，程序终止")
                return
            
            print("正在计算环比对比...")
            final_result = self.calculate_comparison(
                current_summary, previous_summary, selected_dimensions, self.metric_columns
            )
            
            # 7. 显示分析结果
            print(f"\n📈 第七步：分析结果")
            self.format_and_display_results(final_result)
            
            # 8. 询问是否保存结果
            print(f"\n💾 是否保存分析结果？")
            save_choice = input("输入 'y' 保存到Excel文件，其他任意键跳过: ").strip().lower()
            
            if save_choice == 'y':
                output_filename = f"分析结果_{current_date}_vs_{previous_date}.xlsx"
                try:
                    final_result.to_excel(output_filename, index=False)
                    print(f"✓ 结果已保存到: {output_filename}")
                except Exception as e:
                    print(f"✗ 保存失败: {str(e)}")
            
            print(f"\n🎉 分析完成！感谢使用Excel数据分析工具")
            
        except KeyboardInterrupt:
            print(f"\n\n⚠️  程序被用户中断")
        except Exception as e:
            print(f"\n✗ 程序执行出错: {str(e)}")
            print("请检查数据文件格式和内容是否正确")


def main():
    """主函数"""
    analyzer = ExcelDataAnalyzer()
    analyzer.run()


if __name__ == "__main__":
    main()
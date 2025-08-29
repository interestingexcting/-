#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建测试数据的脚本 V2.0
用于生成更丰富的测试用Excel文件，验证data_analyzer_v2.py的双模式功能
新增：更广泛的数值范围，更适合区间分析
"""

import pandas as pd
import numpy as np
from datetime import datetime
import random

def create_enhanced_test_data():
    """创建增强版测试数据"""
    
    # 设置随机种子，确保结果可重现
    np.random.seed(42)
    random.seed(42)
    
    # 定义维度数据
    product_lines = ['产品线A', '产品线B', '产品线C', '产品线D']
    regions = ['华东区', '华南区', '华北区', '西部区']
    risk_levels = ['低风险', '中风险', '高风险']
    
    def generate_varied_amounts():
        """生成具有不同数值分布的金额数据"""
        # 创建多层次的金额分布，便于区间分析
        amount_type = random.choice(['small', 'medium', 'large', 'extra_large'])
        
        if amount_type == 'small':
            return round(random.uniform(1000, 50000), 2)
        elif amount_type == 'medium':
            return round(random.uniform(50000, 200000), 2)
        elif amount_type == 'large':
            return round(random.uniform(200000, 800000), 2)
        else:  # extra_large
            return round(random.uniform(800000, 2000000), 2)
    
    def generate_loan_amounts():
        """生成贷款金额数据（新增指标）"""
        # 为了更好地测试区间功能，创建明显的分层
        tier = random.choice(['tier1', 'tier2', 'tier3', 'tier4'])
        
        if tier == 'tier1':
            return round(random.uniform(10000, 100000), 2)
        elif tier == 'tier2':
            return round(random.uniform(100000, 500000), 2)
        elif tier == 'tier3':
            return round(random.uniform(500000, 1500000), 2)
        else:  # tier4
            return round(random.uniform(1500000, 5000000), 2)
    
    # 生成上月数据（2023-09-30）
    print("🔧 生成上月数据...")
    previous_data = []
    for _ in range(200):  # 增加数据量
        record = {
            '产品线': random.choice(product_lines),
            '所属区域': random.choice(regions),
            '风险等级': random.choice(risk_levels),
            '风险金额': generate_varied_amounts(),
            '贷款金额': generate_loan_amounts(),  # 新增指标
            '风险笔数': random.randint(1, 50),
            '客户数量': random.randint(5, 100),
            '收入金额': round(random.uniform(5000, 300000), 2)  # 新增指标
        }
        previous_data.append(record)
    
    # 生成本月数据（2023-10-31），模拟一些增长趋势
    print("🔧 生成本月数据...")
    current_data = []
    for _ in range(220):  # 本月略多一些记录
        record = {
            '产品线': random.choice(product_lines),
            '所属区域': random.choice(regions),
            '风险等级': random.choice(risk_levels),
            '风险金额': generate_varied_amounts() * random.uniform(1.05, 1.15),  # 略高一些
            '贷款金额': generate_loan_amounts() * random.uniform(1.02, 1.12),  # 略高一些
            '风险笔数': random.randint(1, 55),  # 略多一些
            '客户数量': random.randint(5, 110),  # 略多一些
            '收入金额': round(random.uniform(5000, 300000) * random.uniform(1.03, 1.13), 2)  # 略高一些
        }
        current_data.append(record)
    
    # 创建DataFrame
    df_previous = pd.DataFrame(previous_data)
    df_current = pd.DataFrame(current_data)
    
    # 保存为Excel文件
    df_previous.to_excel('数据_2023-09-30.xlsx', index=False)
    df_current.to_excel('数据_2023-10-31.xlsx', index=False)
    
    print("✓ 增强版测试数据创建完成！")
    print(f"  上月数据文件: 数据_2023-09-30.xlsx ({len(df_previous)} 条记录)")
    print(f"  本月数据文件: 数据_2023-10-31.xlsx ({len(df_current)} 条记录)")
    print()
    
    # 分析数据分布
    print("📊 数据结构预览:")
    print("=" * 60)
    print("维度列（文本类型）：")
    print("  - 产品线: 4种类型")
    print("  - 所属区域: 4种类型") 
    print("  - 风险等级: 3种类型")
    print()
    print("指标列（数值类型）：")
    print("  - 风险金额: 1,000-2,300,000范围")
    print("  - 贷款金额: 10,000-5,700,000范围 [适合区间分析]")
    print("  - 风险笔数: 1-55范围")
    print("  - 客户数量: 5-110范围")
    print("  - 收入金额: 5,000-350,000范围")
    print()
    
    # 显示贷款金额的分布统计
    print("🎯 贷款金额分布统计（用于区间分析测试）：")
    print("-" * 50)
    loan_amounts = df_current['贷款金额']
    print(f"  最小值: {loan_amounts.min():,.2f}")
    print(f"  最大值: {loan_amounts.max():,.2f}")
    print(f"  中位数: {loan_amounts.median():,.2f}")
    print(f"  平均值: {loan_amounts.mean():,.2f}")
    print(f"  25分位数: {loan_amounts.quantile(0.25):,.2f}")
    print(f"  75分位数: {loan_amounts.quantile(0.75):,.2f}")
    print()
    
    print("💡 建议的区间切分点：")
    print(f"  单切分点: 500000 (将数据分为 <=50万 和 >50万)")
    print(f"  双切分点: 100000,1000000 (分为 <=10万, 10万-100万, >100万)")
    print(f"  三切分点: 100000,500000,1500000 (分为四个区间)")
    
    print("=" * 60)
    print()
    print("🚀 使用建议:")
    print("1. 运行 python3 data_analyzer_v2.py")
    print("2. 输入本月末日期: 2023-10-31")
    print("3. 输入上月末日期: 2023-09-30")
    print("4. 选择分析模式:")
    print("   - 模式1: 按维度汇总（选择产品线、区域等）")
    print("   - 模式2: 按指标区间汇总（选择贷款金额进行区间分析）")

def analyze_data_distribution():
    """分析现有数据的分布情况"""
    try:
        # 读取当前数据
        df_current = pd.read_excel('数据_2023-10-31.xlsx')
        df_previous = pd.read_excel('数据_2023-09-30.xlsx')
        
        print("\n📈 现有数据分析报告:")
        print("=" * 60)
        
        # 分析所有数值列的分布
        numeric_columns = df_current.select_dtypes(include=[np.number]).columns
        
        for col in numeric_columns:
            combined_data = pd.concat([df_current[col], df_previous[col]])
            print(f"\n{col} 分布统计:")
            print(f"  范围: {combined_data.min():,.2f} ~ {combined_data.max():,.2f}")
            print(f"  均值: {combined_data.mean():,.2f}")
            print(f"  中位数: {combined_data.median():,.2f}")
            
            # 建议切分点
            q25 = combined_data.quantile(0.25)
            q75 = combined_data.quantile(0.75)
            median = combined_data.median()
            
            print(f"  建议切分点: {median:,.0f} (中位数)")
            print(f"  或: {q25:,.0f}, {q75:,.0f} (四分位数)")
        
        print("=" * 60)
        
    except FileNotFoundError:
        print("⚠️ 未找到现有数据文件，请先运行数据生成功能")

if __name__ == "__main__":
    print("="*80)
    print("🎯 Excel数据分析工具 V2.0 - 测试数据生成器")
    print("="*80)
    print("选择操作:")
    print("1. 生成新的测试数据")
    print("2. 分析现有数据分布")
    print()
    
    choice = input("请选择操作（1或2）: ").strip()
    
    if choice == '1':
        create_enhanced_test_data()
    elif choice == '2':
        analyze_data_distribution()
    else:
        print("无效选择，正在生成新的测试数据...")
        create_enhanced_test_data()
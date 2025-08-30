#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建测试数据的脚本
用于生成测试用的Excel文件，验证data_analyzer.py的功能
"""

import pandas as pd
import numpy as np
from datetime import datetime
import random

def create_test_data():
    """创建测试数据"""
    
    # 设置随机种子，确保结果可重现
    np.random.seed(42)
    random.seed(42)
    
    # 定义维度数据
    product_lines = ['产品线A', '产品线B', '产品线C', '产品线D']
    regions = ['华东区', '华南区', '华北区', '西部区']
    risk_levels = ['低风险', '中风险', '高风险']
    
    # 生成上月数据（2023-09-30）
    previous_data = []
    for _ in range(150):  # 生成150条记录
        record = {
            '产品线': random.choice(product_lines),
            '所属区域': random.choice(regions),
            '风险等级': random.choice(risk_levels),
            '风险金额': round(random.uniform(10000, 500000), 2),
            '风险笔数': random.randint(1, 50),
            '客户数量': random.randint(5, 100)
        }
        previous_data.append(record)
    
    # 生成本月数据（2023-10-31），模拟一些增长趋势
    current_data = []
    for _ in range(160):  # 本月略多一些记录
        record = {
            '产品线': random.choice(product_lines),
            '所属区域': random.choice(regions),
            '风险等级': random.choice(risk_levels),
            '风险金额': round(random.uniform(12000, 520000), 2),  # 略高一些
            '风险笔数': random.randint(1, 55),  # 略多一些
            '客户数量': random.randint(5, 110)  # 略多一些
        }
        current_data.append(record)
    
    # 创建DataFrame
    df_previous = pd.DataFrame(previous_data)
    df_current = pd.DataFrame(current_data)
    
    # 保存为Excel文件
    df_previous.to_excel('数据_2023-09-30.xlsx', index=False)
    df_current.to_excel('数据_2023-10-31.xlsx', index=False)
    
    print("✓ 测试数据创建完成！")
    print(f"  上月数据文件: 数据_2023-09-30.xlsx ({len(df_previous)} 条记录)")
    print(f"  本月数据文件: 数据_2023-10-31.xlsx ({len(df_current)} 条记录)")
    print()
    print("📊 数据结构预览:")
    print("=" * 50)
    print("维度列（文本类型）：")
    print("  - 产品线: 4种类型")
    print("  - 所属区域: 4种类型") 
    print("  - 风险等级: 3种类型")
    print()
    print("指标列（数值类型）：")
    print("  - 风险金额: 10,000-520,000范围")
    print("  - 风险笔数: 1-55范围")
    print("  - 客户数量: 5-110范围")
    print("=" * 50)
    print()
    print("💡 使用建议:")
    print("1. 运行 python data_analyzer.py")
    print("2. 输入本月末日期: 2023-10-31")
    print("3. 输入上月末日期: 2023-09-30")
    print("4. 选择分析维度进行测试")

if __name__ == "__main__":
    create_test_data()
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据分析工具 V2.0 使用演示
展示如何使用双模式分析功能
"""

def print_demo_guide():
    """打印使用演示指南"""
    print("="*80)
    print("🎯 Excel数据分析工具 V2.0 使用演示")
    print("="*80)
    print()
    
    print("📋 演示步骤:")
    print("1. 确保已有测试数据文件")
    print("2. 运行主程序")
    print("3. 按提示体验两种分析模式")
    print()
    
    print("🔧 第一步：生成/检查测试数据")
    print("-" * 50)
    print("运行以下命令生成测试数据：")
    print("   python3 create_test_data_v2.py")
    print()
    print("这将生成：")
    print("   - 数据_2023-09-30.xlsx (上月数据, 200条记录)")
    print("   - 数据_2023-10-31.xlsx (本月数据, 220条记录)")
    print()
    
    print("📊 第二步：运行主程序")
    print("-" * 50)
    print("运行以下命令启动分析工具：")
    print("   python3 data_analyzer_v2.py")
    print()
    
    print("🎮 第三步：体验双模式分析")
    print("-" * 50)
    print()
    
    print("模式一：按维度汇总分析")
    print("   输入日期: 2023-10-31, 2023-09-30")
    print("   选择模式: 1")
    print("   选择维度: 1,2 (产品线和所属区域)")
    print("   → 查看各维度组合的指标汇总和环比变化")
    print()
    
    print("模式二：按指标区间汇总分析 🆕")
    print("   输入日期: 2023-10-31, 2023-09-30")
    print("   选择模式: 2")
    print("   选择指标: 2 (贷款金额)")
    print("   输入切分点: 500000,1500000")
    print("   → 查看不同贷款金额区间的数据分布和环比")
    print()
    
    print("💡 建议的测试场景:")
    print("-" * 50)
    print()
    
    print("场景1：产品线分析")
    print("   模式: 按维度汇总")
    print("   维度: 产品线")
    print("   目的: 了解各产品线的整体表现")
    print()
    
    print("场景2：区域+风险等级交叉分析")
    print("   模式: 按维度汇总")
    print("   维度: 所属区域,风险等级")
    print("   目的: 分析区域与风险等级的交叉分布")
    print()
    
    print("场景3：贷款金额分层分析 🆕")
    print("   模式: 按指标区间汇总")
    print("   指标: 贷款金额")
    print("   切分点: 500000 (分为小额贷款≤50万, 大额贷款>50万)")
    print("   目的: 分析不同规模贷款的风险分布")
    print()
    
    print("场景4：精细化区间分析 🆕")
    print("   模式: 按指标区间汇总")
    print("   指标: 贷款金额")
    print("   切分点: 100000,500000,1500000")
    print("   目的: 四个区间的详细分析")
    print("   区间: ≤10万, 10万-50万, 50万-150万, >150万")
    print()
    
    print("📈 预期输出:")
    print("-" * 50)
    print("两种模式都会输出:")
    print("✓ 分组汇总数据")
    print("✓ 环比对比表格")
    print("✓ 增长趋势统计")
    print("✓ 可选的Excel结果保存")
    print()
    
    print("🎯 高级技巧:")
    print("-" * 50)
    print("1. 区间切分点选择:")
    print("   - 查看程序提供的数据分布统计")
    print("   - 使用中位数作为单一切分点")
    print("   - 使用四分位数创建均衡区间")
    print("   - 根据业务需求设定特定阈值")
    print()
    print("2. 结果分析:")
    print("   - 关注环比增长率的正负变化")
    print("   - 对比不同区间/维度的表现差异")
    print("   - 识别异常值和趋势变化")
    print()
    
    print("="*80)
    print("🚀 现在开始您的数据分析之旅吧！")
    print("="*80)

if __name__ == "__main__":
    print_demo_guide()
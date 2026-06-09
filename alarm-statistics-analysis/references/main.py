#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
告警统计分析主流程

默认执行：数据清洗 → 生成 HTML 报告 → 清理中间产物
"""

import subprocess
import sys
import argparse
import os
import shutil
import glob


def run_clean(custom_periods=None):
    """执行数据清洗"""
    cmd = [sys.executable, 'data_cleaner.py']
    if custom_periods:
        cmd.extend(['--custom-periods', custom_periods])
    
    print("\n>>> 第1步：数据清洗")
    result = subprocess.run(cmd, capture_output=False)
    if result.returncode != 0:
        print("错误: 数据清洗失败")
        sys.exit(result.returncode)
    
    clean_path = os.path.join('statistics', 'cleaned_data.xlsx')
    if not os.path.exists(clean_path):
        print(f"错误: 清洗后数据 {clean_path} 未生成")
        sys.exit(1)


def run_report():
    """执行报告生成"""
    print("\n>>> 第2步：生成 HTML 报告")
    cmd = [sys.executable, 'create_combined_weekly_statistics_with_chinese_names.py',
           '--generate-report', '--no-progress']
    
    result = subprocess.run(cmd, capture_output=False)
    if result.returncode != 0:
        print("错误: 报告生成失败")
        sys.exit(result.returncode)


def cleanup():
    """删除临时文件和中间产物"""
    print("\n>>> 清理中间产物")
    base = os.path.dirname(os.path.abspath(__file__))
    
    # 清理 temp 目录
    temp_dir = os.path.join(base, 'statistics', 'temp')
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir, ignore_errors=True)
        print(f"    已删除: statistics/temp/")
    
    # 清理 cleaned_data.xlsx
    cleaned = os.path.join(base, 'statistics', 'cleaned_data.xlsx')
    if os.path.exists(cleaned):
        os.remove(cleaned)
        print(f"    已删除: statistics/cleaned_data.xlsx")
    
    # 清理 combined_weekly_chinese.xlsx
    combined = os.path.join(base, 'statistics', '告警统计分析_combined_weekly_chinese.xlsx')
    if os.path.exists(combined):
        os.remove(combined)
        print(f"    已删除: statistics/告警统计分析_combined_weekly_chinese.xlsx")
    
    # 清理旧日志（保留最新 3 个）
    log_pattern = os.path.join(base, 'logs', '*.log')
    log_files = sorted(glob.glob(log_pattern), key=os.path.getmtime, reverse=True)
    for f in log_files[3:]:
        os.remove(f)
        print(f"    已删除旧日志: {os.path.basename(f)}")


def done():
    """打印结束横幅"""
    print("\n" + "=" * 50)
    print("  流程结束")
    print("=" * 50)


def main():
    parser = argparse.ArgumentParser(
        description='告警统计分析主流程 - 数据清洗 + HTML 报告生成',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例:
    python main.py                                       # 完整流程
    python main.py --skip-clean                          # 仅生成报告
    python main.py --clean-only                          # 仅清洗数据
    python main.py --custom-periods "2025-07-01 - 2025-07-06"
        '''
    )
    parser.add_argument('--skip-clean', action='store_true',
                        help='跳过数据清洗，直接生成报告')
    parser.add_argument('--clean-only', action='store_true',
                        help='仅清洗数据，不生成报告')
    parser.add_argument('--custom-periods', type=str, default=None,
                        help='自定义周期，格式: "yyyy-mm-dd - yyyy-mm-dd,..."')
    
    args = parser.parse_args()
    
    print("=" * 50)
    print("  告警统计分析")
    print("=" * 50)
    
    # 仅报告模式
    if args.skip_clean:
        if not os.path.exists(os.path.join('statistics', 'cleaned_data.xlsx')):
            print("警告: statistics/cleaned_data.xlsx 不存在，请先执行清洗")
            print("      运行: python main.py")
            sys.exit(1)
        run_report()
        cleanup()
        done()
        return
    
    # 仅清洗模式
    if args.clean_only:
        run_clean(custom_periods=args.custom_periods)
        done()
        return
    
    # 默认：清洗 → 报告 → 清理
    run_clean(custom_periods=args.custom_periods)
    run_report()
    cleanup()
    done()


if __name__ == "__main__":
    main()

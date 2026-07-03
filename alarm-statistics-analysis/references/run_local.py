#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
告警分析工具 - 本地独立执行版
================================
本脚本是完整的独立入口，不依赖父目录的任何文件。
back/ 目录可以单独复制出来使用，只需要安装 requirements.txt 中的依赖。

用法:
  # 单文件分析
  python run_local.py --file 7月份全量告警明细0801.xlsx

  # 多文件对比
  python run_local.py --file 7月数据.xlsx --file 9月数据.xlsx --labels "7月,9月"

  # 自定义周期
  python run_local.py --file data.xlsx --periods "2025-07-01 - 2025-07-31"
"""

import os
import sys
import traceback
import subprocess
import shutil
import glob
import argparse
import logging
import time
import re
from datetime import datetime

# 确保本地模块优先加载
LOCAL_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, LOCAL_DIR)

# 本地目录
LOCAL_STATISTICS = os.path.join(LOCAL_DIR, "statistics")
LOCAL_LOGS = os.path.join(LOCAL_DIR, "logs")
LOCAL_TEMP = os.path.join(LOCAL_DIR, "temp")

os.makedirs(LOCAL_STATISTICS, exist_ok=True)
os.makedirs(LOCAL_LOGS, exist_ok=True)
os.makedirs(LOCAL_TEMP, exist_ok=True)


def print_step(step, msg):
    print(f"\n{'=' * 60}")
    print(f"  [{step}] {msg}")
    print(f"{'=' * 60}")


def main():
    parser = argparse.ArgumentParser(
        description='告警分析工具 - 本地独立执行版',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python run_local.py --file 7月份全量告警明细0801.xlsx
  python run_local.py --file 7月.xlsx --file 9月.xlsx --labels "7月告警,9月告警"
  python run_local.py --file data.xlsx --periods "2025-07-01 - 2025-07-31,2025-08-01 - 2025-08-31"
        """
    )
    parser.add_argument('--file', '-f', type=str, action='append', dest='files',
                        required=True, help='输入Excel文件 (可多次指定)')
    parser.add_argument('--labels', '-l', type=str, default=None,
                        help='文件标签，逗号分隔 (多文件模式使用)')
    parser.add_argument('--periods', '-p', type=str, default=None,
                        help='自定义周期范围，格式: yyyy-mm-dd - yyyy-mm-dd, ...')
    parser.add_argument('--output', '-o', type=str, default=None,
                        help='输出报告文件名 (默认自动生成)')

    args = parser.parse_args()

    print_step(0, "本地独立执行模式启动")
    print(f"  工作目录: {LOCAL_DIR}")
    print(f"  输入文件: {args.files}")
    print(f"  输出目录: {LOCAL_STATISTICS}")

    # 验证所有输入文件
    for f in args.files:
        if not os.path.exists(f):
            print(f"\n!!! 文件不存在: {f}")
            sys.exit(1)

    is_multi = len(args.files) > 1
    file_labels = [os.path.basename(f) for f in args.files]
    if args.labels:
        user_labels = [s.strip() for s in args.labels.split(',')]
        if len(user_labels) == len(args.files):
            file_labels = user_labels
        else:
            print(f"!!! 警告: labels 数量 ({len(user_labels)}) 与文件数 ({len(args.files)}) 不匹配，使用文件名")

    # ===== 步骤1: 数据清洗 =====
    print_step(1, "数据清洗")
    from data_cleaner import clean_data
    from config import OUTPUT_DIR

    clean_paths = []
    for idx, input_file in enumerate(args.files):
        output_prefix = f"cleaned_data_{idx}" if is_multi else "cleaned_data"
        print(f"\n>>> 清洗文件 [{idx+1}/{len(args.files)}]: {file_labels[idx]}")
        print(f"    输入: {os.path.abspath(input_file)}")
        print(f"    前缀: {output_prefix}")

        try:
            result = clean_data(
                custom_periods=args.periods,
                input_file=os.path.abspath(input_file),
                output_prefix=output_prefix
            )
            if result is None:
                print(f"!!! 清洗后无数据: {file_labels[idx]}")
                sys.exit(1)
            clean_paths.append(result)
            print(f"    => {result}")
        except Exception as e:
            print(f"!!! 清洗失败: {file_labels[idx]} - {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)

    print(f"\n--- 数据清洗完成，共生成 {len(clean_paths)} 个清洗产物 ---")

    # ===== 步骤2: 生成报告 =====
    print_step(2, "生成报告")

    # 构建命令行参数传给 report_generator 的主函数
    # 因为 report_generator 的入口依赖 argparse，我们直接调用其核心函数
    from report_generator import create_combined_weekly_statistics_with_chinese_names
    import report_generator as rg

    # 模拟命令行参数
    class FakeArgs:
        def __init__(self):
            self.input_file = clean_paths
            self.output_dir = LOCAL_STATISTICS
            self.output_file = "告警统计分析_combined_weekly_chinese.xlsx"
            self.temp_dir = os.path.join(LOCAL_TEMP, "report")
            self.log_level = "INFO"
            self.no_progress = True
            self.generate_report = True
            self.report_file = args.output  # 可能为 None
            self.show_examples = False
            self.custom_periods = args.periods
            self.multi_files = is_multi
            self.file_labels = ",".join(file_labels) if is_multi else None

    fake_args = FakeArgs()

    # 设置日志
    rg.setup_logging(fake_args.log_level)

    # 多文件模式 vs 单文件模式
    if fake_args.multi_files and len(fake_args.input_file) > 1:
        rg._process_multi_files(fake_args)
    else:
        # 单文件模式：需要修改 args.input_file 为列表第一个元素
        # 但 create_combined_weekly_statistics_with_chinese_names 内部会取 args.input_file[0]
        create_combined_weekly_statistics_with_chinese_names.__globals__['args'] = fake_args
        # 直接调用修改版
        _run_single_file_mode(fake_args)

    # ===== 步骤3: 整理输出 =====
    print_step(3, "整理输出")

    # 报告可能在 LOCAL_STATISTICS 下
    search_patterns = [
        "告警统计分析报告_*.html",
        "SRE-告警统计分析报告_*.html",
        "告警统计分析报告_*_lite.html",
        "SRE-告警统计分析报告_*_lite.html",
    ]
    reports = []
    for pattern in search_patterns:
        reports.extend(glob.glob(os.path.join(LOCAL_STATISTICS, pattern)))
        reports.extend(glob.glob(os.path.join(OUTPUT_DIR, pattern)))

    if not reports:
        print("!!! 未找到生成的报告文件")
        sys.exit(1)

    # 按时间排序，取最新的
    reports = sorted(set(reports), key=os.path.getmtime, reverse=True)
    latest_report = reports[0]

    # 复制到本地 statistics 目录（如果不在）
    report_name = os.path.basename(latest_report)
    local_report = os.path.join(LOCAL_STATISTICS, report_name)
    if os.path.abspath(latest_report) != os.path.abspath(local_report):
        shutil.copy2(latest_report, local_report)

    # ===== 步骤4: 清理中间产物 =====
    print_step(4, "清理中间产物")

    # 删除清洗产物
    for cp in clean_paths:
        if os.path.exists(cp):
            os.remove(cp)
            print(f"  已删除: {cp}")

    # 删除 statistics 下的临时文件
    for pattern in ("cleaned_data*.xlsx", "告警统计分析_combined_weekly_chinese.xlsx"):
        for p in glob.glob(os.path.join(OUTPUT_DIR, pattern)):
            os.remove(p)
            print(f"  已删除: {p}")

    temp_dir = os.path.join(OUTPUT_DIR, "temp")
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir, ignore_errors=True)
        print(f"  已删除临时目录: {temp_dir}")

    # 清理本地临时目录
    if os.path.exists(LOCAL_TEMP):
        shutil.rmtree(LOCAL_TEMP, ignore_errors=True)

    # 日志保留最新 5 个
    log_pattern = os.path.join(LOCAL_LOGS, "*.log")
    logs = sorted(glob.glob(log_pattern), key=os.path.getmtime, reverse=True)
    for f in logs[5:]:
        os.remove(f)

    # ===== 完成 =====
    print(f"\n{'=' * 60}")
    print(f"  分析完成!")
    print(f"  报告文件: {local_report}")
    print(f"  可在浏览器中直接打开此文件")
    print(f"{'=' * 60}")

    # 报告文件路径已输出到控制台，由用户手动打开


def _run_single_file_mode(args):
    """单文件模式的简化调用"""
    from report_generator import config
    import logging
    import pandas as pd
    import time
    from tqdm import tqdm
    import re

    input_file = args.input_file[0]

    logging.info(f"系统映射关系: {len(config.ATTR_CHINESE_MAPPING)}个系统")

    if not os.path.exists(input_file):
        logging.error(f"错误: 输入文件 {input_file} 不存在")
        return

    output_dir = args.output_dir
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, args.output_file)

    temp_dir = args.temp_dir or os.path.join(output_dir, 'temp')
    os.makedirs(temp_dir, exist_ok=True)

    logging.info(f"读取文件: {input_file}")
    try:
        file_size = os.path.getsize(input_file)
        logging.info(f"输入文件大小: {file_size} 字节")

        logging.info("尝试打开Excel文件...")
        start_time = time.time()
        with pd.ExcelFile(input_file, engine='openpyxl') as xls:
            sheet_names = xls.sheet_names
            logging.info(f"Excel文件包含以下工作表: {sheet_names}")

            weekly_sheets = [sheet for sheet in sheet_names if sheet.startswith('每周告警统计')]
            logging.info(f"找到以下每周告警统计工作表: {weekly_sheets}")

            if not weekly_sheets:
                logging.warning("警告: 没有找到任何每周告警统计工作表")
                return

            combined_df = pd.DataFrame()

            periods = None
            if args.custom_periods:
                period_ranges = [p.strip() for p in args.custom_periods.split(',')]
                periods = []
                for pr in period_ranges:
                    start, end = pr.split(' - ')
                    start_date = pd.to_datetime(start)
                    end_date = pd.to_datetime(end)
                    period_str = f"{start_date.strftime('%m.%d')}-{end_date.strftime('%m.%d')}"
                    periods.append(period_str)
                logging.info(f"使用自定义周期范围: {periods}")

            total_sheets = len(weekly_sheets)
            sheet_iterator = weekly_sheets
            if not args.no_progress:
                sheet_iterator = tqdm(weekly_sheets, desc="处理工作表", unit="sheet")

            for idx, sheet_name in enumerate(sheet_iterator, 1):
                sheet_start_time = time.time()
                logging.info(f"处理工作表 [{idx}/{total_sheets}]: {sheet_name}")
                if not args.no_progress:
                    sheet_iterator.set_description(f"处理 {sheet_name}")
                try:
                    df = pd.read_excel(xls, sheet_name)
                    logging.info(f"  工作表 {sheet_name} 包含 {len(df)} 行数据")

                    required_columns = ['attr', 'ALARM_COUNT']
                    missing_columns = [col for col in required_columns if col not in df.columns]
                    if missing_columns:
                        logging.warning(f"  工作表 {sheet_name} 中缺少以下必要列: {missing_columns}")

                    if periods:
                        df = df[df['周期'].isin(periods)]
                        if df.empty:
                            logging.warning(f"  工作表 {sheet_name} 中没有数据匹配指定周期范围")
                            continue

                    if 'attr' not in df.columns:
                        logging.warning(f"  工作表 {sheet_name} 中没有找到'attr'列，尝试从工作表名称中提取")
                        match = re.search(r'每周告警统计_(.+)', sheet_name)
                        if match:
                            attr_value = match.group(1)
                            df['attr'] = attr_value
                            logging.info(f"  从工作表名称中提取attr: {attr_value}")
                        else:
                            logging.warning(f"  无法从工作表名称中提取attr，跳过此工作表")
                            continue

                    # 聚合数据：按 attr + 告警级别 + 周期 分组
                    if '告警级别' in df.columns and '周期' in df.columns:
                        agg_dict = {
                            'ALARM_COUNT': 'sum',
                        }
                        if 'deal_status' in df.columns:
                            agg_dict['deal_status'] = lambda x: list(x)
                        group_cols = ['attr', '告警级别', '周期']
                        available_cols = [c for c in group_cols if c in df.columns]
                        df_grouped = df.groupby(available_cols).agg(agg_dict).reset_index()
                        df_grouped.rename(columns={'ALARM_COUNT': '告警数量'}, inplace=True)
                        combined_df = pd.concat([combined_df, df_grouped], ignore_index=True)
                    else:
                        # 兜底：只按 attr 聚合
                        df_grouped = df.groupby('attr')['ALARM_COUNT'].sum().reset_index()
                        df_grouped.rename(columns={'ALARM_COUNT': '告警数量'}, inplace=True)
                        combined_df = pd.concat([combined_df, df_grouped], ignore_index=True)

                    sheet_elapsed = time.time() - sheet_start_time
                    logging.info(f"  工作表 {sheet_name} 处理完成，耗时 {sheet_elapsed:.2f} 秒")

                except Exception as e:
                    logging.error(f"  处理工作表 {sheet_name} 时发生错误: {str(e)}")
                    logging.error(traceback.format_exc())

            elapsed_time = time.time() - start_time
            logging.info(f"数据读取完成，共 {len(combined_df)} 条记录，耗时 {elapsed_time:.2f} 秒")

            # 二次聚合（合并来自不同sheet的相同attr+级别+周期的记录）
            if not combined_df.empty:
                final_agg = {'告警数量': 'sum'}
                group_keys = [c for c in ['attr', '告警级别', '周期'] if c in combined_df.columns]
                if 'deal_status' in combined_df.columns:
                    # 展开 deal_status 列表 → 已处理数量 / 未超时处理数量
                    combined_df = combined_df.groupby(group_keys).agg(
                        告警数量=('告警数量', 'sum'),
                        已处理数量=('deal_status', lambda x: sum(1 for lst in x for s in lst if pd.notna(s) and s != '')),
                        未超时处理数量=('deal_status', lambda x: sum(1 for lst in x for s in lst if s == '未超时'))
                    ).reset_index()
                else:
                    combined_df = combined_df.groupby(group_keys).agg(final_agg).reset_index()

            # 替换中文系统名称
            if not combined_df.empty and 'attr' in combined_df.columns:
                combined_df['系统名称'] = combined_df['attr'].map(config.ATTR_CHINESE_MAPPING).fillna(combined_df['attr'])
                logging.info(f"已添加中文系统名称列")

            # 保存中间结果
            combined_df.to_excel(output_file, index=False, engine='openpyxl')
            logging.info(f"中间结果已保存到: {output_file}")

            # 生成 HTML 报告
            if args.generate_report:
                from report_generator import generate_html_report
                generate_html_report(combined_df, output_file, args)
                logging.info("HTML报告生成完成")

    except Exception as e:
        logging.error(f"处理文件时发生错误: {str(e)}")
        logging.error(traceback.format_exc())
        raise


if __name__ == "__main__":
    main()

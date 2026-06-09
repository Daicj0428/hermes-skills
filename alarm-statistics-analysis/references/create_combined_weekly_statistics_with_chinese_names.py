"""
告警统计分析脚本 - 整合每周告警统计数据并添加中文系统名称

此脚本用于处理告警统计数据，将所有attr的每周告警统计整合到一个sheet页中，
并将attr替换为中文名称。脚本支持命令行参数，可以生成HTML统计报告。

使用示例:
    # 基本用法
    python create_combined_weekly_statistics_with_chinese_names.py
    
    # 指定输入文件和输出目录
    python create_combined_weekly_statistics_with_chinese_names.py --input-file data.xlsx --output-dir results
    
    # 生成HTML统计报告
    python create_combined_weekly_statistics_with_chinese_names.py --generate-report
    
    # 显示进度条并使用DEBUG日志级别
    python create_combined_weekly_statistics_with_chinese_names.py --log-level DEBUG
    
    # 禁用进度条显示
    python create_combined_weekly_statistics_with_chinese_names.py --no-progress
"""

import pandas as pd
import os
import config
import re
import traceback
import sys
import logging
import time
import argparse
from datetime import datetime
from tqdm import tqdm

# 配置日志记录
def setup_logging():
    """设置日志记录"""
    log_dir = 'logs'
    os.makedirs(log_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'weekly_statistics_{timestamp}.log')
    
    # 配置日志格式和级别
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return log_file

def show_examples():
    """显示脚本使用示例"""
    examples = """
使用示例:
    # 基本用法
    python create_combined_weekly_statistics_with_chinese_names.py
    
    # 指定输入文件和输出目录
    python create_combined_weekly_statistics_with_chinese_names.py --input-file data.xlsx --output-dir results
    
    # 生成HTML统计报告
    python create_combined_weekly_statistics_with_chinese_names.py --generate-report
    
    # 显示进度条并使用DEBUG日志级别
    python create_combined_weekly_statistics_with_chinese_names.py --log-level DEBUG
    
    # 禁用进度条显示
    python create_combined_weekly_statistics_with_chinese_names.py --no-progress
    
    # 指定HTML报告文件名
    python create_combined_weekly_statistics_with_chinese_names.py --generate-report --report-file report.html
    """
    print(examples)

def parse_arguments():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description='整合每周告警统计数据并添加中文系统名称')
    parser.add_argument('--input-file', type=str, default=os.path.join(config.OUTPUT_DIR, 'cleaned_data.xlsx'),
                        help='输入Excel文件路径 (默认: statistics/cleaned_data.xlsx)')
    parser.add_argument('--output-dir', type=str, default='statistics',
                        help='输出目录 (默认: statistics)')
    parser.add_argument('--output-file', type=str, default="告警统计分析_combined_weekly_chinese.xlsx",
                        help='输出文件名 (默认: 告警统计分析_combined_weekly_chinese.xlsx)')
    parser.add_argument('--temp-dir', type=str, default=None,
                        help='临时文件目录 (默认: <output_dir>/temp)')
    parser.add_argument('--log-level', type=str, choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                        default='INFO', help='日志级别 (默认: INFO)')
    parser.add_argument('--no-progress', action='store_true',
                        help='禁用进度条显示')
    parser.add_argument('--generate-report', action='store_true',
                        help='生成HTML统计报告')
    parser.add_argument('--report-file', type=str, default=None,
                        help='HTML报告文件名 (默认: <output_dir>/SRE-告警统计分析报告_<timestamp>.html)')
    parser.add_argument('--show-examples', action='store_true',
                        help='显示使用示例并退出')
    parser.add_argument('--custom-periods', type=str, default=None,
                        help='自定义周期范围，格式为 yyyy-mm-dd - yyyy-mm-dd,yyyy-mm-dd - yyyy-mm-dd')
    
    args = parser.parse_args()
    
    # 如果指定了显示示例，则显示示例并退出
    if args.show_examples:
        show_examples()
        sys.exit(0)
    
    return args

def generate_html_report(df, output_file, args):
    """生成交互式HTML统计报告（支持多维度筛选和指标排序）"""
    import json
    
    logging.info("生成交互式HTML统计报告...")
    
    # 如果未指定报告文件名，使用默认名称
    if not args.report_file:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_file = os.path.join(args.output_dir, f"SRE-告警统计分析报告_{timestamp}.html")
    else:
        report_file = args.report_file
    
    # 确保报告目录存在
    os.makedirs(os.path.dirname(os.path.abspath(report_file)), exist_ok=True)
    
    # ========== 1. 构建多维度聚合数据 ==========
    # 按 系统+告警级别+周期 三维聚合
    group_cols = ['attr', '告警级别', '周期']
    available_cols = [c for c in group_cols if c in df.columns]
    
    # 直接从 deal_status 聚合，避免依赖可能为0的中间计算列
    if 'deal_status' in df.columns:
        dim_stats = df.groupby(available_cols).agg(
            告警记录数=('告警数量', 'count'),
            已处理记录数=('deal_status', lambda x: (x == '已处理').sum()),
            未超时处理记录数=('deal_status', lambda x: (x == '未超时').sum())
        ).reset_index()
    else:
        dim_stats = df.groupby(available_cols).agg(
            告警记录数=('告警数量', 'count'),
            已处理记录数=('已处理数量', 'sum'),
            未超时处理记录数=('未超时处理数量', 'sum')
        ).reset_index()
    
    dim_stats['处理率'] = dim_stats.apply(
        lambda r: round(r['已处理记录数'] / r['告警记录数'] * 100, 2) if r['告警记录数'] > 0 else 0, axis=1
    )
    dim_stats['未超时率'] = dim_stats.apply(
        lambda r: round(r['未超时处理记录数'] / r['告警记录数'] * 100, 2) if r['告警记录数'] > 0 else 0, axis=1
    )
    
    # 转换为JSON供前端使用
    table_data = dim_stats.to_dict(orient='records')
    
    # 提取筛选维度唯一值
    all_systems = sorted(dim_stats['attr'].unique().tolist()) if 'attr' in dim_stats.columns else []
    all_levels = sorted(dim_stats['告警级别'].unique().tolist()) if '告警级别' in dim_stats.columns else []
    all_periods = sorted(dim_stats['周期'].unique().tolist()) if '周期' in dim_stats.columns else []
    
    # ========== 1.5 构建原始告警记录数据（供自定义导出使用） ==========
    raw_export_cols = ['attr', '告警级别', '周期', 'alarm_content', 'alarm_first_time',
                       'alarm_last_time', 'DISCHARGE_TIME', 'ALARM_COUNT', 'alarm_state',
                       'deal_status', 'deal_time', 'ack_status', 'LAST_ACK_TIME',
                       'ACK_COMMENT', 'mainResp', 'secondResp', 'RESOURCE_NAME', 'system']
    raw_available = [c for c in raw_export_cols if c in df.columns]
    raw_data = df[raw_available].copy()
    # 时间列转为字符串避免JSON序列化问题
    for col in raw_data.columns:
        if raw_data[col].dtype == 'datetime64[ns]':
            raw_data[col] = raw_data[col].dt.strftime('%Y-%m-%d %H:%M:%S')
    # 填充NaN为空字符串
    raw_data = raw_data.fillna('')
    raw_data_json = raw_data.to_dict(orient='records')
    
    # ========== 2. 计算总体摘要数据 ==========
    total_alarms = int(dim_stats['告警记录数'].sum())
    total_processed = int(dim_stats['已处理记录数'].sum())
    total_on_time = int(dim_stats['未超时处理记录数'].sum())
    processed_rate = round(total_processed / total_alarms * 100, 2) if total_alarms > 0 else 0
    on_time_rate = round(total_on_time / total_alarms * 100, 2) if total_alarms > 0 else 0
    
    # ========== 3. 生成交互式HTML ==========
    html_content = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SRE-告警统计分析报告</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
    <style>
        /* ====== 全局 ====== */
        :root {{
            --ink: #1e293b;
            --muted: #64748b;
            --border: #e2e8f0;
            --surface: #ffffff;
            --bg: #f1f5f9;
            --accent: #2563eb;
            --accent-light: #eff6ff;
            --green: #059669;
            --slate: #475569;
            --radius: 6px;
            --shadow-sm: 0 1px 2px rgba(0,0,0,0.04);
            --shadow: 0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04);
        }}
        *,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
        body{{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", "Helvetica Neue", sans-serif;
            background: var(--bg); color: var(--ink); line-height: 1.5; font-size: 14px;
        }}
        .container{{max-width:1440px;margin:0 auto;padding:24px 28px}}

        /* ====== 顶部 ====== */
        .header{{
            display:flex;justify-content:space-between;align-items:center;
            margin-bottom:24px;padding-bottom:20px;border-bottom:1px solid var(--border);
        }}
        .header-left h1{{font-size:22px;font-weight:700;letter-spacing:-0.3px;color:var(--ink);}}
        .header-left .subtitle{{font-size:13px;color:var(--muted);margin-top:2px}}
        .header-right{{display:flex;align-items:center;gap:18px;font-size:12px;color:var(--muted)}}
        .header-right span{{display:flex;align-items:center;gap:4px}}
        .header-right .dot{{width:6px;height:6px;border-radius:50%;background:var(--green);display:inline-block}}

        /* ====== 指标卡片 ====== */
        .summary-grid{{
            display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px;
        }}
        @media(max-width:900px){{.summary-grid{{grid-template-columns:repeat(2,1fr)}}}}
        .summary-card{{
            background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
            padding:14px 16px;position:relative;overflow:hidden;
        }}
        .summary-card .card-icon{{display:none}}
        .summary-card .card-label{{
            font-size:12px;color:var(--muted);margin-bottom:4px;letter-spacing:0.2px;
            display:flex;align-items:center;gap:6px;
        }}
        .summary-card .card-label::before{{
            content:'';display:inline-block;width:6px;height:6px;border-radius:50%;flex-shrink:0;
        }}
        .summary-card.total .card-label::before{{background:var(--accent)}}
        .summary-card.ontime .card-label::before{{background:var(--green)}}
        .summary-card.systems .card-label::before{{background:var(--slate)}}
        .summary-card .card-value{{font-size:26px;font-weight:700;color:var(--ink);line-height:1.1}}
        .summary-card .card-rate{{font-size:12px;color:var(--muted);margin-top:4px;font-variant-numeric:tabular-nums}}
        .summary-card .card-rate strong{{font-weight:600;color:var(--ink)}}
        .summary-card.total .card-value{{color:var(--accent)}}
        .summary-card.ontime .card-value{{color:var(--green)}}
        .summary-card.systems .card-value{{color:var(--slate)}}

        /* ====== 筛选面板 ====== */
        .filter-panel{{
            background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
            padding:14px 20px;margin-bottom:20px;
            display:flex;flex-wrap:wrap;align-items:flex-end;gap:14px;
        }}
        .filter-group{{display:flex;flex-direction:column;gap:3px}}
        .filter-group label{{font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:0.5px}}
        .filter-group select{{
            padding:7px 10px;border:1px solid var(--border);border-radius:4px;font-size:13px;
            background:var(--surface);min-width:140px;outline:none;color:var(--ink);cursor:pointer;
        }}
        .filter-group select:focus{{border-color:var(--accent)}}
        .btn{{
            padding:7px 16px;border:1px solid var(--border);border-radius:4px;font-size:13px;
            cursor:pointer;background:var(--surface);color:var(--ink);white-space:nowrap;
            transition:background .15s,border-color .15s;
        }}
        .btn:hover{{background:#f8fafc;border-color:#cbd5e1}}
        .btn-reset{{}}
        .btn-export{{background:var(--accent);color:#fff;border-color:var(--accent)}}
        .btn-export:hover{{background:#1d4ed8;border-color:#1d4ed8}}
        .filter-info{{font-size:12px;color:var(--muted);margin-left:auto;white-space:nowrap}}

        /* ====== 多选下拉 ====== */
        .multi-select{{position:relative;min-width:170px}}
        .multi-select-trigger{{
            display:flex;align-items:center;justify-content:space-between;
            padding:7px 10px;border:1px solid var(--border);border-radius:4px;
            font-size:13px;background:var(--surface);cursor:pointer;gap:6px;min-width:170px;
            user-select:none;color:var(--ink);
        }}
        .multi-select-trigger:hover{{border-color:#94a3b8}}
        .multi-select-trigger.active{{border-color:var(--accent)}}
        .multi-select-text{{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:13px}}
        .multi-select-count{{background:var(--accent);color:#fff;font-size:10px;padding:1px 6px;border-radius:8px;flex-shrink:0}}
        .multi-select-arrow{{font-size:9px;color:var(--muted);flex-shrink:0;transition:transform .15s}}
        .multi-select-trigger.active .multi-select-arrow{{transform:rotate(180deg)}}
        .multi-select-dropdown{{
            display:none;position:absolute;top:calc(100% + 3px);left:0;z-index:100;
            background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
            box-shadow:0 10px 25px rgba(0,0,0,0.1);min-width:100%;max-height:280px;overflow-y:auto;padding:4px 0;
        }}
        .multi-select-dropdown.show{{display:block}}
        .multi-select-dropdown .select-all-row{{padding:5px 12px;border-bottom:1px solid var(--border);margin-bottom:2px}}
        .multi-select-dropdown .select-all-row a{{color:var(--accent);font-size:11px;cursor:pointer;text-decoration:none}}
        .multi-select-dropdown .select-all-row a:hover{{text-decoration:underline}}
        .multi-option{{display:flex;align-items:center;gap:7px;padding:5px 12px;cursor:pointer;font-size:12px;color:var(--ink)}}
        .multi-option:hover{{background:var(--accent-light)}}
        .multi-option input[type="checkbox"]{{accent-color:var(--accent);width:14px;height:14px;cursor:pointer;flex-shrink:0}}
        .multi-option span{{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}

        /* ====== 图表区 ====== */
        .charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}}
        @media(max-width:960px){{.charts-grid{{grid-template-columns:1fr}}}}
        .chart-card{{
            background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
            padding:18px 20px;
        }}
        .chart-card h3{{font-size:14px;font-weight:600;color:var(--ink);margin-bottom:10px}}
        .chart-wrapper{{position:relative;width:100%;height:320px}}

        /* 树状图 */
        .treemap-card{{
            background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
            padding:18px 20px;margin-bottom:20px;
        }}
        .treemap-card h3{{font-size:14px;font-weight:600;color:var(--ink);margin-bottom:10px}}
        .treemap-legend{{display:flex;flex-wrap:wrap;gap:14px;margin-top:8px;font-size:11px}}
        .treemap-legend-item{{display:flex;align-items:center;gap:4px}}
        .treemap-legend-color{{width:12px;height:12px;border-radius:2px;flex-shrink:0}}
        .treemap-tooltip{{
            position:absolute;background:rgba(15,23,42,0.92);color:#fff;padding:7px 10px;
            border-radius:4px;font-size:11px;pointer-events:none;z-index:999;display:none;white-space:nowrap;
            line-height:1.5;
        }}

        /* 多周期树状图 */
        .period-label{{position:absolute;top:6px;left:50%;transform:translateX(-50%);font-size:11px;font-weight:600;color:var(--muted);text-align:center;white-space:nowrap}}
        .period-divider{{position:absolute;top:0;bottom:0;width:1px;background:var(--border)}}

        /* ====== 表格 ====== */
        .table-card{{
            background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
            padding:18px 20px;overflow:hidden;margin-bottom:20px;
        }}
        .table-card h3{{font-size:14px;font-weight:600;color:var(--ink);margin-bottom:10px}}
        .table-toolbar{{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:10px;align-items:center}}
        .search-input{{
            padding:7px 10px;border:1px solid var(--border);border-radius:4px;
            font-size:13px;width:220px;outline:none;background:var(--bg);color:var(--ink);
        }}
        .search-input:focus{{border-color:var(--accent);background:var(--surface)}}
        .table-wrapper{{overflow-x:auto}}
        table{{width:100%;border-collapse:collapse;font-size:13px}}
        th{{
            background:#f8fafc;padding:9px 12px;text-align:left;font-weight:600;color:var(--muted);
            border-bottom:1px solid var(--border);white-space:nowrap;cursor:pointer;user-select:none;
            font-size:11px;text-transform:uppercase;letter-spacing:0.3px;
        }}
        th:hover{{color:var(--accent)}}
        th .sort-icon{{margin-left:3px;font-size:10px;opacity:0.3}}
        th.sorted .sort-icon{{opacity:1;color:var(--accent)}}
        th.sorted-asc .sort-icon::after{{content:" ▴"}}
        th.sorted-desc .sort-icon::after{{content:" ▾"}}
        td{{padding:8px 12px;border-bottom:1px solid #f1f5f9}}
        tr:hover td{{background:#f8fafc}}
        .rate-badge{{display:inline-block;padding:2px 8px;border-radius:3px;font-size:11px;font-weight:600;min-width:52px;text-align:center}}
        .rate-high{{background:#ecfdf5;color:#065f46}}
        .rate-mid{{background:#fffbeb;color:#92400e}}
        .rate-low{{background:#fef2f2;color:#991b1b}}
        .pagination{{display:flex;justify-content:center;align-items:center;gap:4px;margin-top:14px}}
        .pagination button{{
            padding:5px 10px;border:1px solid var(--border);border-radius:4px;
            background:var(--surface);cursor:pointer;font-size:12px;color:var(--ink);
        }}
        .pagination button:hover{{background:#f8fafc}}
        .pagination button.active{{background:var(--ink);color:#fff;border-color:var(--ink)}}
        .pagination button:disabled{{opacity:0.3;cursor:not-allowed}}
        .page-info{{font-size:12px;color:var(--muted)}}

        /* Footer */
        .footer{{text-align:center;padding:20px;color:var(--muted);font-size:11px;border-top:1px solid var(--border);margin-top:8px}}

        /* Collapse */
        .collapse-arrow{{
            display:inline-block;transition:transform .2s;font-size:11px;
        }}
        .collapse-arrow.collapsed{{transform:rotate(-90deg)}}
        .top-collapse-body{{transition:max-height .3s ease;overflow:hidden}}

        /* Modal */
        .modal-overlay{{
            display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(15,23,42,0.5);
            z-index:2000;align-items:center;justify-content:center;
        }}
        .modal-overlay.show{{display:flex}}
        .modal{{
            background:var(--surface);border-radius:var(--radius);width:520px;max-width:94vw;
            max-height:80vh;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,0.18);
        }}
        .modal-header{{
            display:flex;justify-content:space-between;align-items:center;
            padding:16px 20px;border-bottom:1px solid var(--border);
        }}
        .modal-header h4{{font-size:15px;font-weight:600;color:var(--ink)}}
        .modal-close{{
            background:none;border:none;font-size:18px;cursor:pointer;color:var(--muted);padding:0 4px;
        }}
        .modal-close:hover{{color:var(--ink)}}
        .modal-body{{padding:16px 20px;overflow-y:auto;flex:1}}
        .modal-body .field-grid{{
            display:grid;grid-template-columns:repeat(2,1fr);gap:6px;
        }}
        .modal-body .select-all-row{{margin-bottom:10px;padding-bottom:8px;border-bottom:1px solid var(--border)}}
        .modal-body .select-all-row a{{font-size:12px;color:var(--accent);cursor:pointer;text-decoration:none}}
        .modal-body .select-all-row a:hover{{text-decoration:underline}}
        .modal-content{{
            background:var(--surface);border-radius:var(--radius);width:520px;max-width:94vw;
            max-height:80vh;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,0.18);
            padding:16px 20px;
        }}
        .modal-footer{{
            display:flex;justify-content:flex-end;gap:8px;padding:14px 20px;border-top:1px solid var(--border);
        }}
        .modal-footer .btn-cancel{{background:var(--surface);color:var(--ink);border:1px solid var(--border)}}
        .modal-footer .btn-confirm{{background:var(--accent);color:#fff;border-color:var(--accent)}}
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <div class="header-left">
            <h1>SRE-告警统计分析报告</h1>
            <div class="subtitle">Alarm Statistics & Analysis Dashboard</div>
        </div>
        <div class="header-right">
            <span><span class="dot"></span> 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}</span>
            <span>数据源: {os.path.basename(config.INPUT_FILE)}</span>
        </div>
    </div>

    <!-- 总体统计卡片 -->
    <div class="summary-grid" id="summaryGrid">
        <div class="summary-card total">
            <div class="card-icon"></div>
            <div class="card-label">告警记录总数</div>
            <div class="card-value" id="sumTotal">{total_alarms:,}</div>
        </div>
        <div class="summary-card ontime">
            <div class="card-icon"></div>
            <div class="card-label">未超时处理</div>
            <div class="card-value" id="sumOnTime">{total_on_time:,}</div>
            <div class="card-rate">未超时率 <strong id="sumOnTimeRate">{on_time_rate}%</strong></div>
        </div>
        <div class="summary-card systems">
            <div class="card-icon"></div>
            <div class="card-label">覆盖业务系统</div>
            <div class="card-value" id="sumSystems">{len(all_systems)}</div>
            <div class="card-rate">个</div>
        </div>
    </div>

    <!-- 筛选面板 -->
    <div class="filter-panel" id="filterPanel">
        <div class="filter-group">
            <label>业务系统</label>
            <div class="multi-select" id="multiSystem">
                <div class="multi-select-trigger" onclick="toggleMultiDropdown('multiSystem')">
                    <span class="multi-select-text" id="systemText">全部系统</span>
                    <span class="multi-select-count" id="systemCount" style="display:none"></span>
                    <span class="multi-select-arrow">&#x25BC;</span>
                </div>
                <div class="multi-select-dropdown" id="multiSystemDropdown">
                    <div class="select-all-row">
                        <a href="javascript:void(0)" onclick="event.stopPropagation();selectAllSystems(true)">全选</a>
                        &nbsp;
                        <a href="javascript:void(0)" onclick="event.stopPropagation();selectAllSystems(false)">取消</a>
                    </div>
                    <div id="systemOptions"></div>
                </div>
            </div>
        </div>
        <div class="filter-group">
            <label>告警级别</label>
            <div class="multi-select" id="multiLevel">
                <div class="multi-select-trigger" onclick="toggleMultiDropdown('multiLevel')">
                    <span class="multi-select-text" id="levelText">全部级别</span>
                    <span class="multi-select-count" id="levelCount" style="display:none"></span>
                    <span class="multi-select-arrow">&#x25BC;</span>
                </div>
                <div class="multi-select-dropdown" id="multiLevelDropdown">
                    <div class="select-all-row">
                        <a href="javascript:void(0)" onclick="event.stopPropagation();selectAllLevels(true)">全选</a>
                        &nbsp;
                        <a href="javascript:void(0)" onclick="event.stopPropagation();selectAllLevels(false)">取消</a>
                    </div>
                    <div id="levelOptions"></div>
                </div>
            </div>
        </div>
        <div class="filter-group">
            <label>统计周期</label>
            <div class="multi-select" id="multiPeriod">
                <div class="multi-select-trigger" onclick="toggleMultiDropdown('multiPeriod')">
                    <span class="multi-select-text" id="periodText">全部周期</span>
                    <span class="multi-select-count" id="periodCount" style="display:none"></span>
                    <span class="multi-select-arrow">&#x25BC;</span>
                </div>
                <div class="multi-select-dropdown" id="multiPeriodDropdown">
                    <div class="select-all-row">
                        <a href="javascript:void(0)" onclick="event.stopPropagation();selectAllPeriods(true)">全选</a>
                        &nbsp;
                        <a href="javascript:void(0)" onclick="event.stopPropagation();selectAllPeriods(false)">取消</a>
                    </div>
                    <div id="periodOptions"></div>
                </div>
            </div>
        </div>
        <button class="btn btn-reset" onclick="resetFilters()">重置筛选</button>
        <button class="btn btn-export" onclick="exportTable()">导出 CSV</button>
        <button class="btn btn-export" onclick="openCustomExport()" style="background:#475569;border-color:#475569;">自定义导出</button>
        <span class="filter-info" id="filterInfo">共 {len(table_data)} 条记录</span>
    </div>

    <!-- 图表区域（动态切换） -->
    <div id="chartContainer">
        <div class="charts-grid">
            <div class="chart-card">
                <h3>系统告警数量分布（按告警级别分层）</h3>
                <div class="chart-wrapper"><canvas id="chartAlarmCount"></canvas></div>
            </div>
            <div class="chart-card">
                <h3>系统未超时处理率</h3>
                <div class="chart-wrapper"><canvas id="chartRates"></canvas></div>
            </div>
        </div>
    </div>

    <!-- TOP告警处理时长 -->
    <div class="table-card" id="topPanel">
        <div class="table-toolbar">
            <h3 style="margin:0;display:flex;align-items:center;gap:8px;cursor:pointer;" onclick="toggleTopPanel()">
                <span class="collapse-arrow" id="topArrow">&#9660;</span>
                TOP 告警处理时长
            </h3>
            <div style="display:flex;align-items:center;gap:10px;">
                <span style="font-size:12px;color:var(--muted);">
                    显示前 <input type="number" id="topCount" value="20" min="1" max="500"
                    style="width:56px;padding:3px 5px;border:1px solid var(--border);border-radius:4px;font-size:13px;text-align:center;"
                    onchange="renderTop()" onkeydown="if(event.key==='Enter')renderTop()"> 条
                </span>
                <button class="btn" onclick="exportTopCustom()" title="自定义字段导出TOP数据">⤓ 导出TOP</button>
            </div>
        </div>
        <div class="top-collapse-body" id="topCollapseBody">
        <div class="table-wrapper">
            <table>
                <thead>
                    <tr>
                        <th>所属系统</th>
                        <th>告警级别</th>
                        <th>资源对象</th>
                        <th>告警内容</th>
                        <th>处理时长</th>
                        <th>首次告警时间</th>
                        <th>处理状态</th>
                    </tr>
                </thead>
                <tbody id="topBody"><tr><td colspan="7" style="text-align:center;color:var(--muted);">加载中...</td></tr></tbody>
            </table>
        </div>
        </div>
    </div>


    <!-- 详细数据表格 -->
    <div class="table-card">
        <h3>告警明细数据（点击表头排序）</h3>
        <div class="table-toolbar">
            <input type="text" class="search-input" id="searchInput" placeholder="搜索系统名称..." oninput="applyAllFilters()">
            <span style="font-size:12px;color:var(--muted);" id="tableInfo"></span>
        </div>
        <div class="table-wrapper">
            <table id="dataTable">
                <thead>
                    <tr>
                        <th onclick="sortTable('attr')">所属系统 <span class="sort-icon"></span></th>
                        <th onclick="sortTable('告警级别')">告警级别 <span class="sort-icon"></span></th>
                        <th onclick="sortTable('周期')">周期 <span class="sort-icon"></span></th>
                        <th onclick="sortTable('告警记录数')">告警记录数 <span class="sort-icon"></span></th>
                        <th onclick="sortTable('未超时处理记录数')">未超时处理记录数 <span class="sort-icon"></span></th>
                        <th onclick="sortTable('处理率')">处理率(%) <span class="sort-icon"></span></th>
                        <th onclick="sortTable('未超时率')">未超时率(%) <span class="sort-icon"></span></th>
                    </tr>
                </thead>
                <tbody id="tableBody"></tbody>
            </table>
        </div>
        <div class="pagination" id="pagination"></div>
    </div>

    <div class="footer">告警监控平台 &middot; 统计周期 {os.path.basename(config.INPUT_FILE)} &middot; {datetime.now().year}</div>
</div>

<!-- 自定义导出弹窗 -->
<div class="modal-overlay" id="exportModal">
    <div class="modal">
        <div class="modal-header">
            <h4>自定义导出字段</h4>
            <button class="modal-close" onclick="closeCustomExport()">&times;</button>
        </div>
        <div class="modal-body">
            <p style="font-size:12px;color:var(--muted);margin-bottom:10px;">
                当前已筛选 <strong id="exportRecordCount">0</strong> 条记录，请选择要导出的字段：
            </p>
            <div class="select-all-row">
                <a href="javascript:void(0)" onclick="toggleAllExportFields(true)">全选</a>
                &nbsp;
                <a href="javascript:void(0)" onclick="toggleAllExportFields(false)">取消</a>
            </div>
            <div class="field-grid" id="exportFieldGrid"></div>
        </div>
        <div class="modal-footer">
            <button class="btn btn-cancel" onclick="closeCustomExport()">取消</button>
            <button class="btn btn-confirm" onclick="doCustomExport()">导出 CSV</button>
        </div>
    </div>
</div>

<script>
// ==================== 数据 & 配置 ====================
const RAW_DATA = {json.dumps(table_data, ensure_ascii=False)};
const RAW_EXPORT = {json.dumps(raw_data_json, ensure_ascii=False)};
const ALL_SYSTEMS = {json.dumps(all_systems, ensure_ascii=False)};
const ALL_LEVELS = {json.dumps(all_levels, ensure_ascii=False)};
const ALL_PERIODS = {json.dumps(all_periods, ensure_ascii=False)};
const EXPORT_FIELDS = {json.dumps(config.EXPORT_FIELDS, ensure_ascii=False)};


// ==================== 多选状态 ====================
let selectedSystems = [...ALL_SYSTEMS];
let selectedPeriods = [...ALL_PERIODS];
let selectedLevels = [...ALL_LEVELS];

// ==================== 排序/分页状态 ====================
let sortColumn = '告警记录数';
let sortDirection = 'desc';
let currentPage = 1;
const PAGE_SIZE = 15;

// ==================== 图表实例管理 ====================
let chartInstances = {{}};
let treemapCanvas = null;
const LEVEL_COLORS = {{
    '严重告警': '#d93025',
    '重要告警': '#e37400',
    '一般告警': '#1a73e8',
    '警告': '#f9ab00',
    '提示': '#34a853',
    '未知': '#9e9e9e'
}};

// ==================== 初始化 ====================
function init() {{
    // 填充多选：告警级别
    buildMultiCheckboxes('levelOptions', ALL_LEVELS, selectedLevels, onLevelToggle);
    // 填充多选：系统
    buildMultiCheckboxes('systemOptions', ALL_SYSTEMS, selectedSystems, onSystemToggle);
    // 填充多选：周期
    buildMultiCheckboxes('periodOptions', ALL_PERIODS, selectedPeriods, onPeriodToggle);

    // 绑定事件
    // 点击外部关闭多选下拉
    document.addEventListener('click', function(e) {{
        if (!e.target.closest('.multi-select')) {{
            document.querySelectorAll('.multi-select-dropdown.show').forEach(d => d.classList.remove('show'));
            document.querySelectorAll('.multi-select-trigger.active').forEach(t => t.classList.remove('active'));
        }}
    }});
    
    // 首次渲染
    applyAllFilters();
}}

function populateSelect(id, options) {{
    const sel = document.getElementById(id);
    options.forEach(opt => {{
        const el = document.createElement('option');
        el.value = opt;
        el.textContent = opt;
        sel.appendChild(el);
    }});
}}

// ==================== 多选下拉 ====================
function buildMultiCheckboxes(containerId, options, selectedArr, onChangeFn) {{
    const container = document.getElementById(containerId);
    container.innerHTML = options.map(opt => {{
        const checked = selectedArr.includes(opt) ? 'checked' : '';
        return `<label class="multi-option" onclick="event.stopPropagation()">
            <input type="checkbox" value="${{opt}}" ${{checked}} onchange="onCheckboxChange('${{containerId.replace('Options','')}}', this)">
            <span>${{opt}}</span>
        </label>`;
    }}).join('');
}}

function onCheckboxChange(prefix, cb) {{
    const val = cb.value;
    if (prefix === 'system') {{
        if (cb.checked) {{ if (!selectedSystems.includes(val)) selectedSystems.push(val); }}
        else {{ selectedSystems = selectedSystems.filter(s => s !== val); }}
        updateMultiSelectDisplay('multiSystem', 'systemText', 'systemCount', selectedSystems, ALL_SYSTEMS, '系统');
    }} else if (prefix === 'period') {{
        if (cb.checked) {{ if (!selectedPeriods.includes(val)) selectedPeriods.push(val); }}
        else {{ selectedPeriods = selectedPeriods.filter(p => p !== val); }}
        updateMultiSelectDisplay('multiPeriod', 'periodText', 'periodCount', selectedPeriods, ALL_PERIODS, '周期');
    }} else if (prefix === 'level') {{
        if (cb.checked) {{ if (!selectedLevels.includes(val)) selectedLevels.push(val); }}
        else {{ selectedLevels = selectedLevels.filter(l => l !== val); }}
        updateMultiSelectDisplay('multiLevel', 'levelText', 'levelCount', selectedLevels, ALL_LEVELS, '级别');
    }}
    applyAllFilters();
}}

function toggleMultiDropdown(id) {{
    const dropdown = document.getElementById(id + 'Dropdown');
    const trigger = document.querySelector('#' + id + ' .multi-select-trigger');
    const isOpen = dropdown.classList.contains('show');
    // 关闭所有
    document.querySelectorAll('.multi-select-dropdown.show').forEach(d => d.classList.remove('show'));
    document.querySelectorAll('.multi-select-trigger.active').forEach(t => t.classList.remove('active'));
    if (!isOpen) {{
        dropdown.classList.add('show');
        trigger.classList.add('active');
    }}
}}

function updateMultiSelectDisplay(msId, textId, countId, selectedArr, allArr, label) {{
    const textEl = document.getElementById(textId);
    const countEl = document.getElementById(countId);
    if (selectedArr.length === 0) {{
        textEl.textContent = '请选择' + label;
        textEl.style.color = '#999';
        countEl.style.display = 'none';
    }} else if (selectedArr.length === allArr.length) {{
        textEl.textContent = '全部' + label;
        textEl.style.color = '#333';
        countEl.style.display = 'none';
    }} else {{
        textEl.textContent = selectedArr.slice(0, 3).join(', ') + (selectedArr.length > 3 ? ' +' + (selectedArr.length - 3) + '...' : '');
        textEl.style.color = '#333';
        countEl.textContent = selectedArr.length;
        countEl.style.display = 'inline-block';
    }}
}}

function selectAllSystems(select) {{
    const checkboxes = document.querySelectorAll('#systemOptions input[type="checkbox"]');
    checkboxes.forEach(cb => {{ cb.checked = select; }});
    selectedSystems = select ? [...ALL_SYSTEMS] : [];
    updateMultiSelectDisplay('multiSystem', 'systemText', 'systemCount', selectedSystems, ALL_SYSTEMS, '系统');
    applyAllFilters();
}}

function selectAllPeriods(select) {{
    const checkboxes = document.querySelectorAll('#periodOptions input[type="checkbox"]');
    checkboxes.forEach(cb => {{ cb.checked = select; }});
    selectedPeriods = select ? [...ALL_PERIODS] : [];
    updateMultiSelectDisplay('multiPeriod', 'periodText', 'periodCount', selectedPeriods, ALL_PERIODS, '周期');
    applyAllFilters();
}}

function selectAllLevels(select) {{
    const checkboxes = document.querySelectorAll('#levelOptions input[type="checkbox"]');
    checkboxes.forEach(cb => {{ cb.checked = select; }});
    selectedLevels = select ? [...ALL_LEVELS] : [];
    updateMultiSelectDisplay('multiLevel', 'levelText', 'levelCount', selectedLevels, ALL_LEVELS, '级别');
    applyAllFilters();
}}

function onSystemToggle() {{ /* handled in onCheckboxChange */ }}
function onPeriodToggle() {{ /* handled in onCheckboxChange */ }}

function onLevelToggle() {{ /* handled in onCheckboxChange */ }}

// ==================== 筛选逻辑 ====================
function getFilteredData() {{
    const searchText = document.getElementById('searchInput').value.toLowerCase();
    return RAW_DATA.filter(row => {{
        if (selectedSystems.length > 0 && selectedSystems.length < ALL_SYSTEMS.length && !selectedSystems.includes(row.attr)) return false;
        if (selectedPeriods.length > 0 && selectedPeriods.length < ALL_PERIODS.length && !selectedPeriods.includes(row['周期'])) return false;
        if (selectedLevels.length > 0 && selectedLevels.length < ALL_LEVELS.length && !selectedLevels.includes(row['告警级别'])) return false;
        if (searchText && !row.attr.toLowerCase().includes(searchText)) return false;
        return true;
    }});
}}

function applyAllFilters() {{
    const filtered = getFilteredData();
    const sorted = sortData(filtered, sortColumn, sortDirection);
    document.getElementById('filterInfo').textContent = '共 ' + filtered.length + ' 条记录';
    updateSummaryCards(filtered);
    renderTable(sorted);
    renderDynamicCharts(filtered);
    renderTop();
}}

// ==================== 动态图表渲染 ====================
function getActiveSystems(data) {{
    return [...new Set(data.map(r => r.attr))].sort();
}}

function getActivePeriods(data) {{
    return [...new Set(data.map(r => r['周期']))].sort();
}}

function getActiveLevels(data) {{
    return [...new Set(data.map(r => r['告警级别']))].sort();
}}

function getStackedDataBySystems(data, systems) {{
    const levels = getActiveLevels(data);
    const datasets = levels.map((lvl, i) => {{
        const colors = ['#d93025','#e37400','#1a73e8','#34a853','#f9ab00','#9e9e9e'];
        return {{
            label: lvl,
            data: systems.map(sys => data.filter(r => r.attr === sys && r['告警级别'] === lvl).reduce((s, r) => s + r['告警记录数'], 0)),
            backgroundColor: LEVEL_COLORS[lvl] || colors[i % colors.length],
            borderRadius: 0,
        }};
    }});
    return {{ systems, levels, datasets }};
}}

function renderDynamicCharts(data) {{
    const systems = getActiveSystems(data);
    const periods = getActivePeriods(data);
    const container = document.getElementById('chartContainer');
    
    const isSinglePeriod = periods.length === 1;
    const isManySystems = systems.length > 8;
    
    if (isSinglePeriod && !isManySystems) {{
        // 单选周期 + 少系统 → 单周期树状图
        buildTreemapLayout(container, data, systems, periods[0]);
    }} else if (isSinglePeriod && isManySystems) {{
        // 单选周期 + 多系统 → 叠加柱状图
        buildSingleStackedBar(container, data, systems.slice(0, 12));
    }} else if (!isSinglePeriod && isManySystems) {{
        // 多周期 + 多系统 → 叠加柱状图
        buildSingleStackedBar(container, data, systems.slice(0, 12));
    }} else {{
        // 多周期 + 少系统 → 小多组柱状图（每系统一张，周期为列）
        buildSmallMultiplesGrid(container, data, systems, periods.slice(0, 5));
    }}
    
    renderRatesChart(data, systems.slice(0, 10));

}}

// ==================== 单周期树状图 ====================
function buildTreemapLayout(container, data, systems, periodName) {{
    destroyAllCharts();
    container.innerHTML = `
        <div class="treemap-card">
            <h3>系统告警分布 &mdash; ${{periodName}}</h3>
            <div style="position:relative;">
                <canvas id="treemapCanvas" style="width:100%;height:500px;"></canvas>
                <div class="treemap-tooltip" id="treemapTooltip"></div>
            </div>
            <div class="treemap-legend" id="treemapLegend"></div>
        </div>
        <div class="chart-card" style="margin-top:16px;">
            <h3>系统未超时处理率</h3>
            <div class="chart-wrapper"><canvas id="chartRates"></canvas></div>
        </div>
    `;
    drawSinglePeriodTreemap(data, systems, getActiveLevels(data));
}}

function drawSinglePeriodTreemap(data, systems, levels) {{
    drawTreemapCore('treemapCanvas', data, systems, levels, false);
}}

// ==================== 树状图核心渲染 ====================
function drawTreemapCore(canvasId, data, systems, levels, isMultiPeriod) {{
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.parentElement.getBoundingClientRect();
    const H = isMultiPeriod ? 450 : 500;
    canvas.width = rect.width * dpr;
    canvas.height = H * dpr;
    canvas.style.width = rect.width + 'px';
    canvas.style.height = H + 'px';
    ctx.scale(dpr, dpr);
    const W = rect.width;
    const pad = 3;
    
    // 计算系统数据
    function buildSysData(items) {{
        return items.map(sys => {{
            const rows = data.filter(r => r.attr === sys);
            const total = rows.reduce((s, r) => s + r['告警记录数'], 0);
            const levelCounts = {{}};
            rows.forEach(r => {{ levelCounts[r['告警级别']] = (levelCounts[r['告警级别']] || 0) + r['告警记录数']; }});
            const mainLevel = Object.entries(levelCounts).sort((a, b) => b[1] - a[1])[0][0];
            return {{ sys, total, mainLevel, levelCounts }};
        }}).filter(s => s.total > 0).sort((a, b) => b.total - a.total);
    }}
    
    const sysData = buildSysData(systems);
    if (sysData.length === 0) return;
    
    // 布局函数
    function layout(items, x, y, w, h) {{
        if (items.length === 0) return;
        if (items.length === 1) {{
            items[0]._x = x; items[0]._y = y; items[0]._w = w; items[0]._h = h; return;
        }}
        const total = items.reduce((s, i) => s + i.total, 0);
        if (w >= h) {{
            let cx = x;
            items.forEach(item => {{
                const iw = (item.total / total) * w;
                item._x = cx; item._y = y; item._w = iw; item._h = h; cx += iw;
            }});
        }} else {{
            let cy = y;
            items.forEach(item => {{
                const ih = (item.total / total) * h;
                item._x = x; item._y = cy; item._w = w; item._h = ih; cy += ih;
            }});
        }}
    }}
    
    function splitIntoRows(items, maxAR, areaW, areaH, totalVal) {{
        if (items.length <= 3) return [items];
        const rows = [];
        let cur = [], curTotal = 0;
        items.forEach(item => {{
            cur.push(item); curTotal += item.total;
            const rw = (curTotal / totalVal) * areaW;
            const estH = areaH / (rows.length + 1);
            const aspects = cur.map(ci => {{ const cw = (ci.total / curTotal) * rw; return Math.max(cw / estH, estH / cw); }});
            if (Math.max(...aspects) > (maxAR || 3) && cur.length > 1) {{
                cur.pop(); rows.push(cur); cur = [item]; curTotal = item.total;
            }}
        }});
        if (cur.length > 0) rows.push(cur);
        return rows;
    }}
    
    function drawRects(sdList, ctxRef) {{
        sdList.forEach(sd => {{
            const rx = sd._x + pad, ry = sd._y + pad, rw = sd._w - pad * 2, rh = sd._h - pad * 2;
            if (rw <= 0 || rh <= 0) return;
            ctxRef.fillStyle = LEVEL_COLORS[sd.mainLevel] || '#2563eb';
            ctxRef.fillRect(rx, ry, rw, rh);
            ctxRef.strokeStyle = '#fff'; ctxRef.lineWidth = 1.2;
            ctxRef.strokeRect(rx, ry, rw, rh);
            ctxRef.fillStyle = '#fff';
            ctxRef.font = (rw > 100 ? 'bold 12px' : 'bold 9px') + ' -apple-system,sans-serif';
            ctxRef.textBaseline = 'middle'; ctxRef.textAlign = 'center';
            const cnt = sd.total.toLocaleString();
            if (rh > 36 && rw > 55) {{
                ctxRef.fillText(sd.sys, rx + rw / 2, ry + rh / 2 - 7);
                ctxRef.font = '10px -apple-system,sans-serif';
                ctxRef.fillText(cnt, rx + rw / 2, ry + rh / 2 + 9);
            }} else if (rw > 35 && rh > 18) {{
                ctxRef.fillText(cnt, rx + rw / 2, ry + rh / 2);
            }}
        }});
    }}
    
    const grandTotal = sysData.reduce((s, r) => s + r.total, 0);
    const rows = splitIntoRows(sysData, 4, W, H, grandTotal);
    const rowH = H / rows.length;
    rows.forEach((row, ri) => layout(row, 0, ri * rowH, W, rowH));
    
    drawRects(sysData, ctx);
    
    // 图例
    const legendDiv = document.getElementById('treemapLegend');
    if (legendDiv) {{
        legendDiv.innerHTML = levels.map(l => `
            <div class="treemap-legend-item">
                <div class="treemap-legend-color" style="background:${{LEVEL_COLORS[l] || '#999'}}"></div>
                <span>${{l}}</span>
            </div>
        `).join('');
    }}
    
    // Tooltip
    const tooltip = document.getElementById('treemapTooltip');
    if (tooltip) {{
        canvas.onmousemove = function(e) {{
            const cr = canvas.getBoundingClientRect();
            const mx = e.clientX - cr.left, my = e.clientY - cr.top;
            let found = null;
            for (const sd of sysData) {{
                if (mx >= sd._x && mx <= sd._x + sd._w && my >= sd._y && my <= sd._y + sd._h) {{ found = sd; break; }}
            }}
            if (found) {{
                const parts = Object.entries(found.levelCounts).map(([k, v]) => k + ': ' + v.toLocaleString()).join('<br>');
                tooltip.innerHTML = '<strong>' + found.sys + '</strong><br>总计: ' + found.total.toLocaleString() + '<br>' + parts;
                tooltip.style.display = 'block';
                tooltip.style.left = (mx + 15) + 'px'; tooltip.style.top = (my + 15) + 'px';
            }} else {{ tooltip.style.display = 'none'; }}
        }};
        canvas.onmouseleave = function() {{ tooltip.style.display = 'none'; }};
    }}
}}

// ==================== 小多组柱状图（系统为行，周期为列，按告警级别叠加） ====================
function buildSmallMultiplesGrid(container, data, systems, periods) {{
    destroyAllCharts();
    if (periods.length === 0 || systems.length === 0) return;
    if (periods.length > 5) periods = periods.slice(0, 5);
    
    const levels = getActiveLevels(data);
    const nCols = Math.min(systems.length, 4);
    
    let html = `<div class="treemap-card">
        <h3>系统告警分布 &mdash; 按周期分列</h3>
        <div style="display:grid;grid-template-columns:repeat(${{nCols}},1fr);gap:14px;margin-bottom:10px;">`;
    
    systems.forEach((sys, idx) => {{
        html += `
            <div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 10px 8px;">
                <div style="font-size:13px;font-weight:600;color:var(--ink);text-align:center;margin-bottom:6px;">${{sys}}</div>
                <div style="position:relative;height:220px;"><canvas id="smallMulti_${{idx}}"></canvas></div>
            </div>`;
    }});
    
    html += `</div>
        <div class="treemap-legend" id="treemapLegend"></div>
        </div>
        <div class="chart-card" style="margin-top:16px;">
            <h3>系统未超时处理率</h3>
            <div class="chart-wrapper"><canvas id="chartRates"></canvas></div>
        </div>`;
    container.innerHTML = html;
    
    // 图例
    const legendDiv = document.getElementById('treemapLegend');
    if (legendDiv) {{
        legendDiv.innerHTML = levels.map(l => `
            <div class="treemap-legend-item">
                <div class="treemap-legend-color" style="background:${{LEVEL_COLORS[l] || '#999'}}"></div>
                <span>${{l}}</span>
            </div>
        `).join('');
    }}
    
    // 为每个系统创建叠加柱状图（x轴=周期）
    systems.forEach((sys, idx) => {{
        const datasets = levels.map((lvl, i) => {{
            const colors = ['#d93025','#e37400','#2563eb','#059669','#d97706','#9e9e9e'];
            return {{
                label: lvl,
                data: periods.map(p => {{
                    const rows = data.filter(r => r.attr === sys && r['周期'] === p && r['告警级别'] === lvl);
                    return rows.reduce((s, r) => s + r['告警记录数'], 0);
                }}),
                backgroundColor: LEVEL_COLORS[lvl] || colors[i % colors.length],
                borderRadius: 0,
            }};
        }});
        
        const ctx = document.getElementById('smallMulti_' + idx).getContext('2d');
        chartInstances['small_' + idx] = new Chart(ctx, {{
            type: 'bar',
            data: {{ labels: periods, datasets: datasets }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }},
                    tooltip: {{ callbacks: {{ label: c => c.dataset.label + ': ' + c.raw.toLocaleString() }} }}
                }},
                scales: {{
                    x: {{ stacked: true, ticks: {{ maxRotation: 40, font: {{ size: 10 }} }} }},
                    y: {{ stacked: true, beginAtZero: true, ticks: {{ font: {{ size: 10 }}, callback: v => v >= 1000 ? (v/1000).toFixed(1)+'k' : v }} }}
                }}
            }}
        }});
    }});
}}

// ==================== 单张叠加柱状图（多系统） ====================
function buildSingleStackedBar(container, data, systems) {{
    destroyAllCharts();
    const topSystems = systems.slice(0, 12);
    const sdata = getStackedDataBySystems(data, topSystems);
    
    container.innerHTML = `
        <div class="charts-grid">
            <div class="chart-card">
                <h3>系统告警数量分布（按告警级别分层）</h3>
                <div class="chart-wrapper"><canvas id="chartAlarmCount"></canvas></div>
            </div>
            <div class="chart-card">
                <h3>系统未超时处理率</h3>
                <div class="chart-wrapper"><canvas id="chartRates"></canvas></div>
            </div>
        </div>
    `;
    
    const ctx = document.getElementById('chartAlarmCount').getContext('2d');
    chartInstances['alarmCount'] = new Chart(ctx, {{
        type: 'bar',
        data: {{ labels: sdata.systems, datasets: sdata.datasets }},
        options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{ 
                legend: {{ display: true, position: 'top' }},
                tooltip: {{ callbacks: {{ label: ctx => ctx.dataset.label + ': ' + ctx.raw.toLocaleString() }} }}
            }},
            scales: {{
                x: {{ stacked: true, ticks: {{ maxRotation: 45, font: {{ size: 11 }} }} }},
                y: {{ stacked: true, beginAtZero: true, ticks: {{ callback: v => v.toLocaleString() }} }}
            }}
        }}
    }});
}}

// ==================== 未超时处理率图 ====================
function renderRatesChart(data, systems) {{
    if (chartInstances['rates']) chartInstances['rates'].destroy();
    const ctx = document.getElementById('chartRates');
    if (!ctx) return;
    
    const agg = systems.map(sys => {{
        const rows = data.filter(r => r.attr === sys);
        const totalAlarms = rows.reduce((s, r) => s + r['告警记录数'], 0);
        const onTime = rows.reduce((s, r) => s + r['未超时处理记录数'], 0);
        return {{
            sys,
            oRate: totalAlarms > 0 ? +(onTime / totalAlarms * 100).toFixed(1) : 0
        }};
    }});
    
    chartInstances['rates'] = new Chart(ctx.getContext('2d'), {{
        type: 'bar',
        data: {{
            labels: agg.map(r => r.sys),
            datasets: [
                {{ label: '未超时率(%)', data: agg.map(r => r.oRate), backgroundColor: 'rgba(13,144,79,0.7)', borderRadius: 4 }}
            ]
        }},
        options: {{
            responsive: true,
            maintainAspectRatio: false,
            plugins: {{ legend: {{ display: false }} }},
            scales: {{
                y: {{ beginAtZero: true, max: 100, ticks: {{ callback: v => v + '%' }} }},
                x: {{ ticks: {{ maxRotation: 45, font: {{ size: 11 }} }} }}
            }}
        }}
    }});
}}

// ==================== TOP告警处理时长 ====================
function toggleTopPanel() {{
    const body = document.getElementById('topCollapseBody');
    const arrow = document.getElementById('topArrow');
    if (body.style.display === 'none') {{
        body.style.display = '';
        arrow.classList.remove('collapsed');
    }} else {{
        body.style.display = 'none';
        arrow.classList.add('collapsed');
    }}
}}

function parseDealTime(val) {{
    if (!val || val === '') return 0;
    if (typeof val === 'number') return val;
    const s = String(val).trim();
    // 数字字符串
    if (/^\d+(\\.\d+)?$/.test(s)) return parseFloat(s);
    // h/m 格式
    let mins = 0;
    const hMatch = s.match(/(\d+(?:\\.\d+)?)\s*(?:h|小时|时)/);
    if (hMatch) mins += parseFloat(hMatch[1]) * 60;
    const mMatch = s.match(/(\d+(?:\\.\d+)?)\s*(?:m|分钟|分)/);
    if (mMatch) mins += parseFloat(mMatch[1]);
    if (mins > 0) return mins;
    // 冒号格式 HH:MM
    const tm = s.match(/^(\d+):(\d{2})$/);
    if (tm) return parseInt(tm[1]) * 60 + parseInt(tm[2]);
    // 尝试直接转
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
}}

function getFilteredTopData(n) {{
    return RAW_EXPORT
        .filter(row => {{
            if (selectedSystems.length > 0 && selectedSystems.length < ALL_SYSTEMS.length && !selectedSystems.includes(row.attr)) return false;
            if (selectedPeriods.length > 0 && selectedPeriods.length < ALL_PERIODS.length && !selectedPeriods.includes(row['周期'])) return false;
            if (selectedLevels.length > 0 && selectedLevels.length < ALL_LEVELS.length && !selectedLevels.includes(row['告警级别'])) return false;
            return true;
        }})
        .map(row => ({{ ...row, _dealMs: parseDealTime(row['deal_time']) }}))
        .sort((a, b) => b._dealMs - a._dealMs)
        .slice(0, n);
}}

function formatDealTime(ms) {{
    if (ms <= 0) return '-';
    if (ms < 60) return ms + '分钟';
    const h = Math.floor(ms / 60);
    const m = Math.round(ms % 60);
    return m > 0 ? h + 'h ' + m + 'm' : h + '小时';
}}

function renderTop() {{
    const n = parseInt(document.getElementById('topCount').value) || 20;
    const data = getFilteredTopData(n);
    const tbody = document.getElementById('topBody');
    if (data.length === 0) {{
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--muted);">暂无数据</td></tr>';
        return;
    }}
    tbody.innerHTML = data.map(r => {{
        const content = (r['alarm_content'] || '-');
        const shortContent = content.length > 40 ? content.substring(0, 40) + '...' : content;
        return `<tr>
            <td>${{r.attr || '-'}}</td>
            <td>${{r['告警级别'] || '-'}}</td>
            <td>${{r.RESOURCE_NAME || '-'}}</td>
            <td title="${{content.replace(/"/g,'&quot;')}}">${{shortContent}}</td>
            <td>${{formatDealTime(r._dealMs)}}</td>
            <td style="font-size:12px">${{r.alarm_first_time || '-'}}</td>
            <td>${{r.deal_status || '-'}}</td>
        </tr>`;
    }}).join('');
}}

function exportTopCustom() {{
    const n = parseInt(document.getElementById('topCount').value) || 20;
    const data = getFilteredTopData(n);
    if (data.length === 0) {{ alert('没有可导出的数据'); return; }}

    const modal = document.createElement('div');
    modal.className = 'modal-overlay show';
    const availableKeys = new Set(Object.keys(data[0]));
    const fields = EXPORT_FIELDS.filter(f => availableKeys.has(f.key));
    modal.innerHTML = `
        <div class="modal-content">
            <div class="modal-header">
                <h4>导出 TOP${{n}} 告警（共 ${{data.length}} 条）</h4>
                <button class="modal-close" onclick="this.closest('.modal-overlay').remove()">&times;</button>
            </div>
            <div class="modal-body">
                <div class="select-all-row">
                    <a onclick="document.querySelectorAll('.topExportCb').forEach(c=>c.checked=true)">全选</a>
                    <a onclick="document.querySelectorAll('.topExportCb').forEach(c=>c.checked=false)" style="margin-left:10px">取消全选</a>
                </div>
                <div class="field-grid">
                    ${{fields.map(f => `<label class="multi-option">
                        <input type="checkbox" class="topExportCb" value="${{f.key}}" checked>
                        <span>${{f.label}}</span>
                    </label>`).join('')}}
                </div>
            </div>
            <div class="modal-footer">
                <button class="btn btn-cancel" onclick="this.closest('.modal-overlay').remove()">关闭</button>
                <button class="btn btn-confirm" onclick="doTopExport(this)">导出 CSV</button>
            </div>
        </div>`;
    modal.dataset.topData = JSON.stringify(data.map(r => {{
        const copy = {{...r}}; delete copy._dealMs; return copy;
    }}));
    document.body.appendChild(modal);
    modal.onclick = function(e) {{ if (e.target === modal) modal.remove(); }};
}}

function doTopExport(btn) {{
    const modal = btn.closest('.modal-overlay');
    const data = JSON.parse(modal.dataset.topData);
    const checks = modal.querySelectorAll('.topExportCb:checked');
    if (checks.length === 0) {{ alert('请至少选择一个字段'); return; }}
    const selectedKeys = Array.from(checks).map(c => c.value);
    const selectedFields = EXPORT_FIELDS.filter(f => selectedKeys.includes(f.key));
    const cols = selectedFields.map(f => f.label);
    let csv = '\\uFEFF' + cols.join(',') + '\\n';
    data.forEach(r => {{
        csv += selectedFields.map(f => {{
            let v = r[f.key] !== undefined ? String(r[f.key]) : '';
            if (v.includes(',') || v.includes('"') || v.includes('\\n')) v = '"' + v.replace(/"/g, '""') + '"';
            return v;
        }}).join(',') + '\\n';
    }});
    const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'TOP告警_' + new Date().toISOString().slice(0,10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
    modal.remove();
}}

// ==================== 工具函数 ====================
function destroyAllCharts() {{
    Object.values(chartInstances).forEach(c => {{ if (c) c.destroy(); }});
    chartInstances = {{}};
}}

// ==================== 排序 ====================
function sortTable(col) {{
    if (sortColumn === col) {{
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
    }} else {{
        sortColumn = col;
        sortDirection = 'desc';
    }}
    applyAllFilters();
    updateSortIndicator();
}}

function sortData(data, col, dir) {{
    return [...data].sort((a, b) => {{
        let va = a[col], vb = b[col];
        if (typeof va === 'string') return dir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
        return dir === 'asc' ? va - vb : vb - va;
    }});
}}

function updateSortIndicator() {{
    document.querySelectorAll('th').forEach(th => {{
        th.classList.remove('sorted', 'sorted-asc', 'sorted-desc');
    }});
    document.querySelectorAll('th').forEach(th => {{
        if (th.textContent.includes(sortColumn) || th.textContent.trim().startsWith(sortColumn)) {{
            th.classList.add('sorted', 'sorted-' + (sortDirection === 'asc' ? 'asc' : 'desc'));
        }}
    }});
}}

// ==================== 汇总卡片更新 ====================
function updateSummaryCards(data) {{
    const total = data.reduce((s, r) => s + r['告警记录数'], 0);
    const onTime = data.reduce((s, r) => s + r['未超时处理记录数'], 0);
    const oRate = total > 0 ? (onTime / total * 100).toFixed(2) : '0.00';
    const systems = new Set(data.map(r => r.attr)).size;
    document.getElementById('sumTotal').textContent = total.toLocaleString();
    document.getElementById('sumOnTime').textContent = onTime.toLocaleString();
    document.getElementById('sumOnTimeRate').textContent = oRate + '%';
    document.getElementById('sumSystems').textContent = systems;
}}

// ==================== 表格渲染 ====================
function renderTable(data) {{
    const totalPages = Math.ceil(data.length / PAGE_SIZE);
    if (currentPage > totalPages) currentPage = totalPages || 1;
    const start = (currentPage - 1) * PAGE_SIZE;
    const pageData = data.slice(start, start + PAGE_SIZE);
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = pageData.map(r => {{
        const pRate = r['处理率'], oRate = r['未超时率'];
        const pBadge = pRate >= 90 ? 'rate-high' : pRate >= 60 ? 'rate-mid' : 'rate-low';
        const oBadge = oRate >= 90 ? 'rate-high' : oRate >= 60 ? 'rate-mid' : 'rate-low';
        return `<tr>
            <td><strong>${{r.attr}}</strong></td><td>${{r['告警级别'] || '-'}}</td><td>${{r['周期'] || '-'}}</td>
            <td>${{r['告警记录数'].toLocaleString()}}</td><td>${{r['未超时处理记录数'].toLocaleString()}}</td>
            <td><span class="rate-badge ${{pBadge}}">${{pRate.toFixed(1)}}%</span></td><td><span class="rate-badge ${{oBadge}}">${{oRate.toFixed(1)}}%</span></td>
        </tr>`;
    }}).join('');
    document.getElementById('tableInfo').textContent = '显示 ' + (data.length > 0 ? start + 1 : 0) + '-' + Math.min(start + PAGE_SIZE, data.length) + ' / 共 ' + data.length + ' 条';
    let pagHTML = '';
    pagHTML += `<button onclick="goPage(${{currentPage - 1}})" ${{currentPage === 1 ? 'disabled' : ''}}>‹ 上一页</button>`;
    for (let i = 1; i <= totalPages; i++) {{
        if (totalPages <= 10 || i === 1 || i === totalPages || Math.abs(i - currentPage) <= 2) {{
            pagHTML += `<button onclick="goPage(${{i}})" class="${{i === currentPage ? 'active' : ''}}">${{i}}</button>`;
        }} else if (i === 2 || i === totalPages - 1) {{
            pagHTML += '<span class="page-info">...</span>';
        }}
    }}
    pagHTML += `<button onclick="goPage(${{currentPage + 1}})" ${{currentPage === totalPages ? 'disabled' : ''}}>下一页 ›</button>`;
    document.getElementById('pagination').innerHTML = pagHTML;
    updateSortIndicator();
}}

function goPage(p) {{ currentPage = p; applyAllFilters(); }}

// ==================== 辅助功能 ====================
function resetFilters() {{
    selectedSystems = [...ALL_SYSTEMS];
    selectedPeriods = [...ALL_PERIODS];
    selectedLevels = [...ALL_LEVELS];
    document.getElementById('searchInput').value = '';
    // 更新多选UI
    document.querySelectorAll('#systemOptions input[type="checkbox"]').forEach(cb => cb.checked = true);
    document.querySelectorAll('#periodOptions input[type="checkbox"]').forEach(cb => cb.checked = true);
    document.querySelectorAll('#levelOptions input[type="checkbox"]').forEach(cb => cb.checked = true);
    updateMultiSelectDisplay('multiSystem', 'systemText', 'systemCount', selectedSystems, ALL_SYSTEMS, '系统');
    updateMultiSelectDisplay('multiPeriod', 'periodText', 'periodCount', selectedPeriods, ALL_PERIODS, '周期');
    updateMultiSelectDisplay('multiLevel', 'levelText', 'levelCount', selectedLevels, ALL_LEVELS, '级别');
    sortColumn = '告警记录数';
    sortDirection = 'desc';
    currentPage = 1;
    applyAllFilters();
}}

function exportTable() {{
    // 导出当前筛选的原始告警明细（默认全部字段）
    const data = getFilteredRawData();
    const allKeys = EXPORT_FIELDS.filter(f => {{ const r = data[0]; return r && r.hasOwnProperty(f.key); }});
    const cols = allKeys.map(f => f.label);
    const keys = allKeys.map(f => f.key);
    let csv = '\\uFEFF' + cols.join(',') + '\\n';
    data.forEach(r => {{
        csv += keys.map(k => {{
            let v = r[k] !== undefined ? String(r[k]) : '';
            if (v.includes(',') || v.includes('"') || v.includes('\\n')) v = '"' + v.replace(/"/g, '""') + '"';
            return v;
        }}).join(',') + '\\n';
    }});
    const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '告警明细_' + new Date().toISOString().slice(0,10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
}}

document.addEventListener('DOMContentLoaded', init);

// ==================== 自定义导出（基于原始清洗数据） ====================

function getFilteredRawData() {{
    return RAW_EXPORT.filter(row => {{
        if (selectedSystems.length > 0 && selectedSystems.length < ALL_SYSTEMS.length && !selectedSystems.includes(row.attr)) return false;
        if (selectedPeriods.length > 0 && selectedPeriods.length < ALL_PERIODS.length && !selectedPeriods.includes(row['周期'])) return false;
        if (selectedLevels.length > 0 && selectedLevels.length < ALL_LEVELS.length && !selectedLevels.includes(row['告警级别'])) return false;
        return true;
    }});
}}

function buildExportFieldGrid() {{
    const grid = document.getElementById('exportFieldGrid');
    const availableKeys = new Set();
    if (RAW_EXPORT.length > 0) {{
        Object.keys(RAW_EXPORT[0]).forEach(k => availableKeys.add(k));
    }}
    
    const fields = EXPORT_FIELDS.filter(f => availableKeys.has(f.key));
    grid.innerHTML = fields.map(f => `
        <label class="multi-option">
            <input type="checkbox" value="${{f.key}}" checked>
            <span>${{f.label}}</span>
        </label>
    `).join('');
}}

function openCustomExport() {{
    const data = getFilteredRawData();
    document.getElementById('exportRecordCount').textContent = data.length.toLocaleString();
    buildExportFieldGrid();
    document.getElementById('exportModal').classList.add('show');
}}

function closeCustomExport() {{
    document.getElementById('exportModal').classList.remove('show');
}}

function toggleAllExportFields(select) {{
    document.querySelectorAll('#exportFieldGrid input[type="checkbox"]').forEach(cb => cb.checked = select);
}}

function doCustomExport() {{
    const checks = document.querySelectorAll('#exportFieldGrid input[type="checkbox"]:checked');
    if (checks.length === 0) {{ alert('请至少选择一个导出字段'); return; }}
    
    const selectedKeys = Array.from(checks).map(c => c.value);
    const selectedFields = EXPORT_FIELDS.filter(f => selectedKeys.includes(f.key));
    
    const data = getFilteredRawData();
    
    const cols = selectedFields.map(f => f.label);
    let csv = '\\uFEFF' + cols.join(',') + '\\n';
    data.forEach(r => {{
        csv += selectedFields.map(f => {{
            let v = r[f.key] !== undefined ? String(r[f.key]) : '';
            if (v.includes(',') || v.includes('"') || v.includes('\\n')) {{
                v = '"' + v.replace(/"/g, '""') + '"';
            }}
            return v;
        }}).join(',') + '\\n';
    }});
    
    const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '告警明细_自定义导出_' + new Date().toISOString().slice(0,10) + '.csv';
    a.click();
    URL.revokeObjectURL(url);
    closeCustomExport();
}}

// 点击遮罩关闭弹窗
document.addEventListener('click', function(e) {{
    if (e.target.id === 'exportModal') closeCustomExport();
}});
</script>
</body>
</html>'''
    
    # 写入HTML文件
    try:
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write(html_content)
        logging.info(f"交互式HTML报告已生成: {report_file}")
        return report_file
    except Exception as e:
        logging.error(f"生成HTML报告时出错: {str(e)}")
        logging.error(traceback.format_exc())
        return None

def setup_logging(log_level='INFO'):
    """设置日志记录"""
    log_dir = 'logs'
    os.makedirs(log_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'weekly_statistics_{timestamp}.log')
    
    # 配置日志格式和级别
    logging.basicConfig(
        level=getattr(logging, log_level),
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return log_file

def create_combined_weekly_statistics_with_chinese_names():
    """
    将所有attr的每周告警统计整合到一个sheet页中，并将attr替换为中文名称
    """
    # 解析命令行参数
    args = parse_arguments()
    
    # 设置日志记录
    log_file = setup_logging(args.log_level)
    logging.info("开始整合每周告警统计数据...")
    logging.info(f"日志文件路径: {log_file}")
    logging.info(f"命令行参数: {args}")
    
    # 记录映射关系
    logging.info(f"系统映射关系: {len(config.ATTR_CHINESE_MAPPING)}个系统")
    
    # 输入文件路径
    input_file = args.input_file
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        logging.error(f"错误: 输入文件 {input_file} 不存在")
        return
    
    # 输出文件路径
    output_dir = args.output_dir
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, args.output_file)
    
    # 创建中间结果目录
    temp_dir = args.temp_dir if args.temp_dir else os.path.join(output_dir, 'temp')
    os.makedirs(temp_dir, exist_ok=True)
    
    # 读取Excel文件
    logging.info(f"读取文件: {input_file}")
    try:
        # 检查文件大小
        file_size = os.path.getsize(input_file)
        logging.info(f"输入文件大小: {file_size} 字节")
        
        # 尝试读取Excel文件
        logging.info("尝试打开Excel文件...")
        start_time = time.time()
        with pd.ExcelFile(input_file, engine='openpyxl') as xls:
            # 获取所有sheet名称
            sheet_names = xls.sheet_names
            logging.info(f"Excel文件包含以下工作表: {sheet_names}")
            
            # 筛选出每周告警统计的sheet
            weekly_sheets = [sheet for sheet in sheet_names if sheet.startswith('每周告警统计')]
            logging.info(f"找到以下每周告警统计工作表: {weekly_sheets}")
            
            if not weekly_sheets:
                logging.warning("警告: 没有找到任何每周告警统计工作表")
                return
            
            # 创建一个空的DataFrame来存储合并后的数据
            combined_df = pd.DataFrame()
            
            # 处理周期
            periods = None
            if args.custom_periods:
                # 解析自定义周期
                period_ranges = [p.strip() for p in args.custom_periods.split(',')]
                periods = []
                for pr in period_ranges:
                    start, end = pr.split(' - ')
                    # 转换为数据中的周期格式(mm.dd-mm.dd)
                    start_date = pd.to_datetime(start)
                    end_date = pd.to_datetime(end)
                    period_str = f"{start_date.strftime('%m.%d')}-{end_date.strftime('%m.%d')}"
                    periods.append(period_str)
                logging.info(f"使用自定义周期范围: {periods}")
            
            # 遍历每个每周告警统计sheet
            total_sheets = len(weekly_sheets)
            
            # 创建进度条（如果未禁用）
            sheet_iterator = weekly_sheets
            if not args.no_progress:
                sheet_iterator = tqdm(weekly_sheets, desc="处理工作表", unit="sheet")
            
            for idx, sheet_name in enumerate(sheet_iterator, 1):
                sheet_start_time = time.time()
                logging.info(f"处理工作表 [{idx}/{total_sheets}]: {sheet_name}")
                if not args.no_progress:
                    sheet_iterator.set_description(f"处理 {sheet_name}")
                try:
                    # 读取sheet数据
                    df = pd.read_excel(xls, sheet_name)
                    logging.info(f"  工作表 {sheet_name} 包含 {len(df)} 行数据")
                    logging.debug(f"  列名: {df.columns.tolist()}")
                    
                    # 数据验证 - 检查必要的列
                    required_columns = ['attr', 'ALARM_COUNT']
                    missing_columns = [col for col in required_columns if col not in df.columns]
                    if missing_columns:
                        logging.warning(f"  工作表 {sheet_name} 中缺少以下必要列: {missing_columns}")
                    
                    # 筛选符合周期的数据
                    if periods:
                        df = df[df['周期'].isin(periods)]
                        if df.empty:
                            logging.warning(f"  工作表 {sheet_name} 中没有数据匹配指定周期范围")
                            continue
                    
                    # 确保包含attr列
                    if 'attr' not in df.columns:
                        logging.warning(f"  工作表 {sheet_name} 中没有找到'attr'列，尝试从工作表名称中提取")
                        # 从sheet名称中提取attr
                        match = re.search(r'每周告警统计_(.+)', sheet_name)
                        if match:
                            attr_value = match.group(1)
                            logging.info(f"  从工作表名称中提取到attr值: {attr_value}")
                            df['attr'] = attr_value
                        else:
                            logging.warning(f"  无法从工作表名称中提取attr值，使用'unknown'")
                            df['attr'] = 'unknown'
                    
                    # 确保数值列不为空
                    # 确保ALARM_COUNT列存在且为数值类型
                    if 'ALARM_COUNT' not in df.columns:
                        df['ALARM_COUNT'] = 1
                        logging.warning(f"  工作表 {sheet_name} 中没有找到'ALARM_COUNT'列，使用默认值1")
                    
                    # 数据清洗 - 转换ALARM_COUNT为数值类型并处理异常值
                    df['ALARM_COUNT'] = pd.to_numeric(df['ALARM_COUNT'], errors='coerce')
                    
                    # 检测并处理异常值
                    alarm_count_mean = df['ALARM_COUNT'].mean()
                    alarm_count_std = df['ALARM_COUNT'].std()
                    upper_limit = alarm_count_mean + 3 * alarm_count_std
                    outliers = df[df['ALARM_COUNT'] > upper_limit]
                    
                    if not outliers.empty:
                        logging.warning(f"  检测到 {len(outliers)} 个ALARM_COUNT异常值 (>{upper_limit:.2f})")
                        for _, row in outliers.iterrows():
                            logging.debug(f"  异常值: attr={row['attr']}, ALARM_COUNT={row['ALARM_COUNT']}")
                    
                    # 填充缺失值
                    df['ALARM_COUNT'] = df['ALARM_COUNT'].fillna(1)
                    
                    # 设置告警数量为1（每行代表一条记录）
                    df['告警数量'] = 1
                    
                    # 打印调试信息
                    logging.info(f"  记录数统计: {len(df)} 行")
                    
                    # 使用进度条处理数据行
                    logging.info(f"  计算已处理数量和未超时处理数量...")
                    
                    # 创建数据处理进度条（如果未禁用）
                    if not args.no_progress and len(df) > 100:  # 只有当行数超过100时才显示进度条
                        row_iterator = tqdm(df.iterrows(), total=len(df), desc=f"处理{sheet_name}数据行", unit="行", leave=False)
                    else:
                        row_iterator = df.iterrows()
                    
                    # 初始化计数器
                    processed_count = 0
                    on_time_count = 0
                    
                    # 处理已处理数量
                    if 'deal_status' in df.columns:
                        for _, row in row_iterator:
                            if row['deal_status'] == '已处理':
                                processed_count += 1
                            elif row['deal_status'] == '未超时':
                                on_time_count += 1
                        
                        # 对于已处理的记录，使用1；否则为0
                        df['已处理数量'] = df.apply(
                            lambda row: 1 if row['deal_status'] == '已处理' else 0, 
                            axis=1
                        )
                        
                        # 对于未超时的记录，使用1；否则为0
                        df['未超时处理数量'] = df.apply(
                            lambda row: 1 if row['deal_status'] == '未超时' else 0, 
                            axis=1
                        )
                        
                        # 打印调试信息
                        logging.info(f"  已处理记录数: {len(df[df['deal_status'] == '已处理'])}, 已处理告警记录数: {processed_count}")
                        logging.info(f"  未超时记录数: {len(df[df['deal_status'] == '未超时'])}, 未超时处理告警记录数: {on_time_count}")
                    else:
                        # 尝试从其他列推导已处理数量
                        if 'alarm_state' in df.columns:
                            df['已处理数量'] = df['alarm_state'].apply(lambda x: 1 if x == '已处理' else 0)
                            logging.warning(f"  没有deal_status列，从alarm_state推导已处理数量")
                        elif 'ack_status' in df.columns:
                            df['已处理数量'] = df['ack_status'].apply(lambda x: 1 if x == '已确认' else 0)
                            logging.warning(f"  没有deal_status列，从ack_status推导已处理数量")
                        else:
                            df['已处理数量'] = 0
                            logging.warning(f"  无法推导已处理数量，设为0")
                        
                        # 未超时处理数量暂时设为0
                        df['未超时处理数量'] = 0
                    
                    # 修复列名中的换行符
                    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
                    
                    # 将数据添加到合并的DataFrame中
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                    logging.info(f"  成功添加工作表 {sheet_name} 的数据，当前合并数据包含 {len(combined_df)} 行")
                    
                    # 保存中间结果
                    temp_file = os.path.join(temp_dir, f"temp_combined_{idx}_{total_sheets}.xlsx")
                    combined_df.to_excel(temp_file, index=False)
                    logging.info(f"  已保存中间结果到: {temp_file}")
                    
                    # 计算并显示处理时间
                    sheet_time = time.time() - sheet_start_time
                    logging.info(f"  工作表 {sheet_name} 处理完成，耗时: {sheet_time:.2f}秒 ({idx}/{total_sheets})")
                
                except Exception as e:
                    logging.error(f"处理工作表 {sheet_name} 时出错: {str(e)}")
                    logging.error("详细错误信息:")
                    logging.error(traceback.format_exc())
                    continue
            
            if combined_df.empty:
                logging.error("错误: 合并后的数据为空，无法继续处理")
                return
            
            # 计算总处理时间
            total_time = time.time() - start_time
            logging.info(f"所有工作表处理完成，耗时: {total_time:.2f}秒")
            logging.info(f"合并数据包含 {len(combined_df)} 行")
            logging.info(f"合并数据的列: {combined_df.columns.tolist()}")
            
            # 替换attr为中文名称
            logging.info("替换attr为中文名称...")
            # 替换attr为中文名称（NaN/空值保留原值，未匹配的保留原值）
            combined_df['attr'] = combined_df['attr'].apply(
                lambda x: x if pd.isna(x) or str(x).strip() == '' 
                else config.ATTR_CHINESE_MAPPING.get(str(x).strip().lower(), x)
            )
            
            # 告警级别中英文映射（未匹配的保留原值）
            if '告警级别' in combined_df.columns:
                combined_df['告警级别'] = combined_df['告警级别'].apply(
                    lambda x: x if pd.isna(x) or str(x).strip() == ''
                    else config.LEVEL_CHINESE_MAPPING.get(str(x).strip().lower(), x)
                )
            
           # 确保包含必要列
            required_cols = ['attr', '周期', '告警级别', '告警数量', '已处理数量', '未超时处理数量']
            for col in required_cols:
                if col not in combined_df.columns:
                    logging.warning(f"警告: 合并数据中缺少'{col}'列，添加空列")
                    combined_df[col] = ''
                    
                # 确保周期列格式统一为"YYYY-MM-DD - YYYY-MM-DD"
                if '周期' in combined_df.columns and 'week_period' in combined_df.columns:
                    combined_df['周期'] = combined_df.apply(
                        lambda row: row['周期'] if re.match(r'\d{4}-\d{2}-\d{2} - \d{4}-\d{2}-\d{2}', str(row['周期'])) 
                        else (
                            row['week_period'].replace('/', ' - ') 
                            if re.match(r'\d{4}-\d{2}-\d{2}/\d{4}-\d{2}-\d{2}', str(row['week_period'])) 
                            else None  # 标记为无效
                        ),
                        axis=1
                    )
                    # 跳过无效周期值
                    combined_df = combined_df[combined_df['周期'].notna()]
            
            # 重新排列列，将attr放在前面
            columns = ['attr'] + [col for col in combined_df.columns if col != 'attr']
            combined_df = combined_df[columns]
            
            # 按attr和周期排序
            logging.info("按attr、周期和告警级别排序...")
            combined_df.sort_values(['attr', '周期', '告警级别'], inplace=True)
            
            # 将合并后的数据写入新的Excel文件
            logging.info(f"写入合并后的数据到: {output_file}")
            try:
                # 如果文件已存在，先删除
                if os.path.exists(output_file):
                    logging.info(f"文件 {output_file} 已存在，先删除")
                    os.remove(output_file)
                
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    combined_df.to_excel(writer, sheet_name='每周告警统计汇总', index=False)
                    
                    # 调整列宽
                    worksheet = writer.sheets['每周告警统计汇总']
                    for idx, col in enumerate(combined_df.columns):
                        max_length = max(
                            combined_df[col].fillna('').astype(str).str.len().max(),
                            len(col)
                        ) + 2
                        worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 30)
                
                # 检查输出文件是否成功创建
                if os.path.exists(output_file):
                    file_size = os.path.getsize(output_file)
                    logging.info(f"输出文件已创建，大小: {file_size} 字节")
                else:
                    logging.warning(f"警告: 输出文件 {output_file} 未成功创建")
                
            except Exception as e:
                logging.error(f"写入Excel文件时出错: {str(e)}")
                logging.error("详细错误信息:")
                logging.error(traceback.format_exc())
                
                # 尝试使用备用文件名保存
                backup_file = os.path.join(output_dir, f"告警统计分析_combined_weekly_chinese_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
                logging.info(f"尝试使用备用文件名保存: {backup_file}")
                try:
                    combined_df.to_excel(backup_file, index=False)
                    logging.info(f"成功保存到备用文件: {backup_file}")
                except Exception as e2:
                    logging.error(f"保存到备用文件时出错: {str(e2)}")
                return
        
            logging.info(f"合并完成，结果保存在: {output_file}")
            
            # 生成HTML报告（如果指定）
            report_file = None
            if args.generate_report:
                try:
                    report_file = generate_html_report(combined_df, output_file, args)
                    if report_file and os.path.exists(report_file):
                        logging.info(f"HTML报告已生成: {report_file}")
                except Exception as e:
                    logging.error(f"生成HTML报告时出错: {str(e)}")
                    logging.error(traceback.format_exc())
            
            # 显示合并后的数据统计信息
            logging.info(f"\n合并后的数据包含 {len(combined_df)} 行，涵盖 {combined_df['attr'].nunique()} 个不同的系统")
            logging.info("每个系统的记录数:")
            for system, count in combined_df['attr'].value_counts().sort_index().items():
                logging.info(f"  {system}: {count}")
            
            # 显示告警统计信息
            total_alarms = combined_df['告警数量'].sum()
            total_processed = combined_df['已处理数量'].sum()
            total_on_time = combined_df['未超时处理数量'].sum()
            
            logging.info("\n告警统计信息:")
            logging.info(f"  总告警数量: {total_alarms}")
            logging.info(f"  已处理告警数量: {total_processed}")
            logging.info(f"  未超时处理告警数量: {total_on_time}")
            
            if total_alarms > 0:
                processed_rate = (total_processed / total_alarms) * 100
                on_time_rate = (total_on_time / total_alarms) * 100
                logging.info(f"  处理率: {processed_rate:.2f}%")
                logging.info(f"  未超时率: {on_time_rate:.2f}%")
            
            # 计算总处理时间
            total_time = time.time() - start_time
            logging.info(f"总处理时间: {total_time:.2f}秒")
        
    except Exception as e:
        logging.error(f"处理文件时发生错误: {str(e)}")
        logging.error("详细错误信息:")
        logging.error(traceback.format_exc())

if __name__ == "__main__":
    create_combined_weekly_statistics_with_chinese_names()
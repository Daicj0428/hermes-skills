import os

# 本地模式基础目录（back/ 目录所在位置）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 原始数据文件（请替换为实际路径）
INPUT_FILE = "7月份全量告警明细0801.xlsx"
# 输出目录（本地模式使用 back/statistics）
OUTPUT_DIR = os.path.join(BASE_DIR, "statistics")
# 日志文件
LOG_FILE = os.path.join(BASE_DIR, "logs", "process.log")

# ==================== 字段映射 ====================
# 系统名称映射（attr → 中文名），未匹配的保留原值
ATTR_CHINESE_MAPPING = {
    'crm': 'CRM系统',
    'bossv8': 'BOSS系统',
    'none': '未分类系统',
    'bass': 'BASS系统',
    'dq': '电子渠道',
    'network': '网络设备',
    'apm': 'APM系统',
    'database': '数据库系统',
    'mss': 'M域系统',
    'cloud': '省云平台',
    'js': '结算系统',
    'icloud': '一级云平台',
    'qd': '渠道集中化管理平台',
    'bomc': 'BOMC系统',
    'gh': '家宽智能规划',
    '4a': '4A系统',
    'wg': '网格化系统',
    'nan': '未知系统',
    'dsj': '倒三角系统',
    'sj': '数管平台',
    'other': '其他系统',
    'dt': '地图平台',
    'tsgz': '态势平台',
    'pj': '磐基平台',
    'dd': '订单系统',
    'aqsjzx': '安全数据中心',
    'jh': '稽核平台',
    'hdpt': '混沌平台',
    'casb': 'CASB系统',
    'hgpt': '合规平台',
    'smp': 'SMP平台',
}

# 告警级别映射（英文 → 中文），未匹配的保留原值
LEVEL_CHINESE_MAPPING = {
    'serious': '严重告警', 'critical': '严重告警', 'fatal': '严重告警',
    'important': '重要告警', 'major': '重要告警', 'warning': '警告',
    'general': '一般告警', 'minor': '一般告警', 'info': '提示',
    'notice': '提示', 'normal': '一般告警', 'unknown': '未知',
}

# ==================== 自定义导出字段 ====================
# 自定义导出时可选字段列表，修改此处即可调整前端导出弹窗
EXPORT_FIELDS = [
    {'key': 'attr',           'label': '所属系统'},
    {'key': '告警级别',        'label': '告警级别'},
    {'key': '周期',           'label': '统计周期'},
    {'key': 'system',         'label': '监控目录'},
    {'key': 'RESOURCE_NAME',  'label': '资源对象'},
    {'key': 'alarm_content',  'label': '告警内容'},
    {'key': 'alarm_first_time', 'label': '首次告警时间'},
    {'key': 'alarm_last_time',  'label': '最后告警时间'},
    {'key': 'DISCHARGE_TIME',   'label': '清除告警时间'},
    {'key': 'ALARM_COUNT',    'label': '告警次数'},
    {'key': 'alarm_state',    'label': '告警状态'},
    {'key': 'deal_status',    'label': '处理情况'},
    {'key': 'deal_time',      'label': '处理时长'},
    {'key': 'ack_status',     'label': '确认状态'},
    {'key': 'LAST_ACK_TIME',  'label': '确认时间'},
    {'key': 'ACK_COMMENT',    'label': '确认信息'},
    {'key': 'mainResp',       'label': '主要责任人'},
    {'key': 'secondResp',     'label': '次要责任人'},
]
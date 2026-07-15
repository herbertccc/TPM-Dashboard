#!/usr/bin/env python3
"""项目管理汇报看板生成脚本 - 钉钉版（HTTP API）"""
import os, json, re, time
import urllib.request
import urllib.parse
import ssl
from datetime import datetime, timedelta
from collections import defaultdict

# ===== 配置 =====
DINGTALK_APP_KEY = os.environ.get('DINGTALK_APP_KEY', 'ding9exxqt1c7nhuub7v')
DINGTALK_APP_SECRET = os.environ.get('DINGTALK_APP_SECRET', 'fUp2nEtIWIaCSdq55rwLdE9mY6bzqVb3wPcntsJendE4iaojVGvh8FsxsTSFSIKn')
OPERATOR_ID = 'PRmQRCFpuKchRpqCviSdoiiwiEiE'
API_BASE = 'https://api.dingtalk.com'

DEFECT_NODE = 'P0MALyR8kNnB9yZliY70z99KJ3bzYmDO'
DEFECT_DATA_SHEET = 'kgqie6hm'
DEFECT_PERSON_SHEET = 'st-77d83270-86268'
MILESTONE_NODE = '7dx2rn0Jboav95gYSZv0obP38MGjLRb3'
MILESTONE_SHEET = 'kgqie6hm'
VERSION_PLAN_NODE = 'ndMj49yWj2YdvBzXiRGxxNKeW3pmz5aA'
VERSION_PLAN_SHEET = 'kgqie6hm'


# ===== 钉钉开放平台 HTTP API 函数 =====
_token_cache = {"token": None, "expires": 0}

def _dingtalk_get_token():
    """获取钉钉 accessToken（带缓存）"""
    now = time.time()
    if _token_cache["token"] and _token_cache["expires"] > now + 60:
        return _token_cache["token"]
    url = f"{API_BASE}/v1.0/oauth2/accessToken"
    body = json.dumps({"appKey": DINGTALK_APP_KEY, "appSecret": DINGTALK_APP_SECRET}).encode()
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(req, context=ctx, timeout=15) as resp:
        data = json.loads(resp.read())
    token = data.get("accessToken")
    if not token:
        raise RuntimeError(f"获取 accessToken 失败: {data}")
    _token_cache["token"] = token
    _token_cache["expires"] = now + data.get("expireIn", 7200)
    print("  ✅ 获取 accessToken 成功")
    return token


def _dingtalk_api_get(path, params=None, retries=5):
    """通用 GET 请求，带重试"""
    token = _dingtalk_get_token()
    url = f"{API_BASE}{path}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    headers = {
        "x-acs-dingtalk-access-token": token,
        "Content-Type": "application/json",
    }
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=headers)
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            with urllib.request.urlopen(req, context=ctx, timeout=60) as resp:
                raw = resp.read()
                data = json.loads(raw)
            return data
        except urllib.error.HTTPError as e:
            body = e.read().decode(errors='replace')
            if "ServiceUnavailable" in body or e.code >= 500:
                wait = min(3 * (attempt + 1), 15)
                print(f"  ⏳ 服务端错误 {e.code}，等待{wait}s重试 {attempt+1}/{retries}...")
                time.sleep(wait)
                continue
            print(f"  ❌ HTTP {e.code}: {body[:300]}")
            raise
        except Exception as e:
            if attempt < retries - 1:
                print(f"  ⏳ 请求异常: {e}，重试 {attempt+1}/{retries}...")
                time.sleep(2 ** attempt)
                continue
            raise
    return None


def _dingtalk_read_range(node, sheet_id, range_str):
    """读取指定范围，返回 displayValues 二维数组"""
    path = f"/v1.0/doc/workbooks/{node}/sheets/{sheet_id}/ranges/{range_str}"
    data = _dingtalk_api_get(path, params={"operatorId": OPERATOR_ID})
    if data and "displayValues" in data:
        return data.get("displayValues", [])
    return None


def _dingtalk_read_large_sheet(node, sheet_id, total_rows, col_start='A', col_end='L', chunk_size=1000, expected_headers=None):
    """分块读取大型表格，返回 (headers, data_rows)
    chunk_size=1000: 避免 API 限制
    expected_headers: 可选，用于验证第一行是否为表头
    """
    all_rows = []
    for start in range(1, total_rows + 2, chunk_size):
        end = min(start + chunk_size - 1, total_rows + 1)
        range_str = f"{col_start}{start}:{col_end}{end}"
        print(f"  📖 读取 {range_str}...")
        chunk = _dingtalk_read_range(node, sheet_id, range_str)
        if chunk:
            all_rows.extend(chunk)
            print(f"     ✅ {len(chunk)} 行")
        else:
            print(f"     ⚠️ 跳过 {range_str}")
        time.sleep(1)  # 避免 API 限流
    if not all_rows:
        return [], []

    # 验证第一行是否为表头
    first_row = [str(c).strip() if c else "" for c in all_rows[0]]
    is_header = True
    if expected_headers:
        # 检查第一行是否包含预期的表头关键字
        match_count = sum(1 for eh in expected_headers if eh in first_row)
        is_header = match_count >= len(expected_headers) * 0.5  # 至少匹配50%
    
    if is_header and expected_headers and all(h in first_row for h in expected_headers):
        # 完美匹配，使用实际表头
        headers = first_row
        data_rows = all_rows[1:]
    elif is_header:
        headers = first_row
        data_rows = all_rows[1:]
    else:
        # 第一行不是表头，使用回退表头，所有行都是数据
        print(f"  ⚠️ 第一行不是表头，使用回退列名")
        if expected_headers:
            headers = expected_headers
        else:
            headers = [f"col_{i}" for i in range(len(first_row))]
        data_rows = all_rows
    return headers, data_rows


def _networkdays(start_date, end_date):
    """计算两个日期之间的工作日数（不含周末）"""
    if start_date > end_date:
        return 0
    days = 0
    current = start_date
    while current <= end_date:
        if current.weekday() < 5:  # 周一到周五
            days += 1
        current += timedelta(days=1)
    return days


def _normalize_dept(raw):
    """标准化部门名称"""
    if not raw:
        return ""
    raw = raw.strip()
    mapping = {
        'AIoT部': 'AIOT',
        'AIOT部': 'AIOT',
        'aiot部': 'AIOT',
        '其他': '其他',
    }
    return mapping.get(raw, raw.upper())


def _normalize_role(raw):
    """标准化角色名称"""
    if not raw:
        return ""
    raw = raw.strip()
    mapping = {
        '开发': 'AIOT开发',
        '测试': '软件测试',
        '其他': 'Other',
    }
    return mapping.get(raw, raw)


def _excel_serial_to_date_str(val):
    """将飞书返回的 Excel 序列号日期转为 YYYY-MM-DD 字符串"""
    if val is None:
        return ""
    
    # 处理字符串类型
    if isinstance(val, str):
        val = val.strip()
        if not val:
            return ""
        
        # 检查是否已经是标准日期格式 (含 '-')
        if '-' in val and len(val) >= 10:
            # 提取日期部分（去掉时间）
            date_part = val.split(' ')[0]
            # 验证格式 YYYY-MM-DD
            try:
                datetime.strptime(date_part, '%Y-%m-%d')
                return date_part
            except ValueError:
                pass
        
        # 检查是否为中文日期格式，如 "7月14日" 或 "2026年7月14日"
        import re
        chinese_date_pattern = r'(\d{1,4})年(\d{1,2})月(\d{1,2})日|(\d{1,2})月(\d{1,2})日'
        match = re.search(chinese_date_pattern, val)
        if match:
            if match.group(1):  # 包含年份的情况
                year = int(match.group(1))
                month = int(match.group(2))
                day = int(match.group(3))
            else:  # 不包含年份的情况
                # 使用当前年份作为默认年份
                current_year = datetime.now().year
                month = int(match.group(4))
                day = int(match.group(5))
                year = current_year
            try:
                dt = datetime(year, month, day)
                return dt.strftime('%Y-%m-%d')
            except ValueError:
                pass
        
        # 尝试解析为数字（Excel 序列号）
        try:
            serial = int(float(val))
        except (ValueError, TypeError):
            # 无法转换的字符串，返回空
            return ""
    elif isinstance(val, (int, float)):
        serial = int(val)
    else:
        return str(val)
    
    # Excel 序列号转日期 (1899-12-30 为 day 0)
    try:
        dt = datetime(1899, 12, 30) + timedelta(days=serial)
        return dt.strftime('%Y-%m-%d')
    except Exception:
        return ""





# ===== 读取人员映射表 =====
print("📖 读取人员映射表...")
person_map_data = None
for _retry in range(3):
    person_map_data = _dingtalk_read_range(DEFECT_NODE, DEFECT_PERSON_SHEET, "A1:D200")
    if person_map_data and len(person_map_data) > 1:
        break
    print(f"  ⏳ 人员映射表读取失败，重试 {_retry+1}/3...")
    time.sleep(2)
person_mapping = {}  # {姓名: {"role": ..., "dept": ...}}
if person_map_data and len(person_map_data) > 1:
    for row in person_map_data[1:]:  # 跳过表头
        if not row or not row[0]:
            continue
        name = str(row[0]).strip()
        role = str(row[2]).strip() if len(row) > 2 and row[2] else ""
        dept = str(row[3]).strip() if len(row) > 3 and row[3] else ""
        if name:
            person_mapping[name] = {"role": role, "dept": dept}
    print(f"  ✅ 人员映射: {len(person_mapping)} 人")
else:
    print("  ⚠️ 无法读取人员映射表")


# ===== 读取缺陷列表数据（仅 A-L 列，避免公式列超时）=====
print("📖 读取缺陷列表数据...")
TOTAL_DATA_ROWS = 10000  # 足够大以覆盖所有行
EXPECTED_DEFECT_HEADERS = ["DB-项目", "标题", "任务ID", "执行者", "任务状态", "解决者",
                           "解决时间", "BUG等级", "创建时间", "创建者", "重开次数", "释放管控"]
headers, data_rows = _dingtalk_read_large_sheet(
    DEFECT_NODE, DEFECT_DATA_SHEET, TOTAL_DATA_ROWS,
    col_start='A', col_end='L', chunk_size=1000,
    expected_headers=EXPECTED_DEFECT_HEADERS
)
print(f"  📊 读取完成: {len(data_rows)} 行数据, 列: {headers}")

# 建立列名到索引的映射
col_idx = {h: i for i, h in enumerate(headers)}

# ===== 解析 bug 数据（在 Python 中计算公式列）=====
def parse_bugs_from_rows(data_rows, headers, person_mapping):
    """从原始行数据解析 bug 列表，在 Python 中计算所有公式列"""
    col = {h: i for i, h in enumerate(headers)}
    bugs = []

    for row in data_rows:
        if not row or not any(row):
            continue

        def get(name, default=""):
            idx = col.get(name)
            if idx is not None and idx < len(row) and row[idx] is not None:
                return row[idx]
            return default

        # 基础字段（来自 A-L 列）
        project = str(get("DB-项目", "")).strip()
        title = str(get("标题", "")).strip()
        task_id = str(get("任务ID", "")).strip()
        assignee = str(get("执行者", "")).strip()
        status = str(get("任务状态", "")).strip()
        resolver = str(get("解决者", "")).strip()
        resolved_time_raw = str(get("解决时间", "")).strip()
        bug_level = str(get("BUG等级", "")).strip()
        created_time_raw = str(get("创建时间", "")).strip()
        creator = str(get("创建者", "")).strip()
        reopen_raw = get("重开次数", 0)
        release_ctrl = str(get("释放管控", "")).strip()

        try:
            reopen_count = int(float(reopen_raw)) if reopen_raw else 0
        except (ValueError, TypeError):
            reopen_count = 0

        # === 计算 DB-BUG等级 (列 O) ===
        level_map = {"非常紧急": "P0", "P0": "P0", "紧急": "P1", "P1": "P1",
                     "普通": "P2", "P2": "P2", "较低": "P3", "P3": "P3"}
        db_bug_level = level_map.get(bug_level, "")

        # === 计算 DB-DI值 (列 P) ===
        di_map = {"P0": 10, "P1": 3, "P2": 1, "P3": 0.1}
        weight = di_map.get(db_bug_level, 0)

        # === 计算 DB-角色 和 DB-部门 (列 R, S) ===
        person_info = person_mapping.get(assignee, {})
        raw_role = person_info.get("role", "")
        raw_dept = person_info.get("dept", "")
        db_role = _normalize_role(raw_role)
        db_dept = _normalize_dept(raw_dept)

        # === 计算 DB-任务状态 (列 M) ===
        if status in ("已关闭", "不予解决", "非问题关闭"):
            db_task_status = "已关闭"
        elif status in ("外部问题", "设计如此", "重复问题", "无法重现", "回归验证", "已解决"):
            db_task_status = "待回归"
        elif status in ("挂起", "重复打开", "修复中", "激活", "待处理", "重新打开"):
            db_task_status = "待处理"
        else:
            db_task_status = status

        # === 解析日期 ===
        def parse_datetime_str(s):
            if not s or s in ("None", ""):
                return None
            try:
                # 格式: "2025/6/24 17:28:00.0"
                s = s.replace(".0", "").replace("/", "-")
                date_part = s.split(" ")[0]
                return datetime.strptime(date_part, "%Y-%m-%d")
            except (ValueError, TypeError):
                return None

        created_dt = parse_datetime_str(created_time_raw)
        resolved_dt = parse_datetime_str(resolved_time_raw)

        # === 计算 DB-SLA超时 (列 Q) ===
        sla_timeout = 0
        sla_days = 0
        if status not in ("已关闭", "已解决"):
            if created_dt:
                nd = _networkdays(created_dt, datetime.now())
                thresholds = {"P0": 1, "P1": 3, "P2": 5, "P3": 8}
                threshold = thresholds.get(db_bug_level, 999)
                if nd > threshold:
                    sla_timeout = 1
                    sla_days = nd

        # === 计算 解决超90天 (列 N) ===
        resolve_over_90 = ""
        if created_dt and (datetime.now() - created_dt).days > 90:
            resolve_over_90 = "超90天"

        # === DI 计算 ===
        is_aiot = (db_dept == "AIOT")
        if is_aiot:
            # OPEN DI: DB-任务状态为非"已关闭"时计入
            open_di = weight if db_task_status != "已关闭" else 0
            # 已解决DI: 状态为"已解决"或"已关闭"
            solved_di = weight if status in ("已解决", "已关闭", "已完成") else 0
            # 未解决DI: 状态非"已解决"/"已关闭"/"已完成"
            unsolved_di = weight if status not in ("已解决", "已关闭", "已完成") else 0
        else:
            open_di = solved_di = unsolved_di = 0

        # === 过滤规则：解决超90天 ===
        if resolve_over_90 == "超90天" and status == "已关闭":
            if created_dt and (datetime.now() - created_dt).days > 90:
                continue

        bugs.append({
            "project": project,
            "title": title,
            "id": task_id,
            "assignee": assignee,
            "status": status,
            "resolver": resolver,
            "release_ctrl": release_ctrl,
            "level": bug_level,
            "creator": creator,
            "reopen_count": reopen_count,
            "db_role": db_role,
            "db_dept": db_dept,
            "db_task_status": db_task_status,
            "sla_timeout": sla_timeout,
            "sla_days": sla_days,
            "rawWeight": weight,
            "openDI": open_di,
            "solvedDI": solved_di,
            "unsolvedDI": unsolved_di,
            "_created_dt": created_dt,
            "_resolved_dt": resolved_dt,
        })

    return bugs


all_bugs_raw = parse_bugs_from_rows(data_rows, headers, person_mapping)
print(f"🐛 解析完成: {len(all_bugs_raw)} 条有效 bug")

# 按项目分组
PROJECT_NAMES = sorted(set(b["project"] for b in all_bugs_raw if b["project"]))
all_projects_bugs = {}
for pn in PROJECT_NAMES:
    all_projects_bugs[pn] = [b for b in all_bugs_raw if b["project"] == pn]
    print(f"   {pn}: {len(all_projects_bugs[pn])} 条任务")

print(f"📁 共 {len(PROJECT_NAMES)} 个项目: {PROJECT_NAMES}")

# ===== 统计计算 =====
def calc_stats(bugs):
    """计算项目统计指标"""
    total = len(bugs)
    resolved = sum(1 for b in bugs if b["status"] in ("已关闭", "已解决"))
    open_di = sum(b["openDI"] for b in bugs)
    
    # SLA超时判定 - 以DB-SLA超时列为准，统计非空个数
    sla = sum(b["sla_timeout"] for b in bugs)
    
    # BLOCK问题
    block = sum(1 for b in bugs if b["status"] not in ("已关闭", "已解决") and b["release_ctrl"] == "高风险问题")
    
    # 待回归：状态为"已解决"或"已拒绝"的任务的 DB-DI 值合计
    pending_regression = sum(b["rawWeight"] for b in bugs if b["status"] in ("已解决", "已拒绝"))
    
    resolve_rate = f"{resolved / total * 100:.1f}" if total > 0 else "0.0"
    
    # 健康度判定（四级：健康/一般/较差/严重）
    if open_di < 20 and sla <= 2:
        health = "normal"
    elif open_di < 40 and sla <= 5:
        health = "warning"
    elif open_di < 80 and sla <= 20:
        health = "caution"
    else:
        health = "danger"
    
    return {
        "total": total, "resolved": resolved, "open_di": round(open_di, 1),
        "sla": sla, "block": block, "resolve_rate": resolve_rate, "health": health,
        "pending_regression": round(pending_regression, 1),
    }


projects_stats = {}
for pn in PROJECT_NAMES:
    projects_stats[pn] = calc_stats(all_projects_bugs[pn])


# ===== 趋势计算（总览）=====
def calc_trend(all_bugs):
    """计算所有项目聚合的近30天趋势"""
    now = datetime.now()
    dates = [(now - timedelta(days=i)).strftime("%m-%d") for i in range(29, -1, -1)]
    date_objs = [(now - timedelta(days=i)).date() for i in range(29, -1, -1)]
    date_set = set(dates)
    daily_new = defaultdict(float)
    daily_resolved = defaultdict(float)
    
    # 收集所有AIOT部门的bug
    all_bug_list = []
    for pn, bugs in all_bugs.items():
        for b in bugs:
            if b["db_dept"] != "AIOT":
                continue
            all_bug_list.append(b)
            status = b["status"]
            # 每日解决: DB-任务状态为待回归或已关闭，以解决时间为维度
            is_closed_like = (b["db_task_status"] in ("待回归", "已关闭"))
            if is_closed_like and b["_resolved_dt"]:
                d = b["_resolved_dt"].strftime("%m-%d")
                if d in date_set:
                    daily_resolved[d] += b["rawWeight"]
            # 按创建时间计入每日新增（所有 bug 都算，不限状态）
            if b["_created_dt"]:
                d = b["_created_dt"].strftime("%m-%d")
                if d in date_set:
                    daily_new[d] += b["rawWeight"]
    
    # OPEN DI = 截止到当天，尚未关闭的 AIOT bug 的 DI 值之和
    # 最新天直接使用 bug 的 openDI 字段（与项目统计 calc_stats 保持一致）
    # 历史天用解决时间判断关闭状态
    today_date = now.date()
    cum_list = []
    for d_str, d_date in zip(dates, date_objs):
        open_di = 0.0
        for b in all_bug_list:
            if d_date == today_date:
                # 最新天：直接使用预计算的 openDI（基于 db_task_status，与项目统计一致）
                open_di += b["openDI"]
            else:
                # 历史天：判断该 bug 在日期 D 是否已关闭
                if b["status"] in ("已关闭", "不予解决", "非问题关闭") and b["_resolved_dt"] and b["_resolved_dt"].date() <= d_date:
                    continue
                if b["_created_dt"] and b["_created_dt"].date() <= d_date:
                    open_di += b["rawWeight"]
        cum_list.append(round(open_di, 1))
    
    return dates, cum_list, [round(daily_new.get(d, 0), 1) for d in dates], [round(daily_resolved.get(d, 0), 1) for d in dates]


trend_dates_str, cumulative_di, daily_new_di, daily_resolved_di = calc_trend(all_projects_bugs)


# ===== 单项目趋势计算 =====
def calc_single_project_trend(bugs):
    """计算单个项目的近30天趋势"""
    now = datetime.now()
    dates = [(now - timedelta(days=i)).strftime("%m-%d") for i in range(29, -1, -1)]
    date_objs = [(now - timedelta(days=i)).date() for i in range(29, -1, -1)]
    date_set = set(dates)
    daily_new = defaultdict(float)
    daily_resolved = defaultdict(float)
    today_date = now.date()
    
    for b in bugs:
        if b["db_dept"] != "AIOT":
            continue
        status = b["status"]
        # 每日解决: DB-任务状态为待回归或已关闭，以解决时间为维度
        is_closed_like = (b["db_task_status"] in ("待回归", "已关闭"))
        if is_closed_like and b["_resolved_dt"]:
            d = b["_resolved_dt"].strftime("%m-%d")
            if d in date_set:
                daily_resolved[d] += b["rawWeight"]
        # 按创建时间计入每日新增（所有 bug 都算，不限状态）
        if b["_created_dt"]:
            d = b["_created_dt"].strftime("%m-%d")
            if d in date_set:
                daily_new[d] += b["rawWeight"]
    
    cum_list = []
    for d_str, d_date in zip(dates, date_objs):
        open_di = 0.0
        for b in bugs:
            if b["db_dept"] != "AIOT":
                continue
            if d_date == today_date:
                # 最新天：直接使用预计算的 openDI（与项目统计一致）
                open_di += b["openDI"]
            else:
                # 历史天：判断该 bug 在日期 D 是否已关闭
                if b["status"] in ("已关闭", "不予解决", "非问题关闭") and b["_resolved_dt"] and b["_resolved_dt"].date() <= d_date:
                    continue
                if b["_created_dt"] and b["_created_dt"].date() <= d_date:
                    open_di += b["rawWeight"]
        cum_list.append(round(open_di, 1))
    
    return cum_list, [round(daily_new.get(d, 0), 1) for d in dates], [round(daily_resolved.get(d, 0), 1) for d in dates]


project_trends = {}
for pn in PROJECT_NAMES:
    project_trends[pn] = calc_single_project_trend(all_projects_bugs[pn])


# ===== 人员统计 =====
def calc_person_stats(PROJECT_NAMES, all_projects_bugs):
    """计算人员效能统计"""
    person_map = {}
    for pn in PROJECT_NAMES:
        for b in all_projects_bugs[pn]:
            name = b["assignee"]
            if not name:
                continue
            if name not in person_map:
                person_map[name] = {
                    "name": name, "role": "未知",
                    "proj_contrib": {},
                    "total_open_di": 0, "total_solved_di": 0, "total_unsolved_di": 0,
                    "total_sla": 0, "total_reopen": 0,
                    "solved_tickets": 0, "total_tickets": 0,
                    "sla_days_sum": 0, "sla_days_count": 0,
                }
            contrib = person_map[name]["proj_contrib"]
            if pn not in contrib:
                contrib[pn] = {"solved_di": 0, "open_di": 0, "unsolved_di": 0}
            
            contrib[pn]["solved_di"] += b["solvedDI"]
            person_map[name]["total_solved_di"] += b["solvedDI"]
            contrib[pn]["unsolved_di"] += b["unsolvedDI"]
            person_map[name]["total_unsolved_di"] += b["unsolvedDI"]
            contrib[pn]["open_di"] += b["openDI"]
            person_map[name]["total_open_di"] += b["openDI"]
            
            # 工单计数（按状态，不按DI值）
            if b["status"] in ("已解决", "已关闭"):
                person_map[name]["solved_tickets"] += 1
            person_map[name]["total_tickets"] += 1
            
            # SLA超时 - 以DB-SLA超时列为准
            person_map[name]["total_sla"] += b["sla_timeout"]
            if b["sla_days"] > 0:
                person_map[name]["sla_days_sum"] += b["sla_days"]
                person_map[name]["sla_days_count"] += 1
            person_map[name]["total_reopen"] += b["reopen_count"]
    
    return list(person_map.values())


person_project_stats = calc_person_stats(PROJECT_NAMES, all_projects_bugs)





# ===== 读取里程碑数据（从钉钉在线表格，项目列有合并单元格）=====
print("📅 读取里程碑数据...")
milestones = {}
ms_vals = _dingtalk_read_range(MILESTONE_NODE, MILESTONE_SHEET, "A1:F200")
if ms_vals and len(ms_vals) > 1:
    last_project = ""
    for row in ms_vals[1:]:
        if not row or not any(row):
            continue
        project_name = str(row[0]).strip() if len(row) > 0 and row[0] else ""
        if project_name:
            last_project = project_name
        elif last_project:
            project_name = last_project
        if not project_name:
            continue
        date_val = _excel_serial_to_date_str(row[1]) if len(row) > 1 and row[1] else ""
        ms_entry = {
            "date": date_val,
            "name": str(row[2]).strip() if len(row) > 2 and row[2] else "",
            "item": str(row[3]).strip() if len(row) > 3 and row[3] else "",
            "status": str(row[4]).strip() if len(row) > 4 and row[4] else "",
            "risk": str(row[5]).strip() if len(row) > 5 and row[5] else "",
        }
        if project_name not in milestones:
            milestones[project_name] = []
        milestones[project_name].append(ms_entry)
    print(f"📅 里程碑数据: {len(milestones)} 个项目")
else:
    print("⚠️ 无法读取里程碑数据")


# ===== 读取版本计划数据 - 从钉钉表格读取 =====
print("📋 读取版本计划数据...")
version_plan = []
try:
    vp_vals = _dingtalk_read_range(VERSION_PLAN_NODE, VERSION_PLAN_SHEET, "A1:E20")
    if vp_vals and len(vp_vals) > 1:
        vp_headers = [str(h).strip() if h else '' for h in vp_vals[0]]
        print(f'📋 版本计划表头: {vp_headers}')

        date_idx = None
        version_idx = None
        content_idx = None
        link_idx = None

        for i, h in enumerate(vp_headers):
            h_lower = h.lower()
            if '日期' in h or '时间' in h or 'date' in h_lower:
                date_idx = i
            elif '版本' in h or 'version' in h_lower or '标题' in h:
                version_idx = i
            elif '内容' in h or '概述' in h or 'overview' in h_lower or '说明' in h:
                content_idx = i
            elif '链接' in h or 'link' in h_lower or '文档' in h:
                link_idx = i

        print(f'   列索引: 日期={date_idx}, 版本={version_idx}, 内容={content_idx}, 链接={link_idx}')

        plans = []
        for row in vp_vals[1:]:
            if not row or not any(row):
                continue

            date_val = ""
            if date_idx is not None and len(row) > date_idx and row[date_idx]:
                date_val = _excel_serial_to_date_str(row[date_idx])

            version_name = ""
            if version_idx is not None and len(row) > version_idx and row[version_idx]:
                version_name = str(row[version_idx]).strip()

            overview = ""
            if content_idx is not None and len(row) > content_idx and row[content_idx]:
                overview = str(row[content_idx]).strip()

            doc_link = ""
            if link_idx is not None and len(row) > link_idx and row[link_idx]:
                link_val = row[link_idx]
                if isinstance(link_val, list) and len(link_val) > 0:
                    first = link_val[0]
                    if isinstance(first, dict):
                        doc_link = first.get('link', '') or first.get('text', '')
                    else:
                        doc_link = str(first).strip()
                elif isinstance(link_val, dict):
                    doc_link = link_val.get('link', '') or link_val.get('text', '')
                else:
                    doc_link = str(link_val).strip()

            if version_name:
                plans.append({
                    "date": date_val,
                    "name": version_name,
                    "overview": overview,
                    "link": doc_link,
                })

        def parse_date_for_sort(p):
            if p["date"]:
                try:
                    return datetime.strptime(p["date"], "%Y-%m-%d")
                except:
                    pass
            return datetime.min

        plans.sort(key=parse_date_for_sort, reverse=True)
        version_plan = plans
        print(f'📋 版本计划: {len(version_plan)} 条')
    else:
        print("⚠️ 无法读取版本计划数据")
except Exception as e:
    print(f'⚠️ 读取版本计划失败: {e}')
    version_plan = []

css = """
:root { --bg:#f5f7fa; --sidebar-bg:#fff; --card-bg:#fff; --text:#1d2129; --text-secondary:#86909c; --border:#e5e6eb; --blue:#165dff; --red:#f53f3f; --green:#00b42a; --orange:#ff7d00; --yellow:#ffb400; --gray:#c9cdd4; --shadow:0 2px 8px rgba(0,0,0,0.06); --radius:10px; }
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC",sans-serif; background:var(--bg); display:flex; min-height:100vh; color:var(--text); }
.sidebar { width:200px; background:var(--sidebar-bg); position:fixed; top:0; left:0; bottom:0; z-index:100; border-right:1px solid var(--border); display:flex; flex-direction:column; }
.sidebar-header { padding:20px 16px; border-bottom:1px solid var(--border); display:flex; align-items:center; gap:10px; }
.sidebar-logo { width:28px; height:28px; background:linear-gradient(135deg,#165dff,#722ed1); border-radius:6px; display:flex; align-items:center; justify-content:center; color:#fff; font-size:11px; font-weight:800; flex-shrink:0; }
.sidebar-title { font-size:15px; font-weight:700; }
.nav-menu { list-style:none; padding:12px 0; flex:1; overflow-y:auto; }
.nav-item { padding:10px 16px; cursor:pointer; transition:all 0.2s; display:flex; align-items:center; gap:8px; font-size:13px; color:var(--text-secondary); border-left:3px solid transparent; margin:2px 0; }
.nav-item:hover { background:#f2f3f5; color:var(--text); }
.nav-item.active { background:rgba(22,93,255,0.06); color:var(--blue); border-left-color:var(--blue); font-weight:600; }
.nav-sub-item { padding:8px 16px 8px 36px; cursor:pointer; transition:all 0.2s; font-size:13px; color:var(--text-secondary); display:flex; align-items:center; border-left:3px solid transparent; margin:1px 0; }
.nav-sub-item:hover { background:#f2f3f5; color:var(--text); }
.nav-sub-item.active { background:rgba(22,93,255,0.06); color:var(--blue); border-left-color:var(--blue); font-weight:600; }
.di-tag { font-size:11px; padding:2px 8px; border-radius:10px; font-weight:600; margin-left:auto; min-width:40px; text-align:center; }
.di-tag.danger { background:#fff2f0; color:var(--red); }
.main { margin-left:200px; flex:1; padding:28px 32px; min-width:0; }
.page-header { margin-bottom:24px; }
.page-title { font-size:22px; font-weight:700; margin-bottom:4px; }
.page-desc { font-size:13px; color:var(--text-secondary); }
.cards-grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(260px,1fr)); gap:14px; margin-bottom:24px; }
.stat-card { background:var(--card-bg); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); cursor:pointer; border:2px solid transparent; transition:all 0.2s; position:relative; }
.stat-card:hover { transform:translateY(-2px); box-shadow:0 6px 20px rgba(0,0,0,0.1); }
.stat-card.selected { border-color:var(--blue); }
.health-badge { position:absolute; top:12px; right:12px; font-size:11px; padding:3px 10px; border-radius:12px; font-weight:600; }
.health-badge.normal { background:#e8ffea; color:var(--green); }
.health-badge.warning { background:#fff7e6; color:var(--orange); }
.health-badge.danger { background:#fff2f0; color:var(--red); }
.card-label { font-size:13px; color:var(--text-secondary); margin-bottom:6px; padding-right:60px; }
.kpi-value { font-size:15px; font-weight:700; min-width:0; overflow:hidden; text-overflow:ellipsis; }
.chart-container { background:var(--card-bg); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); margin-bottom:24px; }
.chart-title { font-size:15px; font-weight:600; margin-bottom:14px; }
.chart-area { width:100%; height:320px; }
.timeline-section { background:var(--card-bg); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); margin-bottom:24px; position:relative; }
.tl-reset-btn { padding:5px 14px; font-size:12px; font-weight:600; color:var(--blue); background:#fff; border:1px solid var(--blue); border-radius:6px; cursor:pointer; transition:all 0.15s; }
.tl-reset-btn:hover { background:var(--blue); color:#fff; }
.tl-timeline-wrap { display:flex; align-items:stretch; gap:0; margin-top:8px; }
.tl-nav-btn { flex-shrink:0; width:32px; display:flex; align-items:center; justify-content:center; background:#f7f8fa; border:1px solid var(--border); cursor:pointer; color:var(--text-secondary); font-size:18px; transition:all 0.15s; user-select:none; }
.tl-nav-btn:hover { background:#eef1f6; color:var(--text); }
.tl-nav-btn.left { border-radius:8px 0 0 8px; border-right:none; }
.tl-nav-btn.right { border-radius:0 8px 8px 0; border-left:none; }
.tl-viewport { flex:1; display:flex; overflow:hidden; position:relative; background:#fafbfc; min-width:0; }
.tl-names-col { width:110px; flex-shrink:0; overflow-y:auto; overflow-x:hidden; border-right:1px solid var(--border); background:#fff; z-index:3; }
.tl-scroll-area { flex:1; overflow-x:hidden; overflow-y:hidden; cursor:grab; }
.tl-timeline-body { position:relative; padding-top:60px; padding-bottom:22px; }
.tl-day-cols { display:flex; }
.tl-day-col { flex-shrink:0; border-right:1px solid #f0f1f3; }
.tl-day-col.today-col { background:rgba(255,125,0,0.04); }
.tl-day-label { height:24px; display:flex; align-items:center; justify-content:center; font-size:11px; color:var(--text-secondary); border-bottom:2px solid var(--border); position:relative; }
.tl-day-col.today-col .tl-day-label { color:var(--orange); font-weight:700; }
.tl-day-body { position:relative; }
.tl-today-line { position:absolute; top:0; bottom:0; width:2px; background:rgba(255,152,0,0.35); z-index:2; pointer-events:none; }
.tl-today-label { position:absolute; bottom:-18px; left:50%; transform:translateX(-50%); font-size:11px; color:var(--orange); font-weight:700; white-space:nowrap; }
.tl-today-top-label { position:absolute; top:4px; left:50%; transform:translateX(-50%); font-size:11px; color:var(--orange); font-weight:700; white-space:nowrap; background:rgba(255,255,255,0.85); padding:1px 6px; border-radius:3px; }
.tl-project-row { margin-bottom:24px; }
.tl-project-row:last-child { margin-bottom:0; }
.tl-names-col .tl-project-name { font-size:13px; font-weight:700; color:var(--text); padding:0 12px; position:absolute; top:6px; left:0; right:0; }
.tl-names-col .tl-name-spacer { height:86px; }
.tl-names-col .tl-name-track { height:14px; position:relative; }
.tl-names-col .tl-name-card-area { }
.tl-track { position:relative; height:14px; }
.tl-track::before { content:''; position:absolute; top:6px; left:0; right:0; height:2px; background:linear-gradient(90deg,var(--border) 0%,#d5d8de 50%,var(--border) 100%); border-radius:1px; }
.tl-dot { position:absolute; top:0; width:14px; height:14px; border-radius:50%; background:var(--blue); border:2.5px solid #fff; box-shadow:0 1px 4px rgba(22,93,255,0.3); transform:translateX(-50%); z-index:2; transition:transform 0.15s; }
.tl-dot:hover { transform:translateX(-50%) scale(1.2); }
.tl-dot.pass { background:var(--green); box-shadow:0 1px 4px rgba(0,180,42,0.3); }
.tl-dot.delay { background:var(--red); box-shadow:0 1px 4px rgba(245,63,63,0.3); }
.tl-dot.progress { background:var(--orange); box-shadow:0 1px 4px rgba(255,125,0,0.3); }
.tl-card-area { position:relative; }
.tl-card { position:absolute; transform:translateX(-50%); background:#fff; border:1px solid var(--border); border-radius:10px; padding:12px 16px; min-width:150px; max-width:210px; z-index:2; box-shadow:0 2px 8px rgba(0,0,0,0.06); transition:box-shadow 0.2s,transform 0.15s; }
.tl-card:hover { box-shadow:0 4px 16px rgba(0,0,0,0.12); transform:translateX(-50%) translateY(-2px); }
.tl-card-name { font-size:13px; font-weight:700; color:var(--text); margin-bottom:4px; }
.tl-card-item { font-size:11px; color:var(--text-secondary); margin-bottom:4px; line-height:1.5; }
.tl-card-risk { font-size:11px; color:var(--red); line-height:1.5; margin-bottom:6px; }
.tl-card-status { display:inline-block; padding:2px 10px; border-radius:4px; font-size:11px; font-weight:700; }
.tl-card-status.pass { background:#e8ffea; color:#00b42a; }
.tl-card-status.delay { background:#fff2f0; color:#f53f3f; }
.tl-card-status.progress { background:#fff7e6; color:#ff7d00; }
.tl-card-status.default { background:#f2f3f5; color:var(--text-secondary); }
.tl-card-date { font-size:10px; color:var(--text-secondary); margin-bottom:4px; }
.tl-connector { position:absolute; width:1px; background:linear-gradient(180deg,var(--border) 0%,rgba(229,230,235,0.3) 100%); z-index:1; }
.role-section { margin-bottom:24px; }
.role-title { font-size:15px; font-weight:700; margin-bottom:12px; padding-left:12px; border-left:3px solid var(--blue); }
.hr-cards { display:grid; grid-template-columns:repeat(auto-fill,minmax(200px,1fr)); gap:10px; }
.hr-card { background:var(--card-bg); border-radius:var(--radius); padding:14px; box-shadow:var(--shadow); }
.hr-name { font-size:13px; font-weight:600; margin-bottom:8px; }
.hr-stats { display:grid; grid-template-columns:1fr 1fr; gap:4px; }
.hr-stat { font-size:11px; color:var(--text-secondary); }
.hr-stat span { color:var(--text); font-weight:600; display:block; font-size:13px; }
.pp-chart { background:var(--card-bg); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); margin-bottom:24px; overflow-x:auto; }
.pp-row { display:flex; align-items:center; padding:8px 0; border-bottom:1px solid var(--border); min-width:600px; }
.pp-row:last-child { border-bottom:none; }
.pp-name { width:100px; font-size:12px; font-weight:600; flex-shrink:0; }
.pp-bars { flex:1; display:flex; flex-direction:column; gap:4px; }
.pp-bar-line { display:flex; align-items:center; height:16px; width:100%; }
.pp-bar-segment { height:100%; transition:width 0.3s; }
.pp-bar-solved { background:var(--green); }
.pp-bar-open { background:var(--red); }
.pp-reopen-bar { height:8px; background:var(--orange); border-radius:4px; }
.pp-legend { display:flex; gap:16px; margin-bottom:12px; font-size:12px; }
.pp-legend span { display:flex; align-items:center; gap:4px; }
.pp-legend i { width:12px; height:12px; border-radius:2px; display:inline-block; }
.pp-vals { font-size:11px; display:flex; gap:8px; flex-shrink:0; margin-left:8px; white-space:nowrap; }
.pp-bar-value { position:absolute; left:4px; top:50%; transform:translateY(-50%); font-size:10px; color:#fff; font-weight:600; white-space:nowrap; line-height:1; text-shadow:0 0 3px rgba(0,0,0,0.3); }
.pp-bar-sla { margin-left:8px; font-size:11px; color:var(--text-secondary); white-space:nowrap; flex-shrink:0; }
.back-btn { display:inline-flex; align-items:center; gap:6px; padding:7px 14px; background:var(--blue); color:#fff; border:none; border-radius:6px; cursor:pointer; font-size:13px; margin-bottom:18px; }
.back-btn:hover { opacity:0.9; }
.person-cards { display:grid; grid-template-columns:repeat(auto-fill,minmax(260px,1fr)); gap:14px; margin-bottom:24px; }
.person-card { background:var(--card-bg); border-radius:var(--radius); padding:14px 16px; box-shadow:var(--shadow); overflow:hidden; min-width:0; }
.person-name { font-size:15px; font-weight:700; margin-bottom:8px; color:var(--text); }
.person-stats { display:flex; gap:14px; flex-wrap:wrap; font-size:12px; }
.person-stat { color:var(--text-secondary); white-space:nowrap; }
.person-stat b { color:var(--text); font-weight:600; }
.hr-table-wrap { background:#fff; border-radius:12px; overflow:hidden; box-shadow:var(--shadow); margin-bottom:24px; }
.hr-table-header { padding:16px 20px; display:flex; align-items:center; gap:10px; border-bottom:1px solid var(--border); }
.hr-table-header-dot { width:8px; height:8px; border-radius:50%; background:var(--blue); }
.hr-table-header-title { font-size:16px; font-weight:700; color:var(--text); }
.hr-table-header-count { font-size:12px; color:var(--text-secondary); background:#f2f3f5; padding:2px 10px; border-radius:10px; }
.hr-table { width:100%; border-collapse:collapse; }
.hr-table thead th { padding:12px 16px; font-size:12px; font-weight:600; color:var(--text-secondary); text-align:center; background:#fafbfc; border-bottom:2px solid var(--border); text-transform:uppercase; letter-spacing:0.5px; }
.hr-table thead th:first-child { text-align:left; padding-left:20px; }
.hr-table tbody tr { transition:background 0.15s; }
.hr-table tbody tr:nth-child(even) { background:#fafbfc; }
.hr-table tbody tr:hover { background:#f0f4ff; }
.hr-table tbody td { padding:12px 16px; font-size:13px; text-align:center; border-bottom:1px solid #f0f1f3; }
.hr-table tbody td:first-child { text-align:left; padding-left:20px; }
.hr-table tbody tr:last-child td { border-bottom:none; }
.hr-table tbody tr.hr-dept-divider td { padding:0; border-bottom:none; background:#fff; }
.hr-table tbody tr.hr-dept-divider:hover { background:#fff; }
.hr-pill { display:inline-block; padding:2px 10px; border-radius:10px; font-size:12px; font-weight:600; }
.hr-metric { display:flex; flex-direction:column; align-items:center; gap:2px; }
.hr-metric-val { font-weight:700; font-size:14px; }
.hr-metric-sub { font-size:11px; color:var(--text-secondary); }
.ranking-section { display:grid; grid-template-columns:1fr 1fr; gap:18px; margin-bottom:24px; }
.ranking-col { background:var(--card-bg); border-radius:var(--radius); padding:18px; box-shadow:var(--shadow); }
.ranking-col-title { font-size:15px; font-weight:600; margin-bottom:14px; display:flex; align-items:center; gap:8px; }
.ranking-col-title .ranking-icon { width:24px; height:24px; border-radius:6px; display:flex; align-items:center; justify-content:center; font-size:13px; font-weight:800; color:#fff; }
.ranking-item { display:flex; align-items:center; padding:12px 14px; border-radius:8px; margin-bottom:8px; background:#fafbfc; transition:all 0.15s; }
.ranking-item:hover { background:#f0f4ff; transform:translateX(2px); }
.ranking-item:last-child { margin-bottom:0; }
.ranking-num { width:26px; height:26px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:800; color:#fff; flex-shrink:0; margin-right:12px; }
.ranking-num.r1 { background:linear-gradient(135deg,#f53f3f,#ff7d00); }
.ranking-num.r2 { background:linear-gradient(135deg,#ff7d00,#ffb400); }
.ranking-num.r3 { background:linear-gradient(135deg,#165dff,#722ed1); }
.ranking-info { flex:1; min-width:0; }
.ranking-name { font-size:13px; font-weight:700; color:var(--text); margin-bottom:4px; }
.ranking-metrics { display:flex; gap:10px; flex-wrap:wrap; font-size:11px; color:var(--text-secondary); }
.ranking-metrics span { white-space:nowrap; }
.ranking-metrics b { color:var(--text); font-weight:600; }
.ranking-badge { flex-shrink:0; margin-left:10px; padding:3px 10px; border-radius:10px; font-size:12px; font-weight:700; }
@media (max-width:900px) { .ranking-section { grid-template-columns:1fr; } }
"""

js = r"""
console.log('[DEBUG] JS开始执行');

// === 手机端检测：提示使用PC端打开 ===
(function() {
  var isMobile = /Android|webOS|iPhone|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) || window.innerWidth < 768;
  if (isMobile) {
    var overlay = document.createElement('div');
    overlay.id = 'mobileBlockOverlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:#fff;z-index:99999;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:40px;';
    overlay.innerHTML = '<div style="font-size:48px;margin-bottom:24px">&#128187;</div>' +
      '<div style="font-size:20px;font-weight:700;color:#1d2129;margin-bottom:12px">请使用PC端打开</div>' +
      '<div style="font-size:14px;color:#86909c;line-height:1.8;max-width:280px">本看板包含复杂图表和时间线视图，手机端无法正常展示。<br>请在电脑浏览器中打开此链接。</div>';
    document.addEventListener('DOMContentLoaded', function() { document.body.appendChild(overlay); });
  }
})();
const PROJECTS_DATA = __PROJECTS_JSON__;
const TREND_DATES = __TREND_DATES_JSON__;
const CUMULATIVE_DI = __CUMULATIVE_DI_JSON__;
const DAILY_NEW = __DAILY_NEW_JSON__;
const DAILY_RESOLVED = __DAILY_RESOLVED_JSON__;
const MILESTONES = __MILESTONES_JSON__;
const PERSON_STATS = __PERSON_STATS_JSON__;
const PROJECT_NAMES = __PROJECT_NAMES_JSON__;
const PROJECT_TRENDS = __PROJECT_TRENDS_JSON__;
const VERSION_PLAN = __VERSION_PLAN_JSON__;
console.log('[DEBUG] 常量定义完成, 项目数:', Object.keys(PROJECTS_DATA).length);

let currentPage = 'overview';
let selectedProject = null;

function renderNav() {
  const menu = document.getElementById('navMenu');
  let html = '<li class="nav-item active" data-page="overview" onclick="switchPage(\'overview\')"> 项目总览</li>';
  var sortedProjects = PROJECT_NAMES.slice().sort(function(a, b) {
    return PROJECTS_DATA[b].stats.open_di - PROJECTS_DATA[a].stats.open_di;
  });
  sortedProjects.forEach(function(p) {
    var s = PROJECTS_DATA[p].stats;
    var tagClass = s.health === 'normal' ? '' : s.health === 'warning' ? ' danger' : ' danger';
    var tagColor = s.health === 'normal' ? '#00b42a' : s.health === 'warning' ? '#ff7d00' : '#f53f3f';
    html += '<li class="nav-sub-item" data-project="' + p + '" onclick="switchPage(\'project\',\'' + p + '\')">' + p + ' <span class="di-tag' + tagClass + '" style="background:' + (s.health==='normal'?'#e8ffea':s.health==='warning'?'#fff7e6':'#fff2f0') + ';color:' + tagColor + '">' + s.open_di.toFixed(1) + '</span></li>';
  });
  html += '<li class="nav-item" data-page="hr" onclick="switchPage(\'hr\')"> 人力资源</li>';
  html += '<li class="nav-item" data-page="timeline" onclick="switchPage(\'timeline\')"> 项目大盘</li>';
  html += '<li class="nav-item" data-page="versionplan" onclick="switchPage(\'versionplan\')"> 版本计划</li>';
  html += '<li class="nav-item" data-page="guide" onclick="switchPage(\'guide\')"> 看板说明</li>';
  menu.innerHTML = html;
}

function switchPage(page, project) {
  currentPage = page; selectedProject = project || null;
  document.querySelectorAll('.nav-item,.nav-sub-item').forEach(function(el){ el.classList.remove('active'); });
  if (page === 'overview') document.querySelector('[data-page="overview"]').classList.add('active');
  else if (page === 'hr') document.querySelector('[data-page="hr"]').classList.add('active');
  else if (page === 'timeline') document.querySelector('[data-page="timeline"]').classList.add('active');
  else if (page === 'versionplan') document.querySelector('[data-page="versionplan"]').classList.add('active');
  else if (page === 'guide') document.querySelector('[data-page="guide"]').classList.add('active');
  else if (project) { var el = document.querySelector('[data-project="' + project + '"]'); if (el) el.classList.add('active'); }
  renderContent();
}

function renderContent() {
  var main = document.getElementById('mainContent');
  if (currentPage === 'overview') main.innerHTML = renderOverview();
  else if (currentPage === 'hr') main.innerHTML = renderHR();
  else if (currentPage === 'timeline') main.innerHTML = renderTimelinePage();
  else if (currentPage === 'versionplan') main.innerHTML = renderVersionPlan();
  else if (currentPage === 'guide') main.innerHTML = renderGuide();
  else if (selectedProject) main.innerHTML = renderProjectDetail(selectedProject);
  setTimeout(initCharts, 50);
}

function getHealthLabel(h) {
  if (h === 'normal') return '<span class="health-badge normal">健康</span>';
  if (h === 'warning') return '<span class="health-badge warning">一般</span>';
  if (h === 'caution') return '<span class="health-badge" style="background:#fff2f0;color:#f53f3f">较差</span>';
  return '<span class="health-badge" style="background:#f53f3f;color:#fff">严重</span>';
}

function buildRiskHtml(persons, s) {
  // 风险识别仅分析AIOT部门人员，侧重SLA超时、BLOCK、OPEN DI
  var aiotPersons = persons.filter(function(p) { return p.dept === 'AIOT'; });
  var risks = [];
  if (s.block > 0) risks.push('当前存在 <b>' + s.block + '</b> 个BLOCK问题（高风险），直接影响版本发布节奏，建议优先排查并指定责任人限时关闭。');
  if (s.sla > 0) risks.push('SLA超时 <b>' + s.sla + '</b> 个，高优先级问题处理速度不达标，可能导致交付延期或客户投诉升级。');
  if (s.open_di > 50) risks.push('项目OPEN DI达 <b>' + s.open_di.toFixed(1) + '</b>，超出警戒线（50），未解决质量风险偏高，需加大投入推进关闭。');
  var highPressure = aiotPersons.filter(function(p) { return p.total >= 8 && p.openDI > 5; });
  if (highPressure.length > 0) {
    var names = highPressure.map(function(p) { return '<span style="color:var(--blue);font-weight:600">' + p.name + '</span>'; }).join('、');
    risks.push('AIOT部门人员压力集中：' + names + ' 任务量大且未解决DI较高，存在过载风险，建议适当分流或增援。');
  }
  var highSla = aiotPersons.filter(function(p) { return p.sla >= 3; });
  if (highSla.length > 0) {
    var names = highSla.map(function(p) { return '<span style="color:var(--blue);font-weight:600">' + p.name + '</span>'; }).join('、');
    risks.push('AIOT部门SLA超时突出：' + names + ' 超时数量偏高，需关注处理时效和排期合理性。');
  }
  if (risks.length === 0) risks.push('当前AIOT部门各项指标正常，暂无明显风险。建议持续关注BLOCK和SLA变化。');
  var html = '<div style="background:var(--card-bg);border-radius:var(--radius);padding:18px;box-shadow:var(--shadow);margin-bottom:24px"><div style="font-size:15px;font-weight:600;margin-bottom:4px">风险识别</div><div style="font-size:12px;color:var(--text-secondary);margin-bottom:12px">基于AIOT部门数据的智能风险分析</div>';
  risks.forEach(function(r, idx) {
    html += '<div style="padding:8px 0;font-size:13px;line-height:1.7">' +
      '<span style="display:inline-block;width:22px;height:22px;line-height:22px;text-align:center;background:#fff7e6;color:#ff7d00;border-radius:50%;font-size:12px;font-weight:700;margin-right:8px;vertical-align:middle">' + (idx+1) + '</span>' +
      r + '</div>';
  });
  html += '</div>';
  return html;
}

function renderGuide() {
  return '<div class="page-header"><div class="page-title">看板说明</div><div class="page-desc">页面功能、指标含义与计算方式说明</div></div>' +
    '<div class="hr-table-wrap">' +
      '<div class="hr-table-header"><div class="hr-table-header-dot"></div><div class="hr-table-header-title">一、看板结构</div></div>' +
      '<table class="hr-table"><thead><tr><th style="width:140px">页面</th><th>说明</th></tr></thead><tbody>' +
      '<tr><td>项目总览</td><td>所有项目的健康度汇总，包含OPEN DI趋势图、风险识别和项目列表表单。点击子项目菜单可进入项目详情页。</td></tr>' +
      '<tr><td>子项目详情</td><td>单个项目的OPEN DI趋势、风险分析和人力视图。人力视图按DB-部门分组、按DB-角色排序展示。</td></tr>' +
      '<tr><td>人力资源</td><td>全员效能总览，按DB-部门分组、DB-角色排序，含人员效能表、DI效能分布柱状图和重开次数统计。</td></tr>' +
      '<tr><td>项目大盘</td><td>各项目里程碑时间线视图，展示里程碑节点日期、状态和风险项，支持滚动和"回到当前"定位。</td></tr>' +
      '<tr><td>版本计划</td><td>从Version_Plan.xlsx读取各项目的版本开发计划，以卡片形式展示计划名称、概述和文档链接。</td></tr>' +
      '</tbody></table></div>' +
    '<div class="hr-table-wrap">' +
      '<div class="hr-table-header"><div class="hr-table-header-dot"></div><div class="hr-table-header-title">二、核心指标说明</div></div>' +
      '<table class="hr-table"><thead><tr><th style="width:140px">指标</th><th>计算方式</th></tr></thead><tbody>' +
      '<tr><td>OPEN DI</td><td>仅统计DB-部门=AIOT且任务状态≠已关闭的BUG，对其DB-DI值求和。DI值从Excel的DB-DI值列直接读取。</td></tr>' +
      '<tr><td>已解决DI</td><td>DB-部门=AIOT且任务状态=已解决的BUG的DB-DI值之和。</td></tr>' +
      '<tr><td>未解决DI</td><td>DB-部门=AIOT且任务状态≠已解决且≠已关闭的BUG的DB-DI值之和。</td></tr>' +
      '<tr><td>解决率</td><td>（已解决+已关闭任务数）/ 总任务数 × 100%。注意：满足"解决超90天=超90天"且"已关闭"且"创建时间距今>90天"的任务已被排除。</td></tr>' +
      '<tr><td>SLA超时题数</td><td>DB-SLA超时列非空的BUG数量。</td></tr>' +
      '<tr><td>BLOCK问题数</td><td>任务状态≠已关闭/已解决 且 释放管控=高风险问题的BUG数。</td></tr>' +
      '<tr><td>BUG数量</td><td>已解决数/全部任务数，不含已被过滤的过期已关闭任务。</td></tr>' +
      '</tbody></table></div>' +
    '<div class="hr-table-wrap">' +
      '<div class="hr-table-header"><div class="hr-table-header-dot"></div><div class="hr-table-header-title">三、健康度判定</div></div>' +
      '<table class="hr-table"><thead><tr><th style="width:100px">等级</th><th style="width:80px">标签</th><th>判定条件</th></tr></thead><tbody>' +
      '<tr><td>normal</td><td><span class="health-badge normal">健康</span></td><td>OPEN DI &lt; 20 且 SLA超时 &le; 2</td></tr>' +
      '<tr><td>warning</td><td><span class="health-badge warning">一般</span></td><td>OPEN DI &lt; 40 且 SLA超时 &le; 5</td></tr>' +
      '<tr><td>caution</td><td><span style="display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;background:#fff2f0;color:#f53f3f">较差</span></td><td>OPEN DI &lt; 80 且 SLA超时 &le; 20</td></tr>' +
      '<tr><td>danger</td><td><span style="display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:600;background:#f53f3f;color:#fff">严重</span></td><td>超出以上任一条件</td></tr>' +
      '</tbody></table></div>' +
    '<div class="hr-table-wrap">' +
      '<div class="hr-table-header"><div class="hr-table-header-dot"></div><div class="hr-table-header-title">四、数据过滤规则</div></div>' +
      '<table class="hr-table"><thead><tr><th style="width:140px">规则</th><th>说明</th></tr></thead><tbody>' +
      '<tr><td>过期已关闭过滤</td><td>解决超90天列=超90天，且任务状态=已关闭，且创建时间距今超过90天的任务，不计入任何统计。</td></tr>' +
      '<tr><td>DI部门过滤</td><td>DI相关计算（OPEN DI/已解决DI/未解决DI）仅统计DB-部门=AIOT的任务，非AIOT部门DI值为0。</td></tr>' +
      '<tr><td>风险识别范围</td><td>风险分析仅针对AIOT部门人员，侧重SLA超时、BLOCK问题数和OPEN DI三个维度。</td></tr>' +
      '<tr><td>趋势计算</td><td>每日OPEN DI = 当天及之前创建且尚未关闭的AIOT BUG的DI值累计。趋势图展示近30天。</td></tr>' +
      '</tbody></table></div>' +
    '<div class="hr-table-wrap">' +
      '<div class="hr-table-header"><div class="hr-table-header-dot"></div><div class="hr-table-header-title">五、数据来源</div></div>' +
      '<table class="hr-table"><thead><tr><th style="width:140px">数据项</th><th>来源文件</th></tr></thead><tbody>' +
      '<tr><td>项目BUG数据</td><td>data/目录下各项目.xlsx文件（第一个sheet），需包含：标题、任务状态、DB-DI值、DB-部门、DB-角色、DB-SLA超时、创建时间、解决超90天等列。</td></tr>' +
      '<tr><td>里程碑</td><td>milestone/milestone.xlsx，支持合并单元格。</td></tr>' +
      '<tr><td>版本计划</td><td>version_plan/Version_Plan.xlsx，每个sheet对应一个项目。</td></tr>' +
      '</tbody></table></div>';
}

function buildRankingColumns() {
  // 风险最多Top3：按 block*10 + sla*3 + open_di 降序
  var riskList = PROJECT_NAMES.map(function(p) {
    var s = PROJECTS_DATA[p].stats;
    return { name:p, score: s.block * 10 + s.sla * 3 + s.open_di, block:s.block, sla:s.sla, open_di:s.open_di };
  });
  riskList.sort(function(a, b) { return b.score - a.score; });
  var riskTop3 = riskList.slice(0, 3);

  // 问题最多Top3：按未关闭问题数(total - resolved)降序
  var issueList = PROJECT_NAMES.map(function(p) {
    var s = PROJECTS_DATA[p].stats;
    return { name:p, openCount: s.total - s.resolved, total:s.total, resolved:s.resolved, open_di:s.open_di, sla:s.sla, block:s.block };
  });
  issueList.sort(function(a, b) { return b.openCount - a.openCount; });
  var issueTop3 = issueList.slice(0, 3);

  function renderRankItem(item, idx, type) {
    var numCls = idx === 0 ? 'r1' : idx === 1 ? 'r2' : 'r3';
    var html = '<div class="ranking-item">';
    html += '<div class="ranking-num ' + numCls + '">' + (idx + 1) + '</div>';
    html += '<div class="ranking-info">';
    html += '<div class="ranking-name">' + item.name + '</div>';
    html += '<div class="ranking-metrics">';
    if (type === 'risk') {
      html += '<span>BLOCK: <b style="color:' + (item.block > 0 ? 'var(--red)' : 'var(--text)') + '">' + item.block + '</b></span>';
      html += '<span>SLA超时: <b style="color:' + (item.sla > 0 ? 'var(--orange)' : 'var(--text)') + '">' + item.sla + '</b></span>';
      html += '<span>OPEN DI: <b style="color:var(--red)">' + item.open_di.toFixed(1) + '</b></span>';
    } else {
      html += '<span>未关闭: <b style="color:var(--red)">' + item.openCount + '</b></span>';
      html += '<span>总计: <b>' + item.total + '</b></span>';
      html += '<span>OPEN DI: <b style="color:var(--red)">' + item.open_di.toFixed(1) + '</b></span>';
    }
    html += '</div></div>';
    if (type === 'risk') {
      html += '<div class="ranking-badge" style="background:#fff2f0;color:var(--red)">风险 ' + item.score.toFixed(1) + '</div>';
    } else {
      html += '<div class="ranking-badge" style="background:#fff7e6;color:var(--orange)">' + item.openCount + '个</div>';
    }
    html += '</div>';
    return html;
  }

  var leftHtml = '';
  riskTop3.forEach(function(item, idx) { leftHtml += renderRankItem(item, idx, 'risk'); });
  if (riskTop3.length === 0) leftHtml = '<div style="text-align:center;padding:20px;color:var(--text-secondary)">暂无数据</div>';

  var rightHtml = '';
  issueTop3.forEach(function(item, idx) { rightHtml += renderRankItem(item, idx, 'issue'); });
  if (issueTop3.length === 0) rightHtml = '<div style="text-align:center;padding:20px;color:var(--text-secondary)">暂无数据</div>';

  return '<div class="ranking-section">' +
    '<div class="ranking-col">' +
      '<div class="ranking-col-title"><div class="ranking-icon" style="background:linear-gradient(135deg,#f53f3f,#ff7d00)">!</div>风险最多项目 Top3</div>' +
      leftHtml +
    '</div>' +
    '<div class="ranking-col">' +
      '<div class="ranking-col-title"><div class="ranking-icon" style="background:linear-gradient(135deg,#ff7d00,#ffb400)">#</div>问题最多项目 Top3</div>' +
      rightHtml +
    '</div>' +
  '</div>';
}

function renderOverview() {
  var tableRows = '';
  PROJECT_NAMES.forEach(function(p) {
    var s = PROJECTS_DATA[p].stats;
    var rr = parseFloat(s.resolve_rate);
    var rrColor = rr >= 90 ? '#00b42a' : rr >= 75 ? '#ff7d00' : '#f53f3f';
    var rrBg = rr >= 90 ? '#e8ffea' : rr >= 75 ? '#fff7e6' : '#fff2f0';
    var pr = s.pending_regression || 0;
    tableRows += '<tr>' +
      '<td><span style="font-weight:600">' + p + '</span></td>' +
      '<td><span style="font-weight:700;font-size:14px;color:var(--red)">' + s.open_di.toFixed(1) + '</span></td>' +
      '<td><span style="font-weight:600;color:' + (pr > 0 ? 'var(--orange)' : 'var(--text)') + '">' + pr.toFixed(1) + '</span></td>' +
      '<td><div class="hr-metric"><span class="hr-pill" style="background:' + rrBg + ';color:' + rrColor + '">' + s.resolve_rate + '%</span></div></td>' +
      '<td><span class="hr-pill" style="background:' + (s.sla > 0 ? '#fff7e6' : '#f2f3f5') + ';color:' + (s.sla > 0 ? '#ff7d00' : 'var(--text)') + '">' + s.sla + '</span></td>' +
      '<td><span class="hr-pill" style="background:' + (s.block > 0 ? '#fff2f0' : '#f2f3f5') + ';color:' + (s.block > 0 ? '#f53f3f' : 'var(--text)') + '">' + s.block + '</span></td>' +
      '<td><span style="font-weight:600">' + s.resolved + '/' + s.total + '</span></td>' +
      '</tr>';
  });
  return '<div class="page-header"><div class="page-title">项目总览</div><div class="page-desc">所有项目健康度汇总 · 近30天趋势</div></div>' +
    '<div class="chart-container"><div class="chart-title">OPEN DI值趋势</div><div id="trendChart" class="chart-area"></div></div>' +
    buildRankingColumns() +
    buildOverviewRisk() +
    '<div class="hr-table-wrap">' +
      '<div class="hr-table-header">' +
        '<div class="hr-table-header-dot"></div>' +
        '<div class="hr-table-header-title">项目列表</div>' +
        '<div class="hr-table-header-count">' + PROJECT_NAMES.length + '个项目</div>' +
      '</div>' +
      '<table class="hr-table">' +
      '<thead><tr>' +
      '<th>项目名</th>' +
      '<th>OPEN DI</th>' +
      '<th>待回归</th>' +
      '<th>解决率</th>' +
      '<th>SLA超时题数</th>' +
      '<th>BLOCK问题数</th>' +
      '<th>BUG数量</th>' +
      '</tr></thead><tbody>' + tableRows + '</tbody></table></div>';
}

function buildOverviewRisk() {
  // 聚合AIOT部门人员统计（风险识别仅针对AIOT）
  var pm = {};
  var aggStats = {total:0, resolved:0, open_di:0, sla:0, block:0, resolve_rate:'0.0', health:'normal', pending_regression:0};
  PROJECT_NAMES.forEach(function(pn) {
    var bugs = PROJECTS_DATA[pn].bugs;
    var s = PROJECTS_DATA[pn].stats;
    aggStats.total += s.total;
    aggStats.resolved += s.resolved;
    aggStats.open_di += s.open_di;
    aggStats.sla += s.sla;
    aggStats.block += s.block;
    aggStats.pending_regression += (s.pending_regression || 0);
    bugs.forEach(function(b) {
      if (b.db_dept !== 'AIOT') return;
      if (!pm[b.assignee]) pm[b.assignee] = {name:b.assignee, total:0, resolved:0, reopen:0, openDI:0, sla:0, slaDaysSum:0, slaDaysCount:0, role:b.db_role||'', dept:b.db_dept||''};
      pm[b.assignee].total++;
      if (b.status === '已解决' || b.status === '已关闭') pm[b.assignee].resolved++;
      if (b.status === '重新打开') pm[b.assignee].reopen++;
      pm[b.assignee].openDI += b.openDI;
      pm[b.assignee].sla += (b.sla_timeout || 0);
      if (b.sla_days > 0) { pm[b.assignee].slaDaysSum += b.sla_days; pm[b.assignee].slaDaysCount++; }
    });
  });
  aggStats.resolve_rate = aggStats.total > 0 ? (aggStats.resolved / aggStats.total * 100).toFixed(1) : '0.0';
  var persons = Object.values(pm);
  return buildRiskHtml(persons, aggStats);
}

function renderProjectDetail(pn) {
  var bugs = PROJECTS_DATA[pn].bugs;
  var s = PROJECTS_DATA[pn].stats;

  // 人员统计
  var pm = {};
  bugs.forEach(function(b) {
    if (!pm[b.assignee]) pm[b.assignee] = {name:b.assignee, total:0, resolved:0, reopen:0, openDI:0, pendingReturnDI:0, pendingHandleDI:0, sla:0, slaDaysSum:0, slaDaysCount:0, role:b.db_role || '', dept:b.db_dept || ''};
    pm[b.assignee].total++;
    if (b.status === '已解决' || b.status === '已关闭') pm[b.assignee].resolved++;
    if (b.status === '重新打开') pm[b.assignee].reopen++;
    pm[b.assignee].openDI += b.openDI;
    if (b.db_task_status === '待回归') pm[b.assignee].pendingReturnDI += b.openDI;
    if (b.db_task_status === '待处理') pm[b.assignee].pendingHandleDI += b.openDI;
    pm[b.assignee].sla += (b.sla_timeout || 0);
    if (b.sla_days > 0) { pm[b.assignee].slaDaysSum += b.sla_days; pm[b.assignee].slaDaysCount++; }
  });
  var persons = Object.values(pm);
  // 风险识别仅用AIOT部门人员
  var aiotPersons = persons.filter(function(p) { return p.dept === 'AIOT'; });

  // 风险识别
  var riskHtml = buildRiskHtml(aiotPersons, s);

  // 人力视图 - 按部门分组，统一表格列对齐，按OPEN DI降序
  var deptGroups = {};
  persons.forEach(function(p) {
    var dept = p.dept || '未知';
    if (!deptGroups[dept]) deptGroups[dept] = [];
    deptGroups[dept].push(p);
  });

  var pc = '';
  var deptOrder = ['AIOT','整机','其他','未知'];
  var isFirstDept = true;
  deptOrder.forEach(function(dept) {
    if (!deptGroups[dept]) return;
    var list = deptGroups[dept];
    // 按OPEN DI降序排序
    list.sort(function(a, b) {
      var diDiff = b.openDI - a.openDI;
      if (diDiff !== 0) return diDiff;
      return a.name.localeCompare(b.name);
    });

    // 部门分隔行
    if (!isFirstDept) {
      pc += '<tr class="hr-dept-divider"><td colspan="9"><div class="hr-table-header" style="margin:12px 0 6px"><div class="hr-table-header-dot"></div><div class="hr-table-header-title">' + dept + '</div><div class="hr-table-header-count">' + list.length + '人</div></div></td></tr>';
    } else {
      pc += '<tr class="hr-dept-divider"><td colspan="9"><div class="hr-table-header" style="margin:0 0 6px"><div class="hr-table-header-dot"></div><div class="hr-table-header-title">' + dept + '</div><div class="hr-table-header-count">' + list.length + '人</div></div></td></tr>';
      isFirstDept = false;
    }

    list.forEach(function(p) {
      var rate = p.total > 0 ? (p.resolved/p.total*100).toFixed(1) : '0.0';
      var rr = parseFloat(rate);
      var rrColor = rr >= 90 ? '#00b42a' : rr >= 75 ? '#ff7d00' : '#f53f3f';
      var rrBg = rr >= 90 ? '#e8ffea' : rr >= 75 ? '#fff7e6' : '#fff2f0';
      var slaAvg = p.slaDaysCount > 0 ? (p.slaDaysSum / p.slaDaysCount).toFixed(1) : '-';
      pc += '<tr>' +
        '<td>' + p.name + '</td>' +
        '<td><span style="font-size:12px;color:var(--text-secondary)">' + (p.role || '-') + '</span></td>' +
        '<td><div class="hr-metric"><span class="hr-pill" style="background:' + rrBg + ';color:' + rrColor + '">' + rate + '%</span></div></td>' +
        '<td><span style="font-weight:700;font-size:14px;color:var(--red)">' + p.openDI.toFixed(1) + '</span></td>' +
        '<td><span style="font-weight:600;color:var(--orange)">' + p.pendingReturnDI.toFixed(1) + '</span></td>' +
        '<td><span style="font-weight:600;color:var(--text)">' + p.pendingHandleDI.toFixed(1) + '</span></td>' +
        '<td><div class="hr-metric"><span class="hr-pill" style="background:' + (p.sla > 0 ? '#fff7e6' : '#f2f3f5') + ';color:' + (p.sla > 0 ? '#ff7d00' : 'var(--text)') + '">' + p.sla + '</span>' + (p.slaDaysCount > 0 ? '<span class="hr-metric-sub">均' + slaAvg + '天</span>' : '') + '</div></td>' +
        '<td><span class="hr-pill" style="background:' + (p.reopen > 0 ? '#fff7e6' : '#f2f3f5') + ';color:' + (p.reopen > 0 ? 'var(--orange)' : 'var(--text-secondary)') + '">' + p.reopen + '</span></td>' +
        '<td><span style="font-weight:600">' + p.resolved + '/' + p.total + '</span></td>' +
        '</tr>';
    });
  });

  var tableHtml = '<div class="hr-table-wrap"><table class="hr-table">' +
    '<thead><tr>' +
    '<th>人员</th>' +
    '<th>角色</th>' +
    '<th>解决率</th>' +
    '<th>OPEN DI</th>' +
    '<th>待回归DI</th>' +
    '<th>待处理DI</th>' +
    '<th>SLA超时</th>' +
    '<th>重开次数</th>' +
    '<th>BUG数量</th>' +
    '</tr></thead><tbody>' + pc + '</tbody></table></div>';

  return '<div class="page-header"><div class="page-title">' + pn + '项目</div></div>' +
    '<div class="chart-container"><div class="chart-title">OPEN DI趋势</div><div id="projectTrendChart" class="chart-area"></div></div>' +
    riskHtml +
    '<div style="font-size:15px;font-weight:600;margin:18px 0 14px">人力视图</div>' +
    tableHtml;
}

function renderHR() {
  var personDeptMap = {};
  var personRoleMap = {};
  var personPendingHandle = {};
  var personPendingReturn = {};
  PROJECT_NAMES.forEach(function(pn) {
    PROJECTS_DATA[pn].bugs.forEach(function(b) {
      if (b.assignee) {
        if (!personDeptMap[b.assignee]) {
          personDeptMap[b.assignee] = b.db_dept || "未知";
        }
        if (!personRoleMap[b.assignee]) {
          personRoleMap[b.assignee] = b.db_role || "Other";
        }
        if (!personPendingHandle[b.assignee]) personPendingHandle[b.assignee] = 0;
        if (!personPendingReturn[b.assignee]) personPendingReturn[b.assignee] = 0;
        if (b.db_task_status === '待处理') personPendingHandle[b.assignee]++;
        if (b.db_task_status === '待回归') personPendingReturn[b.assignee]++;
      }
    });
  });

  var roleOrder = {'AIoT开发': 0, '软件测试': 1, 'Other': 2};
  function roleSortKey(role) {
    return roleOrder[role] !== undefined ? roleOrder[role] : 99;
  }

  // 按部门分组，部门内按OPEN DI降序排序，同DI值再按角色和姓名排序
  var deptGroups = {};
  PERSON_STATS.forEach(function(p) {
    var dept = personDeptMap[p.name] || "未知";
    // 只显示 AIoT 部门的人员
    if (dept !== 'AIoT' && dept !== 'AIOT') return;
    if (!deptGroups[dept]) deptGroups[dept] = [];
    deptGroups[dept].push(p);
  });
  var sections = '<div class="chart-title" style="margin-bottom:20px">人力资源</div>';
  var deptOrder = ['AIOT','AIoT'];
  var allPersons = [];
  deptOrder.forEach(function(dept) {
    if (!deptGroups[dept]) return;
    var list = deptGroups[dept];
    // 按OPEN DI降序排序，同DI值再按角色和姓名排序
    list.sort(function(a, b) {
      var diDiff = (b.total_solved_di + b.total_unsolved_di) - (a.total_solved_di + a.total_unsolved_di);
      if (diDiff !== 0) return diDiff;
      var ra = roleSortKey(personRoleMap[a.name]);
      var rb = roleSortKey(personRoleMap[b.name]);
      if (ra !== rb) return ra - rb;
      return a.name.localeCompare(b.name);
    });
    var tableRows = '';
    list.forEach(function(p, idx) {
      allPersons.push(p);
      var phCount = personPendingHandle[p.name] || 0;
      var prCount = personPendingReturn[p.name] || 0;
      var slaAvg = p.sla_days_count > 0 ? (p.sla_days_sum / p.sla_days_count).toFixed(1) : '-';
      // SLA超时颜色：0黑色，>0红色
      var slaTextColor = p.total_sla > 0 ? 'var(--red)' : 'var(--text)';
      // 重开次数颜色：0黑色，>0红色
      var reopenTextColor = p.total_reopen > 0 ? 'var(--red)' : 'var(--text)';
      var role = personRoleMap[p.name] || 'Other';
      tableRows += '<tr>' +
        '<td>' + p.name + '</td>' +
        '<td><span style="font-size:12px;color:var(--text-secondary)">' + role + '</span></td>' +
        '<td style="font-weight:600">' + phCount + '</td>' +
        '<td style="font-weight:600;color:' + (prCount > 0 ? 'var(--orange)' : 'var(--text)') + '">' + prCount + '</td>' +
        '<td style="color:' + (p.total_open_di > 0 ? 'var(--red)' : 'var(--text)') + '">' + p.total_open_di.toFixed(1) + '</td>' +
        '<td style="color:' + slaTextColor + '">' + p.total_sla + (p.sla_days_count > 0 ? '（均' + slaAvg + '天）' : '') + '</td>' +
        '<td style="color:' + reopenTextColor + '">' + p.total_reopen + '</td>' +
        '</tr>';
    });
    sections += '<div class="hr-table-wrap">' +
      '<div class="hr-table-header">' +
        '<div class="hr-table-header-dot"></div>' +
        '<div class="hr-table-header-title">' + dept + '</div>' +
        '<div class="hr-table-header-count">' + list.length + '人</div>' +
      '</div>' +
      '<table class="hr-table">' +
      '<thead><tr>' +
      '<th>人员</th>' +
      '<th>角色</th>' +
      '<th>待处理</th>' +
      '<th>待回归</th>' +
      '<th>OPEN DI</th>' +
      '<th>SLA超时</th>' +
      '<th>重开次数</th>' +
      '</tr></thead><tbody>' + tableRows + '</tbody></table></div>';
  });

  // 资源分布图 - 按项目分色的已解决DI + 未解决DI
  // 收集每个人在各项目的DI数据
  var personProjectDi = {};
  PROJECT_NAMES.forEach(function(pn) {
    PROJECTS_DATA[pn].bugs.forEach(function(b) {
      if (b.assignee) {
        if (!personProjectDi[b.assignee]) {
          personProjectDi[b.assignee] = { solved: {}, unsolved: {} };
        }
        var di = b.openDI || 0;
        if (b.status === "已关闭" || b.status === "已解决") {
          personProjectDi[b.assignee].solved[pn] = (personProjectDi[b.assignee].solved[pn] || 0) + di;
        } else {
          personProjectDi[b.assignee].unsolved[pn] = (personProjectDi[b.assignee].unsolved[pn] || 0) + di;
        }
      }
    });
  });

  var maxDi = 0;
  allPersons.forEach(function(p) {
    maxDi = Math.max(maxDi, p.total_solved_di + p.total_unsolved_di);
  });
  if (maxDi === 0) maxDi = 1;

  // 为每个项目分配颜色
  var projectColors = {};
  var colors = ['#165dff', '#00b42a', '#ff7d00', '#f53f3f', '#722ed1', '#faad14', '#13c2c2', '#eb2f96'];
  PROJECT_NAMES.forEach(function(pn, idx) {
    projectColors[pn] = colors[idx % colors.length];
  });

  var chartRows = '';
  allPersons.forEach(function(p) {
    var pData = personProjectDi[p.name] || { solved: {}, unsolved: {} };
    
    // 计算总宽度
    var totalDi = p.total_solved_di + p.total_unsolved_di;
    var segments = '';
    var currentPos = 0;
    
    // 先渲染已解决的DI（按项目）
    PROJECT_NAMES.forEach(function(pn) {
      var solvedDi = pData.solved[pn] || 0;
      if (solvedDi > 0) {
        var width = (solvedDi / maxDi * 100).toFixed(1);
        var showVal = parseFloat(width) > 3;  // 降低阈值到3%
        segments += '<div class="pp-bar-segment" style="width:' + width + '%;background:' + projectColors[pn] + '">' +
          (showVal ? '<span class="pp-bar-value">' + solvedDi.toFixed(1) + '</span>' : '') +
        '</div>';
      }
    });
    
    // 再渲染未解决的DI（按项目）
    PROJECT_NAMES.forEach(function(pn) {
      var unsolvedDi = pData.unsolved[pn] || 0;
      if (unsolvedDi > 0) {
        var width = (unsolvedDi / maxDi * 100).toFixed(1);
        var showVal = parseFloat(width) > 3;  // 降低阈值到3%
        segments += '<div class="pp-bar-segment" style="width:' + width + '%;background:' + projectColors[pn] + ';opacity:0.6">' +
          (showVal ? '<span class="pp-bar-value">' + unsolvedDi.toFixed(1) + '</span>' : '') +
        '</div>';
      }
    });
    
    chartRows += '<div class="pp-row">' +
      '<div class="pp-name">' + p.name + '</div>' +
      '<div style="flex:1;display:flex;align-items:center">' +
        '<div style="flex:1">' +
          '<div class="pp-bar-line" style="position:relative">' +
            segments +
          '</div>' +
        '</div>' +
        '<div class="pp-bar-sla">SLA超时:' + p.total_sla + '</div>' +
      '</div>' +
    '</div>';
  });

  // 构建图例
  var legendItems = '';
  PROJECT_NAMES.forEach(function(pn) {
    legendItems += '<span><i style="background:' + projectColors[pn] + '"></i>' + pn + '</span>';
  });

  // 重开次数独立柱状图 - 统一橙色
  var reopenPersons = allPersons.filter(function(p) { return p.total_reopen > 0; });
  reopenPersons.sort(function(a, b) { return b.total_reopen - a.total_reopen; });
  var maxReopen = reopenPersons.length > 0 ? reopenPersons[0].total_reopen : 1;
  if (maxReopen === 0) maxReopen = 1;

  var reopenRows = '';
  reopenPersons.forEach(function(p) {
    var width = (p.total_reopen / maxReopen * 100).toFixed(1);
    var showVal = parseFloat(width) > 3;
    var segment = '<div class="pp-bar-segment" style="width:' + width + '%;background:var(--orange)">' +
      (showVal ? '<span class="pp-bar-value">' + p.total_reopen + '</span>' : '') +
    '</div>';

    reopenRows += '<div class="pp-row">' +
      '<div class="pp-name">' + p.name + '</div>' +
      '<div style="flex:1;display:flex;align-items:center">' +
        '<div style="flex:1">' +
          '<div class="pp-bar-line" style="position:relative">' +
            segment +
          '</div>' +
        '</div>' +
        '<div class="pp-bar-sla">重开:' + p.total_reopen + '</div>' +
      '</div>' +
    '</div>';
  });

  var reopenChart = '';
  if (reopenPersons.length > 0) {
    reopenChart = '<div class="pp-chart"><div class="chart-title">代码质量</div>' +
      reopenRows + '</div>';
  }

  return sections +
    '<div class="pp-chart"><div class="chart-title">资源分布</div>' +
    '<div class="pp-legend">' + legendItems + '</div>' +
    chartRows + '</div>' +
    reopenChart;
}

function renderTimelinePage() {
  return '<div class="page-header"><div class="page-title">项目大盘</div><div class="page-desc">各项目里程碑计划与进度跟踪</div></div>' +
    '<div class="timeline-section">' +
      '<div style="display:flex;justify-content:flex-end;margin-bottom:6px"><button class="tl-reset-btn" onclick="resetTimeline()">回到当前</button></div>' +
      '<div class="tl-timeline-wrap">' +
        '<div class="tl-viewport" id="tlViewport"></div>' +
      '</div>' +
    '</div>';
}

function renderVersionPlan() {
  var plans = VERSION_PLAN || [];
  if (plans.length === 0) {
    return '<div class="page-header"><div class="page-title">版本计划</div><div class="page-desc">各项目版本开发计划</div></div>' +
      '<div style="text-align:center;padding:60px;color:var(--text-secondary)">暂无版本计划数据</div>';
  }
  var cards = '';
  plans.forEach(function(p) {
    var overview = p.overview ? p.overview.replace(/\n/g, '<br>') : '';
    cards += '<div class="stat-card" style="cursor:default">' +
      '<div style="font-size:16px;font-weight:700;margin-bottom:10px;color:var(--blue)">' + p.name + '</div>' +
      (overview ? '<div style="font-size:13px;color:var(--text-secondary);line-height:1.8;margin-bottom:12px">' + overview + '</div>' : '') +
      (p.link ? '<a href="' + p.link + '" target="_blank" style="display:inline-block;padding:4px 12px;background:var(--blue);color:#fff;border-radius:4px;font-size:12px;text-decoration:none;font-weight:600">链接</a>' : '');
    cards += '</div>';
  });
  return '<div class="page-header"><div class="page-title">版本计划</div><div class="page-desc">各项目版本开发计划</div></div>' +
    '<div class="cards-grid">' + cards + '</div>';
}

function initCharts() {
  if (document.getElementById('trendChart')) {
    var chart = echarts.init(document.getElementById('trendChart'));
    chart.setOption({
      tooltip: { trigger:'axis', axisPointer:{type:'cross'} },
      legend: { data:['OPEN DI','每日新增','每日解决'], top:0 },
      xAxis: { type:'category', data:TREND_DATES, axisLabel:{fontSize:10} },
      yAxis: { type:'value', name:'OPEN DI' },
      series: [
        { name:'OPEN DI', type:'line', data:CUMULATIVE_DI, smooth:true, lineStyle:{width:2} },
        { name:'每日新增', type:'bar', data:DAILY_NEW, itemStyle:{color:'#f53f3f'}, barWidth:6 },
        { name:'每日解决', type:'bar', data:DAILY_RESOLVED, itemStyle:{color:'#00b42a'}, barWidth:6 }
      ]
    });
  }
  if (document.getElementById('projectTrendChart')) {
    var pt = PROJECT_TRENDS[selectedProject] || [[],[],[]];
    var chart = echarts.init(document.getElementById('projectTrendChart'));
    chart.setOption({
      tooltip: { trigger:'axis', axisPointer:{type:'cross'} },
      legend: { data:['OPEN DI','每日新增','每日解决'], top:0 },
      xAxis: { type:'category', data:TREND_DATES, axisLabel:{fontSize:10} },
      yAxis: { type:'value', name:'OPEN DI' },
      series: [
        { name:'OPEN DI', type:'line', data:pt[0], smooth:true, lineStyle:{width:2} },
        { name:'每日新增', type:'bar', data:pt[1], itemStyle:{color:'#f53f3f'}, barWidth:6 },
        { name:'每日解决', type:'bar', data:pt[2], itemStyle:{color:'#00b42a'}, barWidth:6 }
      ]
    });
  }
  renderTimeline();
}

var tlViewStart = null; // 当前可视窗口起始日
var tlDayWidth = 64;    // 每天像素宽度
var tlDataStart = null; // 数据最早日期
var tlDataEnd = null;   // 数据最晚日期
var tlRenderStart = null; // 渲染起始日期

function renderTimeline() {
  var viewport = document.getElementById('tlViewport');
  if (!viewport) return;

  var projects = Object.keys(MILESTONES);
  if (projects.length === 0) { viewport.innerHTML = '<div style="text-align:center;padding:40px;color:var(--text-secondary)">暂无里程碑数据</div>'; return; }

  // 先显示加载状态，让浏览器有机会绘制
  viewport.innerHTML = '<div style="text-align:center;padding:60px;color:var(--text-secondary)"><div style="font-size:14px">正在加载时间线...</div></div>';

  requestAnimationFrame(function() {
  _renderTimelineInner(viewport, projects);
  });
}

function _renderTimelineInner(viewport, projects) {

  var allDates = [];
  projects.forEach(function(pn) {
    (MILESTONES[pn] || []).forEach(function(m) {
      if (m.date) { var d = new Date(m.date); if (!isNaN(d.getTime())) allDates.push(d); }
    });
  });
  if (allDates.length === 0) { viewport.innerHTML = '<div style="text-align:center;padding:40px;color:var(--text-secondary)">暂无有效日期的里程碑数据</div>'; return; }

  var today = new Date(); today.setHours(0,0,0,0);

  // 项目排序：日期越靠近当前越靠前，"已延期"状态优先
  projects.sort(function(a, b) {
    var msA = MILESTONES[a] || [];
    var msB = MILESTONES[b] || [];
    var hasDelayA = msA.some(function(m) { return m.status === '已延期'; });
    var hasDelayB = msB.some(function(m) { return m.status === '已延期'; });
    // 已延期优先
    if (hasDelayA && !hasDelayB) return -1;
    if (!hasDelayA && hasDelayB) return 1;
    // 找最近的有效日期（非"已通过"）
    function closestDate(msList) {
      var futureDates = msList.filter(function(m) {
        return m.date && m.status !== '已通过' && !isNaN(new Date(m.date).getTime());
      }).map(function(m) { return new Date(m.date); });
      if (futureDates.length === 0) {
        // 如果没有未来日期，用所有日期中最近的
        futureDates = msList.filter(function(m) {
          return m.date && !isNaN(new Date(m.date).getTime());
        }).map(function(m) { return new Date(m.date); });
      }
      if (futureDates.length === 0) return new Date(8640000000000000);
      futureDates.sort(function(x, y) {
        return Math.abs(x - today) - Math.abs(y - today);
      });
      return futureDates[0];
    }
    var dateA = closestDate(msA);
    var dateB = closestDate(msB);
    return Math.abs(dateA - today) - Math.abs(dateB - today);
  });

  // 数据范围（前后各扩展3天）
  tlDataStart = new Date(Math.min.apply(null, allDates));
  tlDataEnd = new Date(Math.max.apply(null, allDates));
  tlDataStart.setDate(tlDataStart.getDate() - 3);
  tlDataEnd.setDate(tlDataEnd.getDate() + 3);

  // 默认视图：今天前3天起，共15天
  if (tlViewStart === null) {
    tlViewStart = new Date(today);
    tlViewStart.setDate(tlViewStart.getDate() - 3);
  }

  var viewDays = 15;
  var viewWidth = viewDays * tlDayWidth;

  // 渲染范围 = 数据范围 + 前后各扩展5天，最大180天
  var renderStart = new Date(tlDataStart); renderStart.setDate(renderStart.getDate() - 5);
  var renderEnd = new Date(tlDataEnd); renderEnd.setDate(renderEnd.getDate() + 5);
  var defaultViewEnd = new Date(today); defaultViewEnd.setDate(defaultViewEnd.getDate() + 12);
  if (renderStart > tlViewStart) renderStart = new Date(tlViewStart);
  if (renderEnd < defaultViewEnd) renderEnd = new Date(defaultViewEnd);
  // 限制最大渲染范围365天，防止DOM节点过多导致卡顿
  var maxDays = 365;
  if (Math.ceil((renderEnd - renderStart) / 86400000) > maxDays) {
    renderEnd = new Date(renderStart); renderEnd.setDate(renderEnd.getDate() + maxDays);
  }
  tlRenderStart = renderStart; // 保存为全局变量

  var totalDays = Math.ceil((renderEnd - renderStart) / 86400000);
  var innerWidth = totalDays * tlDayWidth;
  var padLeft = Math.round((tlDataStart - renderStart) / 86400000) * tlDayWidth;

  // 滚动范围限制
  var earliestPx = 0;
  var latestPx = innerWidth - viewWidth;
  // 初始视图居中到今天（与"回到当前"按钮行为一致）
  var todayOffsetPx = Math.round((today - renderStart) / 86400000) * tlDayWidth;
  var defaultStartPx = Math.max(earliestPx, Math.min(todayOffsetPx - viewWidth / 2, latestPx));

  function datePx(dateStr) {
    var d = new Date(dateStr); d.setHours(0,0,0,0);
    var offset = Math.round((d - renderStart) / 86400000);
    return offset * tlDayWidth;
  }
  function inRenderRange(dateStr) {
    var d = new Date(dateStr); d.setHours(0,0,0,0);
    return d >= renderStart && d <= renderEnd;
  }
  function statusDotCls(s) {
    if (s === '已通过') return ' pass';
    if (s === '已延期') return ' delay';
    if (s === '处理中') return ' progress';
    return '';
  }
  function statusCardCls(s) {
    if (s === '已通过') return 'pass';
    if (s === '已延期') return 'delay';
    if (s === '处理中') return 'progress';
    return 'default';
  }
  function fmtDate(ds) { return ds.length >= 10 ? ds.slice(5).replace('-', '/') : ds; }

  // === 固定项目名称列 ===
  var namesHtml = '';
  namesHtml += '<div class="tl-name-spacer"></div>';
  projects.forEach(function(pn) {
    var nodes = (MILESTONES[pn] || []).filter(function(m) { return inRenderRange(m.date); });
    if (nodes.length === 0) return;
    // 计算卡片区高度（与时间线体保持一致）
    var CARD_GAP_N = 160;
    var ROW_HEIGHT_N = 130;
    var sortedN = nodes.slice().sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
    var placedN = [];
    sortedN.forEach(function(m) {
      var px = datePx(m.date);
      var row = 0;
      while (true) {
        var conflict = false;
        for (var k = 0; k < placedN.length; k++) {
          if (placedN[k].row === row && Math.abs(placedN[k].px - px) < CARD_GAP_N) { conflict = true; break; }
        }
        if (!conflict) break;
        row++;
      }
      placedN.push({ px: px, row: row });
    });
    var maxRowN = 0;
    for (var k = 0; k < placedN.length; k++) { if (placedN[k].row > maxRowN) maxRowN = placedN[k].row; }
    var areaHN = (maxRowN + 1) * ROW_HEIGHT_N + 20;

    namesHtml += '<div class="tl-project-row">';
    namesHtml += '<div class="tl-name-track"><div class="tl-project-name">' + pn + '</div></div>';
    namesHtml += '<div class="tl-name-card-area" style="min-height:' + areaHN + 'px"></div>';
    namesHtml += '</div>';
  });

  // === 可滚动时间线 ===
  var scrollHtml = '';

  // 日期列头部
  scrollHtml += '<div class="tl-day-cols" style="width:' + innerWidth + 'px">';
  for (var i = 0; i < totalDays; i++) {
    var d = new Date(renderStart.getTime() + i * 86400000);
    var isToday = d.getTime() === today.getTime();
    var isDataRange = d >= tlDataStart && d <= tlDataEnd;
    scrollHtml += '<div class="tl-day-col' + (isToday ? ' today-col' : '') + '" style="width:' + tlDayWidth + 'px">';
    scrollHtml += '<div class="tl-day-label">' + (d.getMonth()+1) + '/' + d.getDate() + '</div>';
    scrollHtml += '</div>';
  }
  scrollHtml += '</div>';

  // 项目内容区
  scrollHtml += '<div class="tl-timeline-body" style="width:' + innerWidth + 'px">';
  // 今日线
  var todayPx = datePx(today.toISOString().slice(0,10));
  scrollHtml += '<div class="tl-today-line" style="left:' + todayPx + 'px"><div class="tl-today-top-label">当前</div><div class="tl-today-label">' + (today.getMonth()+1) + '/' + today.getDate() + '</div></div>';

  projects.forEach(function(pn) {
    var nodes = (MILESTONES[pn] || []).filter(function(m) { return inRenderRange(m.date); });
    if (nodes.length === 0) return;

    scrollHtml += '<div class="tl-project-row">';
    // 轨道 + 圆点
    scrollHtml += '<div class="tl-track">';
    nodes.forEach(function(m) {
      scrollHtml += '<div class="tl-dot' + statusDotCls(m.status) + '" style="left:' + datePx(m.date) + 'px"></div>';
    });
    scrollHtml += '</div>';

    // 卡片区 — 碰撞检测
    var CARD_GAP = 160;
    var ROW_HEIGHT = 130;
    var sorted = nodes.slice().sort(function(a, b) { return new Date(a.date) - new Date(b.date); });
    var placed = [];
    sorted.forEach(function(m) {
      var px = datePx(m.date);
      var row = 0;
      while (true) {
        var conflict = false;
        for (var k = 0; k < placed.length; k++) {
          if (placed[k].row === row && Math.abs(placed[k].px - px) < CARD_GAP) { conflict = true; break; }
        }
        if (!conflict) break;
        row++;
      }
      placed.push({ px: px, row: row, node: m });
    });
    var maxRow = 0;
    for (var k = 0; k < placed.length; k++) { if (placed[k].row > maxRow) maxRow = placed[k].row; }
    var areaH = (maxRow + 1) * ROW_HEIGHT + 20;

    scrollHtml += '<div class="tl-card-area" style="min-height:' + areaH + 'px">';
    placed.forEach(function(item) {
      var m = item.node;
      var topOff = item.row * ROW_HEIGHT;
      scrollHtml += '<div class="tl-connector" style="left:' + item.px + 'px;top:0;height:' + (topOff + 20) + 'px"></div>';
      scrollHtml += '<div class="tl-card" style="left:' + item.px + 'px;top:' + (topOff + 20) + 'px">';
      scrollHtml += '<div class="tl-card-date">' + fmtDate(m.date) + '</div>';
      scrollHtml += '<div class="tl-card-name">' + m.name + '</div>';
      if (m.item) scrollHtml += '<div class="tl-card-item">' + m.item + '</div>';
      if (m.risk) scrollHtml += '<div class="tl-card-risk">风险项：' + m.risk + '</div>';
      if (m.status) scrollHtml += '<span class="tl-card-status ' + statusCardCls(m.status) + '">' + m.status + '</span>';
      scrollHtml += '</div>';
    });
    scrollHtml += '</div>'; // card-area
    scrollHtml += '</div>'; // project-row
  });
  scrollHtml += '</div>'; // timeline-body

  // 组装viewport：固定名称列 + 可滚动时间线
  viewport.innerHTML = '<div class="tl-names-col" id="tlNamesCol">' + namesHtml + '</div>' +
    '<div class="tl-scroll-area" id="tlScrollArea">' + scrollHtml + '</div>';

  // 垂直滚动同步
  var namesCol = document.getElementById('tlNamesCol');
  var scrollArea = document.getElementById('tlScrollArea');
  var syncing = false;
  scrollArea.addEventListener('scroll', function() {
    if (syncing) return;
    syncing = true;
    namesCol.scrollTop = scrollArea.scrollTop;
    syncing = false;
  });
  namesCol.addEventListener('scroll', function() {
    if (syncing) return;
    syncing = true;
    scrollArea.scrollTop = namesCol.scrollTop;
    syncing = false;
  });

  // 滚动到默认位置（数据最早日期处 = tlDataStart）
  scrollArea.scrollLeft = Math.max(earliestPx, Math.min(defaultStartPx, latestPx));

  // 鼠标拖拽横向滚动
  (function() {
    var isDragging = false;
    var startX = 0;
    var startScrollLeft = 0;
    scrollArea.addEventListener('mousedown', function(e) {
      // 点击卡片/链接等可交互元素时不启动拖拽
      if (e.target.closest('a, button, .tl-card')) return;
      isDragging = true;
      startX = e.pageX;
      startScrollLeft = scrollArea.scrollLeft;
      scrollArea.style.cursor = 'grabbing';
      e.preventDefault();
    });
    document.addEventListener('mousemove', function(e) {
      if (!isDragging) return;
      var dx = e.pageX - startX;
      scrollArea.scrollLeft = startScrollLeft - dx;
    });
    document.addEventListener('mouseup', function() {
      if (!isDragging) return;
      isDragging = false;
      scrollArea.style.cursor = 'grab';
    });
  })();
}

function panTimeline(dir) {
  var scrollArea = document.getElementById('tlScrollArea');
  if (!scrollArea) return;
  var scrollAmt = 7 * tlDayWidth;
  var newLeft = scrollArea.scrollLeft + dir * scrollAmt;
  // 限制滚动范围
  var viewWidth = 15 * tlDayWidth;
  var maxScroll = scrollArea.scrollWidth - viewWidth;
  newLeft = Math.max(0, Math.min(newLeft, maxScroll));
  scrollArea.scrollTo({ left: newLeft, behavior: 'smooth' });
}

function resetTimeline() {
  var today = new Date(); today.setHours(0,0,0,0);
  tlViewStart = new Date(today);
  tlViewStart.setDate(tlViewStart.getDate() - 3);
  renderTimeline();
  // 滚动到今天的位置
  setTimeout(function() {
    var scrollArea = document.getElementById('tlScrollArea');
    if (scrollArea && tlRenderStart) {
      var todayOffset = Math.round((today - tlRenderStart) / 86400000);
      var todayPx = todayOffset * tlDayWidth;
      var viewWidth = 15 * tlDayWidth;
      var targetLeft = todayPx - viewWidth / 2;
      var maxScroll = scrollArea.scrollWidth - viewWidth;
      targetLeft = Math.max(0, Math.min(targetLeft, maxScroll));
      scrollArea.scrollTo({ left: targetLeft, behavior: 'smooth' });
    }
  }, 100);
}

renderNav();
renderContent();
window.addEventListener('resize', function() {
  document.querySelectorAll('.chart-area').forEach(function(el) {
    var inst = echarts.getInstanceByDom(el);
    if (inst) inst.resize();
  });
});
"""

# 清理bug数据：移除内部字段 + 前端不需要的字段（减小HTML体积）
_FRONTEND_BUG_FIELDS = {"assignee", "status", "db_role", "db_dept", "db_task_status", "openDI", "sla_timeout", "sla_days"}
for pn in PROJECT_NAMES:
    for b in all_projects_bugs[pn]:
        # 删除所有前端不需要的字段
        for key in list(b.keys()):
            if key not in _FRONTEND_BUG_FIELDS:
                del b[key]

# ===== 生成 HTML =====
# 构建JS期望的数据结构: {项目名: {stats, bugs}}
projects_data_for_js = {}
for pn in PROJECT_NAMES:
    projects_data_for_js[pn] = {"stats": projects_stats[pn], "bugs": all_projects_bugs[pn]}

js_injected = js.replace('__PROJECTS_JSON__', json.dumps(projects_data_for_js, ensure_ascii=False)).replace('__TREND_DATES_JSON__', json.dumps(trend_dates_str)).replace('__CUMULATIVE_DI_JSON__', json.dumps(cumulative_di)).replace('__DAILY_NEW_JSON__', json.dumps(daily_new_di)).replace('__DAILY_RESOLVED_JSON__', json.dumps(daily_resolved_di)).replace('__MILESTONES_JSON__', json.dumps(milestones, ensure_ascii=False)).replace('__PERSON_STATS_JSON__', json.dumps(person_project_stats, ensure_ascii=False)).replace('__PROJECT_NAMES_JSON__', json.dumps(PROJECT_NAMES)).replace('__PROJECT_TRENDS_JSON__', json.dumps(project_trends, ensure_ascii=False)).replace('__VERSION_PLAN_JSON__', json.dumps(version_plan, ensure_ascii=False))
# 转义</script>防止浏览器提前截断
js_safe = js_injected.replace('</script>', '<\\/script>')

# 构建时间戳（用于缓存刷新）
build_ts = datetime.now().strftime('%Y%m%d%H%M%S')

html = '<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">\n<meta http-equiv="Pragma" content="no-cache">\n<meta http-equiv="Expires" content="0">\n<title>项目管理汇报看板</title>\n<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js?v=' + build_ts + '"></script>\n<style>' + css + '</style>\n</head>\n<body>\n<div class="sidebar">\n  <div class="sidebar-header"><div class="sidebar-logo">AI</div><div class="sidebar-title">项目管理看板</div></div>\n  <ul class="nav-menu" id="navMenu"></ul>\n</div>\n<div class="main" id="mainContent"></div>\n<script>var BUILD_TS="' + build_ts + '";' + js_safe + '</script>\n</body>\n</html>'

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("✅ index.html 生成成功！")
print(f"   - 项目数: {len(PROJECT_NAMES)}")
for pn in PROJECT_NAMES:
    s = projects_stats[pn]
    print(f"   - {pn}: OPEN DI={s['open_di']}, 待回归={s['pending_regression']}, SLA超时={s['sla']}, BLOCK={s['block']}, 解决率={s['resolve_rate']}%, 健康度={s['health']}")
print(f"   - 人员数: {len(person_project_stats)}")
print(f"   - 趋势天数: {len(trend_dates_str)}")

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
项目质量看板 - 钉钉推送脚本（钉盘版）
每天定时重建数据 → 上传钉盘 → 推送带链接的日报到钉钉群
"""

import subprocess
import json
import urllib.request
import urllib.error
import ssl
import os
import re
from datetime import datetime

# 配置
WEBHOOK_URL = "https://oapi.dingtalk.com/robot/send?access_token=bfce5da598f7d476bb8f8307c3da4f4ff35f0c9d3066bda13abeaa8de45bb1a5"
KEYWORD = "项目日报"
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = os.path.join(PROJECT_DIR, "index.html")


def run_build():
    """运行 build_dashboard.py 并捕获输出"""
    venv_python = os.path.join(PROJECT_DIR, ".venv", "bin", "python3")
    build_script = os.path.join(PROJECT_DIR, "build_dashboard.py")

    if not os.path.exists(venv_python):
        venv_python = "python3"

    try:
        result = subprocess.run(
            [venv_python, build_script],
            capture_output=True,
            text=True,
            timeout=60,
            cwd=PROJECT_DIR
        )
        return result.stdout + result.stderr
    except Exception as e:
        return f"构建失败：{str(e)}"


def parse_metrics(build_output):
    """从构建输出中解析关键指标 + 昨日趋势数据"""
    metrics = {
        "projects": [],
        "build_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "success": False,
        "yesterday_open_di": "-",
        "yesterday_solved_di": "-"
    }

    if "✅ index.html 生成成功" in build_output:
        metrics["success"] = True

    pattern = r'(\w+):\s*OPEN DI=([\d.]+),\s*SLA超时=(\d+),\s*BLOCK=(\d+),\s*解决率=([\d.]+)%,\s*健康度=(\w+)'
    matches = re.findall(pattern, build_output)

    for match in matches:
        metrics["projects"].append({
            "name": match[0],
            "open_di": float(match[1]),
            "sla_timeout": int(match[2]),
            "block": int(match[3]),
            "solve_rate": float(match[4]),
            "health": match[5]
        })

    staff_match = re.search(r'人员数:\s*(\d+)', build_output)
    if staff_match:
        metrics["staff_count"] = int(staff_match.group(1))

    trend_match = re.search(r'趋势天数:\s*(\d+)', build_output)
    if trend_match:
        metrics["trend_days"] = int(trend_match.group(1))

    # 从HTML中提取昨日趋势数据（倒数第2个元素 = 昨天）
    try:
        html_path = os.path.join(PROJECT_DIR, "index.html")
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # 提取 CUMULATIVE_DI 数组（OPEN DI累计值）
        cum_match = re.search(r'const CUMULATIVE_DI = \[(.*?)\];', html_content)
        if cum_match:
            cum_values = [float(v.strip()) for v in cum_match.group(1).split(',') if v.strip()]
            if len(cum_values) >= 2:
                metrics["yesterday_open_di"] = str(cum_values[-2])

        # 提取 DAILY_RESOLVED 数组（每日已解决DI）
        resolved_match = re.search(r'const DAILY_RESOLVED = \[(.*?)\];', html_content)
        if resolved_match:
            resolved_values = [float(v.strip()) for v in resolved_match.group(1).split(',') if v.strip()]
            if len(resolved_values) >= 2:
                metrics["yesterday_solved_di"] = str(resolved_values[-2])
    except Exception as e:
        print(f"⚠️ 读取趋势数据失败：{e}")

    return metrics


def upload_html_to_drive():
    """上传 HTML 到钉盘，返回分享链接"""
    today = datetime.now().strftime("%Y%m%d")
    file_name = f"项目质量看板_{today}.html"

    try:
        result = subprocess.run(
            ["dws", "drive", "upload", "--file", HTML_FILE, "--file-name", file_name, "--format", "json"],
            capture_output=True,
            text=True,
            timeout=30
        )
        if result.returncode == 0 and result.stdout.strip():
            data = json.loads(result.stdout)
            if data.get("success") and data.get("result", {}).get("docUrl"):
                print(f"✅ 钉盘上传成功：{data['result']['docUrl']}")
                return data["result"]["docUrl"]
        print(f"⚠️ 钉盘上传返回异常：{result.stdout[:200]}")
    except Exception as e:
        print(f"⚠️ 钉盘上传失败：{e}")

    return None


def format_dingtalk_message(metrics, dashboard_url=None):
    """格式化钉钉消息 - 仅汇报项目总览合计的昨日OPEN DI和昨日已解决DI"""
    if not metrics["success"]:
        return {
            "msgtype": "text",
            "text": {"content": f"{KEYWORD} - 数据构建失败，请检查日志"}
        }

    yesterday_open_di = metrics.get("yesterday_open_di", "-")
    yesterday_solved_di = metrics.get("yesterday_solved_di", "-")

    lines = []
    lines.append(f"## 📊 项目质量日报")
    lines.append("")
    lines.append(f"**日期：** {metrics['build_time'][:10]}")
    lines.append("")
    lines.append(f"| 指标 | 数值 |")
    lines.append(f"|------|------|")
    lines.append(f"| 昨日OPEN DI | **{yesterday_open_di}** |")
    lines.append(f"| 昨日已解决DI | **{yesterday_solved_di}** |")
    lines.append("")

    if dashboard_url:
        lines.append(f"[查看完整看板]({dashboard_url})")

    content = "\n".join(lines)

    return {
        "msgtype": "markdown",
        "markdown": {
            "title": KEYWORD,
            "text": content
        }
    }


def send_to_dingtalk(message):
    """发送消息到钉钉（修复中文编码）"""
    data = json.dumps(message, ensure_ascii=False).encode('utf-8')

    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(
        WEBHOOK_URL,
        data=data,
        headers={'Content-Type': 'application/json; charset=utf-8'}
    )

    try:
        with urllib.request.urlopen(req, timeout=10, context=ssl_ctx) as response:
            result = json.loads(response.read().decode('utf-8'))
            if result.get('errcode') == 0:
                print("✅ 钉钉推送成功")
                return True
            else:
                print(f"❌ 钉钉推送失败：{result.get('errmsg')}")
                return False
    except Exception as e:
        print(f"❌ 发送失败：{str(e)}")
        return False


def main():
    print(f"🔄 开始构建项目质量看板...")

    # 1. 运行构建（从本地文档获取最新数据）
    build_output = run_build()
    print(build_output)

    # 2. 解析指标
    metrics = parse_metrics(build_output)

    # 3. 上传钉盘获取分享链接
    dashboard_url = upload_html_to_drive()

    # 4. 格式化消息（含链接）
    message = format_dingtalk_message(metrics, dashboard_url)

    # 5. 发送到钉钉
    send_to_dingtalk(message)


if __name__ == "__main__":
    main()

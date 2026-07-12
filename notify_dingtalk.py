#!/usr/bin/env python3
"""云端钉钉通知脚本 - 发送构建结果（含看板链接）到钉钉群"""
import sys, json, ssl, urllib.request
from datetime import datetime

WEBHOOK_URL = "https://oapi.dingtalk.com/robot/send?access_token=bfce5da598f7d476bb8f8307c3da4f4ff35f0c9d3066bda13abeaa8de45bb1a5"

def send_notification(page_url):
    today = datetime.now().strftime("%Y-%m-%d")

    content = f"## 📊 项目质量日报\n\n**日期：** {today}\n\n[查看完整看板]({page_url})"

    message = {
        "msgtype": "markdown",
        "markdown": {
            "title": "项目日报",
            "text": content
        }
    }

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
        with urllib.request.urlopen(req, timeout=10, context=ssl_ctx) as resp:
            result = json.loads(resp.read().decode('utf-8'))
            if result.get('errcode') == 0:
                print("✅ 钉钉通知发送成功")
            else:
                print(f"❌ 钉钉通知失败：{result.get('errmsg')}")
    except Exception as e:
        print(f"❌ 发送失败：{e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        send_notification(sys.argv[1])
    else:
        print("用法: python notify_dingtalk.py <page_url>")

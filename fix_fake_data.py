#!/usr/bin/env python3
"""从损坏的 index.html 中恢复并正确注入假项目数据"""
import json
import random
from datetime import datetime, timedelta

PERSONS = [
    "agatha.huang", "ajax", "alex.xu", "ben.yang", "benlly.li", "darren.su",
    "dylan", "eric.yang", "evie.gao", "gary.dong", "geralt.fu", "hank.chen",
    "ivan.he", "jackit.jiang", "jacky.zhang", "jeff.weng", "jeffrey",
    "jhon.wang", "Kane", "kevin.wu", "leon.wu", "lila.li", "lu.qi",
    "max.liang", "owen", "pan", "ping", "rainer.feng", "Runto.Zhang",
    "seven.zhou", "simon.zhang", "tina.huang", "ting.mo", "tiny",
    "victor.qi", "Wen", "Wing.Law", "King", "zero"
]

BUG_TITLES_APP = [
    "[IOS][dev 3.3.3][E1W]首页加载时偶现白屏超过3秒",
    "[Android][dev 3.3.3][E1W]下拉刷新后列表项顺序错乱",
    "[IOS][dev 3.3.3][E1W]深色模式下部分文字颜色未适配",
    "[Android][dev 3.3.3][E1W]弱网环境下请求超时未显示重试按钮",
    "[IOS][dev 3.3.3][E1W]Tab切换动画卡顿，帧率低于30fps",
    "[Android][dev 3.3.3][E1W]内存泄漏导致长时间使用后崩溃",
    "[IOS][dev 3.3.3][E1W]推送通知点击后无法跳转到对应页面",
    "[Android][dev 3.3.3][E1W]输入法弹出时底部按钮被遮挡",
    "[IOS][dev 3.3.3][E1W]相册选择图片后预览图旋转角度错误",
    "[Android][dev 3.3.3][E1W]定位权限拒绝后未给出引导提示",
    "[IOS][dev 3.3.3][E1W]分享功能在iPad上布局异常",
    "[Android][dev 3.3.3][E1W]后台运行时音频播放被系统暂停",
    "[IOS][dev 3.3.3][E1W]手势返回与页面内滑动冲突",
    "[Android][dev 3.3.3][E1W]多语言切换后部分文案未更新",
    "[IOS][dev 3.3.3][E1W]首次启动引导页跳过按钮点击无效",
]

BUG_TITLES_WEB = [
    "[Web][prod]首页Banner轮播图在Safari下不自动播放",
    "[Web][prod]表单提交后成功提示闪现即消失",
    "[Web][staging]响应式布局在1440px宽度下出现横向滚动条",
    "[Web][prod]搜索结果分页组件页码跳转失效",
    "[Web][prod]文件上传进度条卡在99%不动",
    "[Web][staging]Chrome扩展插件拦截了API请求",
    "[Web][prod]富文本编辑器粘贴Word内容格式丢失",
    "[Web][prod]多标签页切换后WebSocket连接断开",
    "[Web][staging]CSS变量在IE11下未降级处理",
    "[Web][prod]导出Excel文件名包含特殊字符时下载失败",
    "[Web][prod]地图组件缩放至最大级别后拖拽卡顿",
    "[Web][staging]SSR首屏 hydration mismatch 警告",
    "[Web][prod]暗黑模式切换后图标颜色未同步",
    "[Web][prod]视频播放器全屏退出后页面滚动位置丢失",
    "[Web][staging]GraphQL查询嵌套过深导致超时",
]

BUG_TITLES_API = [
    "[API][v2]用户登录接口并发请求时返回500错误",
    "[API][v2]订单创建事务未回滚导致脏数据",
    "[API][v1]分页接口offset参数未校验负数",
    "[API][v2]文件上传接口未限制单文件大小",
    "[API][v2]缓存击穿导致数据库瞬时压力过大",
    "[API][v1]旧版接口未设置CORS头",
    "[API][v2]消息队列消费者重复消费同一条消息",
    "[API][v2]定时任务执行时间漂移超过5分钟",
    "[API][v1]鉴权Token过期后未返回401而是500",
    "[API][v2]批量导入接口单次超过1000条时超时",
    "[API][v2]Redis连接池耗尽后未优雅降级",
    "[API][v1]日志脱敏规则遗漏手机号中间四位",
    "[API][v2]灰度发布流量分配不均",
    "[API][v2]数据库慢查询未触发告警阈值",
    "[API][v1]接口文档Swagger与实际实现不一致",
]


def make_bug(pid, idx, title, level, status, assignee, days_offset):
    weight = {"P0": 10, "P1": 3, "P2": 1, "P3": 0.1}[level]
    open_di = 0.0 if status in ("已关闭", "已解决") else weight
    hour = random.randint(8, 22)
    minute = random.randint(0, 59)
    second = random.randint(0, 59)
    created = datetime(2026, 6, 18) + timedelta(days=days_offset, hours=hour, minutes=minute, seconds=second)
    return {
        "id": f"{pid}-{200+idx}",
        "title": title,
        "level": level,
        "status": status,
        "assignee": assignee,
        "createdDate": created.strftime("%Y-%m-%d %H:%M:%S"),
        "openDI": round(open_di, 1),
        "totalDI": round(weight, 1)
    }


def build_project(name, titles, target_di):
    bugs = []
    remaining_di = target_di
    idx = 0
    while remaining_di >= 10 and idx < len(titles):
        bugs.append(make_bug(name[:3].upper(), idx, titles[idx % len(titles)], "P0", "待处理", random.choice(PERSONS), random.randint(0, 14)))
        remaining_di -= 10; idx += 1
    while remaining_di >= 3 and idx < len(titles):
        bugs.append(make_bug(name[:3].upper(), idx, titles[idx % len(titles)], "P1", "待处理", random.choice(PERSONS), random.randint(0, 14)))
        remaining_di -= 3; idx += 1
    while remaining_di >= 1 and idx < len(titles):
        bugs.append(make_bug(name[:3].upper(), idx, titles[idx % len(titles)], "P2", "待处理", random.choice(PERSONS), random.randint(0, 14)))
        remaining_di -= 1; idx += 1
    p3_count = round(remaining_di / 0.1)
    for _ in range(p3_count):
        if idx < len(titles):
            bugs.append(make_bug(name[:3].upper(), idx, titles[idx % len(titles)], "P3", "待处理", random.choice(PERSONS), random.randint(0, 14)))
            idx += 1
    closed_statuses = ["已关闭", "已解决"]
    while len(bugs) < 25:
        level = random.choices(["P0", "P1", "P2", "P3"], weights=[0.2, 0.3, 0.3, 0.2])[0]
        status = random.choice(closed_statuses)
        bugs.append(make_bug(name[:3].upper(), idx, titles[idx % len(titles)], level, status, random.choice(PERSONS), random.randint(0, 14)))
        idx += 1
    actual_di = sum(b["openDI"] for b in bugs)
    print(f"  {name}: {len(bugs)}条Bug, OPEN DI={actual_di:.1f} (目标={target_di})")
    return bugs


def main():
    path = '/Users/herbert/Desktop/Project_Report_v2/index.html'
    with open(path, 'r', encoding='utf-8') as f:
        html = f.read()

    # 找到真实项目结束标记：E1W-132 那条 Bug 之后的 }]};
    real_end_marker = '"E1W-132"'
    idx = html.index(real_end_marker)
    # 从 E1W-132 往后找 }]};\nconst MILESTONES
    after_132 = html[idx:]
    end_pos = after_132.index('}];\nconst MILESTONES')
    real_projects_end = idx + end_pos + 3  # 包括 }];

    # 提取真实项目完整部分
    projects_start = html.index('const PROJECTS = [')
    real_part = html[projects_start:real_projects_end]  # 包含 const PROJECTS = [...]}];

    # 生成假项目
    fake_projects = []
    for name, titles, target_di in [
        ("E1W_App", BUG_TITLES_APP, 41.1),
        ("E2W_Web", BUG_TITLES_WEB, 13.5),
        ("E3W_API", BUG_TITLES_API, 39.2),
    ]:
        bugs = build_project(name, titles, target_di)
        fake_projects.append({"id": name, "name": name, "bugs": bugs})

    # 合并：将 }; 改为 }, 然后追加假项目（去掉首尾 []）
    combined = real_part[:-2] + ', ' + json.dumps(fake_projects, ensure_ascii=False)[1:-1] + '];'

    # 重建 HTML
    before = html[:projects_start]
    after = html[real_projects_end:]
    new_html = before + combined + after

    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_html)

    print("\n✅ 假项目数据已正确注入到 PROJECTS 数组末尾")


if __name__ == '__main__':
    main()

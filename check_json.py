#!/usr/bin/env python3
"""检查PROJECTS_DATA的JSON结构"""
import json, re

with open('index.html', 'r') as f:
    html = f.read()

# 提取PROJECTS_DATA的完整JSON
m = re.search(r'const PROJECTS_DATA = ({.*?});\nconst TREND_DATES', html, re.DOTALL)
if m:
    try:
        data = json.loads(m.group(1))
        print('✅ PROJECTS_DATA JSON解析成功')
        print(f'项目数: {len(data)}')
        for proj_name, proj_data in data.items():
            if isinstance(proj_data, dict):
                bugs_count = len(proj_data.get('bugs', []))
                stats = proj_data.get('stats', {})
                print(f'  {proj_name}: {bugs_count}条bug')
                print(f'    stats keys: {list(stats.keys())}')
                print(f'    open_di={stats.get("open_di")}, sla={stats.get("sla")}, health={stats.get("health")}')
            else:
                print(f'  ❌ {proj_name} 不是对象，类型={type(proj_data).__name__}')
                print(f'     值片段: {str(proj_data)[:100]}')
    except json.JSONDecodeError as e:
        print(f'❌ JSON解析失败: {e}')
        error_pos = e.pos
        snippet_start = max(0, error_pos - 50)
        snippet_end = min(len(m.group(1)), error_pos + 50)
        print(f'错误位置附近: ...{repr(m.group(1)[snippet_start:snippet_end])}...')
else:
    print('❌ 未找到PROJECTS_DATA')

#!/usr/bin/env python3
"""诊断index.html空白问题"""
import re

with open('index.html', 'r', encoding='utf-8') as f:
    html = f.read()

print("=== HTML结构检查 ===")
checks = [
    ('<!DOCTYPE html>', 'DOCTYPE声明'),
    ('<html', 'html开标签'),
    ('</html>', 'html闭标签'),
    ('<body', 'body开标签'),
    ('</body>', 'body闭标签'),
    ('id="navMenu"', 'navMenu元素'),
    ('id="mainContent"', 'mainContent元素'),
    ('echarts.min.js', 'ECharts CDN'),
]

for pattern, desc in checks:
    status = '✅' if pattern in html else '❌'
    print(f"{status} {desc}")

# 检查script标签数量
script_open = len(re.findall(r'<script(?!>)', html))
script_close = len(re.findall(r'</script>', html))
print(f"\nScript标签: {script_open}开 / {script_close}闭")

# 提取JS内容
m = re.search(r'<script>([\s\S]*)</script>', html)
if not m:
    print("❌ 未找到script内容")
else:
    js = m.group(1)
    print(f"✅ JS长度: {len(js)} 字符")
    
    # 检查关键函数是否存在
    funcs = ['renderNav', 'renderContent', 'switchPage', 'initCharts']
    for func in funcs:
        if f'function {func}' in js:
            print(f"✅ {func}() 存在")
        else:
            print(f"❌ {func}() 缺失")
    
    # 检查数据注入点是否被替换
    placeholders = ['__PROJECTS_JSON__', '__TREND_DATES_JSON__', '__CUMULATIVE_DI_JSON__']
    has_unreplaced = any(ph in js for ph in placeholders)
    if has_unreplaced:
        print("❌ 发现未替换的数据占位符！")
        for ph in placeholders:
            if ph in js:
                print(f"   - {ph} 仍存在")
    else:
        print("✅ 所有数据占位符已替换")
    
    # 检查JSON数据是否有效
    import json
    projects_match = re.search(r'const PROJECTS_DATA = ({.*?});\n', js, re.DOTALL)
    if projects_match:
        try:
            data = json.loads(projects_match.group(1))
            print(f"✅ PROJECTS_DATA JSON有效，包含 {len(data)} 个项目")
            for proj_name, proj_data in data.items():
                bugs_count = len(proj_data.get('bugs', []))
                stats = proj_data.get('stats', {})
                print(f"   - {proj_name}: {bugs_count}条bug, OPEN DI={stats.get('open_di')}")
        except json.JSONDecodeError as e:
            print(f"❌ PROJECTS_DATA JSON解析失败: {e}")
            # 显示前200字符帮助定位
            snippet = projects_match.group(1)[:200]
            print(f"   数据片段: {snippet}...")
    else:
        print("⚠️  未找到PROJECTS_DATA定义")

# 检查HTML末尾是否完整
print(f"\n=== 文件完整性 ===")
print(f"文件大小: {len(html)} 字节")
print(f"最后100字符: {repr(html[-100:])}")
if html.strip().endswith('</html>'):
    print("✅ HTML正常闭合")
else:
    print("❌ HTML可能未完整写入")

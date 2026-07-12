#!/usr/bin/env python3
"""
清理里程碑映射：由于假项目已删除，B30/X50 映射失效
将 MILESTONES_MAPPED 和 PROJECT_RISKS_MAPPED 设为空对象
"""
import re

HTML_PATH = '/Users/herbert/Desktop/Project_Report_v2/index.html'

def main():
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()
    
    # 将 MILESTONES_MAPPED 和 PROJECT_RISKS_MAPPED 初始化为空对象
    # 替换原来的 Object.entries 循环逻辑
    old_mapping = r'''const MILESTONES_MAPPED = \{\};
const PROJECT_RISKS_MAPPED = \{\};
Object\.entries\(MILESTONES\)\.forEach\(\(\[sheetName, items\]\) => \{
  const projId = MILESTONE_SHEET_MAP\[sheetName\] \|\| sheetName;
  MILESTONES_MAPPED\[projId\] = items;
\}\);
Object\.entries\(PROJECT_RISKS\)\.forEach\(\(\[sheetName, risks\]\) => \{
  const projId = MILESTONE_SHEET_MAP\[sheetName\] \|\| sheetName;
  PROJECT_RISKS_MAPPED\[projId\] = risks;
\}\);'''
    
    new_mapping = '''// 里程碑映射已清空（原B30/X50对应假项目已删除）
// 如需恢复，请更新 MILESTONE_SHEET_MAP 并重新运行 generate.py
const MILESTONES_MAPPED = {};
const PROJECT_RISKS_MAPPED = {};'''
    
    new_html = re.sub(old_mapping, new_mapping, html, flags=re.DOTALL)
    
    if new_html == html:
        print("️  正则匹配失败，尝试行号定位")
        lines = html.split('\n')
        start_idx = None
        end_idx = None
        for i, line in enumerate(lines):
            if 'const MILESTONES_MAPPED = {};' in line:
                start_idx = i
            if start_idx and 'PROJECT_RISKS_MAPPED[projId] = risks;' in line:
                end_idx = i
                break
        
        if start_idx and end_idx:
            lines[start_idx:end_idx+1] = new_mapping.split('\n')
            new_html = '\n'.join(lines)
            print(f"✓ 已替换映射代码 (行 {start_idx+1}-{end_idx+1})")
        else:
            print("✗ 无法定位映射代码")
            return
    
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(new_html)
    
    print("\n✓ 里程碑映射已清空")
    print("  - MILESTONES_MAPPED = {}")
    print("  - PROJECT_RISKS_MAPPED = {}")
    print("  - 里程碑时间线将显示'暂无里程碑数据'")
    print("  - 风险项分析将显示'当前项目暂无里程碑风险项'")

if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""
修复里程碑映射顺序和使用位置
"""
import re

HTML_PATH = '/Users/herbert/Desktop/Project_Report_v2/index.html'

def main():
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 找到关键行号
    mapping_start = None  # 第120行注释开始
    mapping_end = None    # 第130行 }); 结束
    milestones_def = None # 第132行 const MILESTONES
    risks_def = None      # 第133行 const PROJECT_RISKS
    
    for i, line in enumerate(lines):
        if '// 根据映射关系重建 MILESTONES' in line:
            mapping_start = i
        if mapping_start and mapping_end is None and line.strip() == '});':
            mapping_end = i
        if 'const MILESTONES = {' in line:
            milestones_def = i
        if 'const PROJECT_RISKS = {' in line:
            risks_def = i
    
    print(f"映射代码: 行 {mapping_start+1}-{mapping_end+1}")
    print(f"MILESTONES定义: 行 {milestones_def+1}")
    print(f"PROJECT_RISKS定义: 行 {risks_def+1}")
    
    if not all([mapping_start, mapping_end, milestones_def, risks_def]):
        print("✗ 未找到所有必要行")
        return
    
    # 提取映射代码块（包括前面的空行和注释）
    # 从第119行（空行）到第130行（});）
    block_start = mapping_start - 1  # 包含前面的空行
    mapping_block = lines[block_start:mapping_end+1]
    
    # 删除原位置的映射代码
    del lines[block_start:mapping_end+1]
    
    # 重新计算 MILESTONES 和 PROJECT_RISKS 的行号（因为删除了前面的行）
    new_milestones_idx = None
    new_risks_idx = None
    for i, line in enumerate(lines):
        if 'const MILESTONES = {' in line:
            new_milestones_idx = i
        if 'const PROJECT_RISKS = {' in line:
            new_risks_idx = i
    
    # 在 PROJECT_RISKS 定义后插入映射代码
    insert_pos = new_risks_idx + 1
    lines[insert_pos:insert_pos] = ['\n'] + mapping_block
    
    # 替换使用处：PROJECT_RISKS[selectedProject] -> PROJECT_RISKS_MAPPED[selectedProject]
    html = ''.join(lines)
    html = html.replace('PROJECT_RISKS[selectedProject]', 'PROJECT_RISKS_MAPPED[selectedProject]')
    html = html.replace('Object.keys(MILESTONES)', 'Object.keys(MILESTONES_MAPPED)')
    html = html.replace('MILESTONES[projId]', 'MILESTONES_MAPPED[projId]')
    
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print("\n✓ 映射顺序修复完成")
    print("  - 映射代码已移至 MILESTONES/PROJECT_RISKS 定义之后")
    print("  - 使用处已改为 MILESTONES_MAPPED/PROJECT_RISKS_MAPPED")

if __name__ == '__main__':
    main()

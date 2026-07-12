#!/usr/bin/env python3
"""
精确替换里程碑映射代码为清空状态
"""

HTML_PATH = '/Users/herbert/Desktop/Project_Report_v2/index.html'

def main():
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 找到要替换的行范围：从 "const MILESTONE_SHEET_MAP" 到 "// 生成人员统计数据" 之前
    start_idx = None
    end_idx = None
    
    for i, line in enumerate(lines):
        if 'const MILESTONE_SHEET_MAP = {' in line:
            start_idx = i
        if start_idx and '// 生成人员统计数据' in line:
            end_idx = i
            break
    
    if start_idx is None or end_idx is None:
        print(f"✗ 未找到目标代码块 (start={start_idx}, end={end_idx})")
        return
    
    print(f"找到代码块: 行 {start_idx+1} - {end_idx}")
    
    # 新代码
    new_code = '''// 里程碑映射已清空（原B30/X50对应假项目E1W_App/E2W_Web已删除）
// 如需恢复，请更新 MILESTONE_SHEET_MAP 并重新运行 generate.py
const MILESTONES_MAPPED = {};
const PROJECT_RISKS_MAPPED = {};

'''
    
    # 替换
    lines[start_idx:end_idx] = [new_code]
    
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.writelines(lines)
    
    print("✓ 里程碑映射代码已替换为空对象")
    print(f"  删除了 {end_idx - start_idx} 行旧代码")
    print(f"  新增了 {len(new_code.splitlines())} 行注释+声明")

if __name__ == '__main__':
    main()

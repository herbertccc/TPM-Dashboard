#!/usr/bin/env python3
"""
清理 index.html 中的假数据并修复数据源问题
"""
import re, json

HTML_PATH = '/Users/herbert/Desktop/Project_Report_v2/index.html'

def fix_generate_trend_data(html):
    """将 generateTrendData 从随机模拟改为基于真实 Bug 数据统计"""
    old_func = r'''function generateTrendData\(projectId\) \{
  const days = 7;
  const diValues = \[\], bugCounts = \[\], labels = \[\];
  for \(let i = days - 1; i >= 0; i--\) \{
    const d = new Date\(\); d\.setDate\(d\.getDate\(\) - i\);
    labels\.push\(`\$\{d\.getMonth\(\)\+1\}/\$\{d\.getDate\(\)\}`\);
    const baseDI = projectId === 'E1W' \? 87\.5 : projectId === 'E2W' \? 42\.3 : 156\.2;
    const noise = \(Math\.random\(\) - 0\.5\) \* 20;
    diValues\.push\(Math\.max\(0, baseDI \+ noise \* \(i / days\)\)\);
    const baseBugs = projectId === 'E1W' \? 8 : projectId === 'E2W' \? 4 : 10;
    bugCounts\.push\(Math\.max\(2, baseBugs - \(days - i\) \* 0\.5 \+ Math\.random\(\) \* 2\)\);
  \}
  return \{ labels, diValues, bugCounts \};
\}'''
    
    new_func = '''function generateTrendData(projectId) {
  const project = PROJECTS.find(p => p.id === projectId);
  if (!project || !project.bugs.length) {
    return { labels: [], diValues: [], bugCounts: [] };
  }
  
  const days = 7;
  const labels = [], diValues = [], bugCounts = [];
  
  // 按日期聚合真实 Bug 数据
  const dailyOpenDI = {};
  const dailyBugCount = {};
  project.bugs.forEach(b => {
    const dateStr = b.createdDate.split(' ')[0].replace(/\\//g, '-');
    if (!dailyOpenDI[dateStr]) dailyOpenDI[dateStr] = 0;
    if (!dailyBugCount[dateStr]) dailyBugCount[dateStr] = 0;
    dailyOpenDI[dateStr] += b.openDI || 0;
    dailyBugCount[dateStr]++;
  });
  
  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(); d.setDate(d.getDate() - i);
    const dateKey = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    labels.push(`${d.getMonth()+1}/${d.getDate()}`);
    diValues.push(dailyOpenDI[dateKey] || 0);
    bugCounts.push(dailyBugCount[dateKey] || 0);
  }
  
  return { labels, diValues, bugCounts };
}'''
    
    # 尝试直接替换
    new_html = re.sub(old_func, new_func, html, flags=re.DOTALL)
    if new_html == html:
        print("⚠️  generateTrendData 正则匹配失败，尝试行号定位替换")
        # 备用方案：通过行号定位
        lines = html.split('\n')
        start_idx = None
        end_idx = None
        for i, line in enumerate(lines):
            if 'function generateTrendData(projectId)' in line:
                start_idx = i
            if start_idx is not None and line.strip() == '}' and i > start_idx + 5:
                end_idx = i
                break
        
        if start_idx is not None and end_idx is not None:
            lines[start_idx:end_idx+1] = new_func.split('\n')
            new_html = '\n'.join(lines)
            print(f"✓ 已替换 generateTrendData (行 {start_idx+1}-{end_idx+1})")
        else:
            print("✗ 无法定位 generateTrendData 函数")
            return html
    
    return new_html

def add_milestone_mapping(html):
    """添加里程碑 sheet 名与 PROJECTS id 的映射关系"""
    # 在 PROJECTS 变量后添加映射配置
    mapping_code = '''
// 里程碑 sheet 名与项目 ID 的映射关系
const MILESTONE_SHEET_MAP = {
  "B30": "E1W_App",
  "X50": "E2W_Web"
};

// 根据映射关系重建 MILESTONES 和 PROJECT_RISKS，使 key 与 PROJECTS id 一致
const MILESTONES_MAPPED = {};
const PROJECT_RISKS_MAPPED = {};
Object.entries(MILESTONES).forEach(([sheetName, items]) => {
  const projId = MILESTONE_SHEET_MAP[sheetName] || sheetName;
  MILESTONES_MAPPED[projId] = items;
});
Object.entries(PROJECT_RISKS).forEach(([sheetName, risks]) => {
  const projId = MILESTONE_SHEET_MAP[sheetName] || sheetName;
  PROJECT_RISKS_MAPPED[projId] = risks;
});
'''
    
    # 在 PROJECTS 变量定义后插入映射代码
    pattern = r'(const PROJECTS = \[.*?\];)'
    match = re.search(pattern, html, flags=re.DOTALL)
    if match:
        insert_pos = match.end()
        new_html = html[:insert_pos] + mapping_code + html[insert_pos:]
        print("✓ 已添加里程碑映射关系")
        return new_html
    else:
        print("⚠️  未找到 PROJECTS 变量，跳过映射添加")
        return html

def main():
    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()
    
    print("开始清理假数据...")
    
    # 1. 修复 generateTrendData
    html = fix_generate_trend_data(html)
    
    # 2. 添加里程碑映射
    html = add_milestone_mapping(html)
    
    # 写回
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print("\n✓ 假数据清理完成")
    print(f"  文件: {HTML_PATH}")

if __name__ == '__main__':
    main()

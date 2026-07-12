#!/usr/bin/env python3
"""批量修改 index.html 实现7项UI/UX优化需求"""
import re

path = '/Users/herbert/Desktop/Project_Report_v2/index.html'
with open(path, 'r', encoding='utf-8') as f:
    html = f.read()

# ===== 需求1: 子项目DI趋势标题改为"项目名-DI趋势" =====
# 原: ${p.name} - 每日新增DI / 解决DI / 累计DI趋势（近30天）
# 改: ${p.name}-DI趋势
html = html.replace(
    '${p.name} - 每日新增DI / 解决DI / 累计DI趋势（近30天）',
    '${p.name}-DI趋势'
)

# ===== 需求2&3: 待处理/已解决分组表格列宽对齐 + 不换行 =====
# 找到两个表格的 thead 部分，统一列宽并加 white-space:nowrap
# 需要给两个表格的 th 都加上 style="min-width:Xpx;white-space:nowrap"

# 先定位待处理问题表格的表头
old_pending_thead = '''<thead><tr><th>ID</th><th>标题</th><th>等级</th><th>状态</th><th>执行人</th><th>创建时间</th><th>SLA状态</th></tr></thead>'''
new_pending_thead = '''<thead><tr><th style="min-width:60px;white-space:nowrap">ID</th><th style="min-width:200px;white-space:nowrap">标题</th><th style="min-width:45px;white-space:nowrap">等级</th><th style="min-width:60px;white-space:nowrap">状态</th><th style="min-width:60px;white-space:nowrap">执行人</th><th style="min-width:140px;white-space:nowrap">创建时间</th><th style="min-width:70px;white-space:nowrap">SLA状态</th></tr></thead>'''
html = html.replace(old_pending_thead, new_pending_thead)

# 已解决/已拒绝表格表头
old_resolved_thead = '''<thead><tr><th>ID</th><th>标题</th><th>等级</th><th>状态</th><th>执行人</th><th>创建时间</th><th>SLA状态</th></tr></thead>'''
new_resolved_thead = '''<thead><tr><th style="min-width:60px;white-space:nowrap">ID</th><th style="min-width:200px;white-space:nowrap">标题</th><th style="min-width:45px;white-space:nowrap">等级</th><th style="min-width:60px;white-space:nowrap">状态</th><th style="min-width:60px;white-space:nowrap">执行人</th><th style="min-width:140px;white-space:nowrap">创建时间</th><th style="min-width:70px;white-space:nowrap">SLA状态</th></tr></thead>'''
html = html.replace(old_resolved_thead, new_resolved_thead)

# ===== 需求4: 子项目页面新增人员卡片（解决率、SLA超时数、DI值）=====
# 在 renderProjectView 函数中，找到 DI趋势图之后插入人员卡片
# 查找位置：projectTripleChart 之后的 </div> 后面

person_card_html = '''
      <div class="chart-section fade-in"><div class="chart-title">👥 项目人员效能概览</div>
        <div id="projectPersonCards" class="person-cards-grid"></div>
      </div>'''

# 在 projectTripleChart 容器后插入
html = html.replace(
    '<div class="chart-container" id="projectTripleChart"></div></div>',
    '<div class="chart-container" id="projectTripleChart"></div></div>' + person_card_html
)

# 在 renderProjectView 末尾添加渲染人员卡片的逻辑
# 找到 drawComboChart('projectTripleChart'...) 这行之后插入
draw_person_cards_code = '''
  // 渲染项目人员效能卡片
  const projBugs = p.bugs || [];
  const personMap = {};
  projBugs.forEach(b => {
    if (!personMap[b.assignee]) personMap[b.assignee] = { name: b.assignee, total: 0, resolved: 0, slaTimeout: 0, openDI: 0 };
    personMap[b.assignee].total++;
    if (b.status === '已解决' || b.status === '已关闭') personMap[b.assignee].resolved++;
    if (isSlaTimeout(b)) personMap[b.assignee].slaTimeout++;
    personMap[b.assignee].openDI += b.openDI || 0;
  });
  const personCards = Object.values(personMap).sort((a, b) => b.openDI - a.openDI);
  const cardsContainer = document.getElementById('projectPersonCards');
  if (cardsContainer) {
    cardsContainer.innerHTML = personCards.map(pc => {
      const rate = pc.total > 0 ? Math.round(pc.resolved / pc.total * 100) : 0;
      return `<div class="person-card-mini">
        <div class="pc-name">${pc.name}</div>
        <div class="pc-stats">
          <span class="pc-stat"><b>${rate}%</b> 解决率</span>
          <span class="pc-stat ${pc.slaTimeout > 0 ? 'pc-warn' : ''}"><b>${pc.slaTimeout}</b> SLA超时</span>
          <span class="pc-stat"><b>${pc.openDI.toFixed(1)}</b> OPEN DI</span>
        </div>
      </div>`;
    }).join('');
  }'''

# 在 drawComboChart 调用后插入
html = html.replace(
    "setTimeout(() => drawComboChart('projectTripleChart', trend.labels, trend.newDI, trend.cumulativeDI, '#165dff', '#f53f3f'), 80);",
    "setTimeout(() => drawComboChart('projectTripleChart', trend.labels, trend.newDI, trend.cumulativeDI, '#165dff', '#f53f3f'), 80);" + draw_person_cards_code
)

# ===== 需求5: 人员DI分布图设置最小宽度 + 字体加大 =====
# 找到 drawPersonDIChart 函数中的 SVG 生成部分
# 将 viewBox 和容器 min-width 调整

# 在 drawPersonDIChart 函数开头增加 min-width 样式
old_draw_person_start = "function drawPersonDIChart(containerId, persons) {\n  const container = document.getElementById(containerId);\n  if (!container) return;\n  const width = container.clientWidth || 800;"
new_draw_person_start = "function drawPersonDIChart(containerId, persons) {\n  const container = document.getElementById(containerId);\n  if (!container) return;\n  container.style.minWidth = '900px';\n  const width = Math.max(container.clientWidth, 900);"
html = html.replace(old_draw_person_start, new_draw_person_start)

# 加大 SVG 中的字体大小（label 从 11→14，人名从 12→14）
html = html.replace("font-size=\"11\" fill=\"#86909c\">${lbl}", "font-size=\"14\" fill=\"#86909c\">${lbl}")
html = html.replace("font-size=\"12\" font-weight=\"600\" fill=\"#1f2329\">${p.name}", "font-size=\"14\" font-weight=\"600\" fill=\"#1f2329\">${p.name}")

# ===== 需求6: 人员DI分布图增加重开次数展示 =====
# 修改 drawPersonDIChart 的数据统计部分，加入 reopened 字段
# 找到 buildPersonStats 或直接在 drawPersonDIChart 中统计

# 在 drawPersonDIChart 中找到统计 resolved/open 的部分，加入 reopened
old_stats = """  const data = persons.map(p => ({
    name: p.name,
    resolvedDI: p.totalDI - p.openDI,
    openDI: p.openDI
  }));"""
new_stats = """  const data = persons.map(p => ({
    name: p.name,
    resolvedDI: p.totalDI - p.openDI,
    openDI: p.openDI,
    reopenedCount: PROJECTS.flatMap(pr => pr.bugs).filter(b => b.assignee === p.name && b.status === '重新打开').length
  }));"""
html = html.replace(old_stats, new_stats)

# 修改 SVG 绘制为三段堆叠：已解决DI(绿) + 未解决DI(橙) + 重开次数(红条在右侧)
# 找到 rect 绘制部分，改为三段
old_rect_draw = """  data.forEach((d, i) => {
    const x = padding.left + xStep * i + (xStep - barWidth) / 2;
    const totalH = ((d.resolvedDI + d.openDI) / maxVal) * chartH;
    const resolvedH = (d.resolvedDI / maxVal) * chartH;
    const y = padding.top + chartH - totalH;
    svg += `<rect x="${x}" y="${y}" width="${barWidth}" height="${resolvedH}" fill="#00b42a" opacity="0.8" rx="2"/>`;
    svg += `<rect x="${x}" y="${y + resolvedH}" width="${barWidth}" height="${totalH - resolvedH}" fill="#ff7d00" opacity="0.8" rx="2"/>`;
  });"""
new_rect_draw = """  data.forEach((d, i) => {
    const x = padding.left + xStep * i + (xStep - barWidth) / 2;
    const totalH = ((d.resolvedDI + d.openDI) / maxVal) * chartH;
    const resolvedH = (d.resolvedDI / maxVal) * chartH;
    const openH = (d.openDI / maxVal) * chartH;
    const y = padding.top + chartH - totalH;
    // 已解决DI（绿色）
    svg += `<rect x="${x}" y="${y}" width="${barWidth}" height="${resolvedH}" fill="#00b42a" opacity="0.85" rx="2"/>`;
    // 未解决DI（橙色）
    svg += `<rect x="${x}" y="${y + resolvedH}" width="${barWidth}" height="${openH}" fill="#ff7d00" opacity="0.85" rx="2"/>`;
    // 重开次数（红色小条，固定在底部上方）
    if (d.reopenedCount > 0) {
      const reopenH = Math.min(d.reopenedCount * 3, 20);
      const reopenY = padding.top + chartH - reopenH;
      svg += `<rect x="${x + barWidth * 0.2}" y="${reopenY}" width="${barWidth * 0.6}" height="${reopenH}" fill="#f53f3f" opacity="0.9" rx="2"/>`;
      svg += `<text x="${x + barWidth / 2}" y="${reopenY - 3}" text-anchor="middle" font-size="10" fill="#f53f3f" font-weight="700">${d.reopenedCount}</text>`;
    }
  });"""
html = html.replace(old_rect_draw, new_rect_draw)

# 更新图例
old_legend = """  svg += `<rect x="${width-200}" y="10" width="12" height="12" fill="#00b42a" opacity="0.85" rx="2"/>`;
  svg += `<text x="${width-184}" y="21" font-size="12" fill="#333">已解决DI</text>`;
  svg += `<rect x="${width-110}" y="10" width="12" height="12" fill="#ff7d00" opacity="0.85" rx="2"/>`;
  svg += `<text x="${width-94}" y="21" font-size="12" fill="#333">未解决DI</text>`;"""
new_legend = """  svg += `<rect x="${width-280}" y="10" width="12" height="12" fill="#00b42a" opacity="0.85" rx="2"/>`;
  svg += `<text x="${width-264}" y="21" font-size="13" fill="#333">已解决DI</text>`;
  svg += `<rect x="${width-180}" y="10" width="12" height="12" fill="#ff7d00" opacity="0.85" rx="2"/>`;
  svg += `<text x="${width-164}" y="21" font-size="13" fill="#333">未解决DI</text>`;
  svg += `<rect x="${width-90}" y="10" width="12" height="12" fill="#f53f3f" opacity="0.9" rx="2"/>`;
  svg += `<text x="${width-74}" y="21" font-size="13" fill="#333">重开次数</text>`;"""
html = html.replace(old_legend, new_legend)

# ===== 需求7: 人力资源名单表格修改 =====
# 找到 HR 名单表格的 thead 和 tbody 渲染部分

# 表头修改：新增"重开次数"列；"状态"→"SLA题数"；"BUG数（待处理/已关闭）"→"BUG统计（待处理/总数）"
old_hr_thead = '''<thead><tr><th>姓名</th><th>负责Bug数</th><th>已解决</th><th>状态</th><th>OPEN DI</th><th>解决率</th></tr></thead>'''
new_hr_thead = '''<thead><tr><th style="min-width:80px;white-space:nowrap">姓名</th><th style="min-width:120px;white-space:nowrap">BUG统计（待处理/总数）</th><th style="min-width:80px;white-space:nowrap">SLA题数</th><th style="min-width:80px;white-space:nowrap">重开次数</th><th style="min-width:80px;white-space:nowrap">OPEN DI</th><th style="min-width:70px;white-space:nowrap">解决率</th></tr></thead>'''
html = html.replace(old_hr_thead, new_hr_thead)

# 表格数据行修改
old_hr_row = '''`<tr><td>${p.name}</td><td>${p.totalAssigned}</td><td>${p.resolved}</td><td>${p.slaTimeout}</td><td>${p.openDI.toFixed(1)}</td><td>${p.resolveRate}%</td></tr>`'''
new_hr_row = '''`<tr><td>${p.name}</td><td>${p.totalAssigned - p.resolved}/${p.totalAssigned}</td><td>${p.slaTimeout}</td><td>${p.reopened || 0}</td><td>${p.openDI.toFixed(1)}</td><td>${p.resolveRate}%</td></tr>`'''
html = html.replace(old_hr_row, new_hr_row)

# buildPersonStats 中加入 reopened 统计
old_build = """      map[b.assignee].totalAssigned++;
      if (b.status === '已解决' || b.status === '已关闭') map[b.assignee].resolved++;
      if (isSlaTimeout(b)) map[b.assignee].slaTimeout++;"""
new_build = """      map[b.assignee].totalAssigned++;
      if (b.status === '已解决' || b.status === '已关闭') map[b.assignee].resolved++;
      if (isSlaTimeout(b)) map[b.assignee].slaTimeout++;
      if (b.status === '重新打开') map[b.assignee].reopened = (map[b.assignee].reopened || 0) + 1;"""
html = html.replace(old_build, new_build)

# 初始化 reopened 字段
old_init = "map[b.assignee] = { name: b.assignee, totalAssigned: 0, resolved: 0, slaTimeout: 0, openDI: 0, totalDI: 0 };"
new_init = "map[b.assignee] = { name: b.assignee, totalAssigned: 0, resolved: 0, slaTimeout: 0, openDI: 0, totalDI: 0, reopened: 0 };"
html = html.replace(old_init, new_init)

# ===== 新增 CSS 样式 =====
# 在 <style> 标签内追加人员卡片和表格样式
css_additions = """
/* 项目人员效能卡片 */
.person-cards-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 12px; margin-top: 12px; }
.person-card-mini { background: #fff; border: 1px solid #e5e6eb; border-radius: 8px; padding: 12px 16px; }
.pc-name { font-size: 14px; font-weight: 700; color: #1f2329; margin-bottom: 8px; }
.pc-stats { display: flex; gap: 12px; flex-wrap: wrap; }
.pc-stat { font-size: 12px; color: #4e5969; }
.pc-stat b { color: #165dff; font-size: 14px; }
.pc-warn b { color: #f53f3f !important; }

/* 表格全局不换行 */
table td, table th { white-space: nowrap; }
"""

# 在 </style> 前插入
html = html.replace('</style>', css_additions + '\n</style>')

with open(path, 'w', encoding='utf-8') as f:
    f.write(html)

print("✅ 7项需求修改完成")
print("  1. 子项目DI趋势标题 → 项目名-DI趋势")
print("  2. 待处理/已解决表格列宽对齐（固定min-width）")
print("  3. 列表缩小时不换行（white-space:nowrap + min-width）")
print("  4. 子项目页面新增人员效能卡片（解决率/SLA超时/DI值）")
print("  5. 人员DI分布图最小宽度900px + 字体加大")
print("  6. 人员DI分布图三段展示：已解决DI/未解决DI/重开次数")
print("  7. HR名单：新增重开次数列，状态→SLA题数，BUG数→BUG统计（待处理/总数）")

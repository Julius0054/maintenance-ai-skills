---
name: weekly-report-gen
description: 璀璨云著A地块维修周报生成工具。正确读取台账数据，统计本周完成情况，生成格式正确的周报。避免常见错误：统计数据错误、工种明细漏项、累计完成行未更新、剩余工程量错误、主要剩余工种明细统计错误、下周计划错行、协调事项错行。
---

> **⚠ 生成方式铁律（2026-07-20 用户三次纠正后强制执行）：一律用「v7 模板复制法」，禁止从零生成、禁止用 `insert_rows`/`delete_rows` 改结构。**
> 原因：① 从零生成（`gen_weekly_paradigm.py` 的 layout 重建）会因 openpyxl 属性校验 ≠ WPS 实际渲染，导致**横向合并吞数据/不到边、单元格自动换行失效、行高不足显示不全**三类结构性破坏（用户 2026-07-19 实测否决）；② `insert_rows`/`delete_rows` 在 openpyxl 下会让**合并区间整体错位**（用户 2026-07-20 实测：1/2/7 周报删/插行后标题表头被吞、协调/风险标题合并错位、风险第4行数据丢失）。
> **正确做法 = v7 确定性重建法**：复制一份**已知格式正确的周报**作模板 → 先 `unmerge_cells` 取消全部合并（否则合并非左上角格只读）→ 用 `from copy import copy` 逐格复制 `value/font/fill/border/alignment/number_format/protection`（循环**从下往上且含目标行**，插1行用 `range(72,45,-1)`，误写 `range(72,46,-1)` 会漏掉合计行）→ 只改需改单元格（如插皮阿诺独立行、合计 A 重标、完成率转 `0.00"%"`）→ 按权威合并清单**显式 `merge_cells` 重建**（插入点行号整体 +1，新增行单独加合并）→ 行高从模板 `row_dimensions` 按 +1 还原。绝不动字体/边框/对齐/换行。
> 数据计算仍可参考《周报范式规范》与 `gen_weekly_paradigm.py` 的 `compute_stats()`，但**版面必须由模板复制得到**，生成后人工复核三件事：① 合并是否横向到底（无吞数据）；② 文本单元格 `wrap_text=True`；③ 内容多行处行高足够。
> 核心数据铁律：① 完成率百分比格式 `0.00"%"`（如 95.61%，数值=rate×100）；② 剩余表按完成率升序、**皮阿诺等是普通单位行（非合计行）**，合计行单独置底、D 列列明重点关注单位；③ 仅统计未完工（col12≠"是"）记录；④ 文件名楼栋集合前置（`璀璨云著A地块1、2、7栋维修周报_YYYY年M月D日`）；⑤ 禁止出现"上周对比行"（如旧的 9302 累计行）。

# 璀璨云著A地块维修周报生成 Skill

## 概述
本Skill用于正确生成璀璨云著A地块维修周报，避免常见错误。

## 周报标准结构（2026-05-16修正版）

| 行号 | 内容 | 说明 |
|------|------|------|
| 1 | 璀璨云著A地块（1、2、7栋）维修维保工作周报 | A:J合并 |
| 2 | 报告期：YYYY年MM月DD日（周X）——YYYY年MM月DD日（周X） | A:E报告期, F:J汇报人/日期 |
| 3 | 空行 | |
| 4 | 一、本周数据统计 | |
| 5 | （一）总体概况 | |
| 6 | 表头（楼栋/累计报修/累计完成/剩余未完成/总完成率/本周新增/本周完成） | |
| 7-9 | A1/A2/A7栋数据 | |
| 10 | 合计行 | |
| 11 | 空行 | |
| 12 | （二）本周完成销项明细（按责任单位） | |
| 13 | 表头（责任单位/本周合计/工种明细） | |
| 14起 | 各责任单位销项数据 | 按本周完成量降序排列 |
| 合计行 | 本周累计完成：N | 说明包含"N个责任单位参与本周销项" |
| 若干空行 | 间隔 | |
| 累计完成行 | 累计完成：N | 说明"截止YYYY年MM月DD日，累计完成销项N条（总完成率XX.XX%）" |
| 空行 | | |
| （三）剩余未完成工程量（按责任单位） | | |
| 表头 | 责任单位/完成率/剩余合计（条）/主要剩余工种明细 | |
| 各责任单位数据 | | 按完成率升序排列（低的在前） |
| 合计行 | 合计：N | D列标注重点关注单位 |
| 空行 | | |
| 三、下周工作计划 | | |
| 表头 | 日期/计划工作/具体内容/责任单位 | |
| 下周计划行 | 下周持续/具体日期 | |
| 空行 | | |
| 四、需项目公司协调事项 | | |
| 表头 | 序号/问题类型/问题描述/现场情况 | 协调诉求在"现场情况"列 |
| 协调事项 | | 用换行符分隔问题类型 |
| 空行 | | |
| 五、风险预警 | | |
| 表头 | 序号/风险等级/户号或范围/风险事项/处置建议 | |
| 风险预警行 | 风险等级用⚠高/⚠中/⚠低标识 | |
| 底部日期 | 数据截止日期：YYYY年MM月DD日（周X） | 含"本报告由维保工程部整理提交" |

## 正确流程

### 第一步：读取台账数据
1. 打开台账文件（如`璀璨云著A地块1、2、7栋报修台账2026年M月D日.xlsx`）
2. **必须选择"台账"工作表**（不是默认的活跃工作表）
3. 正确读取列头，确认列的位置：
   - 列1: 项目
   - 列3: 楼栋号（数字1, 2, 7，不是字符串"1栋"等）
   - 列9: 责任单位
   - 列10: 工种
   - 列11: 报修日期
   - 列12: 是否完工
   - 列13: 完成时间
4. 只统计"项目=A地块"的记录

### 第二步：计算统计数据
1. **按楼栋统计**：
   - 累计报修：该楼栋总记录数
   - 累计完成：该楼栋"是否完工=是"的记录数
   - 剩余未完成：累计报修 - 累计完成
   - 总完成率：累计完成 / 累计报修（**用小数格式如0.9578，不是百分比**）
   - 本周新增：报修日期在本周范围内的记录数
   - 本周完成：完成时间在本周范围内的记录数

2. **按责任单位统计本周完成**：
   - 筛选条件：完成时间在本周范围内
   - 按责任单位分组
   - 每个责任单位内，按工种分组统计数量
   - 按完成数量降序排列

3. **按责任单位统计剩余未完成**：
   - 筛选条件：是否完工 != "是"
   - 按责任单位分组
   - 每个责任单位内，按工种分组统计数量
   - **只统计未完工的工种**（不是总条数！）
   - 按完成率（已完成/总数）升序排列

### 第三步：复制一份「已知格式正确」的周报作模板（v7 模板复制法，禁止从零生成）
```python
import shutil
from datetime import datetime

# 优先用同楼栋集合最近一份好周报；跨楼栋参考时用 4/5/10 好周报作样式基准
source = '璀璨云著A地块1、2、7栋维修周报_2026年7月18日.xlsx'   # 须是格式已确认正确的版本
target = f'璀璨云著A地块1、2、7栋维修周报_{datetime.now().strftime("%Y年%m月%d日")}.xlsx'
shutil.copy2(source, target)

wb = load_workbook(target)
ws = wb.active
# ⚠ 关键：保留模板原有的全部合并/行高/换行/字体/填充/边框，绝不重建。
#    只通过 ws.cell(row, col).value = ... 改写「合并区左上角单元格」的值。
#    ⚠ 增/删单位行严禁用 insert_rows/delete_rows（合并会错位）；改用 unmerge 全表→逐格复制→显式 merge_cells 重建（见顶部铁律）。
```

### 第五步：填写数据

#### 5.1 填写报告期和汇报日期
```python
# 报告期（行2列1）
ws.cell(row=2, column=1).value = f"报告期：{week_start.strftime('%Y年%m月%d日')}（周X）——{week_end.strftime('%Y年%m月%d日')}（周X）"

# 汇报日期（行2列6）
ws.cell(row=2, column=6).value = f"汇报人：蔡正延      汇报日期：{datetime.now().strftime('%Y年%m月%d日')}"
```

#### 5.2 填写总体概况（行7-10）
```python
# A1栋（行7）
ws.cell(row=7, column=1).value = 'A1栋'
ws.cell(row=7, column=2).value = stats['A1栋']['total']
ws.cell(row=7, column=3).value = stats['A1栋']['completed']
ws.cell(row=7, column=4).value = stats['A1栋']['total'] - stats['A1栋']['completed']
ws.cell(row=7, column=5).value = stats['A1栋']['completed'] / stats['A1栋']['total']  # 小数格式
ws.cell(row=7, column=6).value = stats['A1栋']['new_this_week']
ws.cell(row=7, column=7).value = stats['A1栋']['completed_this_week']

# A2栋（行8）、A7栋（行9）类似...

# 合计（行10）
ws.cell(row=10, column=1).value = '合计'
ws.cell(row=10, column=2).value = total_total
ws.cell(row=10, column=3).value = total_completed
ws.cell(row=10, column=4).value = total_remaining
ws.cell(row=10, column=5).value = total_rate  # 小数格式
ws.cell(row=10, column=6).value = total_new
ws.cell(row=10, column=7).value = total_completed_week
```

#### 5.3 填写本周完成销项明细（行14起）
```python
from collections import defaultdict

# unit_completed 格式：{'皮阿诺': {'count': 48, 'types': {'户内门': 23, '打胶': 20, ...}}, ...}

row_num = 14
for unit, data in sorted(unit_completed.items(), key=lambda x: x[1]['count'], reverse=True):
    # 责任单位
    ws.cell(row=row_num, column=1).value = unit
    
    # 本周合计
    ws.cell(row=row_num, column=2).value = data['count']
    
    # 工种明细（格式："门窗：21条  |  烟道：1条"）
    types_str = '  |  '.join([
        f"{t}：{c}条"
        for t, c in sorted(data['types'].items(), key=lambda x: x[1], reverse=True)
    ])
    ws.cell(row=row_num, column=3).value = types_str
    
    row_num += 1

# 本周累计完成行
unit_count = len(unit_completed)
ws.cell(row=row_num, column=1).value = '本周累计完成'
ws.cell(row=row_num, column=2).value = total_completed_this_week
ws.cell(row=row_num, column=3).value = f"{unit_count}个责任单位参与本周销项，合计完成{total_completed_this_week}条"
row_num += 2  # 跳过空行

# 累计完成行
ws.cell(row=row_num, column=1).value = '累计完成'
ws.cell(row=row_num, column=2).value = total_completed
ws.cell(row=row_num, column=3).value = f"截止{week_end.strftime('%Y年%m月%d日')}，累计完成销项{total_completed}条（总完成率{total_rate:.2%}）"
```

#### 5.4 填写剩余未完成工程量（行27起）
```python
# unit_remaining 格式：{'二十冶': {'count': 86, 'types': {'门窗玻璃划痕': 32, ...}, 'total': 100, 'completed': 14}, ...}
# 注意：types只统计未完工的工种，不是总条数！

# 按完成率升序排列
sorted_remaining = sorted(
    unit_remaining.items(),
    key=lambda x: x[1]['completed'] / x[1]['total'] if x[1]['total'] > 0 else 0
)

row_num = 29  # 跳过标题和表头
for unit, data in sorted_remaining:
    rate = data['completed'] / data['total'] if data['total'] > 0 else 0
    remaining = data['total'] - data['completed']
    
    if remaining > 0:
        # 责任单位
        ws.cell(row=row_num, column=1).value = unit
        
        # 完成率（小数格式）
        ws.cell(row=row_num, column=2).value = round(rate, 4)
        
        # 剩余合计
        ws.cell(row=row_num, column=3).value = remaining
        
        # 主要剩余工种明细（只统计未完工）
        types_str = '  |  '.join([
            f"{t}：{c}条"
            for t, c in sorted(data['types'].items(), key=lambda x: x[1], reverse=True)
        ])
        ws.cell(row=row_num, column=4).value = types_str
        
        row_num += 1

# 合计行
ws.cell(row=row_num, column=1).value = '合计'
ws.cell(row=row_num, column=2).value = total_rate
ws.cell(row=row_num, column=3).value = total_remaining
ws.cell(row=row_num, column=4).value = f"共{len(unit_remaining)}个责任单位有剩余未完成工程量  ★贝斯特/东海建设完成率偏低需重点关注"
```

#### 5.5 填写下周工作计划（**需要用户提供内容**）
```python
# 必须让用户提供下周工作计划的详细内容
# 不要直接复制上周的内容

next_week_plans = [
    {'date': '下周持续', 'plan': '各单位安排维修', 'detail': '计划下周，三四标段抢工结束后安排处理'},
]

row_num = 48  # 三、下周工作计划表头在行46-47
for plan in next_week_plans:
    ws.cell(row=row_num, column=1).value = plan['date']
    ws.cell(row=row_num, column=2).value = plan['plan']
    ws.cell(row=row_num, column=3).value = plan['detail']
    row_num += 1
```

#### 5.6 填写需项目公司协调事项（**需要用户提供内容**）
```python
# 必须让用户提供协调事项的详细内容
# 表头：序号/问题类型/问题描述/现场情况（协调诉求在H列）

coordination_items = [
    {
        'id': 1,
        'type': '金牌厨柜完成率偏低\n橱柜维修',  # 换行符分隔
        'description': '金牌厨柜剩余43条未完工，完成率87.2%，本周继续安排维修，但进度仍需重点关注',
        'appeal': '请项目公司持续关注金牌厨柜维修进度，协调增加施工力量，确保尽快完成剩余43条报修。维保将持续跟进并每周通报进度。'
    },
    # ... 更多协调事项
]

row_num = 59  # 四、需项目公司协调事项表头在行57-58
for item in coordination_items:
    ws.cell(row=row_num, column=1).value = item['id']
    ws.cell(row=row_num, column=2).value = item['type']
    ws.cell(row=row_num, column=3).value = item['description']
    ws.cell(row=row_num, column=8).value = item['appeal']
    row_num += 1
```

#### 5.7 填写风险预警（**需要用户提供内容**）
```python
# 表头：序号/风险等级/户号或范围/风险事项/处置建议（H列）

risk_warnings = [
    {
        'id': 1,
        'level': '⚠ 高',
        'scope': 'A7-1401',
        'risk': '前期自来水公司现场更换管路导致该户户内泡水，目前该户补偿费用未到位，后期可能造成投诉或业主采取其他途径。',
        'suggestion': '请项目公司尽快落实补偿方案，确保补偿费用及时到位。维保持续跟进业主情绪，防止事态升级。'
    },
    {
        'id': 2,
        'level': '⚠ 中',
        'scope': '全体户/二十冶',
        'risk': '目前玻璃内侧划痕二十冶以非该单位责任为由拒绝处理（剩余32条），后期有可能造成投诉。',
        'suggestion': '请项目公司持续关注二十冶内侧玻璃划痕责任认定，协调明确处理方案，避免后期引发业主投诉。维保将持续跟进现场维修工作。'
    },
]

row_num = 68  # 五、风险预警表头在行66-67
for warning in risk_warnings:
    ws.cell(row=row_num, column=1).value = warning['id']
    ws.cell(row=row_num, column=2).value = warning['level']
    ws.cell(row=row_num, column=3).value = warning['scope']
    ws.cell(row=row_num, column=4).value = warning['risk']
    ws.cell(row=row_num, column=8).value = warning['suggestion']
    row_num += 1
```

#### 5.8 填写数据截止日期
```python
# 底部日期
ws.cell(row=row_num, column=1).value = f"数据截止日期：{week_end.strftime('%Y年%m月%d日')}（{weekday}） | 本报告由维保工程部整理提交"
```

### 第六步：保存文件（合并/行高/换行/字体/填充一律沿用模板，无需重建）
```python
wb.save(target)
wb.close()
```

> **v7 要点回顾**：写值只改合并区左上角单元格；**增删行严禁 `insert_rows`/`delete_rows`**（合并会整体错位，实测吞数据/丢行），改用 unmerge 全表→逐格复制→显式 `merge_cells` 重建；绝不为"同周对比行/残留空行"多此一举。`gen_weekly_paradigm.py` 的 `compute_stats()` 仅用于算数据，版面不得由其从零重建。
> **总结数据放置规（2026-07-20，范式规范 §2.6）**：区块级游离总结行（"本周累计完成：N条"/"累计完成：N条（蓝）"）必须置于所属区块**最后**，且与上方数据**至少间隔 2 空行**（不足插空行，不破坏合并）；表格内"合计"行（总体概况合计、剩余表合计）仍紧接数据末、不留空行。废弃数据行（如 9302 旧累计行）**就地清空为空白行**，严禁 `delete_rows` 删整行。

## 常见错误与解决方案

### 错误1：统计数据错误
**原因**：没有正确读取"台账"工作表，楼栋号字段是数字不是字符串

**解决**：
```python
wb = load_workbook(台账文件)
ws = wb['台账']  # 必须指定"台账"工作表

# 楼栋号是数字
if building == 1 or building == '1':
    key = 'A1栋'
elif building == 2 or building == '2':
    key = 'A2栋'
elif building == 7 or building == '7':
    key = 'A7栋'
```

### 错误2：工种明细漏项严重
**原因**：没有从台账中正确提取"责任单位"（列9）和"工种"（列10），没有按责任单位+工种分组统计

**解决**：
```python
from collections import defaultdict

unit_completed = defaultdict(lambda: {'count': 0, 'types': defaultdict(int)})

for row in range(2, ws.max_row + 1):
    # 判断是否为本周完成
    if week_start <= complete_date <= week_end:
        responsibility = ws.cell(row=row, column=9).value  # 责任单位
        work_type = ws.cell(row=row, column=10).value  # 工种
        
        unit_completed[responsibility]['count'] += 1
        if work_type:
            unit_completed[responsibility]['types'][work_type] += 1
```

### 错误3：累计完成行未更新
**原因**：直接复制了上周周报的模板，但没有更新"累计完成"行的数据

**解决**：
```python
# 累计完成 = 所有责任单位本周完成数之和
total_completed_this_week = sum(data['count'] for data in unit_completed.values())

# 写入周报的"累计完成"行
ws.cell(row=累计完成行号, column=2).value = total_completed_this_week
```

### 错误4：剩余未完成工程量统计数据错误
**原因**：没有正确统计"未完工"的记录，没有按责任单位分组统计剩余工种

**解决**：
```python
unit_remaining = defaultdict(lambda: {'count': 0, 'types': defaultdict(int), 'total': 0, 'completed': 0})

for row in range(2, ws.max_row + 1):
    responsibility = ws.cell(row=row, column=9).value
    work_type = ws.cell(row=row, column=10).value
    completed = ws.cell(row=row, column=12).value
    
    unit_remaining[responsibility]['total'] += 1
    
    if completed == '是':
        unit_remaining[responsibility]['completed'] += 1
    else:
        unit_remaining[responsibility]['count'] += 1  # 剩余数
        if work_type:
            unit_remaining[responsibility]['types'][work_type] += 1  # 只统计未完工的工种！
```

### 错误5：下周工作计划、协调事项、风险预警错行、重复
**原因**：直接复制了上周周报的内容，但没有根据用户要求更新；行号计算错误

**解决**：
1. **必须让用户提供以下内容**：
   - 下周工作计划
   - 需项目公司协调事项
   - 风险预警

2. **不要直接复制上周的内容**，必须根据用户提供的本周数据重新填写

3. **检查所有行号**，确保没有错行

### 错误6：主要剩余工种明细统计错误
**表现**："主要剩余工种明细"列填成了"总问题条数"，而非"剩余未完成条数"

**正确做法**：
```python
# 剩余未完成工种明细：只统计 未完工（是否完工 != "是"）的记录
if completed != '是':
    unit_remaining[responsibility]['count'] += 1
    if work_type:
        unit_remaining[responsibility]['types'][work_type] += 1
```

**关键区别**：
- **本周销项明细**：统计本周完成的记录（完成时间在范围内）
- **剩余未完成工种明细**：只统计"是否完工 != 是"的记录

## 格式要点（2026-05-16修正）

1. **完成率格式**：使用小数格式（如0.9578），合计行0.9572
2. **工种明细格式**：统一为"工种：X条"，多工种用"  |  "分隔
3. **本周累计完成说明**：包含责任单位数量（如"5个责任单位参与本周销项，合计完成42条"）
4. **剩余工程量合计行D列**：列出重点关注单位（如"★贝斯特/东海建设完成率偏低需重点关注"）
5. **协调事项表头**：问题类型/问题描述/现场情况（协调诉求在H列/现场情况列）
6. **风险预警表头**：序号/风险等级/户号或范围/风险事项/处置建议（H列）
7. **底部日期**：格式为"数据截止日期：2026年MM月DD日（周X） | 本报告由维保工程部整理提交"

## 使用要点

1. **必须让用户提供以下内容**：
   - 下周工作计划
   - 需项目公司协调事项
   - 风险预警

2. **不要直接复制上周的内容**，必须根据用户提供的本周数据重新填写

3. **检查所有统计数据**，确保与台账一致

4. **检查所有责任单位和工种**，确保没有漏项

5. **检查所有行号**，确保没有错行

6. **检查主要剩余工种明细**，确保是"剩余未完成"条数而非"总条数"

7. **生成后必须让用户检查**，根据反馈修改

## 2026-05-30 工作经验总结

### 合并单元格处理（关键经验）

**问题**：尝试清除合并单元格区域的数据时，会触发 `AttributeError: 'MergedCell' object attribute 'value' is read-only` 错误。

**正确流程**：
```python
# 1. 读取所有合并单元格范围
merged_ranges = list(ws.merged_cells.ranges)

# 2. 取消所有合并
for mr in merged_ranges:
    ws.unmerge_cells(str(mr))

# 3. 清除旧数据并写入新数据
for row in range(start_row, end_row + 1):
    for col in range(1, 5):
        ws.cell(row=row, column=col).value = None  # 先清除
    # 再写入新数据...

# 4. 恢复合并
for mr in merged_ranges:
    ws.merge_cells(str(mr))
```

### 行号定位策略

**问题**：动态文本搜索定位行号时，合并单元格会导致搜索失败。

**解决方案**：模板结构固定时，使用明确行号比动态检测更可靠：
- 行30：剩余未完成工程量标题
- 行31：表头
- 行32-44：数据行（13个单位）
- 行45：合计行

**代码示例**：
```python
# 明确行号定位（推荐）
remaining_start = 32  # 直接指定，不依赖搜索
remaining_end = 44
summary_row = 45
```

### 数据清除范围

**注意**：清除数据时需覆盖所有可能有数据的行，包括旧模板中的空行：
```python
# 清除范围要足够大，确保覆盖旧数据
for row in range(32, 46):  # 包括可能有数据的行
    for col in range(1, 5):
        ws.cell(row=row, column=col).value = None
```

### 本周完成情况

- 文件：`璀璨云著A地块维修周报_2026年5月30日.xlsx`
- 统计：报修9359条，完成9121条，剩余238条，完成率97.46%
- 本周新增35条，本周完成120条
- 剩余12个责任单位，二十冶(75条)最多

### 2026-06-14 重要更新

#### 合并单元格处理新流程

**问题**：恢复合并单元格时，如果合计行（如行46）使用 `A46:J46` 合并，会导致 C2/C3/C4 列的值被吞掉。

**解决方案**：先写入所有数据，再恢复合并，且剩余合计行使用 `D:J` 合并（不是 `A:J`）。

**正确流程**：
```python
# 1. 取消所有合并单元格
for mr in list(ws.merged_cells.ranges):
    ws.unmerge_cells(str(mr))

# 2. 写入所有数据（包括剩余合计行）
ws.cell(row=remaining_summary_row, column=1).value = '合计'
ws.cell(row=remaining_summary_row, column=2).value = round(total_rate, 4)
ws.cell(row=remaining_summary_row, column=3).value = f'{total_remaining}条'
ws.cell(row=remaining_summary_row, column=4).value = '说明文字...'

# 3. 恢复合并（剩余合计行用 D:J 合并）
ws.merge_cells(f'D{remaining_summary_row}:J{remaining_summary_row}')

# 4. 保存
wb.save(OUTPUT)
```

**关键区别**：
- ❌ 错误：`A46:J46` 合并 → 吞掉 C2/C3/C4 的值
- ✅ 正确：`D46:J46` 合并 → C1/C2/C3/C4 都保留

#### 动态行号处理

**问题**：剩余单位数量会变化（如从12个变成14个），合计行位置不固定。

**解决方案**：动态计算合计行位置。

```python
# 计算剩余合计行位置
remaining_summary_row = 32 + len(remaining_list)  # 32是数据起始行
ws.cell(row=remaining_summary_row, column=1).value = '合计'
ws.cell(row=remaining_summary_row, column=2).value = round(total_rate, 4)
ws.cell(row=remaining_summary_row, column=3).value = f'{total_remaining}条'
ws.cell(row=remaining_summary_row, column=4).value = f'共{len(remaining_list)}个责任单位...'
```

### 本周完成情况（2026-06-14）

- 文件：`璀璨云著A地块维修周报_2026年6月14日.xlsx`
- 统计：报修9463条，完成9188条，剩余275条，完成率97.09%
- 本周新增36条，本周完成6条
- 剩余14个责任单位，二十冶(85条)最多
- 新增单位：滁州志鹏(2条洗碗机)

## 参考资料
- 周报模板：`璀璨云著A地块维修周报_YYYY年MM月DD日.xlsx`
- 台账文件：`璀璨云著A地块1、2、7栋报修台账YYYY年M月D日.xlsx`
- 修复脚本：`gen_weekly_0530_v2.py`（2026-05-30工作成果）
- 修复脚本：`gen_weekly_0606.py`（2026-06-06工作成果）
- 修复脚本：`gen_weekly_0614.py`（2026-06-14工作成果，含合并单元格修复）

## 数据计算参考（生成器仅用于算数，版面必须模板复制）

> **2026-07-19 纠正**：`gen_weekly_paradigm.py` 的**从零 layout 重建已废弃**（会破坏合并/换行/行高）。它只剩 `compute_stats()` 可复用——用来从台账算出各责任单位的累计/完成/剩余/完成率与工种明细。版面一律由「复制已知好周报 + 只改 `.cell.value`」得到。

生成器路径：`d:\WorkBuddyClaw\分析报告归档\周报范式研究\gen_weekly_paradigm.py`
依赖 openpyxl，运行环境：`C:\Users\蔡正延\.workbuddy\binaries\python\envs\default\Scripts\python.exe`
范式规范：`d:\WorkBuddyClaw\分析报告归档\周报范式研究\周报范式规范.md`（数据/格式/着色三大逻辑 + 22 条机器校验不变量，可作人工复核清单）

推荐工作流（v7）：
1. `shutil.copy2` 复制一份格式正确的同楼栋周报（或 4/5/10 好周报）为 `target`。
2. 用 `compute_stats()`（或自制脚本读台账）算出本期数据。
3. 仅改写 `target` 中合并区左上角单元格的值（报告期、总体概况、销项明细、剩余表、下周计划、协调、风险、底部日期）。
4. 若单位数量变化：**禁止 `insert_rows`/`delete_rows`**（合并会错位）。改用 v7 确定性重建法：unmerge 全表→逐格复制样式+值→显式 `merge_cells` 重建合并全集（插入点行号整体+1）→行高按+1 还原。
5. 完成率列统一百分比格式 `number_format='0.00"%"'`（用户 2026-07-20 明确要求转百分比）。
6. 保存。人工复核：合并横向到底无吞数据、文本格 `wrap_text=True`、内容多行处行高足够、合计行非单位行、无"上周对比行"。

修复脚本范例（2026-07-19 实战可用）：`d:\WorkBuddyClaw\分析报告归档\周报范式研究\fix_127_report.py`（还原备份→删旧累计行→插真皮阿诺行→合计改标→转百分比→保合并/换行/行高）。

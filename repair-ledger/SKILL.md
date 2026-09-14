---
name: repair-ledger
description: >
  璀璨云著A地块维保报修台账录入 Skill。当用户输入包含以下内容时自动触发： - 日期格式"YYYY年M月D日"后跟随报修消息 -
  含"C2\4-6\10,A1\2\7 维保工程师:"前缀的报修记录 - 含"楼栋-房号"格式（如"1-201"、"7-2604"）的维修问题描述 -
  "电话改为"或"电话变更为"修改请求 - 报修内容含责任单位（如二十冶、海帛丽、恒鼎泰、皮阿诺等）和工种描述
---

# 报修台账录入 Skill

你是璀璨云著A地块维保工程师蔡正延的助手，负责将日常报修消息整理录入Excel台账。

---

## ⚠️ 核心原则（必须遵守）

### 1. 脚本日志强制要求
**所有Python脚本必须写日志文件，禁止依赖stdout输出。**

原因：PowerShell/终端可能会吞掉Python的stdout输出，导致无法确认脚本是否执行成功。

```python
# 脚本开头
log_lines = []

# 脚本中间
log_lines.append(f"处理: {room}")

# 脚本结尾 - 必须写日志文件
log_file = r"D:\WorkBuddyClaw\script_log.txt"
with open(log_file, 'w', encoding='utf-8') as f:
    f.write('\n'.join(log_lines))
```

### 2. 验证结果方式
**完成脚本执行后，必须读取日志文件确认结果。**

```python
# 读取日志确认
with open(r"D:\WorkBuddyClaw\script_log.txt", 'r', encoding='utf-8') as f:
    log_content = f.read()
# 从日志中提取关键信息：Added X records, Total: XXXX
```

### 3. Python 解释器路径
**必须使用系统Python 3.11**，managed Python环境缺少openpyxl。

```powershell
# 正确
& "C:\Users\蔡正延\AppData\Local\Programs\Python\Python311\python.exe" script.py

# 错误 - managed Python 没有 openpyxl
& "C:\Users\蔡正延\.workbuddy\binaries\python\envs\default\Scripts\python.exe" script.py
```

---

## 工作目录

台账文件位于：`D:\WorkBuddyClaw\`

## 处理流程

收到报修内容后，按以下步骤处理：

1. **解析日期** — 从消息开头提取日期（如"2026年4月28日"），确定源文件和目标文件
2. **识别类型** — 区分「电话变更」和「新报修记录」
3. **查询业主信息** — 优先从台账历史查（同栋同室最近一条），台账无记录再查客情表
4. **编写Python脚本** — **必须包含日志文件写入代码**
5. **执行脚本** — 用系统Python 3.11运行
6. **读取日志验证** — 打开日志文件确认"Added X records. Total: XXXX"
7. **汇总输出** — 以表格形式展示写入结果
8. **等待用户确认** — 询问用户台账记录是否正确
9. **上传腾讯文档** — 用户确认后，将台账文件上传到腾讯文档的"璀璨云著"文件夹

---

## 文件命名规则

- 源文件：`璀璨云著A地块1、2、7栋报修台账2026年M月D日.xlsx`（取上一个工作日文件）
- 目标文件：`璀璨云著A地块1、2、7栋报修台账2026年M月D日.xlsx`（当日日期）

## 字段规范（严格遵守——2026-05-16修正版）

### 台账列布局（openpyxl column索引，1-based）
| 列号 | 字段 | Python代码 | 注意事项 |
|------|------|------------|----------|
| 1 | 项目 | `ws.cell(row,1)` | 固定填 `A地块` |
| 2 | 报修渠道 | `ws.cell(row,2)` | 日常报修→`"日常报修"`，验房报告→`"验房报告"` |
| 3 | 楼栋号 | `ws.cell(row,3)` | 整数 1/2/7 |
| 4 | 室号 | `ws.cell(row,4)` | 整数 |
| 5 | 报修人姓名 | `ws.cell(row,5)` | **取台账已有第一个姓氏**，多人名取第一个 |
| 6 | 联系电话 | `ws.cell(row,6)` | **取已有第一个整数号码**，多号码取第一个 |
| 7 | 报修问题详述 | `ws.cell(row,7)` | |
| 8 | 房修工程师 | `ws.cell(row,8)` | 10栋→`"李世鹤"`，其余栋→`"蔡正延"` |
| 9 | 责任单位 | `ws.cell(row,9)` | 见下方标准名称表 |
| 10 | 工种 | `ws.cell(row,10)` | |
| 11 | 报修日期 | `ws.cell(row,11)` | 字符串 `"2026.06.30"`（YYYY.MM.DD格式） |
| 12 | 是否完工 | `ws.cell(row,12)` | 未完成→留空，已完成→`"是"` |
| 13 | 完成时间 | `ws.cell(row,13)` | 新录入留空 |
| 14 | 备注 | `ws.cell(row,14)` | 密码、特殊说明填此处 |

> **⚠️ 2026-05-16教训**：列布局是 Col5=姓名，Col6=电话，Col7=问题，Col8=工程师，Col9=责任单位，Col10=工种。写反一位会导致整列数据永久错位！录入前务必核对。

## 责任单位标准名称

| 责任单位 | 常见工种 |
|----------|----------|
| 圣象地板 | 地板、踢脚线、地板压条 |
| 金牌厨柜 | 橱柜、台面、美容 |
| 海帛丽 | 淋浴屏 |
| 杭州三角洲 | 新风 |
| 星网天合 | 智能面板 |
| 国网 | 空调 |
| 和乐 | 入户门、智能锁、美容、生锈 |
| 恒鼎泰 | 万能工、水电、油漆、打胶、石材、土建、浴霸 |
| 港华 | 地暖 |
| 恺云居 | 玄关柜、卫浴柜、石材 |
| 皮阿诺 | 户内门、美容、厨房门 |
| 二十冶 | 门窗、土建、防火门 |
| 万华 | 水电、油漆、打胶、瓦工 |

## 消息格式解析规则

原始消息格式：
```
C2\4-6\10,A1\2\7 维保工程师: {楼栋}-{室号} {问题描述}（{责任单位}{工种}）
```

- **楼栋**：取"-"前的数字（1/2/7），忽略"A"前缀（如"A1-2206"→栋=1，室=2206）
- **责任单位**：括号内第一个词，对照标准名称表修正
- **工种**：括号内责任单位之后的词
- **电话变更**：消息含"电话改为"或"电话变更为"→修改该室所有台账记录的联系电话
- **玻璃划痕**：问题括号中含"玻璃划痕"时，备注列填写"备注：玻璃划痕"
- **密码/备注**：问题描述中含密码/特殊说明时，填入备注列

## 业主信息查询逻辑

```python
# 优先从台账历史查（同栋同室）。关键：取「有业主信息的最近一行」，而非机械取底行！
# 某户号底部可能有无业主信息的空行（如未售房 / 历史占位行），会盖住真实业主。
def get_owner_from_ledger(ws, bld, room):
    best_name, best_phone = '', None
    for i in range(ws.max_row, 1, -1):   # 倒序，从最新往回找
        if ws.cell(row=i, column=3).value == bld \
           and ws.cell(row=i, column=4).value == room:
            nm = ws.cell(row=i, column=5).value
            ph = ws.cell(row=i, column=6).value
            name = ''
            if nm is not None:
                parts = str(nm).split()
                if parts: name = parts[0]
            phone = None
            if ph is not None:
                try: phone = int(float(str(ph).split()[0]))
                except: phone = None
            # 命中带业主信息的行立即返回（最近一条有效业主）
            if name or phone is not None:
                return name, phone
            # 否则继续往上找更早的有效记录
            best_name, best_phone = name, phone
    return best_name, best_phone
```

台账无记录时，再从 `1#2#7#客情表1231.xlsx` 查询（注意：客情表对7栋覆盖有限）。

> **⚠️ 2026-08-04 踩坑修正**：原逻辑「倒序取第一条匹配行」会命中底部空行（无业主信息），把真实业主（在更早的行里）盖掉。必须取「有业主信息的最近一行」。

## 密码源读取（A127 / A345610）

- 两密码源均为横向成对布局：列 (1,2)(3,4)... = 户号|密码；行 = 楼层。
- **⚠️ 行号偏移不一致**：A127（1/2/7栋）楼层从 **row2** 起算（row2=1楼）；A345610（4/5/10栋）楼层从 **row3** 起算（row5=3楼），比前者多 1 行表头偏移。
- **不要用 `row = floor + 1` 推算行号**（对 A345610 会整层错位、静默查不到密码）。**改为全表扫描户号**：
```python
def get_pwd_from_source(pwb, bld, room):
    smap = {1:"1栋", 2:"2栋", 7:"7栋", 4:"4栋", 5:"5栋", 10:"10栋"}
    target = smap.get(bld); sn = None
    for s in pwb.sheetnames:
        if s.strip() == target: sn = s; break
    if sn is None: return None
    ws = pwb[sn]
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column, 2):
            if ws.cell(row=r, column=c).value == room:
                return ws.cell(row=r, column=c+1).value
    return None
```
- **密码填写口径（SOP）**：仅当该录入户 `col5姓名 与 col6电话 皆空`（无业主信息）时才补 `密码:xxx`；有业主信息的行一律不填密码（避免污染备注）。
- **4-1102 免密码例外**：永不填密码，备注统一 `不给密码、联系 138****0000开门`。

---

## Python脚本标准模板

```python
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime
import shutil

# === 日志 ===
log_lines = []

SRC = r"D:\WorkBuddyClaw\璀璨云著A地块1、2、7栋报修台账2026年X月X日.xlsx"
DST = r"D:\WorkBuddyClaw\璀璨云著A地块1、2、7栋报修台账2026年X月X日.xlsx"
DATE = "2026.06.30"  # YYYY.MM.DD格式字符串

# 复制源文件
shutil.copy(SRC, DST)
log_lines.append(f"源文件: {SRC}")
log_lines.append(f"目标文件: {DST}")

wb = openpyxl.load_workbook(DST)
ws = wb["台账"]
FILL_NONE = PatternFill(fill_type=None)

def add_record(ws, build, room, owner_name, phone, problem, unit, trade, channel="日常报修", engineer="蔡正延", remark=""):
    new_row = ws.max_row + 1
    ws.cell(new_row, 1).value = "A地块"
    ws.cell(new_row, 2).value = channel  # "日常报修"或"验房报告"
    ws.cell(new_row, 3).value = build
    ws.cell(new_row, 4).value = room
    ws.cell(new_row, 5).value = owner_name
    ws.cell(new_row, 6).value = phone
    ws.cell(new_row, 7).value = problem
    ws.cell(new_row, 8).value = engineer  # 10栋→"李世鹤"，其余→"蔡正延"
    ws.cell(new_row, 9).value = unit
    ws.cell(new_row, 10).value = trade
    ws.cell(new_row, 11).value = DATE
    ws.cell(new_row, 12).value = ""   # 未完成留空，完成填"是"
    ws.cell(new_row, 13).value = ""
    ws.cell(new_row, 14).value = remark
    for c in range(1, 15):
        ws.cell(new_row, c).fill = FILL_NONE
    log_lines.append(f"  Row{new_row}: {build}-{room} | {owner_name} | {phone} | {unit}/{trade}")
    return new_row

# === 在此添加记录 ===

wb.save(DST)
log_lines.append(f"SAVED! Total records: {ws.max_row - 1}")

# === 必须：写日志文件 ===
log_file = r"D:\WorkBuddyClaw\add_records_log.txt"
with open(log_file, 'w', encoding='utf-8') as f:
    f.write('\n'.join(log_lines))

print("日志已写入: " + log_file)
```

---

## 输出格式

完成录入后，输出如下格式：

```
✅ 录入完成！
文件：璀璨云著A地块1、2、7栋报修台账2026年X月X日.xlsx
总记录：XXXX 条（新增 N 条）

🔄 电话变更（如有）：
{室号} 全部 N 条记录，电话统一改为 {新电话}

📋 今日 N 条报修明细：
| # | 楼栋 | 责任单位 | 工种 | 问题 | 业主 |
...
```

## 注意事项

- 同一条消息若含多个问题，每个问题拆分为独立一行
- 原始消息未标注责任单位时，从台账历史同室记录推断，并在汇总时提示用户确认
- **Excel文件被占用时**：如果目标文件被Excel打开，脚本会报`PermissionError`，需要关闭Excel后重试
- **验证结果**：每次执行后必须读取日志文件（如`add_records_log.txt`），确认"SAVED! Total records: XXXX"

---

## 常见问题处理

### Q1: 脚本执行后无输出？
**原因**：PowerShell吞掉了stdout。
**解决**：检查日志文件，如`D:\WorkBuddyClaw\add_records_log.txt`。

### Q2: 提示"Permission denied"？
**原因**：目标Excel文件被Excel占用。
**解决**：关闭Excel，重新运行脚本。

### Q3: 提示"ModuleNotFoundError: No module named 'openpyxl'"？
**原因**：使用了managed Python（缺少openpyxl）。
**解决**：使用系统Python 3.11路径：`C:\Users\蔡正延\AppData\Local\Programs\Python\Python311\python.exe`

---

## 📤 腾讯文档上传流程

### 触发时机
**用户确认台账记录无误后**，立即上传到腾讯文档。

### 上传步骤

1. **确认目标文件夹** — 腾讯文档「璀璨云著」文件夹
2. **调用上传工具** — 使用 `tencent-docs` skill 的 `mcp__tencent-docs__manage.pre_import` 工具
3. **上传文件** — 将Excel文件上传到预导入URL
4. **触发导入** — 使用 `mcp__tencent-docs__manage.async_import`
5. **轮询进度** — 使用 `mcp__tencent-docs__manage.import_progress`，直到 progress=100
6. **确认完成** — 提示用户文件已上传

### 上传脚本示例

```python
# 步骤1：预导入
result = mcp__tencent-docs__manage.pre_import(
    file_name="璀璨云著A地块1、2、7栋报修台账2026年X月X日.xlsx",
    file_size=os.path.getsize(file_path),
    file_md5="..."  # 计算MD5
)
# 返回：upload_url, file_key, task_id

# 步骤2：上传文件到upload_url
import requests
with open(file_path, 'rb') as f:
    requests.put(result['upload_url'], data=f)

# 步骤3：触发导入
mcp__tencent-docs__manage.async_import(
    task_id=result['task_id'],
    file_size=...,
    file_key=result['file_key'],
    file_name=...,
    file_md5=...
)

# 步骤4：轮询进度（每5秒一次）
while True:
    progress = mcp__tencent-docs__manage.import_progress(task_id=result['task_id'])
    if progress['progress'] == 100:
        break
    time.sleep(5)
```

### 注意事项
- **必须等待用户确认后再上传**
- **文件命名**：保持原始命名格式 `璀璨云著A地块1、2、7栋报修台账2026年X月X日.xlsx`
- **上传失败**：如果上传失败，提示用户手动上传
- **确认上传结果**：进度100%后，返回文件URL给用户

---

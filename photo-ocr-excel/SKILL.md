---
name: photo-ocr-excel
description: 照片水印OCR识别并追加到Excel汇总表。适用于施工检查、空鼓检查、物业巡检等带水印照片的批量处理场景。触发词：照片汇总、空鼓检查、施工查验、照片整理、OCR识别、生成汇总表、照片归档、检查照片、水印识别、图片生成Excel
allowed-tools:
  - Bash
  - Read
  - Write
  - Glob
---

# 照片OCR识别合并Excel工具

将带水印的检查照片通过OCR识别后，按规范格式追加到Excel汇总表中。

## 适用场景

- 施工质量检查（空鼓、裂缝、渗漏等）
- 物业巡检照片归档
- 精装修样板检查
- 任何带有时空水印的现场照片批量处理

## 核心工作流程

### 1. 环境准备

```bash
# 创建/激活 Python 虚拟环境
PYTHON="C:/ProgramData/WorkBuddy/chromium-env/6s4p0b/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
PIP="C:/ProgramData/WorkBuddy/chromium-env/6s4p0b/.workbuddy/binaries/python/envs/default/Scripts/pip.exe"

# 安装依赖（首次使用）
$PIP install pandas openpyxl rapidocr_onnxruntime Pillow
```

**依赖说明：**
- `rapidocr_onnxruntime`：轻量级OCR引擎，基于ONNX Runtime，无需PyTorch
- `openpyxl`：Excel读写库，支持嵌入图片
- `Pillow`：图像预处理

**⚠️ 避免使用 easyocr**：依赖 PyTorch (~2GB)，安装极慢且易失败

### 2. 读取现有Excel结构

```python
import openpyxl
from openpyxl.drawing.image import Image as XlImage

wb = openpyxl.load_workbook(excel_path)
ws = wb.active

# 获取表头（通常在第1行）
headers = [cell.value for cell in ws[1]]

# 获取现有数据行数
existing_rows = ws.max_row
```

**典型表结构（空鼓检查）：**
| 序号 | 楼栋 | 户号 | 问题描述 | 照片 |
|------|------|------|----------|------|
| 1 | 10栋 | 3-301 | 厨房墙面空鼓 | [图片] |

### 3. OCR识别与水印解析

```python
from rapidocr_onnxruntime import RapidOCR
import re

ocr = RapidOCR()

def parse_watermark(texts):
    """
    解析水印文本，提取楼栋号、户号、问题描述

    支持的水印格式：
    1. 公区格式：10栋3层 + 东单元公区
    2. 房间格式：5-201次卫 / 5-301厨房
    3. 混合格式：10栋3层西单元公区
    """
    if not texts:
        return None, None, None

    # 合并所有文本并去除空格（解决OCR换行分割问题）
    full_text = ' '.join(texts)
    compact = re.sub(r'\s+', '', full_text)

    building = None
    unit = None
    desc = None

    # 模式1：楼栋+楼层+单元+公区
    m = re.search(r'(\d+栋)(\d+)层(.*?)公区', compact)
    if m:
        building = m.group(1)
        unit = f"{m.group(2)}层{m.group(3)}"
        desc = '公区'
        return building, unit, desc

    # 模式2：户号格式（如 5-201次卫）
    m = re.search(r'(\d+-\d+)(.*)', compact)
    if m:
        unit = m.group(1)
        desc = m.group(2).strip() if m.group(2) else None
        # 从户号推断楼栋
        building_num = unit.split('-')[0]
        building = f"{building_num}栋"
        return building, unit, desc

    # 模式3：仅有楼层信息（如 ·21层西单元公区）
    m = re.search(r'(\d+)层(.*?)公区', compact)
    if m:
        unit = f"{m.group(1)}层{m.group(2)}"
        desc = '公区'
        # 注意：此模式无法确定楼栋号，需人工确认或从文件名推断
        return building, unit, desc

    return building, unit, desc
```

### 4. 批量处理并追加到Excel

```python
from pathlib import Path
import os

def process_photos(photo_dir, excel_path, output_path):
    """批量处理照片并追加到Excel"""

    # 收集所有照片
    photos = sorted(Path(photo_dir).glob('*.jpg')) + \
             sorted(Path(photo_dir).glob('*.jpeg')) + \
             sorted(Path(photo_dir).glob('*.png'))

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    start_row = ws.max_row + 1

    results = []
    failed = []

    for idx, photo_path in enumerate(photos):
        try:
            # OCR识别
            result, _ = ocr(str(photo_path))

            if result:
                texts = [item[1] for item in result]
                building, unit, desc = parse_watermark(texts)

                if building or unit:
                    # 写入数据
                    row = start_row + idx
                    ws.cell(row=row, column=1, value=start_row + idx - 1)  # 序号
                    ws.cell(row=row, column=2, value=building)  # 楼栋
                    ws.cell(row=row, column=3, value=unit)  # 户号
                    ws.cell(row=row, column=4, value=desc)  # 问题描述

                    # 嵌入照片（缩放到合适大小）
                    img = XlImage(str(photo_path))
                    img.width = 150
                    img.height = 200
                    ws.add_image(img, f'E{row}')

                    results.append({
                        'file': photo_path.name,
                        'building': building,
                        'unit': unit,
                        'desc': desc
                    })
                else:
                    failed.append(photo_path.name)
            else:
                failed.append(photo_path.name)

        except Exception as e:
            failed.append(f"{photo_path.name}: {str(e)}")

    wb.save(output_path)
    return results, failed
```

### 5. 异常处理

#### 5.1 无水印照片
- **现象**：OCR返回空结果或无法解析
- **处理**：记录到失败列表，人工查看照片内容后手动补充
- **常见原因**：照片本身无水印、水印区域被裁剪、图片质量过低

#### 5.2 楼栋号误判
- **现象**：水印中无明确楼栋号，解析逻辑使用默认值导致错误
- **处理**：编写修正脚本，批量更新指定行的楼栋号

```python
def fix_building_numbers(excel_path, row_range, new_building):
    """修正指定行范围的楼栋号"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    for row in row_range:
        current = ws.cell(row=row, column=2).value
        # 仅修正符合条件的行（如包含"层"字的错误记录）
        if current and '层' in str(ws.cell(row=row, column=3).value or ''):
            ws.cell(row=row, column=2, value=new_building)

    wb.save(excel_path)
```

#### 5.3 OCR换行分割问题
- **现象**：水印文本被OCR识别为多行，导致正则匹配失败
- **解决**：在解析前使用 `re.sub(r'\s+', '', text)` 去除所有空格

## 踩坑经验

### 1. OCR引擎选择
- ❌ `easyocr`：依赖 PyTorch (~2GB)，安装极慢，Windows下常失败
- ✅ `rapidocr_onnxruntime`：轻量级，基于ONNX Runtime，安装快，识别率高

### 2. 图像预处理效果有限
- 对比度增强、锐化等预处理对水印识别帮助不大
- 水印本身清晰度是关键，预处理无法弥补原图质量问题

### 3. 水印格式多样性
- 不同项目/系统的水印格式差异大
- 建议先手动查看10-20张照片，确认水印格式后再编写解析规则
- 预留人工确认环节，特别是楼栋号等关键字段

### 4. Excel嵌入图片的大小控制
- 图片过大会导致Excel文件体积膨胀（本次220张照片 → 68MB）
- 建议设置固定尺寸：`img.width = 150; img.height = 200`

### 5. 批量操作原则
- **先预览后执行**：首次运行时先处理5-10张照片验证逻辑
- **保留中间结果**：保存OCR原始结果，便于后续排查
- **分批处理**：大量照片可分批处理，避免单次运行时间过长

## 输出规范

### 成功处理报告
```
✅ 处理完成
- 总照片数：220
- 成功识别：218 (99.1%)
- 识别失败：2 (无水印)
- 楼栋号修正：14
- 输出文件：10栋空鼓检查汇总表.xlsx (68.2MB)
```

### 失败记录格式
```csv
文件名,失败原因
微信图片_20260424121108_1062_35.jpg,无水印
微信图片_20260424121121_1064_35.jpg,无水印
```

## 相关工具

- `photo-summary-excel`：照片OCR水印识别生成Excel汇总表（项目级skill）
- `inspection-report`：精装修样板联合检查报告生成
- `repair-ledger`：璀璨云著A地块维保报修台账录入

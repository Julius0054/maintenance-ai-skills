---
name: ledger-summary-formula
description: |
  台账「明细→汇总」公式化工具。将明细表（含 分区/楼栋/金额列）按 分区+楼栋 汇总到汇总表，
  通过「辅助列金额数值 + SUMIFS 公式」实现：明细改金额或新增行后，汇总表自动重算，无需手动刷新。
  当用户有 Excel 台账（明细 sheet + 金额列），需要生成各楼栋费用总计、且希望后续明细变动自动同步汇总时使用。
  典型触发：汇总金额、明细汇总、费用总计、楼栋汇总、公式统计台账、台账汇总公式化。
agent_created: true
---

# 台账 明细→汇总 公式化

## 何时用

- 一份 Excel 台账，含「明细」类工作表（列大致为 序号 / 分区 / 楼栋 / 户号 / 变更名称 / 成本审核金额 / 第三方单位…），以及要生成或刷新的「汇总」目标表（列大致为 分区 / 楼栋 / 费用总计）。
- 用户要求：汇总金额用**公式**自动统计，明细改金额、加行后汇总跟着变（不要写死数字）。

## 通用解法（三段式）

1. **明细加辅助列「金额数值」**：把金额列里的混乱写法统一成可求和的数字。
2. **汇总写 SUMIFS 公式**：按 分区 + 楼栋 跨表求和辅助列。
3. **公共区域 + 合计**：楼栋为空的行单列求和，最后一行 SUM。

辅助列填到足够多行（默认 500），后续加行自动覆盖。

## 标准流程

### 0. 选对源文件 + 备份（最容易翻车的一步）

- **关键**：用户可能在工作区副本或历史台账里手工改过金额。**先用 openpyxl 读两个候选文件比一下哪个是权威版**（看金额列是否最新），别用错源——这次就踩过：历史台账那份是旧金额，导致整版算错。
- 把源文件 `cp` 到工作区临时目录再读（沙箱会**静默拦截**读工作区外路径，无任何输出、exit=1）。
- 任何写入前先 `shutil.copy2` 备份原文件到 `备份/` 或 `tmp_rh/`。

### 1. 探测列位置（不要写死列号）

```python
import openpyxl
from openpyxl.utils import get_column_letter
wb = openpyxl.load_workbook(SRC)
m = wb["明细"]
hdr = [m.cell(1, c).value for c in range(1, m.max_column + 1)]
zone_col = hdr.index(next(x for x in hdr if x and "分区" in str(x))) + 1
bld_col  = hdr.index(next(x for x in hdr if x and "楼栋" in str(x))) + 1
amt_col  = hdr.index(next(x for x in hdr if x and ("金额" in str(x) or "费用" in str(x)))) + 1
helper_col = m.max_column + 1
H  = get_column_letter(helper_col)   # 辅助列字母，如 I
ZC = get_column_letter(zone_col)     # 分区列字母
BC = get_column_letter(bld_col)      # 楼栋列字母
```

### 2. 写辅助列公式（金额数值）

规则：`数字原样取；文本带数字取首个数字（如"报价5107.86，暂未审核"→5107.86）；纯文本（费用未审出/暂无价格）记 0`。

```python
m.cell(1, helper_col).value = "金额数值"
FILL = 500
for r in range(2, FILL + 1):
    f = ('=IF(ISNUMBER({a}{r}),{a}{r},'
         'IFERROR(LOOKUP(9.9E+307,--LEFT(MID({a}{r},'
         'MIN(FIND({{0,1,2,3,4,5,6,7,8,9}},{a}{r}&"0")),999),'
         'ROW($1:$999))),0))').format(a=get_column_letter(amt_col), r=r)
    m.cell(r, helper_col).value = f
```

> 注意：公式里没有 `<` / `>`，可安全写在 Bash heredoc 内。heredoc 内**禁用 f-string 的 `{k:<8}` 之类带 `<`/`>` 的格式串**（会被 shell 当重定向，报 unexpected EOF）——遇到对齐需求改用 `%-` 格式化。

### 3. 写汇总表公式（按列头动态定位）

```python
g = wb["汇总"]
ghdr = [g.cell(1, c).value for c in range(1, g.max_column + 1)]
sz_col = ghdr.index(next(x for x in ghdr if x and "分区" in str(x))) + 1
sb_col = ghdr.index(next(x for x in ghdr if x and "楼栋" in str(x))) + 1
sa_col = ghdr.index(next(x for x in ghdr if x and ("费用" in str(x) or "总计" in str(x) or "金额" in str(x)))) + 1
SZ = get_column_letter(sz_col); SB = get_column_letter(sb_col); SA = get_column_letter(sa_col)

# 楼栋行（假设汇总模板已预填 分区/楼栋，行 2..last_bld）
for r in range(2, last_bld_row + 1):
    g.cell(r, sa_col).value = (
        '=SUMIFS(明细!${h}:${h},明细!${z}:${z},{sz}{r},'
        '明细!${b}:${b},{sb}{r})'
    ).format(h=H, z=ZC, b=BC, sz=SZ, sb=SB, r=r)

# 公共区域（楼栋为空）：地下室大堂等，模板里若无则追加一行 lobby_row
g.cell(lobby_row, sa_col).value = '=SUMIFS(明细!${h}:${h},明细!${b}:${b},"")'.format(h=H, b=BC)

# 合计（含楼栋 + 公共区域）
g.cell(total_row, sa_col).value = '=SUM({sa}2:{sa}{n})'.format(sa=SA, n=lobby_row)
```

> `last_bld_row` / `lobby_row` / `total_row` 视模板而定：若汇总已预填楼栋行，直接填；若无公共区域/合计行，在 `g.max_row+1` 追加。

### 4. 校验（落盘前必做）

- 抽看几条公式文本，确认列引用（H/ZC/BC/SA/SZ/SB）正确。
- 用 Python 把「明细」金额按 分区+楼栋 手工求和，与「汇总」公式应得值对账（公式版没有缓存计算值，`data_only=True` 在 WPS 打开后才有值；对账用脚本另算一份）。
- 差额应等于「无楼栋行（公共区域）」金额，否则漏算。

### 5. 同步 + 清理

- 用户要求覆盖原文件名时：`cp` 回原路径（**先确认 WPS 没占用该文件**，否则 cp 会失败，提示用户关闭 WPS）。
- 删除过时冗余副本（旧名 `_汇总` 等），用 `rm -f`，逐条确认 exit。

## 踩坑清单（这次全踩过，务必看）

1. **判空错**：`str(zone) not in (None, "")` 会把 `str(None)` 当字符串 `"None"` 误判为有效楼栋，导致无楼栋的大堂金额错进楼栋合计、还生成 `N/one` 脏数据。**必须 `zone in (None, "")` 先判 None**。
2. **多户行金额分摊**：一行多户时金额只挂在首行，绝不能直接按行求和。用 SUMIFS 按楼栋求和，分摊自然成立（各栋之和 = 明细总额）。
3. **楼栋归一**：同一栋「地块4-104」和「C4-504」要归一成同一栋；公式化后依赖「明细」已正确填写分区/楼栋列，分列阶段就要归一。
4. **源文件选错**：历史台账那份可能是旧金额，务必先比对哪个是用户改过的权威版。
5. **WPS 占用**：原文件被 WPS 打开时 cp 覆盖会失败（Device or resource busy），提示用户先关 WPS。
6. **公式无缓存值**：openpyxl 写公式不存计算值，WPS 打开会自动重算；若某列显示 0，提示设「自动计算」或按 F9。
7. **shell 单命令长度上限**：约 3KB 以上会被截断（报 unexpected EOF 指向 line 1，是长度不是引号）——长脚本拆多段或写紧凑。

## 交付话术

完成后告诉用户：汇总表已是公式，明细改金额/加行会自动重算；辅助列可隐藏；备份位置。

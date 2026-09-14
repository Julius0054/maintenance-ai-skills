# maintenance-ai-skills · 物业维保场景 AI 效率工具集

一组面向「地产维保 / 房修 / 物业工程」真实业务场景的效率工具（WorkBuddy Skill），覆盖报修台账、工程函件、责任单位划分、周报、查验转录、照片 OCR 汇总等高频重复工作。

## 工具清单
| Skill | 解决什么 |
|---|---|
| repair-ledger | 报修台账自动录入与结构化 |
| ccyz-letter | 工程维修催告函 / 扣款函 / 转第三方函生成 |
| a-building-classify | 报修问题 → 责任单位 / 工种自动分类 |
| weekly-report-gen | 维修周报自动生成 |
| weekly-report-editor | 周报修订与校正 |
| pdf-inspection-to-excel | 扫描版验房 PDF → Excel 台账 |
| photo-ocr-excel | 带水印现场照片 OCR → 汇总表 |
| photo-summary-excel | 检查照片批量汇总 |
| ledger-summary-formula | 明细 → 分区/楼栋汇总（SUMIFS 公式化） |
| repair-notification | 维修上门通知批量生成 |
| inspection-report | 查验报告生成 |

## 本人角色与 AI 协作说明（重要）
- **本人负责**：业务痛点梳理、流程与规则定义（如工种 → 责任单位映射）、成效验收
- **代码实现**：借助 AI 协作生成
- 这些工具在真实多地块维保工作中**日常使用**，支撑单人运维多个地块

> 这是当下真实的 AI-native 工作流：由懂业务的人定义问题，AI 协助落地实现。

## 隐私说明
仅收录工具的定义文档与脚本。任何业主个人信息、联系方式、门禁密码等敏感数据均不在本仓库内；文档中的示例手机号已做脱敏处理。
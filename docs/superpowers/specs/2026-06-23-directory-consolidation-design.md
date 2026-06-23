# BBG 电商做账目录归拢设计

日期：2026-06-23

## 目标

将现有的 `BBG_Ecommerce` 和 `Buchhaltung` 两个完整目录归拢到：

`D:\Workstation_Yuhong\BBG\电商\网买做账\`

迁移后保留两个目录原有名称和内部边界，不合并源代码文件；Amazon、eBay、Kaufland、OTTO、TikTok、XPOLLENS 和月度 Umsatzsteuer 数据目录保持现状。

## 目标结构

```text
D:\Workstation_Yuhong\BBG\电商\网买做账\
├── BBG_Ecommerce\
├── Buchhaltung\
├── Amazon\
├── Ebay\
├── Kaufland\
├── Otto\
├── Tiktok\
├── XPOLLENS\
└── Umsatzsteuer_*.xlsx
```

## 迁移策略

采用复制、改造、验证、切换、清理五阶段流程。首先复制两个目录到新位置，旧目录在全部验证完成前保持不动。随后更新脚本、Claude 命令和 Protocol 中的绝对路径，并把适合的代码路径改为基于脚本位置或统一配置解析，避免依赖当前工作目录。

`BBG_Ecommerce` 的 `.git` 历史随整个目录复制。`Buchhaltung` 当前不是独立 Git 仓库，将作为运行支持目录保留，不擅自创建第二个仓库或合并进主仓库历史。

## Python 环境

两个旧 `.venv` 不作为可移植运行环境直接使用，因为其中的激活脚本和入口文件包含旧绝对路径。迁移时在新位置重建虚拟环境，并根据现有环境包清单安装依赖。旧环境在验证完成前保留用于对照。

## 路径与数据流

- 月度原始数据和最终 `Umsatzsteuer_*.xlsx` 继续位于 `网买做账` 根目录及平台子目录。
- `BBG_Ecommerce` 负责读取平台文件并填写月度报表。
- `Buchhaltung` 继续生成 `Kontoauszug_Otto.xlsx`、`Kontoauszug_Ebay.xlsx` 等中间文件。
- 两个代码目录之间的文件引用统一指向新位置。
- `receipts`、`downloaded_receipts`、`logs`、`Ebay_Rechnung` 和 `Otto` 等运行目录随 `Buchhaltung` 一起迁移。

## 更新范围

需要检查并更新：Python 文件中的硬编码路径、`.claude/commands`、本机 `.claude/settings.local.json`、月度做账 Protocol、可能依赖旧工作目录的启动方式，以及虚拟环境或浏览器驱动相关入口。

不改变各平台的财务计算逻辑、Excel 列结构、匹配规则和文件命名规则。

## 验证与失败处理

迁移后先进行 Python 语法和导入检查，再以只读或副本数据验证各平台路径发现、输入读取和输出定位。涉及写入 Excel 的流程使用测试副本，避免修改正式月度账表。

验证至少覆盖 Amazon、OTTO、eBay、Kaufland、TikTok、XPOLLENS、Sheet1 生成和佣金核查。任何验证失败都停止切换，继续使用旧目录并修复新副本；不删除或覆盖旧目录。

## 完成条件

- 新位置中的 Git 仓库状态和提交历史完整。
- 所有脚本不再引用两个旧目录路径。
- Claude 命令和 Protocol 指向新位置。
- 新虚拟环境可运行，关键流程通过副本验证。
- 新旧关键输入和历史数据数量、大小或校验结果一致。
- 用户确认新位置可正常使用后，旧目录才进入单独的清理步骤。

## 安全边界

本次迁移不会自动删除旧目录。删除旧副本是迁移完成后的独立操作，必须在全部验证通过并再次获得用户确认后执行。

# BBG 电商做账目录归拢实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 `BBG_Ecommerce` 与 `Buchhaltung` 安全复制到 `D:\Workstation_Yuhong\BBG\电商\网买做账`，更新所有路径并验证新位置可独立运行，同时保留旧目录作为回滚副本。

**Architecture:** 保持两个目录边界和月度数据目录不变，只改变代码目录所在位置。先建立可核对清单，再复制除虚拟环境和缓存之外的内容；新位置重建 Python 3.14 虚拟环境，路径引用改为新位置，并通过静态、环境、路径和代表性流程检查后切换使用入口。

**Tech Stack:** Windows PowerShell、Git、Python 3.14、venv、openpyxl、pandas、pdfplumber、requests、Selenium。

## Global Constraints

- 旧目录在迁移和验证期间不得删除、重命名或覆盖。
- 月度原始数据、平台目录和 `Umsatzsteuer_*.xlsx` 保持在 `网买做账` 原位置。
- 不改变财务计算逻辑、Excel 列结构、匹配规则和文件命名规则。
- 两个旧 `.venv` 不复制；在新位置重新创建。
- 正式 Excel 文件不得用于破坏性测试；需要写入时使用副本。
- `.claude/settings.local.json` 与 `.idea` 本机配置不得误加入 Git 提交。

---

### Task 1: 建立迁移基线与依赖清单

**Files:**
- Create: `requirements.txt`
- Create: `../Buchhaltung/requirements.txt`
- Create: `docs/superpowers/migration/2026-06-23-source-manifest.csv`

**Interfaces:**
- Consumes: 两个旧目录和各自现有 `.venv`。
- Produces: 可重建环境的固定依赖清单，以及排除 `.git`、`.venv`、缓存后的源文件大小与 SHA-256 基线。

- [ ] **Step 1: 从两个旧虚拟环境导出依赖**

运行各自的 `python -m pip freeze`，将 BBG 环境保存为 `requirements.txt`，Buchhaltung 环境保存到其目录下同名文件。预期分别包含 `openpyxl==3.1.5`，以及 Buchhaltung 的 pandas、pdfplumber、requests、selenium 等依赖。

- [ ] **Step 2: 生成源文件清单**

递归枚举两个目录，排除 `.git`、`.venv`、`__pycache__` 和临时 IDE 缓存；记录相对路径、字节数和 SHA-256。预期清单非空且两个根目录均有记录。

- [ ] **Step 3: 提交可追溯文件**

仅在 `BBG_Ecommerce` 仓库提交其 `requirements.txt`、迁移清单和本实施计划；Buchhaltung 不是独立仓库，其依赖文件随目录复制。

### Task 2: 复制目录到目标位置

**Files:**
- Create: `D:\Workstation_Yuhong\BBG\电商\网买做账\BBG_Ecommerce\**`
- Create: `D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\**`

**Interfaces:**
- Consumes: Task 1 的源清单。
- Produces: 不含旧虚拟环境和 Python 缓存的新位置副本。

- [ ] **Step 1: 确认目标目录不存在**

若目标 `BBG_Ecommerce` 或 `Buchhaltung` 已存在且非空，停止迁移并报告，不覆盖未知内容。

- [ ] **Step 2: 使用 Robocopy 复制两个目录**

使用 `/E /COPY:DAT /DCOPY:DAT /R:2 /W:1`，排除 `.venv` 与 `__pycache__`。Robocopy 返回码 0–7 视为成功，8 以上停止。

- [ ] **Step 3: 对照源清单验证复制**

在目标位置对同一组文件重新计算相对路径、大小和 SHA-256。除随后才会修改的路径配置外，此时目标清单必须与源清单一致。

### Task 3: 更新路径与运行目录解析

**Files:**
- Modify: `BBG电商月度做账Protocol.md`
- Modify: `.claude/commands/*.md`
- Modify: `.claude/settings.local.json`
- Modify: `*.py` 中引用旧 `BBG_Ecommerce` 或 `Buchhaltung` 的路径
- Modify: `../Buchhaltung/ebay.py`
- Modify: `../Buchhaltung/otto.py`
- Modify: `../Buchhaltung/download_otto_invoice.py`

**Interfaces:**
- Consumes: 目标根目录 `D:\Workstation_Yuhong\BBG\电商\网买做账`。
- Produces: 只引用新代码目录的脚本和操作文档；Buchhaltung 主脚本从脚本位置解析自身工作目录。

- [ ] **Step 1: 机械替换两个旧代码根路径**

将 `D:\Workstation_Yuhong\Yuhong's_File\Python\BBG_Ecommerce` 替换为 `D:\Workstation_Yuhong\BBG\电商\网买做账\BBG_Ecommerce`；将 `D:\Workstation_Yuhong\Yuhong's_File\Python\Buchhaltung` 替换为 `D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung`。只修改文本源文件、命令和文档，不处理二进制财务文件。

- [ ] **Step 2: 消除 Buchhaltung 对当前终端目录的依赖**

在 `ebay.py`、`otto.py`、`download_otto_invoice.py` 中将 `BASE_DIR = os.getcwd()` 改为 `BASE_DIR = os.path.dirname(os.path.abspath(__file__))`。这保持原有子目录结构，同时允许从任意终端位置运行脚本。

- [ ] **Step 3: 扫描残留旧路径**

运行 `rg` 排除 `.git`、`.venv` 和二进制文件。预期新目录中的源码、Protocol 和 Claude 命令不再出现两个旧代码根路径；旧目录本身保持不变。

### Task 4: 重建 Python 环境

**Files:**
- Create: `BBG_Ecommerce/.venv/**`
- Create: `Buchhaltung/.venv/**`

**Interfaces:**
- Consumes: Task 1 的两个 `requirements.txt`。
- Produces: 指向新绝对路径的两个 Python 3.14 虚拟环境。

- [ ] **Step 1: 创建两个虚拟环境**

使用系统 Python 3.14 执行 `python -m venv .venv`。预期两个 `.venv\Scripts\python.exe --version` 均返回 Python 3.14.x。

- [ ] **Step 2: 安装固定依赖**

分别执行 `.venv\Scripts\python.exe -m pip install -r requirements.txt`。预期命令退出码为 0。

- [ ] **Step 3: 核对关键导入**

BBG 环境导入 `openpyxl`；Buchhaltung 环境导入 `openpyxl,pandas,pdfplumber,requests,selenium`。预期无异常。

### Task 5: 静态与只读运行验证

**Files:**
- Test: `BBG_Ecommerce/*.py`
- Test: `Buchhaltung/*.py`
- Test: `BBG电商月度做账Protocol.md`
- Test: `.claude/commands/*.md`

**Interfaces:**
- Consumes: Tasks 2–4 的新目录副本。
- Produces: 语法、路径、Git 和环境验证记录。

- [ ] **Step 1: 解析全部 Python 文件**

使用 `ast.parse` 读取两个目录的顶层 `.py`。预期全部解析成功，不生成业务输出。

- [ ] **Step 2: 验证参数化脚本帮助入口**

在新 BBG 环境运行 `fill_kaufland.py --help`、`fill_tiktok.py --help` 和 `fill_xpollens.py --help`。预期退出码为 0 且显示参数帮助。

- [ ] **Step 3: 验证关键输入定位**

确认 Protocol 所列 `网买做账` 根目录、Amazon、eBay、Kaufland、OTTO、TikTok、XPOLLENS 目录均可从新代码位置解析；缺少某个月份数据不作为迁移失败，但路径不得指回旧代码目录。

- [ ] **Step 4: 验证 Git 完整性**

目标 `BBG_Ecommerce` 的 `git log -1`、分支、远端和源仓库复制时一致；迁移产生的代码路径修改应作为明确的新改动显示，本机配置不纳入提交。

### Task 6: 代表性副本验证与切换

**Files:**
- Create: 系统临时目录中的 `Umsatzsteuer` 测试副本
- Modify: 仅测试副本，不修改正式财务工作簿

**Interfaces:**
- Consumes: 新位置脚本、现有历史数据的只读输入和 Excel 测试副本。
- Produces: 平台流程可启动、可读取输入且输出落在预期位置的证据。

- [ ] **Step 1: 为会写入 Excel 的流程建立临时副本**

从现有最近一期 `Umsatzsteuer_*.xlsx` 复制到系统临时目录，记录源文件 SHA-256，测试后再次核对正式源文件 SHA-256 未变化。

- [ ] **Step 2: 验证平台流程**

对 Amazon、OTTO、eBay、Kaufland、TikTok、XPOLLENS、Sheet1 和佣金核查分别执行安全的帮助、输入发现或测试副本路径检查。任何需要真实凭据、邮件、API 或缺少当月文件的步骤只验证启动与路径，不发送请求、不改正式文件。

- [ ] **Step 3: 提交迁移代码与文档变更**

在目标 `BBG_Ecommerce` 仓库仅提交源码、命令、Protocol 和依赖清单。不得提交 `.venv`、测试副本、财务数据、`.claude/settings.local.json` 或 `.idea`。

- [ ] **Step 4: 输出切换结果和回滚方式**

报告新路径、验证结果、仍需人工登录验证的外部步骤，以及旧目录仍在原位。回滚方式是继续从旧目录运行；旧目录删除必须等待用户另行确认。

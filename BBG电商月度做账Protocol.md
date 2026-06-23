# BBG 电商月度数据报表 Protocol

## 前提条件

用 Claude Code 打开目录：
`D:\Workstation_Yuhong\BBG\电商\网买做账\BBG_Ecommerce`

所有 `/fill-*` 命令均需在此目录下运行。

---

## 一、准备阶段

**步骤 1** — 下载银行对账单

登录 Sparkasse 银行账户，下载当月 Kontodruckansicht，保存至：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Kontodruckansicht_[Monat].pdf`

**步骤 2** — 生成月度报表文件

在 Claude Code 中运行 `/fill-sheet1`，将 PDF 内容读取并填入 Excel，生成：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_[Monat].xlsx`

**步骤 3** — 准备 Kontoauszug.xlsx

打开 `Umsatzsteuer_[Monat].xlsx`，将 Sheet1 的**前 4 列**（Nr、Datum、Zahlungsverkehr、Betrag）复制，粘贴到：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\Kontoauszug.xlsx`（覆盖保存）

---

## 二、平台处理阶段

### OTTO

**步骤 1** — 确认入账笔数

打开 `Umsatzsteuer_[Monat].xlsx` → Sheet1，查看当月 OTTO 入账笔数及日期

**步骤 2** — 逐笔准备凭证文件（每笔重复以下操作）

a. 在以下路径创建文件夹：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Otto\[Jahr]\[Monat]\Zahlungsbeleg_[YYMMDD]`

b. 登录 `bbggermany.service@gmail.com`，搜索对应日期，找到标题为
**"Zahlungsbeleg für die Auszahlung von OTTO Market"** 的邮件，下载附件，
文件名只保留 `AZ-DE-?????????-????-??.pdf` 部分

c. 打开该 PDF，找到 Vorgangstyp 列，从 COMMISSION 行起逐行复制 **Rechnungs-ID**，在邮件中搜索并下载对应附件

d. 将 `AZ-DE-...` 文件放入：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\receipts\`

e. 运行 `download_otto_invoice.py`（同目录下），脚本通过 API 自动下载其余文件至 `downloaded_receipts\`

f. 将 `downloaded_receipts\` 内所有文件**剪切**至 `Zahlungsbeleg_[YYMMDD]` 文件夹，确认 `downloaded_receipts\` 已清空，再清空 `receipts\`

g. 重复 a–f，直到 Sheet1 所有 OTTO 入账均处理完毕

**步骤 3** — 汇总计算

将所有 `Zahlungsbeleg_*` 文件夹复制至：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\Otto\[Jahr]\`
然后运行 `otto.py`，导出 `Kontoauszug_Otto.xlsx`

**步骤 4** — 填入报表

在 Claude Code 中运行 `/fill-otto`

**步骤 5** — 验收

脚本会自动核对每行 Summe 与 Sheet1 Betrag：
- 所有行标记 `[OK]` → 正常，继续
- 出现 `[NG]`（金额不一致）→ 脚本停止，确认差值原因后处理：
  - 若差值应计入 VAT-Korrektur：填入该列，Summe 自动更新
  - 若 Sheet1 Betrag 填错：手动修改文件后重新运行

---

### Amazon

**步骤 1** — 下载并保存文件

在 Amazon 卖家后台下载以下文件，保存至：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Amazon\[YYYY]年[M]月\`

| 来源 | 文件类型 | 说明 |
|------|----------|------|
| 付款 → 所有结算 | `.txt` | 结算报告，每笔入账一个文件 |
| 付款 → 广告发票记录 | `INVOICE-*.pdf` | 广告费发票 |
| 付款 → 报告库 | `*CustomTransaction.csv` | 自定义交易报告 |
| 报告 → Tax Document → 卖家费用税务发票 | `DE-AEU-*.pdf` / `.csv` | 平台服务费税务凭证 |
| 报告 → Tax Document → 亚马逊增值税发票 | `taxInvoice_*.zip` | VAT 发票 |

**步骤 2** — 更新脚本路径

打开 `D:\Workstation_Yuhong\BBG\电商\网买做账\BBG_Ecommerce\fill_amazon.py`，将开头两行改为当月：

```python
FOLDER = r"D:\Workstation_Yuhong\BBG\电商\网买做账\Amazon\[YYYY]年[M]月"
XLSX   = r"D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_[Monat].xlsx"
```

**步骤 3** — 运行脚本

```
python fill_amazon.py
```

**步骤 4** — 验收

检查输出：
- 每笔结算均显示 `✓`，且 Sheet1 核对列有对应匹配
- 无 `[警告] 找不到 Sheet1 匹配行`（若有，说明 Sheet1 有入账尚未匹配，需手动处理）
- 若出现 `[警告] 计算合计 vs Summe 差异`：通常由 Vine / Deal 等特殊发票费引起，属正常，继续即可

---

### eBay

**步骤 1** — 下载税务文件

登录 eBay 卖家后台，下载当月账单发票，保存至：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Ebay\[Jahr]\[Monat]\`

需下载以下两份文件（计算时不使用，备查税用）：
- `invoiceId-[ID]_[YYYY-M].pdf` — 发票 PDF
- `invoiceId-[ID]_[YYYY-M].csv` — 发票明细 CSV

**步骤 2** — 准备 eBay CSV 文件

在以下路径新建文件夹：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\Ebay_Rechnung\bis[YYYYMM]\`

将当月 eBay CSV 账单从以下路径**复制**到该文件夹：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Ebay\[Jahr]\[Monat]\`

**步骤 3** — 更新脚本路径

打开 `D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\ebay.py`，将第 16 行的目录名改为当月：

```python
EBAY_CSV_DIR = os.path.join(BASE_DIR, "Ebay_Rechnung", "bis[YYYYMM]")
```

例如 5 月改为 `"bis202605"`。

**步骤 4** — 运行脚本

在 `D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\` 目录下运行：

```
python ebay.py
```

输出 `Kontoauszug_Ebay.xlsx`（含 `Ebay_Umsatz` sheet）。

**步骤 5** — 填入报表

在 Claude Code 中运行 `/fill-ebay`

**步骤 6** — 验收

- 所有行 `[OK]` → 正常，继续
- 出现 `[NG]`（Summe 与 Sheet1 Betrag 不一致）→ 脚本停止，最常见原因是 **Internationale Gebühr** 遗漏：
  - 在 `Ebay_Rechnung\bis[YYYYMM]\` 中找到对应金额的 CSV，检查 `Internationale Gebühr` 列是否有非零值
  - 在 `Kontoauszug_Ebay.xlsx` 中补录后重新从步骤 4 开始

---

### Kaufland

**步骤 1** — 下载税务文件

登录 Kaufland Seller Portal，下载当月各站点的发票 PDF，保存至：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Kaufland\[YYYY]年\[M]月\`

文件名格式：`kaufland_[xx]_facture_R[YYYYMM]-*.pdf`（`xx` 为国家代码，如 `fr`、`de`）
（计算时不使用，备查税用）

**步骤 2** — 确认文件

- 当月施总结算表：`D:\Workstation_Yuhong\BBG\电商\网买做账\Kaufland\[YYYY]年\[M]月\kaufland-BBG *.xlsx`
- 上月施总结算表（同上，路径改为上月目录）

**步骤 3** — 运行脚本

在 Claude Code 中运行 `/fill-kaufland`，或直接执行：

```
python fill_kaufland.py \
  --umsatz "D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_[Monat].xlsx" \
  --kaufland "<当月施总文件完整路径>" \
  --prev-kaufland "<上月施总文件完整路径>"
```

**注意**：运行前确认 Umsatzsteuer 文件未在 Excel 中打开，否则会报 PermissionError。

**步骤 4** — 验收

检查输出结果：
- `✓ All Sheet1 entries matched`（无 Unmatched 警告）
- 所有行 `InnerChk = 0.00`
- EUR 国家 `S1_Diff = 0.00`；CZK/PLN 允许 ±0.06 EUR 以内偏差（日汇率 vs 月均汇率）

**步骤 5** — 查看待到账记录

脚本运行完毕后会自动生成：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Kaufland_pending_[Monat].txt`

该文件列出本月 SKIP 的所有 payout（即本月银行未到账、预计下月入账的款项）。
**下月执行 Kaufland 步骤时，先打开上月的 pending 文件核对**，确认 Sheet1 中已收到这些款项。

---

### Leroy Merlin（XPOLLENS）

**步骤 1** — 确认文件

以下文件由施总提供，确认已全部保存至：
`D:\Workstation_Yuhong\BBG\电商\网买做账\XPOLLENS\[YYYY.M]\`

| 文件 | 说明 |
|------|------|
| `Leroymerlin-BBGSHOP YY.M.D回款明细.csv` | 每笔回款的明细 CSV（脚本使用） |
| `Leroymerlin-BBGSHOP YY.M.D佣金账单.pdf` | 对应佣金发票（备查税用） |

每笔回款对应一份 CSV 和一份 PDF，数量应一致。

**步骤 2** — 运行脚本

在 Claude Code 中运行 `/fill-xpollens`，或直接执行：

```
python fill_xpollens.py \
  --umsatz "D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_[Monat].xlsx" \
  --year [年份] \
  --month [月份数字]
```

**步骤 3** — 验收

每行 `Check = ✓`（Brutto + Andere = Betrag）

---

### TikTok

**步骤 1** — 下载结算文件

登录 TikTok Shop 卖家后台，下载当月结算报告（含订单明细），保存至：
`D:\Workstation_Yuhong\BBG\电商\网买做账\Tiktok\[YYYY]\[M]月\`

文件名格式：`tiktok-BBGMall YY.M.D-YY.M.D 订单.xlsx`

**注意**：TikTok 结算文件的日期范围可能跨月（如 `26.5.1-26.6.5`），脚本按实际到账（Paid）的 Payment ID 匹配，无需手动裁剪。

**步骤 2** — 运行脚本

在 Claude Code 中运行 `/fill-tiktok`，或直接执行：

```
python fill_tiktok.py \
  --umsatz "D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_[Monat].xlsx" \
  --tiktok "<TikTok 结算文件完整路径>"
```

**步骤 3** — 验收

- 每行 `Check = ✓`（Brutto + Andere = Betrag）
- `✓ All Sheet1 TikTok entries matched`
- 若有 `SKIP`：该笔付款不在当月 Sheet1，通常为跨月结算，下月处理

---

## 三、完成确认

运行 `/calc-provision`，核查当月佣金计算是否正确（进账合计 × 2%）。

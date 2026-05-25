将当月 XPOLLENS（Leroy Merlin）结算数据写入当月 Umsatzsteuer 文件的 Sheet1（Brrutto_Umsatz、Andere Gebühren、Steuersatz、Marktplatz）。

## 所需文件

- **来源文件**：`D:\Workstation_Yuhong\BBG\电商\网买做账\XPOLLENS\<YYYY.M>\` 目录下，文件名含 `YY.M.D回款明细.csv` 的所有 CSV（分号分隔，latin-1 编码）
- **目标文件**：当月 Umsatzsteuer 文件，例如 `Umsatzsteuer_April.xlsx`
  - 如果用户在指令中指定了月份（如"四月"、"April"），自动推断对应文件名
  - 如果未指定，询问用户
- **处理脚本**：`D:\Workstation_Yuhong\Yuhong's_File\Python\BBG_Ecommerce\fill_xpollens.py`

---

## 月份与文件名对照

| 月份 | Umsatzsteuer 文件名 | year / month 参数 |
|------|---------------------|-------------------|
| 一月 / Januar    | Umsatzsteuer_Januar.xlsx    | 2026 / 1 |
| 二月 / Februar   | Umsatzsteuer_Februar.xlsx   | 2026 / 2 |
| 三月 / März      | Umsatzsteuer_März.xlsx      | 2026 / 3 |
| 四月 / April     | Umsatzsteuer_April.xlsx     | 2026 / 4 |
| 五月 / Mai       | Umsatzsteuer_Mai.xlsx       | 2026 / 5 |
| 六月 / Juni      | Umsatzsteuer_Juni.xlsx      | 2026 / 6 |

---

## 执行步骤

### 第一步：确认文件路径

1. 根据月份推断 Umsatzsteuer 文件路径
2. 确认文件存在后继续
3. 提醒用户确保 Umsatzsteuer 文件未在 Excel 中打开

### 第二步：运行脚本

```
python "D:\Workstation_Yuhong\Yuhong's_File\Python\BBG_Ecommerce\fill_xpollens.py" \
  --umsatz "<Umsatzsteuer文件路径>" \
  --year <年份> \
  --month <月份数字>
```

脚本会自动在 `XPOLLENS` 目录的所有子目录中搜索文件名匹配 `YY.M.D回款明细.csv` 的文件。

### 第三步：解读输出并汇报

脚本输出分三部分：

**CSV 识别结果**：列出找到的每个结算周期文件及其计算出的 Betrag / Brutto / Andere。

**匹配日志**：
- `MATCH`：CSV 的 Betrag 与 Sheet1 某行精确匹配（误差 <0.02）→ 写入
- `SKIP`：CSV 的 Betrag 在 Sheet1 中无对应条目 → 可能下月才到账，正常
- `⚠ Unmatched`：Sheet1 有 XPOLLENS 条目但无 CSV 匹配 → 需要确认

**写入明细**：每行显示 Row / Brutto / Andere / Betrag / Check（✓ 表示 Brutto + Andere = Betrag）。

### 第四步：异常处理

**如果 Check 显示 ⚠**：
- Brutto + Andere ≠ Betrag，说明数据有误，检查 CSV 是否包含未识别的 Type
- 列出问题行的 Type，告知用户后修正脚本再重跑

**如果有 Unmatched Sheet1 条目**：
- 可能是跨期到账（上月末结算、本月初到账）
- 检查上月 XPOLLENS CSV 的最后一个结算周期，确认是否匹配

---

## 计算规则

每个 `回款明细.csv` 对应一个结算周期，脚本读取**全部行**（不按 Date received 过滤）：

| CSV Type | 归入 |
|----------|------|
| Order amount / Order amount tax / Shipping charges / Shipping tax | Brutto_Umsatz（正数=销售，含税） |
| Order amount refund / tax refund / Shipping refund | Brutto_Umsatz（负数=退款，从 Brutto 中扣除） |
| Commission / Commission (excl, tax) / Commission refund | Andere Gebühren |
| Subscription fee | Andere Gebühren（月平台订阅费） |
| Payment | 提取 Betrag（实际到账金额 = -Payment） |

**Steuersatz = 0**（Leroy Merlin 佣金 VAT 0%），**Marktplatz = Leroymerlin**。

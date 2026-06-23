将 Kontoauszug_Ebay.xlsx 中的 eBay 销售数据填写到当月 Umsatzsteuer 文件的 Ebay sheet，核对与 Sheet1 Betrag 是否一致，确认无误后填入 Brrutto_Umsatz 和 Andere Gebühren。

## 所需文件

- **来源文件**：`D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\Kontoauszug_Ebay.xlsx`，Sheet 名：`Ebay_Umsatz`
- **目标文件**：当月 Umsatzsteuer 文件，例如 `D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_März.xlsx`
  - 如果用户在指令中指定了月份（如"三月"、"April"），自动推断对应文件名
  - 如果未指定，询问用户
- **参考文件（核查用）**：`D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\Ebay_Rechnung\bis[YYYYMM]\`，CSV 文件名即为对应 Betrag 金额（如 `36.35.csv`）

如果 openpyxl 未安装，先运行 `pip install openpyxl`。

---

## 第一步：从 Kontoauszug_Ebay.xlsx 复制全部数据到 Umsatzsteuer Ebay sheet

Kontoauszug Ebay_Umsatz 结构（12列）：
- Col 1=Nr, Col 2=Marktplatz, Col 3=Netto_Sale, Col 4=Ust, Col 5=Brutto_Sale,
  Col 6=Erstattung(19%), Col 7=Andere Gebühre(19%), Col 8=Einbehalten(19%),
  Col 9=Commission(19%), Col 10=Internationale Gebühr(19%), Col 11=Einzahlung, Col 12=Summe

Umsatzsteuer Ebay sheet 结构（15列）：
- Col 1-12 与 Kontoauszug 相同，Col 13=空, Col 14=check（勿修改）

**处理逻辑**：
- 如果 Ebay sheet 中该 Nr 行已存在：整行更新（col 1-12）
- 如果 Ebay sheet 中该 Nr 行不存在：整行插入（col 1-12），col 13-15 不动

完成后输出写入结果（Nr、action、Summe）。

---

## 第二步：核对 Summe 与 Sheet1 Betrag 是否一致

Sheet1 结构：
- Col 1 = Nr，Col 3 = Zahlungsverkehr（值为 `"eBay"`），Col 4 = Betrag

遍历 Sheet1 中所有 Zahlungsverkehr 为 `"eBay"` 的行，按 Nr 与 Ebay sheet 的 Summe（Col 12）对比：

- **一致**：标记 [OK]
- **Sheet1 中有但 Kontoauszug 中缺失**：标记 [??]，列出 Nr 和 Betrag
- **金额不一致**：标记 [NG]，列出 Nr、Sheet1 Betrag、Ebay Summe、diff（= Betrag − Summe）

**如果存在不一致，停止执行**，汇总所有问题并等待用户说明。

### 差异排查提示
最常见原因是 **Internationale Gebühr** 遗漏，对应场景：
- 瑞士、英国等非欧元区买家订单会产生 Internationale Gebühr（负数）
- 若该订单有退款（Rückerstattung），国际费也会退还（正数）
- 两者需同时录入，净额为 0；若只录了其中一笔就会产生差值

排查方法：在 `Ebay_Rechnung\bis[YYYYMM]\` 中找到对应金额的 CSV 文件（文件名 = Sheet1 Betrag），在数据行中查找 `Internationale Gebühr` 列是否有非零值。

**差异处理**：
- 若确认是 Internationale Gebühr 问题：请用户在 Kontoauszug 中修正，重新跑脚本
- 若是其他原因（如 Sheet1 填错）：按用户指示处理
- **不要自动用差值填写，原因各种各样**

所有行一致后进入第三步。

---

## 第三步：填入 Brrutto_Umsatz 和 Andere Gebühren

Ebay sheet 各列含义：
- Col 5 = Brutto_Sale
- Col 6 = Erstattung(19%)
- Col 7 = Andere Gebühre(19%)
- Col 8 = Einbehalten(19%)
- Col 9 = Commission(19%)
- Col 10 = Internationale Gebühr(19%)
- Col 11 = Einzahlung

计算规则：
- **Brrutto_Umsatz** = Col 5（Brutto_Sale）
- **Andere Gebühren** = sum(Col 6 到 Col 11)

Sheet1 目标列：
- Col 5 = Brrutto_Umsatz
- Col 6 = Andere Gebühren

按 Nr 匹配，将计算结果写入 Sheet1 对应 eBay 行，保存文件，输出汇总表格。

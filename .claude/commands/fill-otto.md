将 Kontoauszug_Otto.xlsx 中计算好的 Summe 拷贝到当月 Umsatzsteuer 文件的 OTTO sheet，核对与 Sheet1 Betrag 是否一致，确认无误后填入 Brrutto_Umsatz 和 Andere Gebühren。

## 所需文件

- **来源文件**：`D:\Workstation_Yuhong\BBG\电商\网买做账\Buchhaltung\Kontoauszug_Otto.xlsx`，Sheet 名：`Otto`
- **目标文件**：当月 Umsatzsteuer 文件，例如 `D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_März.xlsx`
  - 如果用户在指令中指定了月份（如"三月"、"April"），自动推断对应文件名
  - 如果未指定，询问用户

如果 openpyxl 未安装，先运行 `pip install openpyxl`。

---

## 第一步：从 Kontoauszug_Otto.xlsx 拷贝全部数据到 Umsatzsteuer OTTO sheet

Kontoauszug_Otto.xlsx Sheet `Otto` 结构（11列）：
- Col 1=Nr, Col 2=Marktplatz, Col 3=Netto_Sale, Col 4=Ust, Col 5=Brutto_Sale,
  Col 6=Refund(19%), Col 7=Commission(19%), Col 8=Commission_rev(19%),
  Col 9=Paymentfee(19%), Col 10=Paymentfee_ref(19%), Col 11=Summe

Umsatzsteuer OTTO sheet 结构（12列）：
- 同上，但 Col 11=VAT-Korrektur(19%)（额外列），Col 12=Summe

**处理逻辑**：
- 如果 OTTO sheet 中该 Nr 行已存在：仅更新 Col 12（Summe）
- 如果 OTTO sheet 中该 Nr 行不存在（为空）：整行插入，Col 1-10 从 Kontoauszug 复制，Col 11（VAT-Korrektur）留空，Col 12（Summe）= Kontoauszug Summe

完成后输出写入结果。

---

## 第二步：核对 Summe 与 Sheet1 Betrag 是否一致

Sheet1 结构：
- Col 1 = Nr，Col 3 = Zahlungsverkehr，Col 4 = Betrag

遍历 Sheet1 中所有 Zahlungsverkehr 为 "OTTO" 的行，按 Nr 与 OTTO sheet 的 Summe（Col 12）对比：

- **一致**：标记为 [OK]
- **不一致**：标记为 [NG]，列出 Nr、Sheet1 Betrag、OTTO Summe、差值（diff = Betrag − Summe）

**如果存在不一致，停止执行**，告知用户每个差异的 Nr 和 diff，并询问如何处理：
- 例如用户说"差值填入 VAT-Korrektur"：则将 diff 写入该行 OTTO sheet Col 11，并将 Col 12 Summe 更新为 Kontoauszug Summe + VAT-Korrektur
- 例如用户说"Sheet1 Betrag 填错了"：提示用户手动修改文件后重新运行
- 其他情况按用户指示处理

所有行一致后再进入第三步。

---

## 第三步：填入 Brrutto_Umsatz 和 Andere Gebühren（原 fill-otto 逻辑）

OTTO sheet 各列含义：
- Col 5 = Brutto_Sale
- Col 6 = Refund(19%)
- Col 7 = Commission(19%)
- Col 8 = Commission_rev(19%)
- Col 9 = Paymentfee(19%)
- Col 10 = Paymentfee_ref(19%)
- Col 11 = VAT-Korrektur(19%)（可能为空，视为 0）

计算规则：
- **Brrutto_Umsatz** = Col 5（Brutto_Sale）
- **Andere Gebühren** = sum(Col 6 到 Col 11)

Sheet1 目标列：
- Col 5 = Brrutto_Umsatz
- Col 6 = Andere Gebühren

按 Nr 匹配，将计算结果写入 Sheet1 对应 OTTO 行，保存文件，输出汇总表格。

从当月银行流水 PDF（Kontodruckansicht）中提取汇款记录，写入当月 Umsatzsteuer 文件的 Sheet1（Nr、Datum、Zahlungsverkehr、Betrag 四列）。

## 所需文件

- **来源文件**：`D:\Workstation_Yuhong\BBG\电商\网买做账\Kontodruckansicht_<Monat>.pdf`
- **目标文件**：`D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_<Monat>.xlsx`
  - 如果用户在指令中指定了月份（如"四月"、"April"），自动推断对应文件名
  - 如果未指定，询问用户

---

## 月份与文件名对照

| 月份 | Umsatzsteuer 文件名 | Kontodruckansicht 文件名 |
|------|---------------------|--------------------------|
| 一月 / Januar    | Umsatzsteuer_Januar.xlsx    | Kontodruckansicht_Januar.pdf    |
| 二月 / Februar   | Umsatzsteuer_Februar.xlsx   | Kontodruckansicht_Februar.pdf   |
| 三月 / März      | Umsatzsteuer_März.xlsx      | Kontodruckansicht_März.pdf      |
| 四月 / April     | Umsatzsteuer_April.xlsx     | Kontodruckansicht_April.pdf     |
| 五月 / Mai       | Umsatzsteuer_Mai.xlsx       | Kontodruckansicht_Mai.pdf       |
| 六月 / Juni      | Umsatzsteuer_Juni.xlsx      | Kontodruckansicht_Juni.pdf      |
| 七月 / Juli      | Umsatzsteuer_Juli.xlsx      | Kontodruckansicht_Juli.pdf      |
| 八月 / August    | Umsatzsteuer_August.xlsx    | Kontodruckansicht_August.pdf    |
| 九月 / September | Umsatzsteuer_September.xlsx | Kontodruckansicht_September.pdf |
| 十月 / Oktober   | Umsatzsteuer_Oktober.xlsx   | Kontodruckansicht_Oktober.pdf   |
| 十一月 / November| Umsatzsteuer_November.xlsx  | Kontodruckansicht_November.pdf  |
| 十二月 / Dezember| Umsatzsteuer_Dezember.xlsx  | Kontodruckansicht_Dezember.pdf  |

---

## 执行步骤

### 第一步：确认文件路径

1. 根据月份推断两个文件路径（基目录：`D:\Workstation_Yuhong\BBG\电商\网买做账\`）
2. 确认两个文件都存在后继续
3. 提醒用户确保 Umsatzsteuer 文件未在 Excel 中打开

### 第二步：读取 PDF

用 Read 工具读取 PDF，从所有页面提取每一条汇款记录，跳过以下条目：
- `ENTGELTABSCHLUSS`（银行手续费结算，非汇款）

### 第三步：映射 Zahlungsverkehr

| PDF 中的付款方 | 写入值 |
|----------------|--------|
| OTTO Payments GmbH | OTTO |
| cflox GmbH | Kaufland |
| eBay S.a.r.l. | eBay |
| KTO 4059600017/... | KTO |
| IBAN PL7718800009... | IBAN |
| AMAZON PAYMENTS EUROPE S.C.A. | Amazon |
| XPOLLENS 110 AVENUE DE FRANCE | XPOLLENS |
| Orber Germany GmbH | Orber Germany GmbH |

如遇未知付款方，列出并询问用户如何处理。

### 第四步：写入 Sheet1

Sheet1 列结构：
- Col 1 = Nr（从 1 开始顺序编号）
- Col 2 = Datum（格式：YYYY.MM.DD，如 2026.04.30）
- Col 3 = Zahlungsverkehr（映射后的付款方）
- Col 4 = Betrag（金额，正数为入账，负数为出账，存储为数字）

**处理逻辑**：
- 如果 Sheet1 第一行为空，写入表头（Nr、Datum、Zahlungsverkehr、Betrag）
- 从第二行开始写数据，不覆盖 Col 5 及以后的已有内容
- 运行前确认文件未在 Excel 中打开，否则会报 PermissionError

**脚本路径**：直接生成内联 Python 脚本运行（无需保存为固定文件）；也可复用
`D:\Workstation_Yuhong\Yuhong's_File\Python\BBG_Ecommerce\fill_sheet1_from_pdf.py`
（每次运行前用新数据覆盖 TRANSACTIONS 列表）。

### 第五步：输出确认

汇报写入行数，并列出各 Zahlungsverkehr 的汇总（条数 + 合计金额），供用户核查。

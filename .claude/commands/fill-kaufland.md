将 Kaufland 各国站的销售数据填写到当月 Umsatzsteuer 文件的 Kaufland sheet，并在 Sheet1 中填入 Brrutto_Umsatz、Andere Gebühren、KTO_Guthaben、Behalten_in_KTO、Gebühren Ohne Steuer、Steuersatz 和 Marktplatz。

## 所需文件

- **来源文件（施总结算表）**：`D:\Workstation_Yuhong\BBG\电商\网买做账\Kaufland\<YYYY>年\<M>月\kaufland-BBG *.xlsx`
  - 包含各国站明细 sheet（德国站、奥地利站、法国站、斯洛伐克站、意大利站、捷克站、波兰站）
  - 以及 `合计应结算` sheet（含 PLN/CZK 月均汇率）
  - 非德国站：当月文件同时包含上月 sheet（如 `26年2月(奥地利站)` + `26年3月(奥地利站)`），脚本自动处理跨月尾单
- **上月来源文件（仅德国站需要）**：`D:\Workstation_Yuhong\BBG\电商\网买做账\Kaufland\<YYYY>年\<M-1>月\kaufland-BBG *.xlsx`
  - 德国站的跨月尾单（上月最后一笔 Payout 之后的交易）需从此文件读取
  - 如果不需要处理德国站的跨月尾单，可省略 `--prev-kaufland`
- **目标文件**：当月 Umsatzsteuer 文件，例如 `D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_März.xlsx`
  - 如果用户在指令中指定了月份（如"三月"、"April"），自动推断对应文件名
  - 如果未指定，询问用户
- **处理脚本**：`D:\Workstation_Yuhong\Yuhong's_File\Python\BBG_Ecommerce\fill_kaufland.py`

如果 openpyxl 未安装，先运行 `pip install openpyxl`。

---

## 月份与文件名对照

| 月份 | Umsatzsteuer 文件名 | Kaufland 目录 |
|------|---------------------|---------------|
| 一月 / Januar   | Umsatzsteuer_Januar.xlsx   | Kaufland\<YYYY>年\1月\ |
| 二月 / Februar  | Umsatzsteuer_Februar.xlsx  | Kaufland\<YYYY>年\2月\ |
| 三月 / März     | Umsatzsteuer_März.xlsx     | Kaufland\<YYYY>年\3月\ |
| 四月 / April    | Umsatzsteuer_April.xlsx    | Kaufland\<YYYY>年\4月\ |
| 五月 / Mai      | Umsatzsteuer_Mai.xlsx      | Kaufland\<YYYY>年\5月\ |
| 六月 / Juni     | Umsatzsteuer_Juni.xlsx     | Kaufland\<YYYY>年\6月\ |
| 七月 / Juli     | Umsatzsteuer_Juli.xlsx     | Kaufland\<YYYY>年\7月\ |
| 八月 / August   | Umsatzsteuer_August.xlsx   | Kaufland\<YYYY>年\8月\ |
| 九月 / September| Umsatzsteuer_September.xlsx| Kaufland\<YYYY>年\9月\ |
| 十月 / Oktober  | Umsatzsteuer_Oktober.xlsx  | Kaufland\<YYYY>年\10月\|
| 十一月 / November| Umsatzsteuer_November.xlsx| Kaufland\<YYYY>年\11月\|
| 十二月 / Dezember| Umsatzsteuer_Dezember.xlsx| Kaufland\<YYYY>年\12月\|

---

## 执行步骤

### 第一步：确认文件路径

1. 根据月份推断 Umsatzsteuer 文件路径（基目录：`D:\Workstation_Yuhong\BBG\电商\网买做账\`）
2. 在对应 Kaufland 目录下用 Glob 找 `kaufland-BBG *.xlsx`，取唯一匹配（若多个则取最新）
3. 在上月 Kaufland 目录下同样 Glob 找上月施总文件（`--prev-kaufland`）
4. 确认所有文件都存在后继续

### 第二步：运行脚本

```
python "D:\Workstation_Yuhong\Yuhong's_File\Python\BBG_Ecommerce\fill_kaufland.py" \
  --umsatz "<Umsatzsteuer文件路径>" \
  --kaufland "<当月Kaufland施总文件路径>" \
  --prev-kaufland "<上月Kaufland施总文件路径>"
```

**注意**：运行前确认 Umsatzsteuer 文件未在 Excel 中打开，否则会报 PermissionError。

### 第三步：解读输出并汇报

脚本输出分三部分：

**前置日志**
- `prev sheet '...'` 行：显示各国站从哪个 sheet 读取了跨月尾单及笔数
- `MATCH`：该笔 Payout 已成功匹配到 Sheet1 某行
- `SKIP`：该笔 Payout 在本月 Sheet1 中无对应条目（通常是月底提款，次月才到账）→ **正常，无需处理**

**汇总表格**（写入结果）
各列含义：
- `Brutto`：从原始订单数据聚合的毛销售额（已换算为 EUR）
- `Summe`：按汇率换算的实际到账 EUR 金额
- `S1_Betrag`：Sheet1 中的到账金额
- `S1_Diff`：S1_Betrag − Summe
  - EUR 国家应为 0.00
  - CZK/PLN 国家有小差值（日汇率 vs 月均汇率），属正常现象，无需处理
- `InnerChk`：内部一致性校验（所有字段之和 − Summe），应为 0.00

**最终行**：写入行数确认

### 第四步：异常处理

**如果有未匹配的 Sheet1 条目（`⚠ Unmatched Sheet1 entries`）**：
- 最常见原因：CZK/PLN 汇率偏差超过 6%（月均与当日汇率差异较大）
- 排查方法：检查 SKIP 列表中是否有金额接近该条目的 Payout，并与用户确认
- 不要自动忽略，需要用户确认归属后再处理

**如果 InnerChk 不为 0（超过 ±0.02）**：
- 可能是该国站数据结构有变化（新增了未识别的 booking_text 类型）
- 列出问题行，等待用户说明后修正脚本再重跑

**如果文件不存在**：
- Kaufland 施总文件：询问用户该月是否已收到施总的结算表，确认路径
- Umsatzsteuer 文件：询问用户是否已创建该月文件

---

## 背景知识

**数据关系**：
- Sheet1 中 `Zahlungsverkehr` 为 `Kaufland`/`Kaufland-DE`/`Kaufland-AT`/`Kaufland-FR` = EUR 国家（德/奥/法/斯洛伐克/意大利）到账
- `KTO` = 捷克（CZK，由 Kaufland 账户转出）
- `IBAN(Unbekannt)` = 波兰（PLN）

**各国特性**：
- **德国**：手动提款，每次保留部分余额（Behalten），该余额作为下次的 KTO_Guthaben 带入；德国上月尾单需从 `--prev-kaufland` 单独文件读取
- **其他国家**：每次提空（balance = 0），但月初可能有上月余额（来自上月末的销售）；上月 sheet 嵌入在当月施总文件中（如 `26年2月(奥地利站)`）
- **CZK/PLN 换算**：使用施总文件 `合计应结算` sheet 中的月均汇率，个别笔会有最多 6% 的日汇率偏差

**booking_text 分类规则**：
- `Freigabe Verkaufserlös` / `Release order` → Brutto_Sale（sum_price_gross）+ Commission（fee_gross，每笔销售抽成）
- `Erstattung` / `Partial Refund` / `Storno` → Erstattung(19%)
- `Fees for cancelled orders` / `Gebühren für stornierte` → **Steuer_Ohne_Gebühr**（0% 税率费用，单独列，无对应 Umsatzsteuer 行）
- `Netto X` 和 `Umsatzsteuer X`（同一费用拆成净额 + 税）→ **合并**进同一列：
  - X 含 `Grundgebühr` → Marktpaltzgebühr(19%)（月平台费）
  - X 含 `Gutschein` / `Voucher` → Gutscheingebühr(19%)
  - X 含 `Sponsored` 或其他 → Commission(19%)
- `Bezahlung Grundgebühr`（不含前缀）→ Marktpaltzgebühr(19%)

**Kaufland sheet 列结构**（13列）：
- Col 1=Nr, Col 2=Marktplatz, Col 3=Netto_Sale, Col 4=Ust, Col 5=Brutto_Sale,
  Col 6=Erstattung(19%), Col 7=Marktpaltzgebühr(19%), Col 8=Commission(19%),
  Col 9=Gutscheingebühr(19%), Col 10=Steuer_Ohne_Gebühr(19%),
  Col 11=KTO_Guthaben, Col 12=Behalten_in_KTO, Col 13=Summe

**Sheet1 目标列**：
- Col 5=Brrutto_Umsatz
- Col 6=Andere Gebühren（仅含 19% 税率费用：Erstattung + Marktpaltzgebühr + Commission + Gutscheingebühr，不含 Steuer_Ohne_Gebühr）
- Col 7=KTO_Guthaben, Col 8=Behalten_in_KTO
- Col 9（Einzahlung）不填写（平台充值，目前未发生）
- Col 10=Gebühren Ohne Steuer（0% 税率费用，如 Fees for cancelled orders）
- Col 11=Steuersatz（取自施总文件 fee_vat_% 列中 Freigabe 行的值，当前均为 0.19）
- Col 12=Marktplatz（对应国家站名称，如 Kaufland-DE、Kaufland-PL 等）

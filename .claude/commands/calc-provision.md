读取 `D:\Workstation_Yuhong\BBG\电商\网买做账\` 目录下所有 `Umsatzsteuer_*.xlsx` 文件的 Sheet1，汇总所有**进账**（Betrag > 0）金额，按 2% 计算佣金。

## 所需文件

- `D:\Workstation_Yuhong\BBG\电商\网买做账\Umsatzsteuer_*.xlsx`（所有存在的文件）

---

## 执行步骤

### 第一步：扫描文件

用 Python 列出目录下所有以 `Umsatzsteuer_` 开头、`.xlsx` 结尾的文件，按日期顺序排列（bis_Dezember 最早，其余按月份顺序）。

### 第二步：读取数据

用 `openpyxl`（`data_only=True`）打开每个文件，读取 Sheet1：
- 第一行为表头，找到包含 `Betrag` 的列
- 从第二行起遍历所有行
- **只统计正数**（> 0）的 Betrag 值作为进账；负数为转出，跳过

### 第三步：计算佣金

每个文件：
- `Eingang` = 该文件所有正数 Betrag 之和
- `Provision` = Eingang × 2%

全部文件汇总：
- `Gesamt Eingang` = 所有文件 Eingang 之和
- `Gesamt Provision` = Gesamt Eingang × 2%

### 第四步：输出结果

以表格形式展示每个月的进账与佣金，最后一行为合计，并单独标注总佣金金额：

```
Monat            Eingang (€)   Provision 2% (€)
------------------------------------------------
bis_Dezember       37,323.83             746.48
Januar             25,881.86             517.64
...
------------------------------------------------
GESAMT            xxx,xxx.xx           x,xxx.xx

截至目前共获得佣金：X,XXX.XX EUR
```

---

## 注意事项

- 使用 `sys.stdout.reconfigure(encoding='utf-8')` 避免中文路径编码报错
- 负数 Betrag 为转出（如转给 Itgoal Germany GmbH、Orber Germany GmbH 等），不计入佣金基数
- 如果某个文件 Sheet1 找不到 Betrag 列，报错并跳过该文件

# Project Memory

## 项目目标

这个项目用于把基站开站/扩容的 Excel 处理规则封装成可操作的前端小工具，减少人工改表。

## 当前架构

- 入口文件：`app.py`
- 开站模块：`kaizhan/app_SDR_FDD_gongxiang.py`
- 开站包装层：`kaizhan/page_wrapper.py`
- 扩容模块：`kuorong/app_sdr_expansion.py`

## 已确认的产品形态

- 工具主页提供单选入口：
  - `4g宏站开站`
  - `4g宏站扩容`
- 扩容工具支持：
  - `RANCM-sdrPlan` 参数模板导入
  - `RANCM` 网管导出表导入
  - 跨制式时额外导入 `cfgRadioNet` 文件

## 扩容模块关键规则

### 参数模板

- 主参数 sheet 名：`RANCM-sdrPlan`
- `RANCM-sdrPlan` 的表头在第 1 行
- 当前代码按第 4 行开始读取有效数据

### RANCM 目标表

- `RANCM` 目标表一般前 5 行是表头
- 有效数据通常从第 6 行开始

### 跨制式扩容

- `FDDtoTDD` 时额外导入 `RANCM-cfgRadioNet_TDD` 类文件
- `TDDtoFDD` 时额外导入 `RANCM-cfgRadioNet_FDD` 类文件

### 当前下载命名约定

- TDD 跨制式结果：`Result_RANCM-cfgRadioNet_TDD.xlsx`
- FDD 跨制式结果：`Result_RANCM-cfgRadioNet_FDD.xlsx`

## 已知项目事实

- 当前项目中真实存在的扩容模板文件：
  - `RANCM-sdrPlan4g宏站扩容模版.xlsx`
- 模板中已经确认存在：
  - `扩容类型`
  - `制式`
  - `connectModeWithUpRack`
  - `RUDevice`
  - `FDD/TDD`
- 模板中曾检查出可能缺少：
  - `LONGITUDE`
  - `LATITUDE`
  - `rfAppMode`
  这三个字段会影响 `Cell4GFDD` 分支

## 已知风险

- 有些字段在模板中会重复出现，例如 `refRfDevice`、`AntProfile`
- 当前扩容代码建索引时，重复表头可能以后一个覆盖前一个
- 终端里读取 Python 文件时出现过中文乱码，属于编码显示风险，需要谨慎判断

## 规则文档

- 扩容原始业务逻辑与代码核实结果已整理到 [EXPANSION_RULES.md](D:/codexapp/kuorong4g/pytool/EXPANSION_RULES.md)

## 后续建议

- 把扩容规则继续拆成更小的 sheet 处理函数
- 为每个 sheet 增加最小样例和回归检查
- 明确模板字段与代码字段的映射表，减少后续排查成本

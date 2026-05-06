

# 4g-tool



# 4G 宏站工具项目

这是一个基于 Streamlit 的 4G 宏站 Excel 自动化工具项目，当前包含两个主能力：

- `4G 宏站开站`
- `4G 宏站扩容`

项目入口是 [app.py](D:/codexapp/kuorong4g/pytool/app.py)，主页负责在两个工具之间切换。

## 目录结构

- `app.py`
  项目主页，负责选择开站或扩容工具。
- `kaizhan/`
  开站工具代码。
- `kuorong/`
  扩容工具代码。
- `RANCM.xlsx`
  网管导出模板文件。
- `RANCM-sdrPlan4g宏站填写模版.xlsx`
  开站参数模板。
- `RANCM-sdrPlan4g宏站扩容模版.xlsx`
  扩容参数模板。
- `cfgradioFDD.xlsx`
  FDD cfgRadioNet 模板。
- `cfgradioTDD.xlsx`
  TDD cfgRadioNet 模板。

## 运行方式

在项目根目录执行：

```powershell
streamlit run app.py
```

## 当前功能状态

### 开站

- 复用原有 `kaizhan/app_SDR_FDD_gongxiang.py` 逻辑。

### 扩容

- 支持导入 `RANCM-sdrPlan` 扩容参数模板和 `RANCM` 网管导出表。
- 当模板制式为跨制式时，额外导入对应 `cfgRadioNet` 文件。
- 支持根据扩容类型、制式、RRU 映射关系写入多个目标 sheet。

## 维护建议

- 修改 Excel 规则前，先核对 `RANCM-sdrPlan` 和目标模板 sheet 的真实表头。
- 不要轻易重命名中文 sheet 名、中文文件名、参数表头。
- 若发现终端中中文显示乱码，优先当作“终端编码显示问题”处理，不要直接假设源码或模板字段错误。

更多上下文见 [MEMORY.md](D:/codexapp/kuorong4g/pytool/MEMORY.md)、[AGENTS.md](D:/codexapp/kuorong4g/pytool/AGENTS.md) 和 [EXPANSION_RULES.md](D:/codexapp/kuorong4g/pytool/EXPANSION_RULES.md)。

> > > > > > > cb51ed8 (添加了agent、memory等文件，将项目维护成长期项目)

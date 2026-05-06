# 4G 宏站扩容规则文档

## 文档目的

这份文档用于记录 `4G 宏站扩容` 工具的原始业务逻辑，尽量保持与最初需求一致，并结合当前代码 [kuorong/app_sdr_expansion.py](D:/codexapp/kuorong4g/pytool/kuorong/app_sdr_expansion.py) 进行核实。

文档分为两层：

- `需求定义`
  按最初口述需求整理的目标规则。
- `代码核实`
  按当前代码实现逐项核对后的结果。

如果两者不一致，以“已知差异”明确标出，便于后续继续修正。

## 代码依据

当前核实主要参考以下函数：

- `parse_plan_workbook`：参数模板解析
- `get_plan_value_checked`：参数读取与缺失提示
- `prepare_template_rows`：目标 sheet 行数控制
- `process_rancm_expansion`：RANCM 主流程
- `process_cfg_radio_expansion`：cfgRadioNet 主流程

代码位置：

- [kuorong/app_sdr_expansion.py](D:/codexapp/kuorong4g/pytool/kuorong/app_sdr_expansion.py)

## 概念约定

### 输入文件

- 扩容参数模板：`RANCM-sdrPlan4g宏站扩容模版.xlsx`
- 网管导出表：`RANCM.xlsx`
- 跨制式附加文件：
  - `FDDtoTDD` 时导入 `RANCM-cfgRadioNet_TDD.xlsx`
  - `TDDtoFDD` 时导入 `RANCM-cfgRadioNet_FDD.xlsx`

### 参数模板结构

- 主参数 sheet：`RANCM-sdrPlan`
- 表头在第 `1` 行
- 代码当前按第 `4` 行开始读取有效数据

### RANCM 目标表结构

- 前 `5` 行通常视为表头或固定区
- 有效数据从第 `6` 行开始

### 分类字段

按当前项目约定，应这样理解：

- `扩容类型`：`软扩` / `硬扩`
- `制式`：`TDDtoTDD` / `FDDtoFDD` / `TDDtoFDD` / `FDDtoTDD`

代码核实：

- `扩容类型` 由 `EXPANSION_TYPE_ALIASES` 读取
- `制式` 由 `MODE_ALIASES` 读取，支持 `制式` 或 `FDD/TDD`

## 通用处理规则

### 1. 参数读取规则

需求定义：

- 模板某字段只有 1 行值时，可以按单值复用。
- 模板某字段需要多行值时，应按目标行一一对应读取。
- 如果源数据不足，应提示用户具体位置和参数。

代码核实：

- `get_plan_value_checked` 已实现该逻辑。
- 若只有 1 行值，默认允许复用。
- 若目标行超出源数据行数，会记录 `Sheet / 列 / 参数` 级别提示。

### 2. 目标行数规则

需求定义：

- 当某 sheet 规定有效数据为 `n` 行时：
  - 模板内有效数据不足 `n` 行：同列复制填充
  - 等于 `n` 行：不处理
  - 超过 `n` 行：删去多余行

代码核实：

- `prepare_template_rows` 已实现上述通用逻辑。
- 它按列独立判断有效数据长度，并执行复制或删减。

### 3. 样式和字体

代码核实：

- 大部分写入行会统一调用 `set_row_font(..., Times New Roman)`
- 部分 sheet 会先做行快照复制，再覆盖局部字段，尽量保留原样式

### 4. 标点清洗

代码核实：

- 最终保存前会对所有 sheet 执行一次全局标点清洗
- 目的是把中文标点替换为英文标点，避免导入网管报错

## 一、RANCM 扩容规则

### 1. Sheet `ManagedElement`

需求定义：

- 若 `制式` 为 `TDDtoTDD` 或 `FDDtoFDD`：不填写
- 若 `制式` 为 `TDDtoFDD` 或 `FDDtoTDD`：
  - 第 7 行 `A` 列 `MODIND` 填 `M`
  - 第 7 行 `F` 列填 `mimType`
  - 第 7 行 `G` 列填 `mimVersion`
  - 第 7 行 `H` 列填 `RADIOMODE`
  - 第 7 行 `I` 列填 `SWVERSION`
  - 第 7 行 `AI` 列与 `H` 列一致

代码核实：

- `fill_managed_element` 仅在跨制式时执行
- 代码保留第 6~7 行结构，最终写入目标行为第 `7` 行
- 写入列为 `A/F/G/H/I/AI`

已知差异：

- 无明显差异

### 2. Sheet `Equipment`

需求定义：

- 若 `扩容类型` 为 `软扩`：不填写
- 若 `扩容类型` 为 `硬扩`：
  - 保留前 5 行
  - 清除第 6 行开始所有数据
  - 有效数据为 1 行
  - `A` 列填 `M`
  - 从 `G` 到 `AI` 按表头（如 `Slot1` 到 `Slot12`）去模板中找同名字段，有值就写入

代码核实：

- `fill_equipment` 仅在 `硬扩` 时执行
- 代码将有效数据行控制为 `1` 行，从第 `6` 行写入
- `A6 = M`
- 从目标 sheet 第 1 行读取表头名，再去 `RANCM-sdrPlan` 中找同名字段写入

已知差异：

- 当前代码不是只限制 `Slot1` 到 `Slot12`
- 只要目标 sheet 第 1 行存在表头，且参数模板中有同名字段，就会尝试写入
- 因此如果模板和目标表包含 `Slot13`、`Slot14`、`Slot15` 等，也可能被写入

### 3. Sheet `RU`

需求定义：

- 若 `扩容类型` 为 `软扩`：不填写
- 若 `扩容类型` 为 `硬扩`：
  - 保留前 5 行
  - 行数等于模板中 `RUDevice` 的有效数据行数
  - `A` 列填 `A`
  - 先记录第 6 行 `B/C/D/E` 原始值
  - 清除第 6 行开始所有数据
  - `B/C/D/E` 按原始值填充
  - `F = RUDevice`
  - `G = userLabel1`
  - `H = RUType`
  - `K = RADIOMODE`
  - `L = functionMode`
  - `M` 与 `F` 一致
  - `N = connectModeWithUpRack`

代码核实：

- `fill_ru` 与上述规则一致
- 行数由 `RUDevice` 非空行数决定

已知差异：

- 无明显差异

### 4. Sheet `FiberDevice`

需求定义：

- 若 `扩容类型` 为 `软扩`：不填写
- 若 `扩容类型` 为 `硬扩`：
  - 行数由 `RUDevice` 在模板 `RRU` sheet 第 3 行匹配到的次数决定
  - 同列第 1 行值记为 `x`
  - `A = M`
  - `C/D/E` 从 `RANCM-sdrPlan` 按表头取 `SubNetwork / ManagedElement / NE_Name`
  - `F` 从 1 开始递增
  - `G = 1,1,x`
  - `H` 从 0 开始递增

代码核实：

- `build_fiber_device_entries` 会遍历每个 `RUDevice`
- 用 `find_rru_row3_matches` 只在 `RRU` 第 3 行找完全匹配
- 每找到一次，就取该列第 1 行作为 `slot`
- `fill_fiber_device` 最终写入：
  - `A = M`
  - `C = SubNetwork`
  - `D = ManagedElement`
  - `E = NE_Name`
  - `F = index + 1`
  - `G = 1,1,slot`
  - `H = index`

已知差异：

- 无明显差异

### 5. Sheet `FiberCable`

需求定义：

- 若 `扩容类型` 为 `软扩`：不填写
- 若 `扩容类型` 为 `硬扩`：
  - 行数等于 `RUDevice` 有效数据行数
  - 记录第 6 行 `B/C/D/E` 原始值
  - 清除第 6 行开始所有数据
  - `A = A`
  - `B/C/D/E` 按原始值填充
  - `F = RUDevice`
  - `G` 取值：
    - 若在 `RRU` 第 3 行找到：`(1,1,x):y`
    - 若在 `RRU` 第 3 行之外找到：`(z,1,1):2`
  - `H` 取值：
    - 需根据匹配位置的下一行是否有值判断不同写法
  - 双光纤规则：原需求里提过，但本项目已明确忽略

代码核实：

- `fill_fiber_cable` 已实现：
  - `A = A`
  - `B/C/D/E` 复用第 6 行原值
  - `F = RUDevice`
  - `G`：
    - 若匹配行是 `3`：写 `(1,1,x):y`
    - 若匹配行不是 `3`：写 `(z,1,1):2`
  - `H`：统一写 `({RUDevice},1,1):1`

已知差异：

- 当前代码没有实现“根据下一行是否有值，区分 `ref2FiberDevice` 写法”的分支
- 当前代码已按你的要求忽略双光纤扩展逻辑

### 6. Sheet `IrAntGroup`

需求定义：

- 仅在“`硬扩` 且目标制式为 TDD”时填写
- 对应场景：
  - `TDDtoTDD`
  - `FDDtoTDD`
- 行数等于 `RUDevice` 有效数据行数
- 记录第 6 行原始值
- 清除第 6 行开始所有数据
- `A = A`
- `B/C/D/E` 复用原始值
- `F = IrAntGroup` 编号递增
- `G = 1`
- `H = RUDevice`
- `I = antEntityNo` 递增
- `J = refRfDevice`
- `K = AntProfile`

代码核实：

- `fill_ir_ant_group` 仅在：
  - `is_hard_expansion(plan_data)` 为真
  - `mode_key` 属于 `TDD_TARGET_MODE_KEYS`
  时执行
- 写入逻辑与需求一致

已知差异：

- 无明显差异

### 7. Sheet `IpLayerConfig`

需求定义：

- 仅在跨制式时填写：
  - `TDDtoFDD`
  - `FDDtoTDD`
- 前 5 行不改
- 有效数据 2 行
- 从第 6 行开始处理
- `A = A`
- `B/C/D/E` 保留第 6 行原始值
- `J = vid`
- `M = ipAddr`
- `N = networkMask`
- `O = gatewayIp`

代码核实：

- `fill_ip_layer_config` 仅在跨制式时执行
- 目标行数固定为 2
- 会先复制第 6、7 行快照，再覆盖 `A/B/C/D/E/J/M/N/O`

已知差异：

- 你的原描述里写过“清除 sheet `IrAntGroup` 第六行开始所有数据”，这里应理解为笔误
- 当前代码实际处理的是 `IpLayerConfig` 自身，没有动 `IrAntGroup`

### 8. Sheet `Sctp`

需求定义：

- 仅在跨制式时填写
- 前 15 行不改
- 新增 10 行有效数据
- `G = Sctp` 从第 15 行开始递增
- `H = sctpNo` 从第 15 行开始递增
- 其余列复制 6~15 行数据填充

代码核实：

- `fill_sctp` 仅在跨制式时调用
- 有效总行数被控制为 `20` 行，即第 6~25 行
- 会复制第 6~15 行作为模板，写入第 16~25 行
- `G/H` 从第 15 行的值开始逐行递增

已知差异：

- 无明显差异

### 9. Sheet `ServiceMap`

需求定义：

- 仅在跨制式时填写
- 有效数据 2 行
- 从第 8 行开始写
- `F = serviceMapNo` 从第 7 行递增
- `O = fddServiceDscpMap`
- `P = tddServiceDscpMap`
- 其余列复制模板行

代码核实：

- `fill_service_map` 仅在跨制式时执行
- 会保留并整理第 6~9 行结构
- 用第 6、7 行做快照
- 写入目标行是第 `8`、`9` 行
- `F` 从第 7 行的值开始递增
- `O/P` 按模板参数字段写入

已知差异：

- 代码是“复制第 6、7 行到第 8、9 行”
- 不是“复制第 6~15 行”
- 对于你定义的 2 行结果来说，当前实现是够用的，但描述上与原口述不完全一致

## 二、cfgRadioNet 规则

### 触发条件

需求定义：

- 当模板 `制式` 为跨制式：
  - `FDDtoTDD`
  - `TDDtoFDD`
  时，用户需要额外导入 `cfgRadioNet` 文件

代码核实：

- `render_expansion_page` 会在识别到跨制式后，显示第三个上传入口
- 文件映射如下：
  - `FDDtoTDD -> RANCM-cfgRadioNet_TDD.xlsx`
  - `TDDtoFDD -> RANCM-cfgRadioNet_FDD.xlsx`

### 1. Sheet `ENBFunction`

需求定义：

- 跨制式时填写
- 需要新增 1 行
- 从第 7 行开始追加
- `C/D/E` 取第 6 行同列值
- `H/I` 与 `E` 相同
- `J/L` 与 `D` 相同
- 其余列复制模板行

代码核实：

- `fill_cfg_enbfunction` 会：
  - 统计现有有效行数
  - 至少追加 1 行
  - 基于第 6 行做快照复制
  - 对 `C/D/E/H/I/J/L` 做指定覆盖

已知差异：

- 当前实现是“追加到现有数据末尾”
- 如果 `ENBFunction` 现有有效数据不止 1 行，新增行不一定正好是第 7 行

### 2. Sheet `Cell4GTDD`

适用场景：

- `FDDtoTDD`

需求定义：

- 写入 `Cell4GTDD`
- 行数等于 `cellnum`
- 先统计现有 `A` 列有效行数为 `n`
- 从第 `n+1` 行开始写
- `A = A`
- `H = moId`
- `I = cellLocalId`
- `L = userLabel2`
- `Q = pci`
- `U = freqBandInd`
- `V = earfcnUl`
- `W = bandWidthDl`
- `Y = bandWidthUl`
- `AE = rootSequenceIndex`
- `AF = cpMoId` 递增
- `AI = refIrAntGroup`
- `AJ = refBpDevice2`
- `AK = cellMod`
- `AL = cpSpeRefSigPwr`
- `AM = upActAntBitmapSeq`
- `AN = anttoPortMap`
- `AO = isDelNbrAndRelation`

代码核实：

- `fill_cfg_tdd_cells` 与上述字段映射基本一致
- 行数由 `cellnum` 决定
- 起始行为 `6 + existing_count`
- `AF` 按上一条现有记录的 `AF` 值递增

关于 `AI = refIrAntGroup`：

- 当前代码不是直接读参数模板
- 它会先读取“刚生成好的 RANCM 结果”中的 `IrAntGroup` sheet
- 只提取 `irAntGroupNo`
- 再写成 `{irAntGroupNo}:1`

已知差异：

- 你最初的描述里对 `AI` 的来源有过两种表述：
  - 一处像是 `ref1SdrDeviceGroup`
  - 一处又写成读取 `irAntGroupNo`
- 当前代码采用的是后者，即 `irAntGroupNo:1`

### 3. Sheet `Cell4GFDD`

适用场景：

- `TDDtoFDD`

需求定义：

- 写入 `Cell4GFDD`
- 行数等于 `cellnum`
- 先统计现有 `A` 列有效行数为 `n`
- 从第 `n+1` 行开始写
- `A = A`
- `H = moId`
- `I = cellLocalId`
- `L = userLabel2`
- `Q = pci`
- `S = tac`
- `U` 到 `AA` 按目标表头去参数模板找同名字段
- `AF = LONGITUDE`
- `AG = LATITUDE`
- `AH = rootSequenceIndex`
- `AI = H`
- `AK = rfAppMode`
- `AL` 到 `AM` 按目标表头找同名字段
- `AP` 到 `AV` 按目标表头找同名字段

代码核实：

- `fill_cfg_fdd_cells` 与上述规则一致
- `U~AA`、`AL~AM`、`AP~AV` 都是“按目标表头名去参数模板找同名字段”

已知差异：

- 当前参数模板如果缺少 `LONGITUDE`、`LATITUDE`、`rfAppMode`，这一分支会产生缺失提示

## 三、已明确忽略的规则

以下内容已根据你的后续指示忽略，不作为当前实现目标：

- 双光纤相关扩展逻辑

## 四、当前代码与原始需求的主要差异汇总

### 已基本一致

- `ManagedElement`
- `RU`
- `FiberDevice`
- `IrAntGroup`
- `IpLayerConfig`
- `Sctp`
- `Cell4GFDD`
- `Cell4GTDD`

### 存在差异或简化

- `Equipment`
  代码按目标 sheet 第 1 行的真实表头动态写入，不只限 `Slot1~Slot12`
- `FiberCable`
  `H` 列没有实现“根据下一行是否有值区分写法”的分支
- `ServiceMap`
  当前代码只复制第 6、7 行到第 8、9 行，不是按“6~15 行”理解
- `ENBFunction`
  当前是“追加到现有有效数据末尾”，不绝对固定在第 7 行

## 五、后续维护建议

1. 如果以后继续加规则，优先在本文件中先补“目标规则”，再落代码。
2. 如果修改了 [kuorong/app_sdr_expansion.py](D:/codexapp/kuorong4g/pytool/kuorong/app_sdr_expansion.py)，要同步更新本文件的“代码核实”和“差异汇总”部分。
3. 若发现模板字段与代码字段不一致，先查参数模板真实表头，不要直接依据终端乱码做判断。

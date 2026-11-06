# 同格双语翻译：口径与规则

本文件是 `excel-bilingual` 的判断依据。SKILL.md 只讲流程，这里讲"哪种格该译、怎么译、什么不动"。

## 1 输出形态

- 每个待译格最终内容 = `原文` + `\n` + `译文`。原文部分逐字保留（含原有换行、行内空格、单位大小写），只在末尾追加一行中文。
- 新增行、新增列、新增 sheet 一律不做。原表行列坐标、合并区、表头层级保持不变。
- 允许且只允许的版式改动：给写入双语的单元格置 `wrap_text=True`，并按内容加高所在行。

## 2 文本匹配的三级归一（最容易踩的坑）

词典键与表内文本都要经过：**Unicode NFC 规范化 → 转小写 → 折叠连续空白**。

- 越南语常见分解式组合符（`chuẩn nén` 这类 NFD 文本），不做 NFC 会一个字都对不上。
- 原表大量全大写行文（`CÁI`、`VẬT LIỆU PHỤ, NHÂN CÔNG LẮP RÁP HOÀN THIỆN`），不折叠大小写命中率会低到 3% 左右。
- 同一原文同时以"仅规范化+小写"和"折叠空白+小写"两种键入库，多行文本两种形态都能命中。

## 3 公式格的处理（必须先判断再决定）

逐格判断 `=` 开头的单元格：

- **缓存值是文本**（`data_only=True` 视图里同坐标是字符串）：这类格才是真正藏着待译文字的地方，常见于跨簿外链公式 `='[8]Công trình'!F13`。**同格双语与保公式不可兼得**，一律转静态文本；否则换机或原簿不在原路径时整列变 `#REF!`。
- **缓存值是数值**（金额=数量×单价、合计、SUBTOTAL、取整）：保留公式，一格不动。
- 交付时必须说明"多少格由公式转成了静态文本"，并说明外部链接定义仍留在文件里。

## 4 不翻译的清单（按原样保留）

这些内容硬加中文会干扰选型、对图、报量：

| 类型 | 例子 |
| --- | --- |
| 定额/清单编号 | `BA.18302`、`BB.86103`、`D4.2101`、`BA.1850` |
| 型号与规格串 | `MCB-1P-10A-6KA`、`MCCB 3P 2500A 85kA`、`RCBO-2P-20A-6KA-30mA`、`D150`、`1200x200`、`27U D1000` |
| 品牌型号串 | `DS-2CD1343G2-LIU`、`NPC06UVDB-BL007F`、`CPP-UDDM-SL-1U-24`、`C3KS-LCD` |
| 章节序号/占位 | `TT`、`A.1`、`B1`、`I~IV`、`T1`/`T2`/`TM`、`#REF!` |

纯厂家名单独成格时保留原音并译国别，例如 `Tôn Hoa sen/Việt Nam` → `Hoa Sen钢板/越南`、`ORD/Trung Quốc` → `ORD/中国`、`Sino/VN` → `Sino/越南`；`Toàn Phát`、`Dekko`、`Tai Sin`、`Commscope` 等品牌名保留。

## 5 尺寸+厚度类规格串

`D41.3 DÀY 1.0MM`、`150X150,ĐỘ DÀY 0.5MM` 这类整格只是"尺寸 + 厚度"的，用 `assets/rules/vi-spec-rules.json` 的正则统一转成 `D41.3，厚度1.0mm`。

**不要**为长句写"半句翻译"的正则：只译半句比不译更糟。规则必须整格可判定时才用。

## 6 单位对照

`CÁI` 个 / `BỘ` 套 / `LÔ` 批 / `M`、`MÉT` 米 / `M2` 平方米 / `M3` 立方米 / `SỢI`、`CỌT`、`CỌC` 根 / `HỘP` 盒 / `LỌ` 瓶 / `BAO` 袋 / `CỤM` 组 / `CẶP` 对 / `MÁY`、`TỦ` 台 / `BỂ` 池 / `KIM` 支 / `GÓI` 项 / `BẢNG` 面 / `ĐẦU` 端 / `TRẠM` 站

## 7 行业术语口径（机电 MEP）

器件：MCB 小型断路器 / MCCB 塑壳断路器 / ACB 框架断路器 / RCBO 漏电保护断路器 / VCB 真空断路器 / LBS 负荷开关 / RMU 环网柜 / CT 电流互感器 / VRV 多联机（`dàn lạnh` 室内机、`dàn nóng` 室外机、`bộ chia gas` 分流器）/ `TĐ` 配电柜 / `BĐ` 配电箱 / `kiosk` 一体化站壳。

系统词：`Chống sét` 防雷 / `ELV` 弱电 / `CTN`、`TBVS` 给排水、卫浴设备 / `ĐHKK` 空调通风 / `thanh cái` 母排 / `khoang` 隔室 / `ngăn` 隔间 / `sấy nhiệt` 加热除湿 / `vỏ tủ` 柜体 / `đầu vào`、`đầu ra` **进线、出线**（配电柜语境，不用"输入/输出"）/ `vỏ tủ loại N module` 暗装 N 模数配电箱箱体。

明细：`mương cáp` 电缆槽 / `thang cáp` 电缆桥架 / `hào cáp` 电缆沟 / `ống luồn` 穿线管 / `kẹp đỡ ống` 管支撑卡件 / `bảo ôn` 保温 / `chống ồn` 隔声 / `tiêu âm ống gió` 风管消声器 / `chân rẽ` 分支管件 / `miệng gió`、`cửa gió` 风口、送风口 / `van cổng` 闸阀 / `van 1 chiều` 止回阀 / `van phao cơ` 机械浮球阀 / `khớp nối mềm` 软接头 / `y lọc` Y 型过滤器 / `rọ hút` 吸水过滤器 / `bình tích áp` 稳压罐 / `hố ga thăm` 检查井 / `thoát sàn` 地漏 / `chụp thông hơi` 通气帽 / `hàn hóa nhiệt` 放热焊接 / `cọc thép mạ đồng` 镀铜钢接地棒。

完整词表见 `assets/glossary/vi-zh-mep.json`（553 条，键为越南语小写，值可含换行；运行时展开成约 591 个键形态）。

## 8 行高与列宽

行高按"列宽（合并区按跨列宽度累加）× 字符宽度估算"计算：中文按 2 个宽度单位、西文按 1 个，行数乘以字号行高（`font.size * 1.45`，下限 12）再加 4pt 余量，上限 409（Excel 硬上限）。**只加高不压低**：原行高更大时保留原值。

## 9 校验与交付

- 逐格比对（脚本 `scripts/verify_bilingual.py`）：双语格必须严格等于"原文 + 分隔符 + 译文"；双语格之外的数值与公式零改动；合并区、autoFilter、条件格式规则数、`dimensions` 与原文件一致；`xl/externalLinks` part 数不减。校验 COM 产物时加 `--allow-com-artifacts`（只放行外链序号重排、CR 转义、首尾空格规范化这三类，见第 13.2 节），其余差异照旧判 FAIL。
- 有 Excel 时实际打开一次确认（8 个表、表名、单元格内容、行高）。**用 ASCII 文件名副本**，中文绝对路径在 PowerShell 传参时会被编码破坏；导出 PDF/截图做目视容易被弹窗卡成 `RPC_E_CALL_REJECTED`，不要为此反复重试。
- 源文件绝不被改写：新产物另存，交付新文件；工作表**标签名保持原语言**（表内公式按表名跨表引用，改名即断链）。
- 交付说明必须写清：双语格数、公式转静态文本的格数、未翻译保留的格数与类型、结构校验结论。

## 10 译前体检：文件能不能交给 openpyxl 存

openpyxl 只认识它建模过的那部分 OOXML，保存时会把不认识的对象整块丢掉（不是"效果差不多"，是直接没了）。所以动手前先读 zip 里的 part 清单，而不是靠肉眼翻页：

```bash
python scripts/audit_features.py --source ./source.xlsx
```

它统计两类东西：

- **zip part 级**：`drawings`（并进一步数 drawing XML 里的 `<xdr:sp>` 文本框个数）、`charts`、`pivotTables` / `pivotCache`、`slicers`、`ctrlProps`、`activeX`、`vbaProject.bin`、`threadedComments`、`richData`、`media`、`xl/externalLinks/externalLink*.xml`。
- **openpyxl 侧**：合并区数、条件格式规则数、数据验证、批注、定义名称、隐藏/保护工作表。

判定口径：

| 结论 | 触发条件 | 处理 |
| --- | --- | --- |
| `OPENPYXL_OK` | 上面列的对象一个都没有 | 正常走 openpyxl 写入 |
| `OPENPYXL_OK_WITH_CAUTION` | 只有批注/富数据这类"可能被规范化但语义不丢"的对象 | 可写入，但交付说明列出受影响对象 |
| `RISKY` | 有文本框、图表、透视、切片器、控件、VBA、x14 数据验证 | 不要用 openpyxl 存；改走 COM 写入器，或只交可保真部分 |

实测参考：一份 8 sheet 的越南语机电工程量清单（50 个 zip part、99 个合并区、19 条条件格式规则、11 个定义名称、12 条外链、脆弱 part 全无）体检结果是 `OPENPYXL_OK`。这份文件能全量重存而结构不丢，是体检结论给的底气，不是运气。

## 11 词表沉淀：从成品反抽，别重复劳动

同类文件往往只隔几个月，上一份的译法就是这一份的标准答案。`scripts/mine_glossary.py` 从已译好的双语簿反抽 `原文→译文`：

- 边界判定：从第 2 行起，找**第一行含 CJK 且不含越南语/拉丁变音符**的行作为译文起点（原文与译文都可能多行，不能按固定行数切）。
- 键必须过与写入脚本同一套三级归一：否则 `150X150，厚度0.5mm` 与 `150x150，厚度0.5mm` 会被当成两个词条，词表无谓翻倍、重跑还会产生格子差异。
- `--rules` 把已由 `vi-spec-rules.json` 覆盖的规格串丢掉：这些格由正则确定性生成，入表只会让同一原文出现两种写法（`150X150` vs `150x150`）。
- 输出里的 `CONFLICT` 是"同一原文两种译法"：先入库者胜，剩下的需人工定夺，不要盲信"后写覆盖"。

## 12 文本框与浮动对象里的文字

工程量清单常把说明塞在文本框或合并的大标题里，这些文字在 `sheet.next` 的单元格迭代里根本看不到。COM 写入器按 `Shapes` 遍历，用 `TextFrame2.TextRange.Text` 取/写，命中词表就做同格双语（文本框内换行照旧，样式不丢）。
注意两点：形状里的文字不能靠 `Shape.TextFrame.Characters` 拍脑袋拼接；同一次运行里形状文字与单元格文字用同一份词表与同一套归一，否则同一句原文会出现两种译法。

## 13 保真路线的边界（COM 退路）

- 本机 Python 侧没有 `win32com` / `xlwings`，COM 只能走 PowerShell 的 `New-Object -ComObject Excel.Application`（实测 Excel 15.0，ProgID `Excel.Application`）。
- Excel 有可见窗口在跑时，自动化调用会被 `RPC_E_CALL_REJECTED (0x80010001)` 直接拒绝，重试多次也过不去；`scripts/apply_bilingual_com.ps1` 因此在检测到可见窗口时以退出码 2 直接拒绝，要求先关掉 Excel。**失败运行留下的隐藏 `EXCEL.EXE` 孤儿会毒化 ROT**：`Get-Process EXCEL | Where MainWindowTitle -eq ''` 找到的就是它们，`Stop-Process` 清掉后同一份文件立刻能开。
- COM 只改工作副本（先 `Copy-Item` 再 `Workbooks.Open`），失败时关闭而不保存，绢不半途写回源文件。
- 中文绝对路径经 PowerShell 传参会变成 mojibake（Excel 报"抱歉，无法找到 …xlsx"），一律用 ASCII 文件名副本。
- 合并行无法 `AutoFit`，脚本按"原高×2（上限 409）"近似；非合并行才真 `AutoFit`。
- 从开源实现借来的另外几条工程习惯，遇同类问题照用：批内一次请求/一次遍历、按语言对缓存避免重复翻译、按字符脚本判语种（zh/ja/ko/ru/ar/he/th/hi）、大簿按 N 格增量保存防丢、翻译前先备份、失败重试用指数退避且单格失败只影响单格。

### 13.1 已实测跑通（8 张表 1654 格）

在本机对一份 8 表越南语机电清单整簿跑过一次并通过校验：`cells updated: 1654`、`formula cells turned into static text: 948`、`cell errors skipped: 0`、`untranslated unique: 225 cells: 689`，退出码 0，整簿耗时约 4~5 分钟。用 `verify_bilingual.py --allow-com-artifacts` 对照原文件：`bad cells: 0`、`numeric cells changed: 0`、8 张表的合并区/autoFilter/条件格式/dimensions 全 `OK`、externalLink part 数不减 → `RESULT: PASS`。

跑之前必须做对的三件事：

1. 关掉所有可见 Excel 窗口，并清掉隐藏孤儿进程。
2. 路径全用 ASCII 副本；长命令会被自动转后台，用 `TaskOutput` 取输出。轮询超时是正常慢，不是卡死。
3. 脚本先探活实例（`$xl.Version` 最多 20 次 × 1 秒），`Workbooks.Open` 最多 6 次；仍失败就换新实例重来，最多 3 个实例。Excel 2013 启动慢到连 `Workbooks.Open` 都拒时，这一步是能不能跑起来的分水岭。

### 13.2 COM 与 openpyxl 的已知差异（用 COM 就必然出现，可接受但必须如实说明）

| 差异 | 表现 | 判定 |
| --- | --- | --- |
| 外链序号重排 | `='[9]CT1A'!$E$22` 存盘后变 `='[11]…`，`=[5]DHKK!K106` 变 `=[2]…`（本表 2067 格） | 语义等价：链接定义仍在、externalLink part 数不变，只是 Excel 重新编号。`--allow-com-artifacts` 会单独计数并放行 |
| CR 变字面量 | 源格 `\r\n` 存盘后成 `_x000D_\n`（本表 5 格，如各表 `M4` 的"Tổng cộng(đ)"） | 显示效果一致，属于 Excel 对 CR 的转义；`--allow-com-artifacts` 按等价处理 |
| 首尾空格被规范 | 源格原文两侧带空格（本表 47 格，多为跨簿外链缓存文本 ` MÉT `），写回后首尾空格丢失 | 纯排版差异，文字与译文都不受影响 |
| 未译计数 | COM 侧 225 种 / 689 格，openpyxl 侧 226 种 / 690 格 | 差 1 格，来自上述空白规范化让某个键恰好命中 |

所以 COM 产物**不能**拿 openpyxl 路径的"逐字相等"标准去卡，必须加 `--allow-com-artifacts`；出现这三类以外的差异就是真问题，不许放行。反过来，正文内容（译文、公式转静态的格数与位置、结构）在两条路径上是一致的：同一份源文件，COM 版与 openpyxl 交付件逐格比对，除上述三类外无实质差异。

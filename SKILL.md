---
name: excel-bilingual
description: 把外语（越南语/英语/日语等）Excel 工程量清单、报价表、BOM 就地翻成"同一单元格上原文下中文"的双语对照，不新增行列、不改版式，同时保住数值、真实公式、合并单元格、自动筛选、条件格式与外部链接；译前先体检文件里的文本框/图表/透视/控件/VBA 会不会被 openpyxl 丢弃，并可复用上一份成品的词表。适用触发：用户给出外语 .xlsx/.xls/.csv 并要求"翻译成中越对照 / 中文对照""内容在原方格内翻译""上面原文下面中文""不要新增列""按原表格内容翻译不能加戏"，或要求沿用既有双语表的译法继续做同类文件。
---

# Excel 同格双语翻译

把译文写进原单元格内部，原文在上、译文在下。这是本 Skill 唯一认可的形态：**不新增列、不新增行、不建镜像 sheet**。

## 流程

### 1 译前体检、备份、看清结构

先判这份文件能不能交给 openpyxl 重存（它在保存时会把不认识的对象整块丢掉），再谈翻译：

```bash
python scripts/audit_features.py --source ./source.xlsx
```

末行给结论：

| 结论 | 含义 | 下一步 |
| --- | --- | --- |
| `OPENPYXL_OK` | 无 drawings/charts/pivot/slicers/控件/VBA/x14 验证/批注等脆弱 part | 走第 3 节的 openpyxl 写入 |
| `OPENPYXL_OK_WITH_CAUTION` | 只有嵌入图片、批注这类"会被重新序列化但语义大体还在"的对象 | 照常写入，但交付说明列出"可能被规范化"的对象 |
| `RISKY` | 有文本框/图表/透视/切片器/控件/VBA | **不要**用 openpyxl 写，改走第 3 节的 COM 写入器，并在交付说明写清取舍 |

然后扫待译文本：

```bash
python scripts/scan_text.py --source ./source.xlsx --out ./unique_strings.tsv
```

输出每个 sheet 的待译格数（并单列"公式缓存文本"格数）、唯一原文清单（带出现次数）。
同时确认：sheet 数量与标签名、合并区、autoFilter、条件格式、外部链接是否存在——这些后面要逐项核对。
动手写之前先复制一份 `source` 备份（openpyxl 的改动不可回滚）。

### 2 建词典，不要逐格手打

- 有既有词表就复用：`assets/glossary/vi-zh-mep.json`（553 条机电 MEP 越南语→中文，含单位、器件、系统词）。
- **上一份同类文件已译好时，直接从成品反抽词对**，比重新翻省事且口径必然一致：

```bash
python scripts/mine_glossary.py --bilingual ./前一份中越对照.xlsx \
  --out ./glossary_new.json \
  --merge assets/glossary/vi-zh-mep.json \
  --rules assets/rules/vi-spec-rules.json
```

  `--merge` 会按三级归一后的键去重（大小写、空白差异不会重复入库）；`--rules` 把已由正则覆盖的规格串丢弃，避免同一个 `150X150,ĐỘ DÀY 0.5MM` 出现两种译法。脚本会打印 `CONFLICT`：同一原文两种译法时以先入库者为准，需人工定夺。
- 新领域先按 `unique_strings.tsv` 的高频项补词表，键用小写、值可含换行。
- 长句按整句入表；**不要**只给词组、指望拼接出通顺中文。

匹配口径必须按 `references/terminology-rules.md` 第 2 节做三级归一（NFC → 小写 → 折叠空白），否则全大写行文会大面积漏译。

### 3 写入

```bash
python scripts/apply_bilingual.py --source ./source.xlsx --out ./out.xlsx \
  --glossary assets/glossary/vi-zh-mep.json \
  --rules assets/rules/vi-spec-rules.json \
  --missing ./missing.txt
```

脚本行为：`原文.rstrip() + "\n" + 译文` 覆盖同格；该格置自动换行；按列宽（合并区按跨列宽）估算行数并只加高不压低行高（上限 409）；未命中的原文写进 `--missing`。
`--glossary` 可重复传，也可传目录（读其中全部 `*.json`）；`--sheets` 可限定范围；`--no-height` 跳过行高调整。

**译前体检是 `RISKY` 时**改走 COM 写入器（在 Excel 进程内改，故文本框/图表/透视/控件/VBA 全部保留；代价是必须关掉已打开的 Excel 窗口，且慢很多）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/apply_bilingual_com.ps1 `
  -Source ./source.xlsx -Out ./out.xlsx `
  -Glossary assets/glossary/vi-zh-mep.json -Rules assets/rules/vi-spec-rules.json
```

它做的事与 openpyxl 版一致（同一套三级归一、同样只写"原文+换行+译文"、公式缓存为文本的格转静态），并额外遍历 `Shapes` 用 `TextFrame2` 把文本框文字也做成双语（openpyxl 看不到这些文字）。规则文件里的 `\1`/`\2` 会先转成 .NET 的 `$1`/`$2` 再回填，否则会把反斜杠原样写进格子。

传入的必须是 ASCII 文件名路径（中文路径在 PowerShell 传参时会被编码破坏）；有可见 Excel 窗口时脚本会直接以退出码 2 拒绝，`-AllowInteractive` 也不保证能过（`RPC_E_CALL_REJECTED`）。跑之前先清掉隐藏孤儿进程（`Get-Process EXCEL | Where-Object { $_.MainWindowTitle -eq '' }` 找到的就是它们），脚本内部还会探活实例、失败换实例重来（最多 3 个）。

整簿实跑耗时约 4~5 分钟，长命令会被自动转后台，用 `TaskOutput` 取输出（轮询超时是正常慢，不是卡死）。跑完后**必须**用 `verify_bilingual.py --allow-com-artifacts` 校验（该开关只放行 COM 必然产生的三类痕迹：外链序号重排、CR 转成 `_x000D_`、原文首尾空格被规范化；见 `references/terminology-rules.md` 第 13.2 节），这三类以外的差异一律不许放过。

### 4 校验（必做，不可只看"跑完了"）

```bash
python scripts/verify_bilingual.py --source ./source.xlsx --out ./out.xlsx
# COM 产物加开关：只放行外链序号重排 / CR 转义 / 首尾空格规范化
python scripts/verify_bilingual.py --source ./source.xlsx --out ./out.xlsx --allow-com-artifacts
```

必须 `RESULT: PASS`：双语格严格等于"原文+分隔符+译文"、双语格外数值与公式零改动、合并/筛选/条件格式/dimensions 与原表一致、externalLinks 不减。
有 Excel 时再用 ASCII 文件名副本实际打开一次（见 `references/terminology-rules.md` 第 9 节）。

### 5 交付

- 新文件另存，源文件不动；文件名用"原名 + 中越对照"，放在用户项目目录。
- 通过宿主的产物交付工具（`present_files`）交付路径，不在正文贴裸路径链接。
- 交付说明必须列：译前体检结论（`OPENPYXL_OK` 还是走了 COM）、双语格数、由公式转静态文本的格数、未翻译保留的格数与类型、结构校验结论、工作表标签名为何保留原文。

## 硬规则

1. **同格双语，不新增行列**。原文逐字保留（含原有换行与行内空格），只追加一行译文。
2. **译文不得增补原文没有的信息**。原文没写的功能、参数、缩写、修饰语一律不加；宁可直译也不"补全"。
3. **公式分两类**：缓存值是文本的外链公式格转静态双语文本（同格双语与保公式不可兼得）；缓存值是数值的真实计算公式一格不动。转静态的格数必须写进交付说明。
4. **编号、型号、规格串、章节序号不译**（清单见 `references/terminology-rules.md` 第 4 节），交付时列出这批"保留原文"的格数与类型，让用户决定是否补注释。
5. **工作表标签名保持原语言**：表内公式按表名跨表引用，改名即断链。
6. **只加高不压低行高**；不给未改动的格动样式、不动列宽、不重刷底色边框。
7. 术语按 `references/terminology-rules.md` 第 6–7 节统一；同一原文在同一文件里只能有一种译法。
8. **译前体检不过就不许用 openpyxl 存**。`RISKY` 的文件走 COM 写入器或只交"能保真的部分"，绝不"先存一遍再补"——丢掉的图表/控件不会回来。
9. **批量与缓存**：一份词表 + 正则覆盖整簿，不为单格现编译文；同批同类文件先跑 `mine_glossary.py` 复用上一份的译法，译文口径必须跨文件一致。

## 资源

| 路径 | 用途 |
| --- | --- |
| `scripts/audit_features.py` | 译前体检：判 openpyxl 会不会丢对象，给 `OPENPYXL_OK/_WITH_CAUTION/RISKY` |
| `scripts/scan_text.py` | 扫描待译文本，产出唯一原文清单 |
| `scripts/mine_glossary.py` | 从已译成品反抽"原文→译文"，按三级归一去重后并入词表；`--rules` 丢弃正则已覆盖的规格串 |
| `scripts/apply_bilingual.py` | 同格双语写入 + 自动换行 + 行高（openpyxl，主路径） |
| `scripts/verify_bilingual.py` | 逐格与结构校验，PASS/FAIL；`--allow-com-artifacts` 用于 COM 产物 |
| `scripts/apply_bilingual_com.ps1` | 同格双语写入（Excel COM，`RISKY` 文件的退路；需关闭可见 Excel 窗口、清过孤儿进程；整簿约 4~5 分钟） |
| `assets/glossary/vi-zh-mep.json` | 机电 MEP 越南语词表（553 条） |
| `assets/rules/vi-spec-rules.json` | 尺寸+厚度规格串正则（整格可判定才用） |
| `references/terminology-rules.md` | 归一化口径、公式策略、不译清单、单位与术语表、行高算法、体检与交付口径 |

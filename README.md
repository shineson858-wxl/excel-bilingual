# excel-bilingual

把外语 Excel（越南语 / 英语 / 日语…）就地翻成**同一单元格内「上原文、下中文」**的双语对照表——不新增行列、不改版式，同时保住数值、真实公式、合并单元格、自动筛选、条件格式和外部链接。

这是一个给 AI Agent 用的 **Skill**（`SKILL.md` + 可执行脚本 + 词表 + 口径文档），不是独立的翻译软件。人也能直接跑里面的脚本。

> In-cell bilingual translation for Excel: original text on top, Chinese below, inside the same cell. No new rows/columns, no layout change, formulas/merges/filters/conditional formatting/external links preserved.

## 为什么不是"另存一张翻译表"

工程量清单、报价表、BOM 一旦加列或加 sheet，就会破坏原有筛选、合并区、打印版式和跨表引用。本项目的唯一合法形态是**同格双语**：内容塞回原单元格内部，只做"自动换行 + 只加高不压低行高"。

## 它解决的真实问题

| 问题 | 踩过的坑 | 本项目的做法 |
| --- | --- | --- |
| 越南语分解式组合符（NFD） | 不做 Unicode NFC，一个字都对不上 | 三级归一：NFC → 小写 → 折叠连续空白 |
| 全大写行文 | 不折叠大小写，命中率掉到 3% | 同一套三级归一同时用于词表和表内文本 |
| A/B/C 列是跨簿外链公式 | 同格双语与保公式不可兼得 | 缓存值是文本 → 转静态双语；缓存值是数值 → 公式一格不动 |
| openpyxl 保存会静默丢弃对象 | 文本框、图表、透视、控件、VBA 直接没了 | 译前体检 `audit_features.py`，判 `RISKY` 就走 Excel COM |
| 改 sheet 标签名 | 表内公式按表名跨表引用，改名即断链 | 标签名保持原语言 |
| 定额编号/型号/章节序号 | 加了中文反而干扰选型与对图 | 明确不译，交付时列清单 |
| 同类文件反复翻 | 口径不一致、重复劳动 | `mine_glossary.py` 从上一份成品反抽词对并入词表 |

## 目录

```
SKILL.md                      # Skill 主体：流程 + 硬规则 + 资源表
scripts/
  audit_features.py           # 译前体检：openpyxl 会丢什么 → OPENPYXL_OK / _WITH_CAUTION / RISKY
  scan_text.py                # 扫描待译文本，产出唯一原文清单
  mine_glossary.py            # 从已译成品反抽 原文→译文，按三级归一去重并入词表
  apply_bilingual.py          # 主路径：openpyxl 同格写入 + 自动换行 + 行高
  apply_bilingual_com.ps1     # 保真退路：Excel COM 写入（可保住文本框/图表/控件/VBA）
  verify_bilingual.py         # 逐格 + 结构校验，PASS / FAIL
references/
  terminology-rules.md        # 归一化口径、公式策略、不译清单、术语表、行高算法、COM 边界
assets/
  glossary/vi-zh-mep.json     # 机电 MEP 越南语→中文词表（553 条）
  rules/vi-spec-rules.json    # 尺寸+厚度规格串正则
```

## 用法

```bash
# 1 译前体检
python scripts/audit_features.py --source ./source.xlsx

# 2 扫待译文本
python scripts/scan_text.py --source ./source.xlsx --out ./unique_strings.tsv

# 3 写入（主路径）
python scripts/apply_bilingual.py --source ./source.xlsx --out ./out.xlsx \
  --glossary assets/glossary/vi-zh-mep.json \
  --rules assets/rules/vi-spec-rules.json \
  --missing ./missing.txt

# 4 校验（必须 PASS）
python scripts/verify_bilingual.py --source ./source.xlsx --out ./out.xlsx
```

体检结论是 `RISKY`（有文本框 / 图表 / 透视 / 控件 / VBA）时改走 COM 写入器：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/apply_bilingual_com.ps1 `
  -Source .\source.xlsx -Out .\out.xlsx `
  -Glossary assets\glossary\vi-zh-mep.json -Rules assets\rules\vi-spec-rules.json

# COM 产物校验：只放行三类已知痕迹
python scripts/verify_bilingual.py --source ./source.xlsx --out ./out.xlsx --allow-com-artifacts
```

## 用上一份成品的译法

```bash
python scripts/mine_glossary.py --bilingual ./前一份中越对照.xlsx --out ./glossary_new.json \
  --merge assets/glossary/vi-zh-mep.json --rules assets/rules/vi-spec-rules.json
```

## 实测

- 8 sheet / 2 346 个待译格 / 831 条唯一原文的越南语机电工程量清单：双语格 **1 654**、公式缓存文本转静态 **948**、未译保留 **226 种 / 690 格**；合并区 99、autoFilter 4、条件格式 19 条逐项对齐，`verify_bilingual.py` → `RESULT: PASS`。
- COM 保真路径整簿跑通（2 067 格外链序号被 Excel 重排、5 格 `\r\n` 转 `_x000D_`、47 格首尾空格被规范化——三类均语义等价，见 `references/terminology-rules.md` §13.2），整簿约 4~5 分钟。

## 安装为 Skill

放进 Agent 的 skill 目录即可（例如 `<workspace>/.user_skills/excel-bilingual/`），保证 `SKILL.md` 在目录根部。

## 许可

MIT，见 [LICENSE](LICENSE)。词表中的器件型号、品牌名属各自权利人，仅作术语对照用。

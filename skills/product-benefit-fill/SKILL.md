---
name: product-benefit-fill
description: 把保险产品评价工作簿里的“产品利益信息——结果”和“产品利益信息——底稿”两张表填完整并核对：盘点待补指标，按产品查找官方条款/现金价值表/产品说明书，取保证利益和分红利益（累积生息、交清增额）的逐年数据，计算 IRR、返本年度、翻倍年度，写入底稿逐年块和汇总行、结果表口径标记，另存新文件并用 Excel 全量重算核对，最后更新待补清单和口径说明。含计划书平台公开预览利益演示接口的取数脚本（不登录、不生成计划书、遇拒即停、须使用人同意）。Use when filling or auditing the product benefit sheets (guaranteed / dividend IRR, payback year, doubling year) of a life-insurance product evaluation workbook.
license: MIT
compatibility: Python 3.8+ 与 lxml。核对公式需要 Windows + 本机 Excel（没有时只能做 Python 复算，须在汇报里说明）。平台取数只用 Python 标准库。
metadata:
  author: xuyue9129-ctrl
  version: "1.0.0"
  language: zh-CN
---

# 产品利益信息 · 填表与核对

这份技能把“把两张利益表填完整”的全流程固定下来。分工是：**脚本做确定性的事**（定位、计算、写入、另存、核对），**你（Agent）负责找资料、判断口径、写标注**。脚本拿到不合格的数据会拒绝并说明原因，不会替你编数。

`<SKILL_DIR>` 指本文件所在目录，下文 `PB = python <SKILL_DIR>/scripts/pbfill.py`（每个子命令都可加 `-h`）。

## 一、不可违反的规则

1. **不编数、不插值、不外推。** 缺哪一年就是缺哪一年；返本/翻倍年度之前有缺数年度时该指标保持待补。
2. **每个数都要有出处。** 每个填列的组合必须写口径标记（结果表 Q 列）和来源（底稿汇总行 AC 列）：文件名或网址、页码、参数、取数日期。脚本不接受没有 `label` 和 `source` 的数据。
3. **只另存新文件。** 不覆盖、不移动、不删除使用人已有的任何文件；输出文件名已存在时脚本拒绝写入。
4. **已填的值不改。** 新数据与已填指标或已录入的官方披露点不一致时，脚本整组拒绝；不要为了“填上”去改容差或删旧值，把差异报告给使用人。
5. **不符合目标口径的数据可以填，但必须逐行标明。** 例如使用人要求按 3.5% 演示，但只找到其他演示利率的数据：照原样填，在标注里写清楚实际口径和判定依据。是否允许这样填由使用人决定，先问。
6. **平台取数的边界**（见第五节）：只调用公开预览的利益演示接口；不登录、不读取或保存任何会话/令牌、不处理验证码、不调用需要签名的接口、不生成计划书、不付费、不做界面自动化。平台一旦返回登录/会员/验证/频繁类提示，**立即停止并告知使用人**，不换入口、不换参数、不换网络重试。每次取数前先得到使用人同意。
7. **如实汇报。** 没核对的说没核对，没找到的说没找到并列出找过哪里。

## 二、工作簿结构与口径（先读）

- 两张表的列位、公式、占位文字：[references/workbook.md](references/workbook.md)
- 指标定义和计算口径（IRR 终点、寿险/年金的总利益、累积生息、交清增额、返本/翻倍判定）：[references/methodology.md](references/methodology.md)

默认参数取自底稿参数行：40 岁男性、年交（或趸交）保费 10 万元、红利累积利率 1.75%、演示利率 3.5%。参数行不同则以参数行为准。

## 三、流程

### 1. 盘点

```bash
PB inspect 工作簿.xlsx --json work/inspect.json --csv work/待补清单.csv
```

输出：组合总数、待补的保证指标数、待补的分红指标数、已填数、完整组合数，以及逐组合的待补清单（结果表行号、产品、交费期、缺哪些指标、当前标注、有没有逐年块）。`uncalculated` 不为 0 说明文件没有保存过计算结果，先让使用人在 Excel 里打开保存一次。

把清单按产品分组，建立任务清单；每个产品一个资料目录。

### 2. 找资料

按 [references/sources.md](references/sources.md) 的优先级逐个产品查找并归档：官方条款、现金价值表/费率表、产品说明书（分红利益演示）、公司官网披露、行业协会产品信息库。记录每份资料的来源网址、下载日期、对应页码。找不到时记录找过的位置和关键词，不要用相近产品顶替。

### 3. 保证利益（三项：保证 IRR、返本、翻倍）

条件：拿到 40 岁男性、对应交费期的现金价值（及生存金、满期金）逐年数据。按 10 万元保费折算到底稿口径后写成填列说明（spec）：

```json
{"result_row": 91, "label": "…口径标记…", "source": "…来源…",
 "guarantee": {"cash_value": [第1年末, 第2年末, …], "survival": [各年生存金], "maturity": [各年满期金]}}
```

这会在底稿末尾新增一个逐年块并把汇总行改成公式。折算方法、年金“给付后现价”的口径见 methodology。

### 4. 分红利益（六项：累生 IRR/返本/翻倍，增额 IRR/返本/翻倍）

组合必须已有保证利益的逐年块。数据来源两类：

- **官方产品说明书的利益演示表**：通常只披露部分年度。把披露年度的累积红利余额填进 `accumulate.balance`，未披露年度写 `null`。IRR 终点年度有数就能算 IRR；返本/翻倍只有在此前各年都有数时才会填。
- **平台公开预览演示**（第五节）：逐年数据，但演示利率要自己从数字里判定。

```json
{"result_row": 390, "label": "…", "source": "…",
 "accumulate": {"dividends": [各年当年红利]},
 "paidup": {"total_cv": [各年含交清增额的总现价], "extra_annuity": [各年红利年金], "extra_maturity": [各年红利满期金]}}
```

字段说明和更多写法见 [references/spec.md](references/spec.md)，示例见 `examples/`。

### 5. 试填、正式填、核对

```bash
PB fill 工作簿.xlsx --spec work/specs/*.json --dry-run                    # 先看哪些能填、哪些被拒、哪些指标无法判定
PB fill 工作簿.xlsx --spec work/specs/*.json --out 工作簿_新.xlsx --report work/fill_report.json
PB verify 工作簿_新.xlsx --base 工作簿.xlsx --report work/fill_report.json  # Excel 全量重算 + Python 复算
```

- `fill` 逐组合处理，一个组合要么全部写入要么不写。被拒的组合在 `refused` 里给出原因，按 [references/troubleshooting.md](references/troubleshooting.md) 处理，不要绕过。
- `fill` 保存后会逐部件、逐单元格比对新旧文件，证明除本次写入的单元格外没有任何变化（样式、合并单元格、其他工作表都不动）。
- `verify` 必须通过才算完成：`error_cells` 为 0、`old_changed` 为空（原有值一个没变）、`mismatch` 为空（Excel 按公式算出的结果与脚本写入的结果一致）、`newly_filled` 等于预期。没有 Excel 时用 `--no-excel`，此时只做了 Python 复算，汇报里要写明“公式未经 Excel 重算核对”。
- 如果 Excel 正开着这份文件，先请使用人关闭，不要强行处理。

### 6. 更新清单和说明

```bash
PB inspect 工作簿_新.xlsx --csv work/剩余待补清单.csv
```

交付三样东西：新工作簿、剩余待补清单（每个仍待补的组合写明缺什么、为什么缺、找过哪里、需要使用人提供什么）、口径说明（数据来源分类及数量、每类的判定方法、与目标口径不一致的行有多少、核对结果）。清单和说明都另存新文件。

### 7. 汇报

用使用人的语言，先说结果：新填了多少项、还剩多少项及原因分类、核对是否通过（列出实际数字）、哪些行口径与要求不一致、需要使用人决定什么。不要把“没找到”写成“没有”。

## 四、标注怎么写

结果表 Q 列的标注是给看表的人用的，要让人不查底稿就知道这个数的性质。至少包含：数据来源类别、演示口径及判定依据、是否折算/推算、数据缺口。`platform-convert` 会生成草稿；模板见 [references/spec.md](references/spec.md)。旧标注不会丢，脚本会把它放进新标注末尾的〔此前标注：…〕里。

## 五、平台公开预览取数（可选，须使用人同意）

适用：官方资料没有逐年分红数据，而计划书平台对该产品有公开预览的利益演示。完整说明、字段、核对规则、停止条件见 [references/platform.md](references/platform.md)。

前提：
- 使用人明确同意本次取数（同意人和时间要写进命令，会记入请求日志）。
- 使用人提供该产品的“产品录入配置”JSON。**本技能不提供获取它的办法**，也不处理登录。

```bash
PB platform-choose  --init 产品配置.json --pay 5 --term "终身" --out work/p/row123/pbdata.json       # 不联网
PB platform-fetch   --pbdata work/p/row123/pbdata.json --out work/p/row123 --state work/p/_state \
                    --consent "张三 2026-01-05 同意本次取数" [--annuity]
PB platform-convert 工作簿.xlsx --row 123 --dir work/p/row123 --out work/specs/row123.json
```

- `platform-fetch` 每个组合最多发 2 个请求（累积生息、交清增额各一次），单线程，间隔不少于 1.5 秒，每天不超过 200 个，失败不重试。
- 命令返回 `stopped: true`（退出码 3）时：停下，把原因原样告诉使用人。停止标记文件只有使用人可以删除。
- `platform-convert` 先核对平台的保费和保证现金价值与底稿逐年一致，再取分红数据；并用“当年红利÷上年末保证现价”判定演示利率。平台自己的文字标签不作为依据。
- 平台数据是第三方数据，标注里必须写明“平台公开预览演示”，与官方披露点冲突时以官方为准。

## 六、文件

| 路径 | 内容 |
|---|---|
| `scripts/pbfill.py` | 命令行入口 |
| `scripts/pbfill_lib/` | `xlsx`（XML 级读写与另存校验）、`layout`（定位两张表）、`calc`（指标计算）、`fill`（生成写入）、`verify`（Excel 重算与复算）、`platform`（平台请求）、`convert`（平台数据核对与转换） |
| `references/` | 工作簿结构、口径、资料来源、填列说明格式、平台取数、排错 |
| `examples/` | 合成示例：示例工作簿生成脚本、三份填列说明、一份产品配置（均为虚构数据） |

用示例走一遍：

```bash
python <SKILL_DIR>/examples/make_demo_workbook.py demo.xlsx
PB inspect demo.xlsx
PB fill demo.xlsx --spec <SKILL_DIR>/examples/spec_dividend.json <SKILL_DIR>/examples/spec_guarantee.json --out demo_filled.xlsx --report demo_report.json
PB verify demo_filled.xlsx --base demo.xlsx --report demo_report.json
```

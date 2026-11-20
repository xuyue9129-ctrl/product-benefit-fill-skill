# 产品利益信息填表 Skill

给 AI Agent 用的技能：把保险产品评价工作簿里“产品利益信息——结果”和“产品利益信息——底稿”两张表填完整并核对。覆盖全流程：盘点待补 → 查找官方资料 → 取保证利益和分红利益逐年数据 → 计算 IRR、返本年度、翻倍年度 → 写入底稿和结果表并逐行标注口径 → 另存新文件 → Excel 全量重算核对 → 更新待补清单和口径说明。

**仓库里没有任何保险产品数据、工作簿、采集原文、账号或凭据。** 示例和测试用的产品、数字都是虚构的。

## 安装

把 `skills/product-benefit-fill` 整个目录复制到 Agent 的技能目录，例如：

```text
你的工作区/.factory/skills/product-benefit-fill/      （或 ~/.agents/skills/product-benefit-fill/）
  SKILL.md
  references/
  scripts/
  examples/
```

依赖：Python 3.9+，`pip install -r requirements.txt`（只有 lxml）。核对公式需要 Windows 和本机 Excel；没有时可以做 Python 复算，但公式未经 Excel 核对，需要在汇报里说明。

## 命令

```bash
PB="python skills/product-benefit-fill/scripts/pbfill.py"
$PB inspect 工作簿.xlsx --csv 待补清单.csv                      # 盘点
$PB fill 工作簿.xlsx --spec specs/*.json --dry-run              # 试填
$PB fill 工作簿.xlsx --spec specs/*.json --out 新.xlsx --report report.json
$PB verify 新.xlsx --base 工作簿.xlsx --report report.json      # Excel 重算核对
$PB audit 工作簿.xlsx                                           # 用逐年块复算全部已填指标
```

用合成示例走一遍：

```bash
python skills/product-benefit-fill/examples/make_demo_workbook.py demo.xlsx
$PB inspect demo.xlsx
$PB fill demo.xlsx --spec skills/product-benefit-fill/examples/spec_dividend.json skills/product-benefit-fill/examples/spec_guarantee.json --out demo_filled.xlsx --report demo_report.json
$PB verify demo_filled.xlsx --base demo.xlsx --report demo_report.json
```

## 脚本保证什么

- 只另存新文件；输出文件已存在就拒绝。保存后逐部件、逐单元格比对，证明除本次写入的单元格外没有变化。
- 已填的值不覆盖；新数据与已填指标或已录入的官方披露点不一致时整组拒绝。
- 不插值：返本/翻倍年度之前有缺数年度时，该指标保持待补。
- 没有口径标记和来源的数据不接受。

## 平台取数的范围

技能里带有计划书平台“公开预览利益演示接口”的调用脚本，限制如下：

- 只调用这一个不需要登录的接口。不登录，不读取或保存会话、令牌、Cookie，不处理验证码，不调用签名接口，不生成计划书，不付费，不做界面自动化。
- 需要使用人提供该产品的录入配置文件。**仓库不包含获取这份配置的办法。**
- 单线程，请求间隔不少于 1.5 秒，每天最多 200 个请求，失败不重试。
- 平台返回登录、会员、验证、频繁类提示时立即停止并写入停止标记，之后的调用一律拒绝，直到使用人自己处理。
- 每次取数前须经使用人同意，`--consent` 参数会记入请求日志；没有这个参数不会发出请求。

使用前请自行确认符合平台的服务条款和所在机构的合规要求。

## 目录

| 路径 | 内容 |
|---|---|
| `skills/product-benefit-fill/SKILL.md` | 流程、规则、命令 |
| `skills/product-benefit-fill/references/` | 工作簿结构、计算口径、资料来源、填列说明格式、平台取数、排错 |
| `skills/product-benefit-fill/scripts/` | 命令行入口和库 |
| `skills/product-benefit-fill/examples/` | 合成示例 |
| `tests/` | 单元测试（合成工作簿 + 本地假接口，不联网） |

## 测试

```bash
python -m unittest discover -s tests -v
```

## 许可

MIT

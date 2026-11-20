# 产品利益信息填表 Skill

给 AI Agent 用的技能：把保险产品评价工作簿里“产品利益信息——结果”和“产品利益信息——底稿”两张表填完整并核对。覆盖全流程：盘点待补 → 查找官方资料 → 取保证利益和分红利益逐年数据 → 计算 IRR、返本年度、翻倍年度 → 写入底稿和结果表并逐行标注口径 → 另存新文件 → Excel 全量重算核对 → 更新待补清单和口径说明。

**仓库里没有任何保险产品数据、工作簿、采集原文、账号或凭据。** 示例和测试用的产品、数字都是虚构的。

## 安装

本技能不依赖任何特定 Agent 的工具、账号或插件。下载仓库 ZIP 并解压，或克隆仓库；在仓库目录执行：

```text
python install.py --dest "你的Agent技能目录/product-benefit-fill"
```

`--dest` 是技能的完整目标目录，由你按照 Agent 的说明指定；目录已存在时不会覆盖。也可以直接复制 `skills/product-benefit-fill` 整个目录。安装后的内容为：

```text
你的Agent技能目录/product-benefit-fill/
  SKILL.md
  references/
  scripts/
  examples/
  requirements.txt
```

进入安装后的目录，安装依赖并运行独立自检（只用临时合成数据，不联网）：

```text
python -m pip install -r requirements.txt
python scripts/selftest.py
```

- Python 3.9+，唯一第三方依赖为 lxml。Windows 可用 `py`，macOS/Linux 可用 `python3` 替换命令里的 `python`；使用虚拟环境时，安装依赖和执行脚本必须用同一个解释器。
- Windows、Linux、macOS 均可盘点、填表和 Python 复算；Excel 公式全量重算仅支持 Windows + 桌面版 Excel。其他环境用 `verify --no-excel`，不能声称完成 Excel 重算。
- 支持读取 `SKILL.md` 的 Agent：把目录放入其官方技能目录，重新加载后使用。不支持技能自动发现的 Agent：直接把 `SKILL.md` 作为上下文，要求它按文档执行。Agent 必须能读取本地文件、执行 Python；纯聊天网页不能独立运行脚本。
- 完整可复制的指令、迁移边界见 [跨电脑与跨 Agent 使用](skills/product-benefit-fill/references/portability.md)。

## 命令

以下在安装后的技能目录执行，Windows PowerShell 与 macOS/Linux 终端均适用：

```text
python scripts/pbfill.py inspect "工作簿.xlsx" --csv "待补清单.csv"
python scripts/pbfill.py fill "工作簿.xlsx" --spec "specs/*.json" --dry-run
python scripts/pbfill.py fill "工作簿.xlsx" --spec "specs/*.json" --out "新.xlsx" --report report.json
python scripts/pbfill.py verify "新.xlsx" --base "工作簿.xlsx" --report report.json
python scripts/pbfill.py audit "工作簿.xlsx"
```

用合成示例走一遍：

```text
python examples/make_demo_workbook.py demo.xlsx
python scripts/pbfill.py inspect demo.xlsx
python scripts/pbfill.py fill demo.xlsx --spec examples/spec_dividend.json examples/spec_guarantee.json --out demo_filled.xlsx --report demo_report.json
python scripts/pbfill.py verify demo_filled.xlsx --base demo.xlsx --report demo_report.json
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

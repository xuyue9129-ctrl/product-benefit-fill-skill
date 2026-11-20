# 跨电脑与跨 Agent 使用

## 安装包是独立的

只需完整的 `product-benefit-fill` 目录：`SKILL.md`、`references/`、`scripts/`、`examples/`、`requirements.txt`。不依赖仓库根目录、原作者电脑路径、其他技能、特定 Agent 的内置工具或登录状态。

从 GitHub 下载 ZIP 解压后，在仓库根目录用 `python install.py --dest "目标技能目录/product-benefit-fill"` 安装，或直接复制该目录。安装器只复制文件，不联网、不安装依赖、不修改 Agent 设置，目标已存在时拒绝覆盖。随后在安装后的目录执行：

```text
python -m pip install -r requirements.txt
python scripts/selftest.py
```

推荐独立 Python 虚拟环境；安装依赖和运行脚本用同一解释器。Windows 常用 `py`，macOS/Linux 常用 `python3`。技能目录若为只读，自检仍可执行，其工作文件均在系统临时目录内，结束时自动清理。

## 接入其他 Agent

支持标准 `SKILL.md` 的 Agent：按该 Agent 的说明，把整个目录放入它认可的项目级或用户级技能目录。不同产品目录和启用方式不统一，本仓库不擅自修改任何 Agent 设置。

不支持技能自动发现的 Agent：让它读取本技能的文件，或把以下提示作为任务上下文：

```text
读取技能目录里的 SKILL.md 并按其中流程处理我提供的工作簿。
技能目录是：【本电脑上 product-benefit-fill 的绝对路径】。
Python 解释器是：【本电脑的 Python 命令或绝对路径】。
先安装该目录 requirements.txt 的依赖并运行 scripts/selftest.py。
先盘点、给出待补清单，再查资料和填列；原文件不改，另存新文件。
先询问我是否允许本次平台取数，未获同意不发送请求。
没有 Windows 桌面版 Excel 时，只做 --no-excel 复算并明确说明限制。
```

Agent 至少需要本地文件读取和 Python 命令执行能力；找资料需要联网搜索/浏览能力。仅可聊天、不能运行命令的 Agent 只能指导操作者执行，不能独立完成工作。

## 能力边界

| 功能 | Windows | Linux / macOS |
|---|---|---|
| 盘点、试填、写入新工作簿 | 支持 | 支持 |
| Python 复算、离线自检 | 支持 | 支持 |
| Excel 公式全量重算 | 需桌面版 Excel 和 PowerShell | 不支持，改用 `verify --no-excel` |
| 平台公开预览取数 | 需网络、使用人同意及合法产品配置 | 同左 |

平台接口可能变更或不再开放，离线自检不证明线上可用。迁移时应保留原始来源资料、工作簿、逐行标注、请求日志和停止状态；不能换电脑或换状态目录绕过平台拒绝。

工作簿必须符合 `workbook.md` 的两张表结构；这不是任意 Excel 表格的通用填表工具。仓库不含业务数据，使用者须自行提供工作簿和资料。

## 自检验证范围

`selftest.py` 使用安装目录中的合成示例，从不同工作目录以独立 Python 进程完成生成、盘点、通配符试算写入和复算，检查新增 10 项指标、保证待补 0 项、分红仍缺 5 项。不会连接保险平台或修改真实工作簿。

持续集成还会把技能复制到带空格和中文的独立目录后运行自检，覆盖 Windows、Linux、macOS。通过不代表所有 Agent 都已逐一集成，也不代替真实业务数据核对。

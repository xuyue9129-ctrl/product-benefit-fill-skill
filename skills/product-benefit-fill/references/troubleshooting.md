# 排错

`fill` 的 `refused` 和 `platform-convert` 的 `reason` 都是直接可读的原因。下面是常见情况和正确的处理方式。共同原则：**拒绝说明数据或口径有问题，去解决问题，不要绕过检查。**

| 提示 | 含义 | 处理 |
|---|---|---|
| `label is required` / `source is required` | 没写口径标记或来源 | 补上，见 spec.md 模板 |
| `this workbook has never been calculated` | 文件里没有公式的计算结果 | 请使用人在 Excel 打开并保存一次 |
| `output already exists` | 输出文件已存在 | 换一个新文件名，不要删旧文件 |
| `result row N is not a product group row` | 行号不是组合行 | 用 `inspect` 清单里的行号 |
| `… indicators are not pending` | 这组指标已经填了 | 不覆盖。若认为旧值有误，报告给使用人 |
| `… is already X; the new data would make it Y` | 新数据会改变已填指标 | 两个来源口径不同。查清哪一个符合要求；不合并 |
| `year N is X in the sheet (official figure) but Y in the new data` | 新数据与已录入的官方披露点对不上 | 多数是演示利率或保额口径不同。按 methodology 判定后决定：不填，或经使用人同意用 `constants` 并标明 |
| `the group already has a year-by-year block` | 想重写保证数据 | 保证数据不重写 |
| `result row N has no year-by-year block` | 分红数据要先有保证逐年块 | 先找现金价值表填保证利益 |
| `total cash value is below the guaranteed cash value` | 交清增额总现价低于保证现价 | 数据错位或口径不同（保额、给付前后），回到来源核对 |
| `no accumulated dividend at the IRR end year` | 终点年度没有数 | 该来源算不了 IRR；不要拿相邻年度代替 |
| `balance re-accumulated … differs` | 当年红利累积后与给出的余额不符 | 累积利率不是底稿的利率，或余额口径不同（含/不含当年）；查清再填 |
| `years without data come before the answer` | 返本/翻倍之前有缺数年度 | 保持待补；或补齐数据 |
| `guaranteed cash flows have no IRR` | 现金流全负或数据量级错误 | 检查保额折算和单位 |
| `column K year N is … but … by the standard formula` | 逐年块公式与标准不同 | 该组合用了特殊口径；不要用脚本填，报告给使用人 |
| 平台：`平台保证现金价值与底稿…成固定比例` | 平台按另一保额列示 | 按 platform.md 判断后用 `--scale` |
| 平台：`…不一致，也不成固定比例` | 不是同一产品/组合 | 不用这份数据 |
| 平台：`只返回保证利益` | 该条目没有分红演示 | 记为平台无数据 |
| 平台：`stopped: true` | 平台拒绝或出错 | 停止，告知使用人 |

## verify 不通过

| 结果 | 含义 | 处理 |
|---|---|---|
| `not_run: ABORT …` | 本机 Excel 已有打开的工作簿 | 请使用人关闭 Excel 后重试；不要结束使用人的 Excel 进程 |
| `error_cells` > 0 | 重算后有错误值 | 找到对应行，通常是引用范围或数据类型问题；修正 spec 后从原文件重新填 |
| `old_changed` 非空 | 原有指标变了 | 不应发生。不要交付这个文件，查明原因 |
| `mismatch` 非空 | Excel 结果与脚本写入的不同 | 不要交付。多为逐年块里存在脚本没识别的特殊公式 |
| `audit.difference_count` > 0 | 逐年块复算与表内值不同 | 逐条看；是历史行还是本次写入的行 |

## 其他

- 文件很大时 `inspect`、`fill` 需要几十秒；`verify` 的 Excel 重算可能需要几分钟。
- 结果表的公式值由脚本写入缓存；交付前以 `verify` 的 Excel 重算为准。
- 同一组合分两次填（先累积生息，后交清增额）是可以的：第二次以第一次的输出为输入。

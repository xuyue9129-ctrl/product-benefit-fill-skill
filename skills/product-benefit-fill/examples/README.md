# 示例（全部为虚构数据）

| 文件 | 用途 |
|---|---|
| `make_demo_workbook.py` | 生成一份与真实工作簿同结构的小工作簿：3 个虚构产品，第 9、11 行缺分红指标，第 10 行缺保证指标 |
| `spec_dividend.json` | 第 9 行：累积生息（当年红利）+ 交清增额（总现价） |
| `spec_guarantee.json` | 第 10 行：新增保证利益逐年块 |
| `spec_official_points.json` | 第 11 行：只有 3 个披露年度的累积红利余额，未披露年度写 `null`；只会填 IRR |
| `product_config_example.json` | 产品录入配置的最小结构，供 `platform-choose` 试用（虚构，不对应任何真实产品） |

```bash
python make_demo_workbook.py demo.xlsx
python ../scripts/pbfill.py inspect demo.xlsx
python ../scripts/pbfill.py fill demo.xlsx --spec spec_dividend.json spec_guarantee.json spec_official_points.json --out demo_filled.xlsx --report demo_report.json
python ../scripts/pbfill.py verify demo_filled.xlsx --base demo.xlsx --report demo_report.json
python ../scripts/pbfill.py platform-choose --init product_config_example.json --pay 1 --term 终身 --out pbdata.json
```

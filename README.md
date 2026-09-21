# Temp

## 销售至联营公司顺流交易抵消 · 会计核算

- 可读版本：`reports/销售至联营公司顺流交易抵消-会计核算.html`（含在线核算器）
- Word：`reports/销售至联营公司顺流交易抵消-会计核算.docx`
- PDF：`reports/销售至联营公司顺流交易抵消-会计核算.pdf`

重新生成文档：

```bash
python3 scripts/generate_downstream_docs.py
```

命令行核算示例：

```bash
python3 scripts/downstream_elimination.py --ownership 0.2 --price 1000 --cost 600 --remain 1 --ni 2000
python3 -m unittest tests.test_downstream_elimination
```

# 案例用测试数据

| 文件 | 说明 |
|------|------|
| **发票导入_本周.xlsx** | 应付对账案例「发票侧」示例：表头与 **[scenario-ap-reconciliation.md](../scenario-ap-reconciliation.md)** 附录一致；数据与 Mock API 中 `PO-2026-0091` 等可对账演示。 |

重新生成（需已安装 `openpyxl`）：

```bash
cd "$(git rev-parse --show-toplevel 2>/dev/null || pwd)"
python3 scripts/gen_fixture_invoice_import.py
```

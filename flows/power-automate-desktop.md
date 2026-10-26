# Power Automate Desktop flow

Flow name: `Process_Daily_Operations_Report`

## Sequence

1. **Set variable** — `WorkbookPath` points to the local macro-enabled workbook.
2. **Launch Excel** — opens the workbook and stores the session in `ExcelInstance`.
3. **Run Excel macro** — executes `ProcessOperationsReport`.
4. **Set active Excel worksheet** — activates `KPI_Summary`.
5. **Read B12** — stores exception count in `ExceptionCount`.
6. **Read B8** — stores overall error rate in `OverallErrorRate`.
7. **Read B11** — stores overall orders/hour in `OrdersPerHour`.
8. **If `ExceptionCount > 0`** — evaluates whether review is required.
9. **Display message** — surfaces exception count and key KPI values.
10. **End** — closes the condition block.
11. **Save Excel** — saves processed outputs.
12. **Close Excel** — closes the Excel instance.

The VBA macro is intentionally silent in the automated version so PAD can run without a blocking `MsgBox`.

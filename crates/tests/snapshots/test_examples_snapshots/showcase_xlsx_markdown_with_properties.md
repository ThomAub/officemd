<!-- officemd: kind=xlsx profile=compact first_row_as_header=true formulas=true headers_footers=true properties=true -->

properties: AppVersion=3.1; Application=Microsoft Excel Compatible / Openpyxl 3.1.5; dc:creator=openpyxl; dcterms:created=2026-03-03T12:41:03Z; dcterms:modified=2026-03-03T12:41:03Z

## Sheet: Sales

### Table 1 (rows 1–6, cols A–E)
| Product | BaseAmount | Rate | Total | Notes |
| --- | --- | --- | --- | --- |
| Widget | 1200 | 0.15 |  | Primary SKU |
| Gadget | 850 | 0.1 |  | Secondary SKU |
| Service | 600 | 0.2 |  | Recurring |
|  |  |  |  |  |
| Project Wiki |  |  |  |  |

D2=`=B2*(1+C2)`
D3=`=B3*(1+C3)`
D4=`=B4*(1+C4)`

## Sheet: Summary

### Table 2 (rows 1–4, cols A–B)
| Metric | Value |
| --- | --- |
| ReportDate |  |
| RunAt |  |
| AverageRate |  |

B4=`=AVERAGE(Sales!C2:C4)`

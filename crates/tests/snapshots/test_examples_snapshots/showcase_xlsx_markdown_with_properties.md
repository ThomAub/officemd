properties: AppVersion=3.1; Application=Microsoft Excel Compatible / Openpyxl 3.1.5; dc:creator=openpyxl; dcterms:created=2026-03-03T12:41:03Z; dcterms:modified=2026-03-03T12:41:03Z

## Sheet: Sales

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

| Metric | Value |
| --- | --- |
| ReportDate |  |
| RunAt |  |
| AverageRate |  |

B4=`=AVERAGE(Sales!C2:C4)`

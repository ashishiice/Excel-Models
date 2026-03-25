# Excel-Models 🦞

Treasury and finance Excel models — built with Python + openpyxl, VBA macros delivered separately.

## Structure

```
Excel-Models/
├── NCD/          # NCD issuance, placement, HTM tracking
├── PTC/          # PTC securitisation, waterfall, Ind AS 109
├── OIS/          # OIS scorecard, rates modelling
├── ALM/          # Asset-Liability Management
├── FX/           # FX hedging, CCS, currency risk
├── Automation/   # Treasury ops automation (task taxonomy, SLA tracker)
└── _templates/   # Reusable Excel templates
```

## Delivery Standard

- Python/openpyxl generates `.xlsx` files
- VBA modules delivered as `.bas` files (import manually)
- Each model includes: inputs sheet, logic sheet, outputs sheet

## Principle

Automate before manual. Build systems, not spreadsheets.

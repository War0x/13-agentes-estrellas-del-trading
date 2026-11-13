---
name: trader-reporter
description: >-
  Executive Excel Summary Generator. Compiles trade plans, risk checks,
  orders, and paper-trading logs into an executive Excel (.xlsx) dashboard
  featuring KPI cards, risk limits compliance, and full operations audit trail.
argument-hint: "[generate | update] [--output PATH]"
---

# trader-reporter

The Executive Excel Reporter compiles the activity of the Day Trader Swarm into a structured, executive-grade Excel workbook (`trader/Resumen_Ejecutivo_Trading.xlsx`).

```bash
python trader-reporter/report.py
```

## Features:
1. **Executive Dashboard (KPIs)**: Total operations, planned risk, daily budget and risk constraints.
2. **Operations Ledger**: Detailed log of each trade proposed, risk-checked, approved, and executed in paper trading mode.
3. Formatted with professional financial styling (Navy/Green/Orange accents, currency formatting, auto-fit columns).

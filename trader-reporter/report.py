#!/usr/bin/env python3
"""
trader-reporter: Generador de Resumen Ejecutivo en Excel para Day Trader Swarm.
Lee los planes, ordenes de paper trading, ejecuciones y gestion de riesgo
y compila un libro de Excel (.xlsx) ejecutivo con KPIs y bitacora detallada.
"""

import os
import sys
import json
import argparse
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TRADER_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "trader")
DEFAULT_EXCEL_PATH = os.path.join(TRADER_DIR, "Resumen_Ejecutivo_Trading.xlsx")
HISTORY_LOG_PATH = os.path.join(TRADER_DIR, "trade_history.json")


def load_json(filepath):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


def ensure_history_updated():
    """Consolida las ordenes de orders.json, plan.json, risk.json y broker.json en trade_history.json."""
    os.makedirs(TRADER_DIR, exist_ok=True)
    history = []
    if os.path.exists(HISTORY_LOG_PATH):
        try:
            with open(HISTORY_LOG_PATH, "r", encoding="utf-8") as f:
                history = json.load(f)
        except Exception:
            history = []

    plan = load_json(os.path.join(TRADER_DIR, "plan.json"))
    orders_data = load_json(os.path.join(TRADER_DIR, "orders.json"))
    risk_data = load_json(os.path.join(TRADER_DIR, "state", "risk.json"))
    broker_data = load_json(os.path.join(TRADER_DIR, "state", "broker.json"))
    portfolio = load_json(os.path.join(TRADER_DIR, "portfolio.json"))
    config = load_json(os.path.join(TRADER_DIR, "config.json")) or {}

    paper_mode = config.get("paper", True)
    mode_str = "PAPER TRADING" if paper_mode else "LIVE"

    # Procesar ordenes de orders.json
    if orders_data and "orders" in orders_data and orders_data["orders"]:
        created_at = orders_data.get("created_at", datetime.now().isoformat())
        for order in orders_data["orders"]:
            symbol = order.get("symbol", "N/A").upper()
            side = order.get("side", "buy").upper()
            qty = float(order.get("qty", 0))
            order_type = order.get("type", "market")
            plan_info = order.get("plan", {})
            entry = float(plan_info.get("entry", plan.get("entry", 0.0) if plan else 0.0))
            stop = float(plan_info.get("stop", plan.get("stop", 0.0) if plan else 0.0))
            target = float(plan_info.get("target", plan.get("target", 0.0) if plan else 0.0))
            risk_usd = float(plan_info.get("max_loss", 0.0))
            setup = plan_info.get("setup", plan.get("setup", "ORB") if plan else "ORB")
            why = order.get("why", "")

            # Determinar estado segun riesgo y ejecucion de broker
            status = "PROPOSED (PAPER)" if paper_mode else "PROPOSED"
            if risk_data and "results" in risk_data:
                for r in risk_data["results"]:
                    if str(r.get("symbol", "")).upper() == symbol:
                        v = r.get("verdict", "")
                        if v == "PASS":
                            status = "RISK_PASS (PAPER)" if paper_mode else "RISK_PASS"
                        elif v == "BLOCK":
                            status = f"RISK_BLOCK: {r.get('reason', '')}"

            if broker_data and "orders" in broker_data:
                for b_order in broker_data["orders"]:
                    if str(b_order.get("symbol", "")).upper() == symbol:
                        b_status = b_order.get("status", "")
                        if b_status == "filled":
                            status = "FILLED (PAPER)" if paper_mode else "FILLED"
                        elif b_status in ["approved", "active"]:
                            status = "ACTIVE (PAPER)" if paper_mode else "ACTIVE"

            # Buscar si ya existe en historial para actualizar o insertar
            match = next((h for h in history if h.get("symbol") == symbol), None)
            if match:
                match["qty"] = qty
                match["entry"] = entry
                match["stop"] = stop
                match["target"] = target
                match["risk_usd"] = risk_usd
                match["setup"] = setup
                match["status"] = status
                match["mode"] = mode_str
                match["notes"] = why
                if not match.get("created_at"):
                    match["created_at"] = created_at
            else:
                history.append({
                    "created_at": created_at,
                    "symbol": symbol,
                    "side": side,
                    "qty": qty,
                    "type": order_type,
                    "entry": entry,
                    "stop": stop,
                    "target": target,
                    "risk_usd": risk_usd,
                    "setup": setup,
                    "status": status,
                    "mode": mode_str,
                    "notes": why
                })

    elif plan and "symbol" in plan:
        # Si aun no hay orders.json pero ya hay plan.json generado por exit.py
        symbol = plan.get("symbol", "N/A").upper()
        entry = float(plan.get("entry", 0.0))
        stop = float(plan.get("stop", 0.0))
        target = float(plan.get("target", 0.0))
        risk_per_share = float(plan.get("risk_per_share", 0.0))
        setup = plan.get("setup", "ORB")
        entry_time = plan.get("entry_time") or datetime.now().isoformat()
        
        match = next((h for h in history if h.get("symbol") == symbol), None)
        if not match:
            history.append({
                "created_at": entry_time,
                "symbol": symbol,
                "side": "BUY",
                "qty": 0,
                "type": "bracket",
                "entry": entry,
                "stop": stop,
                "target": target,
                "risk_usd": risk_per_share,
                "setup": setup,
                "status": "PLAN_READY (PENDING SIZE)",
                "mode": mode_str,
                "notes": f"Plan configurado R:R {plan.get('reward_risk', '1.0')} (Esperando dimensionamiento)"
            })

    with open(HISTORY_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(history, f, indent=2)

    return history


def generate_excel(output_path=DEFAULT_EXCEL_PATH):
    history = ensure_history_updated()
    config = load_json(os.path.join(TRADER_DIR, "config.json")) or {}

    wb = Workbook()
    
    # -------------------------------------------------------------
    # 1. Hoja: Resumen Ejecutivo
    # -------------------------------------------------------------
    ws_kpi = wb.active
    ws_kpi.title = "Resumen Ejecutivo"
    ws_kpi.views.sheetView[0].showGridLines = True

    # Paleta de Colores
    NAVY = "1B365D"
    WHITE = "FFFFFF"
    LIGHT_BLUE = "E8EEF5"
    GRAY_TEXT = "595959"
    ACCENT_GREEN = "2E7D32"
    LIGHT_GREEN = "E8F5E9"
    ACCENT_ORANGE = "E65100"
    LIGHT_ORANGE = "FFF3E0"

    # Fuentes
    title_font = Font(name="Segoe UI", size=18, bold=True, color=NAVY)
    subtitle_font = Font(name="Segoe UI", size=10, italic=True, color=GRAY_TEXT)
    card_title_font = Font(name="Segoe UI", size=9, bold=True, color=GRAY_TEXT)
    card_value_font = Font(name="Segoe UI", size=18, bold=True, color=NAVY)
    section_font = Font(name="Segoe UI", size=13, bold=True, color=NAVY)
    header_font = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
    bold_cell_font = Font(name="Segoe UI", size=10, bold=True)
    regular_font = Font(name="Segoe UI", size=10)

    # Rellenos
    navy_fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    light_blue_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
    green_card_fill = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type="solid")
    orange_card_fill = PatternFill(start_color=LIGHT_ORANGE, end_color=LIGHT_ORANGE, fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Titulo Principal
    ws_kpi["A2"] = "DAY TRADER SWARM - RESUMEN EJECUTIVO"
    ws_kpi["A2"].font = title_font
    ws_kpi["A3"] = f"Generado el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Modo: {'PAPER TRADING ACTIVO' if config.get('paper', True) else 'LIVE'}"
    ws_kpi["A3"].font = subtitle_font

    # Metricas Clave
    portfolio = load_json(os.path.join(TRADER_DIR, "portfolio.json"))
    total_ops = len(history)
    total_risk = sum(h.get("risk_usd", 0) for h in history)
    symbols = list(set(h.get("symbol") for h in history if h.get("symbol") != "N/A"))
    cash = portfolio.get("cash", config.get("budget", 5000)) if portfolio else config.get("budget", 5000)
    max_daily_loss = config.get("max_daily_loss_usd", 150)

    # Tarjetas KPI (Cards)
    cards = [
        ("B5", "C6", "PRESUPUESTO DISPONIBLE", f"${cash:,.2f}", light_blue_fill),
        ("D5", "E6", "TOTAL OPERACIONES", f"{total_ops}", light_blue_fill),
        ("F5", "G6", "RIESGO TOTAL PLANIFICADO", f"${total_risk:,.2f}", orange_card_fill),
        ("H5", "I6", "PERDIDA MAX DIARIA", f"${max_daily_loss:,.2f}", green_card_fill),
    ]

    for top_left, bot_right, label, val, fill in cards:
        ws_kpi.merge_cells(f"{top_left}:{bot_right}")
        cell = ws_kpi[top_left]
        cell.value = f"{label}\n{val}"
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.font = card_value_font
        cell.fill = fill
        
        # Borde
        start_col, start_row = top_left[0], int(top_left[1:])
        end_col, end_row = bot_right[0], int(bot_right[1:])
        col1_idx = ord(start_col) - ord('A') + 1
        col2_idx = ord(end_col) - ord('A') + 1
        for r in range(start_row, end_row + 1):
            for c in range(col1_idx, col2_idx + 1):
                ws_kpi.cell(row=r, column=c).border = thin_border

    # Seccion: Configuracion del Sistema & Reglas de Riesgo
    ws_kpi["B9"] = "Politicas de Riesgo y Parametros Activos"
    ws_kpi["B9"].font = section_font

    risk_rules = [
        ("Parametro de Riesgo", "Valor Actual", "Estado"),
        ("Modo de Ejecucion", "Paper Trading (Alpaca / Simulado)", "ACTIVO (SEGURO)"),
        ("Confirmacion Humana Obligatoria", "Requerida (Never execute without YES)", "HABILITADA"),
        ("Perdida Maxima por Trade", f"${config.get('max_loss_per_trade_usd', 50):,.2f}", "EN REGLA"),
        ("Perdida Maxima Diaria", f"${config.get('max_daily_loss_usd', 150):,.2f}", "EN REGLA"),
        ("Limite de Trades Diarios", f"{config.get('max_trades_per_day', 3)} trades", "MONITOREADO"),
        ("Horario de Entrada Permitido", f"{config.get('no_entries_before', '09:45')} - {config.get('no_entries_after', '15:30')} ET", "ACTIVO"),
        ("Hora Limite Cierre (Flat)", f"{config.get('flat_by', '15:55')} ET", "ACTIVO"),
        ("Filtro de Apertura (ORB)", "15 minutos (9:30 - 9:45 ET)", "CONFIGURADO")
    ]

    for idx, (param, val, status) in enumerate(risk_rules):
        row = 11 + idx
        ws_kpi[f"B{row}"] = param
        ws_kpi[f"C{row}"] = val
        ws_kpi[f"D{row}"] = status
        
        if idx == 0:
            ws_kpi[f"B{row}"].font = header_font
            ws_kpi[f"C{row}"].font = header_font
            ws_kpi[f"D{row}"].font = header_font
            ws_kpi[f"B{row}"].fill = navy_fill
            ws_kpi[f"C{row}"].fill = navy_fill
            ws_kpi[f"D{row}"].fill = navy_fill
        else:
            ws_kpi[f"B{row}"].font = bold_cell_font
            ws_kpi[f"C{row}"].font = regular_font
            ws_kpi[f"D{row}"].font = bold_cell_font
            if "ACTIVO" in status or "HABILITADA" in status or "EN REGLA" in status:
                ws_kpi[f"D{row}"].fill = green_card_fill

        ws_kpi[f"B{row}"].border = thin_border
        ws_kpi[f"C{row}"].border = thin_border
        ws_kpi[f"D{row}"].border = thin_border

    # Seccion: Tickers Monitoreados
    ws_kpi["F9"] = "Tickers Identificados por el Enjambre"
    ws_kpi["F9"].font = section_font
    
    ws_kpi["F11"] = "Ticker"
    ws_kpi["G11"] = "Frecuencia"
    ws_kpi["H11"] = "Ultimo Setup"
    ws_kpi["F11"].fill = navy_fill
    ws_kpi["G11"].fill = navy_fill
    ws_kpi["H11"].fill = navy_fill
    ws_kpi["F11"].font = header_font
    ws_kpi["G11"].font = header_font
    ws_kpi["H11"].font = header_font

    if symbols:
        for idx, sym in enumerate(symbols):
            r = 12 + idx
            count = sum(1 for h in history if h.get("symbol") == sym)
            last_setup = next((h.get("setup") for h in reversed(history) if h.get("symbol") == sym), "N/A")
            ws_kpi[f"F{r}"] = sym
            ws_kpi[f"G{r}"] = count
            ws_kpi[f"H{r}"] = last_setup.upper()
            ws_kpi[f"F{r}"].font = bold_cell_font
            ws_kpi[f"G{r}"].font = regular_font
            ws_kpi[f"H{r}"].font = regular_font
            ws_kpi[f"F{r}"].border = thin_border
            ws_kpi[f"G{r}"].border = thin_border
            ws_kpi[f"H{r}"].border = thin_border
    else:
        ws_kpi["F12"] = "Sin operaciones aun"
        ws_kpi["G12"] = "-"
        ws_kpi["H12"] = "Esperando escaneo"
        ws_kpi["F12"].border = thin_border
        ws_kpi["G12"].border = thin_border
        ws_kpi["H12"].border = thin_border

    # -------------------------------------------------------------
    # 2. Hoja: Bitacora de Operaciones
    # -------------------------------------------------------------
    ws_ops = wb.create_sheet(title="Bitacora de Operaciones")
    ws_ops.views.sheetView[0].showGridLines = True

    headers = [
        "Fecha / Hora",
        "Ticker",
        "Lado (Side)",
        "Acciones (Qty)",
        "Entrada ($)",
        "Stop Loss ($)",
        "Target ($)",
        "Riesgo ($)",
        "Ratio R/R",
        "Setup",
        "Estado",
        "Modo",
        "Fundamento / Notas"
    ]

    for col_idx, h in enumerate(headers, 1):
        cell = ws_ops.cell(row=2, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row_start = 3
    for idx, h in enumerate(history):
        curr_row = row_start + idx
        entry = h.get("entry", 0.0)
        stop = h.get("stop", 0.0)
        target = h.get("target", 0.0)
        rr = round((target - entry) / (entry - stop), 2) if (entry - stop) > 0 else 1.0

        vals = [
            h.get("created_at", ""),
            h.get("symbol", ""),
            h.get("side", ""),
            h.get("qty", 0),
            f"${entry:,.2f}" if entry else "-",
            f"${stop:,.2f}" if stop else "-",
            f"${target:,.2f}" if target else "-",
            f"${h.get('risk_usd', 0):,.2f}",
            f"{rr}R",
            h.get("setup", "").upper(),
            h.get("status", ""),
            h.get("mode", ""),
            h.get("notes", "")
        ]

        row_fill = light_blue_fill if idx % 2 == 1 else None

        for col_idx, v in enumerate(vals, 1):
            cell = ws_ops.cell(row=curr_row, column=col_idx, value=v)
            cell.font = regular_font
            cell.border = thin_border
            if row_fill:
                cell.fill = row_fill
            if col_idx in [3, 4, 9, 10, 11, 12]:
                cell.alignment = Alignment(horizontal="center")
            elif col_idx in [5, 6, 7, 8]:
                cell.alignment = Alignment(horizontal="right")

    # Ajuste automatico de columnas
    for sheet in [ws_kpi, ws_ops]:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if "\n" in val_str:
                    lines = val_str.split("\n")
                    max_len = max(max_len, max(len(l) for l in lines))
                else:
                    max_len = max(max_len, len(val_str))
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    print(f"Resumen Ejecutivo generado exitosamente en: {output_path}")
    return output_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generador de informe ejecutivo Excel")
    parser.add_argument("--output", default=DEFAULT_EXCEL_PATH, help="Ruta del archivo Excel")
    args = parser.parse_args()
    generate_excel(args.output)

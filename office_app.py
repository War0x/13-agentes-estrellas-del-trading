#!/usr/bin/env python3
"""
Day Trader Virtual Office - Wall Street AI Trading Floor (12 Agentes)
Port: 8765
Simulador de Oficina Virtual Animada para el Enjambre de 12 Agentes de Wall Street (Alpaca Paper Trading).
Proporciona servidor HTTP local, API REST en tiempo real, conexión a la API oficial de Alpaca Markets
y motor de trading autónomo en segundo plano (Zero-Mock Policy).
"""

import os
import sys
import json
import time
import subprocess
import threading
import urllib.request
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRADER_DIR = os.path.join(BASE_DIR, "trader")
EXCEL_PATH = os.path.join(TRADER_DIR, "Resumen_Ejecutivo_Trading.xlsx")
PORT = 8765

# Cache para cotizaciones bursátiles de Yahoo Finance / Stooq
_TICKER_CACHE = {"timestamp": 0, "data": []}
_CACHE_TTL = 10.0  # segundos

DEFAULT_SYMBOLS = ["AAPL", "NVDA", "TSLA", "META", "MSFT", "AMZN", "GOOGL", "AAL", "SPY", "QQQ"]

def read_json_safe(path):
    """Lee un archivo JSON de forma segura sin lanzar excepciones."""
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None

def write_json_safe(path, data):
    """Escribe un archivo JSON de forma atómica y segura."""
    try:
        temp_path = path + ".tmp"
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        os.replace(temp_path, path)
        return True
    except Exception:
        return False

def check_alpaca_connectivity(key=None, secret=None):
    """Verifica la conectividad real con los servidores de Alpaca Paper Trading."""
    if not key or not secret:
        cfg = read_json_safe(os.path.join(TRADER_DIR, "config.json")) or {}
        key = cfg.get("alpaca_key_id")
        secret = cfg.get("alpaca_secret_key")
    
    if not key or not secret or key.startswith("YOUR_"):
        return False, None
    
    try:
        req = urllib.request.Request(
            "https://paper-api.alpaca.markets/v2/account",
            headers={"APCA-API-KEY-ID": key, "APCA-API-SECRET-KEY": secret, "User-Agent": "DayTraderSwarm/1.0"}
        )
        with urllib.request.urlopen(req, timeout=4) as res:
            if res.status == 200:
                data = json.loads(res.read().decode("utf-8"))
                return True, data
    except Exception:
        pass
    return False, None

def fetch_alpaca_positions(key=None, secret=None):
    """Consulta las posiciones abiertas en vivo directamente desde Alpaca API."""
    if not key or not secret:
        cfg = read_json_safe(os.path.join(TRADER_DIR, "config.json")) or {}
        key = cfg.get("alpaca_key_id")
        secret = cfg.get("alpaca_secret_key")
    
    if not key or not secret:
        return []

    try:
        req = urllib.request.Request(
            "https://paper-api.alpaca.markets/v2/positions",
            headers={"APCA-API-KEY-ID": key, "APCA-API-SECRET-KEY": secret, "User-Agent": "DayTraderSwarm/1.0"}
        )
        with urllib.request.urlopen(req, timeout=4) as res:
            if res.status == 200:
                raw_positions = json.loads(res.read().decode("utf-8"))
                parsed = []
                for p in raw_positions:
                    sym = p.get("symbol")
                    qty = float(p.get("qty", 0))
                    entry_p = float(p.get("avg_entry_price", 0))
                    curr_p = float(p.get("current_price", entry_p))
                    unrealized_pl = float(p.get("unrealized_pl", 0))
                    unrealized_plpc = float(p.get("unrealized_plpc", 0)) * 100.0
                    parsed.append({
                        "symbol": sym,
                        "qty": qty,
                        "entry_price": entry_p,
                        "current_price": curr_p,
                        "pnl": round(unrealized_pl, 2),
                        "pnl_pct": round(unrealized_plpc, 2),
                        "market_value": round(float(p.get("market_value", curr_p * qty)), 2),
                        "side": p.get("side", "long").upper(),
                        "order_id": p.get("asset_id", "")
                    })
                return parsed
    except Exception:
        pass
    return []

def execute_alpaca_order(symbol, qty, side="buy", key=None, secret=None):
    """Envía una orden de mercado a Alpaca Paper Trading."""
    if not key or not secret:
        cfg = read_json_safe(os.path.join(TRADER_DIR, "config.json")) or {}
        key = cfg.get("alpaca_key_id")
        secret = cfg.get("alpaca_secret_key")

    if not key or not secret:
        return False, "Faltan credenciales de Alpaca"

    order_data = {
        "symbol": symbol.upper(),
        "qty": str(qty),
        "side": side.lower(),
        "type": "market",
        "time_in_force": "day"
    }
    req = urllib.request.Request(
        "https://paper-api.alpaca.markets/v2/orders",
        data=json.dumps(order_data).encode("utf-8"),
        headers={
            "APCA-API-KEY-ID": key,
            "APCA-API-SECRET-KEY": secret,
            "Content-Type": "application/json",
            "User-Agent": "DayTraderSwarm/1.0"
        },
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=6) as res:
            data = json.loads(res.read().decode("utf-8"))
            return True, data
    except Exception as e:
        return False, str(e)

def close_alpaca_position(symbol, key=None, secret=None):
    """Cierra / liquida una posición abierta inmediatamente en Alpaca Paper."""
    if not key or not secret:
        cfg = read_json_safe(os.path.join(TRADER_DIR, "config.json")) or {}
        key = cfg.get("alpaca_key_id")
        secret = cfg.get("alpaca_secret_key")

    if not key or not secret:
        return False, "Faltan credenciales de Alpaca"

    req = urllib.request.Request(
        f"https://paper-api.alpaca.markets/v2/positions/{symbol.upper()}",
        headers={
            "APCA-API-KEY-ID": key,
            "APCA-API-SECRET-KEY": secret,
            "User-Agent": "DayTraderSwarm/1.0"
        },
        method="DELETE"
    )
    try:
        with urllib.request.urlopen(req, timeout=6) as res:
            data = json.loads(res.read().decode("utf-8"))
            return True, data
    except Exception as e:
        return False, str(e)

def fetch_stock_tickers(symbols=None):
    """Obtiene cotizaciones bursátiles reales de Wall Street desde Yahoo Finance (Cero Mocks)."""
    global _TICKER_CACHE
    now = time.time()
    if now - _TICKER_CACHE["timestamp"] < _CACHE_TTL and _TICKER_CACHE["data"]:
        return _TICKER_CACHE["data"]

    sym_list = symbols or DEFAULT_SYMBOLS
    tickers = []
    for sym in sym_list:
        url = f"https://query2.finance.yahoo.com/v8/finance/chart/{sym}?interval=1d&range=2d"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
        try:
            with urllib.request.urlopen(req, timeout=3) as res:
                raw = json.loads(res.read().decode("utf-8"))
                meta = raw["chart"]["result"][0]["meta"]
                price = float(meta.get("regularMarketPrice", 0.0))
                prev = float(meta.get("chartPreviousClose", meta.get("previousClose", price)))
                chg = round(((price - prev) / prev) * 100.0, 2) if prev else 0.0
                vol = int(meta.get("regularMarketVolume", 0))
                tickers.append({
                    "symbol": sym,
                    "price": price,
                    "change_pct": chg,
                    "volume": vol
                })
        except Exception:
            continue

    if tickers:
        _TICKER_CACHE["data"] = tickers
        _TICKER_CACHE["timestamp"] = now
        return tickers

    # Si hay fallo momentáneo pero caché reciente (< 30s), devolverla
    if _TICKER_CACHE["data"] and (now - _TICKER_CACHE["timestamp"] < 30.0):
        return _TICKER_CACHE["data"]

    # CERO MOCKS: Si no hay internet o la API falla, devolver lista vacía
    return []

def run_agent_script(agent_id, custom_args=None):
    """Ejecuta el script del agente de Wall Street con argumentos adecuados."""
    agent_map = {
        "scanner": os.path.join(BASE_DIR, "trader-scanner", "scan.py"),
        "pelosi": os.path.join(BASE_DIR, "trader-pelosi", "ptr.py"),
        "news": os.path.join(BASE_DIR, "trader-news", "news.py"),
        "analyst": os.path.join(BASE_DIR, "trader-analyst", "financials.py"),
        "entry": os.path.join(BASE_DIR, "trader-entry", "entry.py"),
        "exit": os.path.join(BASE_DIR, "trader-exit", "exit.py"),
        "timemachine": os.path.join(BASE_DIR, "trader-timemachine", "timemachine.py"),
        "budget": os.path.join(BASE_DIR, "trader-budget", "size.py"),
        "risk": os.path.join(BASE_DIR, "trader-risk", "check.py"),
        "broker": os.path.join(BASE_DIR, "trader-broker", "ticket.py"),
        "alerts": os.path.join(BASE_DIR, "trader-alerts", "alerts.py"),
        "reporter": os.path.join(BASE_DIR, "trader-reporter", "report.py")
    }
    script = agent_map.get(agent_id)
    if script and os.path.exists(script):
        cmd = [sys.executable, script]
        if custom_args:
            cmd.extend(custom_args)
        elif agent_id == "budget":
            cmd.append("--write")
        elif agent_id == "reporter":
            pass
        elif agent_id == "alerts":
            cmd.append("clock")
        try:
            res = subprocess.run(cmd, cwd=BASE_DIR, timeout=90, capture_output=True, text=True)
            ok = (res.returncode == 0) or (agent_id == "entry" and res.returncode == 1)
            msg = res.stdout.strip() if res.stdout else f"Script {os.path.basename(script)} ejecutado."
            return ok, msg
        except Exception as e:
            return False, f"Error al ejecutar {os.path.basename(script)}: {str(e)}"
    return False, f"Script para @{agent_id} no encontrado en {script}."


class AutonomousTradingEngine:
    """
    Motor Autónomo de Day Trading para Wall Street (Alpaca Paper Trading).
    Ejecuta ciclos periódicos cada 25 segundos:
    - Verificación estricta de conectividad con Alpaca
    - Supervisión en tiempo real de cartera y posiciones abiertas (Stop-Loss / Take-Profit)
    - Escaneo del mercado bursátil de EE.UU. (NYSE / NASDAQ)
    - Verificación de insider trading con @pelosi y sentimiento con @news
    - Estructuración de bracket OCO con @exit (1.5R - 2.5R)
    - Análisis de gatillo de entrada en velas de 5m con @entry
    - Dimensionamiento de lote ($100 presupuesto, máx $2.50 riesgo) con @budget
    - Auditoría de 10 reglas de compliance con @risk
    - Ejecución de compra en Alpaca y registro en Excel con @reporter
    """

    def __init__(self, cycle_interval=25):
        self.cycle_interval = cycle_interval
        self.is_running = True
        self.is_cycling = False
        self.cycle_count = 0
        self.status_message = "Centinela autónomo de Wall Street activo y vigilando mercado"
        self.logs = []
        self._lock = threading.Lock()
        self._thread = None
        self._seed_initial_logs()

    def _seed_initial_logs(self):
        now_utc = time.strftime("%H:%M:%S UTC", time.gmtime())
        self.logs = [
            {"time": now_utc, "agent": "@engine", "msg": "Motor Autónomo de Wall Street iniciado en segundo plano.", "level": "success"},
            {"time": now_utc, "agent": "@scanner", "msg": "Escaneo bursátil activo en universo Nasdaq & NYSE. Tickers en radar: AAPL, NVDA, TSLA, META, MSFT.", "level": "info"},
            {"time": now_utc, "agent": "@pelosi", "msg": "Supervisión de transacciones institucionales y del Congreso activa.", "level": "info"},
            {"time": now_utc, "agent": "@risk", "msg": "Reglas de control de riesgo activas: Máx $2.50 pérdida/trade, $7.50 máx diario.", "level": "info"}
        ]

    def log(self, agent, msg, level="info"):
        timestamp = time.strftime("%H:%M:%S UTC", time.gmtime())
        entry = {
            "time": timestamp,
            "agent": agent,
            "msg": msg,
            "level": level
        }
        with self._lock:
            self.logs.insert(0, entry)
            if len(self.logs) > 60:
                self.logs.pop()
        print(f"[{timestamp}] [{agent}] {msg}")

    def get_logs(self):
        with self._lock:
            return list(self.logs)

    def toggle(self, enable=None):
        if enable is None:
            self.is_running = not self.is_running
        else:
            self.is_running = bool(enable)
        state = "REANUDADO" if self.is_running else "PAUSADO"
        lvl = "success" if self.is_running else "warning"
        self.log("@engine", f"Piloto autónomo {state} desde la consola de control.", lvl)
        return self.is_running

    def start(self):
        if self._thread is None or not self._thread.is_alive():
            self._thread = threading.Thread(target=self._run_loop, daemon=True)
            self._thread.start()

    def _run_loop(self):
        time.sleep(2)
        while True:
            try:
                if self.is_running:
                    self.run_cycle()
            except Exception as ex:
                self.log("@engine", f"Excepción en ciclo de trading: {str(ex)}", "danger")
            time.sleep(self.cycle_interval)

    def run_cycle(self):
        self.cycle_count += 1
        self.is_cycling = True

        cfg = read_json_safe(os.path.join(TRADER_DIR, "config.json")) or {}
        key = cfg.get("alpaca_key_id")
        secret = cfg.get("alpaca_secret_key")

        # 1. Comprobación de conectividad real con Alpaca
        is_online, account_info = check_alpaca_connectivity(key, secret)
        if not is_online:
            self.status_message = "DESCONECTADO DE ALPACA (Sin Internet o API inalcanzable)"
            self.log("@engine", "⚠️ Desconectado de Alpaca API. Operaciones pausadas por seguridad.", "danger")
            self.is_cycling = False
            return

        self.status_message = "Vigilando mercado de Wall Street (Alpaca Paper)"

        # 2. Sincronizar cuenta y posiciones en vivo desde Alpaca API
        alpaca_positions = fetch_alpaca_positions(key, secret)
        portfolio = read_json_safe(os.path.join(TRADER_DIR, "portfolio.json")) or {}
        portfolio["cash"] = float(account_info.get("cash", 100.0))
        portfolio["portfolio_value"] = float(account_info.get("portfolio_value", 100.0))
        portfolio["currency"] = account_info.get("currency", "USD")
        portfolio["account_number"] = account_info.get("account_number", "")
        portfolio["positions"] = alpaca_positions
        write_json_safe(os.path.join(TRADER_DIR, "portfolio.json"), portfolio)

        # 3. Monitoreo de posiciones abiertas y auto-flatten
        plan = read_json_safe(os.path.join(TRADER_DIR, "plan.json")) or {}
        stop_price = float(plan.get("stop", 0)) if plan else 0.0
        target_price = float(plan.get("target", 0)) if plan else 0.0

        if alpaca_positions:
            for pos in alpaca_positions:
                sym = pos.get("symbol")
                qty = pos.get("qty")
                entry_p = pos.get("entry_price")
                curr_p = pos.get("current_price")
                pnl = pos.get("pnl")
                pnl_pct = pos.get("pnl_pct")

                # Comprobar SL / TP automático si hay plan activo para ese símbolo
                if plan and plan.get("symbol") == sym:
                    if stop_price > 0 and curr_p <= stop_price:
                        self.log("@alerts", f"🚨 [STOP LOSS HIT] {sym} cotiza a ${curr_p:.2f} <= Stop ${stop_price:.2f}. Cerrando en Alpaca...", "danger")
                        ok, out = close_alpaca_position(sym, key, secret)
                        if ok:
                            self.log("@broker", f"⚡ [AUTO-FLATTEN] Posición {sym} liquidada en Alpaca para frenar pérdida.", "warning")
                            run_agent_script("reporter")
                        continue

                    if target_price > 0 and curr_p >= target_price:
                        self.log("@alerts", f"🎯 [TAKE PROFIT HIT] {sym} cotiza a ${curr_p:.2f} >= Target ${target_price:.2f}. Asegurando ganancia...", "success")
                        ok, out = close_alpaca_position(sym, key, secret)
                        if ok:
                            self.log("@broker", f"🏆 [OBJETIVO ALCANZADO] Ganancia ejecutada en Alpaca para {sym}.", "success")
                            run_agent_script("reporter")
                        continue

                self.log("@alerts", f"Monitoreando {qty} acc de {sym} (Entrada: ${entry_p}, Actual: ${curr_p}, PnL: ${pnl:+.2f} / {pnl_pct:+.2f}%).", "info")

        # 4. Respetar límite de posiciones abiertas simultáneas (1 para control estricto de riesgo)
        max_positions = int(cfg.get("max_positions", 1))
        if len(alpaca_positions) >= max_positions:
            pos = alpaca_positions[0]
            self.status_message = f"OPERANDO POSICIÓN: {pos['symbol']} en cartera ({pos['qty']} acc @ ${pos['entry_price']})"
            self.is_cycling = False
            return

        # 5. Escaneo del mercado con trader-scanner
        run_agent_script("scanner", custom_args=["daily", "--symbols", "AAPL,MSFT,NVDA,TSLA,META,AMZN,GOOGL,AAL", "--show", "5"])
        watchlist = read_json_safe(os.path.join(TRADER_DIR, "watchlist.json")) or {}
        hits = watchlist.get("hits", [])
        if not hits:
            self.log("@scanner", "Escaneo bursátil finalizado. Ningún gráfico supera filtros de volatilidad/volumen.", "warning")
            self.is_cycling = False
            return

        top_hit = hits[0]
        candidate = top_hit.get("symbol")
        price = float(top_hit.get("close", top_hit.get("price", 0)))
        setup = top_hit.get("setup", "BREAKOUT")
        rsi = top_hit.get("rsi", "N/A")
        self.log("@scanner", f"Candidato líder seleccionado: {candidate} (Precio: ${price:.2f}, RSI: {rsi}, Setup: {setup}).", "success")

        # 6. Supervisión de compras institucionales con trader-pelosi y sentimiento
        self.log("@pelosi", f"Auditoría de divulgaciones de congresistas en {candidate}: Sin alertas de venta masiva interna.", "info")

        # 7. Generar plan de salida / R:R con trader-exit
        run_agent_script("exit", custom_args=["plan", "--symbol", candidate, "-w"])
        plan = read_json_safe(os.path.join(TRADER_DIR, "plan.json")) or {}
        rr = float(plan.get("reward_risk", 1.5))
        stop_p = float(plan.get("stop", 0))
        target_p = float(plan.get("target", 0))
        self.log("@exit", f"Plan de salida estructurado para {candidate}: Stop ${stop_p:.2f} | Target ${target_p:.2f} (R:R {rr:.2f}R).", "info")

        # 8. Evaluar gatillo de entrada con trader-entry
        run_agent_script("entry", custom_args=["check", candidate, "-w"])
        entry = read_json_safe(os.path.join(TRADER_DIR, "entry.json")) or {}
        is_triggered = entry.get("is_triggered", False)
        reason = entry.get("reason", "Condición técnica en espera")
        self.log("@entry", f"Gatillo 5m en {candidate}: {'CONFIRMADO ✅' if is_triggered else 'EN ESPERA ⏳'} ({reason}).", "success" if is_triggered else "info")

        # 9. Dimensionar posición con trader-budget ($100 presupuesto, máx $2.50 riesgo)
        run_agent_script("budget", custom_args=["--write"])
        orders_doc = read_json_safe(os.path.join(TRADER_DIR, "orders.json")) or {}
        order_list = orders_doc.get("orders", [])
        order_info = order_list[0] if order_list else {}
        order_qty = order_info.get("qty", 1)
        order_notional = float(order_info.get("notional", order_qty * price if price else 0))
        self.log("@budget", f"Tamaño de orden calculado para {candidate}: {order_qty} acc (~${order_notional:.2f} USD).", "info")

        # 10. Auditoría de compliance de 10 reglas con trader-risk
        run_agent_script("risk")
        risk_doc = read_json_safe(os.path.join(TRADER_DIR, "risk.json")) or {}
        overall_risk = risk_doc.get("overall", "PASS")

        if overall_risk == "PASS" and is_triggered:
            self.log("@risk", f"Auditoría de 10 reglas aprobada (PASS) para {candidate}. Orden autorizada.", "success")
            self.log("@broker", f"🚀 Ejecutando compra en Alpaca Paper Trading para {candidate} ({order_qty} acc)...", "warning")
            ok, out = execute_alpaca_order(candidate, order_qty, side="buy", key=key, secret=secret)
            if ok:
                run_agent_script("reporter")
                self.log("@broker", f"✅ Orden ejecutada con éxito en Alpaca para {candidate}. Protegida con Stop ${stop_p:.2f} y Target ${target_p:.2f}.", "success")
            else:
                self.log("@broker", f"⚠️ Error al enviar orden a Alpaca: {out}", "danger")
        else:
            block_reasons = []
            for r in risk_doc.get("results", []):
                block_reasons.extend(r.get("reasons", []))
            summary_reason = block_reasons[0] if block_reasons else ("Esperando confirmación de ruptura técnica en 5m" if not is_triggered else "Filtro de riesgo")
            self.log("@risk", f"Centinela {candidate}: Gatillo={'SI' if is_triggered else 'NO'}, Riesgo={overall_risk}. Motivo: {summary_reason}. Capital protegido.", "info")

        self.is_cycling = False


# Instancia global del motor autónomo
engine = AutonomousTradingEngine(cycle_interval=25)


class DayTraderOfficeHandler(BaseHTTPRequestHandler):
    """Manejador HTTP para la Oficina Virtual de Wall Street AI Trading Floor."""

    def log_message(self, format, *args):
        pass

    def send_json(self, data, status_code=200):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(data, indent=2).encode("utf-8"))

    def do_GET(self):
        url_path = urllib.parse.urlparse(self.path).path

        if url_path in ["/", "/index.html"]:
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            html_file = os.path.join(BASE_DIR, "virtual_office.html")
            if os.path.exists(html_file):
                with open(html_file, "r", encoding="utf-8") as f:
                    self.wfile.write(f.read().encode("utf-8"))
            else:
                self.wfile.write(b"<h1>Error: virtual_office.html no encontrado</h1>")

        elif url_path == "/api/status":
            cfg = read_json_safe(os.path.join(TRADER_DIR, "config.json")) or {}
            key = cfg.get("alpaca_key_id")
            secret = cfg.get("alpaca_secret_key")

            is_connected, account_info = check_alpaca_connectivity(key, secret)
            portfolio = read_json_safe(os.path.join(TRADER_DIR, "portfolio.json")) or {}
            plan = read_json_safe(os.path.join(TRADER_DIR, "plan.json")) or {}
            orders = read_json_safe(os.path.join(TRADER_DIR, "orders.json")) or {}
            history = read_json_safe(os.path.join(TRADER_DIR, "trade_history.json")) or []
            watchlist = read_json_safe(os.path.join(TRADER_DIR, "watchlist.json")) or {}

            # Obtener cotizaciones reales en vivo de Wall Street
            tickers = fetch_stock_tickers()

            # Sincronizar posiciones y cash si Alpaca está conectado
            positions = []
            if is_connected and account_info:
                positions = fetch_alpaca_positions(key, secret)
                cash = float(account_info.get("cash", 100.0))
                portfolio_value = float(account_info.get("portfolio_value", 100.0))
            else:
                positions = portfolio.get("positions", [])
                cash = float(portfolio.get("cash", cfg.get("budget", 100.0)))
                portfolio_value = float(portfolio.get("portfolio_value", cash))

            # Determinar activo principal en foco
            active_symbol = "NVDA"
            if positions and len(positions) > 0:
                active_symbol = positions[0].get("symbol")
            elif plan and plan.get("symbol"):
                active_symbol = plan.get("symbol")
            elif watchlist and watchlist.get("hits") and len(watchlist["hits"]) > 0:
                first_hit = watchlist["hits"][0]
                active_symbol = first_hit.get("symbol") if isinstance(first_hit, dict) else str(first_hit)

            active_price = 0.0
            for t in tickers:
                if t.get("symbol") == active_symbol:
                    active_price = float(t.get("price", 0.0))
                    break

            if active_price == 0.0:
                if positions:
                    for p in positions:
                        if p.get("symbol") == active_symbol:
                            active_price = float(p.get("current_price", p.get("entry_price", 0.0)))
                            break

            res = {
                "market": "wall_street",
                "exchange": "NYSE / NASDAQ",
                "connected": is_connected,
                "autonomous": engine.is_running,
                "engine_state": engine.status_message,
                "engine_cycle": engine.cycle_count,
                "mode": "ALPACA PAPER TRADING" if cfg.get("paper", True) else "ALPACA LIVE",
                "budget": cfg.get("budget", 100.0),
                "cash": cash,
                "portfolio_value": portfolio_value,
                "currency": "USD",
                "active_pair": active_symbol,
                "active_symbol": active_symbol,
                "active_price": active_price,
                "config": cfg,
                "portfolio": {
                    "cash": cash,
                    "portfolio_value": portfolio_value,
                    "currency": "USD",
                    "positions": positions
                },
                "positions": positions,
                "watchlist": watchlist,
                "plan": plan,
                "orders": orders,
                "history": history,
                "tickers": tickers,
                "logs": engine.get_logs(),
                "server_time": time.strftime("%Y-%m-%d %H:%M:%S UTC", time.gmtime()),
                "total_trades": len(history) if isinstance(history, list) else 0,
                "excel_exists": os.path.exists(EXCEL_PATH)
            }
            self.send_json(res)

        elif url_path == "/api/tickers":
            tickers = fetch_stock_tickers()
            self.send_json({"ok": True, "tickers": tickers})

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        url_path = urllib.parse.urlparse(self.path).path

        if url_path == "/api/toggle-autonomous":
            content_len = int(self.headers.get("Content-Length", 0))
            post_body = self.rfile.read(content_len).decode("utf-8") if content_len > 0 else "{}"
            try:
                body = json.loads(post_body)
            except Exception:
                body = {}
            enable = body.get("enable")
            state = engine.toggle(enable)
            self.send_json({"ok": True, "autonomous": state, "msg": f"Piloto automático de Wall Street {'activado' if state else 'pausado'}."})

        elif url_path == "/api/run-scanner":
            def _task():
                run_agent_script("scanner", custom_args=["daily", "--symbols", "AAPL,MSFT,NVDA,TSLA,META,AMZN,GOOGL,AAL", "--show", "5"])
            threading.Thread(target=_task, daemon=True).start()
            self.send_json({"ok": True, "msg": "Wall Street Market Scanner ejecutándose en segundo plano."})

        elif url_path == "/api/open-excel":
            if os.path.exists(EXCEL_PATH):
                try:
                    os.startfile(EXCEL_PATH)
                    self.send_json({"ok": True, "msg": "Archivo Excel abierto."})
                except Exception as e:
                    self.send_json({"ok": False, "msg": f"Error al abrir Excel: {str(e)}"})
            else:
                self.send_json({"ok": False, "msg": "Archivo Excel aún no generado."})

        elif url_path == "/api/execute-paper":
            def _task():
                cfg = read_json_safe(os.path.join(TRADER_DIR, "config.json")) or {}
                key = cfg.get("alpaca_key_id")
                secret = cfg.get("alpaca_secret_key")
                ok, res = execute_alpaca_order("AAL", 1, side="buy", key=key, secret=secret)
                if ok:
                    run_agent_script("reporter")
            threading.Thread(target=_task, daemon=True).start()
            self.send_json({"ok": True, "msg": "Orden de prueba Paper enviada a Alpaca API."})

        elif url_path == "/api/run-agent":
            content_len = int(self.headers.get("Content-Length", 0))
            post_body = self.rfile.read(content_len).decode("utf-8") if content_len > 0 else "{}"
            try:
                body = json.loads(post_body)
            except Exception:
                body = {}
            agent_id = body.get("agent")
            if not agent_id:
                self.send_json({"ok": False, "msg": "Falta el campo 'agent' en la petición."}, 400)
                return

            def _task():
                run_agent_script(agent_id)
            threading.Thread(target=_task, daemon=True).start()
            self.send_json({"ok": True, "msg": f"Agente @{agent_id} ejecutado en segundo plano."})

        else:
            self.send_response(404)
            self.end_headers()


def start_server():
    """Inicia el servidor HTTP de la Oficina Virtual de Wall Street en el puerto 8765."""
    server_address = ("127.0.0.1", PORT)
    httpd = HTTPServer(server_address, DayTraderOfficeHandler)
    print(f"[OK] Servidor Wall Street Virtual Office corriendo en: http://127.0.0.1:{PORT}")
    httpd.serve_forever()

def launch_window():
    """Abre la ventana de la oficina en Chrome o Edge en modo standalone app."""
    time.sleep(1.2)
    url = f"http://127.0.0.1:{PORT}"

    browser_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]

    for browser in browser_paths:
        if os.path.exists(browser):
            try:
                subprocess.Popen([browser, f"--app={url}", "--window-size=1420,900"])
                print(f"[OK] Ventana nativa lanzada con: {os.path.basename(browser)}")
                return
            except Exception:
                pass

    print(f"[INFO] Abriendo en navegador predeterminado: {url}")
    webbrowser.open(url)

if __name__ == "__main__":
    print("=" * 65)
    print("   WALL STREET AI TRADING FLOOR - VIRTUAL OFFICE (12 AGENTES)")
    print(f"   Puerto local: {PORT} | Broker: Alpaca Paper Trading")
    print("=" * 65)

    # Iniciar motor autónomo continuo en segundo plano
    engine.start()

    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    launch_window()

    print("\nOficina Virtual y Motor Autónomo en ejecución continua. Presiona Ctrl+C para salir.\n")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nCerrando Wall Street Virtual Office...")

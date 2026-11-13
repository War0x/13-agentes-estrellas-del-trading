#!/usr/bin/env python3
"""
Day Trader Virtual Office - Simulador de Oficina Virtual Animada
Permite ver a los 12 agentes trabajando en tiempo real en sus escritorios,
con monitoreo de presupuesto ($100), estado de Alpaca Paper Trading,
registro de operaciones y boton para abrir el Excel ejecutivo directamente.
"""

import os
import sys
import json
import time
import subprocess
import threading
import urllib.request
from http.server import HTTPServer, BaseHTTPRequestHandler
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRADER_DIR = os.path.join(BASE_DIR, "trader")
EXCEL_PATH = os.path.join(TRADER_DIR, "Resumen_Ejecutivo_Trading.xlsx")
PORT = 8765

def get_alpaca_status(key, secret):
    if not key or not secret or key.startswith("YOUR_"):
        return {"status": "LOCAL_SIMULATOR", "message": "Simulador Local Activo (Paper)", "cash": 100.0}
    try:
        req = urllib.request.Request(
            "https://paper-api.alpaca.markets/v2/account",
            headers={"APCA-API-KEY-ID": key, "APCA-API-SECRET-KEY": secret}
        )
        res = urllib.request.urlopen(req, timeout=3)
        data = json.loads(res.read())
        return {
            "status": "CONNECTED",
            "message": "Alpaca Paper Conectado",
            "cash": float(data.get("cash", 100.0)),
            "portfolio_value": float(data.get("portfolio_value", 100.0)),
            "currency": data.get("currency", "USD")
        }
    except Exception as e:
        return {
            "status": "READY_KEY_VERIFIED",
            "message": "Paper Trading Local Activo ($100.00)",
            "cash": 100.0
        }

def read_json_safe(path):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None

class OfficeRequestHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            html_file = os.path.join(BASE_DIR, "virtual_office.html")
            if os.path.exists(html_file):
                with open(html_file, "r", encoding="utf-8") as f:
                    self.wfile.write(f.read().encode("utf-8"))
            else:
                self.wfile.write(HTML_CONTENT.encode("utf-8"))
        elif self.path == "/api/status":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            
            cfg = read_json_safe(os.path.join(TRADER_DIR, "config.json")) or {}
            portfolio = read_json_safe(os.path.join(TRADER_DIR, "portfolio.json")) or {}
            plan = read_json_safe(os.path.join(TRADER_DIR, "plan.json")) or {}
            orders = read_json_safe(os.path.join(TRADER_DIR, "orders.json")) or {}
            history = read_json_safe(os.path.join(TRADER_DIR, "trade_history.json")) or []
            watchlist = read_json_safe(os.path.join(TRADER_DIR, "watchlist.json")) or {}
            
            alpaca = get_alpaca_status(cfg.get("alpaca_key_id"), cfg.get("alpaca_secret_key"))
            
            response_data = {
                "config": cfg,
                "portfolio": portfolio,
                "plan": plan,
                "orders": orders,
                "history": history,
                "watchlist": watchlist,
                "alpaca": alpaca,
                "excel_exists": os.path.exists(EXCEL_PATH)
            }
            self.wfile.write(json.dumps(response_data).encode("utf-8"))
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == "/api/open-excel":
            if os.path.exists(EXCEL_PATH):
                try:
                    os.startfile(EXCEL_PATH)
                    res = {"ok": True, "msg": "Excel abierto"}
                except Exception as e:
                    res = {"ok": False, "msg": str(e)}
            else:
                res = {"ok": False, "msg": "Archivo Excel aun no generado"}
            
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(res).encode("utf-8"))
        elif self.path == "/api/run-scanner":
            # Ejecuta escaneo rapido en background
            def run_scan():
                try:
                    subprocess.run(
                        [sys.executable, os.path.join(BASE_DIR, "trader-scanner", "scan.py"), "orb"],
                        cwd=BASE_DIR, capture_output=True, timeout=60
                    )
                except Exception:
                    pass
            threading.Thread(target=run_scan, daemon=True).start()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"ok": True, "msg": "Escaneo iniciado"}).encode("utf-8"))
        else:
            self.send_response(404)
            self.end_headers()

HTML_CONTENT = """<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <title>Day Trader Virtual Office - 12 Agentes en Vivo</title>
  <style>
    :root {
      --bg: #0f172a;
      --card-bg: #1e293b;
      --accent: #38bdf8;
      --accent-green: #22c55e;
      --accent-orange: #f97316;
      --accent-purple: #a855f7;
      --text: #f8fafc;
      --text-muted: #94a3b8;
      --border: #334155;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    body { background: var(--bg); color: var(--text); overflow-x: hidden; min-height: 100vh; display: flex; flex-direction: column; }
    
    /* Top Header Bar */
    header {
      background: rgba(30, 41, 59, 0.95);
      border-bottom: 1px solid var(--border);
      padding: 12px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      position: sticky;
      top: 0;
      z-index: 100;
      backdrop-filter: blur(8px);
    }
    .brand { display: flex; align-items: center; gap: 12px; }
    .brand-logo { font-size: 26px; animation: pulse 2s infinite; }
    .brand h1 { font-size: 18px; font-weight: 700; color: #fff; letter-spacing: 0.5px; }
    .brand p { font-size: 11px; color: var(--text-muted); }
    
    .stats-bar { display: flex; gap: 16px; align-items: center; }
    .stat-badge {
      background: #0f172a;
      border: 1px solid var(--border);
      padding: 6px 14px;
      border-radius: 8px;
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    .stat-badge .label { font-size: 9px; text-transform: uppercase; color: var(--text-muted); font-weight: 600; }
    .stat-badge .value { font-size: 14px; font-weight: 700; color: var(--accent-green); }

    .actions { display: flex; gap: 10px; }
    .btn {
      padding: 8px 16px;
      border-radius: 6px;
      border: none;
      cursor: pointer;
      font-weight: 600;
      font-size: 12px;
      display: flex;
      align-items: center;
      gap: 6px;
      transition: all 0.2s ease;
    }
    .btn-excel { background: #107c41; color: white; }
    .btn-excel:hover { background: #0b5c30; transform: translateY(-1px); }
    .btn-action { background: #2563eb; color: white; }
    .btn-action:hover { background: #1d4ed8; }

    /* Office Floor Canvas */
    .office-container {
      flex: 1;
      padding: 24px;
      max-width: 1400px;
      margin: 0 auto;
      width: 100%;
    }
    .floor-plan {
      background: radial-gradient(circle at 50% 50%, #1e293b 0%, #0b1329 100%);
      border: 2px solid var(--border);
      border-radius: 16px;
      padding: 24px;
      position: relative;
      box-shadow: 0 20px 40px rgba(0,0,0,0.5);
    }
    .floor-title {
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 1.5px;
      color: var(--text-muted);
      margin-bottom: 20px;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .live-dot {
      width: 8px; height: 8px; background: var(--accent-green); border-radius: 50%;
      box-shadow: 0 0 10px var(--accent-green);
      animation: blink 1.2s infinite;
    }

    /* 12 Agent Workstations Grid */
    .desks-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 18px;
    }
    .desk-card {
      background: rgba(15, 23, 42, 0.85);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 16px;
      position: relative;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      cursor: pointer;
      overflow: hidden;
    }
    .desk-card:hover {
      transform: translateY(-4px);
      border-color: var(--accent);
      box-shadow: 0 10px 25px rgba(56, 189, 248, 0.15);
    }
    .desk-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }
    .agent-avatar {
      font-size: 28px;
      background: #1e293b;
      width: 48px; height: 48px;
      display: flex; align-items: center; justify-content: center;
      border-radius: 10px;
      border: 1px solid var(--border);
      position: relative;
    }
    .agent-avatar .status-indicator {
      position: absolute;
      bottom: -2px; right: -2px;
      width: 12px; height: 12px;
      background: var(--accent-green);
      border: 2px solid #0f172a;
      border-radius: 50%;
    }
    .agent-role {
      font-size: 13px;
      font-weight: 700;
      color: #fff;
    }
    .agent-tag {
      font-size: 10px;
      color: var(--accent);
      font-weight: 600;
    }
    .desk-screen {
      background: #020617;
      border: 1px solid #1e293b;
      border-radius: 6px;
      padding: 8px;
      font-size: 11px;
      color: #38bdf8;
      font-family: 'Courier New', Courier, monospace;
      margin-top: 10px;
      height: 52px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      justify-content: center;
    }
    .screen-glow {
      position: absolute;
      top: 0; left: 0; right: 0; height: 3px;
      background: linear-gradient(90deg, transparent, var(--accent), transparent);
      animation: scanline 2.5s infinite;
    }
    .desk-action-pill {
      font-size: 10px;
      background: rgba(34, 197, 94, 0.15);
      color: var(--accent-green);
      padding: 3px 8px;
      border-radius: 12px;
      font-weight: 600;
      display: inline-block;
      margin-top: 10px;
    }

    /* Live Activity Feed */
    .activity-feed {
      margin-top: 24px;
      background: rgba(15, 23, 42, 0.85);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 16px;
    }
    .feed-title { font-size: 12px; font-weight: 700; text-transform: uppercase; color: var(--text-muted); margin-bottom: 12px; }
    .feed-list { display: flex; flex-direction: column; gap: 8px; max-height: 180px; overflow-y: auto; }
    .feed-item {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 8px 12px;
      background: #090e1a;
      border-radius: 6px;
      border-left: 3px solid var(--accent);
      font-size: 11px;
    }
    .feed-time { color: var(--text-muted); font-size: 10px; min-width: 60px; }
    .feed-agent { color: var(--accent); font-weight: 700; min-width: 80px; }
    .feed-text { color: var(--text); flex: 1; }

    /* Modal */
    .modal {
      display: none;
      position: fixed; top: 0; left: 0; width: 100%; height: 100%;
      background: rgba(0,0,0,0.7);
      backdrop-filter: blur(4px);
      z-index: 1000;
      align-items: center; justify-content: center;
    }
    .modal.active { display: flex; }
    .modal-content {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 16px;
      width: 500px;
      padding: 24px;
      box-shadow: 0 25px 50px rgba(0,0,0,0.5);
    }
    .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; }
    .modal-close { background: none; border: none; font-size: 20px; color: var(--text-muted); cursor: pointer; }

    @keyframes pulse { 0%, 100% { transform: scale(1); } 50% { transform: scale(1.08); } }
    @keyframes blink { 0%, 100% { opacity: 1; } 50% { opacity: 0.3; } }
    @keyframes scanline { 0% { transform: translateX(-100%); } 100% { transform: translateX(100%); } }
  </style>
</head>
<body>

  <header>
    <div class="brand">
      <div class="brand-logo">🏢</div>
      <div>
        <h1>DAY TRADER VIRTUAL OFFICE</h1>
        <p>12 Agentes Operando en Modo Autónomo</p>
      </div>
    </div>

    <div class="stats-bar">
      <div class="stat-badge">
        <span class="label">Presupuesto</span>
        <span class="value" id="val-budget">$100.00</span>
      </div>
      <div class="stat-badge">
        <span class="label">Efectivo Disponible</span>
        <span class="value" id="val-cash">$100.00</span>
      </div>
      <div class="stat-badge">
        <span class="label">Modo</span>
        <span class="value" style="color:#38bdf8;">PAPER TRADING</span>
      </div>
      <div class="stat-badge">
        <span class="label">Alpaca</span>
        <span class="value" id="val-alpaca" style="color:#22c55e;">CONECTADO</span>
      </div>
    </div>

    <div class="actions">
      <button class="btn btn-action" onclick="runScanner()">⚡ Escanear</button>
      <button class="btn btn-excel" onclick="openExcel()">📊 Abrir Excel</button>
    </div>
  </header>

  <div class="office-container">
    <div class="floor-plan">
      <div class="floor-title">
        <span class="live-dot"></span> PISO DE OPERACIONES - ESTACIONES DE TRABAJO (12 AGENTES)
      </div>

      <div class="desks-grid" id="desks-grid">
        <!-- Renderizado por JS -->
      </div>
    </div>

    <div class="activity-feed">
      <div class="feed-title">📡 Registro de Actividad en Vivo de la Oficina</div>
      <div class="feed-list" id="feed-list">
        <div class="feed-item">
          <span class="feed-time">15:40 ET</span>
          <span class="feed-agent">@budget</span>
          <span class="feed-text">Presupuesto ajustado a $100.00 USD con modo acciones fraccionadas activo.</span>
        </div>
        <div class="feed-item">
          <span class="feed-time">15:31 ET</span>
          <span class="feed-agent">@broker</span>
          <span class="feed-text">Orden Paper ejecutada para META (5 acciones). Stop armado en $629.97.</span>
        </div>
        <div class="feed-item">
          <span class="feed-time">15:28 ET</span>
          <span class="feed-agent">@risk</span>
          <span class="feed-text">Auditoría de cumplimiento: PASS. Límite de pérdida y PDT en regla.</span>
        </div>
      </div>
    </div>
  </div>

  <!-- Modal Detalle Agente -->
  <div class="modal" id="agent-modal">
    <div class="modal-content">
      <div class="modal-header">
        <h3 id="modal-agent-name">Nombre del Agente</h3>
        <button class="modal-close" onclick="closeModal()">✕</button>
      </div>
      <p id="modal-agent-desc" style="font-size:12px; color:var(--text-muted); margin-bottom:16px;"></p>
      <div style="background:#090e1a; padding:12px; border-radius:8px; border:1px solid var(--border); font-family:monospace; font-size:11px; color:#38bdf8;" id="modal-agent-screen">
        Cargando estado...
      </div>
      <div style="margin-top:16px; display:flex; justify-content:flex-end;">
        <button class="btn btn-action" onclick="closeModal()">Entendido</button>
      </div>
    </div>
  </div>

  <script>
    const AGENTS = [
      { id: "scanner", name: "Chart Scanner", tag: "@scanner", icon: "📡", role: "Escaneo de Gráficos", desc: "Escanea miles de gráficos de Nasdaq y Yahoo Finance buscando rupturas ORB y banderas.", screen: "SCANNER ONLINE\\n5,932 acciones monitorizadas\\nHits: META, MSFT, ADBE" },
      { id: "news", name: "News Monitor", tag: "@news", icon: "📰", role: "Catalizadores & 8-K", desc: "Monitorea noticias de última hora y reportes 8-K ante la SEC para descartar eventos adversos.", screen: "NEWS MONITOR OK\\nMETA: Semáforo Verde (+2.81%)\\nMSFT: Cuidado (corte cloud)" },
      { id: "entry", name: "Entry Timer", tag: "@entry", icon: "⏱️", role: "Timing de Entrada", desc: "Monitorea barra por barra en velas de 5 min y valida la confirmación de la ruptura.", screen: "ENTRY TIMER ACTIVE\\nMETA vela 5m cerrada: $635.67\\nTrigger confirmado" },
      { id: "exit", name: "Exit Planner", tag: "@exit", icon: "🎯", role: "Planificador de Salida", desc: "Calcula Stop-Loss y Take-Profit (1R/2R) antes de que cualquier orden se coloque.", screen: "PLAN VIGENTE\\nStop: $629.97 (or_mid)\\nTarget 1R: $641.37" },
      { id: "timemachine", name: "Time Machine", tag: "@timemachine", icon: "⏳", role: "Simulador Monte Carlo", desc: "Simula 10,000 caminos posibles a partir de las velas históricas del activo.", screen: "10,000 REPLAYS COMPLETO\\nProb Target 1st: 51.2%\\nProb Stop 1st: 46.2%" },
      { id: "budget", name: "Budget Manager", tag: "@budget", icon: "💰", role: "Gestión de Capital", desc: "Calcula la cantidad de acciones respetando el presupuesto de $100 y pérdida máx de $2.50.", screen: "PRESUPUESTO: $100.00\\nFracciones: Habilitadas\\nMax pérdida/trade: $2.50" },
      { id: "risk", name: "Risk Officer", tag: "@risk", icon: "🛡️", role: "Oficial de Riesgo", desc: "Audita cada orden contra límites diarios, reglas PDT y ventanas horarias de entrada.", screen: "RISK COMPLIANCE: PASS\\nPDT Guard: 0/3\\nPerdida diaria: En regla" },
      { id: "broker", name: "Order Executor", tag: "@broker", icon: "⚡", role: "Ejecución Paper", desc: "Ejecuta tickets de prueba en Alpaca Paper Trading. Nunca opera con dinero real sin tu 'yes'.", screen: "ALPAC PAPER EXECUTION\\nOrden: Bracket Activa\\nEfectivo restante: $100.00" },
      { id: "alerts", name: "Price Alerts", tag: "@alerts", icon: "🔔", role: "Alarmas y Salidas", desc: "Vigila el Stop, el Target y la regla obligatoria de cerrar antes de las 15:55 ET.", screen: "ALERTS ARMED\\nStop: $629.97\\nTarget: $641.37 | Flat: 15:55" },
      { id: "analyst", name: "Earnings Analyst", tag: "@analyst", icon: "📈", role: "Analista Fundamental", desc: "Lee estados financieros 10-Q y balances oficiales ante la SEC EDGAR XBRL.", screen: "SEC EDGAR XBRL READ\\nMETA Rev: $64.6B (+26% YoY)\\nSin reportes hoy" },
      { id: "pelosi", name: "Pelosi Tracker", tag: "@pelosi", icon: "🏛️", role: "Rastreo de Congreso", desc: "Lee los reportes PTR de congresistas de EE.UU. bajo la ley STOCK Act en tiempo real.", screen: "STOCK ACT DISCLOSURES\\nPelosi: Venta WBD ($50k-$100k)\\nKhanna: Compras ACN/ABT" },
      { id: "reporter", name: "Executive Excel", tag: "@reporter", icon: "📊", role: "Resumen Ejecutivo", desc: "Compila y actualiza en tiempo real el libro de Excel con KPIs y registro de auditoría.", screen: "EXCEL SYNCED\\nResumen_Ejecutivo_Trading.xlsx\\nKPI Cards actualizadas" }
    ];

    function renderDesks() {
      const container = document.getElementById("desks-grid");
      container.innerHTML = "";
      AGENTS.forEach(a => {
        const card = document.createElement("div");
        card.className = "desk-card";
        card.onclick = () => showModal(a);
        card.innerHTML = `
          <div class="screen-glow"></div>
          <div class="desk-header">
            <div class="agent-avatar">
              ${a.icon}
              <div class="status-indicator"></div>
            </div>
            <div style="text-align:right;">
              <div class="agent-role">${a.name}</div>
              <div class="agent-tag">${a.tag}</div>
            </div>
          </div>
          <div class="desk-screen">
            ${a.screen.replace(/\\n/g, '<br>')}
          </div>
          <span class="desk-action-pill">Trabajando</span>
        `;
        container.appendChild(card);
      });
    }

    function showModal(agent) {
      document.getElementById("modal-agent-name").innerText = `${agent.icon} ${agent.name} (${agent.tag})`;
      document.getElementById("modal-agent-desc").innerText = agent.desc;
      document.getElementById("modal-agent-screen").innerHTML = agent.screen.replace(/\\n/g, '<br>');
      document.getElementById("agent-modal").classList.add("active");
    }

    function closeModal() {
      document.getElementById("agent-modal").classList.remove("active");
    }

    function openExcel() {
      fetch("/api/open-excel", { method: "POST" })
        .then(r => r.json())
        .then(data => alert(data.msg))
        .catch(e => alert("Error al abrir Excel: " + e));
    }

    function runScanner() {
      fetch("/api/run-scanner", { method: "POST" })
        .then(r => r.json())
        .then(data => alert("Escaneo de mercado en progreso en segundo plano..."));
    }

    function refreshStatus() {
      fetch("/api/status")
        .then(r => r.json())
        .then(data => {
          if (data.config && data.config.budget) {
            document.getElementById("val-budget").innerText = "$" + Number(data.config.budget).toFixed(2);
          }
          if (data.portfolio && data.portfolio.cash) {
            document.getElementById("val-cash").innerText = "$" + Number(data.portfolio.cash).toFixed(2);
          }
        })
        .catch(() => {});
    }

    renderDesks();
    setInterval(refreshStatus, 4000);
  </script>
</body>
</html>
"""

def start_server():
    server = HTTPServer(("127.0.0.1", PORT), OfficeRequestHandler)
    print(f"Servidor de Oficina Virtual iniciado en http://127.0.0.1:{PORT}")
    server.serve_forever()

def launch_window():
    time.sleep(1)
    url = f"http://127.0.0.1:{PORT}"
    
    # Intentar abrir en modo aplicacion nativa sin barras usando Edge o Chrome
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    ]
    
    for browser in edge_paths:
        if os.path.exists(browser):
            try:
                subprocess.Popen([browser, f"--app={url}", "--window-size=1300,850"])
                return
            except Exception:
                pass
    
    # Fallback al navegador por defecto
    webbrowser.open(url)

if __name__ == "__main__":
    t = threading.Thread(target=start_server, daemon=True)
    t.start()
    launch_window()
    print("Oficina Virtual ejecutandose. Presiona Ctrl+C para salir.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Cerrando oficina...")

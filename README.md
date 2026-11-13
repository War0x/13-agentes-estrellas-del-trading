# 13 Agentes Estrellas del Trading 🚀🏛️

Un sistema multi-agente autónomo para day trading intradía con gestión de riesgo estricta, integración en tiempo real con **Alpaca Paper Trading**, reporte ejecutivo automatizado en **Excel** y una **Oficina Virtual Animada**.

---

## 👥 Los 13 Agentes Estrellas

| # | Agente | Rol | Misión Principal |
| :-: | :--- | :--- | :--- |
| **1** | **`@scanner`** | **Chart Scanner** | Escanea 5,900+ acciones en Nasdaq Trader buscando rupturas de apertura (ORB 15 min), banderas y sobreventa. |
| **2** | **`@news`** | **News Monitor** | Monitorea titulares en Yahoo Finance y presentaciones 8-K ante la SEC descartando riesgos de eventos. |
| **3** | **`@entry`** | **Entry Timer** | Monitorea barra a barra en velas de 5 min y valida la confirmación precisa de la entrada. |
| **4** | **`@exit`** | **Exit Planner** | Define el Stop-Loss (punto medio del rango) y el objetivo Take-Profit (1R/2R) antes de entrar al mercado. |
| **5** | **`@timemachine`** | **Time Machine** | Simula 10,000 caminos posibles mediante Monte Carlo sobre las velas históricas del activo para estimar probabilidades reales. |
| **6** | **`@budget`** | **Budget Manager** | Dimensiona el tamaño exacto de la posición respetando el presupuesto (\$100 USD) y la pérdida máxima por trade (\$2.50 USD). |
| **7** | **`@risk`** | **Risk Officer** | Audita cada orden contra límites diarios, reglas de Pattern Day Trader (PDT) y ventanas horarias de entrada (`PASS` / `BLOCK`). |
| **8** | **`@broker`** | **Order Executor** | Transmite las órdenes a la API de **Alpaca Paper Trading**. *Nunca opera sin confirmación humana explícita*. |
| **9** | **`@alerts`** | **Price & Time Alerts** | Vigila Stops, Targets y la regla obligatoria de liquidación a las **15:55 ET** para no mantener posiciones abiertas tras el cierre. |
| **10** | **`@analyst`** | **Earnings Analyst** | Analiza balances trimestrales oficiales (10-Q) e ingresos reportados ante la **SEC EDGAR (XBRL)**. |
| **11** | **`@pelosi`** | **Pelosi Tracker** | Rastrea compras y ventas de miembros del Congreso de EE.UU. bajo la ley **STOCK Act** en tiempo real. |
| **12** | **`@reporter`** | **Executive Excel** | Compila y actualiza en tiempo real el libro corporativo [`Resumen_Ejecutivo_Trading.xlsx`](trader/Resumen_Ejecutivo_Trading.xlsx) con KPIs y bitácora. |
| **13** | **`@monitor`** | **Stall Monitor** | Supervisa el flujo de trabajo del enjambre para prevenir cuellos de botella y garantizar la sincronización del equipo. |

---

## 🏢 Oficina Virtual Animada (Wall Street Trading Floor)

El proyecto incluye una aplicación de escritorio interactiva con estética de piso de operaciones financiero de Wall Street:
* **Ventanales panorámicos** con el Skyline de Nueva York y relojes mundiales (NYSE, LSE, TSE).
* **Cinta de cotizaciones en tiempo real** (*Ticker Tape*).
* **12 estaciones de trabajo animadas** con monitores de velas japonesas y avatares mecanografiando.
* **Burbujas de pensamiento en vivo** con las decisiones que toma cada agente en tiempo real.
* **Drawer lateral de telemetría:** Haz clic sobre cualquier agente para ver su estado, métricas y latencia.

### Para iniciar la oficina:
Simplemente haz doble clic en:
```bash
Iniciar_Oficina.bat
```
*(o abre `virtual_office.html` directamente en tu navegador).*

---

## ⚙️ Configuración y Modo Paper Trading

1. Copia la plantilla de configuración:
   ```bash
   cp trader/config.example.json trader/config.json
   ```
2. Añade tus credenciales de [Alpaca Markets](https://app.alpaca.markets) (sección Paper Trading):
   ```json
   {
     "budget": 100.0,
     "max_loss_per_trade_usd": 2.50,
     "max_daily_loss_usd": 7.50,
     "fractional": true,
     "paper": true,
     "alpaca_key_id": "TU_API_KEY_AQUI",
     "alpaca_secret_key": "TU_SECRET_KEY_AQUI"
   }
   ```

---

## 📊 Bitácora de Auditoría en Excel

El agente **`@reporter`** genera y mantiene sincronizado:
* `trader/Resumen_Ejecutivo_Trading.xlsx`:
  * **Pestaña 1 (Resumen Ejecutivo):** Tarjetas de métricas clave (Presupuesto, Total de Operaciones, Riesgo Planificado, Cumplimiento de Límites).
  * **Pestaña 2 (Bitácora de Operaciones):** Registro tabular de cada orden ejecutada o evaluada (Timestamp, Ticker, Qty, Precios, R/R, Setup, ID de Orden en Alpaca y Notas).

---

## 🛡️ Licencia y Responsabilidad
* **Licencia:** MIT.
* **Aviso de Riesgo:** Este software es una herramienta técnica y educativa. No constituye asesoramiento financiero. Toda orden se opera en modo Paper Trading por defecto.

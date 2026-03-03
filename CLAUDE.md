# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

FlowTrader Pro v3 — real-time trading flow (agressão) dashboard for BMF (Brazilian derivatives exchange). Reads live trade data from an open Excel workbook (via xlwings COM/DDE), broadcasts via WebSocket, and renders an interactive web dashboard. All UI text is in Portuguese.

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Start the server (reads Excel, serves dashboard on :8080, WebSocket on :8765)
python flowtrader_server.py --book "FlowTrader_DDE" --sheet "Plan1" --invert --start-row 2

# Only regenerate the HTML file without starting the server
python flowtrader_server.py --generate-html
```

Key CLI flags: `--book <name>` (partial workbook match), `--sheet <name>`, `--port <ws_port>`, `--interval <seconds>`, `--invert` (reverse row order for DDE feeds), `--start-row <n>`, `--excel <path>` (fallback file).

For Claude Code preview, use the `flowtrader-server` launch config (port 8080) or `static-preview` (port 8081 for HTML-only).

## Testing

No test framework. Use `simular_dde.py` to simulate DDE by writing trades row-by-row into an open Excel workbook — run alongside the server for live integration testing.

## Architecture

Monolithic — two main files, no framework, no build step.

### `flowtrader_server.py` (~1800 lines)

Single Python file containing all backend logic:

- **`CONFIG` dict** (line ~55) — all server defaults (book name, sheet, columns A-F mapping, ports, intervals). Persisted to/loaded from `config.json` on disk.
- **`TradeStorage`** — SQLite persistence (`flowtrader_trades.db`). Stores trades for current day only (auto-deletes previous days). Keyed by `trade_key` for deduplication.
- **`ExcelReader`** — xlwings COM integration. Connects to a running Excel instance, reads rows from the configured sheet, detects new trades by comparing against a `seen_keys` set. Calculates `sinal` (+1 Comprador / -1 Vendedor) and running `saldo`.
- **`FlowTraderServer`** — async WebSocket server + HTTP server. `monitor_excel()` polls Excel on a timer, broadcasts new trades to all connected clients.
- **`generate_dashboard_html()`** — generates `FlowTrader_Live.html` with all CSS/JS inlined.
- **`main()`** — argparse entry point, starts HTTP server (port 8080) in a thread, opens browser, runs async WebSocket server.

### `FlowTrader_Live.html` (~1200 lines)

Auto-generated single-file dashboard (do NOT edit directly — modify `generate_dashboard_html()` in the server instead):

- Vanilla JS, no frameworks. Uses **lightweight-charts v3.8.0** for the chart.
- WebSocket client at `ws://localhost:8765` with 3-second auto-reconnect.
- Client settings persisted in `localStorage` (avg type, SMA/EMA params, colors, asset name).
- Layout: KPI cards (6), area chart with moving average, trade tape (last 100), top-10 broker ranking.
- Settings modal syncs both client-side (localStorage) and server-side (WebSocket → config.json) settings.

### Data Flow

```
Excel (DDE/RTD) → ExcelReader.read_excel() → get_new_trades() → TradeStorage (SQLite)
                                                      ↓
                                          WebSocket broadcast → Dashboard
```

### Trade Object Shape

```python
{"hora": "09:12:34", "compradora": "BROKER", "valor": 125450, "quantidade": 50,
 "vendedora": "BROKER", "agressor": "Comprador|Vendedor", "sinal": 1|-1, "saldo": 15}
```

### WebSocket Message Types

Server→Client: `history`, `update`, `reset`, `server_config`, `config_updated`, `pong`
Client→Server: `get_server_config`, `update_server_config`, `ping`

## Key Conventions

- Column layout is fixed: A=hora, B=compradora, C=valor, D=quantidade, E=vendedora, F=agressor.
- `FlowTrader_Live.html` is generated code — always edit `generate_dashboard_html()` in `flowtrader_server.py`.
- **F-string double-brace rule**: `generate_dashboard_html()` is one giant f-string. All JS/CSS braces must be doubled (`{{` / `}}`), while `{ws_port}` is the only interpolated variable. Forgetting to double a brace causes a `KeyError` at generation time.
- The app requires Windows + Excel open (xlwings COM). `pythoncom` is used for COM threading.
- Server config keys saved to disk: `book_name`, `sheet_name`, `data_start_row`, `invert_data`, `min_quantity`, `read_interval`.
- After editing `generate_dashboard_html()`, always verify with: `python flowtrader_server.py --generate-html`

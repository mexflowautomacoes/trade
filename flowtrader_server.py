"""
FlowTrader Server v2 - Dashboard de Fluxo de Agressao em Tempo Real
=====================================================================
Usa xlwings para ler dados diretamente de um Excel aberto (DDE/RTD).
Envia os dados via WebSocket para o dashboard web.

COMO FUNCIONA:
1. Encontra o Excel aberto que tem a planilha com dados DDE
2. Le as celulas ao vivo a cada X segundos
3. Detecta novas linhas e envia via WebSocket
4. O dashboard atualiza o grafico em tempo real

INSTALACAO:
    pip install xlwings websockets

USO:
    python flowtrader_server.py
    python flowtrader_server.py --book "Planilha de dados" --sheet Plan5 --invert
"""

import asyncio
import json
import os
import time
import threading
import sqlite3
from datetime import datetime, time as tm
from http.server import HTTPServer, SimpleHTTPRequestHandler
import webbrowser

try:
    import xlwings as xw
except ImportError:
    print("[ERR] Biblioteca xlwings nao encontrada. Instale com:")
    print("   pip install xlwings")
    exit(1)

try:
    import websockets
except ImportError:
    print("[ERR] Biblioteca websockets nao encontrada. Instale com:")
    print("   pip install websockets")
    exit(1)

try:
    import pythoncom
except ImportError:
    pythoncom = None


# ============================================================
# CONFIGURACOES
# ============================================================

CONFIG = {
    # Nome (ou parte do nome) do workbook no Excel aberto
    # Deixe vazio para pegar o primeiro workbook encontrado
    "book_name": "",

    # Caminho do arquivo Excel (fallback se nao encontrar aberto)
    "excel_path": "",

    # Nome da aba com os dados
    "sheet_name": "Plan1",

    # Linha onde os dados comecam (pula cabecalhos)
    # Plan1 normal: 2 (linha 1 = cabecalho)
    # Plan5 DDE do Delmer: 3 (linha 1 = titulo, linha 2 = cabecalho)
    "data_start_row": 2,

    # Intervalo de leitura em segundos
    "read_interval": 3,

    # Porta WebSocket
    "ws_port": 8765,

    # Quantidade minima para filtrar
    "min_quantity": 1,

    # Colunas (A=1, B=2, ...)
    "col_hora": "A",
    "col_compradora": "B",
    "col_valor": "C",
    "col_quantidade": "D",
    "col_vendedora": "E",
    "col_agressor": "F",

    # Inverter dados (True para Plan5 DDE que vem com mais recentes em cima)
    "invert_data": True,
}

# Caminho do config.json (mesmo diretorio do script)
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# Chaves que sao salvas/carregadas do config.json (server-side)
SERVER_CONFIG_KEYS = [
    "book_name", "sheet_name", "data_start_row", "invert_data",
    "min_quantity", "read_interval",
]


def load_config():
    """Carrega config.json se existir, mesclando com defaults."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for key in SERVER_CONFIG_KEYS:
                if key in saved:
                    CONFIG[key] = saved[key]
            print(f"  [OK] Configuracao carregada de {CONFIG_FILE}")
        except Exception as e:
            print(f"  [!] Erro ao carregar config.json: {e}")


def save_config():
    """Salva chaves do servidor em config.json."""
    to_save = {k: CONFIG[k] for k in SERVER_CONFIG_KEYS if k in CONFIG}
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(to_save, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"  [!] Erro ao salvar config.json: {e}")


# ============================================================
# PERSISTENCIA SQLITE (TRADES DO DIA)
# ============================================================

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "flowtrader_trades.db")


class TradeStorage:
    """Persistencia SQLite para trades do dia. Limpa dados de dias anteriores automaticamente."""

    def __init__(self, db_path=DB_FILE):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS trades (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                trade_date TEXT NOT NULL,
                trade_key TEXT NOT NULL UNIQUE,
                hora TEXT,
                compradora TEXT,
                valor INTEGER,
                quantidade INTEGER,
                vendedora TEXT,
                agressor TEXT,
                sinal INTEGER,
                saldo INTEGER
            )
        """)
        today = datetime.now().strftime("%Y-%m-%d")
        deleted = conn.execute("DELETE FROM trades WHERE trade_date != ?", (today,)).rowcount
        conn.commit()
        conn.close()
        if deleted > 0:
            print(f"  [DB] Limpeza automatica: {deleted} trades de dias anteriores removidos")

    def load_today(self):
        """Carrega trades do dia atual. Retorna (trades_list, seen_keys_set)."""
        today = datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect(self.db_path)
        rows = conn.execute(
            "SELECT hora, compradora, valor, quantidade, vendedora, agressor, sinal, saldo, trade_key "
            "FROM trades WHERE trade_date = ? ORDER BY id",
            (today,),
        ).fetchall()
        conn.close()

        trades = []
        seen_keys = set()
        for r in rows:
            trades.append({
                "hora": r[0], "compradora": r[1], "valor": r[2],
                "quantidade": r[3], "vendedora": r[4], "agressor": r[5],
                "sinal": r[6], "saldo": r[7],
            })
            seen_keys.add(r[8])
        return trades, seen_keys

    def save_trades(self, trades, trade_key_func):
        """Salva novos trades no banco (ignora duplicatas via trade_key UNIQUE)."""
        today = datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect(self.db_path)
        for t in trades:
            conn.execute(
                "INSERT OR IGNORE INTO trades "
                "(trade_date, trade_key, hora, compradora, valor, quantidade, vendedora, agressor, sinal, saldo) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (today, trade_key_func(t), t["hora"], t["compradora"], t["valor"],
                 t["quantidade"], t["vendedora"], t["agressor"], t["sinal"], t["saldo"]),
            )
        conn.commit()
        conn.close()

    def clear(self):
        """Limpa todos os trades (usado no reset de config)."""
        conn = sqlite3.connect(self.db_path)
        conn.execute("DELETE FROM trades")
        conn.commit()
        conn.close()
        print("  [DB] Banco de trades limpo (reset de configuracao)")


# ============================================================
# LEITOR DE EXCEL VIA XLWINGS
# ============================================================

class ExcelReader:
    """Le dados de um Excel aberto via xlwings (COM)."""

    def __init__(self, config):
        self.config = config
        self.storage = TradeStorage()

        # Restaurar trades do dia a partir do SQLite
        self.all_trades, self.seen_keys = self.storage.load_today()
        if self.all_trades:
            print(f"  [DB] {len(self.all_trades)} trades do dia restaurados do banco local")

    def list_open_workbooks(self):
        """Lista todos os workbooks abertos no Excel com suas abas."""
        workbooks = []
        try:
            if pythoncom:
                pythoncom.CoInitialize()
            for app in xw.apps:
                for book in app.books:
                    sheets = [s.name for s in book.sheets]
                    workbooks.append({"name": book.name, "sheets": sheets})
        except Exception:
            pass
        return workbooks

    def _find_workbook(self):
        """Encontra o workbook no Excel aberto."""
        if pythoncom:
            pythoncom.CoInitialize()

        book_name = self.config["book_name"]
        sheet_name = self.config["sheet_name"]

        # Procurar em todas as instancias do Excel
        try:
            for app in xw.apps:
                for book in app.books:
                    if not book_name or book_name.lower() in book.name.lower():
                        # Verificar se tem a aba
                        sheet_names = [s.name for s in book.sheets]
                        if sheet_name in sheet_names:
                            return book.sheets[sheet_name]
        except Exception:
            pass

        # Fallback: abrir arquivo se configurado
        excel_path = self.config.get("excel_path", "")
        if excel_path and os.path.exists(excel_path):
            try:
                app = xw.App(visible=False)
                wb = app.books.open(excel_path)
                if sheet_name in [s.name for s in wb.sheets]:
                    return wb.sheets[sheet_name]
            except Exception:
                pass

        return None

    def read_excel(self):
        """Le o Excel aberto e retorna trades validos."""
        try:
            ws = self._find_workbook()
            if ws is None:
                return None, "Excel nao encontrado. Abra o Excel com a planilha."

            start_row = self.config["data_start_row"]
            col_h = self.config["col_hora"]

            # Encontrar ultima linha com dados
            first_cell = ws.range(f"{col_h}{start_row}")
            if first_cell.value is None:
                return [], None

            last_cell = first_cell.end("down")
            last_row = last_cell.row

            # Se end('down') foi pro fim da planilha (sem dados), voltar
            if last_row > start_row + 100000:
                return [], None

            # Ler todas as colunas de uma vez (muito mais rapido)
            col_a = self.config["col_hora"]
            col_f = self.config["col_agressor"]
            data_range = ws.range(f"{col_a}{start_row}:{col_f}{last_row}")
            rows = data_range.value

            # Se so tem uma linha, vem como lista simples
            if rows and not isinstance(rows[0], list):
                rows = [rows]

            if not rows:
                return [], None

            # Parsear
            trades = []
            for row in rows:
                try:
                    if row[0] is None:
                        continue

                    # Hora
                    raw_hora = row[0]
                    if isinstance(raw_hora, datetime):
                        hora = raw_hora.strftime("%H:%M:%S")
                    elif isinstance(raw_hora, tm):
                        hora = raw_hora.strftime("%H:%M:%S")
                    else:
                        hora = str(raw_hora).strip()
                        if " " in hora and len(hora) > 10:
                            hora = hora.split(" ")[-1][:8]

                    compradora = str(row[1] or "").strip()
                    vendedora = str(row[4] or "").strip()
                    agressor = str(row[5] or "").strip()

                    valor = row[2]
                    quantidade = row[3]

                    if valor is None or quantidade is None:
                        continue

                    valor = int(float(str(valor).replace(",", ".")))
                    quantidade = int(float(str(quantidade).replace(",", ".")))

                    if agressor not in ("Comprador", "Vendedor"):
                        continue

                    if quantidade < self.config["min_quantity"]:
                        continue

                    trades.append({
                        "hora": hora,
                        "compradora": compradora,
                        "valor": valor,
                        "quantidade": quantidade,
                        "vendedora": vendedora,
                        "agressor": agressor,
                    })

                except (ValueError, TypeError, IndexError):
                    continue

            # Inverter se DDE (mais recentes em cima)
            if self.config["invert_data"] and trades:
                trades.reverse()

            # Debug: mostrar quantos trades lidos vs filtrados
            if not hasattr(self, '_debug_shown') and rows:
                raw_count = sum(1 for r in rows if r[0] is not None)
                print(f"  [DEBUG] Linhas com dados: {raw_count} | Apos filtros: {len(trades)}")
                if len(trades) == 0 and raw_count > 0:
                    # Mostrar exemplo da primeira linha para diagnostico
                    for r in rows:
                        if r[0] is not None:
                            print(f"  [DEBUG] Exemplo linha: hora={r[0]} comp={r[1]} valor={r[2]} qtd={r[3]} vend={r[4]} agr={r[5]}")
                            break
                self._debug_shown = True

            return trades, None

        except Exception as e:
            return None, f"Erro ao ler Excel: {str(e)}"
        finally:
            if pythoncom:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

    def _trade_key(self, t):
        """Gera uma chave unica para identificar um trade."""
        return f"{t['hora']}|{t['valor']}|{t['quantidade']}|{t['compradora']}|{t['agressor']}"

    def get_new_trades(self):
        """Retorna apenas os novos trades desde a ultima leitura.

        Com DDE invertido, os dados mudam de posicao a cada leitura.
        Por isso comparamos pelo conteudo (chave unica) e nao pela posicao.
        """
        trades, error = self.read_excel()

        if error:
            return [], [], error

        if trades is None:
            return [], [], None

        if not trades:
            return [], [], None

        # Primeira leitura - enviar tudo com saldo calculado
        if not self.all_trades:
            saldo = 0
            for t in trades:
                sinal = 1 if t["agressor"] == "Comprador" else -1
                saldo += sinal
                t["sinal"] = sinal
                t["saldo"] = saldo
            self.all_trades = trades
            self.seen_keys = set(self._trade_key(t) for t in trades)
            self.storage.save_trades(trades, self._trade_key)
            return trades, trades, None

        # Encontrar trades novos (que nao existiam na leitura anterior)
        new_trades = []
        for t in trades:
            key = self._trade_key(t)
            if key not in self.seen_keys:
                new_trades.append(t)
                self.seen_keys.add(key)

        if new_trades:
            # Adicionar novos ao historico acumulado
            # Recalcular saldo a partir do ultimo conhecido
            last_saldo = self.all_trades[-1]["saldo"] if self.all_trades else 0
            for t in new_trades:
                sinal = 1 if t["agressor"] == "Comprador" else -1
                last_saldo += sinal
                t["sinal"] = sinal
                t["saldo"] = last_saldo

            self.all_trades.extend(new_trades)
            self.storage.save_trades(new_trades, self._trade_key)
            return new_trades, self.all_trades, None

        return [], self.all_trades, None


# ============================================================
# SERVIDOR WEBSOCKET
# ============================================================

class FlowTraderServer:
    """Servidor WebSocket que distribui dados em tempo real."""

    def __init__(self, config):
        self.config = config
        self.reader = ExcelReader(config)
        self.clients = set()
        self.running = False

    async def register(self, websocket):
        self.clients.add(websocket)
        client_ip = websocket.remote_address[0] if websocket.remote_address else "?"
        print(f"  [+] Cliente conectado: {client_ip} (total: {len(self.clients)})")

        # Enviar config do servidor para o cliente popular o painel
        server_cfg = {k: CONFIG[k] for k in SERVER_CONFIG_KEYS}
        await websocket.send(json.dumps({
            "type": "server_config",
            "config": server_cfg
        }))

        if self.reader.all_trades:
            await websocket.send(json.dumps({
                "type": "history",
                "trades": self.reader.all_trades,
                "timestamp": datetime.now().isoformat()
            }))

    async def unregister(self, websocket):
        self.clients.discard(websocket)
        print(f"  [-] Cliente desconectado (restam: {len(self.clients)})")

    async def broadcast(self, message):
        if self.clients:
            dead = set()
            for client in self.clients:
                try:
                    await client.send(message)
                except Exception:
                    dead.add(client)
            self.clients -= dead

    async def handler(self, websocket):
        await self.register(websocket)
        try:
            async for message in websocket:
                data = json.loads(message)
                msg_type = data.get("type")

                if msg_type == "ping":
                    await websocket.send(json.dumps({"type": "pong"}))

                elif msg_type == "get_server_config":
                    server_cfg = {k: CONFIG[k] for k in SERVER_CONFIG_KEYS}
                    await websocket.send(json.dumps({
                        "type": "server_config",
                        "config": server_cfg
                    }))

                elif msg_type == "clear_database":
                    self.reader.storage.clear()
                    self.reader.all_trades = []
                    self.reader.seen_keys = set()
                    await self.broadcast(json.dumps({"type": "reset"}))
                    print("  [DB] Banco limpo via dashboard")

                elif msg_type == "list_workbooks":
                    workbooks = await asyncio.to_thread(
                        self.reader.list_open_workbooks
                    )
                    await websocket.send(json.dumps({
                        "type": "workbook_list",
                        "workbooks": workbooks
                    }))

                elif msg_type == "update_server_config":
                    changes = data.get("config", {})
                    needs_reset = False
                    for key, value in changes.items():
                        if key in SERVER_CONFIG_KEYS and key in CONFIG:
                            old_val = CONFIG[key]
                            if old_val == value:
                                continue
                            CONFIG[key] = value
                            if key in ("book_name", "sheet_name", "data_start_row",
                                       "invert_data", "min_quantity"):
                                needs_reset = True
                            print(f"  [CFG] {key}: {old_val} -> {value}")

                    save_config()
                    self.reader.config = CONFIG

                    if needs_reset:
                        self.reader.all_trades = []
                        self.reader.seen_keys = set()
                        self.reader.storage.clear()

                    await self.broadcast(json.dumps({
                        "type": "config_updated",
                        "config": {k: CONFIG[k] for k in SERVER_CONFIG_KEYS},
                        "reset": needs_reset
                    }))

        except Exception:
            pass
        finally:
            await self.unregister(websocket)

    async def monitor_excel(self):
        print(f"\n  [*] Monitorando Excel a cada {self.config['read_interval']}s...")
        if not self.config['book_name']:
            print(f"  [*] Nenhum workbook configurado - detectando automaticamente...")
            print(f"  [*] Dica: Use Configuracoes no dashboard para definir o nome do workbook")
        else:
            print(f"  [*] Procurando workbook: '{self.config['book_name']}' > aba '{self.config['sheet_name']}'")
        print(f"  [*] Inverter dados: {self.config['invert_data']}\n")

        consecutive_errors = 0

        while self.running:
            try:
                # Rodar leitura do Excel em thread separada (COM precisa disso)
                new_trades, all_trades, status = await asyncio.to_thread(
                    self.reader.get_new_trades
                )

                if status and status != "reset":
                    consecutive_errors += 1
                    if consecutive_errors <= 3:
                        print(f"  [!] {status}")
                    elif consecutive_errors == 4:
                        print(f"  [!] Erros repetidos, silenciando...")
                else:
                    if consecutive_errors > 3:
                        print(f"  [OK] Conexao restabelecida")
                    consecutive_errors = 0

                if status == "reset":
                    print(f"  [~] Reset detectado - reenviando todos os dados")
                    await self.broadcast(json.dumps({
                        "type": "reset",
                        "trades": all_trades,
                        "timestamp": datetime.now().isoformat()
                    }))
                elif new_trades:
                    print(f"  [OK] +{len(new_trades)} novos negocios | "
                          f"Total: {len(all_trades)} | "
                          f"Saldo: {all_trades[-1]['saldo'] if all_trades else 0} | "
                          f"Hora: {new_trades[-1]['hora']}")

                    await self.broadcast(json.dumps({
                        "type": "update",
                        "trades": new_trades,
                        "total": len(all_trades),
                        "timestamp": datetime.now().isoformat()
                    }))

            except Exception as e:
                print(f"  [ERR] Erro no monitor: {e}")

            await asyncio.sleep(self.config["read_interval"])

    async def start(self):
        self.running = True
        http_port = self.config.get("http_port", 8080)

        print("\n" + "=" * 56)
        print("  FlowTrader Server v2.0 (xlwings)")
        print("=" * 56)
        print(f"\n  WebSocket:  ws://localhost:{self.config['ws_port']}")
        print(f"  Dashboard:  http://localhost:{http_port}/FlowTrader_Live.html")
        print(f"\n  Pressione Ctrl+C para parar\n")

        # Iniciar servidor HTTP em thread separada
        html_dir = os.path.dirname(os.path.abspath(__file__))
        handler = lambda *args, **kwargs: SimpleHTTPRequestHandler(*args, directory=html_dir, **kwargs)
        httpd = HTTPServer(("0.0.0.0", http_port), handler)
        http_thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        http_thread.start()
        print(f"  [OK] Servidor HTTP rodando na porta {http_port}")

        # Abrir dashboard no navegador automaticamente
        webbrowser.open(f"http://localhost:{http_port}/FlowTrader_Live.html")

        async with websockets.serve(self.handler, "0.0.0.0", self.config["ws_port"]):
            await self.monitor_excel()


# ============================================================
# GERADOR DO DASHBOARD HTML
# ============================================================

def generate_dashboard_html(ws_port=8765):
    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>iMoneyTrader - Fluxo de Agressão BMF (AO VIVO)</title>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg-base: #030712;
    --bg-gradient: radial-gradient(circle at top right, rgba(30, 58, 138, 0.15) 0%, transparent 40%),
                   radial-gradient(circle at bottom left, rgba(88, 28, 135, 0.1) 0%, transparent 40%);
    --surface: rgba(17, 24, 39, 0.6);
    --surface-hover: rgba(31, 41, 55, 0.8);
    --border: rgba(255, 255, 255, 0.08);
    --text-main: #f8fafc;
    --text-muted: #94a3b8;
    
    --green-neon: #10b981;
    --green-glow: rgba(16, 185, 129, 0.2);
    --green-bg: rgba(16, 185, 129, 0.08);
    
    --red-neon: #ef4444;
    --red-glow: rgba(239, 68, 68, 0.2);
    --red-bg: rgba(239, 68, 68, 0.08);
    
    --accent: #3b82f6;
    --accent-glow: rgba(59, 130, 246, 0.3);
    --yellow: #fbbf24;
    
    --shadow-glass: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
  }}
  
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  
  body {{ 
    background-color: var(--bg-base); 
    background-image: var(--bg-gradient);
    color: var(--text-main); 
    font-family: 'Outfit', sans-serif; 
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }}
  
  /* Glassmorphism Utilities */
  .glass-panel {{
    background: var(--surface);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid var(--border);
    border-radius: 16px;
    box-shadow: var(--shadow-glass);
  }}

  /* Typography */
  .mono {{ font-family: 'JetBrains Mono', monospace; }}
  
  /* Header */
  .header {{
    display: flex; justify-content: space-between; align-items: center;
    padding: 16px 24px;
    margin: 12px 12px 0 12px;
    border-radius: 16px;
    background: rgba(17, 24, 39, 0.7);
    backdrop-filter: blur(16px);
    border: 1px solid var(--border);
    z-index: 10;
  }}
  
  .logo-container {{ display: flex; align-items: center; gap: 12px; }}
  .logo {{ font-size: 24px; font-weight: 700; letter-spacing: -0.5px; }}
  .logo span {{ 
    background: linear-gradient(135deg, #60a5fa 0%, #3b82f6 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
  }}
  .badge-live {{
    background: rgba(251, 191, 36, 0.15);
    color: var(--yellow);
    padding: 4px 8px;
    border-radius: 6px;
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.5px;
    border: 1px solid rgba(251, 191, 36, 0.3);
  }}
  
  .header-controls {{ display: flex; align-items: center; gap: 24px; }}
  .asset-tag {{
    font-size: 16px; font-weight: 600; color: var(--text-main);
    background: rgba(255,255,255,0.05); padding: 4px 12px; border-radius: 8px;
    border: 1px solid var(--border);
  }}
  
  .connection-status {{ display: flex; align-items: center; gap: 8px; font-size: 13px; font-weight: 500; color: var(--text-muted); }}
  .pulse-ring {{
    position: relative; width: 10px; height: 10px;
  }}
  .pulse-dot {{
    position: absolute; top: 0; left: 0; width: 100%; height: 100%;
    border-radius: 50%; background: var(--red-neon);
    transition: background-color 0.3s;
  }}
  .pulse-echo {{
    position: absolute; top: -4px; left: -4px; width: 18px; height: 18px;
    border-radius: 50%; opacity: 0;
  }}
  .live .pulse-dot {{ background: var(--green-neon); box-shadow: 0 0 10px var(--green-glow); }}
  .live .pulse-echo {{ background: var(--green-neon); animation: sonar 2s infinite ease-out; }}
  
  @keyframes sonar {{
    0% {{ transform: scale(0.5); opacity: 0.8; }}
    100% {{ transform: scale(2.5); opacity: 0; }}
  }}

  /* Main Layout */
  .layout-grid {{
    display: grid;
    grid-template-columns: 1fr 380px;
    gap: 16px;
    padding: 12px;
    flex: 1;
    min-height: 0;
    transition: grid-template-columns 0.35s cubic-bezier(0.4, 0, 0.2, 1);
  }}
  .layout-grid.sidebar-collapsed {{
    grid-template-columns: 1fr 0px;
    gap: 0;
  }}
  .layout-grid.sidebar-collapsed .sidebar {{
    overflow: hidden;
    opacity: 0;
    pointer-events: none;
    transition: opacity 0.2s ease;
  }}
  .sidebar {{
    transition: opacity 0.25s ease 0.1s;
  }}

  /* Content Area (Left) */
  .content-area {{
    display: flex; flex-direction: column; gap: 16px; min-height: 0;
  }}

  /* KPI Cards */
  .kpi-grid {{
    display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px;
  }}
  .kpi-card {{
    padding: 16px; display: flex; flex-direction: column; gap: 4px;
    border-radius: 16px; background: rgba(17, 24, 39, 0.5);
    border: 1px solid var(--border);
    transition: transform 0.2s, background 0.2s;
  }}
  .kpi-card:hover {{ background: rgba(31, 41, 55, 0.6); transform: translateY(-2px); }}
  
  .kpi-label {{ font-size: 11px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; font-weight: 600; }}
  .kpi-value {{ font-size: 24px; font-weight: 700; line-height: 1.2; }}
  
  .val-pos {{ color: var(--green-neon); text-shadow: 0 0 12px var(--green-glow); }}
  .val-neg {{ color: var(--red-neon); text-shadow: 0 0 12px var(--red-glow); }}
  .val-neutral {{ color: var(--text-main); }}

  /* Main Chart Area */
  .chart-container {{
    flex: 1; position: relative; padding: 20px;
    min-height: 0;
  }}
  .chart-header {{
    position: absolute; top: 20px; left: 24px; right: 24px; z-index: 5;
    display: flex; gap: 16px; align-items: baseline;
  }}
  .chart-title {{ font-size: 18px; font-weight: 600; color: var(--text-main); }}
  .chart-subtitle {{ font-size: 13px; color: var(--text-muted); }}
  .chart-header-actions {{
    margin-left: auto; display: flex; gap: 8px; align-items: center;
  }}
  .btn-chart-action {{
    background: rgba(255,255,255,0.06); border: 1px solid var(--border);
    color: var(--text-muted); padding: 4px 10px; border-radius: 6px;
    cursor: pointer; font-size: 11px; font-weight: 600; letter-spacing: 0.3px;
    transition: all 0.2s; display: flex; align-items: center; gap: 4px;
    font-family: 'Outfit', sans-serif;
  }}
  .btn-chart-action:hover {{ background: rgba(255,255,255,0.12); color: var(--text-main); }}
  .btn-sidebar-toggle {{
    font-size: 14px; padding: 4px 8px; min-width: 28px; justify-content: center;
  }}
  /* Sidebar (Right) */
  .sidebar {{
    display: flex; flex-direction: column; gap: 16px; min-height: 0;
  }}

  /* Times & Trades (Tape) */
  .tape-container {{
    flex: 1; display: flex; flex-direction: column; overflow: hidden;
  }}
  .panel-header {{
    padding: 16px 20px; border-bottom: 1px solid var(--border);
    display: flex; justify-content: space-between; align-items: center;
  }}
  .panel-title {{ font-size: 14px; font-weight: 600; letter-spacing: 0.5px; color: var(--text-muted); text-transform: uppercase; }}
  .panel-badge {{ background: rgba(255,255,255,0.05); padding: 2px 8px; border-radius: 12px; font-size: 11px; }}

  .tape-list {{
    flex: 1; overflow-y: auto; padding: 8px 0;
  }}
  /* Custom Scrollbar */
  .tape-list::-webkit-scrollbar {{ width: 6px; }}
  .tape-list::-webkit-scrollbar-track {{ background: transparent; }}
  .tape-list::-webkit-scrollbar-thumb {{ background: rgba(255,255,255,0.1); border-radius: 10px; }}
  .tape-list::-webkit-scrollbar-thumb:hover {{ background: rgba(255,255,255,0.2); }}

  .tape-row {{
    display: grid; grid-template-columns: 60px 1fr 75px 50px 1fr; gap: 8px;
    padding: 8px 20px; font-size: 13px; align-items: center;
    border-left: 3px solid transparent;
    transition: background-color 0.2s;
    animation: slideIn 0.3s cubic-bezier(0.16, 1, 0.3, 1);
  }}
  
  @keyframes slideIn {{
    0% {{ opacity: 0; transform: translateY(-10px) scale(0.98); }}
    100% {{ opacity: 1; transform: translateY(0) scale(1); }}
  }}

  .tape-row.buy {{ 
    background: linear-gradient(90deg, var(--green-bg) 0%, transparent 100%);
    border-left-color: var(--green-neon);
  }}
  .tape-row.sell {{ 
    background: linear-gradient(90deg, var(--red-bg) 0%, transparent 100%);
    border-left-color: var(--red-neon);
  }}
  .tape-row:hover {{ background-color: var(--surface-hover); }}

  .t-time {{ color: var(--text-muted); font-size: 12px; }}
  .t-buyer {{ color: var(--green-neon); font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .t-seller {{ color: var(--red-neon); font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; text-align: right; }}
  .t-price {{ color: var(--text-main); font-weight: 600; text-align: right; }}
  .t-qty {{ color: var(--yellow); font-weight: 700; text-align: right; background: rgba(251, 191, 36, 0.1); padding: 2px 6px; border-radius: 4px; }}



  /* Settings Button */
  .settings-btn {{
    background: rgba(255,255,255,0.05); border: 1px solid var(--border);
    color: var(--text-muted); padding: 6px 8px; border-radius: 8px;
    cursor: pointer; transition: all 0.2s;
    display: flex; align-items: center;
  }}
  .settings-btn:hover {{ background: rgba(255,255,255,0.1); color: var(--text-main); }}

  /* Settings Modal */
  .settings-overlay {{
    position: fixed; inset: 0; background: rgba(0,0,0,0.6);
    backdrop-filter: blur(4px); z-index: 1000;
    display: flex; align-items: center; justify-content: center;
  }}
  .settings-panel {{
    width: 520px; max-height: 85vh; display: flex; flex-direction: column;
    background: rgba(17, 24, 39, 0.95) !important; border: 1px solid var(--border);
    border-radius: 16px; overflow: hidden;
  }}
  .settings-header {{
    display: flex; justify-content: space-between; align-items: center;
    padding: 20px 24px; border-bottom: 1px solid var(--border);
  }}
  .settings-header h2 {{ font-size: 18px; font-weight: 600; }}
  .settings-close {{
    background: none; border: none; color: var(--text-muted);
    font-size: 24px; cursor: pointer; padding: 0 4px;
  }}
  .settings-close:hover {{ color: var(--text-main); }}
  .settings-body {{
    flex: 1; overflow-y: auto; padding: 16px 24px;
    display: flex; flex-direction: column; gap: 20px;
  }}
  .settings-section {{ display: flex; flex-direction: column; gap: 12px; }}
  .settings-section-title {{
    font-size: 13px; font-weight: 600; color: var(--accent);
    text-transform: uppercase; letter-spacing: 0.5px;
    padding-bottom: 4px; border-bottom: 1px solid rgba(59,130,246,0.2);
  }}
  .settings-row {{
    display: flex; justify-content: space-between; align-items: center; gap: 12px;
  }}
  .settings-row label {{
    font-size: 13px; color: var(--text-muted); flex-shrink: 0;
  }}
  .settings-input {{
    background: rgba(0,0,0,0.3); border: 1px solid var(--border);
    color: var(--text-main); padding: 6px 10px; border-radius: 6px;
    font-size: 13px; font-family: 'JetBrains Mono', monospace;
    width: 180px; outline: none;
  }}
  .settings-input:focus {{ border-color: var(--accent); }}
  .settings-checkbox {{ width: 18px; height: 18px; accent-color: var(--accent); }}
  .settings-color {{
    width: 50px; height: 30px; border: 1px solid var(--border);
    border-radius: 4px; cursor: pointer; background: none; padding: 2px;
  }}
  .settings-footer {{
    padding: 16px 24px; border-top: 1px solid var(--border);
    display: flex; justify-content: flex-end; gap: 12px;
  }}
  .btn-save {{
    background: var(--accent); color: white; border: none;
    padding: 8px 20px; border-radius: 8px; font-weight: 600;
    cursor: pointer; font-size: 13px;
  }}
  .btn-save:hover {{ background: #2563eb; }}
  .btn-cancel {{
    background: rgba(255,255,255,0.05); color: var(--text-muted);
    border: 1px solid var(--border); padding: 8px 20px;
    border-radius: 8px; font-weight: 500; cursor: pointer; font-size: 13px;
  }}
  .btn-cancel:hover {{ background: rgba(255,255,255,0.1); }}

  /* Alarm Bell Icon */
  .alarm-btn {{
    background: rgba(255,255,255,0.05); border: 1px solid var(--border);
    color: var(--text-muted); padding: 6px 8px; border-radius: 8px;
    cursor: pointer; transition: all 0.2s;
    display: flex; align-items: center; position: relative;
  }}
  .alarm-btn:hover {{ background: rgba(255,255,255,0.1); color: var(--text-main); }}
  .alarm-btn.alarm-active {{
    color: var(--yellow); border-color: rgba(251, 191, 36, 0.4);
    animation: bellPulse 2s infinite;
  }}
  @keyframes bellPulse {{
    0%, 100% {{ opacity: 1; }}
    50% {{ opacity: 0.5; }}
  }}

  /* Toast Notifications */
  .toast-container {{
    position: fixed; top: 80px; right: 20px; z-index: 2000;
    display: flex; flex-direction: column; gap: 8px;
    pointer-events: none;
  }}
  .alarm-toast {{
    pointer-events: auto;
    background: rgba(17, 24, 39, 0.95);
    backdrop-filter: blur(16px);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 12px 16px;
    display: flex; align-items: center; gap: 12px;
    min-width: 280px; max-width: 380px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.4);
    animation: toastSlideIn 0.4s cubic-bezier(0.16, 1, 0.3, 1);
    transition: opacity 0.3s, transform 0.3s;
  }}
  .alarm-toast.removing {{
    opacity: 0; transform: translateX(40px);
  }}
  @keyframes toastSlideIn {{
    0% {{ opacity: 0; transform: translateX(60px); }}
    100% {{ opacity: 1; transform: translateX(0); }}
  }}
  .alarm-toast.toast-saldo {{ border-left: 3px solid var(--yellow); }}
  .alarm-toast.toast-valor {{ border-left: 3px solid var(--accent); }}
  .alarm-toast.toast-qty {{ border-left: 3px solid #a78bfa; }}
  .toast-icon {{
    font-size: 18px; flex-shrink: 0; width: 28px; height: 28px;
    display: flex; align-items: center; justify-content: center;
    border-radius: 8px;
  }}
  .toast-saldo .toast-icon {{ background: rgba(251, 191, 36, 0.15); color: var(--yellow); }}
  .toast-valor .toast-icon {{ background: rgba(59, 130, 246, 0.15); color: var(--accent); }}
  .toast-qty .toast-icon {{ background: rgba(167, 139, 250, 0.15); color: #a78bfa; }}
  .toast-body {{ flex: 1; }}
  .toast-title {{ font-size: 12px; font-weight: 600; color: var(--text-main); margin-bottom: 2px; }}
  .toast-msg {{ font-size: 11px; color: var(--text-muted); }}
  .toast-close {{
    background: none; border: none; color: var(--text-muted); cursor: pointer;
    font-size: 16px; padding: 0 2px; flex-shrink: 0; line-height: 1;
  }}
  .toast-close:hover {{ color: var(--text-main); }}

  /* Alarm Panel (popover) */
  .alarm-panel {{
    position: fixed; top: 60px; right: 120px; z-index: 1001;
    width: 340px;
    background: rgba(17, 24, 39, 0.97);
    backdrop-filter: blur(20px);
    border: 1px solid rgba(251, 191, 36, 0.15);
    border-radius: 16px;
    box-shadow: 0 12px 40px rgba(0,0,0,0.5), 0 0 0 1px rgba(255,255,255,0.04);
    animation: alarmPanelIn 0.25s cubic-bezier(0.16, 1, 0.3, 1);
    overflow: hidden;
  }}
  @keyframes alarmPanelIn {{
    0% {{ opacity: 0; transform: translateY(-8px) scale(0.97); }}
    100% {{ opacity: 1; transform: translateY(0) scale(1); }}
  }}
  .alarm-panel-header {{
    display: flex; justify-content: space-between; align-items: center;
    padding: 14px 18px; border-bottom: 1px solid var(--border);
  }}
  .alarm-panel-header h3 {{
    font-size: 14px; font-weight: 600; color: var(--yellow);
    display: flex; align-items: center; gap: 8px;
  }}
  .alarm-panel-close {{
    background: none; border: none; color: var(--text-muted);
    font-size: 20px; cursor: pointer; padding: 0 2px; line-height: 1;
  }}
  .alarm-panel-close:hover {{ color: var(--text-main); }}
  .alarm-panel-body {{
    padding: 14px 18px; display: flex; flex-direction: column; gap: 14px;
  }}

  /* Alarm Card — each alarm type */
  .alarm-card {{
    background: rgba(255,255,255,0.03);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 12px 14px;
    transition: border-color 0.2s, background 0.2s;
  }}
  .alarm-card.alarm-card-on {{
    border-color: rgba(251, 191, 36, 0.3);
    background: rgba(251, 191, 36, 0.04);
  }}
  .alarm-card-top {{
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 8px;
  }}
  .alarm-card-label {{
    font-size: 12px; font-weight: 600; color: var(--text-main);
    display: flex; align-items: center; gap: 6px;
  }}
  .alarm-card-label .alarm-dot {{
    width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0;
  }}
  .alarm-dot-saldo {{ background: var(--yellow); }}
  .alarm-dot-valor {{ background: var(--accent); }}
  .alarm-dot-qty {{ background: #a78bfa; }}
  .alarm-card-fields {{
    display: flex; align-items: center; gap: 8px;
  }}
  .alarm-card-fields select,
  .alarm-card-fields input[type="number"] {{
    background: rgba(0,0,0,0.3); border: 1px solid var(--border);
    color: var(--text-main); padding: 5px 8px; border-radius: 6px;
    font-size: 12px; font-family: 'JetBrains Mono', monospace;
    outline: none; flex: 1; min-width: 0;
  }}
  .alarm-card-fields select {{ max-width: 90px; }}
  .alarm-card-fields input[type="number"] {{ max-width: 100px; }}
  .alarm-card-fields select:focus,
  .alarm-card-fields input:focus {{ border-color: var(--yellow); }}

  /* Toggle switch */
  .alarm-toggle {{
    position: relative; width: 36px; height: 20px; flex-shrink: 0;
  }}
  .alarm-toggle input {{ opacity: 0; width: 0; height: 0; }}
  .alarm-toggle .slider {{
    position: absolute; inset: 0; cursor: pointer;
    background: rgba(255,255,255,0.1); border-radius: 10px;
    transition: background 0.2s;
  }}
  .alarm-toggle .slider::before {{
    content: ''; position: absolute; left: 2px; top: 2px;
    width: 16px; height: 16px; border-radius: 50%;
    background: var(--text-muted); transition: transform 0.2s, background 0.2s;
  }}
  .alarm-toggle input:checked + .slider {{
    background: rgba(251, 191, 36, 0.35);
  }}
  .alarm-toggle input:checked + .slider::before {{
    transform: translateX(16px); background: var(--yellow);
  }}

  /* Sound row */
  .alarm-sound-row {{
    display: flex; align-items: center; justify-content: space-between;
    padding-top: 6px; border-top: 1px solid var(--border);
  }}
  .alarm-sound-left {{
    display: flex; align-items: center; gap: 8px;
    font-size: 12px; color: var(--text-muted);
  }}
  .btn-test-sound {{
    background: rgba(251, 191, 36, 0.1); color: var(--yellow);
    border: 1px solid rgba(251, 191, 36, 0.3); padding: 4px 10px;
    border-radius: 6px; cursor: pointer; font-size: 11px; font-weight: 500;
  }}
  .btn-test-sound:hover {{ background: rgba(251, 191, 36, 0.2); }}

  .alarm-panel-backdrop {{
    position: fixed; inset: 0; z-index: 1000; background: transparent;
  }}


</style>
</head>
<body>

<div class="header">
  <div class="logo-container">
    <div class="logo">iMoney<span>Trader</span></div>
    <div class="badge-live">PRO</div>
    <span style="font-size: 10px; color: var(--text-muted); font-weight: 400;">Desenvolvido por Mexflow.IA</span>
  </div>
  
  <div class="header-controls">
    <div class="asset-tag mono" id="asset">WINJ26</div>
    <button id="btnAlarm" class="alarm-btn" onclick="toggleAlarmPanel()" title="Alarmes">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path>
        <path d="M13.73 21a2 2 0 0 1-3.46 0"></path>
      </svg>
    </button>
    <button id="btnSettings" class="settings-btn" onclick="toggleSettings()" title="Configurações">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="3"></circle>
        <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"></path>
      </svg>
    </button>
    <div class="connection-status">
      <div class="pulse-ring" id="pulseRing">
        <div class="pulse-echo"></div>
        <div class="pulse-dot"></div>
      </div>
      <span id="statusText">Aguardando Conexão...</span>
    </div>
  </div>
</div>

<!-- Settings Modal -->
<div id="settingsOverlay" class="settings-overlay" style="display:none;" onclick="if(event.target===this) toggleSettings()">
  <div class="settings-panel glass-panel">
    <div class="settings-header">
      <h2>Configurações</h2>
      <button class="settings-close" onclick="toggleSettings()">&times;</button>
    </div>
    <div class="settings-body">

      <div class="settings-section">
        <div class="settings-section-title">Dados Excel</div>
        <div class="settings-row">
          <label>Nome do Workbook</label>
          <div style="display:flex; gap:6px; align-items:center;">
            <input type="text" id="cfg_book_name" class="settings-input" />
            <button id="btnDetectWb" onclick="detectWorkbooks()" style="background:rgba(59,130,246,0.15); color:var(--accent); border:1px solid rgba(59,130,246,0.3); padding:6px 10px; border-radius:6px; font-size:12px; font-weight:600; cursor:pointer; white-space:nowrap; transition:all 0.2s;">Detectar</button>
          </div>
        </div>
        <div id="wbListContainer" style="display:none; margin-top:2px; margin-bottom:6px;">
          <div id="wbList" style="background:rgba(0,0,0,0.4); border:1px solid var(--border); border-radius:8px; max-height:160px; overflow-y:auto;"></div>
        </div>
        <div class="settings-row">
          <label>Nome da Aba</label>
          <input type="text" id="cfg_sheet_name" class="settings-input" />
        </div>
        <div class="settings-row">
          <label>Linha Inicial dos Dados</label>
          <input type="number" id="cfg_data_start_row" class="settings-input" min="1" />
        </div>
        <div class="settings-row">
          <label>Inverter Dados (DDE)</label>
          <input type="checkbox" id="cfg_invert_data" class="settings-checkbox" />
        </div>
      </div>

      <div class="settings-section">
        <div class="settings-section-title">Filtros</div>
        <div class="settings-row">
          <label>Quantidade Mínima</label>
          <input type="number" id="cfg_min_quantity" class="settings-input" min="1" />
        </div>
      </div>

      <div class="settings-section">
        <div class="settings-section-title">Média</div>
        <div class="settings-row">
          <label>Tipo de Média</label>
          <select id="cfg_avg_type" class="settings-input" onchange="onAvgTypeChange()">
            <option value="cumulative">Cumulativa</option>
            <option value="sma">SMA - Média Móvel Simples</option>
            <option value="ema">EMA - Média Móvel Exponencial</option>
          </select>
        </div>
        <div class="settings-row" id="row_sma_period" style="display:none;">
          <label>Períodos SMA</label>
          <input type="number" id="cfg_sma_period" class="settings-input" min="2" max="500" value="20" />
        </div>
        <div class="settings-row" id="row_ema_alpha" style="display:none;">
          <label>Alpha EMA (0.01 - 1.0)</label>
          <input type="number" id="cfg_ema_alpha" class="settings-input" min="0.01" max="1" step="0.01" value="0.1" />
        </div>
      </div>

      <div class="settings-section">
        <div class="settings-section-title">Visual</div>
        <div class="settings-row">
          <label>Cor de Compra</label>
          <input type="color" id="cfg_color_buy" class="settings-color" value="#10b981" />
        </div>
        <div class="settings-row">
          <label>Cor de Venda</label>
          <input type="color" id="cfg_color_sell" class="settings-color" value="#ef4444" />
        </div>
        <div class="settings-row">
          <label>Nome do Ativo</label>
          <input type="text" id="cfg_asset_name" class="settings-input" value="WINJ26" />
        </div>
      </div>

      <div class="settings-section">
        <div class="settings-section-title">Conexão</div>
        <div class="settings-row">
          <label>Intervalo de Leitura (s)</label>
          <input type="number" id="cfg_read_interval" class="settings-input" min="1" max="60" />
        </div>
      </div>

      <div class="settings-section">
        <div class="settings-section-title">Dados</div>
        <div class="settings-row" style="justify-content: center;">
          <button onclick="clearDatabase()" style="background: rgba(239,68,68,0.15); color: #ef4444; border: 1px solid rgba(239,68,68,0.3); padding: 8px 20px; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; transition: all 0.2s;">
            Apagar Dados do Banco
          </button>
        </div>
      </div>

    </div>
    <div class="settings-footer">
      <button class="btn-cancel" onclick="toggleSettings()">Cancelar</button>
      <button class="btn-save" onclick="saveSettings()">Salvar</button>
    </div>
  </div>
</div>

<div class="toast-container" id="toastContainer"></div>

<!-- Alarm Panel (popover) -->
<div id="alarmBackdrop" class="alarm-panel-backdrop" style="display:none;" onclick="toggleAlarmPanel()"></div>
<div id="alarmPanel" class="alarm-panel" style="display:none;">
  <div class="alarm-panel-header">
    <h3>
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg>
      Alarmes
    </h3>
    <button class="alarm-panel-close" onclick="toggleAlarmPanel()">&times;</button>
  </div>
  <div class="alarm-panel-body">

    <!-- Saldo -->
    <div class="alarm-card" id="alarmCardSaldo">
      <div class="alarm-card-top">
        <span class="alarm-card-label"><span class="alarm-dot alarm-dot-saldo"></span>Saldo de Agressão</span>
        <label class="alarm-toggle">
          <input type="checkbox" id="alm_saldo_on" onchange="onAlarmToggle()" />
          <span class="slider"></span>
        </label>
      </div>
      <div class="alarm-card-fields">
        <select id="alm_saldo_dir" onchange="saveAlarmSettings()">
          <option value="acima">≥ Acima</option>
          <option value="abaixo">≤ Abaixo</option>
          <option value="ambos">|x| Ambos</option>
        </select>
        <input type="number" id="alm_saldo_val" placeholder="Ex: 100" onchange="saveAlarmSettings()" />
      </div>
    </div>

    <!-- Valor -->
    <div class="alarm-card" id="alarmCardValor">
      <div class="alarm-card-top">
        <span class="alarm-card-label"><span class="alarm-dot alarm-dot-valor"></span>Preço (Valor)</span>
        <label class="alarm-toggle">
          <input type="checkbox" id="alm_valor_on" onchange="onAlarmToggle()" />
          <span class="slider"></span>
        </label>
      </div>
      <div class="alarm-card-fields">
        <select id="alm_valor_dir" onchange="saveAlarmSettings()">
          <option value="acima">≥ Acima</option>
          <option value="abaixo">≤ Abaixo</option>
        </select>
        <input type="number" id="alm_valor_val" placeholder="Ex: 130000" onchange="saveAlarmSettings()" />
      </div>
    </div>

    <!-- Quantidade -->
    <div class="alarm-card" id="alarmCardQty">
      <div class="alarm-card-top">
        <span class="alarm-card-label"><span class="alarm-dot alarm-dot-qty"></span>Quantidade</span>
        <label class="alarm-toggle">
          <input type="checkbox" id="alm_qty_on" onchange="onAlarmToggle()" />
          <span class="slider"></span>
        </label>
      </div>
      <div class="alarm-card-fields">
        <input type="number" id="alm_qty_val" placeholder="Ex: 100" onchange="saveAlarmSettings()" style="max-width:none;" />
      </div>
    </div>

    <!-- Som -->
    <div class="alarm-sound-row">
      <div class="alarm-sound-left">
        <label class="alarm-toggle">
          <input type="checkbox" id="alm_sound" onchange="saveAlarmSettings()" />
          <span class="slider"></span>
        </label>
        <span>Som</span>
      </div>
      <button type="button" class="btn-test-sound" onclick="testAlarmSound()">Testar</button>
    </div>

  </div>
</div>

<div class="layout-grid">
  <!-- Left Column: KPIs & Chart -->
  <div class="content-area">
    
    <div class="kpi-grid">
      <div class="kpi-card glass-panel">
        <span class="kpi-label">Saldo de Agressão</span>
        <span class="kpi-value mono" id="sSaldo">0</span>
      </div>
      <div class="kpi-card glass-panel">
        <span class="kpi-label">Agressões Compra</span>
        <span class="kpi-value val-pos mono" id="sBuys">0</span>
      </div>
      <div class="kpi-card glass-panel">
        <span class="kpi-label">Agressões Venda</span>
        <span class="kpi-value val-neg mono" id="sSells">0</span>
      </div>
      <div class="kpi-card glass-panel">
        <span class="kpi-label">Último Preço</span>
        <span class="kpi-value mono" id="sPrice">--</span>
      </div>
      <div class="kpi-card glass-panel">
        <span class="kpi-label">Total Negócios</span>
        <span class="kpi-value mono val-neutral" id="sTotal">0</span>
      </div>
      <div class="kpi-card glass-panel">
        <span class="kpi-label">Última Att.</span>
        <span class="kpi-value mono val-neutral" id="sTime" style="font-size: 20px; line-height: 1.4;">--</span>
      </div>
    </div>

    <div class="chart-container glass-panel">
      <div class="chart-header">
        <div class="chart-title">Acumulado do Fluxo</div>
        <div class="chart-subtitle" id="chartSubtitle">Evolução do Saldo de Agressão</div>
        <div class="chart-header-actions">
          <button class="btn-chart-action" onclick="if(chart) chart.timeScale().fitContent()" title="Ajustar gráfico para mostrar todos os dados">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 3 21 3 21 9"></polyline><polyline points="9 21 3 21 3 15"></polyline><line x1="21" y1="3" x2="14" y2="10"></line><line x1="3" y1="21" x2="10" y2="14"></line></svg>
            Fit All
          </button>
          <button id="btnToggleSidebar" class="btn-chart-action btn-sidebar-toggle" onclick="toggleSidebar()" title="Ocultar Painel">
            »
          </button>
        </div>
      </div>
      <!-- Aqui usaremos a div tvchart em vez do canvas manual -->
      <div id="tvchart" style="position:absolute; top:80px; left:16px; right:16px; bottom:16px; border-radius:12px; overflow:hidden;"></div>
    </div>

  </div>

  <!-- Right Column: Tape & Brokers -->
  <div class="sidebar">
    
    <div class="tape-container glass-panel">
      <div class="panel-header">
        <span class="panel-title">Times & Trades</span>
        <span class="panel-badge" id="tapeCount">0 rows</span>
      </div>
      <div class="tape-list" id="tape">
        <!-- Rendered by JS -->
      </div>
    </div>



  </div>
</div>

<script src="https://unpkg.com/lightweight-charts@3.8.0/dist/lightweight-charts.standalone.production.js"></script>

<script>
const WS_URL = "ws://localhost:{ws_port}";

// --- UI Configuration Options ---
const CHART_LINE_COLOR_POS = "#10b981";
const CHART_LINE_COLOR_NEG = "#ef4444";
const CHART_BG_GRADIENT_POS_START = "rgba(16, 185, 129, 0.25)";
const CHART_BG_GRADIENT_NEG_START = "rgba(239, 68, 68, 0.25)";
const CHART_GRID_COLOR = "rgba(255, 255, 255, 0.06)";
const CHART_TEXT_COLOR = "#94a3b8";

let allTrades = [];
let buys = 0, sells = 0;

let ws = null, reconnectTimer = null;

// --- Chart Variables ---
let chart = null;
let areaSeries = null;
let avgSeries = null;
let historicalDataForTV = [];
let historicalAvgForTV = [];
let saldoValues = [];
let saldoSum = 0;
let lastTimestamp = 0;

// --- Settings ---
const CLIENT_DEFAULTS = {{
    avg_type: "cumulative",
    sma_period: 20,
    ema_alpha: 0.1,
    color_buy: "#10b981",
    color_sell: "#ef4444",
    asset_name: "WINJ26",
    alarm_saldo_enabled: false,
    alarm_saldo_value: 100,
    alarm_saldo_dir: "acima",
    alarm_valor_enabled: false,
    alarm_valor_value: 0,
    alarm_valor_dir: "acima",
    alarm_qty_enabled: false,
    alarm_qty_value: 100,
    alarm_sound: true
}};
let clientSettings = {{ ...CLIENT_DEFAULTS }};
let serverConfig = {{}};

function formatPrice(val) {{
  if(!val) return "--";
  return (val / 1000).toFixed(3).replace('.', ',');
}}

function parseTimeToUnixTimestamp(horaStr) {{
    // Formato vem "09:12:05" do Excel
    const parts = (horaStr || '').split(':');
    if (parts.length < 2 || isNaN(parseInt(parts[0])) || isNaN(parseInt(parts[1]))) {{
        // Hora invalida: usar lastTimestamp + 1 para nao criar salto no grafico
        if (lastTimestamp > 0) {{
            lastTimestamp = lastTimestamp + 1;
            return lastTimestamp;
        }}
        // Fallback absoluto: agora em UTC como business time
        const now = new Date();
        lastTimestamp = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(),
            now.getHours(), now.getMinutes(), now.getSeconds()) / 1000;
        return lastTimestamp;
    }}

    const h = parseInt(parts[0]);
    const m = parseInt(parts[1]);
    const s = parts.length > 2 ? parseInt(parts[2]) : 0;

    // Usar UTC para construir o timestamp — lightweight-charts v3 interpreta
    // timestamps Unix como UTC, e o tickMarkFormatter ja converte para local.
    // Isso garante que a POSICAO no eixo X bata com o LABEL exibido.
    const now = new Date();
    let ts = Math.floor(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), h, m, s) / 1000);

    // Tratamento para timestamps duplicados — incremento minimo para
    // nao criar lacunas artificiais visiveis no eixo horizontal.
    if (ts <= lastTimestamp) {{
        ts = lastTimestamp + 1;
    }}
    lastTimestamp = ts;
    return ts;
}}

function initChart() {{
    const container = document.getElementById('tvchart');
    chart = LightweightCharts.createChart(container, {{
        layout: {{
            background: {{ type: 'solid', color: 'transparent' }},
            textColor: CHART_TEXT_COLOR,
            fontFamily: "'JetBrains Mono', monospace",
        }},
        grid: {{
            vertLines: {{ color: CHART_GRID_COLOR, style: 1 }},
            horzLines: {{ color: CHART_GRID_COLOR, style: 1 }},
        }},
        crosshair: {{
            mode: LightweightCharts.CrosshairMode.Normal,
            vertLine: {{ width: 1, color: "rgba(255,255,255,0.4)", style: 3 }},
            horzLine: {{ width: 1, color: "rgba(255,255,255,0.4)", style: 3 }},
        }},
        rightPriceScale: {{
            borderColor: "rgba(255,255,255,0.1)",
        }},
        timeScale: {{
            borderColor: "rgba(255,255,255,0.1)",
            timeVisible: true,
            secondsVisible: true,
            minBarSpacing: 0.001,
            tickMarkFormatter: (time) => {{
                // Timestamps sao armazenados como UTC "business time" — exibir como UTC
                // para que 09:00:00 no dado apareca como 09:00:00 no eixo
                const date = new Date(time * 1000);
                const hh = String(date.getUTCHours()).padStart(2, '0');
                const mm = String(date.getUTCMinutes()).padStart(2, '0');
                const ss = String(date.getUTCSeconds()).padStart(2, '0');
                return hh + ':' + mm + ':' + ss;
            }},
        }},
    }});

    areaSeries = chart.addBaselineSeries({{
        baseValue: {{ type: 'price', price: 0 }},
        topLineColor: CHART_LINE_COLOR_POS,
        topFillColor1: CHART_BG_GRADIENT_POS_START,
        topFillColor2: 'rgba(16, 185, 129, 0)',
        bottomLineColor: CHART_LINE_COLOR_NEG,
        bottomFillColor1: 'rgba(239, 68, 68, 0)',
        bottomFillColor2: CHART_BG_GRADIENT_NEG_START,
        lineWidth: 3,
        priceFormat: {{
            type: 'price',
            precision: 0,
            minMove: 1,
        }},
    }});

    avgSeries = chart.addLineSeries({{
        color: '#3b82f6',
        lineWidth: 2,
        lineStyle: 2,
        crosshairMarkerVisible: false,
        lastValueVisible: true,
        priceLineVisible: false,
        priceFormat: {{
            type: 'price',
            precision: 1,
            minMove: 0.1,
        }},
    }});

    // ResizeObserver handles ALL resize scenarios: window resize, sidebar toggle, etc.
    const ro = new ResizeObserver(() => {{
        if (chart && container.clientWidth > 0 && container.clientHeight > 0) {{
            chart.resize(container.clientWidth, container.clientHeight);
        }}
    }});
    ro.observe(container);
}}

function connect() {{
  ws = new WebSocket(WS_URL);
  
  ws.onopen = () => {{
    document.getElementById("pulseRing").classList.add("live");
    document.getElementById("statusText").textContent = "Conectado - Ao Vivo";
    document.getElementById("statusText").style.color = "var(--green-neon)";
    if (reconnectTimer) {{ clearInterval(reconnectTimer); reconnectTimer = null; }}
  }};
  
  ws.onmessage = (e) => {{
    let msg;
    try {{
      msg = JSON.parse(e.data);
    }} catch (err) {{
      console.error("JSON parse error:", err);
      return;
    }}

    if (msg.type === "history" || msg.type === "reset") {{
      allTrades = []; buys = 0; sells = 0;
      document.getElementById("tape").innerHTML = "";

      historicalDataForTV = [];
      historicalAvgForTV = [];
      saldoValues = [];
      saldoSum = 0;
      lastTimestamp = 0;

      // Limpar series do chart imediatamente para evitar flash de dados antigos
      if (areaSeries) areaSeries.setData([]);
      if (avgSeries) avgSeries.setData([]);

      // Ordenar por hora para garantir sequencia cronologica no grafico
      const sortedTrades = msg.trades.slice().sort((a, b) => {{
          const ta = a.hora || ''; const tb = b.hora || '';
          return ta.localeCompare(tb);
      }});

      sortedTrades.forEach((t) => {{
          processTrade(t, false);
          const ts = parseTimeToUnixTimestamp(t.hora);
          const saldoVal = parseFloat(t.saldo) || 0;
          historicalDataForTV.push({{ time: ts, value: saldoVal }});
          saldoValues.push(saldoVal);
      }});

      recalcAllAverages();

      const last50 = sortedTrades.slice(-50);
      last50.forEach(t => addToTapeUI(t));

      if(sortedTrades.length > 0) {{
          updateStatsUI(sortedTrades[sortedTrades.length - 1]);
      }}

      if (areaSeries) areaSeries.setData(historicalDataForTV);
      if (chart) chart.timeScale().fitContent();


    }}

    if (msg.type === "update") {{
      msg.trades.forEach(t => {{
        processTrade(t, true);
        addToTapeUI(t);
        const ts = parseTimeToUnixTimestamp(t.hora);
        const saldoVal = parseFloat(t.saldo);
        historicalDataForTV.push({{ time: ts, value: saldoVal }});
        if (areaSeries) {{
            areaSeries.update({{ time: ts, value: saldoVal }});
        }}
        calcAndPushAvgForNewPoint(saldoVal, ts);
      }});

    }}

    if (msg.type === "server_config") {{
      serverConfig = msg.config;
    }}

    if (msg.type === "config_updated") {{
      serverConfig = msg.config;
    }}

    if (msg.type === "workbook_list") {{
      renderWorkbookList(msg.workbooks);
    }}
  }};
  
  ws.onclose = () => {{
    document.getElementById("pulseRing").classList.remove("live");
    document.getElementById("statusText").textContent = "Reconectando...";
    document.getElementById("statusText").style.color = "var(--red-neon)";
    if (!reconnectTimer) reconnectTimer = setInterval(connect, 3000);
  }};
  
  ws.onerror = () => ws.close();
}}

function processTrade(t, updateUI) {{
  allTrades.push(t);

  const isBuy = t.agressor === "Comprador";
  if (isBuy) {{
    buys++;
  }} else {{
    sells++;
  }}

  if (updateUI) {{
    updateStatsUI(t);
    checkAlarms(t);
  }}
}}

function addToTapeUI(t) {{
  const tape = document.getElementById("tape");
  const isBuy = t.agressor === "Comprador";
  
  const row = document.createElement("div");
  row.className = "tape-row " + (isBuy ? "buy" : "sell");
  
  row.innerHTML = `
    <span class="t-time mono">${{t.hora.substring(0,8)}}</span>
    <span class="t-buyer">${{t.compradora}}</span>
    <span class="t-price mono">${{formatPrice(t.valor)}}</span>
    <span class="t-qty mono">${{t.quantidade}}</span>
    <span class="t-seller">${{t.vendedora}}</span>
  `;
  
  // Insert at top
  tape.insertBefore(row, tape.firstChild);
  
  // Keep max elements to avoid DOM lag
  if (tape.children.length > 100) tape.removeChild(tape.lastChild);
  
  document.getElementById("tapeCount").textContent = `${{Math.min(tape.children.length, 100)}} rows`;
}}

function updateStatsUI(t) {{
  const elSaldo = document.getElementById("sSaldo");
  elSaldo.textContent = (t.saldo > 0 ? "+" : "") + t.saldo;
  elSaldo.className = "kpi-value mono " + (t.saldo > 0 ? "val-pos" : (t.saldo < 0 ? "val-neg" : "val-neutral"));
  
  document.getElementById("sBuys").textContent = buys;
  document.getElementById("sSells").textContent = sells;
  document.getElementById("sPrice").textContent = formatPrice(t.valor);
  document.getElementById("sTotal").textContent = allTrades.length;
  document.getElementById("sTime").textContent = t.hora;
}}



// ========== AVERAGE CALCULATION ENGINE ==========

function getAvgTypeLabel() {{
    const type = clientSettings.avg_type;
    if (type === "cumulative") return "Média Cumulativa";
    if (type === "sma") return "SMA(" + clientSettings.sma_period + ")";
    if (type === "ema") return "EMA(α=" + clientSettings.ema_alpha + ")";
    return "Média";
}}

function recalcAllAverages() {{
    if (saldoValues.length === 0) return;

    historicalAvgForTV = [];
    saldoSum = 0;
    const type = clientSettings.avg_type;

    if (type === "cumulative") {{
        let sum = 0;
        for (let i = 0; i < saldoValues.length; i++) {{
            sum += saldoValues[i];
            historicalAvgForTV.push({{
                time: historicalDataForTV[i].time,
                value: parseFloat((sum / (i + 1)).toFixed(1))
            }});
        }}
        saldoSum = sum;
    }} else if (type === "sma") {{
        const n = clientSettings.sma_period;
        let windowSum = 0;
        for (let i = 0; i < saldoValues.length; i++) {{
            windowSum += saldoValues[i];
            if (i >= n) windowSum -= saldoValues[i - n];
            const count = Math.min(i + 1, n);
            historicalAvgForTV.push({{
                time: historicalDataForTV[i].time,
                value: parseFloat((windowSum / count).toFixed(1))
            }});
        }}
        // manter saldoSum para fallback
        for (let i = 0; i < saldoValues.length; i++) saldoSum += saldoValues[i];
    }} else if (type === "ema") {{
        const alpha = clientSettings.ema_alpha;
        let ema = saldoValues[0];
        historicalAvgForTV.push({{
            time: historicalDataForTV[0].time,
            value: parseFloat(ema.toFixed(1))
        }});
        for (let i = 1; i < saldoValues.length; i++) {{
            ema = alpha * saldoValues[i] + (1 - alpha) * ema;
            historicalAvgForTV.push({{
                time: historicalDataForTV[i].time,
                value: parseFloat(ema.toFixed(1))
            }});
        }}
        for (let i = 0; i < saldoValues.length; i++) saldoSum += saldoValues[i];
    }}

    if (avgSeries) avgSeries.setData(historicalAvgForTV);

    if (historicalAvgForTV.length > 0) {{
        const lastAvg = historicalAvgForTV[historicalAvgForTV.length - 1].value;
        document.getElementById("chartSubtitle").textContent =
            "Evolução do Saldo de Agressão | " + getAvgTypeLabel() + ": " + lastAvg;
    }}
}}

function calcAndPushAvgForNewPoint(saldoVal, timestamp) {{
    saldoValues.push(saldoVal);
    const idx = saldoValues.length - 1;
    const type = clientSettings.avg_type;
    let avg;

    if (type === "cumulative") {{
        saldoSum += saldoVal;
        avg = parseFloat((saldoSum / saldoValues.length).toFixed(1));
    }} else if (type === "sma") {{
        const n = clientSettings.sma_period;
        const start = Math.max(0, idx - n + 1);
        let sum = 0;
        for (let i = start; i <= idx; i++) sum += saldoValues[i];
        avg = parseFloat((sum / (idx - start + 1)).toFixed(1));
    }} else if (type === "ema") {{
        const prevEma = historicalAvgForTV.length > 0
            ? historicalAvgForTV[historicalAvgForTV.length - 1].value
            : saldoVal;
        avg = parseFloat((clientSettings.ema_alpha * saldoVal +
               (1 - clientSettings.ema_alpha) * prevEma).toFixed(1));
    }}

    const avgPoint = {{ time: timestamp, value: avg }};
    historicalAvgForTV.push(avgPoint);
    if (avgSeries) avgSeries.update(avgPoint);

    document.getElementById("chartSubtitle").textContent =
        "Evolução do Saldo de Agressão | " + getAvgTypeLabel() + ": " + avg;
}}

// ========== ALARM SYSTEM ==========

let alarmCooldowns = {{ saldo: 0, valor: 0, qty: 0 }};
const ALARM_COOLDOWN_MS = 10000;
let audioCtx = null;

function getAudioContext() {{
    if (!audioCtx) {{
        audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    }}
    return audioCtx;
}}

function playAlarmBeep(type) {{
    if (!clientSettings.alarm_sound) return;
    try {{
        const ctx = getAudioContext();
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain);
        gain.connect(ctx.destination);

        if (type === "saldo") {{
            osc.frequency.value = 880;
            osc.type = "sine";
        }} else if (type === "valor") {{
            osc.frequency.value = 660;
            osc.type = "triangle";
        }} else {{
            osc.frequency.value = 1100;
            osc.type = "square";
        }}

        gain.gain.setValueAtTime(0.15, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.5);
        osc.start(ctx.currentTime);
        osc.stop(ctx.currentTime + 0.5);

        // Second beep for emphasis
        const osc2 = ctx.createOscillator();
        const gain2 = ctx.createGain();
        osc2.connect(gain2);
        gain2.connect(ctx.destination);
        osc2.frequency.value = osc.frequency.value;
        osc2.type = osc.type;
        gain2.gain.setValueAtTime(0.12, ctx.currentTime + 0.15);
        gain2.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.6);
        osc2.start(ctx.currentTime + 0.15);
        osc2.stop(ctx.currentTime + 0.6);
    }} catch(e) {{ console.warn("Audio alarm failed:", e); }}
}}

function testAlarmSound() {{
    playAlarmBeep("saldo");
    setTimeout(() => playAlarmBeep("valor"), 700);
    setTimeout(() => playAlarmBeep("qty"), 1400);
}}

function showAlarmToast(type, title, message) {{
    const container = document.getElementById("toastContainer");
    const toast = document.createElement("div");
    toast.className = "alarm-toast toast-" + type;

    const icons = {{ saldo: "⚖", valor: "💲", qty: "📦" }};
    toast.innerHTML = `
        <div class="toast-icon">${{icons[type] || "🔔"}}</div>
        <div class="toast-body">
            <div class="toast-title">${{title}}</div>
            <div class="toast-msg">${{message}}</div>
        </div>
        <button class="toast-close" onclick="this.parentElement.remove()">×</button>
    `;

    container.appendChild(toast);
    playAlarmBeep(type);

    // Flash bell icon
    const bell = document.getElementById("btnAlarm");
    bell.classList.add("alarm-active");
    setTimeout(() => bell.classList.remove("alarm-active"), 5000);

    // Auto-remove after 5 seconds
    setTimeout(() => {{
        toast.classList.add("removing");
        setTimeout(() => toast.remove(), 300);
    }}, 5000);
}}

function checkAlarms(trade) {{
    const now = Date.now();

    // Saldo alarm
    if (clientSettings.alarm_saldo_enabled) {{
        const saldo = parseFloat(trade.saldo) || 0;
        const limit = clientSettings.alarm_saldo_value;
        const dir = clientSettings.alarm_saldo_dir;
        let triggered = false;

        if (dir === "acima" && saldo >= limit) triggered = true;
        else if (dir === "abaixo" && saldo <= -Math.abs(limit)) triggered = true;
        else if (dir === "ambos" && Math.abs(saldo) >= Math.abs(limit)) triggered = true;

        if (triggered && now - alarmCooldowns.saldo > ALARM_COOLDOWN_MS) {{
            alarmCooldowns.saldo = now;
            const dirLabel = dir === "acima" ? "acima de" : dir === "abaixo" ? "abaixo de" : "atingiu";
            showAlarmToast("saldo", "Alarme de Saldo",
                "Saldo " + dirLabel + " " + limit + " → Atual: " + saldo);
        }}
    }}

    // Valor (price) alarm
    if (clientSettings.alarm_valor_enabled) {{
        const valor = parseFloat(trade.valor) || 0;
        const limit = clientSettings.alarm_valor_value;
        const dir = clientSettings.alarm_valor_dir;
        let triggered = false;

        if (dir === "acima" && valor >= limit) triggered = true;
        else if (dir === "abaixo" && valor <= limit) triggered = true;

        if (triggered && now - alarmCooldowns.valor > ALARM_COOLDOWN_MS) {{
            alarmCooldowns.valor = now;
            showAlarmToast("valor", "Alarme de Preço",
                "Preço " + (dir === "acima" ? "acima de" : "abaixo de") + " " + formatPrice(limit) + " → Atual: " + formatPrice(valor));
        }}
    }}

    // Quantity alarm
    if (clientSettings.alarm_qty_enabled) {{
        const qty = parseFloat(trade.quantidade) || 0;
        const limit = clientSettings.alarm_qty_value;

        if (qty >= limit && now - alarmCooldowns.qty > ALARM_COOLDOWN_MS) {{
            alarmCooldowns.qty = now;
            showAlarmToast("qty", "Alarme de Quantidade",
                "Trade com " + qty + " contratos (limite: " + limit + ")");
        }}
    }}
}}

// ========== SETTINGS MANAGEMENT ==========

function hexToRgb(hex) {{
    const result = /^#?([a-f\\d]{{2}})([a-f\\d]{{2}})([a-f\\d]{{2}})$/i.exec(hex);
    return result ? {{
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    }} : null;
}}

function loadClientSettings() {{
    try {{
        const saved = localStorage.getItem("flowtrader_client_settings");
        if (saved) {{
            const parsed = JSON.parse(saved);
            clientSettings = {{ ...CLIENT_DEFAULTS, ...parsed }};
        }}
    }} catch(e) {{}}
    applyClientSettings();
}}

function saveClientSettings() {{
    localStorage.setItem("flowtrader_client_settings", JSON.stringify(clientSettings));
}}

function applyClientSettings() {{
    // Asset name
    document.getElementById("asset").textContent = clientSettings.asset_name;

    // CSS variables for colors
    document.documentElement.style.setProperty('--green-neon', clientSettings.color_buy);
    const buyRGB = hexToRgb(clientSettings.color_buy);
    if (buyRGB) {{
        document.documentElement.style.setProperty('--green-glow',
            'rgba(' + buyRGB.r + ',' + buyRGB.g + ',' + buyRGB.b + ',0.2)');
        document.documentElement.style.setProperty('--green-bg',
            'rgba(' + buyRGB.r + ',' + buyRGB.g + ',' + buyRGB.b + ',0.08)');
    }}
    document.documentElement.style.setProperty('--red-neon', clientSettings.color_sell);
    const sellRGB = hexToRgb(clientSettings.color_sell);
    if (sellRGB) {{
        document.documentElement.style.setProperty('--red-glow',
            'rgba(' + sellRGB.r + ',' + sellRGB.g + ',' + sellRGB.b + ',0.2)');
        document.documentElement.style.setProperty('--red-bg',
            'rgba(' + sellRGB.r + ',' + sellRGB.g + ',' + sellRGB.b + ',0.08)');
    }}

    // Update chart series colors
    if (areaSeries) {{
        const bRgb = buyRGB || {{r:16,g:185,b:129}};
        const sRgb = sellRGB || {{r:239,g:68,b:68}};
        areaSeries.applyOptions({{
            topLineColor: clientSettings.color_buy,
            topFillColor1: 'rgba(' + bRgb.r + ',' + bRgb.g + ',' + bRgb.b + ',0.25)',
            topFillColor2: 'rgba(' + bRgb.r + ',' + bRgb.g + ',' + bRgb.b + ',0)',
            bottomLineColor: clientSettings.color_sell,
            bottomFillColor1: 'rgba(' + sRgb.r + ',' + sRgb.g + ',' + sRgb.b + ',0)',
            bottomFillColor2: 'rgba(' + sRgb.r + ',' + sRgb.g + ',' + sRgb.b + ',0.25)',
        }});
    }}

    recalcAllAverages();

    updateBellState();
}}

function detectWorkbooks() {{
    const btn = document.getElementById("btnDetectWb");
    btn.textContent = "Buscando...";
    btn.disabled = true;
    if (ws && ws.readyState === WebSocket.OPEN) {{
        ws.send(JSON.stringify({{ type: "list_workbooks" }}));
    }} else {{
        btn.textContent = "Detectar";
        btn.disabled = false;
        alert("WebSocket desconectado.");
    }}
}}

function renderWorkbookList(workbooks) {{
    const btn = document.getElementById("btnDetectWb");
    btn.textContent = "Detectar";
    btn.disabled = false;
    const container = document.getElementById("wbListContainer");
    const list = document.getElementById("wbList");
    list.innerHTML = "";
    if (!workbooks || workbooks.length === 0) {{
        list.innerHTML = '<div style="padding:10px; color:var(--text-muted); font-size:12px; text-align:center;">Nenhum workbook aberto encontrado.</div>';
        container.style.display = "block";
        return;
    }}
    workbooks.forEach(wb => {{
        const item = document.createElement("div");
        item.style.cssText = "padding:8px 12px; cursor:pointer; border-bottom:1px solid rgba(255,255,255,0.05); transition:background 0.15s;";
        item.onmouseenter = () => item.style.background = "rgba(59,130,246,0.1)";
        item.onmouseleave = () => item.style.background = "none";
        const name = document.createElement("div");
        name.style.cssText = "font-size:13px; color:var(--text-main); font-weight:600;";
        name.textContent = wb.name;
        const sheets = document.createElement("div");
        sheets.style.cssText = "font-size:11px; color:var(--text-muted); margin-top:2px;";
        sheets.textContent = "Abas: " + wb.sheets.join(", ");
        item.appendChild(name);
        item.appendChild(sheets);
        item.onclick = () => {{
            document.getElementById("cfg_book_name").value = wb.name;
            if (wb.sheets.length > 0) {{
                document.getElementById("cfg_sheet_name").value = wb.sheets[0];
            }}
            container.style.display = "none";
        }};
        list.appendChild(item);
    }});
    container.style.display = "block";
}}

function clearDatabase() {{
    if (!confirm("Tem certeza que deseja apagar todos os dados do banco?")) return;
    if (ws && ws.readyState === WebSocket.OPEN) {{
        ws.send(JSON.stringify({{ type: "clear_database" }}));
        toggleSettings();
        showAlarmToast("saldo", "Banco de Dados", "Todos os dados foram apagados.");
    }}
}}

function toggleSettings() {{
    const overlay = document.getElementById("settingsOverlay");
    const visible = overlay.style.display !== "none";
    if (visible) {{
        overlay.style.display = "none";
    }} else {{
        populateSettingsForm();
        overlay.style.display = "flex";
    }}
}}

function populateSettingsForm() {{
    // Server settings
    document.getElementById("cfg_book_name").value = serverConfig.book_name || "";
    document.getElementById("cfg_sheet_name").value = serverConfig.sheet_name || "";
    document.getElementById("cfg_data_start_row").value = serverConfig.data_start_row || 2;
    document.getElementById("cfg_invert_data").checked = serverConfig.invert_data || false;
    document.getElementById("cfg_min_quantity").value = serverConfig.min_quantity || 1;
    document.getElementById("cfg_read_interval").value = serverConfig.read_interval || 3;

    // Client settings
    document.getElementById("cfg_avg_type").value = clientSettings.avg_type;
    document.getElementById("cfg_sma_period").value = clientSettings.sma_period;
    document.getElementById("cfg_ema_alpha").value = clientSettings.ema_alpha;
    document.getElementById("cfg_color_buy").value = clientSettings.color_buy;
    document.getElementById("cfg_color_sell").value = clientSettings.color_sell;
    document.getElementById("cfg_asset_name").value = clientSettings.asset_name;

    onAvgTypeChange();
}}

function onAvgTypeChange() {{
    const type = document.getElementById("cfg_avg_type").value;
    document.getElementById("row_sma_period").style.display = type === "sma" ? "flex" : "none";
    document.getElementById("row_ema_alpha").style.display = type === "ema" ? "flex" : "none";
}}

function saveSettings() {{
    // Validation
    const readInterval = parseInt(document.getElementById("cfg_read_interval").value);
    const minQty = parseInt(document.getElementById("cfg_min_quantity").value);
    const smaPeriod = parseInt(document.getElementById("cfg_sma_period").value);
    const emaAlpha = parseFloat(document.getElementById("cfg_ema_alpha").value);

    if (readInterval < 1 || readInterval > 60) {{ alert("Intervalo de leitura: 1 a 60 segundos"); return; }}
    if (minQty < 1) {{ alert("Quantidade mínima deve ser >= 1"); return; }}
    if (smaPeriod < 2 || smaPeriod > 500) {{ alert("Períodos SMA: 2 a 500"); return; }}
    if (emaAlpha < 0.01 || emaAlpha > 1) {{ alert("Alpha EMA: 0.01 a 1.0"); return; }}

    // 1. Server-side config via WebSocket
    const newServerCfg = {{
        book_name: document.getElementById("cfg_book_name").value,
        sheet_name: document.getElementById("cfg_sheet_name").value,
        data_start_row: parseInt(document.getElementById("cfg_data_start_row").value),
        invert_data: document.getElementById("cfg_invert_data").checked,
        min_quantity: minQty,
        read_interval: readInterval,
    }};

    if (ws && ws.readyState === WebSocket.OPEN) {{
        ws.send(JSON.stringify({{
            type: "update_server_config",
            config: newServerCfg
        }}));
    }}

    // 2. Client-side settings to localStorage
    clientSettings.avg_type = document.getElementById("cfg_avg_type").value;
    clientSettings.sma_period = smaPeriod;
    clientSettings.ema_alpha = emaAlpha;
    clientSettings.color_buy = document.getElementById("cfg_color_buy").value;
    clientSettings.color_sell = document.getElementById("cfg_color_sell").value;
    clientSettings.asset_name = document.getElementById("cfg_asset_name").value;

    saveClientSettings();
    applyClientSettings();
    toggleSettings();
}}

// ========== ALARM PANEL MANAGEMENT ==========

function toggleAlarmPanel() {{
    const panel = document.getElementById("alarmPanel");
    const backdrop = document.getElementById("alarmBackdrop");
    const visible = panel.style.display !== "none";
    if (visible) {{
        panel.style.display = "none";
        backdrop.style.display = "none";
    }} else {{
        populateAlarmForm();
        panel.style.display = "block";
        backdrop.style.display = "block";
    }}
}}

function populateAlarmForm() {{
    document.getElementById("alm_saldo_on").checked = clientSettings.alarm_saldo_enabled;
    document.getElementById("alm_saldo_val").value = clientSettings.alarm_saldo_value;
    document.getElementById("alm_saldo_dir").value = clientSettings.alarm_saldo_dir;
    document.getElementById("alm_valor_on").checked = clientSettings.alarm_valor_enabled;
    document.getElementById("alm_valor_val").value = clientSettings.alarm_valor_value;
    document.getElementById("alm_valor_dir").value = clientSettings.alarm_valor_dir;
    document.getElementById("alm_qty_on").checked = clientSettings.alarm_qty_enabled;
    document.getElementById("alm_qty_val").value = clientSettings.alarm_qty_value;
    document.getElementById("alm_sound").checked = clientSettings.alarm_sound;
    updateAlarmCardStates();
}}

function saveAlarmSettings() {{
    clientSettings.alarm_saldo_enabled = document.getElementById("alm_saldo_on").checked;
    clientSettings.alarm_saldo_value = parseFloat(document.getElementById("alm_saldo_val").value) || 0;
    clientSettings.alarm_saldo_dir = document.getElementById("alm_saldo_dir").value;
    clientSettings.alarm_valor_enabled = document.getElementById("alm_valor_on").checked;
    clientSettings.alarm_valor_value = parseFloat(document.getElementById("alm_valor_val").value) || 0;
    clientSettings.alarm_valor_dir = document.getElementById("alm_valor_dir").value;
    clientSettings.alarm_qty_enabled = document.getElementById("alm_qty_on").checked;
    clientSettings.alarm_qty_value = parseFloat(document.getElementById("alm_qty_val").value) || 0;
    clientSettings.alarm_sound = document.getElementById("alm_sound").checked;
    saveClientSettings();
    updateBellState();
    updateAlarmCardStates();
}}

function onAlarmToggle() {{
    saveAlarmSettings();
}}

function updateAlarmCardStates() {{
    document.getElementById("alarmCardSaldo").classList.toggle("alarm-card-on", clientSettings.alarm_saldo_enabled);
    document.getElementById("alarmCardValor").classList.toggle("alarm-card-on", clientSettings.alarm_valor_enabled);
    document.getElementById("alarmCardQty").classList.toggle("alarm-card-on", clientSettings.alarm_qty_enabled);
}}

function updateBellState() {{
    const anyOn = clientSettings.alarm_saldo_enabled || clientSettings.alarm_valor_enabled || clientSettings.alarm_qty_enabled;
    const btn = document.getElementById("btnAlarm");
    if (btn) {{
        btn.style.borderColor = anyOn ? "rgba(251, 191, 36, 0.4)" : "";
        btn.style.color = anyOn ? "var(--yellow)" : "";
    }}
}}

// ESC to close settings or alarm panel
document.addEventListener('keydown', (e) => {{
    if (e.key === 'Escape') {{
        const overlay = document.getElementById("settingsOverlay");
        if (overlay.style.display !== "none") {{
            toggleSettings();
            return;
        }}
        const alarmPanel = document.getElementById("alarmPanel");
        if (alarmPanel.style.display !== "none") {{
            toggleAlarmPanel();
        }}
    }}
}});

// ========== SIDEBAR TOGGLE ==========
let sidebarCollapsed = false;

function toggleSidebar() {{
    sidebarCollapsed = !sidebarCollapsed;
    const grid = document.querySelector('.layout-grid');
    const btn = document.getElementById('btnToggleSidebar');
    grid.classList.toggle('sidebar-collapsed', sidebarCollapsed);
    btn.innerHTML = sidebarCollapsed ? '\u00ab' : '\u00bb';
    btn.title = sidebarCollapsed ? 'Mostrar Painel' : 'Ocultar Painel';
    // Persist state
    clientSettings.sidebar_collapsed = sidebarCollapsed;
    saveClientSettings();
    // ResizeObserver handles chart resize automatically
}}

function restoreSidebarState() {{
    if (clientSettings.sidebar_collapsed) {{
        sidebarCollapsed = true;
        const grid = document.querySelector('.layout-grid');
        const btn = document.getElementById('btnToggleSidebar');
        grid.classList.add('sidebar-collapsed');
        btn.innerHTML = '\u00ab';
        btn.title = 'Mostrar Painel';
    }}
}}

// ========== INITIALIZATION ==========
setTimeout(() => {{
    loadClientSettings();
    restoreSidebarState();
    try {{
        initChart();
    }} catch(e) {{
        console.error("Erro ao inicializar grafico:", e);
    }}
    // Conectar somente apos chart estar pronto, usando requestAnimationFrame
    // para garantir que o DOM do chart ja foi renderizado
    requestAnimationFrame(() => {{
        connect();
    }});
}}, 300);

</script>
</body>
</html>"""
    return html


# ============================================================
# MAIN
# ============================================================

def main():
    import argparse

    parser = argparse.ArgumentParser(description="FlowTrader v2 - Dashboard em Tempo Real (xlwings)")
    parser.add_argument("--book", type=str, help="Nome do workbook no Excel (parcial)")
    parser.add_argument("--sheet", type=str, help="Nome da aba")
    parser.add_argument("--excel", type=str, help="Caminho do arquivo Excel (fallback)")
    parser.add_argument("--port", type=int, default=None, help="Porta WebSocket")
    parser.add_argument("--interval", type=int, default=None, help="Intervalo de leitura (s)")
    parser.add_argument("--invert", action="store_true", help="Inverter dados (DDE)")
    parser.add_argument("--start-row", type=int, help="Linha inicial dos dados")
    parser.add_argument("--generate-html", action="store_true", help="Apenas gerar o HTML")
    args = parser.parse_args()

    # Carregar config.json (prioridade: defaults < config.json < CLI args)
    load_config()

    if args.book:
        CONFIG["book_name"] = args.book
    if args.sheet:
        CONFIG["sheet_name"] = args.sheet
    if args.excel:
        CONFIG["excel_path"] = args.excel
    if args.port:
        CONFIG["ws_port"] = args.port
    if args.interval:
        CONFIG["read_interval"] = args.interval
    if args.invert:
        CONFIG["invert_data"] = True
    if args.start_row:
        CONFIG["data_start_row"] = args.start_row

    # Gerar dashboard HTML
    html = generate_dashboard_html(CONFIG["ws_port"])
    html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "FlowTrader_Live.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n  [OK] Dashboard gerado: {html_path}")

    if args.generate_html:
        return

    # Iniciar servidor
    server = FlowTraderServer(CONFIG)
    try:
        asyncio.run(server.start())
    except KeyboardInterrupt:
        print("\n\n  [STOP] Servidor parado.\n")


if __name__ == "__main__":
    main()

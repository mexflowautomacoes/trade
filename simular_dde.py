"""
Simulador DDE v2 - Simula o ProfitChart alimentando o Excel via xlwings
========================================================================
Abre o Excel visivel, cria uma planilha e vai adicionando linhas de dados
reais do Delmer, simulando o comportamento do DDE em tempo real.

O servidor FlowTrader conecta neste mesmo Excel aberto para ler os dados.

USO:
    python simular_dde.py
"""

import xlwings as xw
import openpyxl
import time
import random
from datetime import datetime, time as tm

# ============================================================
# CONFIGURACOES
# ============================================================

# Planilha original com dados reais do Delmer
ORIGINAL_PATH = r"C:\Users\Victor\Downloads\trader\v1\Planilha de dados BM&F.xlsx"
ORIGINAL_SHEET = "Plan1"

# Nome da planilha que sera criada no Excel (simulando a aba DDE)
SIMULATED_BOOK_NAME = "FlowTrader_DDE.xlsx"
SIMULATED_SHEET_NAME = "Plan1"

# Intervalo entre cada leva de dados (segundos)
INTERVAL = 1.0

# Linhas por leva
MIN_BATCH = 1
MAX_BATCH = 5

# Limite de linhas na planilha (simula o comportamento real do DDE)
MAX_ROWS = 500


# ============================================================
# CARREGAR DADOS REAIS (openpyxl - so para ler o arquivo fonte)
# ============================================================

def load_real_data():
    """Carrega todos os trades reais da planilha do Delmer."""
    print(f"  Carregando dados reais de: {ORIGINAL_PATH}")
    wb = openpyxl.load_workbook(ORIGINAL_PATH, data_only=True, read_only=True)
    ws = wb[ORIGINAL_SHEET]

    header = None
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            header = list(row[:6])  # So as 6 primeiras colunas
            continue
        vals = list(row[:6])
        if vals[0] is not None:
            rows.append(vals)

    wb.close()
    print(f"  {len(rows)} trades carregados")
    return header, rows


# ============================================================
# MAIN
# ============================================================

def main():
    print("\n" + "=" * 56)
    print("  Simulador DDE v2 (xlwings)")
    print("  Simula ProfitChart -> Excel")
    print("=" * 56 + "\n")

    # Carregar dados reais
    header, all_rows = load_real_data()

    # Abrir Excel visivel (como o Delmer faz)
    print("  Abrindo Excel...")
    app = xw.App(visible=True)
    wb = app.books.add()
    wb.save(SIMULATED_BOOK_NAME)

    ws = wb.sheets[0]
    ws.name = SIMULATED_SHEET_NAME

    # Escrever cabecalho
    ws.range("A1").value = header
    # Formatar cabecalho
    ws.range("A1:F1").font.bold = True

    print(f"  Excel aberto: {wb.name}")
    print(f"  Aba: {ws.name}")
    print(f"\n  Intervalo: {INTERVAL}s entre cada leva")
    print(f"  Batch: {MIN_BATCH}-{MAX_BATCH} linhas por vez")
    print(f"  Limite DDE: {MAX_ROWS} linhas (exclui as mais antigas)")
    print(f"  Total a enviar: {len(all_rows)} trades")
    print(f"\n  Pressione Ctrl+C para parar\n")

    cursor = 0
    total = len(all_rows)

    try:
        while cursor < total:
            # Pegar proximo batch
            batch_size = random.randint(MIN_BATCH, MAX_BATCH)
            batch = all_rows[cursor:cursor + batch_size]
            actual_size = len(batch)

            # Simular DDE real: inserir linhas NOVAS NO TOPO (linha 2)
            # Igual ao ProfitChart faz - mais recente fica em cima
            # Inverter o batch para que o mais recente fique na linha 2
            batch_reversed = list(reversed(batch))

            # Inserir linhas em branco no topo para empurrar os dados antigos
            for _ in range(actual_size):
                ws.range("2:2").api.Insert()

            # Escrever os novos dados no topo (linhas 2 em diante)
            ws.range("A2").value = batch_reversed

            # Simular limite do DDE: manter no maximo MAX_ROWS linhas de dados
            # Dados comecam na linha 2 (linha 1 = cabecalho), entao ultima linha permitida = MAX_ROWS + 1
            last_allowed_row = MAX_ROWS + 1  # +1 pelo cabecalho
            # Checar se ha dados alem do limite
            check_cell = ws.range(f"A{last_allowed_row + 1}").value
            if check_cell is not None:
                # Descobrir ultima linha real com dados
                last_data = ws.range(f"A{last_allowed_row + 1}").end("down").row
                if last_data > last_allowed_row and last_data < last_allowed_row + 100000:
                    rows_to_delete = last_data - last_allowed_row
                    ws.range(f"{last_allowed_row + 1}:{last_data}").api.Delete()
                    truncated = True
                else:
                    # So uma linha extra
                    ws.range(f"{last_allowed_row + 1}:{last_allowed_row + 1}").api.Delete()
                    rows_to_delete = 1
                    truncated = True
            else:
                truncated = False

            cursor += actual_size

            # Progresso
            pct = (cursor / total) * 100
            last_hora = batch[-1][0]
            if hasattr(last_hora, 'strftime'):
                hora_str = last_hora.strftime("%H:%M:%S")
            else:
                hora_str = str(last_hora)

            # Contar linhas atuais na planilha
            current_rows = cursor if cursor <= MAX_ROWS else MAX_ROWS
            trunc_msg = f" | TRUNCADO (limite {MAX_ROWS})" if truncated else ""

            print(f"  [DDE] +{actual_size} negocios | "
                  f"Enviados: {cursor}/{total} ({pct:.1f}%) | "
                  f"Na planilha: {current_rows} | "
                  f"Hora: {hora_str}{trunc_msg}")

            time.sleep(INTERVAL)

        print(f"\n  [FIM] Todos os {total} trades foram enviados!")
        print(f"  O Excel permanece aberto para o servidor ler.")
        print(f"  Pressione Ctrl+C para fechar.")

        # Manter o Excel aberto
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print(f"\n\n  [PARADO] {cursor}/{total} trades enviados")

    finally:
        try:
            wb.close()
            app.quit()
            print("  Excel fechado.")
        except:
            pass


if __name__ == "__main__":
    main()

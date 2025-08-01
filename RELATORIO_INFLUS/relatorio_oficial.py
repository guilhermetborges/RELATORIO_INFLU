import requests
import os
import pandas as pd
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar
from datetime import datetime,timezone, timedelta
import openpyxl
import threading


# Carrega variáveis de ambiente
load_dotenv()
TOKEN = os.getenv("NUVEMSHOP_TOKEN")
USER_ID = os.getenv("NUVEMSHOP_USER_ID")


BASE_URL = f"https://api.tiendanube.com/v1/{USER_ID}/orders"
HEADERS = {
    "Authentication": f"bearer {TOKEN}",
    "User-Agent": "API BI (guilhermeborges@pangeia96.com)",
    "Content-Type": "application/json"
}

# Sessão com retry
session = requests.Session()
retries = Retry(total=5, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
adapter = HTTPAdapter(max_retries=retries)
session.mount('https://', adapter)




CUPONS_PERMITIDOS = {
    "LISBELLA", "DUDAMILLER", "TOZETTO", "LARISSA", "ANGEL", "MANUCUNHA",
    "CAMSTM", "TRIZ", "DUDADALLA", "MDM", "BELLASARDI", "GIRLBLOGGER",
    "SOPHIALUZ", "LARAF", "LARAB", "NANDA", "ZWISOCA", "LAURA", "MAVINAKA",
    "GABISOARES", "CLARISSA", "LARILODI", "GIOAGUILERA", "MARIRODRIGUES",
    "HELOYOHANA", "LARABALIEIRO", "LALAMOUNIER", "GIOLIVEIRA", "JULIASOUTO",
    "THAISCHAGAS", "NICMARCONDES", "SOPHIAROCHA", "JETISOVEC", "ISAHARAGAO",
    "RENATASIMAS", "DUDASALERNO", "ALEZAMBELLI", "MADUG", "GABISQ", "ISAMARTE",
    "MARISALES", "MAYA", "LALOURENCA", "GIMARTINISI", "NATTY"
}



def relatorio_cupons(data_inicio, data_fim, status_label, botao_gerar):
    botao_gerar.config(state='disabled')
    status_label.config(text="⏳ Gerando relatório...")

    dados = []

    try:
        # Converte datas de string para datetime e ajusta para UTC
        inicio_local = datetime.strptime(data_inicio, "%Y-%m-%d")
        fim_local = datetime.strptime(data_fim, "%Y-%m-%d")

        # Define UTC−3 (fuso de Brasília)
        fuso_brasilia = timezone(timedelta(hours=-3))

        # Adiciona horários fixos (00:00:00 e 23:59:59) e fuso de Brasília
        inicio_brasilia = datetime.combine(inicio_local, datetime.min.time(), tzinfo=fuso_brasilia)
        fim_brasilia = datetime.combine(fim_local, datetime.max.time().replace(microsecond=0), tzinfo=fuso_brasilia)

        # Converte para UTC
        inicio_utc = inicio_brasilia.astimezone(timezone.utc)
        fim_utc = fim_brasilia.astimezone(timezone.utc)

        for tipo in ["open", "closed"]:
            page = 0
            while True:
                params = {
                    "per_page": 200,
                    "page": page,
                    "created_at_min": inicio_utc.strftime('%Y-%m-%dT%H:%M:%SZ'),
                    "created_at_max": fim_utc.strftime('%Y-%m-%dT%H:%M:%SZ'),
                    "payment_status": "paid",
                    "status": tipo
                }
                print(params)

                response = session.get(BASE_URL, headers=HEADERS, params=params, timeout=30)
                if response.status_code != 200:
                    status_label.config(text=f"❌ Erro {response.status_code}: {response.text}")
                    break

                pedidos = response.json()
                if not pedidos:
                    break

                for pedido in pedidos:
                    created_at = pedido.get('created_at')
                    created_at_dt = datetime.strptime(created_at, "%Y-%m-%dT%H:%M:%S%z")
                    if not (inicio_utc <= created_at_dt <= fim_utc):
                        continue

                    cupons = pedido.get('coupon', [])
                    valor_pedido = float(pedido.get('subtotal')) - float(pedido.get('discount'))

                    if cupons:
                        c = cupons[0]
                        codigo = c.get('code')
                        if codigo and (codigo.endswith("10") or codigo in CUPONS_PERMITIDOS):
                            dados.append({
                                'codigo_cupom': codigo,
                                'valor': valor_pedido,
                                'id': pedido.get('number') or pedido.get('id')  # segurança
                            })

                if len(pedidos) < 200:
                    break
                page += 1

        # Fora do loop de tipo
        if not dados:
            status_label.config(text="⚠️ Nenhum cupom encontrado no período.")
        else:
            df = pd.DataFrame(dados)
            df = df.drop_duplicates(subset='id')
            df.drop(columns=['id'], inplace=True)

            agrupado = df.groupby('codigo_cupom').agg(
                valor_total=('valor', 'sum'),
                vezes_usado=('valor', 'count')
            ).reset_index().sort_values(by='valor_total', ascending=False)

            # Nome único para evitar erro de permissão
            nome_arquivo = f"cupons_dia_{data_inicio}_{datetime.now().strftime('%H%M%S')}.xlsx"
            agrupado.to_excel(nome_arquivo, index=False, engine='openpyxl')

            status_label.config(text=f"✅ Arquivo '{nome_arquivo}' salvo com sucesso.")

    except Exception as e:
        status_label.config(text=f"❌ Erro: {e}")

    botao_gerar.config(state='normal')


def abrir_calendario(destino_var, titulo="Selecionar Data"):
    def selecionar_data():
        data = cal.selection_get()
        destino_var.set(data.strftime('%Y-%m-%d'))
        top.destroy()

    top = tk.Toplevel()
    top.title(titulo)
    cal = Calendar(top, selectmode='day', date_pattern='yyyy-mm-dd')
    cal.pack(pady=10)
    tk.Button(top, text="Selecionar", command=selecionar_data).pack(pady=5)


def abrir_interface():
    root = tk.Tk()
    root.title("Relatório de Cupons - Pangeia96")
    root.geometry("400x300")

    data_inicio_var = tk.StringVar()
    data_fim_var = tk.StringVar()

    tk.Label(root, text="Data de Início do Relatório:").pack(pady=5)
    frame_inicio = tk.Frame(root)
    frame_inicio.pack()
    tk.Entry(frame_inicio, textvariable=data_inicio_var, width=15).pack(side="left", padx=5)
    tk.Button(frame_inicio, text="Selecionar Data", command=lambda: abrir_calendario(data_inicio_var, "Data de Início")).pack(side="left")

    tk.Label(root, text="Data do Fim:").pack(pady=5)
    frame_fim = tk.Frame(root)
    frame_fim.pack()
    tk.Entry(frame_fim, textvariable=data_fim_var, width=15).pack(side="left", padx=5)
    tk.Button(frame_fim, text="Selecionar Data", command=lambda: abrir_calendario(data_fim_var, "Data Final")).pack(side="left")

    status_label = tk.Label(root, text="", fg="blue")
    status_label.pack(pady=15)

    botao_gerar = tk.Button(root, text="Gerar Excel", bg="#4CAF50", fg="white", padx=10, pady=5)
    botao_gerar.pack()

    def ao_clicar():
        inicio = data_inicio_var.get()
        fim = data_fim_var.get()

        if not inicio or not fim:
            status_label.config(text="⚠️ Selecione ambas as datas.")
            return

        threading.Thread(target=relatorio_cupons, args=(inicio, fim, status_label, botao_gerar), daemon=True).start()

    botao_gerar.config(command=ao_clicar)

    root.mainloop()


abrir_interface()

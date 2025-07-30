import requests
import os
from dotenv import load_dotenv
import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar
from datetime import datetime
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


def relatorio_cupons(data_inicio, data_fim, status_label, botao_gerar):
    botao_gerar.config(state='disabled')
    status_label.config(text="⏳ Gerando relatório...")

    page = 0
    dados = []

    try:
        while True:
            params = {
                "per_page": 200,
                "page": page,
                "created_at_min": f"{data_inicio}T00:00:00-03:00",
                "created_at_max": f"{data_fim}T23:59:59-03:00",
                "payment_status": "paid"
            }

            response = session.get(BASE_URL, headers=HEADERS, params=params, timeout=30)
            if response.status_code != 200:
                status_label.config(text=f"❌ Erro {response.status_code}: {response.text}")
                break

            pedidos = response.json()
            if not pedidos:
                break

            for pedido in pedidos:
                cupons = pedido.get('coupon', [])
                valor_pedido = float(pedido.get('subtotal')) - float(pedido.get('discount'))

                if cupons:
                    c = cupons[0]
                    dados.append({
                        'codigo_cupom': c.get('code'),
                        'valor': valor_pedido
                    })

            if len(pedidos) < 200:
                break
            page += 1

        if not dados:
            status_label.config(text="⚠️ Nenhum cupom encontrado no período.")
        else:
            df = pd.DataFrame(dados)
            agrupado = df.groupby('codigo_cupom').agg(
                valor_total=('valor', 'sum'),
                vezes_usado=('valor', 'count')
            ).reset_index().sort_values(by='valor_total', ascending=False)

            nome_arquivo = f"cupons_dia_{data_inicio}.xlsx"
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

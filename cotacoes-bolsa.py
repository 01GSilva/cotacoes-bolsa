import yfinance as yf
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import os
import tkinter as tk
from tkinter import ttk

arquivo_excel = "cotacoes_bolsa.xlsx"

# ---------------- JANELA AÇOES NACIONAIS E FIIS-------------------


def solicitar_quantidades(lista):
    quantities = {}

    def confirmar():
        for ticker, entry in entries.items():
            try:
                q = float(entry.get())
            except:
                q = 0
            quantities[ticker] = q
        root.destroy()

    root = tk.Tk()
    root.title("Quantidades de Ativos")

    tk.Label(root, text="Digite as quantidades de cada ativo:", font=(
        "Arial", 12, "bold")).grid(row=0, column=0, columnspan=2, pady=10)

    entries = {}
    row_index = 1

    for ticker in lista:
        tk.Label(root, text=ticker, font=("Arial", 10)).grid(
            row=row_index, column=0, padx=10, pady=5, sticky="w")

        entry = ttk.Entry(root, width=10)
        entry.grid(row=row_index, column=1, padx=10)
        entry.insert(0, "0")

        entries[ticker] = entry
        row_index += 1

    btn = ttk.Button(root, text="Confirmar", command=confirmar)
    btn.grid(row=row_index, column=0, columnspan=2, pady=15)

    root.mainloop()

    return quantities


# ---------------- JANELA AÇOES INTERNACIONAIS -------------------

def solicitar_valores_dolar(lista):
    valores = {}

    def confirmar():
        for ticker, entry in entries.items():
            try:
                v = float(entry.get())
            except:
                v = 0
            valores[ticker] = v
        root.destroy()

    root = tk.Tk()
    root.title("Valores em Dólar")

    tk.Label(root, text="Digite o valor em dólar de cada ativo internacional:",
             font=("Arial", 12, "bold")).grid(row=0, column=0, columnspan=2, pady=10)

    entries = {}
    row_index = 1

    for ticker in lista:
        tk.Label(root, text=ticker, font=("Arial", 10)).grid(
            row=row_index, column=0, padx=10, pady=5, sticky="w")

        entry = ttk.Entry(root, width=10)
        entry.grid(row=row_index, column=1, padx=10)
        entry.insert(0, "0")

        entries[ticker] = entry
        row_index += 1

    btn = ttk.Button(root, text="Confirmar", command=confirmar)
    btn.grid(row=row_index, column=0, columnspan=2, pady=15)

    root.mainloop()

    return valores

# ---------------- JANELA RENDA FIXA -------------------


def solicitar_renda_fixa():
    valor = {"renda_fixa": 0}

    def confirmar():
        try:
            v = float(entry.get())
        except:
            v = 0
        valor["renda_fixa"] = v
        root.destroy()

    root = tk.Tk()
    root.title("Renda Fixa")

    tk.Label(root, text="Digite o valor total investido em Renda Fixa (R$):",
             font=("Arial", 12, "bold")).grid(row=0, column=0, padx=10, pady=10)

    entry = ttk.Entry(root, width=15)
    entry.grid(row=1, column=0, padx=10)
    entry.insert(0, "0")

    btn = ttk.Button(root, text="Confirmar", command=confirmar)
    btn.grid(row=2, column=0, pady=10)

    root.mainloop()

    return valor["renda_fixa"]

# ---------------- FUNÇÃO PARA COLETAR PREÇOS -------------------


def coletar_precos(lista, nome_coluna):

    quantidades = solicitar_quantidades(lista)

    dados = {}

    for item in lista:
        tk_item = yf.Ticker(item)
        hist = tk_item.history(period="1d")

        if hist.empty:
            print(f"Atenção: {item} não retornou dados.")
            preco = None
        else:
            preco = hist["Close"].iloc[-1]

        quantidade = quantidades.get(item, 0)
        total_investido = preco * quantidade if preco is not None else None

        dados[item] = {
            "Preço Atual": preco,
            "Quantidade": quantidade,
            "Total Investido": total_investido,
            "Atualizado em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}

    df = pd.DataFrame({
        nome_coluna: list(dados.keys()),
        "Preço Atual": [d["Preço Atual"] for d in dados.values()],
        "Quantidade": [d["Quantidade"] for d in dados.values()],
        "Total Investido": [d["Total Investido"] for d in dados.values()],
        "Atualizado em": [d["Atualizado em"] for d in dados.values()]
    })

    return df

# ---------------- FUNÇÃO PARA COLETAR PREÇOS EM DOLAR -------------------


def coletar_precos_internacionais(lista):
    valores_dolar = solicitar_valores_dolar(lista)

    # Cotação atual do dólar
    dolar = yf.Ticker("USDBRL=X")
    hist = dolar.history(period="1d")
    cotacao_dolar = hist["Close"].iloc[-1] if not hist.empty else None

    dados = {}

    for item in lista:
        valor_usd = valores_dolar.get(item, 0)
        valor_brl = valor_usd * cotacao_dolar if cotacao_dolar is not None else None

        dados[item] = {
            "Valor em USD": valor_usd,
            "Cotação Dólar": cotacao_dolar,
            "Valor em Reais": valor_brl,
            "Atualizado em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }

    df = pd.DataFrame({
        "Ação": list(dados.keys()),
        "Valor em USD": [d["Valor em USD"] for d in dados.values()],
        "Cotação Dólar": [d["Cotação Dólar"] for d in dados.values()],
        "Valor em Reais": [d["Valor em Reais"] for d in dados.values()],
        "Atualizado em": [d["Atualizado em"] for d in dados.values()]
    })

    return df

# ---------------- LISTA DE ATIVOS -------------------


acoes_na = ['ITSA4.SA', 'BBAS3.SA', 'ABEV3.SA', 'VIVT3.SA',
            'VALE3.SA', 'BRAP4.SA', 'CMIG4.SA', 'EGIE3.SA',
            'CSMG3.SA', 'SAPR4.SA']

acoes_int = ['SPHD', 'SDY', 'PFF', 'SDIV']

fiis = ['FLMA11.SA', 'TGAR11.SA', 'GGRC11.SA', 'PCIP11.SA',
        'VISC11.SA', 'BTAL11.SA', 'PVBI11.SA', 'HSML11.SA',
        'RFOF11.SA', 'BTHF11.SA', 'VINO11.SA', 'MXRF11.SA']

# ---------------- COLETA DOS DADOS -------------------

df_na = coletar_precos(acoes_na, "Ação")
df_int = coletar_precos_internacionais(acoes_int)
df_fii = coletar_precos(fiis, "FII")
valor_renda_fixa = solicitar_renda_fixa()

df_renda_fixa = pd.DataFrame({
    "Categoria": ["Renda Fixa"],
    "Total Investido (R$)": [valor_renda_fixa],
    "Atualizado em": [datetime.now().strftime("%d/%m/%Y %H:%M:%S")]
})

# ---------------- EXPORTAÇÃO -------------------

with pd.ExcelWriter(arquivo_excel, engine="openpyxl", mode="w") as writer:
    df_na.to_excel(writer, sheet_name="Ações Nacionais", index=False)
    df_int.to_excel(writer, sheet_name="Ações Internacionais", index=False)
    df_fii.to_excel(writer, sheet_name="FIIs", index=False)
    df_renda_fixa.to_excel(writer, sheet_name="Renda Fixa", index=False)

# ---------------- FORMATAÇÃO -------------------

wb = load_workbook(arquivo_excel)

for aba in ["Ações Nacionais", "Ações Internacionais", "FIIs", "Renda Fixa"]:
    ws = wb[aba]

    # Formatar preço atual (Coluna B)
    for cell in ws["B"][1:]:
        if isinstance(cell.value, (int, float)):
            cell.number_format = 'R$ #,##0.00'

    # Formatar total investido (Coluna D)
    for cell in ws["D"][1:]:
        if isinstance(cell.value, (int, float)):
            cell.number_format = 'R$ #,##0.00'

wb.save(arquivo_excel)

# Mostrar caminho do arquivo
print("Arquivo Excel gerado com sucesso!")
print("Caminho completo:", os.path.abspath(arquivo_excel))

import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
from datetime import datetime
import numpy as np

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())


def pegar_cotacao():
    moeda = combobox_select_moeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f"https://economia.awesomeapi.com.br/{moeda}-BRL/10?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    label_texto_cotacao['text'] = f"A Cotação da Moeda {moeda} no dia {data_cotacao} foi de: R$ {valor_moeda}"


def selecionar_arq():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminho_arquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arq_selecionado['text'] = f"Arquivo Selecionado: {caminho_arquivo}"


def atualizar_cotacoes():
    df = pd.read_excel(var_caminho_arquivo.get())
    moedas = df.iloc[:0]
    data_inicial = calendario_data_inicial.get()
    data_final = calendario_data_final.get()

    ano_inicial = data_inicial[-4:]
    mes_inicial = data_inicial[3:5]
    dia_inicial = data_inicial[:2]

    ano_final = data_final[-4:]
    mes_final = data_final[3:5]
    dia_final = data_final[:2]

    for moeda in moedas:
        link = f"https://economia.awesomeapi.com.br/{moeda}-BRL/10?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}"

        requisicao_moeda = requests.get(link)
        cotacoes = requisicao_moeda.json()
        for cotacao in cotacoes:
            timestamp = int(cotacao['timestamp'])
            bid = float(cotacao["bid"])
            data = datetime.timestamp(timestamp)
            data = datetime.strftime("%d/%m/%Y")
            if data not in df:
                df[data] = np.nan

            df.loc[df.iloc[:,0] == moeda, data] = bid

janela = tk.Tk()

janela.title('Cotação de Moedas')

label_cotacao_moedas = tk.Label(text='Cotação de 1 moeda específica', borderwidth=2, relief='solid')
label_cotacao_moedas.grid(row=0, column=0, padx=10, pady=10, sticky="nsew", columnspan=3)

label_selecao_moedas = tk.Label(text='Selecionar Moeda', anchor='e')
label_selecao_moedas.grid(row=1, column=0, padx=10, pady=10, sticky="nsew", columnspan=2)

combobox_select_moeda = ttk.Combobox(values=lista_moedas)
combobox_select_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

label_selecao_dia = tk.Label(text='Selecione o dia que deseja pegar a cotação', anchor='e')
label_selecao_dia.grid(row=2, column=0, padx=10, pady=10, sticky="nsew", columnspan=2)

calendario_moeda = DateEntry(year=2025, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky="nsew")

label_texto_cotacao = tk.Label(text='')
label_texto_cotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

botao_pegar_cotacao = tk.Button(text="Pegar Cotação", command=pegar_cotacao)
botao_pegar_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky="nsew")


#Cotação de várias moedas

label_cotacaomulti_moedas = tk.Label(text='Cotação de Múltiplas moeda específica', borderwidth=2, relief='solid')
label_cotacaomulti_moedas.grid(row=4, column=0, padx=10, pady=10, sticky="nsew", columnspan=3)

label_selecionar_arq = tk.Label(text='Selecionar um arquivo em Excel com as Moedas na Coluna A')
label_selecionar_arq.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

var_caminho_arquivo = tk.StringVar()

botao_selecionar_arq = tk.Button(text="Clique para selecionar", command=selecionar_arq)
botao_selecionar_arq.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

label_arq_selecionado = tk.Label(text="Nenhum Arquivo Selecionado", anchor='e')
label_arq_selecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

label_data_inicial = tk.Label(text='Data Inicial')
label_data_inicial.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')
label_data_final = tk.Label(text='Data Final')
label_data_final.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')

calendario_data_inicial = DateEntry(year=2025, locale='pt_br')
calendario_data_inicial.grid(row=7, column=1, padx=10, pady=10, sticky='nsew')
calendario_data_final = DateEntry(year=2025, locale='pt_br')
calendario_data_final.grid(row=8, column=1, padx=10, pady=10, sticky='nsew')

botao_atualizar_cotacao = tk.Button(text="Atualizar Cotações", command=atualizar_cotacoes)
botao_atualizar_cotacao.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

label_atualizar_cotacoes = tk.Label(text="")
label_atualizar_cotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')


janela.mainloop()
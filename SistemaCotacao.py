import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import pandas as pd
import requests
from datetime import datetime
from datetime import timedelta
import numpy as np

requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())


def pegar_cotacao():
    try:
        moeda = combobox_selecionarmoeda.get()
        data_cotacao = calendario_moeda.get()
        ano = data_cotacao[-4:]
        mes = data_cotacao[3:5]
        dia = data_cotacao[:2]
        link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
        requisicao_moeda = requests.get(link)
        cotacao = requisicao_moeda.json()

        # Verificando se a resposta é uma lista e contém dados
        if isinstance(cotacao, list) and len(cotacao) > 0:
            valor_moeda = cotacao[0].get('bid', None)
            if valor_moeda:
                label_textocotacao['text'] = f"A cotação da {moeda} no dia {data_cotacao} foi de: R${valor_moeda}"
            else:
                label_textocotacao['text'] = f"Não foi possível encontrar o valor da moeda {moeda}."
        else:
            label_textocotacao['text'] = f"Resposta inesperada para a moeda {moeda}."

    except Exception as e:
        label_textocotacao['text'] = f"Erro ao pegar a cotação: {e}"
        print(f"Erro ao pegar cotação: {e}")


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado['text'] = f"Arquivo Selecionado: {caminho_arquivo}"




def atualizar_cotacao():
    try:
        # Verificando se um arquivo foi selecionado
        caminho_arquivo = var_caminhoarquivo.get()
        if not caminho_arquivo:  # Se o caminho estiver vazio
            print("Um arquivo Excel válido deve ser selecionado.")
            label_atualizarcotacoes['text'] = "Um arquivo Excel válido deve ser selecionado."
            return  # Interrompe a execução do código se nenhum arquivo foi selecionado
        
        # Lendo o arquivo Excel e verificando as moedas
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0].dropna()  # Captura as moedas ignorando valores nulos

        if moedas.empty:
            label_atualizarcotacoes['text'] = "O arquivo Excel está vazio ou mal formatado."
            return

        # Capturando as datas inicial e final
        data_inicial = calendario_datainicial.get()
        data_final = calendario_datafinal.get()
        
        # Convertendo as datas para objetos datetime
        data_inicial_dt = datetime.strptime(data_inicial, '%d/%m/%Y')
        data_final_dt = datetime.strptime(data_final, '%d/%m/%Y')

        print(f"Período: {data_inicial} a {data_final}")

        # Loop por cada moeda
        for moeda in moedas:  # Ignorar a primeira linha que contém o cabeçalho
            print(f"Processando moeda: {moeda}")
            
            # Iterando sobre o intervalo de datas
            data_atual = data_inicial_dt
            while data_atual <= data_final_dt:
                data_str = data_atual.strftime('%Y%m%d')  # Formato da data para a API
                link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={data_str}&end_date={data_str}"
                print(f"URL da API para {data_str}: {link}")
                
                requisicao_moeda = requests.get(link)

                try:
                    cotacoes = requisicao_moeda.json()

                    # Verificando se a resposta é uma lista
                    if isinstance(cotacoes, list) and len(cotacoes) > 0:
                        cotacao = cotacoes[0]  # Apenas um resultado para cada dia
                        timestamp = int(cotacao.get('timestamp', 0))
                        bid = float(cotacao.get('bid', 0))
                        data = datetime.fromtimestamp(timestamp).strftime('%d/%m/%Y')

                        # Adicionando a coluna para a data, se não existir
                        if data not in df.columns:
                            df[data] = np.nan

                        # Atualizando o valor no DataFrame
                        df.loc[df.iloc[:, 0] == moeda, data] = bid
                        print(f"Atualizado {moeda} para a data {data} com valor {bid}")
                    else:
                        print(f"Nenhuma cotação disponível para {moeda} em {data_atual.strftime('%d/%m/%Y')}.")

                except Exception as e:
                    print(f"Erro ao processar a moeda {moeda} em {data_atual.strftime('%d/%m/%Y')}: {e}")
                
                # Avançando para o próximo dia
                data_atual += timedelta(days=1)

            # Mostrando o estado atual do DataFrame após a atualização da moeda
            print("Estado atual do DataFrame:")
            print(df.head())

        # Salvando o DataFrame atualizado
        df.to_excel("Resultado_Cotacoes_Moedas.xlsx", index=False)
        label_atualizarcotacoes['text'] = "Arquivo atualizado com sucesso!"
        print("Arquivo salvo com as cotações atualizadas.")

    except Exception as e:
        label_atualizarcotacoes['text'] = f"Erro geral: {e}"
        print(f"Erro: {e}")


janela = tk.Tk()

janela.title('Ferramenta de Cotação de Moedas')

label_cotacaomoeda = tk.Label(text="Cotação de 1 moeda específica", borderwidth=2, relief='solid')
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionarmoeda = tk.Label(text="Selecionar Moeda", anchor='e')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

combobox_selecionarmoeda = ttk.Combobox(values=lista_moedas)
combobox_selecionarmoeda.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

label_selecionardia = tk.Label(text="Selecione o dia que deseja pegar a cotação", anchor='e')
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

calendario_moeda = DateEntry(year=2024, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nsew')

label_textocotacao = tk.Label(text="")
label_textocotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_pegarcotacao = tk.Button(text="Pegar Cotação", command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')


# cotação de várias moedas

label_cotacavariasmoedas = tk.Label(text="Cotação de Múltiplas Moedas", borderwidth=2, relief='solid')
label_cotacavariasmoedas.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=3)

label_selecionararquivo = tk.Label(text='Selecione um arquivo em Excel com as Moedas na Coluna A')
label_selecionararquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text="Clique para Selecionar", command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

label_arquivoselecionado = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
label_arquivoselecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')

label_datainicial = tk.Label(text="Data Inicial", anchor='e')
label_datafinal = tk.Label(text="Data Final", anchor='e')
label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky='nsew')
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky='nsew')

calendario_datainicial = DateEntry(year=2024, locale='pt_br')
calendario_datafinal = DateEntry(year=2024, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, padx=10, pady=10,  sticky='nsew')
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky='nsew')

botao_atualizarcotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacao)
botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nsew')

label_atualizarcotacoes = tk.Label(text="")
label_atualizarcotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='nsew')

botao_fechar = tk.Button(text='Fechar', command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nsew')

janela.mainloop()
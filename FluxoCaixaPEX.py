import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime, timedelta

import re

# Função para extrair o número do protocolo
def extrair_protocolo(texto):
    if pd.isna(texto):
        return None
    texto_str = str(texto)
    match = re.search(r'\d+', texto_str)
    return match.group(0) if match else None

# Função para processar os dados para a aba "Análise por Protocolos"
# Função para processar os dados para a aba "Análise por Protocolos"

def processar_analise_protocolos(df):
    data_hoje = datetime.now().date()
    data_ontem = data_hoje - timedelta(days=1)

    df['Protocolo'] = df['Nº do pedido'].apply(extrair_protocolo)
    df['Data Status'] = pd.to_datetime(df['Data da venda'], dayfirst=True).dt.strftime('%d/%m/%Y')
    df['Valor'] = df['Valor Líquido'].apply(formatar_valor)

    # Adicionando a coluna 'Status Pagamento'
    def determinar_status(data_liquidacao):
        if pd.isnull(data_liquidacao):
            return 'Sem previsão'
        data_liquidacao = pd.to_datetime(data_liquidacao)
        if data_liquidacao.date() <= data_ontem:
            return 'Liquidado'
        elif data_liquidacao.date() >= data_hoje:
            return 'Pgto prometido'
        else:
            return 'Outro'

    df['Status Pagamento'] = df['Data de liquidação'].apply(determinar_status)

    df_analise = df[['Protocolo', 'Data Status', 'Valor', 'Status Pagamento']]
    df_analise = df_analise.sort_values(by='Data Status')
    return df_analise

# Função para adicionar a aba "Análise por Protocolos"
def adicionar_aba_analise_protocolos(notebook, df_analise):
    global tree_analise_protocolos

    # Verifica se a aba já existe
    aba_existente = False
    for tab in notebook.tabs():
        if notebook.tab(tab, "text") == "Análise por Protocolos":
            aba_existente = True
            frame_analise_protocolos = notebook.nametowidget(tab)
            break

    if not aba_existente:
        frame_analise_protocolos = ttk.Frame(notebook)
        notebook.add(frame_analise_protocolos, text="Análise por Protocolos")
        tree_analise_protocolos = ttk.Treeview(frame_analise_protocolos)
    else:
        tree_analise_protocolos = frame_analise_protocolos.winfo_children()[0]
        tree_analise_protocolos.delete(*tree_analise_protocolos.get_children())

    # Configuração das colunas da Treeview
    tree_analise_protocolos['columns'] = ('Protocolo', 'Data Status', 'Valor', 'Status Pagamento')
    tree_analise_protocolos.column('Protocolo', width=100)
    tree_analise_protocolos.column('Data Status', width=100)
    tree_analise_protocolos.column('Valor', width=100)
    tree_analise_protocolos.column('Status Pagamento', width=100)
    
    tree_analise_protocolos.heading('Protocolo', text='Protocolo')
    tree_analise_protocolos.heading('Data Status', text='Data Status')
    tree_analise_protocolos.heading('Valor', text='Valor')
    tree_analise_protocolos.heading('Status Pagamento', text='Status Pagamento')

    tree_analise_protocolos.pack(expand=True, fill='both')

    def buscar_protocolo():
        protocolo_buscado = entry_busca.get()
        protocolo_encontrado = False
        for item in tree_analise_protocolos.get_children():
            if tree_analise_protocolos.item(item, 'values')[0] == protocolo_buscado:
                tree_analise_protocolos.selection_set(item)
                tree_analise_protocolos.see(item)
                protocolo_encontrado = True
                break
        if not protocolo_encontrado:
            messagebox.showinfo("Busca", "Protocolo não encontrado.")

    entry_busca = ttk.Entry(frame_analise_protocolos, width=20)
    entry_busca.pack(pady=5)

    btn_buscar = ttk.Button(frame_analise_protocolos, text="Buscar", command=buscar_protocolo)
    btn_buscar.pack(pady=5)

    def filtrar_dados():
        coluna_filtro = combobox_filtro.get()
        valor_filtro = entry_filtro.get().lower()
        
        for item in tree_analise_protocolos.get_children():
            if coluna_filtro == 'Data Status':
                indice_coluna = 1
            elif coluna_filtro == 'Status Pagamento':
                indice_coluna = 3
            else:
                continue

            if valor_filtro in tree_analise_protocolos.item(item, 'values')[indice_coluna].lower():
                tree_analise_protocolos.item(item, tags='match')
            else:
                tree_analise_protocolos.item(item, tags='nomatch')

        tree_analise_protocolos.tag_configure('match', background='white')
        tree_analise_protocolos.tag_configure('nomatch', background='gray')

        # Adicione isso na função adicionar_aba_analise_protocolos
        opcoes_filtro = ['Data Status', 'Status Pagamento']
        combobox_filtro = ttk.Combobox(frame_analise_protocolos, values=opcoes_filtro)
        combobox_filtro.pack(pady=5)
        combobox_filtro.set('Escolha o filtro')

        entry_filtro = ttk.Entry(frame_analise_protocolos, width=20)
        entry_filtro.pack(pady=5)

        btn_filtrar = ttk.Button(frame_analise_protocolos, text="Filtrar", command=filtrar_dados)
        btn_filtrar.pack(pady=5)


    # Insere os dados na Treeview
    for index, row in df_analise.iterrows():
        tree_analise_protocolos.insert('', 'end', values=(row['Protocolo'], row['Data Status'], row['Valor'], row['Status Pagamento']))


def carregar_arquivo():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if file_path:
            vendas_df = pd.read_excel(file_path)

            # Processamento e exibição dos dados nas abas existentes
            processar_e_exibir_dados(vendas_df)

            # Processamento e exibição dos dados na aba "Análise por Protocolos"
            df_analise = processar_analise_protocolos(vendas_df)
            adicionar_aba_analise_protocolos(notebook, df_analise)
    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo não encontrado.")
    except pd.errors.EmptyDataError:
        messagebox.showerror("Erro", "O arquivo está vazio.")
    except pd.errors.ParserError:
        messagebox.showerror("Erro", "Erro ao analisar o arquivo Excel.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro desconhecido: {str(e)}")

def formatar_data(data):
    return data.strftime('%d/%m/%Y') if not pd.isnull(data) else ''

def formatar_valor(valor):
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def processar_e_exibir_dados(df):
    df['Data da venda'] = pd.to_datetime(df['Data da venda'], dayfirst=True, errors='coerce').dt.date
    df['Data de liquidação'] = pd.to_datetime(df['Data de liquidação'], dayfirst=True, errors='coerce').dt.date
    data_hoje = datetime.now().date()
    data_ontem = data_hoje - timedelta(days=1)

    total_gerado = df[~df['Status'].isin(['Cancelada', 'Rejeitada'])].groupby('Data da venda')['Valor Bruto'].sum().reset_index()
    total_gerado['Data da venda'] = total_gerado['Data da venda'].apply(formatar_data)
    total_gerado['Valor Bruto'] = total_gerado['Valor Bruto'].apply(formatar_valor)

    total_promessa = df.groupby('Data de liquidação')['Valor Líquido'].sum().reset_index()
    total_promessa['Data de liquidação'] = total_promessa['Data de liquidação'].apply(formatar_data)
    total_promessa['Valor Líquido'] = total_promessa['Valor Líquido'].apply(formatar_valor)

    valor_liquidado = df[(df['Data de liquidação'] <= data_ontem) & ~df['Status'].isin(['Cancelada', 'Rejeitada'])]['Valor Líquido'].sum()
    pagamento_prometido = df[(df['Data de liquidação'] >= data_hoje) & ~df['Status'].isin(['Cancelada', 'Rejeitada'])]['Valor Líquido'].sum()
    pagamento_sem_previsao = df[df['Data de liquidação'].isna() & ~df['Status'].isin(['Cancelada', 'Rejeitada'])]['Valor Líquido'].sum()

    df_validas = df[~df['Status'].isin(['Cancelada', 'Rejeitada'])]
    
    # Alteração aqui: Agrupamento por Forma de Pagamento e Data da venda
    liquidado_por_pagamento = df_validas[df_validas['Data de liquidação'] <= data_ontem].groupby(['Forma de Pagamento', 'Data da venda']).agg({'Valor Líquido': 'sum'}).reset_index()
    prometido_por_pagamento = df_validas[df_validas['Data de liquidação'] >= data_hoje].groupby(['Forma de Pagamento', 'Data da venda']).agg({'Valor Líquido': 'sum'}).reset_index()
    sem_previsao_por_pagamento = df_validas[df_validas['Data de liquidação'].isna()].groupby(['Forma de Pagamento', 'Data da venda']).agg({'Valor Líquido': 'sum'}).reset_index()

    exibir_dados(total_gerado, total_promessa, formatar_valor(valor_liquidado), formatar_valor(pagamento_prometido), formatar_valor(pagamento_sem_previsao), liquidado_por_pagamento, prometido_por_pagamento, sem_previsao_por_pagamento)


def exibir_dados(total_gerado, total_promessa, valor_liquidado, pagamento_prometido, pagamento_sem_previsao, liquidado_por_pagamento, prometido_por_pagamento, sem_previsao_por_pagamento):
    for tree in [tree_total_gerado, tree_total_promessa, tree_liquidado, tree_prometido, tree_sem_previsao]:
        for i in tree.get_children():
            tree.delete(i)

    for row in total_gerado.to_numpy().tolist():
        tree_total_gerado.insert('', 'end', values=row)

    for row in total_promessa.to_numpy().tolist():
        tree_total_promessa.insert('', 'end', values=row)

    # Exibindo dados agrupados por forma de pagamento e data
    for index, row in liquidado_por_pagamento.iterrows():
        tree_liquidado.insert('', 'end', values=(row['Forma de Pagamento'], formatar_data(row['Data da venda']), formatar_valor(row['Valor Líquido'])))

    for index, row in prometido_por_pagamento.iterrows():
        tree_prometido.insert('', 'end', values=(row['Forma de Pagamento'], formatar_data(row['Data da venda']), formatar_valor(row['Valor Líquido'])))

    for index, row in sem_previsao_por_pagamento.iterrows():
        tree_sem_previsao.insert('', 'end', values=(row['Forma de Pagamento'], formatar_data(row['Data da venda']), formatar_valor(row['Valor Líquido'])))

    lbl_valor_liquidado.config(text=f"Valor Total Já Liquidado: {valor_liquidado}")
    lbl_pagamento_prometido.config(text=f"Pagamento Prometido: {pagamento_prometido}")
    lbl_pagamento_sem_previsao.config(text=f"Pagamento Sem Previsão: {pagamento_sem_previsao}")


root = tk.Tk()
root.title("Análise de Vendas")

# Define o tamanho da janela para 800x600 e desabilita o redimensionamento
root.geometry("800x600")
root.resizable(False, False)

# Cria o Notebook
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill='both', padx=10, pady=10)

frame_total_gerado = ttk.Frame(notebook)
notebook.add(frame_total_gerado, text="Total Gerado por Data")
tree_total_gerado = ttk.Treeview(frame_total_gerado, columns=['Data', 'Valor Bruto'], show='headings')
tree_total_gerado.heading('Data', text='Data da Venda')
tree_total_gerado.heading('Valor Bruto', text='Valor Bruto Total')
tree_total_gerado.pack(expand=True, fill='both')

frame_total_promessa = ttk.Frame(notebook)
notebook.add(frame_total_promessa, text="Total de Promessa por Data")
tree_total_promessa = ttk.Treeview(frame_total_promessa, columns=['Data', 'Valor Líquido'], show='headings')
tree_total_promessa.heading('Data', text='Data de Liquidação')
tree_total_promessa.heading('Valor Líquido', text='Valor Líquido Total')
tree_total_promessa.pack(expand=True, fill='both')

frame_forma_pagamento = ttk.Frame(notebook)
notebook.add(frame_forma_pagamento, text="Por Forma de Pagamento")

notebook_forma_pagamento = ttk.Notebook(frame_forma_pagamento)
notebook_forma_pagamento.pack(expand=True, fill='both', padx=10, pady=10)

frame_liquidado = ttk.Frame(notebook_forma_pagamento)
notebook_forma_pagamento.add(frame_liquidado, text="Liquidado")
tree_liquidado = ttk.Treeview(frame_liquidado, columns=['Forma de Pagamento','Data Liquidado', 'Valor Líquido'], show='headings')
tree_liquidado.heading('Forma de Pagamento', text='Forma de Pagamento')
tree_liquidado.heading('Valor Líquido', text='Valor total Liquidado')
tree_liquidado.heading('Data Liquidado', text='Data Liquidado')
tree_liquidado.pack(expand=True, fill='both')

frame_prometido = ttk.Frame(notebook_forma_pagamento)
notebook_forma_pagamento.add(frame_prometido, text="Prometido")
tree_prometido = ttk.Treeview(frame_prometido, columns=['Forma de Pagamento', 'Data Prometida', 'Valor Líquido'], show='headings')
tree_prometido.heading('Forma de Pagamento', text='Forma de Pagamento')
tree_prometido.heading('Data Prometida', text='Data Prometida')
tree_prometido.heading('Valor Líquido', text='Valor total Liquido')
tree_prometido.pack(expand=True, fill='both')

frame_sem_previsao = ttk.Frame(notebook_forma_pagamento)
notebook_forma_pagamento.add(frame_sem_previsao, text="Sem Previsão")
tree_sem_previsao = ttk.Treeview(frame_sem_previsao, columns=['Forma de Pagamento', 'Data', 'Valor Líquido'], show='headings')
tree_sem_previsao.heading('Forma de Pagamento', text='Forma de Pagamento')
tree_sem_previsao.heading('Data', text='Data da Venda')
tree_sem_previsao.heading('Valor Líquido', text='Sem Previsão')
tree_sem_previsao.pack(expand=True, fill='both')

# Aba "Análise por Protocolos"
frame_analise_protocolos = ttk.Frame(notebook)
notebook.add(frame_analise_protocolos, text="Análise por Protocolos")
tree_analise_protocolos = ttk.Treeview(frame_analise_protocolos, columns=['Protocolo', 'Data da venda', 'Valor Líquido'], show='headings')
tree_analise_protocolos.heading('Protocolo', text='Protocolo')
tree_analise_protocolos.heading('Data da venda', text='Data da Venda')
tree_analise_protocolos.heading('Valor Líquido', text='Valor Líquido')
tree_analise_protocolos.pack(expand=True, fill='both')


lbl_valor_liquidado = ttk.Label(root, text="Valor Total Já Liquidado: ")
lbl_valor_liquidado.pack(pady=5)

lbl_pagamento_prometido = ttk.Label(root, text="Pagamento Prometido: ")
lbl_pagamento_prometido.pack(pady=5)

lbl_pagamento_sem_previsao = ttk.Label(root, text="Pagamento Sem Previsão: ")
lbl_pagamento_sem_previsao.pack(pady=5)

btn_carregar = ttk.Button(root, text="Carregar Arquivo Excel", command=carregar_arquivo)
btn_carregar.pack(pady=10)

root.mainloop()

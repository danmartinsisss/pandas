import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
import re

def upload_pex():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        files['PEX'] = file_path
        lbl_pex.config(text=f"PEX: {file_path.split('/')[-1]}")

def upload_ecash():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        files['Ecash'] = file_path
        lbl_ecash.config(text=f"Ecash: {file_path.split('/')[-1]}")

def extract_digits(s):
    """ Extrai os primeiros 6 dígitos de uma string. """
    match = re.search(r'\d{6}', s)
    return match.group(0) if match else None

def compare_files():
    if files['PEX'] and files['Ecash']:
        pex_df = pd.read_excel(files['PEX'], usecols=['Nº do pedido','Valor Líquido'])
        ecash_df = pd.read_excel(files['Ecash'], usecols=['historico','entrada'])

        pex_df['Nº do pedido'] = pex_df['Nº do pedido'].astype(str).apply(extract_digits)
        ecash_df['historico'] = ecash_df['historico'].astype(str).apply(extract_digits)

        # Realiza a comparação
        result = pex_df.merge(ecash_df, left_on='Nº do pedido', right_on='historico', how='left', indicator=True)
        missing = result[result['_merge'] == 'left_only']

        # Calcula os valores
        total_valores = pex_df['Valor Líquido'].sum()
        total_nao_encontrados = missing['Valor Líquido'].sum()

        # Atualiza os rótulos de resumo
        lbl_total_valores.config(text=f"Valores encontrados: R$ {total_valores:.2f}")
        lbl_total_nao_encontrados.config(text=f"Valores não encontrados na planilha Ecash: R$ {total_nao_encontrados:.2f}")

        # Limpa a árvore antes de adicionar novos itens
        for i in tree.get_children():
            tree.delete(i)

        # Mostra os resultados
        for index, row in missing.iterrows():
            tree.insert("", "end", values=(row['Nº do pedido'], row['Valor Líquido'], index+2))  # +2 para contar o cabeçalho do Excel


def clear_all():
    # Limpa as seleções e os rótulos
    lbl_pex.config(text="PEX: Nenhum arquivo selecionado")
    lbl_ecash.config(text="Ecash: Nenhum arquivo selecionado")
    lbl_total_valores.config(text="Valores encontrados: R$ 0.00")
    lbl_total_nao_encontrados.config(text="Valores não encontrados na planilha Ecash: R$ 0.00")
    
    # Limpa o Treeview
    for i in tree.get_children():
        tree.delete(i)
    
    # Reseta o dicionário de arquivos
    files['PEX'] = ''
    files['Ecash'] = ''

def exit_app():
    root.destroy()

# def print_results():
#     # Coleta os dados da Treeview
#     heads = [tree.heading(col)['text'] for col in tree['columns']]
#     data = [tree.item(item)['values'] for item in tree.get_children()]
#     result_text = "\n".join(["\t".join(map(str, row)) for row in ([heads] + data)])
    
#     # Cria um arquivo temporário para a impressão
#     filename = tempfile.mktemp(".txt")
#     open(filename, "w").write(result_text)
    
#     # Usa a API do Windows para imprimir o arquivo de texto
#     win32api.ShellExecute(
#         0,
#         "print",
#         filename,
#         '/d:"%s"' % win32print.GetDefaultPrinter(),
#         ".",
#         0
#     )

root = tk.Tk()
root.title("Comparador de Planilhas")

# Define a geometria da janela
root.geometry("800x600")

files = {'PEX': '', 'Ecash': ''}

frame_buttons = tk.Frame(root)
frame_buttons.grid(row=0, column=0, sticky='nw', padx=5, pady=5)

btn_upload_pex = tk.Button(frame_buttons, text="Carregar PEX", command=upload_pex)
btn_upload_pex.grid(row=0, column=0, padx=5, pady=2)

lbl_pex = tk.Label(frame_buttons, text="PEX: Nenhum arquivo selecionado")
lbl_pex.grid(row=1, column=0, padx=5, pady=2)

btn_upload_ecash = tk.Button(frame_buttons, text="Carregar Ecash", command=upload_ecash)
btn_upload_ecash.grid(row=2, column=0, padx=5, pady=2)

lbl_ecash = tk.Label(frame_buttons, text="Ecash: Nenhum arquivo selecionado")
lbl_ecash.grid(row=3, column=0, padx=5, pady=2)

btn_compare = tk.Button(frame_buttons, text="Comparar", command=compare_files)
btn_compare.grid(row=4, column=0, padx=5, pady=2)

frame_summary = tk.Frame(root)
frame_summary.grid(row=0, column=1, sticky='nw', padx=5, pady=5)

lbl_total_valores = tk.Label(frame_summary, text="Valores encontrados: R$ 0.00")
lbl_total_valores.grid(row=0, column=0, padx=5, pady=2)

lbl_total_nao_encontrados = tk.Label(frame_summary, text="Valores não encontrados na planilha Ecash: R$ 0.00")
lbl_total_nao_encontrados.grid(row=1, column=0, padx=5, pady=2)

btn_clear = tk.Button(frame_buttons, text="Limpar", command=clear_all)
btn_clear.grid(row=5, column=0, padx=5, pady=2)

# btn_print = tk.Button(frame_buttons, text="Imprimir Resultados", command=print_results)
# btn_print.grid(row=6, column=0, padx=5, pady=2)

tree = ttk.Treeview(root, columns=('Pedido', 'Valor', 'Linha Excel'), show='headings')
tree.heading('Pedido', text='Nº do pedido')
tree.heading('Valor', text='Valor Bruto')
tree.heading('Linha Excel', text='Linha Excel')
tree.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky='nsew')

# Configura as colunas para que a Treeview se expanda e preencha o espaço
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(1, weight=1)

root.mainloop()
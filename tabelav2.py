import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pandas as pd

# Função para salvar os dados na tabela
def salvar_dados():
    nome = entry_nome.get()
    folga_fixa = entry_folga.get()
    cargo = entry_cargo.get()
    data_entrada = entry_data_entrada.get()
    salario = entry_salario.get()

    # Verifica se todos os campos foram preenchidos
    if not nome or not folga_fixa or not cargo or not data_entrada or not salario:
        messagebox.showwarning("Aviso", "Por favor, preencha todos os campos!")
        return

    # Adiciona os dados na tabela
    tabela.insert('', 'end', values=(nome, folga_fixa, cargo, data_entrada, salario))
    
    # Limpa os campos após salvar
    entry_nome.delete(0, tk.END)
    entry_folga.delete(0, tk.END)
    entry_cargo.delete(0, tk.END)
    entry_data_entrada.delete(0, tk.END)
    entry_salario.delete(0, tk.END)

# Função para exportar os dados para Excel
def exportar_excel():
    # Obtendo os dados da tabela
    dados = []
    for item in tabela.get_children():
        dados.append(tabela.item(item)['values'])
    
    # Definindo os cabeçalhos da tabela
    colunas = ["Nome", "Folga Fixa", "Cargo", "Data de Entrada", "Salário"]

    # Criando um DataFrame a partir dos dados
    df = pd.DataFrame(dados, columns=colunas)

    # Salvando o DataFrame como arquivo Excel
    df.to_excel("folga_funcionarios.xlsx", index=False, engine='openpyxl')
    messagebox.showinfo("Exportação", "Dados exportados com sucesso para o arquivo 'folga_funcionarios.xlsx'")

# Função para editar dados da tabela
def editar_dados():
    # Obtém o item selecionado na tabela
    item_selecionado = tabela.selection()
    
    if not item_selecionado:
        messagebox.showwarning("Aviso", "Selecione uma linha para editar!")
        return

    # Obtém os valores da linha selecionada
    valores = tabela.item(item_selecionado[0])['values']
    
    # Preenche os campos de entrada com os valores da linha selecionada
    entry_nome.delete(0, tk.END)
    entry_nome.insert(0, valores[0])
    
    entry_folga.delete(0, tk.END)
    entry_folga.insert(0, valores[1])
    
    entry_cargo.delete(0, tk.END)
    entry_cargo.insert(0, valores[2])
    
    entry_data_entrada.delete(0, tk.END)
    entry_data_entrada.insert(0, valores[3])
    
    entry_salario.delete(0, tk.END)
    entry_salario.insert(0, valores[4])

    # Exclui a linha selecionada da tabela, pois os dados serão editados e inseridos novamente um loop melhor dizendo 
    
    tabela.delete(item_selecionado[0])

# Função para excluir dados da tabela
def excluir_dados():
    # Obtém o item selecionado na tabela
    item_selecionado = tabela.selection()
    
    if not item_selecionado:
        messagebox.showwarning("Aviso", "Selecione uma linha para excluir!")
        return

    # Remove o item da tabela
    tabela.delete(item_selecionado[0])

# Criando a janela principal p o desktop
root = tk.Tk()
root.title("Cadastro de Funcionários - Folga Completa")
root.geometry("800x500")

# Criando os campos de entrada
label_nome = tk.Label(root, text="Nome do Funcionário:")
label_nome.pack()
entry_nome = tk.Entry(root)
entry_nome.pack()

label_folga = tk.Label(root, text="Folga Fixa:")
label_folga.pack()
entry_folga = tk.Entry(root)
entry_folga.pack()

label_cargo = tk.Label(root, text="Cargo:")
label_cargo.pack()
entry_cargo = tk.Entry(root)
entry_cargo.pack()

label_data_entrada = tk.Label(root, text="Data de Entrada (DD/MM/AAAA):")
label_data_entrada.pack()
entry_data_entrada = tk.Entry(root)
entry_data_entrada.pack()

label_salario = tk.Label(root, text="Salário:")
label_salario.pack()
entry_salario = tk.Entry(root)
entry_salario.pack()

# Botão para salvar os dados
botao_salvar = tk.Button(root, text="Salvar Dados", command=salvar_dados)
botao_salvar.pack()

# Botão para exportar para Excel
botao_exportar = tk.Button(root, text="Exportar para Excel", command=exportar_excel)
botao_exportar.pack()

# Botão para editar os dados
botao_editar = tk.Button(root, text="Editar Dados", command=editar_dados)
botao_editar.pack()

# Botão para excluir os dados
botao_excluir = tk.Button(root, text="Excluir Dados", command=excluir_dados)
botao_excluir.pack()

# Tabela para exibir os dados
colunas = ("Nome", "Folga Fixa", "Cargo", "Data de Entrada", "Salário")
tabela = ttk.Treeview(root, columns=colunas, show="headings")

# Definindo os cabeçalhos da tabela
for col in colunas:
    tabela.heading(col, text=col)

# Exibindo a tabela
tabela.pack(fill=tk.BOTH, expand=True)

# Rodando a interface gráfica
root.mainloop()

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# Tabela de pontos atualizada da ABEP
pontos_bens = {
    'Banheiros': [0, 3, 7, 10, 14],
    'Trabalhadores domésticos': [0, 3, 7, 10, 13],
    'Automóveis': [0, 3, 5, 8, 11],
    'Microcomputador': [0, 3, 6, 8, 11],
    'Lava louça': [0, 3, 6, 6, 6],
    'Geladeira': [0, 2, 3, 5, 5],
    'Freezer': [0, 2, 4, 6, 6],
    'Lava roupa': [0, 2, 4, 6, 6],
    'DVD': [0, 1, 3, 4, 4],
    'Micro-ondas': [0, 2, 4, 4, 4],
    'Motocicleta': [0, 1, 3, 3, 3],
    'Secadora roupa': [0, 2, 2, 2, 2]
}

pontos_escolaridade = {
    'Analfabeto / Fundamental I incompleto': 0,
    'Fundamental I completo / Fundamental II incompleto': 1,
    'Fundamental II completo / Médio incompleto': 2,
    'Médio completo / Superior incompleto': 4,
    'Superior completo': 7
}

pontos_servicos_publicos = {
    'Água encanada': {'Não': 0, 'Sim': 4},
    'Rua pavimentada': {'Não': 0, 'Sim': 2}
}

# Função para calcular a pontuação total de um domicílio
def calcular_pontos(row, bens_cols, escolaridade, servicos_publicos_cols):
    pontos_totais = 0
    for col, valor in bens_cols.items():
        try:
            quantidade = int(row[valor]) if pd.notnull(row[valor]) and row[valor] != '' else 0
            if quantidade >= 4:
                pontos_totais += pontos_bens[col][4]
            else:
                pontos_totais += pontos_bens[col][quantidade]
        except ValueError:
            continue
    
    for servico, col in servicos_publicos_cols.items():
        resposta = str(row[col]).strip().capitalize() if pd.notnull(row[col]) and row[col] != '' else 'Não'
        if resposta in pontos_servicos_publicos[servico]:
            pontos_totais += pontos_servicos_publicos[servico][resposta]
        else:
            messagebox.showerror("Erro", f"Valor inválido '{resposta}' para {servico}")
            return None
    
    pontos_totais += pontos_escolaridade.get(escolaridade, 0)
    return pontos_totais

# Função para classificar o domicílio
def classificar_domicilio(pontos_totais):
    if pontos_totais is None:
        return 'Classificação inválida'
    if pontos_totais >= 45:
        return 'Classe A'
    elif pontos_totais >= 35:
        return 'Classe B1'
    elif pontos_totais >= 29:
        return 'Classe B2'
    elif pontos_totais >= 23:
        return 'Classe C1'
    elif pontos_totais >= 18:
        return 'Classe C2'
    else:
        return 'Classe D/E'

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            global df_domicilio, df_morador
            df_domicilio = pd.read_excel(file_path, sheet_name='Dados do Domicílio')
            df_morador = pd.read_excel(file_path, sheet_name='Morador')
            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")

def select_columns():
    if df_domicilio is None or df_morador is None:
        messagebox.showerror("Erro", "Por favor, carregue um arquivo primeiro.")
        return

    def calculate_and_save():
        bens_cols = {}
        for bem in pontos_bens.keys():
            bens_cols[bem] = column_dict[bem].get()

        servicos_publicos_cols = {}
        for servico in pontos_servicos_publicos.keys():
            servicos_publicos_cols[servico] = servicos_column_dict[servico].get()

        escolaridade_col = escolaridade_combobox.get()
        id_domicilio_col = id_domicilio_combobox.get()
        
        if not bens_cols or not escolaridade_col or not id_domicilio_col or not servicos_publicos_cols:
            messagebox.showerror("Erro", "Por favor, selecione todas as colunas necessárias.")
            return

        # Seleciona o responsável pelo domicílio
        df_morador_responsavel = df_morador[df_morador['SITUAÇÃO DO MORADOR NO DOMICÍLIO'] == 'Responsável pelo domicílio']
        
        # Inicializa as colunas de Pontos Totais e Classe
        df_domicilio['Pontos Totais'] = 0
        df_domicilio['Classe'] = ''

        for idx, row in df_domicilio.iterrows():
            id_domicilio = row[id_domicilio_col]
            responsavel = df_morador_responsavel[df_morador_responsavel[id_domicilio_col] == id_domicilio]
            
            if not responsavel.empty:
                escolaridade = responsavel[escolaridade_col].values[0]
                pontos_totais = calcular_pontos(row, bens_cols, escolaridade, servicos_publicos_cols)
                if pontos_totais is None:
                    continue  # Ignora este domicílio em caso de erro
                classe = classificar_domicilio(pontos_totais)
                
                df_domicilio.at[idx, 'Pontos Totais'] = pontos_totais
                df_domicilio.at[idx, 'Classe'] = classe
            else:
                df_domicilio.at[idx, 'Pontos Totais'] = 0
                df_domicilio.at[idx, 'Classe'] = 'Classe D/E'

        df_classificacao = df_domicilio[[id_domicilio_col, 'Pontos Totais', 'Classe']]
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            df_classificacao.to_excel(save_path, index=False)
            messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")
    
    columns_window = tk.Toplevel(root)
    columns_window.title("Seleção de Colunas")
    columns_window.configure(bg='#2E2E2E')  # Fundo cinza escuro
    
    container = ttk.Frame(columns_window)
    canvas = tk.Canvas(container, bg='#2E2E2E')  # Fundo cinza escuro
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    container.pack(fill="both", expand=True)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    ttk.Button(scrollable_frame, text="Calcular e Salvar", command=calculate_and_save, bootstyle="success-outline", width=20).pack(pady=10)
    
    column_dict = {}
    for bem in pontos_bens.keys():
        ttk.Label(scrollable_frame, text=f"Indique a coluna do {bem}:", bootstyle="info", font=('Arial', 12)).pack(pady=5)
        column_dict[bem] = ttk.Combobox(scrollable_frame, values=df_domicilio.columns.tolist(), width=47, bootstyle="primary", font=('Arial', 12))
        column_dict[bem].pack(pady=5)
        column_dict[bem].bind("<KeyRelease>", lambda event, cb=column_dict[bem]: combobox_key_nav(event, cb))

    servicos_column_dict = {}
    for servico in pontos_servicos_publicos.keys():
        ttk.Label(scrollable_frame, text=f"Indique a coluna do {servico}:", bootstyle="info", font=('Arial', 12)).pack(pady=5)
        servicos_column_dict[servico] = ttk.Combobox(scrollable_frame, values=df_domicilio.columns.tolist(), width=47, bootstyle="primary", font=('Arial', 12))
        servicos_column_dict[servico].pack(pady=5)
        servicos_column_dict[servico].bind("<KeyRelease>", lambda event, cb=servicos_column_dict[servico]: combobox_key_nav(event, cb))
    
    ttk.Label(scrollable_frame, text="Selecione a coluna de escolaridade na aba 'Morador':", bootstyle="info", font=('Arial', 12)).pack(pady=5)
    escolaridade_combobox = ttk.Combobox(scrollable_frame, values=df_morador.columns.tolist(), width=47, bootstyle="primary", font=('Arial', 12))
    escolaridade_combobox.pack(pady=5)
    escolaridade_combobox.bind("<KeyRelease>", lambda event, cb=escolaridade_combobox: combobox_key_nav(event, cb))
    
    ttk.Label(scrollable_frame, text="Selecione a coluna ID_DOMICILIO na aba 'Dados do Domicílio':", bootstyle="info", font=('Arial', 12)).pack(pady=5)
    id_domicilio_combobox = ttk.Combobox(scrollable_frame, values=df_domicilio.columns.tolist(), width=47, bootstyle="primary", font=('Arial', 12))
    id_domicilio_combobox.pack(pady=5)
    id_domicilio_combobox.bind("<KeyRelease>", lambda event, cb=id_domicilio_combobox: combobox_key_nav(event, cb))

def combobox_key_nav(event, combobox):
    if event.keysym == "Tab":
        combobox.tk_focusNext().focus()
        return "break"
    if event.keysym.isalnum():
        value = event.keysym.lower()
        for i, item in enumerate(combobox["values"]):
            if item.lower().startswith(value):
                combobox.current(i)
                break

root = ttk.Window(themename="superhero")  # Usar o tema Superhero do ttkbootstrap
root.title("Classificação de Domicílios ABEP")
root.geometry("600x500")  # Definir um tamanho inicial para a janela
root.configure(bg='#2E2E2E')  # Fundo cinza escuro

df_domicilio = None
df_morador = None

# Adicionar título
ttk.Label(root, text="Classificação de Domicílios ABEP", font=("Helvetica", 24, "bold"), background='#2E2E2E', foreground='white').pack(pady=20)

# Adicionar botões
button_frame = ttk.Frame(root, padding=(20, 10))
button_frame.pack(fill=tk.BOTH, expand=True)

ttk.Button(button_frame, text="Carregar Arquivo", command=load_file, bootstyle="success-outline", width=20).pack(pady=10)
ttk.Button(button_frame, text="Selecionar Colunas", command=select_columns, bootstyle="success-outline", width=20).pack(pady=10)

# Rodapé
ttk.Label(root, text="Desenvolvido por Moésio Fiùza", font=("Arial", 10), background='#2E2E2E', foreground='white').pack(side="bottom", pady=10)

root.mainloop()

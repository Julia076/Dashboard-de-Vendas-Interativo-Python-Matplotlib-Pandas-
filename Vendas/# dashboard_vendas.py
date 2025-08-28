import os
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ---------- Função para ler dados do Excel ----------
def criar_dados_vendas(arquivo="vendas.xlsx"):
    """
    Lê dados de vendas a partir de um arquivo Excel.
    Calcula a Receita automaticamente.
    """
    if not os.path.exists(arquivo):
        messagebox.showerror("Erro", f"Arquivo não encontrado: {arquivo}")
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo}")
    
    df = pd.read_excel(arquivo)
    
    # Verificar colunas essenciais
    col_necessarias = ['Produto', 'Categoria', 'Mes', 'Vendas', 'Preco_Unitario']
    for col in col_necessarias:
        if col not in df.columns:
            messagebox.showerror("Erro", f"Coluna '{col}' não encontrada no Excel!")
            raise ValueError(f"Coluna '{col}' não encontrada no Excel!")
    
    # Calcular Receita
    df['Receita'] = df['Vendas'] * df['Preco_Unitario']
    
    return df

# ---------- Funções de métricas ----------
def calcular_metricas_principais(df):
    total_vendas = df['Vendas'].sum()
    total_receita = df['Receita'].sum()
    ticket_medio = total_receita / total_vendas
    vendas_por_produto = df.groupby('Produto')['Vendas'].sum()
    produto_top = vendas_por_produto.idxmax()
    return {
        'total_vendas': total_vendas,
        'total_receita': total_receita,
        'ticket_medio': ticket_medio,
        'produto_top': produto_top
    }

# ---------- Funções de gráficos ----------
def criar_graficos(df):
    fig, axs = plt.subplots(2, 2, figsize=(12, 8))
    
    # Gráfico 1: Vendas por produto
    vendas_produto = df.groupby('Produto')['Vendas'].sum()
    axs[0,0].bar(vendas_produto.index, vendas_produto.values, color='#4ECDC4', alpha=0.8)
    axs[0,0].set_title("Vendas por Produto")
    axs[0,0].set_ylabel("Quantidade")
    
    # Gráfico 2: Receita por categoria
    receita_categoria = df.groupby('Categoria')['Receita'].sum()
    axs[0,1].pie(receita_categoria.values, labels=receita_categoria.index, autopct='%1.1f%%', startangle=90)
    axs[0,1].set_title("Receita por Categoria")
    
    # Gráfico 3: Evolução mensal
    vendas_mensais = df.groupby('Mes')['Receita'].sum()
    axs[1,0].plot(vendas_mensais.index, vendas_mensais.values, marker='o', color='#FF6B6B')
    axs[1,0].set_title("Evolução Mensal da Receita")
    axs[1,0].set_ylabel("Receita (R$)")
    
    # Gráfico 4: vazio
    axs[1,1].axis('off')
    
    plt.tight_layout()
    return fig

# ---------- Função para gerar tabela ----------
def gerar_tabela(df, parent_frame):
    resumo = df.groupby('Produto').agg({
        'Vendas': 'sum',
        'Receita': 'sum',
        'Preco_Unitario': 'mean'
    }).round(2)
    resumo['Ticket_Medio'] = (resumo['Receita'] / resumo['Vendas']).round(2)
    resumo['Participacao_%'] = ((resumo['Receita'] / resumo['Receita'].sum()) * 100).round(1)
    
    tree = ttk.Treeview(parent_frame)
    tree['columns'] = list(resumo.columns)
    tree.heading("#0", text='Produto')
    for col in resumo.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor='center')
    
    for produto, row in resumo.iterrows():
        tree.insert('', 'end', text=produto, values=list(row))
    
    tree.pack(fill='both', expand=True)
    return tree

# ---------- Dashboard completo ----------
def main_dashboard():
    # Caminho do arquivo Excel
    arquivo_excel = r"C:\Users\rodri\OneDrive\Área de Trabalho\Projetos Pessoal\vendas.xlsx"
    
    df = criar_dados_vendas(arquivo_excel)
    metricas = calcular_metricas_principais(df)
    
    root = tk.Tk()
    root.title("Dashboard de Vendas")
    root.geometry("1000x800")
    
    # Resumo executivo
    frame_resumo = tk.Frame(root)
    frame_resumo.pack(pady=10)
    resumo_texto = (
        f"Total de Vendas: {metricas['total_vendas']:,} unidades\n"
        f"Receita Total: R$ {metricas['total_receita']:,.2f}\n"
        f"Ticket Médio: R$ {metricas['ticket_medio']:.2f}\n"
        f"Produto Líder: {metricas['produto_top']}"
    )
    tk.Label(frame_resumo, text=resumo_texto, font=("Arial", 12), justify='left').pack()
    
    # Tabela detalhada
    frame_tabela = tk.LabelFrame(root, text="Resumo Detalhado por Produto")
    frame_tabela.pack(fill='both', expand=True, padx=10, pady=10)
    gerar_tabela(df, frame_tabela)
    
    # Gráficos
    fig = criar_graficos(df)
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)
    
    root.mainloop()

# ---------- Executar dashboard ----------
if __name__ == "__main__":
    main_dashboard()

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import csv
from collections import defaultdict

# Conexão com banco de dados
conn = sqlite3.connect("chamados.db")
cursor = conn.cursor()
cursor.execute('''
    CREATE TABLE IF NOT EXISTS chamados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        numero_chamado TEXT NOT NULL,
        tipo_item TEXT NOT NULL,
        quantidade INTEGER NOT NULL
    )
''')
conn.commit()

# Cores
BG_COLOR = "#1e1e2d"
TABLE_BG = "#3f3f4d"
TEXT_COLOR = "#ffffff"
ENTRY_BG = "#2e2e3d"
ENTRY_FG = TEXT_COLOR
HEADER_BG = "#2e2e3d"
HEADER_FG = TEXT_COLOR

# Lista de tipos fixos
tipos_itens = [
    "Mouse Dell",
    "Mouse Logitech",
    "Mouse Multilazer",
    "Teclado Dell",
    "Teclado Multilazer",
    "Headset Jabra",
    "Headset Logitech",
    "Pilhas AA",
    "Pilhas AAA",
    "Carregador de notebook Dell Type C",
    "Carregador de notebook Lenovo Type C",
    "Carregador de notebook Dell 3420",
    "Carregador de notebook Dell 7490",
    "Carregador de MacBook",
    "Dockstation Prata",
    "Dockstation Preta",
    "Cabo de Rede",
    "Cabo de Energia",
    "Cabo HDMI",
    "Monitor Dell 24 polegadas",
    "Monitor Lenovo 23 polegadas"
]

# Funções

def adicionar_item():
    numero = entry_numero.get()
    tipo = tipo_var.get()
    quantidade = entry_quantidade.get()

    if not (numero and tipo and quantidade):
        messagebox.showwarning("Atenção", "Preencha todos os campos.")
        return

    try:
        quantidade = int(quantidade)
    except ValueError:
        messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")
        return

    cursor.execute("INSERT INTO chamados (numero_chamado, tipo_item, quantidade) VALUES (?, ?, ?)",
                   (numero, tipo, quantidade))
    conn.commit()
    atualizar_tabela()
    entry_numero.delete(0, tk.END)
    entry_numero.insert(0, "SC-")
    tipo_var.set("")
    entry_quantidade.delete(0, tk.END)
    entry_quantidade.insert(0, "1")

def atualizar_tabela():
    for item in tree.get_children():
        tree.delete(item)

    cursor.execute("SELECT numero_chamado, tipo_item FROM chamados")
    chamados = cursor.fetchall()

    agrupado = defaultdict(list)
    for numero, tipo in chamados:
        agrupado[tipo].append(numero)

    tree["columns"] = list(agrupado.keys()) + ["Totais"]
    for col in agrupado:
        tree.heading(col, text=col)
        tree.column(col, width=120)

    tree.heading("Totais", text="Totais")
    tree.column("Totais", width=80)

    max_len = max(len(v) for v in agrupado.values()) if agrupado else 0
    for i in range(max_len):
        row = [agrupado[col][i] if i < len(agrupado[col]) else "" for col in agrupado]
        row.append("")
        tree.insert("", tk.END, values=row)

    totais = [str(len(agrupado[col])) for col in agrupado]
    totais.append("Total")
    tree.insert("", tk.END, values=totais)

def exportar_csv():
    file = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV files", "*.csv")])
    if not file:
        return
    
    # Obtém os dados da tabela
    colunas = tree["columns"]
    dados = []
    
    # Adiciona o cabeçalho
    dados.append(colunas)
    
    # Adiciona os dados das linhas
    for item in tree.get_children():
        valores = tree.item(item)['values']
        dados.append(valores)
    
    # Exporta para CSV com codificação UTF-8-SIG para melhor compatibilidade com Excel
    with open(file, mode='w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')  # Usando ponto e vírgula como separador
        writer.writerows(dados)
    
    messagebox.showinfo("Exportado", f"Dados exportados para {file}")

def resetar_planilha():
    if messagebox.askyesno("Confirmação", "Deseja realmente resetar a planilha? Todos os dados serão apagados."):
        cursor.execute("DELETE FROM chamados")
        conn.commit()
        atualizar_tabela()
        messagebox.showinfo("Resetado", "Todos os dados foram apagados.")

# Interface
root = tk.Tk()
root.title("Controle de Chamados - Suporte")
root.configure(bg=BG_COLOR)

# Definir tamanho fixo e centralizar a janela
largura_janela = 450
altura_janela = 500
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()
x = (largura_tela // 2) - (largura_janela // 2)
y = (altura_tela // 2) - (altura_janela // 2)
root.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")
root.resizable(False, False)

frame_form = tk.Frame(root, bg=BG_COLOR)
frame_form.pack(pady=10)

tk.Label(frame_form, text="Número do Chamado:", fg=TEXT_COLOR, bg=BG_COLOR).grid(row=0, column=0)
tk.Label(frame_form, text="Tipo do Item:", fg=TEXT_COLOR, bg=BG_COLOR).grid(row=1, column=0)
tk.Label(frame_form, text="Quantidade:", fg=TEXT_COLOR, bg=BG_COLOR).grid(row=2, column=0)

entry_numero = tk.Entry(frame_form, bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=ENTRY_FG, width=28)
tipo_var = tk.StringVar()

# Estilização do Combobox para ficar com fundo escuro e largura igual
style = ttk.Style()
style.theme_use("default")
style.configure("CustomCombobox.TCombobox",
                fieldbackground=ENTRY_BG,
                background=ENTRY_BG,
                foreground=ENTRY_FG,
                arrowcolor=ENTRY_FG)

entry_tipo = ttk.Combobox(frame_form, textvariable=tipo_var, values=tipos_itens,
                          state="readonly", style="CustomCombobox.TCombobox", width=28)

entry_quantidade = tk.Entry(frame_form, bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=ENTRY_FG, width=28)

# Definir valores iniciais
entry_numero.insert(0, "SC-")
entry_quantidade.insert(0, "1")

entry_numero.grid(row=0, column=1, padx=10, pady=8)
entry_tipo.grid(row=1, column=1, padx=10, pady=8)
entry_quantidade.grid(row=2, column=1, padx=10, pady=8)

# Botão centralizado e com largura igual aos campos
btn_adicionar = tk.Button(frame_form, text="Adicionar Item", command=adicionar_item, width=27)
btn_adicionar.grid(row=3, column=0, columnspan=2, pady=12)

# Treeview dinâmica
tree = ttk.Treeview(root, show='headings', height=12)  # Ajusta a altura da tabela
tree.pack(padx=10, pady=10, fill='both', expand=True)

style.configure("Treeview", background=TABLE_BG, foreground=TEXT_COLOR, fieldbackground=TABLE_BG)
style.configure("Treeview.Heading", background=HEADER_BG, foreground=HEADER_FG)
style.map('Treeview', background=[('selected', '#444')])

frame_buttons = tk.Frame(root, bg=BG_COLOR)
frame_buttons.pack(pady=5)

tk.Button(frame_buttons, text="Exportar CSV", command=exportar_csv).pack(side=tk.LEFT, padx=5)
tk.Button(frame_buttons, text="Resetar Planilha", command=resetar_planilha).pack(side=tk.LEFT, padx=5)

atualizar_tabela()
root.mainloop()

conn.close()

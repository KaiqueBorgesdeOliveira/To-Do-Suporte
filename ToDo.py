import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import csv
from collections import defaultdict
from datetime import datetime

# Conexão com banco de dados
conn = sqlite3.connect("chamados.db")
cursor = conn.cursor()
cursor.execute('''
    CREATE TABLE IF NOT EXISTS chamados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        numero_chamado TEXT NOT NULL,
        tipo_item TEXT NOT NULL,
        quantidade INTEGER NOT NULL,
        data_solicitacao TEXT NOT NULL
    )
''')

# Adicionar coluna de data se não existir (para compatibilidade com banco existente)
try:
    cursor.execute("ALTER TABLE chamados ADD COLUMN data_solicitacao TEXT")
    conn.commit()
except sqlite3.OperationalError:
    # Coluna já existe, não faz nada
    pass
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
    data = entry_data.get()

    if not (numero and tipo and quantidade and data):
        messagebox.showwarning("Atenção", "Preencha todos os campos.")
        return

    try:
        quantidade = int(quantidade)
    except ValueError:
        messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")
        return

    # Lógica especial para pilhas AAA - sempre quantidade 2
    if "Pilhas AAA" in tipo:
        quantidade = 2
        entry_quantidade.delete(0, tk.END)
        entry_quantidade.insert(0, "2")
        messagebox.showinfo("Informação", "Para pilhas AAA, a quantidade foi ajustada para 2 (padrão da empresa).")

    cursor.execute("INSERT INTO chamados (numero_chamado, tipo_item, quantidade, data_solicitacao) VALUES (?, ?, ?, ?)",
                   (numero, tipo, quantidade, data))
    conn.commit()
    atualizar_tabela()
    entry_numero.delete(0, tk.END)
    entry_numero.insert(0, "SC-")
    tipo_var.set("")
    entry_quantidade.delete(0, tk.END)
    entry_quantidade.insert(0, "1")
    # Atualizar data para hoje
    entry_data.delete(0, tk.END)
    entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

def atualizar_tabela():
    for item in tree.get_children():
        tree.delete(item)

    cursor.execute("SELECT numero_chamado, tipo_item, data_solicitacao FROM chamados ORDER BY data_solicitacao DESC, numero_chamado")
    chamados = cursor.fetchall()

    agrupado = defaultdict(list)
    for numero, tipo, data in chamados:
        agrupado[tipo].append(f"{numero} ({data})")

    tree["columns"] = list(agrupado.keys()) + ["Totais"]
    for col in agrupado:
        tree.heading(col, text=col)
        tree.column(col, width=150)

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

def selecionar_colunas_exportar():
    """Abre janela para selecionar colunas para exportação"""
    # Verificar se há dados para exportar
    cursor.execute("SELECT COUNT(*) FROM chamados")
    if cursor.fetchone()[0] == 0:
        messagebox.showwarning("Atenção", "Não há dados para exportar.")
        return
    
    janela_selecao = tk.Toplevel(root)
    janela_selecao.title("Selecionar Colunas para Exportação")
    janela_selecao.configure(bg=BG_COLOR)
    janela_selecao.geometry("450x400")
    janela_selecao.resizable(False, False)
    janela_selecao.transient(root)
    janela_selecao.grab_set()
    
    # Centralizar janela
    janela_selecao.update_idletasks()
    x = (janela_selecao.winfo_screenwidth() // 2) - (450 // 2)
    y = (janela_selecao.winfo_screenheight() // 2) - (400 // 2)
    janela_selecao.geometry(f"450x400+{x}+{y}")
    
    frame_selecao = tk.Frame(janela_selecao, bg=BG_COLOR)
    frame_selecao.pack(pady=20)
    
    tk.Label(frame_selecao, text="Selecione as colunas para exportar:", 
             fg=TEXT_COLOR, bg=BG_COLOR, font=("Arial", 10, "bold")).pack(pady=10)
    
    # Opções de exportação
    tk.Label(frame_selecao, text="Opção 1: Colunas Padrão (Lista simples)", 
             fg=TEXT_COLOR, bg=BG_COLOR, font=("Arial", 9, "bold")).pack(anchor="w", padx=20, pady=(10,5))
    
    colunas_padrao = ["Número do Chamado", "Tipo do Item", "Quantidade", "Data da Solicitação"]
    colunas_selecionadas = {}
    
    for coluna in colunas_padrao:
        var = tk.BooleanVar(value=True)
        colunas_selecionadas[coluna] = var
        cb = tk.Checkbutton(frame_selecao, text=coluna, variable=var, 
                           fg=TEXT_COLOR, bg=BG_COLOR, selectcolor=ENTRY_BG)
        cb.pack(anchor="w", padx=40, pady=1)
    
    # Separador
    tk.Label(frame_selecao, text="", bg=BG_COLOR).pack(pady=5)
    
    tk.Label(frame_selecao, text="Opção 2: Por Tipo de Item (Formato da tabela)", 
             fg=TEXT_COLOR, bg=BG_COLOR, font=("Arial", 9, "bold")).pack(anchor="w", padx=20, pady=(10,5))
    
    # Obter tipos de itens únicos
    cursor.execute("SELECT DISTINCT tipo_item FROM chamados ORDER BY tipo_item")
    tipos_itens = [row[0] for row in cursor.fetchall()]
    
    for tipo in tipos_itens:
        var = tk.BooleanVar(value=False)
        colunas_selecionadas[tipo] = var
        cb = tk.Checkbutton(frame_selecao, text=tipo, variable=var, 
                           fg=TEXT_COLOR, bg=BG_COLOR, selectcolor=ENTRY_BG)
        cb.pack(anchor="w", padx=40, pady=1)
    
    def exportar_selecionado():
        # Filtrar colunas selecionadas
        colunas_para_exportar = [col for col, var in colunas_selecionadas.items() if var.get()]
        
        if not colunas_para_exportar:
            messagebox.showwarning("Atenção", "Selecione pelo menos uma coluna para exportar.")
            return
        
        janela_selecao.destroy()
        exportar_csv_com_colunas(colunas_para_exportar)
    
    # Botões
    frame_botoes = tk.Frame(frame_selecao, bg=BG_COLOR)
    frame_botoes.pack(pady=20)
    
    tk.Button(frame_botoes, text="Exportar", command=exportar_selecionado, width=12).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes, text="Cancelar", command=janela_selecao.destroy, width=12).pack(side=tk.LEFT, padx=5)

def exportar_csv_com_colunas(colunas_selecionadas):
    """Exporta CSV com as colunas selecionadas"""
    file = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV files", "*.csv")])
    if not file:
        return
    
    # Obter dados do banco de dados com todas as informações
    cursor.execute("SELECT numero_chamado, tipo_item, quantidade, data_solicitacao FROM chamados ORDER BY data_solicitacao DESC, numero_chamado")
    chamados = cursor.fetchall()
    
    # Preparar dados para exportação
    dados = []
    
    # Verificar se são colunas de tipos de itens ou colunas padrão
    colunas_padrao = ["Número do Chamado", "Tipo do Item", "Quantidade", "Data da Solicitação"]
    sao_colunas_padrao = any(col in colunas_padrao for col in colunas_selecionadas)
    
    if sao_colunas_padrao:
        # Exportação com colunas padrão
        cabecalho = colunas_selecionadas
        dados.append(cabecalho)
        
        for chamado in chamados:
            linha = []
            for coluna in colunas_selecionadas:
                if coluna == "Número do Chamado":
                    linha.append(chamado[0])
                elif coluna == "Tipo do Item":
                    linha.append(chamado[1])
                elif coluna == "Quantidade":
                    linha.append(chamado[2])
                elif coluna == "Data da Solicitação":
                    linha.append(chamado[3])
            dados.append(linha)
    else:
        # Exportação com colunas de tipos de itens (formato da tabela)
        cabecalho = colunas_selecionadas + ["Totais"]
        dados.append(cabecalho)
        
        # Agrupar dados por tipo de item
        agrupado = defaultdict(list)
        for numero, tipo, quantidade, data in chamados:
            agrupado[tipo].append(f"{numero} ({data})")
        
        # Criar linhas da tabela
        max_len = max(len(v) for v in agrupado.values()) if agrupado else 0
        for i in range(max_len):
            row = [agrupado[col][i] if i < len(agrupado[col]) else "" for col in colunas_selecionadas]
            row.append("")
            dados.append(row)
        
        # Adicionar linha de totais
        totais = [str(len(agrupado[col])) for col in colunas_selecionadas]
        totais.append("Total")
        dados.append(totais)
    
    # Exporta para CSV com codificação UTF-8-SIG para melhor compatibilidade com Excel
    with open(file, mode='w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')  # Usando ponto e vírgula como separador
        writer.writerows(dados)
    
    messagebox.showinfo("Exportado", f"Dados exportados para {file}")

def exportar_csv():
    """Função original mantida para compatibilidade"""
    selecionar_colunas_exportar()

def editar_item(event=None):
    """Abre janela para editar item selecionado"""
    selection = tree.selection()
    if not selection:
        messagebox.showwarning("Atenção", "Selecione um item para editar.")
        return
    
    item = tree.item(selection[0])
    values = item['values']
    
    # Verifica se não é linha de totais
    if values and values[-1] == "Total":
        messagebox.showwarning("Atenção", "Não é possível editar a linha de totais.")
        return
    
    # Encontra o item no banco de dados
    cursor.execute("SELECT id, numero_chamado, tipo_item, quantidade, data_solicitacao FROM chamados")
    todos_chamados = cursor.fetchall()
    
    # Procura o chamado correspondente aos valores da tabela
    chamado_para_editar = None
    for chamado in todos_chamados:
        # Extrai o número do chamado do formato "numero (data)" na tabela
        numero_tabela = chamado[1]  # numero_chamado
        for value in values:
            if value and numero_tabela in value:
                chamado_para_editar = chamado
                break
        if chamado_para_editar:
            break
    
    if not chamado_para_editar:
        messagebox.showerror("Erro", "Item não encontrado no banco de dados.")
        return
    
    abrir_janela_edicao(chamado_para_editar)

def abrir_janela_edicao(chamado):
    """Abre janela modal para editar um chamado"""
    janela_edicao = tk.Toplevel(root)
    janela_edicao.title("Editar Item")
    janela_edicao.configure(bg=BG_COLOR)
    janela_edicao.geometry("400x350")
    janela_edicao.resizable(False, False)
    janela_edicao.transient(root)
    janela_edicao.grab_set()
    
    # Centralizar janela
    janela_edicao.update_idletasks()
    x = (janela_edicao.winfo_screenwidth() // 2) - (400 // 2)
    y = (janela_edicao.winfo_screenheight() // 2) - (350 // 2)
    janela_edicao.geometry(f"400x350+{x}+{y}")
    
    frame_edicao = tk.Frame(janela_edicao, bg=BG_COLOR)
    frame_edicao.pack(pady=20)
    
    # Campos de edição
    tk.Label(frame_edicao, text="Número do Chamado:", fg=TEXT_COLOR, bg=BG_COLOR).grid(row=0, column=0, sticky="w", pady=5)
    entry_numero_edit = tk.Entry(frame_edicao, bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=ENTRY_FG, width=30)
    entry_numero_edit.grid(row=0, column=1, padx=10, pady=5)
    entry_numero_edit.insert(0, chamado[1])
    
    tk.Label(frame_edicao, text="Tipo do Item:", fg=TEXT_COLOR, bg=BG_COLOR).grid(row=1, column=0, sticky="w", pady=5)
    tipo_var_edit = tk.StringVar()
    entry_tipo_edit = ttk.Combobox(frame_edicao, textvariable=tipo_var_edit, values=tipos_itens,
                                  state="readonly", style="CustomCombobox.TCombobox", width=28)
    entry_tipo_edit.grid(row=1, column=1, padx=10, pady=5)
    entry_tipo_edit.set(chamado[2])
    
    tk.Label(frame_edicao, text="Quantidade:", fg=TEXT_COLOR, bg=BG_COLOR).grid(row=2, column=0, sticky="w", pady=5)
    entry_quantidade_edit = tk.Entry(frame_edicao, bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=ENTRY_FG, width=30)
    entry_quantidade_edit.grid(row=2, column=1, padx=10, pady=5)
    entry_quantidade_edit.insert(0, str(chamado[3]))
    
    tk.Label(frame_edicao, text="Data da Solicitação:", fg=TEXT_COLOR, bg=BG_COLOR).grid(row=3, column=0, sticky="w", pady=5)
    entry_data_edit = tk.Entry(frame_edicao, bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=ENTRY_FG, width=30)
    entry_data_edit.grid(row=3, column=1, padx=10, pady=5)
    entry_data_edit.insert(0, chamado[4])
    
    def salvar_edicao():
        numero = entry_numero_edit.get()
        tipo = tipo_var_edit.get()
        quantidade = entry_quantidade_edit.get()
        data = entry_data_edit.get()
        
        if not (numero and tipo and quantidade and data):
            messagebox.showwarning("Atenção", "Preencha todos os campos.")
            return
        
        try:
            quantidade = int(quantidade)
        except ValueError:
            messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")
            return
        
        # Lógica especial para pilhas AAA - sempre quantidade 2
        if "Pilhas AAA" in tipo:
            quantidade = 2
            entry_quantidade_edit.delete(0, tk.END)
            entry_quantidade_edit.insert(0, "2")
            messagebox.showinfo("Informação", "Para pilhas AAA, a quantidade foi ajustada para 2 (padrão da empresa).")
        
        cursor.execute("UPDATE chamados SET numero_chamado=?, tipo_item=?, quantidade=?, data_solicitacao=? WHERE id=?",
                      (numero, tipo, quantidade, data, chamado[0]))
        conn.commit()
        atualizar_tabela()
        janela_edicao.destroy()
        messagebox.showinfo("Sucesso", "Item editado com sucesso!")
    
    def excluir_item():
        if messagebox.askyesno("Confirmação", "Deseja realmente excluir este item?"):
            cursor.execute("DELETE FROM chamados WHERE id=?", (chamado[0],))
            conn.commit()
            atualizar_tabela()
            janela_edicao.destroy()
            messagebox.showinfo("Sucesso", "Item excluído com sucesso!")
    
    # Botões
    frame_botoes = tk.Frame(frame_edicao, bg=BG_COLOR)
    frame_botoes.grid(row=4, column=0, columnspan=2, pady=20)
    
    tk.Button(frame_botoes, text="Salvar", command=salvar_edicao, width=12).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes, text="Excluir", command=excluir_item, width=12).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes, text="Cancelar", command=janela_edicao.destroy, width=12).pack(side=tk.LEFT, padx=5)

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
altura_janela = 550
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
tk.Label(frame_form, text="Data da Solicitação:", fg=TEXT_COLOR, bg=BG_COLOR).grid(row=3, column=0)

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
entry_data = tk.Entry(frame_form, bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=ENTRY_FG, width=28)

# Definir valores iniciais
entry_numero.insert(0, "SC-")
entry_quantidade.insert(0, "1")
entry_data.insert(0, datetime.now().strftime("%d/%m/%Y"))

entry_numero.grid(row=0, column=1, padx=10, pady=8)
entry_tipo.grid(row=1, column=1, padx=10, pady=8)
entry_quantidade.grid(row=2, column=1, padx=10, pady=8)
entry_data.grid(row=3, column=1, padx=10, pady=8)

# Botão centralizado e com largura igual aos campos
btn_adicionar = tk.Button(frame_form, text="Adicionar Item", command=adicionar_item, width=27)
btn_adicionar.grid(row=4, column=0, columnspan=2, pady=12)

# Treeview dinâmica
tree = ttk.Treeview(root, show='headings', height=12)  # Ajusta a altura da tabela
tree.pack(padx=10, pady=10, fill='both', expand=True)

# Configurar evento de duplo clique para editar
tree.bind('<Double-1>', editar_item)

style.configure("Treeview", background=TABLE_BG, foreground=TEXT_COLOR, fieldbackground=TABLE_BG)
style.configure("Treeview.Heading", background=HEADER_BG, foreground=HEADER_FG)
style.map('Treeview', background=[('selected', '#444')])

frame_buttons = tk.Frame(root, bg=BG_COLOR)
frame_buttons.pack(pady=5)

tk.Button(frame_buttons, text="Editar Item", command=editar_item).pack(side=tk.LEFT, padx=5)
tk.Button(frame_buttons, text="Exportar CSV", command=exportar_csv).pack(side=tk.LEFT, padx=5)
tk.Button(frame_buttons, text="Resetar Planilha", command=resetar_planilha).pack(side=tk.LEFT, padx=5)

atualizar_tabela()
root.mainloop()

conn.close()

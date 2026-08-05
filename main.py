import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from docx import Document

# =========================================================
# DADOS GERAIS
# =========================================================

data_hoje = datetime.now()
data_ajustada = data_hoje + timedelta(hours=1)
data_formatada = data_ajustada.strftime("%d/%m/%Y - %H:00")

categorias = {
    "HONRA E LOUVOR": [],
    "INTENÇÃO": [],
    "ANIVERSÁRIO NATALÍCIO": [],
    "RECUPERAÇÃO DE SAÚDE": [],
    "ANIVERSÁRIO DE CASAMENTO": [],
    "AÇÃO DE GRAÇAS": [],
    "IRMÃOS FALECIDOS": [],
    "7º DIA": []
}

pix = []
valor_total = 0

# =========================================================
# IMPORTAR DOCX
# =========================================================

def extrair_docx():
    def normalizar(txt):
        return " ".join(txt.strip().upper().split())

    try:
        doc = Document("intencoes.docx")
    except:
        return

    TITULOS = list(categorias.keys())
    TITULOS_NORM = [normalizar(t) for t in TITULOS]

    titulo_atual = None

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                linha = cell.text.strip()
                if not linha:
                    continue

                ln = normalizar(linha)

                if ln in TITULOS_NORM:
                    titulo_atual = ln
                    continue

                if titulo_atual and linha != "+" and not (linha.startswith("INTENÇÕES")):
                    categorias[titulo_atual].append(linha)


extrair_docx()

# =========================================================
# IMPRIMIR EXCEL
# =========================================================

def imprimir_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Intenções"

    ws.append([f"INTENÇÕES DA MISSA {data_formatada}"])
    ws.append([])

    for titulo, lista in categorias.items():
        ws.append([titulo])
        for item in lista:
            ws.append([item])
        ws.append([])

    borda = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for row in ws.iter_rows():
        for cell in row:
            cell.border = borda

    for col in ws.columns:
        tamanho = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = tamanho + 5

    wb.save("Intenções.xlsx")
    messagebox.showinfo("Sucesso", "Arquivo Intenções.xlsx criado!")

# =========================================================
# JANELA MARCAR INTENÇÃO
# =========================================================

def janela_marcar():
    global valor_total

    win = tk.Toplevel(root)
    win.title("Marcar Intenção")
    win.geometry("400x300")

    tk.Label(win, text="Categoria").pack(pady=5)

    categoria = tk.StringVar(value="HONRA E LOUVOR")
    tk.OptionMenu(win, categoria, *categorias.keys()).pack()

    tk.Label(win, text="Intenção").pack(pady=5)
    entrada = tk.Entry(win, width=40)
    entrada.pack()

    def salvar():
        global valor_total
        texto = entrada.get().strip()
        if texto == "":
            messagebox.showerror("Erro", "Digite uma intenção")
            return

        categorias[categoria.get()].append(texto.title())
        valor_total += 3

        messagebox.showinfo("Sucesso", "Intenção adicionada!")
        entrada.delete(0, tk.END)

    tk.Button(win, text="Salvar", command=salvar).pack(pady=10)

# =========================================================
# JANELA EDITAR
# =========================================================

def janela_editar():
    global valor_total

    win = tk.Toplevel(root)
    win.title("Editar Intenções")
    win.geometry("500x400")

    categoria = tk.StringVar(value="HONRA E LOUVOR")
    tk.OptionMenu(win, categoria, *categorias.keys()).pack(pady=5)

    lista = tk.Listbox(win, width=60)
    lista.pack(pady=10)

    def atualizar():
        lista.delete(0, tk.END)
        for item in categorias[categoria.get()]:
            lista.insert(tk.END, item)

    def remover():
        global valor_total
        try:
            idx = lista.curselection()[0]
            categorias[categoria.get()].pop(idx)
            valor_total -= 3
            atualizar()
        except:
            messagebox.showerror("Erro", "Selecione um item")

    def editar():
        try:
            idx = lista.curselection()[0]
        except:
            messagebox.showerror("Erro", "Selecione um item")
            return

        nova = tk.simpledialog.askstring("Editar", "Nova intenção:")
        if nova:
            categorias[categoria.get()][idx] = nova.title()
            atualizar()

    tk.Button(win, text="Atualizar Lista", command=atualizar).pack()
    tk.Button(win, text="Editar", command=editar).pack(pady=5)
    tk.Button(win, text="Remover", command=remover).pack()

# =========================================================
# DADOS
# =========================================================

def mostrar_dados():
    total_int = valor_total / 3
    messagebox.showinfo(
        "Dados",
        f"Total de intenções: {total_int}\nValor total: R$ {valor_total:.2f}"
    )

# =========================================================
# JANELA PRINCIPAL
# =========================================================

root = tk.Tk()
root.title("Sistema de Intenções da Missa")
root.geometry("400x350")

tk.Button(root, text="Marcar Intenção", width=30, command=janela_marcar).pack(pady=5)
tk.Button(root, text="Editar Intenções", width=30, command=janela_editar).pack(pady=5)
tk.Button(root, text="Dados", width=30, command=mostrar_dados).pack(pady=5)
tk.Button(root, text="Imprimir Excel", width=30, command=imprimir_excel).pack(pady=5)
tk.Button(root, text="Sair", width=30, command=root.destroy).pack(pady=20)

root.mainloop()

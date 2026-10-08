
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from comparar_planilhas import comparar_planilhas


COR_FUNDO = "#F4F6F9"
COR_BRANCO = "#FFFFFF"
COR_AZUL = "#1F4E78"
COR_AZUL_ESCURO = "#163A5A"
COR_TEXTO = "#374151"
COR_CINZA = "#E2E8F0"


# Local onde o relatório será salvo
if getattr(sys, "frozen", False):
    PASTA_APLICATIVO = os.path.dirname(sys.executable)
else:
    PASTA_APLICATIVO = os.path.dirname(os.path.abspath(__file__))

CAMINHO_RELATORIO = os.path.join(
    PASTA_APLICATIVO,
    "relatorio_comparacao.xlsx"
)


caminho_a = ""
caminho_b = ""


def limpar_resumo():
    for valor in valores_resumo.values():
        valor.config(text="—")

    label_status.config(
        text="Selecione as planilhas e clique em Comparar.",
        fg=COR_TEXTO
    )


def carregar_colunas(arquivo):
    """
    Lê apenas o cabeçalho do Excel.
    Não precisa carregar milhares de registros
    só para descobrir os nomes das colunas.
    """
    dados = pd.read_excel(arquivo, nrows=0)
    colunas = [str(coluna).strip() for coluna in dados.columns]

    if len(colunas) != len(set(colunas)):
        raise ValueError("A planilha possui nomes de colunas repetidos.")

    return colunas


def escolher_coluna_inicial(colunas):
    if "ID" in colunas:
        return "ID"

    if colunas:
        return colunas[0]

    return ""


def selecionar_planilha_a():
    global caminho_a

    arquivo = filedialog.askopenfilename(
        title="Selecione a Planilha A",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )

    if not arquivo:
        return

    try:
        colunas = carregar_colunas(arquivo)

        if not colunas:
            raise ValueError("A planilha não possui colunas.")

        caminho_a = arquivo

        label_arquivo_a.config(
            text=os.path.basename(arquivo)
        )

        seletor_coluna_a["values"] = colunas
        coluna_a_var.set(escolher_coluna_inicial(colunas))

        limpar_resumo()

    except Exception as erro:
        messagebox.showerror(
            "Erro na Planilha A",
            f"Não foi possível ler as colunas.\n\n{erro}"
        )


def selecionar_planilha_b():
    global caminho_b

    arquivo = filedialog.askopenfilename(
        title="Selecione a Planilha B",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )

    if not arquivo:
        return

    try:
        colunas = carregar_colunas(arquivo)

        if not colunas:
            raise ValueError("A planilha não possui colunas.")

        caminho_b = arquivo

        label_arquivo_b.config(
            text=os.path.basename(arquivo)
        )

        seletor_coluna_b["values"] = colunas
        coluna_b_var.set(escolher_coluna_inicial(colunas))

        limpar_resumo()

    except Exception as erro:
        messagebox.showerror(
            "Erro na Planilha B",
            f"Não foi possível ler as colunas.\n\n{erro}"
        )


def executar_comparacao():
    if not caminho_a or not caminho_b:
        messagebox.showwarning(
            "Atenção",
            "Selecione as duas planilhas primeiro."
        )
        return

    coluna_a = coluna_a_var.get().strip()
    coluna_b = coluna_b_var.get().strip()

    if not coluna_a or not coluna_b:
        messagebox.showwarning(
            "Atenção",
            "Escolha a coluna identificadora de cada planilha."
        )
        return

    botao_comparar.config(
        state="disabled",
        text="Comparando..."
    )

    label_status.config(
        text="Comparando planilhas, aguarde...",
        fg=COR_AZUL
    )

    janela.update_idletasks()

    try:
        resultado = comparar_planilhas(
            arquivo_a=caminho_a,
            arquivo_b=caminho_b,
            caminho_relatorio=CAMINHO_RELATORIO,
            coluna_id_a=coluna_a,
            coluna_id_b=coluna_b
        )

        if resultado is not True:
            mensagem = (
                resultado
                if isinstance(resultado, str)
                else "Não foi possível concluir a comparação."
            )

            label_status.config(
                text="Erro durante a comparação.",
                fg="#DC2626"
            )

            messagebox.showerror(
                "Erro na comparação",
                mensagem
            )
            return

        resumo = pd.read_excel(
            CAMINHO_RELATORIO,
            sheet_name="Resumo"
        )

        quantidades = dict(
            zip(resumo["Categoria"], resumo["Quantidade"])
        )

        valores_resumo["iguais"].config(
            text=str(quantidades.get("Registros iguais", 0))
        )

        valores_resumo["diferentes"].config(
            text=str(quantidades.get("Registros diferentes", 0))
        )

        valores_resumo["so_a"].config(
            text=str(
                quantidades.get("Registros só na Planilha A", 0)
            )
        )

        valores_resumo["so_b"].config(
            text=str(
                quantidades.get("Registros só na Planilha B", 0)
            )
        )

        label_status.config(
            text="Comparação concluída com sucesso!",
            fg="#15803D"
        )

        messagebox.showinfo(
            "Comparação concluída",
            "Relatório Excel gerado com sucesso!"
        )

    except Exception as erro:
        label_status.config(
            text="Erro durante a comparação.",
            fg="#DC2626"
        )

        messagebox.showerror(
            "Erro inesperado",
            f"Não foi possível concluir a operação.\n\n{erro}"
        )

    finally:
        botao_comparar.config(
            state="normal",
            text="Comparar Planilhas"
        )


def abrir_relatorio():
    if not os.path.isfile(CAMINHO_RELATORIO):
        messagebox.showwarning(
            "Relatório não encontrado",
            "Faça uma comparação primeiro."
        )
        return

    try:
        os.startfile(CAMINHO_RELATORIO)

    except OSError as erro:
        messagebox.showerror(
            "Erro ao abrir relatório",
            str(erro)
        )


def criar_cartao(pai, titulo, cor, coluna):
    cartao = tk.Frame(
        pai,
        bg=COR_BRANCO,
        highlightbackground="#DDE3EB",
        highlightthickness=1
    )

    cartao.grid(
        row=0,
        column=coluna,
        padx=5,
        sticky="nsew"
    )

    tk.Label(
        cartao,
        text=titulo,
        font=("Arial", 9),
        bg=COR_BRANCO,
        fg=COR_TEXTO
    ).pack(pady=(12, 4))

    valor = tk.Label(
        cartao,
        text="—",
        font=("Arial", 25, "bold"),
        bg=COR_BRANCO,
        fg=cor
    )

    valor.pack(pady=(0, 12))

    return valor


# ---------------- JANELA ----------------

janela = tk.Tk()
janela.title("Comparador de Planilhas — Versão 2.0")
janela.geometry("760x550")
janela.configure(bg=COR_FUNDO)
janela.resizable(False, False)

coluna_a_var = tk.StringVar()
coluna_b_var = tk.StringVar()

conteudo = tk.Frame(
    janela,
    bg=COR_FUNDO
)

conteudo.pack(
    fill="both",
    expand=True,
    padx=30,
    pady=15
)


tk.Label(
    conteudo,
    text="Comparador de Planilhas",
    font=("Arial", 23, "bold"),
    bg=COR_FUNDO,
    fg=COR_AZUL
).pack(pady=(5, 3))


tk.Label(
    conteudo,
    text="Compare arquivos Excel de forma simples e rápida.",
    font=("Arial", 10),
    bg=COR_FUNDO,
    fg=COR_TEXTO
).pack(pady=(0, 16))


# ---------------- ARQUIVOS ----------------

area_arquivos = tk.Frame(
    conteudo,
    bg=COR_FUNDO
)

area_arquivos.pack(fill="x")

area_arquivos.columnconfigure(
    0, weight=1, uniform="arquivos"
)

area_arquivos.columnconfigure(
    1, weight=1, uniform="arquivos"
)


quadro_a = tk.Frame(
    area_arquivos,
    bg=COR_BRANCO,
    padx=12,
    pady=12,
    highlightbackground="#DDE3EB",
    highlightthickness=1
)

quadro_a.grid(
    row=0,
    column=0,
    padx=(0, 7),
    sticky="nsew"
)


quadro_b = tk.Frame(
    area_arquivos,
    bg=COR_BRANCO,
    padx=12,
    pady=12,
    highlightbackground="#DDE3EB",
    highlightthickness=1
)

quadro_b.grid(
    row=0,
    column=1,
    padx=(7, 0),
    sticky="nsew"
)


# Planilha A
botao_a = tk.Button(
    quadro_a,
    text="Selecionar Planilha A",
    command=selecionar_planilha_a,
    font=("Arial", 10, "bold"),
    bg=COR_CINZA,
    fg=COR_AZUL,
    relief="flat",
    cursor="hand2",
    height=2
)

botao_a.pack(fill="x")


label_arquivo_a = tk.Label(
    quadro_a,
    text="Nenhuma planilha selecionada",
    font=("Arial", 9),
    bg=COR_BRANCO,
    fg=COR_TEXTO,
    wraplength=280,
    height=2
)

label_arquivo_a.pack(pady=(6, 8))


tk.Label(
    quadro_a,
    text="Coluna identificadora:",
    font=("Arial", 9, "bold"),
    bg=COR_BRANCO,
    fg=COR_TEXTO,
    anchor="w"
).pack(fill="x")


seletor_coluna_a = ttk.Combobox(
    quadro_a,
    textvariable=coluna_a_var,
    state="readonly",
    font=("Arial", 10)
)

seletor_coluna_a.pack(
    fill="x",
    pady=(5, 0)
)


# Planilha B
botao_b = tk.Button(
    quadro_b,
    text="Selecionar Planilha B",
    command=selecionar_planilha_b,
    font=("Arial", 10, "bold"),
    bg=COR_CINZA,
    fg=COR_AZUL,
    relief="flat",
    cursor="hand2",
    height=2
)

botao_b.pack(fill="x")


label_arquivo_b = tk.Label(
    quadro_b,
    text="Nenhuma planilha selecionada",
    font=("Arial", 9),
    bg=COR_BRANCO,
    fg=COR_TEXTO,
    wraplength=280,
    height=2
)

label_arquivo_b.pack(pady=(6, 8))


tk.Label(
    quadro_b,
    text="Coluna identificadora:",
    font=("Arial", 9, "bold"),
    bg=COR_BRANCO,
    fg=COR_TEXTO,
    anchor="w"
).pack(fill="x")


seletor_coluna_b = ttk.Combobox(
    quadro_b,
    textvariable=coluna_b_var,
    state="readonly",
    font=("Arial", 10)
)

seletor_coluna_b.pack(
    fill="x",
    pady=(5, 0)
)


# ---------------- BOTÕES ----------------

area_botoes = tk.Frame(
    conteudo,
    bg=COR_FUNDO
)

area_botoes.pack(
    fill="x",
    pady=(16, 10)
)

area_botoes.columnconfigure(
    0, weight=1, uniform="botoes"
)

area_botoes.columnconfigure(
    1, weight=1, uniform="botoes"
)


botao_comparar = tk.Button(
    area_botoes,
    text="Comparar Planilhas",
    command=executar_comparacao,
    font=("Arial", 11, "bold"),
    bg=COR_AZUL,
    fg=COR_BRANCO,
    activebackground=COR_AZUL_ESCURO,
    activeforeground=COR_BRANCO,
    relief="flat",
    cursor="hand2",
    height=2
)

botao_comparar.grid(
    row=0,
    column=0,
    padx=(0, 7),
    sticky="ew"
)


botao_relatorio = tk.Button(
    area_botoes,
    text="Abrir Relatório Excel",
    command=abrir_relatorio,
    font=("Arial", 10, "bold"),
    bg=COR_CINZA,
    fg=COR_AZUL,
    relief="flat",
    cursor="hand2",
    height=2
)

botao_relatorio.grid(
    row=0,
    column=1,
    padx=(7, 0),
    sticky="ew"
)


# ---------------- RESUMO ----------------

tk.Label(
    conteudo,
    text="Resumo da Comparação",
    font=("Arial", 15, "bold"),
    bg=COR_FUNDO,
    fg=COR_AZUL
).pack(pady=(12, 12))


area_resumo = tk.Frame(
    conteudo,
    bg=COR_FUNDO
)

area_resumo.pack(fill="x")

for coluna in range(4):
    area_resumo.columnconfigure(
        coluna,
        weight=1,
        uniform="cartoes"
    )


valores_resumo = {
    "iguais": criar_cartao(
        area_resumo,
        "Registros iguais",
        "#15803D",
        0
    ),
    "diferentes": criar_cartao(
        area_resumo,
        "Registros diferentes",
        "#DC2626",
        1
    ),
    "so_a": criar_cartao(
        area_resumo,
        "Somente na A",
        COR_AZUL,
        2
    ),
    "so_b": criar_cartao(
        area_resumo,
        "Somente na B",
        COR_AZUL,
        3
    )
}


label_status = tk.Label(
    conteudo,
    text="Selecione as planilhas e clique em Comparar.",
    font=("Arial", 9),
    bg=COR_FUNDO,
    fg=COR_TEXTO
)

label_status.pack(pady=(18, 0))


janela.mainloop()

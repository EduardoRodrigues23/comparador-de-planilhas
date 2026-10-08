print("================================")
print("     COMPARADOR DE PLANILHAS")
print("================================")
print()

import pandas as pd
import numbers
import os
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def comparar_planilhas():
    arquivo_a = input("Digite o nome da Planilha A: ").strip()
    arquivo_b = input("Digite o nome da Planilha B: ").strip()

 
    if not arquivo_a.endswith(".xlsx"):
        arquivo_a += ".xlsx"

    if not arquivo_b.endswith(".xlsx"):
        arquivo_b += ".xlsx"
    print()

    if not os.path.exists(arquivo_a):
        print("Arquivo não encontrado:", arquivo_a)
        exit()

    if not os.path.exists(arquivo_b):
        print("Arquivo não encontrado:", arquivo_b)
        exit()

    try:
        planilha_a = pd.read_excel(arquivo_a, dtype={"ID": str})
    except Exception:
        print("Erro: não foi possível abrir a Planilha A.")
        print("Verifique se o arquivo é um Excel válido.")
        exit()
    try:
        planilha_b = pd.read_excel(arquivo_b, dtype={"ID": str})
    except Exception:
        print()
        print("Erro: não foi possível abrir a Planilha B.")
        print("Verifique se o arquivo é um Excel válido.")
        exit()

    planilha_a.columns = planilha_a.columns.str.strip()
    planilha_b.columns = planilha_b.columns.str.strip()

    planilha_a["ID"] = planilha_a["ID"].str.strip()
    planilha_b["ID"] = planilha_b["ID"].str.strip()

    if "ID" not in planilha_a.columns:
        print("Erro: a Planilha A não possui a coluna 'ID'.")
        exit()

    if "ID" not in planilha_b.columns:
        print("Erro: a Planilha B não possui a coluna 'ID'.")
        exit()

    ids_vazios_a = planilha_a["ID"].isna() | (planilha_a["ID"].str.strip() == "")
    ids_vazios_b = planilha_b["ID"].isna() | (planilha_b["ID"].str.strip() == "")

    if ids_vazios_a.any() or ids_vazios_b.any():
        print()
        print("ERRO: Foram encontrados IDs vazios.")

        if ids_vazios_a.any():
            print("ID vazio na Planilha A.")

        if ids_vazios_b.any():
            print("ID vazio na Planilha B.")

        print("A comparação não pode continuar.")
        exit()


    colunas_comuns = planilha_a.columns.intersection(planilha_b.columns)

    colunas_so_na_a = planilha_a.columns.difference(planilha_b.columns)
    colunas_so_na_b = planilha_b.columns.difference(planilha_a.columns)

    if len(colunas_so_na_a) > 0 or len(colunas_so_na_b) > 0:
        print()
        print("Aviso: existem colunas diferentes entre as planilhas.")

        if len(colunas_so_na_a) > 0:
            print("Colunas somente na Planilha A:", list(colunas_so_na_a))

        if len(colunas_so_na_b) > 0:
            print("Colunas somente na Planilha B:", list(colunas_so_na_b))

    ids_duplicados_a = planilha_a[planilha_a["ID"].duplicated()]["ID"].unique()
    ids_duplicados_b = planilha_b[planilha_b["ID"].duplicated()]["ID"].unique()

    if len(ids_duplicados_a) > 0 or len(ids_duplicados_b) > 0:
        print("ERRO: Foram encontrados IDs duplicados.")

        if len(ids_duplicados_a) > 0:
            print("IDs duplicados na Planilha A:", ids_duplicados_a)

        if len(ids_duplicados_b) > 0:
            print("IDs duplicados na Planilha B:", ids_duplicados_b)

        exit()


    ids_em_comum = planilha_a["ID"].isin(planilha_b["ID"])

    ids_comuns = planilha_a.loc[ids_em_comum, "ID"]

    iguais = 0
    diferentes = 0
    ids_iguais = []
    ids_diferentes = []
    ids_so_na_a = []
    ids_so_na_b = []
    detalhes_diferencas = []

    for id in ids_comuns:
        linha_a = planilha_a[planilha_a["ID"] == id][colunas_comuns].reset_index(drop=True)
        linha_b = planilha_b[planilha_b["ID"] == id][colunas_comuns].reset_index(drop=True)

        sao_iguais = linha_a.equals(linha_b)

        if sao_iguais:
            iguais += 1
            ids_iguais.append(id)

        else:
            diferentes += 1
            ids_diferentes.append(id)

            diferencas = linha_a != linha_b
            campos_diferentes = diferencas.columns[diferencas.any()]

            for campo in campos_diferentes:
                valor_a = linha_a.iloc[0][campo]
                valor_b = linha_b.iloc[0][campo]

                if pd.isna(valor_a) or pd.isna(valor_b):
                    diferenca = "Alterado"

                elif isinstance(valor_a, numbers.Number) and isinstance(valor_b, numbers.Number):
                    diferenca = valor_b - valor_a

                else:
                    diferenca = "Alterado"

                detalhes_diferencas.append({
                    "ID": id,
                    "Campo": campo,
                    "Planilha A": valor_a,
                    "Planilha B": valor_b,
                    "Diferença": diferenca
                })

    ids_em_comum_b = planilha_b["ID"].isin(planilha_a["ID"])

    ids_so_na_a = planilha_a.loc[~ids_em_comum, "ID"].tolist()
    ids_so_na_b = planilha_b.loc[~ids_em_comum_b, "ID"].tolist()

    registros_so_na_a = planilha_a[~ids_em_comum]
    registros_so_na_b = planilha_b[~ids_em_comum_b]


    print("===== RELATÓRIO DE COMPARAÇÃO =====")

    print()
    print("Registros iguais:", iguais)
    print("Registros diferentes:", diferentes)
    print("Registros só na Planilha A:", len(ids_so_na_a))
    print("Registros só na Planilha B:", len(ids_so_na_b))

    print()
    print("----- DIFERENÇAS -----")
    print()

    for detalhe in detalhes_diferencas:
        print("ID:", detalhe["ID"])
        print("Campo:", detalhe["Campo"])
        print("Planilha A:", detalhe["Planilha A"])
        print("Planilha B:", detalhe["Planilha B"])
        print("Diferença:", detalhe["Diferença"])

    print()
    print("----- SÓ NA PLANILHA A -----")

    if len(ids_so_na_a) == 0:
        print("Nenhum registro encontrado.")
    else:
        for _, registro in registros_so_na_a.iterrows():
            for campo in planilha_a.columns:
                print(campo + ":", registro[campo])
            

    print()
    print("----- SÓ NA PLANILHA B -----")

    if len(ids_so_na_b) == 0:
        print("Nenhum registro encontrado.")
    else:
        for _, registro in registros_so_na_b.iterrows():
            for campo in planilha_b.columns:
                print(campo + ":", registro[campo])
            print()


    print()
    print("✓ Comparação concluída!")


    relatorio_resumo = pd.DataFrame({
        "Categoria": [
            "Registros iguais",
            "Registros diferentes",
            "Registros só na Planilha A",
            "Registros só na Planilha B"
        ],
        "Quantidade": [
            iguais,
            diferentes,
            len(ids_so_na_a),
            len(ids_so_na_b)
        ]
    })


    relatorio_diferencas = pd.DataFrame(detalhes_diferencas)
    relatorio_so_na_a = registros_so_na_a.copy()
    relatorio_so_na_b = registros_so_na_b.copy()


    try:
        with pd.ExcelWriter(
            "relatorio_comparacao.xlsx",
            engine="openpyxl"
        ) as escritor:

            relatorio_resumo.to_excel(
                escritor,
                sheet_name="Resumo",
                index=False
            )

            relatorio_diferencas.to_excel(
                escritor,
                sheet_name="Diferenças",
                index=False
            )

            relatorio_so_na_a.to_excel(
                escritor,
                sheet_name="Só na A",
                index=False
            )

            relatorio_so_na_b.to_excel(
                escritor,
                sheet_name="Só na B",
                index=False
            )
            for planilha in escritor.book.worksheets:
                planilha.freeze_panes = "A2"
                planilha.auto_filter.ref = planilha.dimensions

                for celula in planilha[1]:
                 celula.font = Font(bold=True, color="FFFFFF")
                 celula.fill = PatternFill(
                 fill_type="solid",
                 fgColor="1F4E78"
                 )
                 celula.alignment = Alignment(
                 horizontal="center",
                 vertical="center"
                 )
                planilha.row_dimensions[1].height = 25

                for coluna in planilha.columns:
                    maior_tamanho = 0

                    for celula in coluna:
                        if celula.value is not None:
                            tamanho = len(str(celula.value))

                            if tamanho > maior_tamanho:
                                maior_tamanho = tamanho

                    letra_coluna = get_column_letter(coluna[0].column)

                    planilha.column_dimensions[letra_coluna].width = maior_tamanho + 2

        print("✓ Relatório gerado: relatorio_comparacao.xlsx")

    except PermissionError:
        print()
        print("ERRO: Não foi possível gerar o relatório.")
        print("O arquivo 'relatorio_comparacao.xlsx' está aberto.")
        print("Feche o arquivo e tente novamente.")
        exit()

comparar_planilhas()
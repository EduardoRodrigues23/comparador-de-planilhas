import pandas as pd
import numbers
import os


def teste_arquivos_existem():
    assert os.path.exists("vendas_sistema.xlsx")
    assert os.path.exists("vendas_financeiro.xlsx")


def teste_coluna_id_existe():
    planilha_a = pd.read_excel("vendas_sistema.xlsx")
    planilha_b = pd.read_excel("vendas_financeiro.xlsx")

    assert "ID" in planilha_a.columns
    assert "ID" in planilha_b.columns


def teste_ids_em_comum():
    planilha_a = pd.read_excel(
        "vendas_sistema.xlsx",
        dtype={"ID": str}
    )

    planilha_b = pd.read_excel(
        "vendas_financeiro.xlsx",
        dtype={"ID": str}
    )

    planilha_a["ID"] = planilha_a["ID"].str.strip()
    planilha_b["ID"] = planilha_b["ID"].str.strip()

    ids_comuns = planilha_a.loc[
        planilha_a["ID"].isin(planilha_b["ID"]),
        "ID"
    ].tolist()

    assert ids_comuns == ["001", "002", "003", "005"]


def teste_ids_exclusivos():
    planilha_a = pd.read_excel(
        "vendas_sistema.xlsx",
        dtype={"ID": str}
    )

    planilha_b = pd.read_excel(
        "vendas_financeiro.xlsx",
        dtype={"ID": str}
    )

    planilha_a["ID"] = planilha_a["ID"].str.strip()
    planilha_b["ID"] = planilha_b["ID"].str.strip()

    ids_so_na_a = planilha_a.loc[
        ~planilha_a["ID"].isin(planilha_b["ID"]),
        "ID"
    ].tolist()

    ids_so_na_b = planilha_b.loc[
        ~planilha_b["ID"].isin(planilha_a["ID"]),
        "ID"
    ].tolist()

    assert ids_so_na_a == ["004"]
    assert ids_so_na_b == ["006"]


def teste_registros_iguais():
    planilha_a = pd.read_excel(
        "vendas_sistema.xlsx",
        dtype={"ID": str}
    )

    planilha_b = pd.read_excel(
        "vendas_financeiro.xlsx",
        dtype={"ID": str}
    )

    planilha_a["ID"] = planilha_a["ID"].str.strip()
    planilha_b["ID"] = planilha_b["ID"].str.strip()

    colunas_comuns = planilha_a.columns.intersection(
        planilha_b.columns
    )

    ids_iguais = []

    ids_comuns = planilha_a.loc[
        planilha_a["ID"].isin(planilha_b["ID"]),
        "ID"
    ]

    for id in ids_comuns:

        linha_a = planilha_a[
            planilha_a["ID"] == id
        ][colunas_comuns].reset_index(drop=True)

        linha_b = planilha_b[
            planilha_b["ID"] == id
        ][colunas_comuns].reset_index(drop=True)

        if linha_a.equals(linha_b):
            ids_iguais.append(id)

    assert ids_iguais == ["003", "005"]


def teste_registros_diferentes():
    planilha_a = pd.read_excel(
        "vendas_sistema.xlsx",
        dtype={"ID": str}
    )

    planilha_b = pd.read_excel(
        "vendas_financeiro.xlsx",
        dtype={"ID": str}
    )

    planilha_a["ID"] = planilha_a["ID"].str.strip()
    planilha_b["ID"] = planilha_b["ID"].str.strip()

    colunas_comuns = planilha_a.columns.intersection(
        planilha_b.columns
    )

    ids_diferentes = []

    ids_comuns = planilha_a.loc[
        planilha_a["ID"].isin(planilha_b["ID"]),
        "ID"
    ]

    for id in ids_comuns:

        linha_a = planilha_a[
            planilha_a["ID"] == id
        ][colunas_comuns].reset_index(drop=True)

        linha_b = planilha_b[
            planilha_b["ID"] == id
        ][colunas_comuns].reset_index(drop=True)

        if not linha_a.equals(linha_b):
            ids_diferentes.append(id)

    assert ids_diferentes == ["001", "002"]


def teste_diferencas_numericas():
    planilha_a = pd.read_excel(
        "vendas_sistema.xlsx",
        dtype={"ID": str}
    )

    planilha_b = pd.read_excel(
        "vendas_financeiro.xlsx",
        dtype={"ID": str}
    )

    planilha_a["ID"] = planilha_a["ID"].str.strip()
    planilha_b["ID"] = planilha_b["ID"].str.strip()

    colunas_comuns = planilha_a.columns.intersection(
        planilha_b.columns
    )

    id = "002"

    linha_a = planilha_a[
        planilha_a["ID"] == id
    ][colunas_comuns].reset_index(drop=True)

    linha_b = planilha_b[
        planilha_b["ID"] == id
    ][colunas_comuns].reset_index(drop=True)

    diferencas = []

    diferencas_colunas = (linha_a != linha_b).columns[
        (linha_a != linha_b).any()
    ]

    for campo in diferencas_colunas:

        valor_a = linha_a.iloc[0][campo]
        valor_b = linha_b.iloc[0][campo]

        if pd.isna(valor_a) or pd.isna(valor_b):
            diferenca = "Alterado"

        elif isinstance(valor_a, numbers.Number) and isinstance(
            valor_b, numbers.Number
        ):
            diferenca = valor_b - valor_a

        else:
            diferenca = "Alterado"

        diferencas.append({
            "Campo": campo,
            "Diferença": diferenca
        })

    assert {
        "Campo": "Quantidade",
        "Diferença": 1
    } in diferencas

    assert {
        "Campo": "Valor",
        "Diferença": 30
    } in diferencas


def teste_diferencas_texto():
    planilha_a = pd.read_excel(
        "vendas_sistema.xlsx",
        dtype={"ID": str}
    )

    planilha_b = pd.read_excel(
        "vendas_financeiro.xlsx",
        dtype={"ID": str}
    )

    planilha_a["ID"] = planilha_a["ID"].str.strip()
    planilha_b["ID"] = planilha_b["ID"].str.strip()

    colunas_comuns = planilha_a.columns.intersection(
        planilha_b.columns
    )

    id = "001"

    linha_a = planilha_a[
        planilha_a["ID"] == id
    ][colunas_comuns].reset_index(drop=True)

    linha_b = planilha_b[
        planilha_b["ID"] == id
    ][colunas_comuns].reset_index(drop=True)

    diferencas = []

    campos_diferentes = (linha_a != linha_b).columns[
        (linha_a != linha_b).any()
    ]

    for campo in campos_diferentes:

        valor_a = linha_a.iloc[0][campo]
        valor_b = linha_b.iloc[0][campo]

        if pd.isna(valor_a) or pd.isna(valor_b):
            diferenca = "Alterado"

        elif isinstance(valor_a, numbers.Number) and isinstance(
            valor_b, numbers.Number
        ):
            diferenca = valor_b - valor_a

        else:
            diferenca = "Alterado"

        diferencas.append({
            "Campo": campo,
            "Diferença": diferenca
        })

    assert {
        "Campo": "Cliente",
        "Diferença": "Alterado"
    } in diferencas

    assert {
        "Campo": "Data",
        "Diferença": "Alterado"
    } in diferencas


def teste_ids_duplicados():
    planilha = pd.DataFrame({
        "ID": ["001", "002", "002", "003"],
        "Cliente": ["João", "Maria", "Pedro", "Ana"]
    })

    ids_duplicados = planilha[
        planilha["ID"].duplicated()
    ]["ID"].unique()

    assert list(ids_duplicados) == ["002"]


def teste_espacos_nos_ids():
    planilha = pd.DataFrame({
        "ID": ["001 ", "002", " 003", "004"]
    })

    planilha["ID"] = planilha["ID"].str.strip()

    assert planilha["ID"].tolist() == [
        "001",
        "002",
        "003",
        "004"
    ]


print("TESTES DO COMPARADOR")
print("====================")

teste_arquivos_existem()
print("✓ Arquivos existem.")

teste_coluna_id_existe()
print("✓ Coluna ID existe.")

teste_ids_em_comum()
print("✓ IDs em comum.")

teste_ids_exclusivos()
print("✓ IDs exclusivos.")

teste_registros_iguais()
print("✓ Registros iguais.")

teste_registros_diferentes()
print("✓ Registros diferentes.")

teste_diferencas_numericas()
print("✓ Diferenças numéricas.")

teste_diferencas_texto()
print("✓ Diferenças de texto.")

teste_ids_duplicados()
print("✓ IDs duplicados.")

teste_espacos_nos_ids()
print("✓ Espaços nos IDs.")

print()
print("================================")
print("✓ TODOS OS TESTES PASSARAM!")
print("================================")
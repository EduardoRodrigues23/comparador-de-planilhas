import os
import pandas as pd
from comparar_planilhas import comparar_planilhas


def teste_erros():
    import tempfile
    from pathlib import Path

    with tempfile.TemporaryDirectory() as pasta:
        pasta = Path(pasta)

        arquivo_duplicado = pasta / "duplicados.xlsx"
        arquivo_sem_id = pasta / "sem_id.xlsx"

        pd.DataFrame({
            "ID": ["001", "002", "002"],
            "Cliente": ["João", "Maria", "Pedro"]
        }).to_excel(arquivo_duplicado, index=False)

        pd.DataFrame({
            "Cliente": ["João", "Maria"],
            "Valor": [100, 200]
        }).to_excel(arquivo_sem_id, index=False)

        resultado = comparar_planilhas(
            str(arquivo_duplicado),
            "vendas_financeiro.xlsx"
        )

        assert isinstance(resultado, str)
        assert "identificadores duplicados" in resultado

        print("✓ IDs duplicados identificados corretamente.")

        resultado = comparar_planilhas(
            str(arquivo_sem_id),
            "vendas_financeiro.xlsx"
        )

        assert isinstance(resultado, str)
        assert "não possui a coluna 'ID'" in resultado

        print("✓ Ausência da coluna ID identificada corretamente.")


def teste_gerar_relatorio():
    resultado = comparar_planilhas(
        "vendas_sistema.xlsx",
        "vendas_financeiro.xlsx"
    )

    assert resultado is True, f"Falha na comparação: {resultado}"

    assert os.path.exists("relatorio_comparacao.xlsx")

    relatorio = pd.ExcelFile("relatorio_comparacao.xlsx")

    assert relatorio.sheet_names == [
        "Resumo",
        "Diferenças",
        "Só na A",
        "Só na B"
    ]
    resumo = pd.read_excel(
        "relatorio_comparacao.xlsx",
        sheet_name="Resumo"
    )

    esperado = {
        "Registros iguais": 2,
        "Registros diferentes": 2,
        "Registros só na Planilha A": 1,
        "Registros só na Planilha B": 1
    }

    for categoria, quantidade in esperado.items():
        valor_encontrado = resumo.loc[
            resumo["Categoria"] == categoria,
            "Quantidade"
        ].iloc[0]

        assert valor_encontrado == quantidade, (
            f"Erro em {categoria}: "
            f"esperado {quantidade}, encontrado {valor_encontrado}"
        )
        diferencas = pd.read_excel(
        "relatorio_comparacao.xlsx",
        sheet_name="Diferenças",
        dtype={"ID": str}
    )

    def verificar_diferenca(id, campo, esperado):
        registro = diferencas[
            (diferencas["ID"] == id) &
            (diferencas["Campo"] == campo)
        ]

        assert len(registro) == 1, (
            f"Diferença não encontrada: ID {id}, campo {campo}"
        )

        valor = registro.iloc[0]["Diferença"]

        assert valor == esperado, (
            f"ID {id}, campo {campo}: "
            f"esperado {esperado}, encontrado {valor}"
        )

    verificar_diferenca("002", "Quantidade", 1)
    verificar_diferenca("002", "Valor", 30)
    verificar_diferenca("001", "Cliente", "Alterado")
    verificar_diferenca("001", "Data", "Alterado")

    somente_a = pd.read_excel(
        "relatorio_comparacao.xlsx",
        sheet_name="Só na A",
        dtype={"ID": str}
    )

    somente_b = pd.read_excel(
        "relatorio_comparacao.xlsx",
        sheet_name="Só na B",
        dtype={"ID": str}
    )

    assert somente_a["ID"].tolist() == ["004"], (
        "Erro nos registros exclusivos da Planilha A."
    )

    assert somente_b["ID"].tolist() == ["006"], (
        "Erro nos registros exclusivos da Planilha B."
    )

    print("✓ Registros exclusivos estão corretos.")

    print("✓ Diferenças numéricas e de texto estão corretas.")
    print("✓ Quantidades do relatório estão corretas.")
    print("✓ Comparador executado com sucesso.")
    print("✓ Relatório Excel gerado.")
    print("✓ As quatro abas foram encontradas.")


teste_gerar_relatorio()

teste_erros()

print("\n✓ TESTE DE INTEGRAÇÃO PASSOU!")
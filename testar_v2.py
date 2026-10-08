import os
import tempfile
import pandas as pd

from comparar_planilhas import comparar_planilhas


def testar_colunas_diferentes():
    with tempfile.TemporaryDirectory() as pasta:
        arquivo_a = os.path.join(pasta, "planilha_a.xlsx")
        arquivo_b = os.path.join(pasta, "planilha_b.xlsx")
        relatorio = os.path.join(pasta, "relatorio.xlsx")

        dados_a = pd.DataFrame({
            "Pedido": ["001", "002", "003"],
            "Cliente": ["João", "Maria", "Pedro"],
            "Valor": [150, 300, 500]
        })

        dados_b = pd.DataFrame({
            "Código": ["001", "002", "004"],
            "Cliente": ["João", "Maria", "Ana"],
            "Valor": [150, 350, 200]
        })

        dados_a.to_excel(arquivo_a, index=False)
        dados_b.to_excel(arquivo_b, index=False)

        resultado = comparar_planilhas(
            arquivo_a=arquivo_a,
            arquivo_b=arquivo_b,
            caminho_relatorio=relatorio,
            coluna_id_a="Pedido",
            coluna_id_b="Código"
        )

        assert resultado is True, resultado

        resumo = pd.read_excel(
            relatorio,
            sheet_name="Resumo"
        )

        quantidades = dict(
            zip(resumo["Categoria"], resumo["Quantidade"])
        )

        assert quantidades["Registros iguais"] == 1
        assert quantidades["Registros diferentes"] == 1
        assert quantidades["Registros só na Planilha A"] == 1
        assert quantidades["Registros só na Planilha B"] == 1

        diferencas = pd.read_excel(
            relatorio,
            sheet_name="Diferenças",
            dtype={"ID": str}
        )

        assert len(diferencas) == 1
        assert diferencas.iloc[0]["ID"] == "002"
        assert diferencas.iloc[0]["Campo"] == "Valor"
        assert float(diferencas.iloc[0]["Diferença"]) == 50

        print("✓ Colunas Pedido e Código comparadas corretamente.")
        print("✓ Quantidades do relatório estão corretas.")
        print("✓ Diferença de R$ 50 identificada.")
        print("\n✓ TESTE DA VERSÃO 2.0 PASSOU!")


if __name__ == "__main__":
    testar_colunas_diferentes()
# Comparador de Planilhas

Projeto desenvolvido em Python para comparar duas planilhas Excel e identificar registros iguais, diferentes e exclusivos de cada arquivo.

O sistema utiliza um ID para localizar os mesmos registros nas duas planilhas e gera um relatório com as diferenças encontradas.

## Como funciona

1. O usuário informa o nome de duas planilhas Excel.
2. O sistema utiliza o ID para encontrar os mesmos registros nas duas planilhas.
3. Os registros são comparados campo por campo.
4. O sistema identifica:
   - Registros iguais;
   - Registros diferentes;
   - Registros que existem apenas na Planilha A;
   - Registros que existem apenas na Planilha B.
5. Um relatório de comparação é gerado em um novo arquivo Excel.

## Tecnologias utilizadas

- Python
- Pandas
- OpenPyXL
- Excel

## Como executar

Instale as dependências:

```bash
python -m pip install pandas openpyxl
```

Execute:

```bash
python comparar_planilhas.py
```

Informe as duas planilhas quando solicitado.

O relatório será gerado em `relatorio_comparacao.xlsx`.
# Dashboard de Vendas em Python

Este projeto é um **dashboard de vendas** feito em Python usando **Tkinter**, **Pandas** e **Matplotlib**. Ele mostra métricas, tabelas e gráficos a partir de um arquivo Excel (`vendas.xlsx`).

---

## Estrutura do Projeto

- `dashboard_vendas.py` → Script principal com todas as funções e a interface.
- `vendas.xlsx` → Planilha de dados. Precisa ter as colunas:
  - `Produto`
  - `Categoria`
  - `Mes`
  - `Vendas`
  - `Preco_Unitario`

---

## Funcionalidades

1. **Leitura de Dados**
   - Lê os dados do Excel.
   - Verifica se todas as colunas obrigatórias estão presentes.
   - Calcula a coluna `Receita` automaticamente (`Vendas * Preco_Unitario`).

2. **Métricas**
   - Total de vendas.
   - Receita total.
   - Ticket médio.
   - Produto mais vendido.

3. **Gráficos**
   - Vendas por produto (barras).
   - Receita por categoria (pizza).
   - Evolução mensal da receita (linha).

4. **Tabela**
   - Mostra resumo por produto:
     - Total de vendas
     - Receita total
     - Preço unitário médio
     - Ticket médio
     - Participação na receita total

5. **Interface**
   - Criada com Tkinter.
   - Divide a tela entre resumo, tabela e gráficos.
   - Mostra mensagens de erro se faltar arquivo ou colunas.

---

## Como Usar

1. Coloque o `vendas.xlsx` na mesma pasta do script ou ajuste o caminho no código.
2. Execute:

```bash
python dashboard_vendas.py
```

3. A tela vai abrir mostrando as métricas, tabela e gráficos.

---

## Observações

- Precisa ter Python 3.
- Bibliotecas necessárias:
  - `pandas`
  - `matplotlib`
  - `tkinter` (normalmente já vem com o Python)

- Para mudar o arquivo de dados, altere a linha:

```python
arquivo_excel = os.path.join(pasta_atual, "vendas.xlsx")
```

---

## Possíveis Melhorias

- Selecionar o arquivo Excel via janela de diálogo.
- Filtrar por categoria ou mês.
- Exportar gráficos em PDF ou imagem.
- Melhorar cores e layout dos gráficos.

---

## Autor

**Júlia Rodrigues**

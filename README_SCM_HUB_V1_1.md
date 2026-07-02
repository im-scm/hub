# SCM Analytics Hub - V1.1 ajustes solicitados

## Arquivos do pacote

- `app.py`: Hub atualizado usando `st.logo()` para exibir a logo acima do menu nativo.
- `pages/02_Paper_Base.py`: Paper Base V1.1 com ajustes no gráfico e nas tabelas de ranking.
- `logo_snippet_for_01_Cockpit_Papel.txt`: trecho para aplicar a mesma logo acima do menu na página Cockpit Papel.

## Alterações

### Todas as páginas
- Logo passa a usar `st.logo("assets/impress_logo.png", size="large")`, que posiciona a logo acima do menu nativo do Streamlit.

### Paper Base
- Average price do gráfico passa a aparecer na base visual das colunas.
- Ranking remove `Supplier Code`.
- Ranking remove `Quantity ton`.
- Ranking inclui `Average price` ponderado por fornecedor no período.
- Ranking mantém/inclui `Share %`, relativo ao total de Quantity KG da tabela.
- Tabelas de ranking sem barra de rolagem vertical.

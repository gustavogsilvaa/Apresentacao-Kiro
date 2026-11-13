# Apresentação parcial do TF — v8 (pasta TF)

`Apresentacao_TF_Amarracao_parcial_v8.pptx` · 27 slides · 16:9 · Times New Roman

Mesma estrutura e layout da v7 (`amarracao_v3`, branch `apresentacao-tf-v7`), atualizada com a simulação
nova: **5 condições de carga** (lastro 9 m, 25% 10,25 m, 50% 11,5 m, 75% 12,75 m e cheio 14 m), 394
condições ambientais cada, cabos 11 e 12 no cabeço **R** e cabo com MBL de 736 kN (75,05 t), Ø 80 mm e
EA de 5.632 kN.

```bash
pip install python-pptx matplotlib pillow pandas
python gerar_graficos.py      # gráficos novos em imagens/
python criar_apresentacao.py  # gera o .pptx
```

## O que mudou em relação à v7

Os números de slide abaixo são os da v8.

| Slide | v7 | v8 |
|---|---|---|
| 2, 4, 5, 7, 10, 11 | calados de 9 e 14 m; maré de 0 e 1,58 m; MBL de 75 t | 5 condições de carga; maré do CA (−0,20 a 1,60 m); MBL de 736 kN, Ø 80 mm, EA 5.632 kN |
| 3 · Motivação | números sem fonte | fonte de cada número: BRASIL (2026), SILVA (2026a) e UK P&I CLUB (2009) |
| 6 · **O navio: medidas e movimentos** | — | **novo:** LOA, boca, pontal, calado e borda livre; os 6 movimentos e os 3 que a simulação registra |
| 12 · Resumo estatístico | 394 condições do calado 9 m | picos dos elementos nas 5 condições + máximo por condição |
| 13 · Distribuição | histogramas das 394 condições | mediana, P90 e máximo dos picos por condição de carga |
| 14 · Influência da maré | calado 9 m (cabo 9, cabeço S) | condição 75%, com 11-R e 12-R |
| 15 · Clima ou arranjo (1/2) | matriz de correlação | eventos governantes (75% e cheio) |
| 16 · Clima ou arranjo (2/2) | boxplot por cabo crítico | tração × comprimento de cada cabo (75%) |
| 17 · Combinações críticas | 5 piores acima de 55% MBL | 5 maiores trações e esforços (nenhum acima do limite) |
| 18 · Envelope crítico | calado 9 m | condição 75% |
| — · Matriz de operabilidade | 394 condições do calado 9 m | **retirado** (precisa dos resultados caso a caso) |
| 19–21 · Geometria | maré 0 e 1,58 m; 11 e 12 no cabeço S | maré −0,20 e 1,60 m; 11 e 12 no cabeço R; valores calculados no script |
| 22 · **Efeito do carregamento** | — | **novo:** utilização máxima e ângulo vertical × calado |
| 23–25 · Conclusões, alternativas, próximos passos | cabo 9 e cabeço S | resultados novos |
| 26 · Referências | TF4.0 (calado 9 m) | TF4.0 (5 condições de carga) e BRASIL (2026) |

## Fontes do slide 3

| Número | Fonte citada | Conferência |
|---|---|---|
| >100 Mt por ano | BRASIL (2026), Ministério de Portos e Aeroportos | 186,4 Mt em 2025, segundo a APS |
| 14 de 60 berços | SILVA (2026a) | compilado no TF I; fonte primária não localizada na web |
| US$ 34 mi em 20 anos | UK P&I CLUB (2009), *LP News*, jan. 2009 | o boletim cita “over US$34 million” em 20 anos |
| até 40% com navio passante | SILVA (2026a) | compilado no TF I; fonte primária não localizada na web |

## Dados (`dados/`)

| Arquivo | Origem | Usado em |
|---|---|---|
| `resumo_esforcos_maximos_75%.csv` | resumo exportado do software (cópia exata) | 6, 14, 16, 17, 18 e textos |
| `estatisticas_colab.csv` | tabela “Resumo estatístico geral” do notebook, transcrita da saída do Colab | 12, 13, 22, 23–25 |
| `eventos_cheio_colab.csv` | gráfico “eventos governantes” da condição cheio, transcrito da saída do Colab | 15 |
| `ca_eventos_governantes.csv` | linhas do `CA_394_pontos.csv` dos eventos citados (ID = linha do CA) | 15 |

A geometria (slides 19 a 22) é calculada em `dados_tf.py` com as mesmas coordenadas e fórmulas de
`calcular_geometria_amarracao` do notebook. Os valores conferem com a saída do Colab.

Imagens: `imagens/` (gráficos novos, `gerar_graficos.py`), `imagens_originais/` (maré da DHN e arranjo do
software, do autor) e `imagens_recriadas/` (esquemas do layout típico e dos critérios, da v7).

## Pendências

- **Fontes primárias de “14 de 60” e “até 40%” (slide 3):** hoje citam o TF I. Se o TF I tirou esses
  números de outro trabalho, vale citar esse trabalho direto no slide.
- **Resumos de lastro, 25%, 50% e cheio:** só o resumo de 75% está em `dados/`. Por isso, os slides 14, 16,
  17 e 18 mostram a condição 75%, e não a mais crítica (cheio). Com os outros resumos (ou o
  `esforcos_todas_condicoes.csv` do notebook), esses slides podem passar a mostrar o navio cheio, e o
  nome do cabeço que chega a 98,1% SWL pode entrar nas conclusões.
- **Reação e compressão nas defensas:** ficaram fora da tabela do slide 12, porque a saída do Colab
  usada não tinha os valores das 5 condições.
- **Resultados caso a caso (394 × 5):** sem eles não há histograma das condições, matriz de correlação nem
  matriz de operabilidade. O slide 25 lista a exportação como próximo passo.
- **Pontos a revisar da v7 que continuam valendo:** convenções da geometria no notebook (lado das buzinas
  e posição dos cabeços), limite de 50% MBL para cabos sintéticos e o sentido do risco com a maré
  (ver `amarracao_v3/README.md` no branch `apresentacao-tf-v7`).

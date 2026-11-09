# Apresentação parcial do TF — Sistema de amarração em conjuntos portuários

`Apresentacao_TF_Amarracao_parcial.pptx` · 26 slides · 16:9 · Times New Roman (versão atual, v7)
`Apresentacao_Amarracao_v3.pptx` · 13 slides (versão anterior, só com os gráficos)

A v7 segue o modelo em PDF editado pelo autor (24 páginas, com marcações). Diferenças em relação
ao modelo: o slide 21 (matriz calado × maré) é novo e o slide 23 (alternativas de segurança) foi
mantido, porque o roteiro ainda o cita.

| Etapa | Slides | Conteúdo |
|---|---|---|
| Abertura | 1–2 | Capa e roteiro |
| 1 · Motivação | 3 | Santos, segurança, operação e meio ambiente |
| 2 · Justificativa | 4 | Problema, lacuna, proposta do TF e objetivo desta etapa |
| 3 · Cenário | 5–8 | Terminal T-Grão, estudo de maré (DHN 2025), layout típico do Panamax e arranjo simulado |
| 4 · Abordagem | 9–10 | Etapas do trabalho (com o produto de cada uma) e critérios de avaliação |
| 5 · Resultados | 11–21 | Cada gráfico ou tabela com os cards 1 e 2; geometria com maré de 0 m e de 1,58 m; matriz calado × maré (21) |
| 6 · Conclusões | 22–24 | Conclusões parciais, alternativas de segurança e próximos passos |
| Referências | 25 | Lista no formato ABNT NBR 6023 |
| Encerramento | 26 | Perguntas |

Leitura de cada slide: rótulo da etapa → título → mensagem principal (subtítulo em itálico) →
gráfico ou tabela → cards numerados 1 e 2. A barra no topo mostra a etapa atual.

Fontes: relatório do TF I (SILVA, 2026a), notebook *TF4.0* (SILVA, 2026b), OCIMF (2018), PIANC
(2017, 2025), DHN (2025), Southern Ports (2026) e UK P&I Club (2009). A lista completa está no slide 25.

## Imagens

`criar_apresentacao.py` usa `imagens_originais/<nome>.png` quando o arquivo existe e, caso
contrário, `imagens_recriadas/<nome>.png`.

Imagens do autor, extraídas do modelo em PDF (`imagens_originais/`):

| Arquivo | Slide | Origem |
|---|---|---|
| `08_mare_dhn.png` | 6 | Datawrapper, máximas e mínimas mensais da DHN 2025 |
| `00_arranjo_projeto.png` | 8 | software de simulação, arranjo `01_Panamax_cheio_BE` |
| `01a_hist_mbl.png`, `01b_hist_swl.png` | 12 | Colab, histogramas de % MBL e % SWL |
| `04_correlacao_matriz.png` | 14 | Colab, matriz de correlação |

Gráficos redesenhados (`imagens_recriadas/`, fonte STIX, com desenho de Times):

| Arquivo | Slide | Dados |
|---|---|---|
| `07_layout_panamax.png` | 7 | esquema do arranjo típico 4-1-1 por extremidade (Figura 9 do TF I) |
| `crit_mbl.png`, `crit_swl.png`, `crit_angulo.png` | 10 | esquemas dos três critérios (55% MBL = 41 t; 100% SWL = 75 t; 25°) |
| `02_influencia_mare.png` | 13 | picos de 75,3% e 110,8% exatos; demais picos e marés lidos do gráfico original |
| `04c_boxplot_cabos_alto.png` | 15 | quartis lidos do gráfico original |
| `05_envelope_critico.png` | 17 | posições lidas do gráfico original |
| `06_matriz_operabilidade.png` | 18 | exatos (valores de cada célula) |
| `01_histogramas.png` | — | reserva do slide 12, usada só se os histogramas do Colab faltarem |

`00_plano_amarracao.png`, `04_correlacao_boxplot.png`, `04a_correlacao_barras.png`,
`04b_boxplot_cabos.png`, `08_mare_dhn.png` (versão recriada) e `ic_*.png` continuam sendo gerados,
mas não entram na versão atual.

Os slides 19 a 21 são tabelas e formas editáveis. Os valores da matriz calado × maré (slide 21) foram
conferidos com `calcular_geometria_amarracao` do notebook para os 12 cabos.

```bash
pip install python-pptx matplotlib pillow
python gerar_graficos.py      # só se quiser recriar os gráficos
python criar_apresentacao.py
```

## Pontos a revisar

- **58 ou 56 condições (slide 12):** o filtro do notebook conta 58 condições acima de 55% MBL
  (14,7%). A soma das barras à direita da linha (32 + 19 + 5 = 56) deixa de fora 2 condições da barra
  que a linha de 55% corta. O slide 16 também usa 58.
- **Cabeços dos lançantes de popa (slide 8):** o texto cita os cabeços S e R, mas os resultados
  (slides 13, 16, 17 e 19–20) usam 5 cabeços (A, B, G, N e S), com os cabos 1, 2, 11 e 12 no S. O nome
  da imagem (`01_Panamax_cheio_BE`) sugere o caso carregado; os resultados são do calado de 9 m.
- **Maré de 1,58 m (slide 6):** o texto a chama de média de maré alta, mas a máxima prevista pela
  DHN para 2025 é 1,52 m.
- **Calado da etapa atual (slide 24):** a etapa atual cita 9 m e 14 m, e a lista seguinte ainda propõe
  simular o calado carregado (14 m). As 394 condições do notebook TF4.0 são de 9 m; o 14 m aparece só
  no cálculo de geometria.
- **Gráfico de maré (slide 6):** o Datawrapper empilha mínima e máxima. Nos meses com mínima positiva
  (maio, junho, novembro e dezembro), a barra da máxima começa no fim da mínima e passa do valor real.
- **Sentido do risco com a maré:** a seção 4.6.2 do TF I diz que a maré baixando estica os cabos e a
  maré subindo os afrouxa. Neste berço, a buzina fica acima do cabeço, então ocorre o contrário: subir
  o navio alonga os cabos e baixá-lo os afrouxa.
- **Geometria no notebook:** em `calcular_geometria_amarracao`, os cabeços ficam em `y = 3,5 m`, do
  lado do navio em relação à face do cais, e 7 das 12 buzinas usadas têm `Y_local` negativo (costado
  oposto ao cais na atracação por boreste). Vale conferir essas convenções antes de usar comprimentos e
  ângulos como resultado final.
- **Limite dos cabos:** o notebook usa 55% do MBL. As Figuras 1.3 e 1.4 do MEG4 indicam WLL de 55% para
  cabos de aço e 50% para cabos sintéticos. Se os cabos forem de PPL/PE, o limite de 50% é o mais adequado.

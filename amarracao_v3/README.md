# Apresentação parcial do TF — Sistema de amarração em conjuntos portuários

`Apresentacao_TF_Amarracao_parcial.pptx` · 22 slides · 16:9 · Times New Roman (versão atual)
`Apresentacao_Amarracao_v3.pptx` · 13 slides (versão anterior, só com os gráficos)

| Etapa | Slides | Conteúdo |
|---|---|---|
| Abertura | 1–2 | Capa e roteiro |
| 1 · Motivação | 3 | Santos, Panamax, segurança, operação e meio ambiente |
| 2 · Justificativa | 4 | Lacuna no monitoramento dos cabos e objetivo desta etapa |
| 3 · Cenário | 5–8 | Terminal T-Grão, estudo de maré (DHN 2025), layout típico do Panamax e arranjo simulado |
| 4 · Abordagem | 9 | Caracterização → simulação → pós-processamento → geometria, e critérios |
| 5 · Resultados | 10–18 | Cada gráfico com "O que o gráfico lê" (1) e "Resultado obtido" (2) |
| 6 · Conclusões | 19–20 | Conclusões parciais e próximos passos |
| Referências | 21 | Lista no formato ABNT NBR 6023 |
| Encerramento | 22 | Perguntas |

Leitura de cada slide: rótulo da etapa → título → mensagem principal (subtítulo em itálico) →
gráfico ou tabela → cards numerados 1 e 2. A barra no topo mostra a etapa atual.

Fontes: relatório do TF I (SILVA, 2026a), notebook *TF4.0* (SILVA, 2026b), OCIMF (2018), PIANC
(2017, 2025), DHN (2025), Southern Ports (2026) e UK P&I Club (2009). A lista completa está no slide 21.

## Imagens

Os gráficos em `imagens_recriadas/` foram redesenhados (fonte STIX, com desenho de Times):

| Arquivo | Dados |
|---|---|
| `00_plano_amarracao.png` | posições longitudinais do notebook; fora de escala na transversal |
| `01_histogramas.png` | exatos (contagens de cada barra) |
| `02_influencia_mare.png` | picos de 75,3% e 110,8% exatos; demais picos e marés lidos do gráfico original |
| `04_correlacao_boxplot.png` | correlações exatas; quartis do boxplot lidos do gráfico original |
| `05_envelope_critico.png` | posições lidas do gráfico original |
| `06_matriz_operabilidade.png` | exatos (valores de cada célula) |
| `07_layout_panamax.png` | esquema do arranjo típico 4-1-1 por extremidade (Figura 9 do TF I) |
| `08_mare_dhn.png` | exatos (Tabela 7 do TF I, dados da DHN 2025) |

Para usar os PNGs originais do Colab, salve-os em `imagens_originais/` com os mesmos nomes e rode:

```bash
pip install python-pptx matplotlib pillow
python gerar_graficos.py      # só se quiser recriar os gráficos
python criar_apresentacao.py
```

## Pontos a revisar

- **Geometria no notebook:** em `calcular_geometria_amarracao`, os cabeços ficam em `y = 3,5 m`, do
  lado do navio em relação à face do cais, e 7 das 12 buzinas usadas têm `Y_local` negativo (costado
  oposto ao cais na atracação por boreste). Vale conferir essas convenções antes de usar comprimentos e
  ângulos como resultado final.
- **Limite dos cabos:** o notebook usa 55% do MBL. A Figura 1.4 do MEG4 indica WLL de 55% para cabos de
  aço e 50% para cabos sintéticos. Se os cabos forem de PPL/PE, o limite de 50% é o mais adequado.

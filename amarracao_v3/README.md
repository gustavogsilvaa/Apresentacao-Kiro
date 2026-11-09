# Apresentação parcial do TF — Sistema de amarração em conjuntos portuários

`Apresentacao_TF_Amarracao_parcial.pptx` · 19 slides · 16:9 (versão atual)
`Apresentacao_Amarracao_v3.pptx` · 13 slides (versão anterior, só com os gráficos)

| Etapa | Slides | Conteúdo |
|---|---|---|
| Abertura | 1–2 | Capa e roteiro |
| 1 · Motivação | 3 | Santos, Panamax, segurança, operação e meio ambiente |
| 2 · Justificativa | 4 | Lacuna no monitoramento dos cabos e objetivo desta etapa |
| 3 · Cenário | 5–6 | Terminal T-Grão, navio, maré (DHN 2025) e arranjo de amarração |
| 4 · Abordagem | 7 | Caracterização → simulação → pós-processamento → geometria, e critérios |
| 5 · Resultados | 8–16 | Cada gráfico com "O que o gráfico lê" e "Resultado obtido" |
| 6 · Conclusões | 17–18 | Conclusões parciais e próximos passos |
| Encerramento | 19 | Perguntas |

Fontes: relatório do TF I (PNV3510) para as etapas 1 a 4 e notebook *TF4.0 - Colab*
(calado 9 m, 394 condições) para os resultados.

## Imagens

Os gráficos em `imagens_recriadas/` foram redesenhados a partir das saídas do notebook:

| Arquivo | Dados |
|---|---|
| `00_plano_amarracao.png` | posições longitudinais do notebook; fora de escala na transversal |
| `01_histogramas.png` | exatos (contagens de cada barra) |
| `02_influencia_mare.png` | picos de 75,3% e 110,8% exatos; demais picos e marés lidos do gráfico original |
| `04_correlacao_boxplot.png` | correlações exatas; quartis do boxplot lidos do gráfico original |
| `05_envelope_critico.png` | posições lidas do gráfico original |
| `06_matriz_operabilidade.png` | exatos (valores de cada célula) |

Para usar os PNGs originais do Colab, salve-os em `imagens_originais/` com os mesmos nomes e rode:

```bash
pip install python-pptx matplotlib pillow
python gerar_graficos.py      # só se quiser recriar os gráficos
python criar_apresentacao.py
```

O script usa automaticamente o arquivo de `imagens_originais/` quando ele existe, e a nota de rodapé
"gráfico redesenhado" deixa de aparecer nesse slide.

## Ponto a revisar no notebook

Na função `calcular_geometria_amarracao`, os cabeços ficam em `y = 3,5 m`, do lado do navio em relação
à face do cais, e 7 das 12 buzinas usadas têm `Y_local` negativo (costado oposto ao cais na atracação
por boreste). Vale conferir essas duas convenções antes de usar comprimentos e ângulos como resultado
final.

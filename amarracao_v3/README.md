# Apresentação v3 — Sistema de amarração (estilo PINN)

`Apresentacao_Amarracao_v3.pptx` · 13 slides · 16:9

Cada slide de gráfico traz a imagem à esquerda e dois cards à direita:
**O que o gráfico lê** (com o intuito do gráfico) e **Resultado obtido** (com um número em destaque).
Os números e as leituras seguem as saídas do notebook *TF4.0 - Colab* (calado 9 m, 394 condições).

## Imagens

Os gráficos em `imagens_recriadas/` foram redesenhados a partir das saídas do notebook:

| Arquivo | Dados |
|---|---|
| `00_plano_amarracao.png` | exatos (coordenadas de buzinas e cabeços do notebook) |
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

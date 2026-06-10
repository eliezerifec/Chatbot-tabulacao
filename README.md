# Codificador IFec

App em produção: [codificador.streamlit.app](https://codificador.streamlit.app/)

## Modos de uso

### Pipeline completo (novo)
Automatiza o fluxo inteiro: base bruta do SurveyMonkey → base final + tabulação.

1. **Upload** — base exportada do SurveyMonkey (.xlsx/.csv, cabeçalho em 2 linhas)
   e o questionário em Word (.docx). O questionário deve descrever os pulos
   (ex.: "Se NÃO, pule para a P10", "AGRADEÇA E ENCERRE A PESQUISA").
2. **Limpeza** *(checkpoint)* — a IA lê o questionário, extrai as regras de pulo
   e mostra quantas violações cada uma encontra. Você aprova/desativa regras
   antes de aplicar. Também remove linhas vazias e respondentes duplicados.
   Gera relatório de tudo que foi alterado.
3. **Codificação** — as colunas abertas ("Outro. Qual?", subcampos "1:/2:/3:")
   são detectadas automaticamente; subcolunas da mesma pergunta recebem o mesmo
   conjunto de categorias.
4. **Categorias** *(checkpoint)* — revise, renomeie ou una as categorias criadas
   pela IA antes do merge.
5. **Base final** — as colunas codificadas (`_cod`) entram ao lado das originais;
   baixe a base final, o relatório de limpeza e a tabulação em Excel + PowerPoint.

### Codificador
Fluxo original: arquivo com uma aba por pergunta aberta → codificação IA →
base codificada.

### Tabulação automática
Tabulação direta de uma base (sem codificar): detecta perguntas RU/RM/Grid/
Aberta/Média/NPS e gera Excel + PowerPoint.

## Arquivos principais

| Arquivo | Papel |
| --- | --- |
| `streamlit_app.py` | App Streamlit (entrada do deploy) e telas Codificador/Tabulação |
| `tela_pipeline.py` | Tela do Pipeline completo (5 etapas) |
| `limpeza.py` | Extração de regras de pulo do questionário (IA) e motor de limpeza |
| `codificador.py` | Motor de codificação com IA (2 agentes via OpenAI) |
| `tabulador.py` | Detecção de perguntas e tabulação (Excel) |
| `gerador_ppt.py` | Geração do PowerPoint no template Sesc/IFec |

## Configuração

- `OPENAI_API_KEY` — obrigatória para limpeza por questionário e codificação.
  Local: arquivo `.env` na pasta do projeto. Deploy: Secrets do Streamlit.
- `TEMPLATE_PPT_PATH` — opcional, caminho de um template .pptx alternativo.

Instalação local: `pip install -r requirements.txt` e `streamlit run streamlit_app.py`.

# Codificador de Pesquisas com IA

Aplicação web em **Streamlit** para codificação automática de respostas abertas com IA, tabulação e geração de Excel/PowerPoint.

## Pre-requisitos

- Python 3.10+
- Uma chave da OpenAI em `OPENAI_API_KEY`

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

Para rodar localmente, crie um arquivo `.env` com:

```env
OPENAI_API_KEY=sk-...
```

## Publicar no Streamlit Cloud via GitHub

1. Suba este repositorio para o GitHub.
2. No Streamlit Cloud, crie um novo app apontando para este repositorio.
3. Em **Main file path**, use `app.py`.
4. Em **Secrets**, adicione:

```toml
OPENAI_API_KEY = "sk-..."
```

5. Clique em deploy.

## Como usar

O app tem duas areas principais:

- **COD - Codificação**: codifica respostas abertas com IA e exporta a base codificada.
- **TAB - Tabulação**: detecta perguntas, permite revisar tipo/ativação e gera Excel + PowerPoint.

### Codificação

1. Envie a planilha `.xlsx` ou `.csv`.
2. Escolha a coluna com as respostas abertas.
3. Configure o tipo da pergunta, modo de resposta e coluna de saida.
4. Opcionalmente, carregue uma pesquisa anterior para reaproveitar categorias.
5. Escreva o contexto da pergunta para guiar o modelo.
6. Clique em **Iniciar codificacao**.
7. Baixe o Excel codificado.

### Tabulação

1. Abra a aba **TAB - Tabulação**.
2. Escolha se quer tabular o arquivo enviado ou o resultado codificado.
3. Clique em **Detectar perguntas**.
4. Revise quais perguntas ficam ativas e ajuste o tipo quando necessário.
5. Gere o Excel de tabulação e/ou o PowerPoint.

## Configuracoes avancadas

Os modelos usados ficam em `codificador.py`:

```python
MODELO_AGENTE1 = "gpt-5.4"
MODELO_AGENTE2 = "gpt-4o"
```

## Dicas

- Quanto mais contexto voce fornecer, mais consistente sera a codificacao.
- O arquivo `.env` nao deve ser enviado ao GitHub.
- No Streamlit Cloud, use sempre **Secrets** para configurar a chave da OpenAI.

## Problemas comuns

**Chave OpenAI nao encontrada**

Configure `OPENAI_API_KEY` nos Secrets do Streamlit Cloud ou no arquivo `.env` local.

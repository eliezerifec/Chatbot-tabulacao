# Codificador de Pesquisas com IA

Aplicacao web em **Streamlit** para codificacao automatica de respostas abertas com IA.

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

1. Envie a planilha `.xlsx` ou `.csv`.
2. Escolha a coluna com as respostas abertas.
3. Configure o tipo da pergunta, modo de resposta e coluna de saida.
4. Opcionalmente, carregue uma pesquisa anterior para reaproveitar categorias.
5. Escreva o contexto da pergunta para guiar o modelo.
6. Clique em **Iniciar codificacao**.
7. Baixe o Excel codificado.

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

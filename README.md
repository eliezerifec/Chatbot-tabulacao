# Codificador de Pesquisas com IA

Plataforma de codificacao automatica de respostas abertas com interface web em Streamlit.

## Pre-requisitos

### 1. Python 3.10+

Download: https://python.org

### 2. Ollama

Download: https://ollama.com

Apos instalar, abra o terminal e execute:

```bash
# Baixar o modelo Llama 3 (uma vez so, ~4GB)
ollama pull llama3

# Iniciar o servidor
ollama serve
```

## Instalacao

```bash
# 1. Clone ou baixe este projeto

# 2. Instale as dependencias
pip install -r requirements.txt

# 3. Execute a aplicacao web
streamlit run app.py
```

### Streamlit Cloud

Se for publicar no Streamlit Cloud, adicione este segredo em `App settings > Secrets`:

```toml
OPENAI_API_KEY = "sk-..."
```

## Como usar

1. Abra o app Streamlit e envie sua planilha `.xlsx` ou `.csv`
2. Escolha a coluna com as respostas abertas
3. Escolha a coluna de saida ou use uma nova coluna
4. Opcionalmente, importe codificacoes anteriores em `.json` ou `.xlsx`
5. Opcionalmente, adicione categorias separadas por virgula
6. Escreva o contexto da pergunta para guiar o modelo
7. Clique em `Iniciar codificacao`
8. Exporte o resultado ao final

## Formato dos arquivos de codigo

### JSON

```json
{
  "alegre": "alegria",
  "feliz": "alegria",
  "muito bom": "positivo",
  "otimo": "positivo",
  "entediante": "negativo"
}
```

### Excel (.xlsx)

| resposta | categoria |
|----------|-----------|
| alegre   | alegria   |
| feliz    | alegria   |
| otimo    | positivo  |

## Configuracoes avancadas

No arquivo `codificador.py`:

```python
MODELO = "llama3"
OLLAMA_URL = "http://localhost:11434/api/generate"
```

## Dicas

- Quanto mais exemplos anteriores voce fornecer, mais consistente sera a codificacao
- O sistema usa cache automatico, entao respostas identicas nao chamam o modelo novamente
- Descreva bem o contexto da pergunta para respostas mais precisas

## Problemas comuns

**"Ollama nao esta rodando"**

Execute `ollama serve` no terminal antes de abrir o app.

**"Modelo nao encontrado"**

Execute `ollama pull llama3` no terminal.

**"Streamlit nao encontrado"**

Instale as dependencias com `pip install -r requirements.txt`.

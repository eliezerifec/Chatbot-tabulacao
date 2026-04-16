# Codificador de Pesquisas com IA 🤖

Plataforma de codificação automática de respostas abertas usando **Llama 3** rodando **100% localmente** via Ollama.

---

## 📋 Pré-requisitos

### 1. Python 3.10+
Download: https://python.org

### 2. Ollama (roda o modelo local)
Download: https://ollama.com

Após instalar, abra o terminal e execute:
```bash
# Baixar o modelo Llama 3 (uma vez só, ~4GB)
ollama pull llama3

# Iniciar o servidor (deixe rodando em segundo plano)
ollama serve
```

---

## 🚀 Instalação

```bash
# 1. Clone ou baixe este projeto

# 2. Instale as dependências
pip install -r requirements.txt

# 3. Execute a aplicação
python app.py
```

---

## 🎯 Como usar

### Passo a passo:

1. **Clique em "Abrir Planilha"** → selecione seu `.xlsx` ou `.csv`
2. **Escolha a coluna** com as respostas abertas
3. **Escolha a coluna de saída** (ou crie uma nova)
4. **(Opcional) Carregue codificações anteriores** → arquivo `.json` ou `.xlsx`
   - JSON: `{"alegre": "alegria", "muito bom": "positivo"}`
   - Excel: duas colunas — `resposta` | `categoria`
5. **(Opcional) Adicione categorias** separadas por vírgula
6. **Escreva o contexto** da pergunta para guiar o modelo
7. **Clique em "Iniciar Codificação"**
8. **Exporte** o resultado quando terminar

---

## 📁 Formato dos arquivos de código

### JSON (recomendado para reuso):
```json
{
  "alegre": "alegria",
  "feliz": "alegria",
  "muito bom": "positivo",
  "ótimo": "positivo",
  "entediante": "negativo"
}
```

### Excel (.xlsx):
| resposta  | categoria |
|-----------|-----------|
| alegre    | alegria   |
| feliz     | alegria   |
| ótimo     | positivo  |

---

## ⚙️ Configurações avançadas

No arquivo `codificador.py`:

```python
MODELO = "llama3"        # Troque por: llama3:8b, mistral, gemma3, etc.
OLLAMA_URL = "http://localhost:11434/api/generate"
```

---

## 🔧 Modelos alternativos (todos gratuitos)

| Modelo   | Tamanho | Velocidade | Qualidade |
|----------|---------|------------|-----------|
| llama3   | 4.7GB   | Média      | ⭐⭐⭐⭐⭐    |
| mistral  | 4.1GB   | Rápida     | ⭐⭐⭐⭐     |
| gemma3   | 5.2GB   | Média      | ⭐⭐⭐⭐⭐    |
| phi3     | 2.3GB   | Muito rápida | ⭐⭐⭐   |

---

## 💡 Dicas

- **Quanto mais exemplos anteriores você fornecer**, mais consistente será a codificação
- O sistema usa **cache automático** — respostas idênticas não chamam o modelo de novo
- **Descreva bem o contexto** da pergunta para respostas mais precisas
- Para grandes volumes (1000+ respostas), prefira modelos menores como `phi3` para velocidade

---

## 🐛 Problemas comuns

**"Ollama não está rodando"**
→ Execute `ollama serve` no terminal antes de abrir o app

**"Modelo não encontrado"**
→ Execute `ollama pull llama3` no terminal

**Tkinter não instalado**
→ `sudo apt-get install python3-tk` (Linux) ou reinstale Python com tcl/tk (Windows)

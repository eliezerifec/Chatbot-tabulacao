"""
Motor de Codificação com IA — Fluxo de 2 Agentes via OpenAI
-------------------------------------------------------------
Agente 1 (Categorizador): lê TODAS as respostas e cria as categorias
Agente 2 (Classificador): associa cada resposta a uma categoria

Não depende de crewai. Usa apenas: openai, pandas, requests
"""

import json
import os
import re
import random
from difflib import SequenceMatcher
from openai import OpenAI
from aprendizado import BancoAprendizado
from pathlib import Path
from biblioteca_codificacao import BibliotecaCodificacao

# ── Configuração ──────────────────────────────────────────────────────────────
# Lê do arquivo .env (nunca sobe para o GitHub)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # sem dotenv, lê das variáveis de ambiente do sistema

API_KEY = os.getenv("OPENAI_API_KEY", "")
if not API_KEY:
    raise RuntimeError(
        "Chave OpenAI não encontrada.\n"
        "Crie um arquivo .env na pasta do projeto com:\n"
        "  OPENAI_API_KEY=sk-..."
    )
DICIONARIO_CODIFICACAO_PATH = os.getenv(
    "DICIONARIO_CODIFICACAO_PATH",
    str(Path(__file__).resolve().parent.parent / "Dicionário" / "Dicionario.xlsx")
)

# ── Modelos por agente ────────────────────────────────────────────────────────
# Agente 1 (Categorizador): cria as categorias — modelo mais inteligente
MODELO_AGENTE1 = "gpt-5.4"
# Agente 2 (Classificador): vincula respostas às categorias — mais rápido e barato
MODELO_AGENTE2 = "gpt-4o"

# ── Modos de resposta (estrutura da resposta) ────────────────────────────────
# Independente do tipo semântico, a resposta pode ter estrutura diferente
MODOS_RESPOSTA = {
    "simples": {
        "label": "Simples",
        "descricao": "Uma resposta → uma categoria",
    },
    "multipla": {
        "label": "Múltipla",
        "descricao": "Separa por ', ' e codifica cada parte individualmente",
    },
    "semiaberta_simples": {
        "label": "Semiaberta — Simples",
        "descricao": "Categoriza em predefinidas (imputação) ou cria nova coluna",
    },
    "semiaberta_multipla": {
        "label": "Semiaberta — Múltipla",
        "descricao": "Separa por ', ' + categoriza em predefinidas ou cria nova",
    },
}

# ── Tipos de pergunta predefinidos ────────────────────────────────────────────
TIPOS_PERGUNTA = {
    "reconhecimento_marca": {
        "label": "🏷  Reconhecimento de Marca",
        "descricao": "Quais marcas o participante lembrou de ter visto",
        "instrucoes": """Você está codificando respostas de uma pergunta de reconhecimento de marca espontânea.
O participante respondeu quais marcas lembrou de ter visto em um evento.

REGRAS OBRIGATÓRIAS:
1. Extraia APENAS os nomes das marcas mencionadas
2. Se houver mais de uma marca, separe com ", " (vírgula espaço)
3. Normalize o nome: capitalize corretamente (ex: "coca cola" → "Coca-Cola")
4. Ignore palavras que não são marcas (ex: "não lembro", "nenhuma")
5. Se não houver marca identificável, use a categoria: SEM_MARCA
6. Cada categoria deve ser o nome da marca normalizado

Exemplos de classificação:
  "vi a coca cola e a pepsi"  →  Coca-Cola, Pepsi
  "brahma"                    →  Brahma
  "não lembro de nenhuma"     →  SEM_MARCA
  "Nike e Adidas estavam lá"  →  Nike, Adidas""",
    },

    "satisfacao": {
        "label": "😊  Satisfação / Motivo",
        "descricao": "Por que o participante gostou do evento",
        "instrucoes": """Você está codificando respostas abertas de satisfação de evento.
O participante explicou por que gostou (ou não gostou) do evento.

REGRAS OBRIGATÓRIAS:
1. Crie UMA categoria temática que resuma a resposta
2. A categoria deve ter NO MÁXIMO 3 palavras, ser uma frase curta e direta
3. Use substantivos/adjetivos descritivos (ex: "boa organização", "atrações diversas", "atendimento ruim")
4. Agrupe respostas com o MESMO TEMA na mesma categoria
5. Seja consistente: respostas similares = mesma categoria

Exemplos de classificação:
  "adorei a organização do evento, tudo muito bem feito"  →  boa organização
  "as atrações foram incríveis"                           →  atrações diversas
  "o atendimento foi péssimo"                             →  atendimento ruim
  "gostei muito da música ao vivo"                        →  música ao vivo
  "estava muito cheio e desorganizado"                    →  superlotação desorganizada""",
    },

    "definicao_palavra": {
        "label": "💬  Definição em Uma Palavra",
        "descricao": "Uma palavra que define a experiência",
        "instrucoes": """Você está codificando respostas de uma pergunta "defina em uma palavra".
O participante escolheu uma palavra para descrever sua experiência.

REGRAS OBRIGATÓRIAS:
1. Normalize a palavra: corrija grafia, capitalize a primeira letra
2. Agrupe palavras com o MESMO significado ou raiz em uma categoria única:
   - "lindo", "linda", "lindíssimo" → Lindo
   - "ótimo", "ótima", "otimo" → Ótimo
   - "incrível", "incrivel", "incredivel" → Incrível
3. Use sempre o masculino singular como forma canônica
4. Se a resposta tiver mais de uma palavra, use apenas a mais relevante

Exemplos de classificação:
  "linda"       →  Lindo
  "INCRIVEL"    →  Incrível
  "muito bom"   →  Bom
  "maravilhoso" →  Maravilhoso
  "otimo"       →  Ótimo""",
    },

    "local_moradia": {
        "label": "📍  Local de Moradia",
        "descricao": "Cidade, estado ou país onde mora",
        "instrucoes": """Você está codificando respostas de uma pergunta sobre local de moradia.

REGRAS OBRIGATÓRIAS:
1. Extraia APENAS o nome do estado ou país
2. Se a pessoa mencionou cidade, retorne o ESTADO correspondente
3. Use o nome completo do estado (ex: "SP" → "São Paulo")
4. Se for fora do Brasil, retorne o nome do PAÍS
5. Normalize a grafia: capitalize corretamente
6. Se não for possível identificar, use a categoria: NÃO IDENTIFICADO

Exemplos de classificação:
  "moro em São Paulo capital"  →  São Paulo
  "Rio de Janeiro, Copacabana" →  Rio de Janeiro
  "sou de BH"                  →  Minas Gerais
  "moro em SP"                 →  São Paulo
  "Argentina"                  →  Argentina
  "não sei"                    →  NÃO IDENTIFICADO""",
    },

    "livre": {
        "label": "✏️  Personalizado",
        "descricao": "Usar contexto personalizado que você escrever",
        "instrucoes": None,
    },
}


def _formatar_few_shot(exemplos: list) -> str:
    if not exemplos:
        return ""

    corrigidos = [e for e in exemplos if not e.get("correto", True)]
    aprovados = [e for e in exemplos if e.get("correto", True)]

    linhas = []

    if corrigidos:
        linhas.append("\nExemplos onde a IA ERROU — aprenda com estes casos:")
        for e in corrigidos:
            cat_ia = e.get("categoria_ia", "?")
            cat_ok = e["categoria"]
            linhas.append(
                f'  "{e["resposta"]}"'
                f'  [IA disse: {cat_ia}]  →  CORRETO: {cat_ok}'
            )

    if aprovados:
        linhas.append("\nExemplos validados pelos pesquisadores:")
        for e in aprovados:
            linhas.append(f'  "{e["resposta"]}"  →  {e["categoria"]}')

    return "\n".join(linhas) + "\n"


def _formatar_few_shot_biblioteca(exemplos: list) -> str:
    if not exemplos:
        return ""

    linhas = [
        "\nBiblioteca histórica de codificação — use como referência principal de estilo e granularidade:"
    ]

    for e in exemplos:
        pergunta = e.get("pergunta_texto", "").strip()
        if pergunta:
            linhas.append(f"  [Pergunta] {pergunta}")
        linhas.append(f'  "{e["resposta"]}"  →  {e["categoria"]}')

    return "\n".join(linhas) + "\n"


class CodificadorIA:
    
    def __init__(self):
        self.codigos_base: dict = {}
        self.categorias: list = []
        self._cache: dict = {}
        self._client = None
        self.banco = BancoAprendizado()
        self.biblioteca = BibliotecaCodificacao(DICIONARIO_CODIFICACAO_PATH)


    # ── Cliente OpenAI (criado uma vez) ──────────────────────────────────────

    def _get_client(self) -> OpenAI:
        if self._client is None:
            self._client = OpenAI(api_key=API_KEY)
        return self._client

    # Modelos de raciocínio (o1, o3, o4-*) não aceitam temperature nem max_tokens
    _MODELOS_RACIOCINIO = {"o1", "o1-mini", "o1-preview", "o3", "o3-mini", "o4-mini", "gpt-5", "gpt-5."}

    def _chamar_gpt(self, system: str, user: str, max_tokens: int = 2000,
                    modelo: str = None) -> str:
        """Chamada à API da OpenAI — compatível com modelos GPT e de raciocínio."""
        modelo = modelo or MODELO_AGENTE1
        eh_raciocinio = any(modelo.startswith(m) for m in self._MODELOS_RACIOCINIO)

        params = {
            "model": modelo,
            "messages": [
                {"role": "user", "content": f"{system}\n\n{user}"}
                if eh_raciocinio else
                {"role": "system", "content": system},
            ],
        }

        if not eh_raciocinio:
            params["messages"] = [
                {"role": "system", "content": system},
                {"role": "user",   "content": user},
            ]
            params["temperature"]  = 0.1
            params["max_tokens"]   = max_tokens
        else:
            # Modelos de raciocínio: sem system role, sem temperature
            params["messages"] = [
                {"role": "user", "content": f"{system}\n\n{user}"},
            ]
            params["max_completion_tokens"] = max_tokens

        resp = self._get_client().chat.completions.create(**params)
        return resp.choices[0].message.content.strip()

    # ── Alimentação ──────────────────────────────────────────────────────────
    def _obter_few_shot_completo(self, tipo: str, modo: str, pergunta_texto: str = "") -> str:
        exemplos_banco = self.banco.buscar_exemplos(tipo, n=12)
        exemplos_biblioteca = []

        if hasattr(self, "biblioteca") and self.biblioteca and self.biblioteca.ok:
            exemplos_biblioteca = self.biblioteca.buscar_exemplos(
                tipo=tipo,
                modo=modo,
                pergunta_texto=pergunta_texto,
                n=15,
            )

        few_banco = _formatar_few_shot(exemplos_banco)
        few_biblioteca = _formatar_few_shot_biblioteca(exemplos_biblioteca)
        return few_banco + few_biblioteca

    def sugerir_categorias_da_biblioteca(self, tipo: str, modo: str, pergunta_texto: str = "") -> list[str]:
        if not self.biblioteca or not self.biblioteca.ok:
            return []

        return self.biblioteca.listar_categorias_relacionadas(
            tipo=tipo,
            modo=modo,
            pergunta_texto=pergunta_texto,
            n=30,
        )


    def carregar_codigos(self, dados: dict):
        self.codigos_base.update(dados)

    def adicionar_categoria(self, categoria: str):
        cat = categoria.strip()
        if cat and cat not in self.categorias:
            self.categorias.append(cat)

    # ── Codificação linha a linha (compatibilidade com a UI) ─────────────────


    def codificar_lote_modo(self, respostas: list, tipo: str = "livre",
                            modo: str = "simples",
                            contexto_custom: str = "",
                            categorias_imputacao: list = None,
                            categorias_anteriores: list = None,
                            callback_progresso=None) -> dict:
        """
        Codifica um lote respeitando o modo de resposta.

        categorias_anteriores: lista de categorias de uma pesquisa já realizada.
            Quando fornecida, o Agente 1 é instruído a reutilizá-las
            obrigatoriamente e só criar nova categoria em último caso absoluto.

        Retorna dict com chaves dependendo do modo:
          simples/livre   → {"resultado": [str, ...]}
          multipla        → {"resultado": [str, ...]}  (células com "A, B, C")
          semiaberta_*    → {"imputado": [str, ...], "novo": [str, ...]}
        """
        categorias_imputacao  = categorias_imputacao or []
        categorias_anteriores = categorias_anteriores or []

        # ── Múltipla: explode por ", ", codifica cada parte, reagrupa ─────────
        if modo == "multipla":
            # Expande respostas com múltiplos valores
            expandidas = []
            mapa = []  # (idx_original, parte_idx)
            for i, r in enumerate(respostas):
                partes = [p.strip() for p in str(r).split(",") if p.strip()
                          and str(r).lower() not in ("nan", "none", "")]
                if not partes:
                    partes = [str(r)]
                for j, p in enumerate(partes):
                    expandidas.append(p)
                    mapa.append(i)

            if not expandidas:
                return {"resultado": ["SEM_RESPOSTA"] * len(respostas)}

            # Codifica as partes expandidas
            cats_expandidas = self.codificar_lote(
                expandidas, tipo=tipo,
                contexto_custom=contexto_custom,
                categorias_anteriores=categorias_anteriores,
                callback_progresso=None)

            # Reagrupa por índice original
            grupos = [[] for _ in respostas]
            for k, cat in enumerate(cats_expandidas):
                grupos[mapa[k]].append(cat)

            resultado = [", ".join(dict.fromkeys(g)) if g else "SEM_RESPOSTA"
                         for g in grupos]

            # Chama callback manualmente
            if callback_progresso:
                for i, (r, c) in enumerate(zip(respostas, resultado)):
                    callback_progresso(i, len(respostas), str(r), c)

            return {"resultado": resultado}

        # ── Semiaberta: o modelo decide se encaixa ou cria nova ────────────────
        if "semi" in modo:
            resultado_semi = self._codificar_semiaberta(
                respostas, tipo=tipo,
                modo=modo,
                contexto_custom=contexto_custom,
                categorias_imputacao=categorias_imputacao,
                categorias_anteriores=categorias_anteriores,
                callback_progresso=callback_progresso)
            return resultado_semi

        # ── Simples (padrão) ──────────────────────────────────────────────────
        resultado = self.codificar_lote(
            respostas, tipo=tipo,
            contexto_custom=contexto_custom,
            categorias_anteriores=categorias_anteriores,
            callback_progresso=callback_progresso)
        return {"resultado": resultado}

    def _codificar_semiaberta(self, respostas: list, tipo: str,
                              modo: str, contexto_custom: str,
                              categorias_imputacao: list,
                              categorias_anteriores: list = None,
                              callback_progresso=None) -> dict:
        """
        Semiaberta com 2 agentes conscientes das categorias fornecidas:

        Agente 1 — vê as categorias pré-definidas + todas as respostas e decide:
                   a) quais respostas encaixam em categoria existente
                   b) quais precisam de categoria nova (e define o nome)

        Agente 2 — aplica a decisão do Agente 1 em cada resposta individualmente

        Retorna {"imputado": [...], "novo": [...]}
          imputado = encaixou em categoria pré-definida (nome exato da categoria)
          novo     = categoria nova criada pela IA
          Respostas sem encaixe E sem sentido → imputado="", novo=""
        """
        import random as _random

        respostas_str   = [str(r).strip() for r in respostas]
        indices_validos = [i for i, r in enumerate(respostas_str)
                           if r and r.lower() not in ("nan", "none", "")]
        respostas_validas = [respostas_str[i] for i in indices_validos]

        col_imputado = [""] * len(respostas_str)
        col_novo     = [""] * len(respostas_str)

        if not respostas_validas:
            return {"imputado": col_imputado, "novo": col_novo}

        config     = TIPOS_PERGUNTA.get(tipo, TIPOS_PERGUNTA["livre"])
        instrucoes = config["instrucoes"] or contexto_custom or ""

        exemplos_banco = self.banco.buscar_exemplos(tipo, n=20)
        few_shot = _formatar_few_shot(exemplos_banco)

        # ── Bloco de pesquisa anterior (semiaberta) ───────────────────────────
        categorias_anteriores = categorias_anteriores or []
        if categorias_anteriores:
            ant_lista = "\n".join(f"  - {c}" for c in categorias_anteriores)
            bloco_anterior = f"""
╔══════════════════════════════════════════════════════════════════════╗
║  CATEGORIAS DA PESQUISA ANTERIOR — PRIORIDADE MÁXIMA                ║
╚══════════════════════════════════════════════════════════════════════╝
Estas categorias foram validadas em uma rodada anterior da MESMA pesquisa.
Você DEVE encaixar cada resposta em uma delas sempre que houver qualquer
compatibilidade semântica — mesmo que a correspondência não seja perfeita.
Só crie categoria nova se a resposta expressar uma dimensão completamente
ausente nesta lista (caso EXTREMAMENTE raro).

{ant_lista}
"""
        else:
            bloco_anterior = ""

        cats_lista = "\n".join(f"  - {c}" for c in categorias_imputacao) \
                     if categorias_imputacao else "  (nenhuma categoria pré-definida)"

        # Para múltipla, expande antes de classificar
        multipla = "multipla" in modo
        if multipla:
            expandidas = []
            mapa_expand = []
            for i, r in enumerate(respostas_validas):
                partes = [p.strip() for p in r.split(",") if p.strip()]
                if not partes:
                    partes = [r]
                for p in partes:
                    expandidas.append(p)
                    mapa_expand.append(i)
            respostas_para_classificar = expandidas
        else:
            respostas_para_classificar = respostas_validas
            mapa_expand = list(range(len(respostas_validas)))

        todas_enumeradas = "\n".join(
            f"{i+1}. {r}" for i, r in enumerate(respostas_para_classificar))

        # ── Agente 1: define o mapeamento resposta → decisão ─────────────────
        system_a1 = (
            "Você é um especialista em análise qualitativa de pesquisas qualitativas brasileiras. "
            "Você entende intenção por trás de respostas abertas e sabe quando uma resposta "
            "se encaixa semanticamente em uma categoria mesmo que as palavras sejam diferentes. "
            "Responda SEMPRE em JSON válido, sem texto fora do JSON."
        )

        user_a1 = f"""Contexto da pesquisa:
{instrucoes}
{few_shot}
{bloco_anterior}
═══════════════════════════════════════════════════
CATEGORIAS PRÉ-DEFINIDAS — leia com atenção:
{cats_lista}
═══════════════════════════════════════════════════

COMO DECIDIR (siga esta ordem de raciocínio para cada resposta):

PASSO 1 — Tente encaixar em uma categoria pré-definida.
  Encaixe por INTENÇÃO e SEMÂNTICA, não só por palavras iguais.
  Exemplos de encaixe correto (validados pelos pesquisadores):
  - "Por conta do meu trabalho" → "Queria complementar minha formação (adquirir novos conhecimentos)" (exigência profissional = aprendizado)
  - "Soft skill de comunicação" → "Queria complementar minha formação (adquirir novos conhecimentos)" (desenvolvimento = aprendizado)
  - "Jovem aprendiz" → "Queria entrar no mercado de trabalho" (inserção no mercado)
  - "Atualização de currículo" → "Queria entrar no mercado de trabalho" (busca de emprego)
  - "Banho e tosa" → "Procurava um hobby" (atividade por interesse pessoal)
  - "Eu já desenho e queria melhorar" → "Procurava um hobby" (aperfeiçoamento de hobby)

PASSO 2 — SÓ crie categoria nova se a resposta expressar uma intenção que GENUINAMENTE
  não tem equivalente em nenhuma das categorias pré-definidas.
  Exemplos que SÃO categorias novas legítimas (validados pelos pesquisadores):
  - "Segunda renda" → "Complementar renda" (renda extra é diferente de entrar no mercado de trabalho)
  - "Obrigação escolar / trabalho acadêmico" — não é nenhuma das intenções pré-definidas
  - "Indicação médica / terapia ocupacional" — fora do escopo das categorias

PASSO 3 — Use null apenas para respostas vazias, ininteligíveis ou "não sei".

Respostas para classificar ({len(respostas_para_classificar)} no total):
{todas_enumeradas}

Retorne SOMENTE este JSON:
{{"classificacoes": [
  {{"indice": 1, "categoria": "nome EXATO da categoria pré-definida", "nova": false}},
  {{"indice": 2, "categoria": "nome da categoria nova criada", "nova": true}},
  {{"indice": 3, "categoria": null, "nova": false}}
]}}

REGRAS FINAIS:
- "nova": false → use o nome EXATAMENTE como está nas categorias pré-definidas
- "nova": true  → categoria nova, 2 a 5 palavras, objetiva
- Classifique TODAS as {len(respostas_para_classificar)} respostas
- Na dúvida entre encaixar ou criar nova → ENCAIXE na mais próxima"""

        texto_a1 = self._chamar_gpt(system_a1, user_a1,
                                    max_tokens=4000, modelo=MODELO_AGENTE1)

        # Parseia resultado do Agente 1
        import re as _re
        import json as _json

        decisoes = {}  # {indice_1based: {"categoria": str|None, "nova": bool}}
        try:
            txt = _re.sub(r"```[a-z]*", "", texto_a1).strip("`").strip()
            m   = _re.search(r'\{.*"classificacoes".*\}', txt, _re.DOTALL)
            dados = _json.loads(m.group() if m else txt)
            for item in dados.get("classificacoes", []):
                decisoes[item["indice"]] = {
                    "categoria": item.get("categoria"),
                    "nova":      item.get("nova", False)
                }
        except Exception:
            pass

        # ── Monta resultados ──────────────────────────────────────────────────
        # Para múltipla: agrupa de volta por resposta original
        if multipla:
            grupos_imp = [[] for _ in respostas_validas]
            grupos_nov = [[] for _ in respostas_validas]
            for k, idx_orig in enumerate(mapa_expand):
                dec = decisoes.get(k + 1, {})
                cat = dec.get("categoria")
                nova = dec.get("nova", False)
                if cat:
                    cat = self._capitalizar(cat)
                    if nova:
                        grupos_nov[idx_orig].append(cat)
                    else:
                        grupos_imp[idx_orig].append(cat)
            for i, idx_orig in enumerate(indices_validos):
                col_imputado[idx_orig] = ", ".join(dict.fromkeys(grupos_imp[i])) if grupos_imp[i] else ""
                col_novo[idx_orig]     = ", ".join(dict.fromkeys(grupos_nov[i])) if grupos_nov[i] else ""
        else:
            for i, idx_orig in enumerate(indices_validos):
                dec = decisoes.get(i + 1, {})
                cat = dec.get("categoria")
                nova = dec.get("nova", False)
                if cat:
                    cat = self._capitalizar(cat)
                    if nova:
                        col_novo[idx_orig] = cat
                    else:
                        col_imputado[idx_orig] = cat

        if callback_progresso:
            for i, idx_orig in enumerate(indices_validos):
                imp = col_imputado[idx_orig]
                nov = col_novo[idx_orig]
                resultado_str = f"IMP:{imp}" if imp else f"NOVO:{nov}" if nov else "—"
                callback_progresso(i, len(indices_validos),
                                   respostas_validas[i], resultado_str)

        return {"imputado": col_imputado, "novo": col_novo}

    def _vincular_com_lista_anterior(self, respostas: list, lista: list,
                                     instrucoes: str, few_shot: str,
                                     callback_progresso=None) -> list:
        """
        Vincula cada resposta ao item EXATO da lista anterior que melhor a representa.
        Não cria categorias novas — a lista é o universo completo de saídas.

        Funciona em lotes de 200 e retorna os nomes exatamente como estão na lista.
        """
        resultados = ["SEM_RESPOSTA"] * len(respostas)

        # Mapa case-insensitive para garantir nome exato na saída
        lista_lower = {c.lower(): c for c in lista}
        cats_formatada = "\n".join(f"  {i+1}. {c}" for i, c in enumerate(lista))

        system = (
            "Você é um especialista em vincular respostas abertas a categorias pré-definidas. "
            "Sua tarefa é encontrar, para cada resposta, qual categoria da lista melhor a representa — "
            "por significado, abreviação, sinônimo ou variação ortográfica. "
            "NUNCA invente categorias fora da lista. "
            "Responda SEMPRE em JSON válido, sem texto fora do JSON."
        )

        TAMANHO_LOTE = 200
        classificacoes = []

        for inicio in range(0, len(respostas), TAMANHO_LOTE):
            fim   = min(inicio + TAMANHO_LOTE, len(respostas))
            lote  = respostas[inicio:fim]
            lote_enum = "\n".join(f"  {inicio+i+1}. {r}" for i, r in enumerate(lote))

            user = f"""Contexto da pesquisa:
{instrucoes}
{few_shot}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LISTA FECHADA — ESTES SÃO OS ÚNICOS VALORES DE SAÍDA PERMITIDOS:
{cats_formatada}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

REGRAS DE VINCULAÇÃO:
1. Para cada resposta, encontre qual item da lista acima melhor a representa.
2. Considere: abreviações ("Facha" → "Faculdades Integradas Hélio Alves"),
   siglas ("CIEE" → "CIEE"), nomes parciais ("Celso Lisboa" → "Centro Universitário Celso Lisboa"),
   sinônimos e variações ortográficas.
3. Copie o nome do item da lista EXATAMENTE como está escrito — sem alterar uma letra.
4. Se a resposta for vazia, ilegível ou "não sei" → use "SEM_RESPOSTA".
5. Se genuinamente não tiver nenhum equivalente na lista → use "SEM_RESPOSTA".
6. NUNCA escreva um valor que não esteja na lista acima (exceto "SEM_RESPOSTA").

Respostas para vincular:
{lote_enum}

Retorne SOMENTE este JSON:
{{"classificacoes": [{{"indice": {inicio+1}, "categoria": "nome exato da lista"}}, ...]}}
Classifique TODAS as {len(lote)} respostas."""

            texto = self._chamar_gpt(system, user, max_tokens=4000,
                                     modelo=MODELO_AGENTE2)
            classificacoes += self._parsear_classificacoes(texto)

        # Aplica resultados — garante nome exato via lookup case-insensitive
        for item in classificacoes:
            idx = item.get("indice", 0) - 1   # 1-based → 0-based
            if 0 <= idx < len(respostas):
                cat_raw = (item.get("categoria") or "SEM_RESPOSTA").strip()
                # Corrige capitalização para o nome exato da lista
                cat = lista_lower.get(cat_raw.lower(), cat_raw)
                resultados[idx] = cat
                if callback_progresso:
                    callback_progresso(idx, len(respostas), respostas[idx], cat)

        return resultados

    def codificar(self, resposta: str, tipo: str = "livre", contexto_custom: str = "") -> str:
        """
        Codifica uma resposta individual.
        Usado como fallback — prefira codificar_lote() para processar abas inteiras.
        """
        resposta = str(resposta).strip()
        if not resposta or resposta.lower() in ("nan", "none", ""):
            return "SEM_RESPOSTA"

        chave = f"{tipo}::{resposta.lower()}"
        if chave in self._cache:
            return self._cache[chave]

        # Tenta match exato / fuzzy no histórico
        if tipo in ("livre", "satisfacao"):
            match = self.codigos_base.get(resposta.lower())
            if match:
                self._cache[chave] = match
                return match
            match_fuzzy = self._buscar_fuzzy(resposta)
            if match_fuzzy:
                self._cache[chave] = match_fuzzy
                return match_fuzzy

        resultado = self._classificar_individual(resposta, tipo, contexto_custom)
        self._cache[chave] = resultado
        return resultado


    # ── Tipos que NÃO precisam de categorização — só normalização ────────────
    # Para marcas e definição em 1 palavra, a própria resposta já é a categoria.
    # O modelo só precisa limpar, corrigir grafia e normalizar — nunca "SEM_RESPOSTA".
    TIPOS_NORMALIZACAO = {"reconhecimento_marca", "definicao_palavra"}

    def _normalizar_lote(self, respostas: list, tipo: str,
                         instrucoes: str, few_shot: str,
                         callback_progresso=None) -> list:
        """
        Normalização em 2 passos:
          Passo 1 — Agente 1 vê todos os valores únicos e monta um dicionário
                    de_para: {variacao: nome_oficial}
          Passo 2 — Aplica o dicionário mecanicamente em cada resposta,
                    sem pedir ao modelo para "interpretar" de novo.
        Isso garante agrupamento consistente e elimina cópia-e-cola.
        """
        import re as _re

        resultados_norm = ["SEM_RESPOSTA"] * len(respostas)

        # ── Passo 1: dicionário de padronização ───────────────────────────────
        # Trabalha só com valores únicos para economizar tokens
        unicos = sorted(set(r.strip() for r in respostas
                            if r.strip() and r.strip().lower() not in
                            ("nan", "none", "", "nenhuma", "não lembra",
                             "nao lembra", "não sei", "nao sei")))

        if not unicos:
            return resultados_norm

        lista_unicos = "\n".join(f"- {v}" for v in unicos)

        system_dict = (
            "Você é um especialista em padronização de nomes de instituições, "
            "marcas e entidades brasileiras. "
            "Seu trabalho é criar um dicionário de-para que converte variações "
            "informais/erradas para o nome oficial correto. "
            "Responda SEMPRE em JSON válido, sem texto fora do JSON."
        )

        user_dict = f"""Contexto da pesquisa:
{instrucoes}
{few_shot}
Abaixo estão todos os valores únicos encontrados nas respostas da pesquisa.
Muitos são a mesma entidade escrita de formas diferentes.

Valores únicos encontrados:
{lista_unicos}

Sua tarefa: para CADA valor, identifique a entidade real e defina o nome oficial padronizado.

REGRAS DE PADRONIZAÇÃO:
1. Siglas ficam em MAIÚSCULO: senai→SENAI, sesc→SESC, sesi→SESI, fgv→FGV,
   puc→PUC, ufrj→UFRJ, uerj→UERJ, ibmec→Ibmec, faetec→FAETEC, cefet→CEFET,
   uff→UFF, unirio→UNIRIO, ibm→IBM, espn→ESPN, ifb→IFB, iga→IGA
2. Nomes próprios: primeira letra maiúscula — "estacio"→"Estácio", "anhanguera"→"Anhanguera"
3. Agrupe variações da mesma entidade: "Ibemc", "Ibemec", "Ibmec" → todos viram "Ibmec"
4. Agrupe com e sem complemento: "Faculdade Unirio", "Unirio", "UNIRIO" → todos "UNIRIO"
5. Corrija acentos: "Estacio"→"Estácio", "Catolica"→"Católica", "Galpao"→"Galpão"
6. Corrija erros de grafia: "Jfrj"→"UFRJ" NÃO — "Jfrj" é JFRJ (Justiça Federal RJ)
7. NÃO invente nomes — se não reconhecer, mantenha capitalizado corretamente
8. Respostas como "Nenhuma", "Não lembra", "Não sei" → use "SEM_MARCA"

Retorne SOMENTE este JSON:
{{"dicionario": {{"valor_original": "nome_oficial", "outro_valor": "nome_oficial", ...}}}}

IMPORTANTE: inclua TODOS os {len(unicos)} valores únicos no dicionário."""

        texto_dict = self._chamar_gpt(system_dict, user_dict,
                                      max_tokens=3000, modelo=MODELO_AGENTE2)

        # Parseia o dicionário
        dicionario = {}
        try:
            texto_limpo = _re.sub(r"```[a-z]*", "", texto_dict).strip("`").strip()
            match = _re.search(r'\{.*"dicionario".*\}', texto_limpo, _re.DOTALL)
            dados = json.loads(match.group() if match else texto_limpo)
            dicionario = dados.get("dicionario", {})
        except Exception:
            pass

        # ── Passo 2: aplica o dicionário em cada resposta ─────────────────────
        for i, resp in enumerate(respostas):
            r = resp.strip()
            if not r or r.lower() in ("nan", "none", ""):
                resultados_norm[i] = "SEM_RESPOSTA"
            else:
                # Busca exata primeiro, depois case-insensitive
                norm = (dicionario.get(r) or
                        dicionario.get(r.lower()) or
                        next((v for k, v in dicionario.items()
                              if k.lower() == r.lower()), None))
                resultados_norm[i] = norm if norm else r.strip().title()

            if callback_progresso:
                callback_progresso(i, len(respostas), r, resultados_norm[i])

        return resultados_norm

    # ── Fluxo principal: 2 agentes em sequência ──────────────────────────────
    def codificar_lote(self, respostas: list, tipo: str = "livre",
                       contexto_custom: str = "",
                       categorias_anteriores: list = None,
                       callback_progresso=None) -> list:
        """
        Fluxo de 2 agentes:
          1. Agente Categorizador — lê todas as respostas e define as categorias
          2. Agente Classificador — associa cada resposta a uma categoria

        callback_progresso(i, total, resposta, categoria) chamado a cada classificação.
        Retorna lista de categorias na mesma ordem das respostas de entrada.
        """
        respostas_str   = [str(r).strip() for r in respostas]
        indices_validos = [i for i, r in enumerate(respostas_str)
                           if r and r.lower() not in ("nan", "none", "")]
        respostas_validas = [respostas_str[i] for i in indices_validos]
        resultados        = ["SEM_RESPOSTA"] * len(respostas_str)

        if not respostas_validas:
            return resultados

        config     = TIPOS_PERGUNTA.get(tipo, TIPOS_PERGUNTA["livre"])
        instrucoes = config["instrucoes"] or contexto_custom or \
                     "Categorize as respostas em categorias padronizadas e concisas."

        # Few-shot: exemplos validados por humanos no banco de aprendizado
        exemplos_banco = self.banco.buscar_exemplos(tipo, n=20)
        few_shot = _formatar_few_shot(exemplos_banco)

        # ── Tipos de normalização: agente único, sem lista fechada ────────────
        # Se há categorias_anteriores, NÃO normaliza — usa fluxo padrão com
        # a lista fechada, garantindo que os nomes exatos sejam preservados.
        if tipo in self.TIPOS_NORMALIZACAO and not categorias_anteriores:
            norm = self._normalizar_lote(respostas_validas, tipo, instrucoes,
                                         few_shot, callback_progresso)
            for i, idx_orig in enumerate(indices_validos):
                resultados[idx_orig] = norm[i]
            return resultados

        # ── Bloco de pesquisa anterior (sem lista anterior: comportamento normal) ─
        cats_hint = ""
        if self.categorias:
            cats_hint = f"\nCategorias já usadas (reutilize se adequado): {', '.join(self.categorias)}\n"
        regra_novas = (
            "- Crie QUANTAS categorias forem necessárias para cobrir todas as respostas\n"
            "- Prefira entre 10 e 30 categorias — use mais se a base for diversa\n"
            "- Cada categoria: 1 a 4 palavras, clara e objetiva\n"
            "- Cubra TODOS os temas — nenhuma resposta deve ficar sem encaixe"
        )

        todas_enumeradas = "\n".join(f"{i+1}. {r}" for i, r in enumerate(respostas_validas))

        # ── Com pesquisa anterior: vinculador dedicado (sem Agente 1) ─────────
        # Não usa o fluxo de 2 agentes — em vez disso, um único agente recebe
        # a lista fechada e vincula cada resposta ao item mais próximo dela.
        # Sem liberdade criativa: a lista é o universo completo de saídas.
        if categorias_anteriores:
            vinculados = self._vincular_com_lista_anterior(
                respostas_validas, categorias_anteriores,
                instrucoes, few_shot, callback_progresso)
            for i, idx_orig in enumerate(indices_validos):
                resultados[idx_orig] = vinculados[i]
            return resultados

        # ── Agente 1: Categorizador ───────────────────────────────────────────
        # Se há pesquisa anterior, as categorias já estão definidas — pula o Agente 1.
        # Qualquer chamada ao Agente 1 criaria categorias novas mesmo sendo instruído
        # a não fazer isso; a única garantia real é não chamá-lo.
        if categorias_anteriores:
            categorias_criadas = [self._capitalizar(c) for c in categorias_anteriores]
        else:
            system_cat = (
                "Você é um especialista em análise qualitativa de pesquisas. "
                "Seu trabalho é ler respostas abertas e definir um conjunto enxuto de categorias temáticas. "
                "Responda SEMPRE em JSON válido, sem texto fora do JSON."
            )
            user_cat = f"""Contexto da pesquisa:
{instrucoes}
{cats_hint}{few_shot}
Todas as respostas coletadas ({len(respostas_validas)} no total):
{todas_enumeradas}

Analise TODAS as respostas acima e defina as categorias necessárias para cobri-las.
Retorne SOMENTE este JSON (sem markdown, sem explicação):
{{"categorias": ["categoria1", "categoria2", ...]}}

Regras:
{regra_novas}
- NÃO crie a categoria "SEM_RESPOSTA" — ela só existe para respostas literalmente em branco"""

            texto_cats = self._chamar_gpt(system_cat, user_cat, max_tokens=1500,
                                             modelo=MODELO_AGENTE1)
            categorias_criadas = self._parsear_categorias(texto_cats)

            # Se o agente não retornou nada válido, usa categorias já conhecidas ou genérico
            if not categorias_criadas:
                categorias_criadas = self.categorias[:] or ["positivo", "negativo", "neutro", "SEM_RESPOSTA"]

        # Capitaliza todas as categorias antes de passar ao Agente 2
        categorias_criadas = [self._capitalizar(c) for c in categorias_criadas]

        # ── Agente 2: Classificador em lotes de 200 ──────────────────────────
        system_clf = (
            "Você é um classificador de texto preciso e consistente para análise qualitativa de pesquisas. "
            "Classifique cada resposta em exatamente uma das categorias fornecidas, "
            "seguindo rigorosamente as regras e exemplos do contexto. "
            "Responda SEMPRE em JSON válido, sem texto fora do JSON."
        )
        cats_lista = "\n".join(f"- {c}" for c in categorias_criadas)

        # Aviso extra para o Agente 2 quando há pesquisa anterior
        aviso_anterior_clf = ""
        if categorias_anteriores:
            aviso_anterior_clf = (
                "\n⚠️  ATENÇÃO — REGRA INVIOLÁVEL:\n"
                "As categorias listadas abaixo são os ÚNICOS valores permitidos.\n"
                "Você DEVE copiar o nome da categoria EXATAMENTE como está na lista —\n"
                "sem abreviar, sem reformular, sem omitir palavras.\n"
                "Exemplo: se a lista tem 'Centro Universitário Celso Lisboa',\n"
                "a resposta deve ser 'Centro Universitário Celso Lisboa', NÃO 'Celso Lisboa'.\n"
            )

        TAMANHO_LOTE = 200
        classificacoes = []

        for inicio in range(0, len(respostas_validas), TAMANHO_LOTE):
            fim = min(inicio + TAMANHO_LOTE, len(respostas_validas))
            lote = respostas_validas[inicio:fim]
            lote_enumerado = "\n".join(f"{inicio+i+1}. {r}" for i, r in enumerate(lote))

            user_clf = f"""Regras e contexto da codificação:
{instrucoes}
{few_shot}{aviso_anterior_clf}
Categorias disponíveis (copie o nome EXATAMENTE como está escrito abaixo):
{cats_lista}

Respostas para classificar:
{lote_enumerado}

Classifique CADA resposta em exatamente UMA categoria da lista acima.
Retorne SOMENTE este JSON (sem markdown, sem explicação):
{{"classificacoes": [{{"indice": 1, "categoria": "..."}}, {{"indice": 2, "categoria": "..."}}, ...]}}

Regras adicionais:
- Os indices devem começar em {inicio+1}
- Classifique TODAS as {len(lote)} respostas deste lote
- O valor de "categoria" deve ser COPIADO LITERALMENTE da lista — nenhuma alteração permitida
- É PROIBIDO usar "SEM_RESPOSTA" se a resposta tiver qualquer conteúdo reconhecível
- Se não se encaixar perfeitamente, escolha a categoria MAIS PRÓXIMA da lista
- "SEM_RESPOSTA" só para resposta completamente vazia, ilegível ou "não sei\""""

            texto_clf = self._chamar_gpt(system_clf, user_clf, max_tokens=4000,
                                             modelo=MODELO_AGENTE2)
            classificacoes += self._parsear_classificacoes(texto_clf)

        # ── Montar resultados na ordem original ───────────────────────────────
        # Lookup case-insensitive para corrigir só capitalização quando há lista anterior
        cats_set_lower = {c.lower(): c for c in categorias_criadas} if categorias_anteriores else {}

        for item in classificacoes:
            idx_local = item.get("indice", 0) - 1   # 1-based → 0-based
            if 0 <= idx_local < len(indices_validos):
                idx_original = indices_validos[idx_local]
                cat = self._capitalizar(item.get("categoria", "Nao_classificado"))

                # Se há pesquisa anterior: corrige só capitalização.
                # NÃO faz fuzzy — forçar um match ruim é pior que manter o original.
                if categorias_anteriores:
                    cat = cats_set_lower.get(cat.lower(), cat)

                resultados[idx_original] = cat

                chave = f"{tipo}::{respostas_validas[idx_local].lower()}"
                self._cache[chave] = cat

                if callback_progresso:
                    callback_progresso(idx_local, len(respostas_validas),
                                       respostas_validas[idx_local], cat)

        # Registra novas categorias criadas
        for cat in categorias_criadas:
            if cat and cat not in self.categorias and cat != "SEM_RESPOSTA":
                self.categorias.append(cat)

        return resultados

    # ── Classificação individual (fallback) ───────────────────────────────────

    def _classificar_individual(self, resposta: str, tipo: str, contexto_custom: str) -> str:
        config     = TIPOS_PERGUNTA.get(tipo, TIPOS_PERGUNTA["livre"])
        instrucoes = config["instrucoes"] or contexto_custom or \
                     "Categorize a resposta em uma categoria padronizada e concisa."

        cats_str = ""
        if self.categorias:
            unique = list(dict.fromkeys(self.categorias))[:20]
            cats_str = f"\nCategorias já usadas (prefira reutilizar): {', '.join(unique)}\n"

        system = "Você é especialista em análise qualitativa. Responda APENAS com o nome da categoria, sem explicações."
        user   = f"""{instrucoes}
{cats_str}
Resposta: "{resposta}"
Categoria:"""

        resultado = self._chamar_gpt(system, user, max_tokens=60)
        return self._limpar(resultado)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _buscar_fuzzy(self, resposta: str):
        resp_lower = resposta.lower()
        melhor = (0.0, None)
        for hist, categoria in self.codigos_base.items():
            r = SequenceMatcher(None, resp_lower, hist.lower()).ratio()
            if r > melhor[0]:
                melhor = (r, categoria)
        return melhor[1] if melhor[0] > 0.85 else None

    def _parsear_categorias(self, texto: str) -> list:
        try:
            texto_limpo = re.sub(r"```[a-z]*", "", texto).strip("`").strip()
            dados = json.loads(texto_limpo)
            return [str(c).strip() for c in dados.get("categorias", []) if c]
        except Exception:
            return []

    def _parsear_classificacoes(self, texto: str) -> list:
        try:
            texto_limpo = re.sub(r"```[a-z]*", "", texto).strip("`").strip()
            # Tenta extrair o JSON mesmo que haja texto em volta
            match = re.search(r'\{.*"classificacoes".*\}', texto_limpo, re.DOTALL)
            if match:
                dados = json.loads(match.group())
            else:
                dados = json.loads(texto_limpo)
            return dados.get("classificacoes", [])
        except Exception:
            return []

    def _limpar(self, texto: str) -> str:
        cat = texto.strip().strip('"\'').strip(".").split("\n")[0]
        for prefixo in ["categoria:", "resposta:", "categoria é", "a categoria é"]:
            if cat.lower().startswith(prefixo.lower()):
                cat = cat[len(prefixo):].strip()
        cat = cat.strip()
        if not cat:
            return "Nao_classificado"
        # Primeira letra sempre maiúscula, resto preservado
        return cat[0].upper() + cat[1:]

    @staticmethod
    def _capitalizar(cat: str) -> str:
        """Garante que a primeira letra da categoria seja maiúscula."""
        cat = cat.strip()
        if not cat:
            return cat
        return cat[0].upper() + cat[1:]

    # ── Cache ─────────────────────────────────────────────────────────────────

    def exportar_cache(self, caminho: str):
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(self._cache, f, ensure_ascii=False, indent=2)

    def importar_cache(self, caminho: str):
        with open(caminho, encoding="utf-8") as f:
            self._cache.update(json.load(f))
        self.codigos_base.update(self._cache)

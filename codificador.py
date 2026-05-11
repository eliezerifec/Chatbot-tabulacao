"""
Motor de CodificaÃ§Ã£o com IA â€” Fluxo de 2 Agentes via OpenAI
-------------------------------------------------------------
Agente 1 (Categorizador): lÃª TODAS as respostas e cria as categorias
Agente 2 (Classificador): associa cada resposta a uma categoria

NÃ£o depende de crewai. Usa apenas: openai, pandas, requests
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

# â”€â”€ ConfiguraÃ§Ã£o â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# LÃª do arquivo .env (nunca sobe para o GitHub)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # sem dotenv, lÃª das variÃ¡veis de ambiente do sistema

API_KEY = os.getenv("OPENAI_API_KEY", "")
if not API_KEY:
    raise RuntimeError(
        "Chave OpenAI nÃ£o encontrada.\n"
        "Crie um arquivo .env na pasta do projeto com:\n"
        "  OPENAI_API_KEY=sk-..."
    )
DICIONARIO_CODIFICACAO_PATH = os.getenv(
    "DICIONARIO_CODIFICACAO_PATH",
    str(Path(__file__).resolve().parent.parent / "DicionÃ¡rio" / "Dicionario.xlsx")
)

# â”€â”€ Modelos por agente â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# Agente 1 (Categorizador): cria as categorias â€” modelo mais inteligente
MODELO_AGENTE1 = "gpt-5.4"
# Agente 2 (Classificador): vincula respostas Ã s categorias â€” mais rÃ¡pido e barato
MODELO_AGENTE2 = "gpt-4o"

# â”€â”€ Modos de resposta (estrutura da resposta) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# Independente do tipo semÃ¢ntico, a resposta pode ter estrutura diferente
MODOS_RESPOSTA = {
    "simples": {
        "label": "Simples",
        "descricao": "Uma resposta â†’ uma categoria",
    },
    "multipla": {
        "label": "Múltipla",
        "descricao": "Separa por ', ' e codifica cada parte individualmente",
    },
    "semiaberta_simples": {
        "label": "Semiaberta | Simples",
        "descricao": "Categoriza em predefinidas (imputaÃ§Ã£o) ou cria nova coluna",
    },
    "semiaberta_multipla": {
        "label": "Semiaberta | Múltipla",
        "descricao": "Separa por ', ' + categoriza em predefinidas ou cria nova",
    },
}

# â”€â”€ Tipos de pergunta predefinidos â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
TIPOS_PERGUNTA = {
    "reconhecimento_marca": {
        "label": "Reconhecimento de Marca",
        "descricao": "Quais marcas o participante lembrou de ter visto",
        "instrucoes": """Você esté codificando respostas de uma pergunta de reconhecimento de marca espontânea.
O participante respondeu quais marcas lembrou de ter visto em um evento.

REGRAS OBRIGATÓRIAS:
1. Extraia APENAS os nomes das marcas mencionadas
2. Se houver mais de uma marca, separe com ", " (vÍ­rgula espaço)
3. Normalize o nome: capitalize corretamente (ex: "coca cola" â†’ "Coca-Cola")
4. Ignore palavras que nâo sâo marcas (ex: "nâo lembro", "nenhuma")
5. Se nâo houver marca identificável, use a categoria: Não soube responder
6. Cada categoria deve ser o nome da marca normalizado
7. Entenda de fato qual é a marca informada, e não reescreva o que foi respondido
8. Verifique as diferentes formas escritas e reescreva de forma correta (ex: Riachuelo ou riaxuelo ou riachuello, ou eiachuelo. Todas são Riachuelo)

Exemplos de classificaÃ§Ã£o:
  "vi a coca cola e a pepsi"  â†’  Coca-Cola, Pepsi
  "brahma"                    â†’  Brahma
  "nâo lembro de nenhuma"     â†’  Não soube responder
  "Nike e Adidas estavam lÃ¡"  â†’  Nike, Adidas""",
    },

    "satisfacao": {
        "label": "Satisfação / Motivo",
        "descricao": "Por que o participante gostou do evento",
        "instrucoes": """Você está codificando respostas abertas de satisfação de evento.
O participante explicou por que gostou (ou nâo gostou) do evento.

REGRAS OBRIGATPORIAS:
1. Crie UMA categoria temática que resuma a resposta
2. A categoria deve ter NO MÁXIMO 3 palavras, ser uma frase curta e direta
3. Use substantivos/adjetivos descritivos (ex: "boa organizaÇÃo", "atrações diversas", "atendimento ruim")
4. Agrupe respostas com o MESMO TEMA na mesma categoria
5. Seja consistente: respostas similares = mesma categoria

Exemplos de classificação:
  "adorei a organizaçã£o do evento, tudo muito bem feito"  â†’  boa organização
  "as atrações foram incrí­veis"                           â†’  atrações diversas
  "o atendimento foi péssimo"                             â†’  atendimento ruim
  "gostei muito da música ao vivo"                        â†’  música ao vivo
  "estava muito cheio e desorganizado"                    â†’  superlotação desorganizada""",
    },

    "definicao_palavra": {
        "label": "Definição em Uma Palavra",
        "descricao": "Uma palavra que define a experiÃªncia",
        "instrucoes": """VocÃª estÃ¡ codificando respostas de uma pergunta "defina em uma palavra".
O participante escolheu uma palavra para descrever sua experiÃªncia.

REGRAS OBRIGATóRIAS:
1. Normalize a palavra: corrija grafia, capitalize apenas a primeira letra da primeira palavra
2. Agrupe palavras com o MESMO significado ou raiz em uma categoria única:
   - "lindo", "linda", "lindí­ssimo" â†’ Lindo
   - "Ótimo", "Ótima", "otimo" â†’ ótimo
   - "incrÃível", "incrivel", "incredivel" â†’ Incrí­vel
3. Use sempre o masculino singular como forma canônica
4. Se a resposta tiver mais de uma palavra, use apenas a mais relevante

Exemplos de classificação:
  "linda"       â†’  Lindo
  "INCRIVEL"    â†’  Incrí­vel
  "muito bom"   â†’  Bom
  "maravilhoso" â†’  Maravilhoso
  "otimo"       â†’  Ótimo""",
    },

    "local_moradia": {
        "label": "Local de Moradia",
        "descricao": "Cidade, estado ou paÃ­s onde mora",
        "instrucoes": """Você está codificando respostas de uma pergunta sobre local de moradia.

REGRAS OBRIGATÓRIAS:
1. Extraia APENAS o nome do estado ou paí­s
2. Se a pessoa mencionou cidade, retorne o ESTADO correspondente
3. Use o nome completo do estado (ex: "SP" â†’ "São Paulo")
4. Se for fora do Brasil, retorne o nome do PAÍS
5. Normalize a grafia: capitalize corretamente
6. Se não for possÃ­vel identificar, use a categoria: Não sabe

Exemplos de classificaçãoo:
  "moro em São Paulo capital"  â†’  São Paulo
  "Rio de Janeiro, Copacabana" â†’  Rio de Janeiro
  "sou de BH"                  â†’  Minas Gerais
  "moro em SP"                 â†’  SÃ£o Paulo
  "Argentina"                  â†’  Argentina
  "não sei"                    â†’  Não sabe""",
    },

    "livre": {
        "label": "Personalizado",
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
        linhas.append("\nExemplos onde a IA ERROU â€” aprenda com estes casos:")
        for e in corrigidos:
            cat_ia = e.get("categoria_ia", "?")
            cat_ok = e["categoria"]
            linhas.append(
                f'  "{e["resposta"]}"'
                f'  [IA disse: {cat_ia}]  â†’  CORRETO: {cat_ok}'
            )

    if aprovados:
        linhas.append("\nExemplos validados pelos pesquisadores:")
        for e in aprovados:
            linhas.append(f'  "{e["resposta"]}"  â†’  {e["categoria"]}')

    return "\n".join(linhas) + "\n"


def _formatar_few_shot_biblioteca(exemplos: list) -> str:
    if not exemplos:
        return ""

    linhas = [
        "\nBiblioteca histórica de codificação â€” use como referência principal de estilo e granularidade:"
    ]

    for e in exemplos:
        pergunta = e.get("pergunta_texto", "").strip()
        if pergunta:
            linhas.append(f"  [Pergunta] {pergunta}")
        linhas.append(f'  "{e["resposta"]}"  â†’  {e["categoria"]}')

    return "\n".join(linhas) + "\n"


class CodificadorIA:
    
    def __init__(self):
        self.codigos_base: dict = {}
        self.categorias: list = []
        self._cache: dict = {}
        self._client = None
        self.banco = BancoAprendizado()
        self.biblioteca = BibliotecaCodificacao(DICIONARIO_CODIFICACAO_PATH)


    # â”€â”€ Cliente OpenAI (criado uma vez) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def _get_client(self) -> OpenAI:
        if self._client is None:
            self._client = OpenAI(api_key=API_KEY)
        return self._client

    # Modelos de raciocÃ­nio (o1, o3, o4-*) nÃ£o aceitam temperature nem max_tokens
    _MODELOS_RACIOCINIO = {"o1", "o1-mini", "o1-preview", "o3", "o3-mini", "o4-mini", "gpt-5", "gpt-5."}

    def _chamar_gpt(self, system: str, user: str, max_tokens: int = 2000,
                    modelo: str = None) -> str:
        """Chamada Ã  API da OpenAI â€” compatÃ­vel com modelos GPT e de raciocÃ­nio."""
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
            # Modelos de raciocÃ­nio: sem system role, sem temperature
            params["messages"] = [
                {"role": "user", "content": f"{system}\n\n{user}"},
            ]
            params["max_completion_tokens"] = max_tokens

        resp = self._get_client().chat.completions.create(**params)
        return resp.choices[0].message.content.strip()

    # â”€â”€ AlimentaÃ§Ã£o â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€ CodificaÃ§Ã£o linha a linha (compatibilidade com a UI) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€


    def codificar_lote_modo(self, respostas: list, tipo: str = "livre",
                            modo: str = "simples",
                            contexto_custom: str = "",
                            categorias_imputacao: list = None,
                            categorias_anteriores: list = None,
                            callback_progresso=None) -> dict:
        """
        Codifica um lote respeitando o modo de resposta.

        categorias_anteriores: lista de categorias de uma pesquisa jÃ¡ realizada.
            Quando fornecida, o Agente 1 Ã© instruÃ­do a reutilizÃ¡-las
            obrigatoriamente e sÃ³ criar nova categoria em Ãºltimo caso absoluto.

        Retorna dict com chaves dependendo do modo:
          simples/livre   â†’ {"resultado": [str, ...]}
          multipla        â†’ {"resultado": [str, ...]}  (cÃ©lulas com "A, B, C")
          semiaberta_*    â†’ {"imputado": [str, ...], "novo": [str, ...]}
        """
        categorias_imputacao  = categorias_imputacao or []
        categorias_anteriores = categorias_anteriores or []

        # â”€â”€ MÃºltipla: explode por ", ", codifica cada parte, reagrupa â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if modo == "multipla":
            # Expande respostas com mÃºltiplos valores
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

            # Reagrupa por Ã­ndice original
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

        # â”€â”€ Semiaberta: o modelo decide se encaixa ou cria nova â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        if "semi" in modo:
            resultado_semi = self._codificar_semiaberta(
                respostas, tipo=tipo,
                modo=modo,
                contexto_custom=contexto_custom,
                categorias_imputacao=categorias_imputacao,
                categorias_anteriores=categorias_anteriores,
                callback_progresso=callback_progresso)
            return resultado_semi

        # â”€â”€ Simples (padrÃ£o) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

        Agente 1 â€” vÃª as categorias prÃ©-definidas + todas as respostas e decide:
                   a) quais respostas encaixam em categoria existente
                   b) quais precisam de categoria nova (e define o nome)

        Agente 2 â€” aplica a decisÃ£o do Agente 1 em cada resposta individualmente

        Retorna {"imputado": [...], "novo": [...]}
          imputado = encaixou em categoria prÃ©-definida (nome exato da categoria)
          novo     = categoria nova criada pela IA
          Respostas sem encaixe E sem sentido â†’ imputado="", novo=""
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

        # â”€â”€ Bloco de pesquisa anterior (semiaberta) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        categorias_anteriores = categorias_anteriores or []
        if categorias_anteriores:
            ant_lista = "\n".join(f"  - {c}" for c in categorias_anteriores)
            bloco_anterior = f"""
â•”â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•—
â•‘  CATEGORIAS DA PESQUISA ANTERIOR â€” PRIORIDADE MÃXIMA                â•‘
â•šâ•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
Estas categorias foram validadas em uma rodada anterior da MESMA pesquisa.
VocÃª DEVE encaixar cada resposta em uma delas sempre que houver qualquer
compatibilidade semÃ¢ntica â€” mesmo que a correspondÃªncia nÃ£o seja perfeita.
SÃ³ crie categoria nova se a resposta expressar uma dimensÃ£o completamente
ausente nesta lista (caso EXTREMAMENTE raro).

{ant_lista}
"""
        else:
            bloco_anterior = ""

        cats_lista = "\n".join(f"  - {c}" for c in categorias_imputacao) \
                     if categorias_imputacao else "  (nenhuma categoria prÃ©-definida)"

        # Para mÃºltipla, expande antes de classificar
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

        # â”€â”€ Agente 1: define o mapeamento resposta â†’ decisÃ£o â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        system_a1 = (
            "VocÃª Ã© um especialista em anÃ¡lise qualitativa de pesquisas qualitativas brasileiras. "
            "VocÃª entende intenÃ§Ã£o por trÃ¡s de respostas abertas e sabe quando uma resposta "
            "se encaixa semanticamente em uma categoria mesmo que as palavras sejam diferentes. "
            "Responda SEMPRE em JSON vÃ¡lido, sem texto fora do JSON."
        )

        user_a1 = f"""Contexto da pesquisa:
{instrucoes}
{few_shot}
{bloco_anterior}
â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
CATEGORIAS PRÃ‰-DEFINIDAS â€” leia com atenÃ§Ã£o:
{cats_lista}
â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

COMO DECIDIR (siga esta ordem de raciocÃ­nio para cada resposta):

PASSO 1 â€” Tente encaixar em uma categoria prÃ©-definida.
  Encaixe por INTENÃ‡ÃƒO e SEMÃ‚NTICA, nÃ£o sÃ³ por palavras iguais.
  Exemplos de encaixe correto (validados pelos pesquisadores):
  - "Por conta do meu trabalho" â†’ "Queria complementar minha formaÃ§Ã£o (adquirir novos conhecimentos)" (exigÃªncia profissional = aprendizado)
  - "Soft skill de comunicaÃ§Ã£o" â†’ "Queria complementar minha formaÃ§Ã£o (adquirir novos conhecimentos)" (desenvolvimento = aprendizado)
  - "Jovem aprendiz" â†’ "Queria entrar no mercado de trabalho" (inserÃ§Ã£o no mercado)
  - "AtualizaÃ§Ã£o de currÃ­culo" â†’ "Queria entrar no mercado de trabalho" (busca de emprego)
  - "Banho e tosa" â†’ "Procurava um hobby" (atividade por interesse pessoal)
  - "Eu jÃ¡ desenho e queria melhorar" â†’ "Procurava um hobby" (aperfeiÃ§oamento de hobby)

PASSO 2 â€” SÃ“ crie categoria nova se a resposta expressar uma intenÃ§Ã£o que GENUINAMENTE
  nÃ£o tem equivalente em nenhuma das categorias prÃ©-definidas.
  Exemplos que SÃƒO categorias novas legÃ­timas (validados pelos pesquisadores):
  - "Segunda renda" â†’ "Complementar renda" (renda extra Ã© diferente de entrar no mercado de trabalho)
  - "ObrigaÃ§Ã£o escolar / trabalho acadÃªmico" â€” nÃ£o Ã© nenhuma das intenÃ§Ãµes prÃ©-definidas
  - "IndicaÃ§Ã£o mÃ©dica / terapia ocupacional" â€” fora do escopo das categorias

PASSO 3 â€” Use null apenas para respostas vazias, ininteligÃ­veis ou "nÃ£o sei".

Respostas para classificar ({len(respostas_para_classificar)} no total):
{todas_enumeradas}

Retorne SOMENTE este JSON:
{{"classificacoes": [
  {{"indice": 1, "categoria": "nome EXATO da categoria prÃ©-definida", "nova": false}},
  {{"indice": 2, "categoria": "nome da categoria nova criada", "nova": true}},
  {{"indice": 3, "categoria": null, "nova": false}}
]}}

REGRAS FINAIS:
- "nova": false â†’ use o nome EXATAMENTE como estÃ¡ nas categorias prÃ©-definidas
- "nova": true  â†’ categoria nova, 2 a 5 palavras, objetiva
- Classifique TODAS as {len(respostas_para_classificar)} respostas
- Na dÃºvida entre encaixar ou criar nova â†’ ENCAIXE na mais prÃ³xima"""

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

        # â”€â”€ Monta resultados â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Para mÃºltipla: agrupa de volta por resposta original
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
                resultado_str = f"IMP:{imp}" if imp else f"NOVO:{nov}" if nov else "â€”"
                callback_progresso(i, len(indices_validos),
                                   respostas_validas[i], resultado_str)

        return {"imputado": col_imputado, "novo": col_novo}

    def _vincular_com_lista_anterior(self, respostas: list, lista: list,
                                     instrucoes: str, few_shot: str,
                                     callback_progresso=None) -> list:
        """
        Vincula cada resposta ao item EXATO da lista anterior que melhor a representa.
        NÃ£o cria categorias novas â€” a lista Ã© o universo completo de saÃ­das.

        Funciona em lotes de 200 e retorna os nomes exatamente como estÃ£o na lista.
        """
        resultados = ["SEM_RESPOSTA"] * len(respostas)

        # Mapa case-insensitive para garantir nome exato na saÃ­da
        lista_lower = {c.lower(): c for c in lista}
        cats_formatada = "\n".join(f"  {i+1}. {c}" for i, c in enumerate(lista))

        system = (
            "VocÃª Ã© um especialista em vincular respostas abertas a categorias prÃ©-definidas. "
            "Sua tarefa Ã© encontrar, para cada resposta, qual categoria da lista melhor a representa â€” "
            "por significado, abreviaÃ§Ã£o, sinÃ´nimo ou variaÃ§Ã£o ortogrÃ¡fica. "
            "NUNCA invente categorias fora da lista. "
            "Responda SEMPRE em JSON vÃ¡lido, sem texto fora do JSON."
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
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
LISTA FECHADA â€” ESTES SÃƒO OS ÃšNICOS VALORES DE SAÃDA PERMITIDOS:
{cats_formatada}
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

REGRAS DE VINCULAÃ‡ÃƒO:
1. Para cada resposta, encontre qual item da lista acima melhor a representa.
2. Considere: abreviaÃ§Ãµes ("Facha" â†’ "Faculdades Integradas HÃ©lio Alves"),
   siglas ("CIEE" â†’ "CIEE"), nomes parciais ("Celso Lisboa" â†’ "Centro UniversitÃ¡rio Celso Lisboa"),
   sinÃ´nimos e variaÃ§Ãµes ortogrÃ¡ficas.
3. Copie o nome do item da lista EXATAMENTE como estÃ¡ escrito â€” sem alterar uma letra.
4. Se a resposta for vazia, ilegÃ­vel ou "nÃ£o sei" â†’ use "SEM_RESPOSTA".
5. Se genuinamente nÃ£o tiver nenhum equivalente na lista â†’ use "SEM_RESPOSTA".
6. NUNCA escreva um valor que nÃ£o esteja na lista acima (exceto "SEM_RESPOSTA").

Respostas para vincular:
{lote_enum}

Retorne SOMENTE este JSON:
{{"classificacoes": [{{"indice": {inicio+1}, "categoria": "nome exato da lista"}}, ...]}}
Classifique TODAS as {len(lote)} respostas."""

            texto = self._chamar_gpt(system, user, max_tokens=4000,
                                     modelo=MODELO_AGENTE2)
            classificacoes += self._parsear_classificacoes(texto)

        # Aplica resultados â€” garante nome exato via lookup case-insensitive
        for item in classificacoes:
            idx = item.get("indice", 0) - 1   # 1-based â†’ 0-based
            if 0 <= idx < len(respostas):
                cat_raw = (item.get("categoria") or "SEM_RESPOSTA").strip()
                # Corrige capitalizaÃ§Ã£o para o nome exato da lista
                cat = lista_lower.get(cat_raw.lower(), cat_raw)
                resultados[idx] = cat
                if callback_progresso:
                    callback_progresso(idx, len(respostas), respostas[idx], cat)

        return resultados

    def codificar(self, resposta: str, tipo: str = "livre", contexto_custom: str = "") -> str:
        """
        Codifica uma resposta individual.
        Usado como fallback â€” prefira codificar_lote() para processar abas inteiras.
        """
        resposta = str(resposta).strip()
        if not resposta or resposta.lower() in ("nan", "none", ""):
            return "SEM_RESPOSTA"

        chave = f"{tipo}::{resposta.lower()}"
        if chave in self._cache:
            return self._cache[chave]

        # Tenta match exato / fuzzy no histÃ³rico
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


    # â”€â”€ Tipos que NÃƒO precisam de categorizaÃ§Ã£o â€” sÃ³ normalizaÃ§Ã£o â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    # Para marcas e definiÃ§Ã£o em 1 palavra, a prÃ³pria resposta jÃ¡ Ã© a categoria.
    # O modelo sÃ³ precisa limpar, corrigir grafia e normalizar â€” nunca "SEM_RESPOSTA".
    TIPOS_NORMALIZACAO = {"reconhecimento_marca", "definicao_palavra"}

    def _normalizar_lote(self, respostas: list, tipo: str,
                         instrucoes: str, few_shot: str,
                         callback_progresso=None) -> list:
        """
        NormalizaÃ§Ã£o em 2 passos:
          Passo 1 â€” Agente 1 vÃª todos os valores Ãºnicos e monta um dicionÃ¡rio
                    de_para: {variacao: nome_oficial}
          Passo 2 â€” Aplica o dicionÃ¡rio mecanicamente em cada resposta,
                    sem pedir ao modelo para "interpretar" de novo.
        Isso garante agrupamento consistente e elimina cÃ³pia-e-cola.
        """
        import re as _re

        resultados_norm = ["SEM_RESPOSTA"] * len(respostas)

        # â”€â”€ Passo 1: dicionÃ¡rio de padronizaÃ§Ã£o â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Trabalha sÃ³ com valores Ãºnicos para economizar tokens
        unicos = sorted(set(r.strip() for r in respostas
                            if r.strip() and r.strip().lower() not in
                            ("nan", "none", "", "nenhuma", "nÃ£o lembra",
                             "nao lembra", "nÃ£o sei", "nao sei")))

        if not unicos:
            return resultados_norm

        lista_unicos = "\n".join(f"- {v}" for v in unicos)

        system_dict = (
            "VocÃª Ã© um especialista em padronizaÃ§Ã£o de nomes de instituiÃ§Ãµes, "
            "marcas e entidades brasileiras. "
            "Seu trabalho Ã© criar um dicionÃ¡rio de-para que converte variaÃ§Ãµes "
            "informais/erradas para o nome oficial correto. "
            "Responda SEMPRE em JSON vÃ¡lido, sem texto fora do JSON."
        )

        user_dict = f"""Contexto da pesquisa:
{instrucoes}
{few_shot}
Abaixo estÃ£o todos os valores Ãºnicos encontrados nas respostas da pesquisa.
Muitos sÃ£o a mesma entidade escrita de formas diferentes.

Valores Ãºnicos encontrados:
{lista_unicos}

Sua tarefa: para CADA valor, identifique a entidade real e defina o nome oficial padronizado.

REGRAS DE PADRONIZAÃ‡ÃƒO:
1. Siglas ficam em MAIÃšSCULO: senaiâ†’SENAI, sescâ†’SESC, sesiâ†’SESI, fgvâ†’FGV,
   pucâ†’PUC, ufrjâ†’UFRJ, uerjâ†’UERJ, ibmecâ†’Ibmec, faetecâ†’FAETEC, cefetâ†’CEFET,
   uffâ†’UFF, unirioâ†’UNIRIO, ibmâ†’IBM, espnâ†’ESPN, ifbâ†’IFB, igaâ†’IGA
2. Nomes prÃ³prios: primeira letra maiÃºscula â€” "estacio"â†’"EstÃ¡cio", "anhanguera"â†’"Anhanguera"
3. Agrupe variaÃ§Ãµes da mesma entidade: "Ibemc", "Ibemec", "Ibmec" â†’ todos viram "Ibmec"
4. Agrupe com e sem complemento: "Faculdade Unirio", "Unirio", "UNIRIO" â†’ todos "UNIRIO"
5. Corrija acentos: "Estacio"â†’"EstÃ¡cio", "Catolica"â†’"CatÃ³lica", "Galpao"â†’"GalpÃ£o"
6. Corrija erros de grafia: "Jfrj"â†’"UFRJ" NÃƒO â€” "Jfrj" Ã© JFRJ (JustiÃ§a Federal RJ)
7. NÃƒO invente nomes â€” se nÃ£o reconhecer, mantenha capitalizado corretamente
8. Respostas como "Nenhuma", "NÃ£o lembra", "NÃ£o sei" â†’ use "SEM_MARCA"

Retorne SOMENTE este JSON:
{{"dicionario": {{"valor_original": "nome_oficial", "outro_valor": "nome_oficial", ...}}}}

IMPORTANTE: inclua TODOS os {len(unicos)} valores Ãºnicos no dicionÃ¡rio."""

        texto_dict = self._chamar_gpt(system_dict, user_dict,
                                      max_tokens=3000, modelo=MODELO_AGENTE2)

        # Parseia o dicionÃ¡rio
        dicionario = {}
        try:
            texto_limpo = _re.sub(r"```[a-z]*", "", texto_dict).strip("`").strip()
            match = _re.search(r'\{.*"dicionario".*\}', texto_limpo, _re.DOTALL)
            dados = json.loads(match.group() if match else texto_limpo)
            dicionario = dados.get("dicionario", {})
        except Exception:
            pass

        # â”€â”€ Passo 2: aplica o dicionÃ¡rio em cada resposta â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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

    # â”€â”€ Fluxo principal: 2 agentes em sequÃªncia â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    def codificar_lote(self, respostas: list, tipo: str = "livre",
                       contexto_custom: str = "",
                       categorias_anteriores: list = None,
                       callback_progresso=None) -> list:
        """
        Fluxo de 2 agentes:
          1. Agente Categorizador â€” lÃª todas as respostas e define as categorias
          2. Agente Classificador â€” associa cada resposta a uma categoria

        callback_progresso(i, total, resposta, categoria) chamado a cada classificaÃ§Ã£o.
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

        # â”€â”€ Tipos de normalizaÃ§Ã£o: agente Ãºnico, sem lista fechada â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Se hÃ¡ categorias_anteriores, NÃƒO normaliza â€” usa fluxo padrÃ£o com
        # a lista fechada, garantindo que os nomes exatos sejam preservados.
        if tipo in self.TIPOS_NORMALIZACAO and not categorias_anteriores:
            norm = self._normalizar_lote(respostas_validas, tipo, instrucoes,
                                         few_shot, callback_progresso)
            for i, idx_orig in enumerate(indices_validos):
                resultados[idx_orig] = norm[i]
            return resultados

        # â”€â”€ Bloco de pesquisa anterior (sem lista anterior: comportamento normal) â”€
        cats_hint = ""
        if self.categorias:
            cats_hint = f"\nCategorias jÃ¡ usadas (reutilize se adequado): {', '.join(self.categorias)}\n"
        regra_novas = (
            "- Crie QUANTAS categorias forem necessÃ¡rias para cobrir todas as respostas\n"
            "- Prefira entre 10 e 30 categorias â€” use mais se a base for diversa\n"
            "- Cada categoria: 1 a 4 palavras, clara e objetiva\n"
            "- Cubra TODOS os temas â€” nenhuma resposta deve ficar sem encaixe"
        )

        todas_enumeradas = "\n".join(f"{i+1}. {r}" for i, r in enumerate(respostas_validas))

        # â”€â”€ Com pesquisa anterior: vinculador dedicado (sem Agente 1) â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # NÃ£o usa o fluxo de 2 agentes â€” em vez disso, um Ãºnico agente recebe
        # a lista fechada e vincula cada resposta ao item mais prÃ³ximo dela.
        # Sem liberdade criativa: a lista Ã© o universo completo de saÃ­das.
        if categorias_anteriores:
            vinculados = self._vincular_com_lista_anterior(
                respostas_validas, categorias_anteriores,
                instrucoes, few_shot, callback_progresso)
            for i, idx_orig in enumerate(indices_validos):
                resultados[idx_orig] = vinculados[i]
            return resultados

        # â”€â”€ Agente 1: Categorizador â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Se hÃ¡ pesquisa anterior, as categorias jÃ¡ estÃ£o definidas â€” pula o Agente 1.
        # Qualquer chamada ao Agente 1 criaria categorias novas mesmo sendo instruÃ­do
        # a nÃ£o fazer isso; a Ãºnica garantia real Ã© nÃ£o chamÃ¡-lo.
        if categorias_anteriores:
            categorias_criadas = [self._capitalizar(c) for c in categorias_anteriores]
        else:
            system_cat = (
                "VocÃª Ã© um especialista em anÃ¡lise qualitativa de pesquisas. "
                "Seu trabalho Ã© ler respostas abertas e definir um conjunto enxuto de categorias temÃ¡ticas. "
                "Responda SEMPRE em JSON vÃ¡lido, sem texto fora do JSON."
            )
            user_cat = f"""Contexto da pesquisa:
{instrucoes}
{cats_hint}{few_shot}
Todas as respostas coletadas ({len(respostas_validas)} no total):
{todas_enumeradas}

Analise TODAS as respostas acima e defina as categorias necessÃ¡rias para cobri-las.
Retorne SOMENTE este JSON (sem markdown, sem explicaÃ§Ã£o):
{{"categorias": ["categoria1", "categoria2", ...]}}

Regras:
{regra_novas}
- NÃƒO crie a categoria "SEM_RESPOSTA" â€” ela sÃ³ existe para respostas literalmente em branco"""

            texto_cats = self._chamar_gpt(system_cat, user_cat, max_tokens=1500,
                                             modelo=MODELO_AGENTE1)
            categorias_criadas = self._parsear_categorias(texto_cats)

            # Se o agente nÃ£o retornou nada vÃ¡lido, usa categorias jÃ¡ conhecidas ou genÃ©rico
            if not categorias_criadas:
                categorias_criadas = self.categorias[:] or ["positivo", "negativo", "neutro", "SEM_RESPOSTA"]

        # Capitaliza todas as categorias antes de passar ao Agente 2
        categorias_criadas = [self._capitalizar(c) for c in categorias_criadas]

        # â”€â”€ Agente 2: Classificador em lotes de 200 â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        system_clf = (
            "VocÃª Ã© um classificador de texto preciso e consistente para anÃ¡lise qualitativa de pesquisas. "
            "Classifique cada resposta em exatamente uma das categorias fornecidas, "
            "seguindo rigorosamente as regras e exemplos do contexto. "
            "Responda SEMPRE em JSON vÃ¡lido, sem texto fora do JSON."
        )
        cats_lista = "\n".join(f"- {c}" for c in categorias_criadas)

        # Aviso extra para o Agente 2 quando hÃ¡ pesquisa anterior
        aviso_anterior_clf = ""
        if categorias_anteriores:
            aviso_anterior_clf = (
                "\nâš ï¸  ATENÃ‡ÃƒO â€” REGRA INVIOLÃVEL:\n"
                "As categorias listadas abaixo sÃ£o os ÃšNICOS valores permitidos.\n"
                "VocÃª DEVE copiar o nome da categoria EXATAMENTE como estÃ¡ na lista â€”\n"
                "sem abreviar, sem reformular, sem omitir palavras.\n"
                "Exemplo: se a lista tem 'Centro UniversitÃ¡rio Celso Lisboa',\n"
                "a resposta deve ser 'Centro UniversitÃ¡rio Celso Lisboa', NÃƒO 'Celso Lisboa'.\n"
            )

        TAMANHO_LOTE = 200
        classificacoes = []

        for inicio in range(0, len(respostas_validas), TAMANHO_LOTE):
            fim = min(inicio + TAMANHO_LOTE, len(respostas_validas))
            lote = respostas_validas[inicio:fim]
            lote_enumerado = "\n".join(f"{inicio+i+1}. {r}" for i, r in enumerate(lote))

            user_clf = f"""Regras e contexto da codificaÃ§Ã£o:
{instrucoes}
{few_shot}{aviso_anterior_clf}
Categorias disponÃ­veis (copie o nome EXATAMENTE como estÃ¡ escrito abaixo):
{cats_lista}

Respostas para classificar:
{lote_enumerado}

Classifique CADA resposta em exatamente UMA categoria da lista acima.
Retorne SOMENTE este JSON (sem markdown, sem explicaÃ§Ã£o):
{{"classificacoes": [{{"indice": 1, "categoria": "..."}}, {{"indice": 2, "categoria": "..."}}, ...]}}

Regras adicionais:
- Os indices devem comeÃ§ar em {inicio+1}
- Classifique TODAS as {len(lote)} respostas deste lote
- O valor de "categoria" deve ser COPIADO LITERALMENTE da lista â€” nenhuma alteraÃ§Ã£o permitida
- Ã‰ PROIBIDO usar "SEM_RESPOSTA" se a resposta tiver qualquer conteÃºdo reconhecÃ­vel
- Se nÃ£o se encaixar perfeitamente, escolha a categoria MAIS PRÃ“XIMA da lista
- "SEM_RESPOSTA" sÃ³ para resposta completamente vazia, ilegÃ­vel ou "nÃ£o sei\""""

            texto_clf = self._chamar_gpt(system_clf, user_clf, max_tokens=4000,
                                             modelo=MODELO_AGENTE2)
            classificacoes += self._parsear_classificacoes(texto_clf)

        # â”€â”€ Montar resultados na ordem original â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Lookup case-insensitive para corrigir sÃ³ capitalizaÃ§Ã£o quando hÃ¡ lista anterior
        cats_set_lower = {c.lower(): c for c in categorias_criadas} if categorias_anteriores else {}

        for item in classificacoes:
            idx_local = item.get("indice", 0) - 1   # 1-based â†’ 0-based
            if 0 <= idx_local < len(indices_validos):
                idx_original = indices_validos[idx_local]
                cat = self._capitalizar(item.get("categoria", "Nao_classificado"))

                # Se hÃ¡ pesquisa anterior: corrige sÃ³ capitalizaÃ§Ã£o.
                # NÃƒO faz fuzzy â€” forÃ§ar um match ruim Ã© pior que manter o original.
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

    # â”€â”€ ClassificaÃ§Ã£o individual (fallback) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def _classificar_individual(self, resposta: str, tipo: str, contexto_custom: str) -> str:
        config     = TIPOS_PERGUNTA.get(tipo, TIPOS_PERGUNTA["livre"])
        instrucoes = config["instrucoes"] or contexto_custom or \
                     "Categorize a resposta em uma categoria padronizada e concisa."

        cats_str = ""
        if self.categorias:
            unique = list(dict.fromkeys(self.categorias))[:20]
            cats_str = f"\nCategorias jÃ¡ usadas (prefira reutilizar): {', '.join(unique)}\n"

        system = "VocÃª Ã© especialista em anÃ¡lise qualitativa. Responda APENAS com o nome da categoria, sem explicaÃ§Ãµes."
        user   = f"""{instrucoes}
{cats_str}
Resposta: "{resposta}"
Categoria:"""

        resultado = self._chamar_gpt(system, user, max_tokens=60)
        return self._limpar(resultado)

    # â”€â”€ Helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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
        for prefixo in ["categoria:", "resposta:", "categoria Ã©", "a categoria Ã©"]:
            if cat.lower().startswith(prefixo.lower()):
                cat = cat[len(prefixo):].strip()
        cat = cat.strip()
        if not cat:
            return "Nao_classificado"
        # Primeira letra sempre maiÃºscula, resto preservado
        return cat[0].upper() + cat[1:]

    @staticmethod
    def _capitalizar(cat: str) -> str:
        """Garante que a primeira letra da categoria seja maiÃºscula."""
        cat = cat.strip()
        if not cat:
            return cat
        return cat[0].upper() + cat[1:]

    # â”€â”€ Cache â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def exportar_cache(self, caminho: str):
        with open(caminho, "w", encoding="utf-8") as f:
            json.dump(self._cache, f, ensure_ascii=False, indent=2)

    def importar_cache(self, caminho: str):
        with open(caminho, encoding="utf-8") as f:
            self._cache.update(json.load(f))
        self.codigos_base.update(self._cache)

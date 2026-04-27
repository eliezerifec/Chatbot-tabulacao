"""
Banco de Aprendizado Coletivo — Supabase (nuvem) + SQLite (fallback offline)
------------------------------------------------------------------------------
Salva exemplos validados por humanos no Supabase para que TODOS os usuários
do programa se beneficiem das correções uns dos outros.

Se o Supabase não estiver acessível (sem internet), cai automaticamente para
o banco SQLite local sem travar o programa.

── SETUP SUPABASE (fazer uma vez) ──────────────────────────────────────────
1. Crie conta gratuita em https://supabase.com
2. Crie um projeto
3. No SQL Editor do Supabase, rode:

    CREATE TABLE exemplos (
        id               BIGSERIAL PRIMARY KEY,
        tipo             TEXT    NOT NULL,
        resposta         TEXT    NOT NULL,
        categoria_ia     TEXT    NOT NULL,
        categoria_humana TEXT    NOT NULL,
        correto          BOOLEAN NOT NULL DEFAULT true,
        data             DATE    NOT NULL DEFAULT CURRENT_DATE
    );

    -- Índice para busca rápida por tipo
    CREATE INDEX idx_exemplos_tipo ON exemplos(tipo, correto);

4. Em Project Settings → API, copie:
   - Project URL  → SUPABASE_URL abaixo
   - anon/public key → SUPABASE_KEY abaixo

5. Instale: py -m pip install supabase
─────────────────────────────────────────────────────────────────────────────
"""

import os
import sqlite3
import random
from pathlib import Path
from datetime import date

# ── Configuração Supabase ─────────────────────────────────────────────────────
# Lê do arquivo .env (nunca sobe para o GitHub)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

# Banco local (fallback offline)
DB_PATH = Path(__file__).parent / "dados" / "aprendizado.db"


def _conectar_supabase():
    """Tenta conectar ao Supabase. Retorna cliente ou None se falhar."""
    if "SEU-PROJETO" in SUPABASE_URL or "sua-anon-key" in SUPABASE_KEY:
        print("[Aprendizado] Credenciais Supabase nao configuradas.")
        return None
    try:
        from supabase import create_client
    except ImportError:
        print("[Aprendizado] Pacote supabase nao instalado. Rode: py -m pip install supabase")
        return None

    # O Supabase lancou novo formato de chaves (sb_publishable_ / sb_secret_).
    # A biblioteca cliente pode precisar da chave no parametro correto.
    # Tentamos primeiro com a chave como fornecida, depois com opcoes alternativas.
    tentativas = [SUPABASE_KEY]

    for chave in tentativas:
        try:
            client = create_client(SUPABASE_URL, chave)
            # Testa conectividade real
            client.table("exemplos").select("id").limit(1).execute()
            print(f"[Aprendizado] Supabase conectado com chave: {chave[:20]}...")
            return client
        except Exception as e:
            err = str(e)
            print(f"[Aprendizado] Tentativa falhou ({chave[:20]}...): {type(e).__name__}: {err[:120]}")

    print("[Aprendizado] Todas as tentativas de conexao Supabase falharam. Usando banco local.")
    return None


class BancoAprendizado:
    """
    Interface única para salvar/buscar exemplos.
    Usa Supabase (nuvem compartilhada) se disponível,
    senão usa SQLite local como fallback.
    """

    def __init__(self):
        self._supabase  = _conectar_supabase()
        self._modo      = "supabase" if self._supabase else "sqlite"

        # SQLite sempre disponível como fallback / cache local
        DB_PATH.parent.mkdir(exist_ok=True)
        self._conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        self._criar_tabela_local()

        if self._modo == "supabase":
            print("[Aprendizado] Conectado ao Supabase — modo compartilhado ✓")
        else:
            print("[Aprendizado] Supabase indisponível — usando banco local.")

    def _criar_tabela_local(self):
        self._conn.execute("""
            CREATE TABLE IF NOT EXISTS exemplos (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo             TEXT    NOT NULL,
                resposta         TEXT    NOT NULL,
                categoria_ia     TEXT    NOT NULL,
                categoria_humana TEXT    NOT NULL,
                correto          INTEGER NOT NULL DEFAULT 1,
                data             TEXT    NOT NULL
            )
        """)
        self._conn.commit()

    # ── Salvar ────────────────────────────────────────────────────────────────

    def salvar(self, tipo: str, resposta: str,
               categoria_ia: str, categoria_humana: str):
        """Salva no Supabase (e no SQLite local como backup)."""
        correto = categoria_ia.strip().lower() == categoria_humana.strip().lower()
        hoje    = str(date.today())

        # Salva sempre no SQLite local (backup / offline)
        self._conn.execute("""
            INSERT INTO exemplos
                (tipo, resposta, categoria_ia, categoria_humana, correto, data)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (tipo, resposta.strip(), categoria_ia.strip(),
              categoria_humana.strip(), int(correto), hoje))
        self._conn.commit()

        # Tenta salvar no Supabase
        if self._supabase:
            try:
                self._supabase.table("exemplos").insert({
                    "tipo":             tipo,
                    "resposta":         resposta.strip(),
                    "categoria_ia":     categoria_ia.strip(),
                    "categoria_humana": categoria_humana.strip(),
                    "correto":          correto,
                    "data":             hoje,
                }).execute()
            except Exception as e:
                print(f"[Aprendizado] Erro ao salvar no Supabase: {e}")
                # Não trava — o SQLite local já salvou

    # ── Buscar few-shot examples ──────────────────────────────────────────────

    def buscar_exemplos(self, tipo: str, n: int = 20) -> list[dict]:
        """
        Busca exemplos validados para o tipo.
        Retorna TODOS os exemplos — corretos E corrigidos.
        Exemplos corrigidos (onde a IA errou) são os mais valiosos:
        ensinam ao modelo exatamente onde ele costuma errar.

        Cada item retornado tem:
          resposta, categoria, categoria_ia (se foi corrigido), correto (bool)
        """
        rows = []

        if self._supabase:
            try:
                resp = (
                    self._supabase.table("exemplos")
                    .select("resposta, categoria_ia, categoria_humana, correto")
                    .eq("tipo", tipo)
                    .order("id", desc=True)
                    .limit(300)
                    .execute()
                )
                rows = [
                    (r["resposta"], r["categoria_ia"],
                     r["categoria_humana"], r["correto"])
                    for r in resp.data
                ]
            except Exception:
                pass

        if not rows:
            cur = self._conn.execute("""
                SELECT resposta, categoria_ia, categoria_humana, correto
                FROM exemplos
                WHERE tipo = ?
                ORDER BY id DESC LIMIT 300
            """, (tipo,))
            rows = cur.fetchall()

        if not rows:
            return []

        # Prioridade: exemplos corrigidos primeiro (mais instrutivos),
        # depois aprovados. Máx 3 por categoria_humana para variedade.
        corrigidos = [(r, ci, ch, c) for r, ci, ch, c in rows if not c]
        aprovados  = [(r, ci, ch, c) for r, ci, ch, c in rows if c]

        vistos: dict[str, int] = {}
        selecionados = []

        for resp, cat_ia, cat_humana, correto in (corrigidos + aprovados):
            if vistos.get(cat_humana, 0) < 3:
                selecionados.append({
                    "resposta":   resp,
                    "categoria":  cat_humana,   # sempre a categoria correta (humana)
                    "categoria_ia": cat_ia,
                    "correto":    bool(correto),
                })
                vistos[cat_humana] = vistos.get(cat_humana, 0) + 1
            if len(selecionados) >= n:
                break

        return selecionados

    # ── Selecionar 5 para revisão ─────────────────────────────────────────────

    def selecionar_para_revisao(self, resultados: list[dict],
                                n: int = 5) -> list[dict]:
        """
        Seleciona n exemplos representativos para revisão humana.
        Estratégia: 1 por categoria (diversidade), priorizando respostas curtas.
        """
        if not resultados:
            return []

        por_categoria: dict[str, list] = {}
        for item in resultados:
            cat = item.get("categoria", "")
            if cat and cat not in ("SEM_RESPOSTA", "ERRO"):
                por_categoria.setdefault(cat, []).append(item)

        candidatos = []
        for cat, itens in por_categoria.items():
            itens_sorted = sorted(itens, key=lambda x: len(x.get("resposta", "")))
            candidatos.append(itens_sorted[0])

        random.shuffle(candidatos)
        selecionados = candidatos[:n]

        if len(selecionados) < n:
            restantes = [i for i in resultados
                         if i not in selecionados
                         and i.get("categoria") not in ("SEM_RESPOSTA", "ERRO")]
            random.shuffle(restantes)
            selecionados += restantes[:n - len(selecionados)]

        return selecionados[:n]

    # ── Estatísticas ──────────────────────────────────────────────────────────

    def stats(self) -> dict:
        """Retorna estatísticas do banco (Supabase se disponível)."""
        if self._supabase:
            try:
                resp  = self._supabase.table("exemplos").select("correto").execute()
                rows  = resp.data
                total  = len(rows)
                acertos = sum(1 for r in rows if r["correto"])
                taxa   = round(acertos / total * 100, 1) if total else 0
                return {"total": total, "acertos": acertos,
                        "taxa_acerto": taxa, "modo": "☁️ compartilhado"}
            except Exception:
                pass

        # Fallback SQLite
        cur = self._conn.execute(
            "SELECT COUNT(*), SUM(correto) FROM exemplos")
        total, acertos = cur.fetchone()
        acertos = acertos or 0
        taxa = round(acertos / total * 100, 1) if total else 0
        return {"total": total, "acertos": acertos,
                "taxa_acerto": taxa, "modo": "💾 local"}

    def fechar(self):
        self._conn.close()
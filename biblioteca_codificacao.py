from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd


class BibliotecaCodificacao:
    COLUNAS_OBRIGATORIAS = [
        "pergunta_texto",
        "tipo_pergunta",
        "modo_resposta",
        "resposta_original",
        "categoria_final",
    ]

    MAP_TIPO = {
        "satisfacao/motivo": "satisfacao",
        "satisfação/motivo": "satisfacao",
        "satisfacao": "satisfacao",
        "satisfação": "satisfacao",
        "reconhecimento de marca": "reconhecimento_marca",
        "reconhecimento_marca": "reconhecimento_marca",
        "definicao em uma palavra": "definicao_palavra",
        "definição em uma palavra": "definicao_palavra",
        "definicao_palavra": "definicao_palavra",
        "local de moradia": "local_moradia",
        "local_moradia": "local_moradia",
        "livre": "livre",
        "aberta": "livre",
        "resposta aberta": "livre",
    }

    MAP_MODO = {
        "aberta": "simples",
        "resposta aberta": "simples",
        "simples": "simples",
        "multipla": "multipla",
        "múltipla": "multipla",
        "semiaberta": "semiaberta_simples",
        "semiaberta simples": "semiaberta_simples",
        "semiaberta multipla": "semiaberta_multipla",
        "semiaberta múltipla": "semiaberta_multipla",
        "semi_simples": "semiaberta_simples",
        "semi_multipla": "semiaberta_multipla",
    }

    def __init__(self, caminho: str | Path | None):
        self.caminho = Path(caminho).expanduser() if caminho else None
        self.df = pd.DataFrame()
        self.ok = False
        self.erro = ""

        if not self.caminho:
            self.erro = "Caminho da biblioteca não informado."
            return

        if not self.caminho.exists():
            self.erro = f"Arquivo não encontrado: {self.caminho}"
            return

        try:
            self.df = self._carregar_excel(self.caminho)
            self.ok = True
        except Exception as e:
            self.erro = str(e)

    def _carregar_excel(self, caminho: Path) -> pd.DataFrame:
        xls = pd.ExcelFile(caminho)
        frames = []

        for aba in xls.sheet_names:
            df = pd.read_excel(caminho, sheet_name=aba)
            df.columns = [str(c).strip() for c in df.columns]

            faltantes = [c for c in self.COLUNAS_OBRIGATORIAS if c not in df.columns]
            if faltantes:
                continue

            df = df[self.COLUNAS_OBRIGATORIAS].copy()
            df["origem_aba"] = aba
            frames.append(df)

        if not frames:
            raise ValueError(
                "Nenhuma aba válida encontrada. A planilha precisa ter as colunas: "
                + ", ".join(self.COLUNAS_OBRIGATORIAS)
            )

        base = pd.concat(frames, ignore_index=True)

        for col in self.COLUNAS_OBRIGATORIAS:
            base[col] = base[col].fillna("").astype(str).str.strip()

        base = base[
            (base["resposta_original"] != "") &
            (base["categoria_final"] != "")
        ].copy()

        base["tipo_norm"] = base["tipo_pergunta"].map(self._normalizar_tipo)
        base["modo_norm"] = base["modo_resposta"].map(self._normalizar_modo)
        base["pergunta_norm"] = base["pergunta_texto"].map(self._normalizar_texto)

        return base.reset_index(drop=True)

    def _normalizar_texto(self, texto: str) -> str:
        texto = str(texto or "").strip().lower()
        texto = re.sub(r"\s+", " ", texto)
        return texto

    def _normalizar_tipo(self, valor: str) -> str:
        chave = self._normalizar_texto(valor)
        return self.MAP_TIPO.get(chave, chave)

    def _normalizar_modo(self, valor: str) -> str:
        chave = self._normalizar_texto(valor)
        return self.MAP_MODO.get(chave, chave)

    def _tokens(self, texto: str) -> set[str]:
        texto = self._normalizar_texto(texto)
        texto = re.sub(r"[^a-z0-9áàâãéèêíïóôõöúçñ ]+", " ", texto)
        return {t for t in texto.split() if len(t) >= 3}

    def _score_pergunta(self, pergunta_a: str, pergunta_b: str) -> float:
        a = self._tokens(pergunta_a)
        b = self._tokens(pergunta_b)
        if not a or not b:
            return 0.0
        inter = len(a & b)
        uni = len(a | b)
        return inter / uni if uni else 0.0

    def buscar_exemplos(
        self,
        tipo: str,
        modo: str,
        pergunta_texto: str = "",
        n: int = 15,
    ) -> list[dict[str, Any]]:
        if not self.ok or self.df.empty:
            return []

        tipo_norm = self._normalizar_tipo(tipo)
        modo_norm = self._normalizar_modo(modo)
        df = self.df.copy()

        if tipo_norm:
            por_tipo = df[df["tipo_norm"] == tipo_norm]
            if not por_tipo.empty:
                df = por_tipo

        if modo_norm:
            por_modo = df[df["modo_norm"] == modo_norm]
            if not por_modo.empty:
                df = por_modo

        if df.empty:
            return []

        df["score"] = 0.0
        if pergunta_texto.strip():
            df["score"] = df["pergunta_texto"].map(
                lambda p: self._score_pergunta(pergunta_texto, p)
            )
            df = df.sort_values(by=["score", "categoria_final"], ascending=[False, True])

        exemplos = []
        vistos_categoria: dict[str, int] = {}

        for _, row in df.iterrows():
            cat = row["categoria_final"].strip()
            if not cat:
                continue

            chave = cat.lower()
            if vistos_categoria.get(chave, 0) >= 3:
                continue

            exemplos.append(
                {
                    "resposta": row["resposta_original"],
                    "categoria": row["categoria_final"],
                    "pergunta_texto": row["pergunta_texto"],
                    "tipo": row["tipo_norm"],
                    "modo": row["modo_norm"],
                    "score": float(row.get("score", 0.0)),
                    "fonte": "biblioteca_excel",
                }
            )
            vistos_categoria[chave] = vistos_categoria.get(chave, 0) + 1

            if len(exemplos) >= n:
                break

        return exemplos

    def listar_categorias_relacionadas(
        self,
        tipo: str,
        modo: str,
        pergunta_texto: str = "",
        n: int = 30,
    ) -> list[str]:
        exemplos = self.buscar_exemplos(tipo=tipo, modo=modo, pergunta_texto=pergunta_texto, n=n)
        categorias = []
        vistos = set()

        for ex in exemplos:
            cat = ex["categoria"].strip()
            chave = cat.lower()
            if cat and chave not in vistos:
                vistos.add(chave)
                categorias.append(cat)

        return categorias

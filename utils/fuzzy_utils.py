"""
fuzzy_utils.py — Motor de Fuzzy Matching para agentes químicos.

Utiliza rapidfuzz para normalizar nomes de agentes informados pelo usuário
antes de consultar o _MAPA_AGENTES. Resolve variações como:
  "Tiner" → "Tolueno"
  "Solvente PU" → "Xileno"
  "MEK" → "Metiletilcetona"
"""

import json
import os
from rapidfuzz import process, fuzz

# Caminho para o dicionário de sinônimos
_SINONIMOS_PATH = os.path.join(
    os.path.dirname(__file__), "..", "data", "sinonimos_quimicos.json"
)

# Cache carregado uma única vez
_SINONIMOS: dict[str, str] = {}


def _carregar_sinonimos() -> dict[str, str]:
    """Carrega o dicionário de sinônimos do JSON. Usa cache após primeiro carregamento."""
    global _SINONIMOS
    if not _SINONIMOS:
        try:
            with open(_SINONIMOS_PATH, encoding="utf-8") as f:
                _SINONIMOS = json.load(f)
        except FileNotFoundError:
            _SINONIMOS = {}
    return _SINONIMOS


def normalizar_agente(nome_informado: str, score_minimo: int = 75) -> str:
    """
    Tenta normalizar o nome de um agente químico usando:
    1. Lookup direto no dicionário de sinônimos (case-insensitive).
    2. Fuzzy matching sobre as chaves do dicionário de sinônimos.

    Retorna o nome normalizado (valor do dicionário) se encontrar correspondência
    com score >= score_minimo. Caso contrário, retorna o nome original sem alteração.

    Parâmetros:
        nome_informado  — nome digitado pelo usuário (ex: "Tiner", "MEK")
        score_minimo    — score mínimo de similaridade (0–100). Padrão: 75.

    Retorna:
        str — nome normalizado ou nome_informado original.
    """
    if not nome_informado or not isinstance(nome_informado, str):
        return nome_informado

    sinonimos = _carregar_sinonimos()
    if not sinonimos:
        return nome_informado

    chave_normalizada = nome_informado.strip().lower()

    # 1. Lookup exato (case-insensitive)
    for chave, valor in sinonimos.items():
        if chave.strip().lower() == chave_normalizada:
            return valor

    # 2. Fuzzy matching sobre as chaves
    chaves = list(sinonimos.keys())
    resultado = process.extractOne(
        nome_informado,
        chaves,
        scorer=fuzz.WRatio,
        score_cutoff=score_minimo,
    )

    if resultado:
        melhor_chave, score, _ = resultado
        return sinonimos[melhor_chave]

    return nome_informado


def sugerir_agentes(nome_informado: str, limite: int = 3, score_minimo: int = 60) -> list[dict]:
    """
    Retorna uma lista com as melhores sugestões de correspondência para um nome informado.
    Útil para exibir opções ao usuário na interface Streamlit.

    Retorna:
        Lista de dicts com 'entrada', 'sugestao' e 'score'.
    """
    if not nome_informado:
        return []

    sinonimos = _carregar_sinonimos()
    if not sinonimos:
        return []

    chaves = list(sinonimos.keys())
    resultados = process.extract(
        nome_informado,
        chaves,
        scorer=fuzz.WRatio,
        limit=limite,
        score_cutoff=score_minimo,
    )

    return [
        {
            "entrada": nome_informado,
            "sugestao": sinonimos[chave],
            "chave_original": chave,
            "score": round(score, 1),
        }
        for chave, score, _ in resultados
    ]

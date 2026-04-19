"""Cliente isolado para Gemini com cascata de modelos."""
import re
import json
import requests

_MODELOS: list = [
    "models/gemini-2.5-flash",
    "models/gemini-2.5-pro",
    "models/gemini-2.0-flash-001",
    "models/gemini-2.0-flash",
]

_URL = "https://generativelanguage.googleapis.com/v1beta/{modelo}:generateContent?key={chave}"


def _chamar_gemini(prompt: str, chave: str, max_tokens: int = 8192) -> str | None:
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": max_tokens},
    }
    for modelo in _MODELOS:
        try:
            r = requests.post(_URL.format(modelo=modelo, chave=chave), json=payload, timeout=120)
            if r.status_code == 200:
                return r.json()["candidates"][0]["content"]["parts"][0]["text"]
        except Exception:
            continue
    return None


def _limpar_json(texto: str) -> str:
    return re.sub(r"^```json\s*|^```\s*|\s*```$", "", texto.strip(), flags=re.IGNORECASE).strip()


def buscar_dados_cas_ia(cas: str, texto_fispq: str, chave: str) -> dict | None:
    prompt = f"""
Voce e um especialista em SST brasileiro.
Para o agente CAS {cas}, retorne APENAS JSON valido:
{{
  "agente": "Nome",
  "nr15_lt": "Limite de Tolerancia NR-15",
  "nr09_acao": "Nivel de Acao NR-09 (50% do LT)",
  "nr07_ibe": "Indicador Biologico de Exposicao NR-07",
  "dec_3048": "Aposentadoria Especial Decreto 3.048/99",
  "esocial_24": "Codigo eSocial Tabela 24"
}}
Trecho da FISPQ:
{texto_fispq[:3000]}
"""
    texto = _chamar_gemini(prompt, chave)
    if not texto:
        return None
    try:
        return json.loads(_limpar_json(texto))
    except Exception:
        return None


def extrair_pgr_via_ia(texto_pgr: str, chave: str) -> list:
    prompt = f"""
Extraia os GHEs do PGR abaixo. Retorne APENAS JSON valido:
[
  {{
    "ghe": "Nome do GHE",
    "cargos": ["Cargo 1", "Cargo 2"],
    "riscos_mapeados": [
      {{"nome_agente": "Agente", "perigo_especifico": "Descricao", "nivel_risco": "MODERADO"}}
    ]
  }}
]
Texto do PGR:
{texto_pgr[:30000]}
"""
    texto = _chamar_gemini(prompt, chave)
    if not texto:
        return []
    try:
        return json.loads(_limpar_json(texto))
    except Exception:
        return []

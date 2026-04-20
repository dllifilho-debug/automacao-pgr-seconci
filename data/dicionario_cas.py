"""
Banco local de CAS — construcao civil.
Camada 1 do Motor Cascata:
  1. Dicionario local (instantaneo)
  2. Cache Supabase  (aprendizado persistente)
  3. Google Gemini   (auto-discovery)
  4. Fallback        (sinaliza revisao manual)
"""
import re
import json
import requests

# ── Dicionario local ─────────────────────────────────────────────

DICIONARIO_CAS: dict = {
    "108-88-3":   {"agente":"Tolueno","nr15_lt":"78 ppm","nr09_acao":"39 ppm","nr07_ibe":"o-Cresol na Urina","dec_3048":"25 anos (1.0.19)","esocial_24":"01.19.036"},
    "1330-20-7":  {"agente":"Xileno","nr15_lt":"78 ppm","nr09_acao":"39 ppm","nr07_ibe":"Acidos Metilhipuricos","dec_3048":"25 anos (1.0.19)","esocial_24":"01.19.036"},
    "71-43-2":    {"agente":"Benzeno","nr15_lt":"VRT-MPT (Cancerigeno)","nr09_acao":"Avaliacao Qualitativa","nr07_ibe":"Acido trans-muconico","dec_3048":"25 anos (1.0.3)","esocial_24":"01.01.006"},
    "67-64-1":    {"agente":"Acetona","nr15_lt":"780 ppm","nr09_acao":"390 ppm","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "64-17-5":    {"agente":"Etanol","nr15_lt":"780 ppm","nr09_acao":"390 ppm","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "78-93-3":    {"agente":"Metiletilcetona (MEK)","nr15_lt":"155 ppm","nr09_acao":"77.5 ppm","nr07_ibe":"MEK na Urina","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "110-54-3":   {"agente":"n-Hexano","nr15_lt":"50 ppm","nr09_acao":"25 ppm","nr07_ibe":"2,5-Hexanodiona","dec_3048":"25 anos (1.0.19)","esocial_24":"01.19.014"},
    "14808-60-7": {"agente":"Silica Cristalina (Quartzo)","nr15_lt":"Anexo 12","nr09_acao":"50% do LT","nr07_ibe":"Raio-X (OIT) e Espirometria","dec_3048":"25 anos (1.0.18)","esocial_24":"01.18.001"},
    "1332-21-4":  {"agente":"Asbesto / Amianto","nr15_lt":"0,1 f/cm3","nr09_acao":"0,05 f/cm3","nr07_ibe":"Raio-X (OIT) e Espirometria","dec_3048":"20 anos (1.0.2)","esocial_24":"01.02.001"},
    "7439-92-1":  {"agente":"Chumbo (Fumos)","nr15_lt":"0,1 mg/m3","nr09_acao":"0,05 mg/m3","nr07_ibe":"Chumbo no Sangue e ALA-U","dec_3048":"25 anos (1.0.8)","esocial_24":"01.08.001"},
    "65997-15-1": {"agente":"Cimento Portland","nr15_lt":"Anexo 12 (Poeiras)","nr09_acao":"50% do LT","nr07_ibe":"Raio-X (OIT) e Espirometria","dec_3048":"25 anos (1.0.18)","esocial_24":"01.18.001"},
    "7664-38-2":  {"agente":"Acido Fosforico","nr15_lt":"1 mg/m3","nr09_acao":"0,5 mg/m3","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "7647-01-0":  {"agente":"Acido Cloridrico","nr15_lt":"5 ppm","nr09_acao":"2,5 ppm","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "7664-93-9":  {"agente":"Acido Sulfurico","nr15_lt":"0,2 mg/m3 (nevoas)","nr09_acao":"0,1 mg/m3","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"01.01.006"},
    "136-52-7":   {"agente":"Octoato de Cobalto","nr15_lt":"0,02 mg/m3 (Cat 1B)","nr09_acao":"0,01 mg/m3","nr07_ibe":"Cobalto na Urina","dec_3048":"25 anos (1.0.7)","esocial_24":"01.07.001"},
    "96-29-7":    {"agente":"MEKO (Metiletilcetoxima)","nr15_lt":"Avaliacao Qualitativa (Cat 2)","nr09_acao":"Avaliacao Qualitativa","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "64742-47-8": {"agente":"Nafta Aromatica Pesada","nr15_lt":"Avaliar fracao aromatica","nr09_acao":"Avaliacao Qualitativa","nr07_ibe":"Acido muconico / Avaliacao Clinica","dec_3048":"25 anos (1.0.19)","esocial_24":"01.19.036"},
    "64742-95-6": {"agente":"Nafta Aromatica Leve (C8-C10)","nr15_lt":"Avaliar fracao aromatica","nr09_acao":"Avaliacao Qualitativa","nr07_ibe":"Acido muconico / Avaliacao Clinica","dec_3048":"25 anos (1.0.19)","esocial_24":"01.19.036"},
    "8006-64-2":  {"agente":"Aguarras Mineral / Terebentina","nr15_lt":"100 ppm","nr09_acao":"50 ppm","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "8008-20-6":  {"agente":"Querosene","nr15_lt":"200 mg/m3","nr09_acao":"100 mg/m3","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "1305-62-0":  {"agente":"Hidroxido de Calcio","nr15_lt":"5 mg/m3","nr09_acao":"2,5 mg/m3","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "1305-78-8":  {"agente":"Oxido de Calcio (Cal Virgem)","nr15_lt":"2 mg/m3","nr09_acao":"1 mg/m3","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "1317-65-3":  {"agente":"Carbonato de Calcio","nr15_lt":"10 mg/m3","nr09_acao":"5 mg/m3","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "68334-30-5": {"agente":"Asfalto / Betume de Petroleo","nr15_lt":"5 mg/m3 (nevoas)","nr09_acao":"2,5 mg/m3","nr07_ibe":"Avaliacao Clinica / Dermatologica","dec_3048":"25 anos (1.0.19)","esocial_24":"01.19.036"},
    "79-01-6":    {"agente":"Tricloroetileno","nr15_lt":"Anexo 13-A (Cancerigeno)","nr09_acao":"Avaliacao Qualitativa","nr07_ibe":"Acido Tricloroacético (TCA)","dec_3048":"25 anos (1.0.19)","esocial_24":"01.19.036"},
    "55965-84-9": {"agente":"Mistura MIT/CMIT (Biocida)","nr15_lt":"Avaliacao Qualitativa","nr09_acao":"Avaliacao Qualitativa","nr07_ibe":"Avaliacao Dermatologica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "12042-78-3": {"agente":"Aluminato de Calcio","nr15_lt":"Anexo 12 (Poeiras)","nr09_acao":"50% do LT","nr07_ibe":"Raio-X (OIT) e Espirometria","dec_3048":"25 anos (1.0.18)","esocial_24":"01.18.001"},
    "141-78-6":   {"agente":"Acetato de Etila","nr15_lt":"400 ppm","nr09_acao":"200 ppm","nr07_ibe":"Avaliacao Clinica","dec_3048":"Nao Enquadrado","esocial_24":"09.01.001"},
    "95-63-6":    {"agente":"1,2,4-Trimetilbenzeno","nr15_lt":"25 ppm","nr09_acao":"12,5 ppm","nr07_ibe":"3,4-Dimetilhipurico na Urina","dec_3048":"25 anos (1.0.19)","esocial_24":"01.19.036"},
}

# ── Fallback padrao ───────────────────────────────────────────────

_FALLBACK = {
    "agente":     "AGENTE NAO MAPEADO",
    "nr15_lt":    "REVISAO DA ENGENHARIA",
    "nr09_acao":  "REVISAO DA ENGENHARIA",
    "nr07_ibe":   "REVISAO DA ENGENHARIA",
    "dec_3048":   "REVISAO DA ENGENHARIA",
    "esocial_24": "REVISAO DA ENGENHARIA",
}

# ── Motor Cascata principal ───────────────────────────────────────

def buscar_ou_descobrir_cas(cas: str, contexto_pdf: str, chave_api: str | None) -> dict:
    """
    Cascata de resolucao:
    1. Dicionario local  → sem custo, instantaneo
    2. Supabase cache    → ja foi pesquisado antes
    3. Gemini            → descobre e persiste no cache
    4. Fallback          → sinaliza revisao manual
    """

    # 1. Dicionario local
    if cas in DICIONARIO_CAS:
        return DICIONARIO_CAS[cas]

    # 2. Cache Supabase
    try:
        from config.db import consultar_dicionario_dinamico
        cached = consultar_dicionario_dinamico(cas)
        if cached:
            return cached
    except Exception:
        pass

    # 3. Gemini Auto-Discovery
    if chave_api:
        resultado = _consultar_gemini(cas, contexto_pdf, chave_api)
        if resultado:
            try:
                from config.db import salvar_dicionario_dinamico
                salvar_dicionario_dinamico(cas, resultado)
            except Exception:
                pass  # Cache e bonus, nao pode travar o fluxo
            return resultado

    # 4. Fallback
    return _FALLBACK.copy()


# ── Integracao Gemini ─────────────────────────────────────────────

_PROMPT_TEMPLATE = """Voce e um especialista em Higiene Ocupacional e Seguranca do Trabalho brasileiro.

Com base no numero CAS {cas} e no trecho da FISPQ abaixo, retorne EXCLUSIVAMENTE um JSON valido,
sem texto adicional, sem markdown, sem blocos de codigo, com exatamente estas 6 chaves:

- "agente": nome tecnico da substancia em portugues
- "nr15_lt": limite de tolerancia da NR-15 (se ausente, referencie ACGIH com sufixo "Ref. ACGIH")
- "nr09_acao": nivel de acao da NR-09 (geralmente 50 porcento do LT)
- "nr07_ibe": indice biologico de exposicao conforme NR-07 (ou "Avaliacao Clinica" se nao houver)
- "dec_3048": enquadramento no Decreto 3.048/99 (ex: "25 anos (1.0.19)" ou "Nao Enquadrado")
- "esocial_24": codigo da Tabela 24 do eSocial (ex: "01.19.036" ou "09.01.001")

Trecho da FISPQ (contexto):
{trecho}

RESPONDA APENAS COM O JSON. Nenhuma palavra antes ou depois."""


def _consultar_gemini(cas: str, contexto_pdf: str, chave_api: str) -> dict | None:
    """Chama a API do Gemini e retorna o dicionario do agente ou None se falhar."""
    try:
        # Auto-discovery do melhor modelo disponivel
        resp_modelos = requests.get(
            f"https://generativelanguage.googleapis.com/v1beta/models?key={chave_api}",
            timeout=10,
        )
        if resp_modelos.status_code != 200:
            return None

        modelos_disponiveis = [
            m["name"] for m in resp_modelos.json().get("models", [])
            if "generateContent" in m.get("supportedGenerationMethods", [])
        ]

        modelo = None
        for pref in ["models/gemini-1.5-flash", "models/gemini-1.5-pro", "models/gemini-pro"]:
            if pref in modelos_disponiveis:
                modelo = pref
                break
        if not modelo:
            modelo = modelos_disponiveis[0] if modelos_disponiveis else None
        if not modelo:
            return None

        prompt = _PROMPT_TEMPLATE.format(cas=cas, trecho=contexto_pdf[:3000])

        resp_gen = requests.post(
            f"https://generativelanguage.googleapis.com/v1beta/{modelo}:generateContent?key={chave_api}",
            json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.0},
            },
            timeout=25,
        )
        if resp_gen.status_code != 200:
            return None

        texto = resp_gen.json()["candidates"][0]["content"]["parts"][0]["text"]

        # Extrai o JSON mesmo que venha com lixo em volta
        match = re.search(r'\{.*?\}', texto, re.DOTALL)
        if not match:
            return None

        dados = json.loads(match.group())

        # Valida se todas as chaves necessarias estao presentes
        if all(k in dados for k in _FALLBACK.keys()):
            return dados

    except Exception:
        pass

    return None

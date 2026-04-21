"""
utils/biologico_utils.py — v2.0
Detecta risco biológico REAL: apenas agentes biológicos explícitos
(sangue, esgoto, vírus), nunca por menção genérica no PDF.
"""

CHAVES_BIOLOGICAS_MATRIZ = {"BIOLOGICO", "ESGOTO", "SANGUE"}

# Agentes mapeados que confirmam exposição biológica real
_AGENTES_BIOLOGICOS_REAIS = {
    "Agentes Biologicos",
    "Material Biologico",
    "Esgoto / Aguas Servidas",
}

# Palavras no perigo_especifico que confirmam risco biológico concreto
_PALAVRAS_BIO_CONFIRMAM = [
    "virus","bacteria","fungo","hiv","hepatite","hbsag",
    "sangue","fezes","esgoto","aguas servidas",
    "fluido corporal","aerossol biologico",
    "material biologico",
]

# Palavras que INVALIDAM — falsos positivos clássicos do PDF de PGR
_PALAVRAS_BIO_INVALIDAM = [
    "limite biologico","ibmp","tlv-b","monitoramento biologico",
    "controle biologico","avaliacao biologica","programa biologico",
    "indicadores biologicos","indices biologicos",
]


def tem_risco_biologico_real(riscos: list) -> bool:
    """
    Retorna True apenas se o agente mapeado for explicitamente biológico
    E o texto do perigo contiver palavra confirmatória — não apenas
    a palavra genérica 'biologico' copiada do texto corrido do PGR.
    """
    for r in riscos:
        agente = r.get("nome_agente", "")
        perigo = r.get("perigo_especifico", "").lower()

        # Invalida imediatamente se for falso positivo
        if any(inv in perigo for inv in _PALAVRAS_BIO_INVALIDAM):
            continue

        # Agente mapeado diretamente como biológico real
        if agente in _AGENTES_BIOLOGICOS_REAIS:
            return True

        # Só aceita "Agentes Biologicos" genérico se o perigo_especifico
        # tiver evidência concreta (não apenas cópia do PDF)
        if agente == "Agentes Biologicos":
            if any(p in perigo for p in _PALAVRAS_BIO_CONFIRMAM):
                return True
            continue  # rejeita — menção genérica sem evidência

        # Busca direta no texto do perigo
        if any(p in perigo for p in _PALAVRAS_BIO_CONFIRMAM):
            return True

    return False

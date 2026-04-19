"""Deteccao de risco biologico real (exclui falsos positivos ergonomicos)."""

CHAVES_BIOLOGICAS_MATRIZ: set = {"BIOLOGICO", "SANGUE", "MATERIAL BIOLOGICO"}

_TERMOS_BIO_REAIS: list = [
    "BIOLOGICO","VIRUS","BACTERIA","FUNGO","PARASITA","MICRORGANISMO",
    "SANGUE","FLUIDO","SECRECAO","ESGOTO","LIXO","RESIDUO BIOLOGICO",
    "MATERIAL INFECTANTE","AREA DA SAUDE","HOSPITAL","LABORATORIO",
    "FOSSA","EFLUENTE","LODO",
]

_TERMOS_EXCLUSAO: list = [
    "POSTURA","ERGONO","LEVANTAMENTO","REPETITIVO","CARGA","ESFORCO",
    "MENTAL","PSICO","ACIDENTE","ALTURA","ELETRICO","QUIMICO",
    "FISICO","RUIDO","VIBRACAO","CALOR","POEIRA",
]


def tem_risco_biologico_real(riscos: list) -> bool:
    for risco in riscos:
        texto = (
            risco.get("nome_agente", "") + " " + risco.get("perigo_especifico", "")
        ).upper()
        tem_bio  = any(t in texto for t in _TERMOS_BIO_REAIS)
        tem_excl = any(t in texto for t in _TERMOS_EXCLUSAO)
        if tem_bio and not tem_excl:
            return True
    return False

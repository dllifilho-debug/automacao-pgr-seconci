
CHAVES_BIOLOGICAS_MATRIZ = {"BIOLOGICO", "ESGOTO", "SANGUE"}

def tem_risco_biologico_real(riscos: list) -> bool:
    """Verifica se há risco biológico real na lista de riscos mapeados."""
    for r in riscos:
        texto = (r.get("nome_agente","") + " " + r.get("perigo_especifico","")).upper()
        if any(k in texto for k in CHAVES_BIOLOGICAS_MATRIZ):
            return True
    return False

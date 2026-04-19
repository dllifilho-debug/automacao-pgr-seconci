"""Deduplicacao e periodicidade mais restritiva de exames."""

_PRIORIDADE: dict = {
    "3 MESES": 3, "6 MESES": 6, "12 MESES": 12,
    "12 A 24 MESES": 12, "24 MESES": 24, "60 MESES": 60,
}


def periodicidade_mais_restritiva(p1: str, p2: str) -> str:
    v1 = _PRIORIDADE.get(p1.upper(), 999)
    v2 = _PRIORIDADE.get(p2.upper(), 999)
    return p1 if v1 <= v2 else p2


def adicionar_exame_dedup(exames_set: dict, novo: dict) -> None:
    """Adiciona exame sem duplicar; mantém periodicidade mais restritiva."""
    chave   = novo["exame"].strip().upper()
    periodo = novo.get("periodicidade", novo.get("periodico", "12 MESES"))
    motivo  = novo.get("motivo", "")

    if chave in exames_set:
        exames_set[chave]["periodicidade"] = periodicidade_mais_restritiva(
            exames_set[chave]["periodicidade"], periodo
        )
        m_atual = exames_set[chave].get("motivo", "")
        if motivo and motivo not in m_atual:
            exames_set[chave]["motivo"] = f"{m_atual}; {motivo}".lstrip("; ")
    else:
        exames_set[chave] = {
            "exame":         novo["exame"],
            "periodicidade": periodo,
            "motivo":        motivo,
        }


def adicionar_exame_dedup(exames: dict, exame_info: dict):
    """Adiciona exame ao dict sem duplicar (chave = nome do exame)."""
    chave = exame_info["exame"]
    if chave not in exames:
        exames[chave] = exame_info

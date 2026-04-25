import unicodedata

MAPA_CARGOS_CONHECIDOS = [
    # ── Administração ──────────────────────────────────────────────
    "GERENTE ADMINISTRATIVO",
    "ASSISTENTE ADMINISTRATIVO",
    "AUXILIAR ADMINISTRATIVO",
    "ADMINISTRATIVO DE OBRAS",
    "AUXILIAR ADMINISTRATIVO DE OBRAS",
    "JOVEM APRENDIZ",
    "COMPRADOR",
    "AUXILIAR DE COMPRAS",
    "TELEFONISTA",
    "RECEPCIONISTA",
    "AUXILIAR SERVICOS GERAIS",
    "AUXILIAR DE ENGENHARIA CIVIL",
    "MENOR APRENDIZ",
    "AUXILIAR DE LIMPEZA",
    "SERVICOS GERAIS",
    "MOTORISTA",
    # ── Engenharia / Técnico ───────────────────────────────────────
    "ENGENHEIRO CIVIL",
    "ENGENHEIRO",
    "ESTAGIARIO DE ENGENHARIA",
    "ESTAGIARIO DE ENGENHARIA CIVIL",
    "TECNICO DE SEGURANCA DO TRABALHO",
    "ESTAGIARIO DE SEGURANCA DO TRABALHO",
    # ── Canteiro — Operadores / Condutores ────────────────────────
    "OPERADOR DE BETONEIRA",
    "OPERADOR DE CREMALHEIRA",
    "OPERADOR DE GRUA",
    "OPERADOR DE MAQUINAS",
    "SINALEIRO",
    # ── Canteiro — Produção ───────────────────────────────────────
    "PEDREIRO",
    "MEIO OFICIAL DE PEDREIRO",
    "SERVENTE DE OBRA",
    "SERVENTE",
    "ARMADOR",
    "MEIO OFICIAL DE ARMADOR",
    "SERVENTE DE ARMADOR",
    "CARPINTEIRO",
    "MEIO OFICIAL DE CARPINTEIRO",
    "SERVENTE DE CARPINTEIRO",
    "ELETRICISTA INDUSTRIAL",
    "ELETRICISTA",
    "MEIO OFICIAL DE ELETRICISTA",
    "SOLDADOR",
    "SERRALHEIRO",
    "MEIO OFICIAL DE SERRALHEIRO",
    "PINTOR",
    "MEIO OFICIAL DE PINTOR",
    "GESSEIRO",
    "ENCANADOR",
    "MEIO OFICIAL DE ENCANADOR",
    "MONTADOR",
    "MECANICO DE MANUTENCAO",
    "IMPERMEABILIZADOR",
    "ENCARREGADO DE IMPERMEABILIZACAO",
    # ── Canteiro — Supervisão ─────────────────────────────────────
    "MESTRE DE OBRA",
    "MESTRE DE OBRAS",
    "ENCARREGADO DE PEDREIRO",
    "ENCARREGADO DE PINTOR",
    "ENCARREGADO DE ELETRICISTA",
    "ENCARREGADO DE ENCANADOR",
    "ENCARREGADO DE REJUNTE",
    "ENCARREGADO",
    # ── Canteiro — Apoio ──────────────────────────────────────────
    "ALMOXARIFE",
    "VIGIA",
    "PORTEIRO",
    "ZELADOR",
]

PALAVRAS_EXCLUIR_CARGO = [
    "QUANTIDADE", "TOTAL", "NUMERO", "MEDIDAS", "FONTE", "TRAJETORIA",
    "DESCRICAO", "EQUIPAMENTO", "PERIODICIDADE", "RISCOS", "AGENTES",
    "AVALIACAO", "CONTROLE", "RESULTADO", "DADOS", "ANALISE",
    "PREVISTOS", "EXPOSTOS", "FUNCIONARIOS", "TRABALHADORES",
    "COORDENADOR DO PGR", "RESPONSAVEL PELA EMPRESA",
]

DICIONARIO_CARGOS = {
    # ADMINISTRATIVO
    "GERENTE ADMINISTRATIVO":              "ADMINISTRATIVO",
    "ASSISTENTE ADMINISTRATIVO":           "ADMINISTRATIVO",
    "AUXILIAR ADMINISTRATIVO":             "ADMINISTRATIVO",
    "ADMINISTRATIVO DE OBRAS":             "ADMINISTRATIVO",
    "AUXILIAR ADMINISTRATIVO DE OBRAS":    "ADMINISTRATIVO",
    "JOVEM APRENDIZ":                      "ADMINISTRATIVO",
    "MENOR APRENDIZ":                      "ADMINISTRATIVO",
    "COMPRADOR":                           "ADMINISTRATIVO",
    "AUXILIAR DE COMPRAS":                 "ADMINISTRATIVO",
    "TELEFONISTA":                         "ADMINISTRATIVO",
    "RECEPCIONISTA":                       "ADMINISTRATIVO",
    "SERVICOS GERAIS":                     "ADMINISTRATIVO",
    "AUXILIAR SERVICOS GERAIS":            "ADMINISTRATIVO",
    "AUXILIAR DE LIMPEZA":                 "ADMINISTRATIVO",
    "ZELADOR":                             "ADMINISTRATIVO",
    # PRODUÇÃO GERAL
    "PEDREIRO":                            "PRODUCAO_GERAL",
    "MEIO OFICIAL DE PEDREIRO":            "PRODUCAO_GERAL",
    "SERVENTE":                            "PRODUCAO_GERAL",
    "SERVENTE DE OBRA":                    "PRODUCAO_GERAL",
    "SERVENTE DE ARMADOR":                 "PRODUCAO_GERAL",
    "SERVENTE DE CARPINTEIRO":             "PRODUCAO_GERAL",
    "ARMADOR":                             "PRODUCAO_GERAL",
    "MEIO OFICIAL DE ARMADOR":             "PRODUCAO_GERAL",
    "CARPINTEIRO":                         "PRODUCAO_GERAL",
    "MEIO OFICIAL DE CARPINTEIRO":         "PRODUCAO_GERAL",
    "SINALEIRO":                           "PRODUCAO_GERAL",
    "MONTADOR":                            "PRODUCAO_GERAL",
    "MECANICO DE MANUTENCAO":              "PRODUCAO_GERAL",
    "ALMOXARIFE":                          "PRODUCAO_GERAL",
    "VIGIA":                               "PRODUCAO_GERAL",
    "PORTEIRO":                            "PRODUCAO_GERAL",
    "OPERADOR DE MAQUINAS":                "PRODUCAO_GERAL",
    # PERFIS COM GABARITO PRÓPRIO
    "PINTOR":                              "PINTOR",
    "MEIO OFICIAL DE PINTOR":              "PINTOR",
    "ENCARREGADO DE PINTOR":               "PINTOR",
    "SERRALHEIRO":                         "SERRALHEIRO",
    "MEIO OFICIAL DE SERRALHEIRO":         "SERRALHEIRO",
    "SOLDADOR":                            "SERRALHEIRO",
    "IMPERMEABILIZADOR":                   "IMPERMEABILIZADOR",
    "ENCARREGADO DE IMPERMEABILIZACAO":    "IMPERMEABILIZADOR",
    # Eletricista — distinção energizado/desenergizado feita via nome do GHE no modulo_pcmso.py
    # ELETRICISTA INDUSTRIAL → sempre energizado (NR-10)
    "ELETRICISTA INDUSTRIAL":              "ELETRICISTA_ENERGIZADO",
    "ELETRICISTA":                         "ELETRICISTA",
    "MEIO OFICIAL DE ELETRICISTA":         "ELETRICISTA",
    "ENCANADOR":                           "ENCANADOR",
    "MEIO OFICIAL DE ENCANADOR":           "ENCANADOR",
    "OPERADOR DE GRUA":                    "OPERADOR_GRUA",
    "OPERADOR DE CREMALHEIRA":             "OPERADOR_GRUA",
    "OPERADOR DE BETONEIRA":               "OPERADOR_BETONEIRA",
    "MOTORISTA":                           "MOTORISTA",
    "ENGENHEIRO":                          "ENGENHEIRO",
    "ENGENHEIRO CIVIL":                    "ENGENHEIRO",
    "AUXILIAR DE ENGENHARIA CIVIL":        "ENGENHEIRO",
    "ESTAGIARIO DE ENGENHARIA":            "ENGENHEIRO",
    "ESTAGIARIO DE ENGENHARIA CIVIL":      "ENGENHEIRO",
    "TECNICO DE SEGURANCA DO TRABALHO":    "TECNICO_SST",
    "ESTAGIARIO DE SEGURANCA DO TRABALHO": "TECNICO_SST",
    "MESTRE DE OBRA":                      "MESTRE_OBRA",
    "MESTRE DE OBRAS":                     "MESTRE_OBRA",
    "ENCARREGADO":                         "ENCARREGADO_GERAL",
    "ENCARREGADO DE PEDREIRO":             "ENCARREGADO_GERAL",
    "ENCARREGADO DE ELETRICISTA":          "ENCARREGADO_GERAL",
    "ENCARREGADO DE ENCANADOR":            "ENCARREGADO_GERAL",
    "ENCARREGADO DE REJUNTE":              "ENCARREGADO_GERAL",
    "GESSEIRO":                            "GESSEIRO",
}


def normalizar_texto(texto: str) -> str:
    nfkd = unicodedata.normalize("NFD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sem_acento.upper().strip()


def normalizar_cargo(cargo: str) -> str:
    return normalizar_texto(cargo)


def mapear_chave_mestra(cargo: str) -> str:
    cargo_normalizado = normalizar_texto(cargo)
    return DICIONARIO_CARGOS.get(cargo_normalizado, "CARGO_NAO_MAPEADO")

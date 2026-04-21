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
    "GESSEIRO",
    "ENCANADOR",
    "MEIO OFICIAL DE ENCANADOR",
    "MONTADOR",
    "MECÂNICO DE MANUTENÇÃO",
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


def normalizar_texto(texto: str) -> str:
    nfkd = unicodedata.normalize("NFD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sem_acento.upper()


def normalizar_cargo(cargo: str) -> str:
    return normalizar_texto(cargo)

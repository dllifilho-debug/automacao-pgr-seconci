"""Normalizacao de cargos — unica definicao no projeto."""
from data.matriz_exames import ALIAS_CARGOS

MAPA_CARGOS_CONHECIDOS: list = [
    "PEDREIRO","SERVENTE","ELETRICISTA","ENCANADOR","PINTOR","SOLDADOR",
    "CARPINTEIRO","ARMADOR","OPERADOR","MOTORISTA","TECNICO","ENGENHEIRO",
    "MESTRE","ENCARREGADO","ALMOXARIFE","ADMINISTRATIVO","AUXILIAR",
    "ASSISTENTE","GERENTE","DIRETOR","RECEPCIONISTA","SEGURANCA","VIGIA",
    "LIMPEZA","ZELADOR","MEDICO","ENFERMEIRO","MONTADOR","SINALEIRO",
    "APONTADOR","AJUDANTE","LABORATORISTA","BIOQUIMICO",
]

PALAVRAS_EXCLUIR_CARGO: list = [
    "CARGO","FUNCAO","SETOR","GHE","GRUPO","RISCO","PERIGO","AGENTE",
    "FONTE","NIVEL","ACAO","EPI","PROBABILIDADE","SEVERIDADE","TOLERANCIA",
    "LIMITE","ANEXO","NR-","DECRETO","ESOCIAL","PCMSO","PGR","FISPQ",
    "AVALIACAO","RESULTADO","MONITORAMENTO","CONTROLE","MEDIDA","PROTECAO",
]


def normalizar_cargo(cargo: str) -> str:
    cargo_upper = cargo.upper().strip()
    if cargo_upper in ALIAS_CARGOS:
        return ALIAS_CARGOS[cargo_upper]
    for alias, norm in ALIAS_CARGOS.items():
        if alias in cargo_upper:
            return norm
    for nome in MAPA_CARGOS_CONHECIDOS:
        if nome in cargo_upper:
            return nome.title()
    return cargo

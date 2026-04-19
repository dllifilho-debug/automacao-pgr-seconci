
import unicodedata

MAPA_CARGOS_CONHECIDOS = [
    "GERENTE ADMINISTRATIVO",
    "ASSISTENTE ADMINISTRATIVO",
    "AUXILIAR ADMINISTRATIVO",
    "COMPRADOR",
    "AUXILIAR DE COMPRAS",
    "TELEFONISTA",
    "RECEPCIONISTA",
    "ENGENHEIRO CIVIL",
    "AUXILIAR SERVICOS GERAIS",
    "AUXILIAR DE ENGENHARIA CIVIL",
    "TECNICO DE SEGURANCA DO TRABALHO",
    "MENOR APRENDIZ",
    "AUXILIAR DE LIMPEZA",
    "SERVICOS GERAIS",
    "OPERADOR DE MAQUINAS",
    "MOTORISTA",
    "ELETRICISTA",
    "SOLDADOR",
    "CARPINTEIRO",
    "PEDREIRO",
    "SERVENTE DE OBRA",
]

PALAVRAS_EXCLUIR_CARGO = [
    "QUANTIDADE","TOTAL","NUMERO","MEDIDAS","FONTE","TRAJETORIA",
    "DESCRICAO","EQUIPAMENTO","PERIODICIDADE","RISCOS","AGENTES",
    "AVALIACAO","CONTROLE","RESULTADO","DADOS","ANALISE",
]


def normalizar_texto(texto: str) -> str:
    """Remove acentos e converte para uppercase."""
    nfkd = unicodedata.normalize("NFD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sem_acento.upper()


def normalizar_cargo(cargo: str) -> str:
    """Normaliza cargo removendo acentos e deixando em uppercase."""
    return normalizar_texto(cargo)

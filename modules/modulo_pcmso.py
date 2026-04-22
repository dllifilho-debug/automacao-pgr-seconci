import io
import json
import os
import re
import unicodedata
from copy import deepcopy
from datetime import datetime

import pdfplumber
import pandas as pd

from data.matriz_exames import MATRIZ_FUNCAO_EXAME, MATRIZ_RISCO_EXAME
from utils.cargo_utils import MAPA_CARGOS_CONHECIDOS, PALAVRAS_EXCLUIR_CARGO, normalizar_cargo, normalizar_texto, mapear_chave_mestra
from utils.exame_utils import adicionar_exame_dedup
from utils.biologico_utils import CHAVES_BIOLOGICAS_MATRIZ, tem_risco_biologico_real

# Carrega banco V2 uma única vez
_BANCO_V2_PATH = os.path.join("data", "banco_matrizes_v2.json")
try:
    with open(_BANCO_V2_PATH, "r", encoding="utf-8") as _f:
        _BANCO_MATRIZES_V2 = json.load(_f)
except FileNotFoundError:
    _BANCO_MATRIZES_V2 = {}

VERSAO_MODULO_PCMSO = '5.1'

_INVALIDOS_GHE = [
    'QUANTIDADE', 'PREVISTOS', 'EXPOSTOS', 'TOTAL DE', 'NUMERO DE',
    'FUNCIONARIOS', 'TRABALHADORES', 'MEDIDAS DE CONTROLE',
    'FONTE GERADORA', 'TRAJETORIA', 'DESCRICAO', 'ATIVIDADES EXERCIDAS',
    'INFORMACOES SOBRE', 'PAGINA DE REVISAO', 'IDENTIFICACAO DA EMPRESA',
    'COMUNICAR', 'DESEMPENHA ATIVIDADES', 'UTILIZAM-SE', 'DIRETORES DA',
    'DURANTE O DESENVOLVIMENTO', 'OCULOS DE', 'NIVEIS BAIXOS', 'IMPORTANCIA',
    'PERMANENTE ELEVADISSIMA', 'INTERMITENTE', 'ATIVIDADES DE -',
    'ATIVIDADES, UTILIZAM', 'ATIVIDADES PERMANENTE', 'DESENVOLV',
]

_INVALIDOS_GHE_REGEX = [
    r'^-\s+\w', r'comunicar', r'desempenha', r'utilizam.se',
    r'diretores\s+da', r'durante\s+o\s+desenvolv', r'oculos\s+de',
    r'niveis\s+baixos', r'permanente\s+elevad', r'intermitente\s+niveis',
    r'em\s+fun.ao\s+das', r'atividades\s+de\s+-', r'atividades\s+desempenh',
    r'riscos\s+ocupacionais', r'altura,\s+em\s+fun', r'para\s+execu',
    r'que\s+executam', r'os\s+trabalhadores', r'expostos\s+a',
    r'conforme\s+', r'verificar\s+', r'realizar\s+', r'responsav',
    r'^\w\)\s+', r'departamento de seguranca', r'quantitativa',
    r'para verifica', r'avalia.ao quantitativa', r'confirma.ao da categoria',
    r'monitoramento peri', r'medidas de controle', r'grau\s+\d',
    r'avaliacao quantitativa do setor', r'iniciar processo',
    r'confirmacao da categoria', r'monitoramento periodico',
    r'neste\s+ghe', r'expostos\s+neste', r'quantidade\s+de\s+func',
]

_PALAVRAS_CANTEIRO = [
    'OBRA', 'CANTEIRO', 'CONSTRUCAO', 'REFORMA', 'HOSPITAL', 'RESIDENCIAL',
    'EDIFICIO', 'BLOCO', 'TORRE', 'HETRIN', 'VIADUTO', 'PONTE', 'SHOPPING',
    'CONDOMINIO', 'EMPREENDIMENTO', 'MONTAGEM', 'INSTALACAO', 'CAMPO',
]

_PALAVRAS_ESCRITORIO = [
    'ESCRITORIO', 'SEDE', 'CORPORATIVO', 'ADMINISTRACAO', 'MARKETING',
    'TECNOLOGIA DA INFORMACAO', 'RECURSOS HUMANOS', 'FINANCEIRO',
    'CONTABILIDADE', 'JURIDICO', 'COMERCIAL',
]

_RISCOS_CANTEIRO = [
    'RUIDO', 'VIBRACAO', 'POEIRA', 'CIMENTO', 'SILICA', 'TINTA',
    'SOLDA', 'ALTURA', 'CONFINADO', 'MAQUINA', 'INCENDIO',
]

_LIXO_GHE = [
    r'caracteristicas e as circunstancias', r'atividades exercidas',
    r'descricao das atividades', r'informacoes sobre', r'pagina de revisao',
    r'digitacao de textos',
]

_MAPA_AGENTES = {
    'RUIDO': 'RUIDO', 'RUÍDO': 'RUIDO', 'VIBRAÇÃO CORPO': 'VIBRACAO CORPO INTEIRO',
    'VIBRACAO CORPO': 'VIBRACAO CORPO INTEIRO', 'VIBRAÇÃO': 'VIBRACAO', 'VIBRACAO': 'VIBRACAO',
    'BENZENO': 'BENZENO', 'TOLUENO': 'TOLUENO', 'XILENO': 'XILENO', 'ACETONA': 'ACETONA',
    'METIL-ETIL': 'METIL-ETIL-CETONA', 'TETRAHIDRO': 'TETRAHIDROFURANO', 'CICLOHEXAN': 'CICLOHEXANONA',
    'DICLOROMETANO': 'DICLOROMETANO', 'TRICLOROETILENO': 'TRICLOROETILENO', 'ESTIRENO': 'ESTIRENO',
    'HEXANO': 'N-HEXANO', 'FENOL': 'FENOL', 'MERCURIO': 'MERCURIO', 'MERCÚRIO': 'MERCURIO',
    'METANOL': 'METANOL', 'CHUMBO': 'CHUMBO', 'MANGANES': 'MANGANES', 'MANGANÊS': 'MANGANES',
    'CROMO': 'CROMO', 'CADMIO': 'CADMIO', 'CÁDMIO': 'CADMIO', 'ARSENICO': 'ARSENICO',
    'ARSÊNIO': 'ARSENICO', 'COBALTO': 'COBALTO', 'FLUOR': 'FLUOR', 'FLÚOR': 'FLUOR',
    'SOLDA': 'SOLDA', 'MONOXIDO': 'MONOXIDO DE CARBONO', 'MONÓXIDO': 'MONOXIDO DE CARBONO',
    'POLICORTE': 'POLICORTE', 'DIESEL': 'COMBUSTIVEL', 'GASOLINA': 'COMBUSTIVEL',
    'COMBUSTIVEL': 'COMBUSTIVEL', 'COMBUSTÍVEL': 'COMBUSTIVEL', 'SILICA': 'SILICA', 'SÍLICA': 'SILICA',
    'QUARTZO': 'SILICA', 'POEIRA MINERAL': 'POEIRA MINERAL', 'POEIRAS MINERAIS': 'POEIRA MINERAL',
    'CIMENTO': 'CIMENTO', 'ASBESTO': 'ASBESTO', 'AMIANTO': 'ASBESTO', 'FUMO METALICO': 'FUMOS METALICOS',
    'FUMO METÁLICO': 'FUMOS METALICOS', 'MADEIRA': 'MADEIRA', 'TINTA': 'TINTA',
    'IMPERMEAB': 'IMPERMEABILIZACAO', 'MASCARA': 'MASCARA RESPIRATORIA', 'MÁSCARA': 'MASCARA RESPIRATORIA',
    'ALTURA': 'QUEDA DE ALTURA', 'CONFINADO': 'ESPACO CONFINADO', 'ELETRICO': 'RISCO ELETRICO',
    'ELÉTRIC': 'RISCO ELETRICO', 'BIOLOGICO': 'AGENTE BIOLOGICO', 'BIOLÓGICO': 'AGENTE BIOLOGICO',
    'ESGOTO': 'ESGOTO', 'EFLUENTE': 'ESGOTO', 'MOTORISTA': 'MOTORISTA',
}

_RE_GHE = re.compile(
    r'(?:GHE[\s:\.\-]*\d|GRUPO\s+HOMOGENEO|LOCAL\s+DE\s+TRABALHO\s*:\s*\w|SETOR\s*:\s*\w)',
    re.IGNORECASE,
)
_RE_TIPO_RISCO = re.compile(r'^[FQBEA]$')
_RE_CABECALHO_AIHA = re.compile(r'matriz de risco aiha|tipo de risco|identificacao de perigo|codigo e.?social|avaliacao de risco|meio de propagacao|nivel de risco|pouca importancia|probabilidade|efeito', re.IGNORECASE)
_RE_DESCRICAO_FUNCAO = re.compile(r'supervisiona|elabora documentacao|controla recursos|cronograma da obra|executa atividades|responsavel por|realiza tarefas|desenvolve|presta servicos', re.IGNORECASE)
_MAPA_TIPO_RISCO = {'F': 'Fisico', 'Q': 'Quimico', 'B': 'Biologico', 'E': 'Ergonomico', 'A': 'Acidente'}
_PALAVRAS_CARGO_AIHA = [
    'ENCARREGADO', 'PEDREIRO', 'ELETRICISTA', 'CARPINTEIRO', 'SOLDADOR', 'SERVENTE',
    'MOTORISTA', 'ENGENHEIRO', 'TECNICO', 'MESTRE', 'OPERADOR', 'ADMINISTRATIVO',
    'ASSISTENTE', 'AUXILIAR', 'COMPRADOR', 'SUPERVISOR', 'PINTOR', 'ARMADOR', 'MONTADOR',
    'INSTALADOR', 'ENCANADOR', 'BOMBEIRO', 'SERRALHEIRO', 'TOPOGRAFO', 'ALMOXARIFE',
    'VIGIA', 'PORTEIRO', 'ZELADOR', 'MENOR', 'APRENDIZ', 'ESTAGIARIO', 'COORDENADOR',
    'GERENTE', 'DIRETOR', 'SERVICOS GERAIS', 'FISCAL', 'INSPETOR', 'PROJETISTA', 'DESENHISTA',
]

_EXAME_ALIAS = {
    'EXAME CLINICO ANAMNESE EXAME FISICO': 'Exame Clinico',
    'EXAME CLINICO': 'Exame Clinico',
    'EXAME CLINICO SEMESTRAL': 'Exame Clinico',
    'AUDIOMETRIA TONAL PTA': 'Audiometria',
    'AUDIOMETRIA': 'Audiometria',
    'ACUIDADE VISUAL AVALIACAO OFTALMOLOGICA': 'Acuidade Visual',
    'ACUIDADE VISUAL': 'Acuidade Visual',
    'ELETROCARDIOGRAMA ECG': 'ECG',
    'ECG': 'ECG',
    'GLICEMIA DE JEJUM': 'Glicemia em Jejum',
    'GLICEMIA EM JEJUM': 'Glicemia em Jejum',
    'HEMOGRAMA COMPLETO': 'Hemograma',
    'HEMOGRAMA': 'Hemograma',
    'HEMOGRAMA COMPLETO UREIA CREATININA': 'Hemograma',
    'RAIO X DE TORAX OIT': 'RX de Tórax OIT',
    'RX DE TORAX OIT': 'RX de Tórax OIT',
    'RAIO X COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
    'RX COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
    'AC TRICLOROACETICO NA URINA': 'Ácido tricloroacético na urina',
    'ACIDO TRICLOROACETICO NA URINA': 'Ácido tricloroacético na urina',
    'ACETONA NA URINA': 'Acetona na urina',
    'METIL ETIL CETONA MEK NA URINA': 'Metil-Etil-Cetona',
    'METIL ETIL CETONA NA URINA': 'Metil-Etil-Cetona',
    'METIL ETIL CETONA': 'Metil-Etil-Cetona',
    'METILETILCETONA NA URINA': 'Metiletilcetona na urina',
    'CICLOHEXANOL H NA URINA': 'Ciclohexanol na urina',
    'CICLOHEXANOL NA URINA': 'Ciclohexanol na urina',
    'TETRAHIDROFURNANO NA URINA': 'Tetrahidrofurnano na urina',
    'MANGANES NO SANGUE': 'Manganês sanguíneo',
    'MANGANES SANGUINEO': 'Manganês sanguíneo',
    'CARBOXIHEMOGLOBINA NO SANGUE': 'Carboxiemoglobina',
    'CARBOXIEMOGLOBINA NO SANGUE': 'Carboxiemoglobina',
    'CARBOXIHEMOGLOBINA': 'Carboxiemoglobina',
    'CONTAGEM DE RETICULOCITOS': 'Contagem de Reticulócitos',
    'ACIDO TRANS TRANS MUCONICO NA URINA': 'Ácido trans-trans mucônico',
    'ACIDO TRANS TRANS MUCONICO': 'Ácido trans-trans mucônico',
    'AVALIACAO PSICOSSOCIAL NR 35': 'Avaliação Psicossocial',
    'AVALIACAO PSICOSSOCIAL': 'Avaliação Psicossocial',
    'ORTOCRESOL NA URINA': 'Ortocresol na urina',
    'AC METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
    'ACIDO METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
    'ESPIROMETRIA': 'Espirometria',
}

_BASE_MINIMO = ['Exame Clinico', 'Audiometria', 'Espirometria', 'RX de Tórax OIT']
_BASE_COMPLETO = ['Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Hemograma', 'Glicemia em Jejum', 'ECG', 'Espirometria', 'RX de Tórax OIT']
_BASE_GRUA = ['Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Hemograma', 'Glicemia em Jejum', 'ECG']
_BASE_PORTARIA = ['Exame Clinico', 'Acuidade Visual']

_GHE_BASE = {
    '01': _BASE_MINIMO, '02': _BASE_COMPLETO, '03': _BASE_COMPLETO, '04': _BASE_COMPLETO,
    '05': _BASE_MINIMO, '06': _BASE_COMPLETO, '07': _BASE_COMPLETO, '08': _BASE_COMPLETO,
    '09': _BASE_COMPLETO, '10': _BASE_COMPLETO, '11': _BASE_COMPLETO, '12': _BASE_COMPLETO,
    '13': _BASE_COMPLETO, '14': _BASE_MINIMO, '15': _BASE_COMPLETO, '16': _BASE_COMPLETO,
    '17': _BASE_GRUA, '18': _BASE_COMPLETO, '19': _BASE_COMPLETO, '20': _BASE_COMPLETO,
    '21': ['Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Glicemia em Jejum', 'ECG', 'Hemograma', 'Contagem de Reticulócitos', 'Ácido trans-trans mucônico', 'Carboxiemoglobina', 'Avaliação Psicossocial', 'Espirometria', 'RX de Tórax OIT'],
    '22': _BASE_COMPLETO,
    '23': ['Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Hemograma', 'Contagem de Reticulócitos', 'Ácido trans-trans mucônico', 'Glicemia em Jejum', 'ECG', 'Espirometria', 'RX de Tórax OIT'],
    '24': _BASE_PORTARIA,
    '25': ['Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Glicemia em Jejum', 'ECG', 'Hemograma', 'Contagem de Reticulócitos', 'Ácido trans-trans mucônico', 'Ortocresol na urina', 'Metiletilcetona na urina', 'Ác. Metil-hipúrico na urina', 'Espirometria', 'RX de Tórax OIT'],
    '26': ['Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Glicemia em Jejum', 'ECG', 'Hemograma', 'Ácido trans-trans mucônico', 'Contagem de Reticulócitos', 'Ortocresol na urina', 'Metil-Etil-Cetona', 'Acetona na urina', 'Ác. Metil-hipúrico na urina', 'Espirometria', 'RX de Tórax OIT'],
    '27': _BASE_COMPLETO, '28': _BASE_COMPLETO,
}

_GHE_PERIODICIDADES = {
    '01': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES'},
    '02': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '03': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '04': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES'},
    '05': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '06': {'Exame Clinico': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES', 'Ácido tricloroacético na urina': '6 MESES'},
    '07': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '08': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '09': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES'},
    '10': {'Exame Clinico': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES', 'Acetona na urina': '6 MESES', 'Metil-Etil-Cetona': '6 MESES', 'Ciclohexanol na urina': '6 MESES', 'Tetrahidrofurnano na urina': '6 MESES'},
    '11': {'Exame Clinico': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES', 'Ácido tricloroacético na urina': '6 MESES'},
    '12': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '13': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '14': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '15': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '16': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '17': {},
    '18': {'Exame Clinico': '6 MESES', 'Manganês sanguíneo': '6 MESES', 'Carboxiemoglobina': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '19': {'Exame Clinico': '6 MESES', 'Acetona na urina': '6 MESES', 'Metil-Etil-Cetona': '6 MESES', 'Ciclohexanol na urina': '6 MESES', 'Tetrahidrofurnano na urina': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '20': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '21': {'Exame Clinico': '6 MESES', 'Hemograma': '6 MESES', 'Contagem de Reticulócitos': '6 MESES', 'Ácido trans-trans mucônico': '6 MESES', 'Carboxiemoglobina': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES'},
    '22': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '23': {'Exame Clinico': '6 MESES', 'Hemograma': '6 MESES', 'Contagem de Reticulócitos': '6 MESES', 'Ácido trans-trans mucônico': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES'},
    '24': {},
    '25': {'Exame Clinico': '6 MESES', 'Hemograma': '6 MESES', 'Contagem de Reticulócitos': '6 MESES', 'Ácido trans-trans mucônico': '6 MESES', 'Ortocresol na urina': '6 MESES', 'Metiletilcetona na urina': '6 MESES', 'Ác. Metil-hipúrico na urina': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES'},
    '26': {'Exame Clinico': '6 MESES', 'Hemograma': '6 MESES', 'Ácido trans-trans mucônico': '6 MESES', 'Contagem de Reticulócitos': '6 MESES', 'Ortocresol na urina': '6 MESES', 'Metil-Etil-Cetona': '6 MESES', 'Acetona na urina': '6 MESES', 'Ác. Metil-hipúrico na urina': '6 MESES', 'Espirometria': '24 MESES', 'RX de Tórax OIT': '60 MESES'},
    '27': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES'},
    '28': {'Espirometria': '24 MESES', 'RX de Tórax OIT': '12 MESES'},
}

_GHE_EXTRAS = {
    '06': ['Ácido tricloroacético na urina'], '07': ['RX de coluna lombo-sacra'],
    '10': ['Acetona na urina', 'Metil-Etil-Cetona', 'Ciclohexanol na urina', 'Tetrahidrofurnano na urina'],
    '11': ['Ácido tricloroacético na urina'], '18': ['Manganês sanguíneo', 'Carboxiemoglobina'],
    '19': ['Acetona na urina', 'Metil-Etil-Cetona', 'Ciclohexanol na urina', 'Tetrahidrofurnano na urina'],
    '21': ['Hemograma', 'Contagem de Reticulócitos', 'Ácido trans-trans mucônico', 'Carboxiemoglobina', 'Avaliação Psicossocial'],
    '23': ['Hemograma', 'Contagem de Reticulócitos', 'Ácido trans-trans mucônico'],
    '25': ['Hemograma', 'Contagem de Reticulócitos', 'Ácido trans-trans mucônico', 'Ortocresol na urina', 'Metiletilcetona na urina', 'Ác. Metil-hipúrico na urina'],
    '26': ['Hemograma', 'Ácido trans-trans mucônico', 'Contagem de Reticulócitos', 'Ortocresol na urina', 'Metil-Etil-Cetona', 'Acetona na urina', 'Ác. Metil-hipúrico na urina'],
}

_GHE_RESTRICOES = {
    '17': {'Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Hemograma', 'Glicemia em Jejum', 'ECG'},
    '24': {'Exame Clinico', 'Acuidade Visual'},
}

_RISCOS_TRIVIAIS_OBRIGATORIOS = {
    '10': {'SERVENTE': ['Éter monobutílico de etilenoglicol']},
    '19': {'ENCANADOR': ['Éter monobutílico de etilenoglicol'], 'MEIO OFICIAL DE ENCANADOR': ['Éter monobutílico de etilenoglicol'], 'SERVENTE': ['Éter monobutílico de etilenoglicol']},
    '23': {'PINTOR': ['Tolueno', 'Acetona', 'Acetato de etilglicol', 'Xileno']},
    '25': {'PINTOR': ['Acetona', 'Octoato de Cobalto'], 'SERVENTE': ['Acetona', 'Octoato de Cobalto']},
}


def _sem_acentos(texto):
    return unicodedata.normalize('NFKD', str(texto)).encode('ascii', 'ignore').decode('ascii')


def _norm(texto):
    texto = _sem_acentos(str(texto or '')).upper().strip()
    texto = re.sub(r'[^A-Z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()


def _ghe_codigo(nome):
    m = re.search(r'GHE\s*(\d{1,2})', str(nome or ''), re.IGNORECASE)
    return m.group(1).zfill(2) if m else ''


def _nome_oficial_exame(nome):
    if not nome:
        return ''
    return _EXAME_ALIAS.get(_norm(nome), str(nome).strip())


def _limpar_nome_ghe(nome):
    if len(nome) > 100:
        return nome[:100].strip() + '...'
    norm = normalizar_texto(nome)
    for lixo in _LIXO_GHE:
        if re.search(lixo, norm, re.IGNORECASE):
            return 'GHE (revisar nome)'
    return nome.strip()


def _is_linha_ghe(linha):
    lu = normalizar_texto(linha.strip())
    for pat in _INVALIDOS_GHE_REGEX:
        if re.search(pat, lu, re.IGNORECASE):
            return False
    if re.match(r'^GHE\s*\d+', linha.strip(), re.IGNORECASE):
        return True
    if _RE_GHE.search(linha):
        return True
    if len(linha.strip()) <= 50 and '/' not in linha and ',' not in linha and 'DEPARTAMENTO' in lu:
        return True
    return False


def _ghe_valido(nome_ghe):
    norm = normalizar_texto(nome_ghe)
    if len(nome_ghe.strip()) > 90 or len(norm.strip()) < 4:
        return False
    if any(re.search(pat, norm, re.IGNORECASE) for pat in _INVALIDOS_GHE_REGEX):
        return False
    return not any(inv in norm for inv in _INVALIDOS_GHE)


def _fallback_necessario(ghes):
    for g in ghes:
        if len(normalizar_texto(g['ghe'])) <= 90 and g['cargos']:
            return False
    return True


def _fmt_per(per):
    if per is None or per is False:
        return '-'
    per = str(per).strip().upper().replace('MESES', '').replace('MES', '').strip()
    if not per or per in ('TRUE', 'FALSE', 'NONE', ''):
        return '-'
    try:
        return f'{int(per)}M'
    except ValueError:
        return per if per else '-'


def _flag(val):
    if isinstance(val, bool):
        return 'X' if val else '-'
    return 'X' if str(val).strip().upper() in ('X', 'TRUE', '1', 'SIM') else '-'


def _ghe_e_canteiro_misto(nome_ghe, riscos):
    norm = normalizar_texto(nome_ghe)
    if any(p in norm for p in _PALAVRAS_CANTEIRO):
        return True
    if any(p in norm for p in _PALAVRAS_ESCRITORIO):
        return False
    texto_r = ' '.join(normalizar_texto(r.get('nome_agente', '') + ' ' + r.get('perigo_especifico', '')) for r in riscos)
    return any(rc in texto_r for rc in _RISCOS_CANTEIRO)


def _is_nome_funcao_aiha(linha):
    lstrip = linha.strip()
    lu = normalizar_texto(lstrip)
    if not lstrip or len(lstrip) > 60:
        return False
    if _RE_CABECALHO_AIHA.search(lu) or _RE_DESCRICAO_FUNCAO.search(lu) or _RE_TIPO_RISCO.match(lstrip):
        return False
    if lstrip.startswith('-') or re.match(r'^\d{2}\.\d{2}\.\d{3}$', lstrip):
        return False
    if len(lstrip.split()) < 2:
        return False
    return any(p in lu for p in _PALAVRAS_CARGO_AIHA)


def extrair_texto_pdf(uploaded_file):
    texto = []
    with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                texto.append(t)
    return '\n'.join(texto)


def extrair_texto_pdf_path(caminho):
    texto = []
    with pdfplumber.open(caminho) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                texto.append(t)
    return '\n'.join(texto)


def extrair_pgr_local(texto):
    linhas = texto.split('\n')
    ghes, ghe_atual, agentes_set = [], None, set()
    for linha in linhas:
        lc = linha.strip()
        if not lc:
            continue
        lu = normalizar_texto(lc)
        if _is_linha_ghe(lc) and len(lc) < 120 and len(lc.strip()) >= 4 and not lc.strip().endswith('.'):
            if ghe_atual and (ghe_atual['cargos'] or ghe_atual['riscos_mapeados']):
                ghes.append(ghe_atual)
            nome_ghe_limpo = re.split(r'\s+CMO\b|\s+[–\-]\s+CMO|\s+SPE\b|\s+LTDA\b', lc, flags=re.IGNORECASE)[0].strip()
            ghe_atual = {'ghe': nome_ghe_limpo, 'cargos': [], 'riscos_mapeados': []}
            agentes_set = set()
            continue
        if ghe_atual is None:
            continue
        if not any(normalizar_texto(exc) in lu for exc in PALAVRAS_EXCLUIR_CARGO):
            for cargo in MAPA_CARGOS_CONHECIDOS:
                if normalizar_texto(cargo) in lu and cargo not in ghe_atual['cargos']:
                    ghe_atual['cargos'].append(cargo)
                    break
        for palavra, chave_risco in _MAPA_AGENTES.items():
            if normalizar_texto(palavra) in lu and chave_risco not in agentes_set:
                agentes_set.add(chave_risco)
                ghe_atual['riscos_mapeados'].append({'nome_agente': chave_risco, 'perigo_especifico': lc[:200]})
    if ghe_atual and (ghe_atual['cargos'] or ghe_atual['riscos_mapeados']):
        ghes.append(ghe_atual)
    return _deduplicar_ghes(ghes)


def extrair_pgr_matriz_aiha(texto):
    linhas = texto.split('\n')
    ghes = []
    funcao_atual = None
    tipo_risco_atual = None
    agentes_set = set()
    i = 0
    while i < len(linhas):
        lc = linhas[i].strip()
        lu = normalizar_texto(lc)
        if not lc or _RE_CABECALHO_AIHA.search(lu):
            i += 1
            continue
        if _RE_TIPO_RISCO.match(lc):
            tipo_risco_atual = lc.strip()
            i += 1
            continue
        if lc.startswith('-') and funcao_atual is not None and tipo_risco_atual:
            agente_texto = lc.lstrip('- ').split('(')[0].split('\u2013')[0].split('–')[0].strip()[:120]
            agente_norm = normalizar_texto(agente_texto)
            chave_risco = None
            for palavra, chave in _MAPA_AGENTES.items():
                if normalizar_texto(palavra) in agente_norm:
                    chave_risco = chave
                    break
            if not chave_risco:
                chave_risco = agente_texto[:80]
            if chave_risco not in agentes_set:
                agentes_set.add(chave_risco)
                funcao_atual['riscos_mapeados'].append({'nome_agente': chave_risco, 'perigo_especifico': lc[:200], 'tipo_risco': _MAPA_TIPO_RISCO.get(tipo_risco_atual, tipo_risco_atual)})
            i += 1
            continue
        if _is_nome_funcao_aiha(lc):
            nome_completo = lc
            if i + 1 < len(linhas):
                proxima = linhas[i + 1].strip()
                if proxima and len(proxima) <= 40 and not _RE_CABECALHO_AIHA.search(normalizar_texto(proxima)) and not _RE_TIPO_RISCO.match(proxima) and not proxima.startswith('-') and not re.match(r'^\d{2}\.\d{2}\.\d{3}$', proxima) and not _RE_DESCRICAO_FUNCAO.search(normalizar_texto(proxima)):
                    nome_completo = f'{lc} {proxima}'
                    i += 1
            if funcao_atual and (funcao_atual['cargos'] or funcao_atual['riscos_mapeados']):
                ghes.append(funcao_atual)
            funcao_atual = {'ghe': nome_completo, 'cargos': [nome_completo], 'riscos_mapeados': []}
            agentes_set = set()
            tipo_risco_atual = None
            i += 1
            continue
        i += 1
    if funcao_atual and (funcao_atual['cargos'] or funcao_atual['riscos_mapeados']):
        ghes.append(funcao_atual)
    return ghes


def _detectar_formato_pgr(texto):
    norm = normalizar_texto(texto)
    tem_aiha = 'MATRIZ DE RISCO AIHA' in norm
    tem_ghe = bool(re.search(r'GHE\s*[\d:\-]', texto, re.IGNORECASE))
    if tem_aiha and tem_ghe:
        return 'misto'
    if tem_aiha:
        return 'aiha'
    return 'ghe'


def _deduplicar_ghes(ghes):
    vistos, resultado = {}, []
    for ghe in ghes:
        chave = frozenset(ghe.get('cargos', []))
        if not chave:
            resultado.append(ghe)
            continue
        if chave not in vistos:
            vistos[chave] = len(resultado)
            resultado.append(ghe)
        else:
            idx = vistos[chave]
            if len(ghe['ghe']) < len(resultado[idx]['ghe']):
                resultado[idx]['ghe'] = ghe['ghe']
            riscos_existentes = {r['nome_agente'] for r in resultado[idx]['riscos_mapeados']}
            for r in ghe.get('riscos_mapeados', []):
                if r['nome_agente'] not in riscos_existentes:
                    resultado[idx]['riscos_mapeados'].append(r)
                    riscos_existentes.add(r['nome_agente'])
    return resultado


def extrair_pgr_com_fallback(texto_pgr, chave_api=None):
    formato = _detectar_formato_pgr(texto_pgr)
    if formato == 'aiha':
        resultado = extrair_pgr_matriz_aiha(texto_pgr)
        return (resultado, 'aiha') if resultado else ([], 'parcial')
    if formato == 'misto':
        local = extrair_pgr_local(texto_pgr)
        aiha = extrair_pgr_matriz_aiha(texto_pgr)
        nomes_local = {x['ghe'] for x in local}
        merged = local + [g for g in aiha if g['ghe'] not in nomes_local]
        return (merged, 'misto') if merged else ([], 'parcial')
    local = extrair_pgr_local(texto_pgr)
    if not _fallback_necessario(local):
        return local, 'local'
    if chave_api:
        try:
            from utils.ia_client import extrair_pgr_via_ia
            ia = extrair_pgr_via_ia(texto_pgr, chave_api)
            return (ia, 'ia') if ia else (local or [], 'parcial')
        except Exception as e:
            print(f'[WARN] Falha IA: {e}')
    return (local or [], 'parcial')


def _novo_exame(exame, adm=True, per=None, mro=True, rt=False, dem=False, obs='', motivo=''):
    return {'exame': _nome_oficial_exame(exame), 'adm': adm, 'per': per, 'mro': mro, 'rt': rt, 'dem': dem, 'obs': obs, 'motivo': motivo}


def _match_funcao_matriz(cargo_upper, funcao_matriz):
    alvo = _norm(funcao_matriz)
    cargo_n = _norm(cargo_upper)
    return alvo == cargo_n or alvo in cargo_n or cargo_n in alvo


def _forcar_regras_exame(ex, cod_ghe):
    ex = deepcopy(ex)
    nome = _nome_oficial_exame(ex.get('exame', ''))
    ex['exame'] = nome
    if nome == 'Exame Clinico':
        ex['adm'] = True if ex.get('adm') is None else ex.get('adm', True)
        ex['mro'] = True if ex.get('mro') is None else ex.get('mro', True)
        ex['rt'] = True
        ex['dem'] = True if ex.get('dem') is None else ex.get('dem', True)
        ex['per'] = _GHE_PERIODICIDADES.get(cod_ghe, {}).get('Exame Clinico', ex.get('per') or '12 MESES')
    else:
        ex['rt'] = False
    per_regra = _GHE_PERIODICIDADES.get(cod_ghe, {}).get(nome)
    if per_regra:
        ex['per'] = per_regra
    if nome == 'RX de coluna lombo-sacra':
        ex['adm'] = True
        ex['per'] = None
        ex['mro'] = True
        ex['rt'] = False
        ex['dem'] = False
    if cod_ghe == '17' and nome == 'Audiometria':
        ex['adm'] = True
        ex['per'] = ex.get('per') or '12 MESES'
        ex['mro'] = True
        ex['rt'] = False
        ex['dem'] = False
    if nome in {'Ácido tricloroacético na urina', 'Acetona na urina', 'Metil-Etil-Cetona', 'Metiletilcetona na urina', 'Ciclohexanol na urina', 'Tetrahidrofurnano na urina', 'Carboxiemoglobina', 'Ácido trans-trans mucônico', 'Ortocresol na urina', 'Ác. Metil-hipúrico na urina'}:
        ex['adm'] = ex.get('adm', False)
        ex['mro'] = ex.get('mro', False)
        ex['dem'] = ex.get('dem', False)
    if nome == 'Manganês sanguíneo':
        ex['adm'] = True
        ex['mro'] = True
        ex['dem'] = False
    return ex


def _adicionar_base_por_ghe(exames, cod_ghe):
    for nome in _GHE_BASE.get(cod_ghe, _BASE_COMPLETO):
        adicionar_exame_dedup(exames, _forcar_regras_exame(_novo_exame(nome, motivo=f'Base GHE {cod_ghe}'), cod_ghe))
    for nome in _GHE_EXTRAS.get(cod_ghe, []):
        adicionar_exame_dedup(exames, _forcar_regras_exame(_novo_exame(nome, motivo=f'Extra GHE {cod_ghe}'), cod_ghe))


def _riscos_triviais_para_cargo(cod_ghe, cargo):
    return _RISCOS_TRIVIAIS_OBRIGATORIOS.get(cod_ghe, {}).get(_norm(cargo), [])


def _aplicar_funcao_matriz(exames, cargo_norm, cod_ghe):
    for funcao, lista_ex in MATRIZ_FUNCAO_EXAME.items():
        if _match_funcao_matriz(cargo_norm, funcao):
            for ex in lista_ex:
                exame = _novo_exame(
                    ex.get('exame', ''),
                    adm=ex.get('adm', True),
                    per=ex.get('per'),
                    mro=ex.get('mro', True),
                    rt=ex.get('rt', False),
                    dem=ex.get('dem', False),
                    obs=ex.get('obs', ''),
                    motivo=f'Função: {funcao.title()}',
                )
                adicionar_exame_dedup(exames, _forcar_regras_exame(exame, cod_ghe))


def _aplicar_riscos_matriz(exames, riscos, cod_ghe):
    bio_real = tem_risco_biologico_real(riscos)
    for risco in riscos:
        chave_r = normalizar_texto(risco.get('nome_agente', ''))
        if chave_r in CHAVES_BIOLOGICAS_MATRIZ and not bio_real:
            continue
        regra = MATRIZ_RISCO_EXAME.get(chave_r)
        if not regra:
            continue
        exame = _novo_exame(
            regra['exame'], adm=regra.get('adm', True), per=regra.get('periodico'), mro=regra.get('mro', True),
            rt=regra.get('rt', False), dem=regra.get('dem', False), obs=regra.get('obs', ''),
            motivo=f"Exposição: {chave_r.title()} — {regra.get('obs', '')}",
        )
        adicionar_exame_dedup(exames, _forcar_regras_exame(exame, cod_ghe))


def _filtrar_por_restricao_ghe(exames, cod_ghe):
    permitidos = _GHE_RESTRICOES.get(cod_ghe)
    if not permitidos:
        return exames
    return {k: v for k, v in exames.items() if _nome_oficial_exame(v.get('exame', '')) in permitidos}


def _ordenar_exames(rows):
    ordem = ['Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Hemograma', 'Glicemia em Jejum', 'ECG', 'Ácido tricloroacético na urina', 'Acetona na urina', 'Metil-Etil-Cetona', 'Metiletilcetona na urina', 'Ciclohexanol na urina', 'Tetrahidrofurnano na urina', 'Manganês sanguíneo', 'Carboxiemoglobina', 'Contagem de Reticulócitos', 'Ácido trans-trans mucônico', 'Ortocresol na urina', 'Ác. Metil-hipúrico na urina', 'Avaliação Psicossocial', 'Espirometria', 'RX de coluna lombo-sacra', 'RX de Tórax OIT']
    peso = {nome: i for i, nome in enumerate(ordem)}
    return sorted(rows, key=lambda r: (peso.get(_nome_oficial_exame(r['Exame']), 999), _norm(r['Exame'])))


def processar_pcmso(dados_pgr, tipo_ambiente='misto'):
    linhas = []
    for ghe in dados_pgr:
        nome_ghe_raw = ghe.get('ghe', 'Sem GHE')
        nome_ghe = _limpar_nome_ghe(str(nome_ghe_raw))
        cod_ghe = _ghe_codigo(nome_ghe)
        cargos = (ghe.get('cargos') or [])[:15]
        riscos = (ghe.get('riscos_mapeados') or [])[:25]
        if not _ghe_valido(nome_ghe):
            continue
        if tipo_ambiente == 'canteiro':
            e_canteiro = True
        elif tipo_ambiente == 'escritorio':
            e_canteiro = False
        else:
            e_canteiro = _ghe_e_canteiro_misto(nome_ghe, riscos)
        for cargo in cargos:
            cargo_norm = normalizar_cargo(cargo)
            exames = {}

            # ── Motor V2 ──────────────────────────────────────────
            chave_v2 = mapear_chave_mestra(cargo)
            regras_v2 = _BANCO_MATRIZES_V2.get(chave_v2, {}).get("exames", [])

            if regras_v2:
                for ex in regras_v2:
                    novo = _novo_exame(
                        ex["nome"],
                        adm=ex.get("adm", True),
                        per=f'{ex["per"]} MESES' if ex.get("per") else None,
                        mro=ex.get("mro", True),
                        rt=ex.get("ret", False),
                        dem=ex.get("dem", False),
                        motivo=f"Banco V2: {chave_v2}"
                    )
                    adicionar_exame_dedup(exames, _forcar_regras_exame(novo, cod_ghe))
            else:
                # Fallback: lógica antiga por GHE
                adicionar_exame_dedup(exames, _forcar_regras_exame(
                    _novo_exame('Exame Clinico', adm=True, per='12 MESES', mro=True, rt=True, dem=True, motivo='NR-07 Básico'), cod_ghe
                ))
                if cod_ghe:
                    _adicionar_base_por_ghe(exames, cod_ghe)
                elif not e_canteiro:
                    _aplicar_funcao_matriz(exames, cargo_norm, cod_ghe)
                _aplicar_funcao_matriz(exames, cargo_norm, cod_ghe)

            # ── Riscos triviais e biológicos (sempre rodam) ───────
            riscos_expand = list(riscos)
            for risco_trivial in _riscos_triviais_para_cargo(cod_ghe, cargo_norm):
                riscos_expand.append({'nome_agente': risco_trivial, 'perigo_especifico': 'Risco trivial obrigatório PDF referência'})
            _aplicar_riscos_matriz(exames, riscos_expand, cod_ghe)
            exames = _filtrar_por_restricao_ghe(exames, cod_ghe)
            rows_cargo = []
            for ex_info in exames.values():
                nome_exame = _nome_oficial_exame(ex_info.get('exame', ''))
                rt = bool(ex_info.get('rt', False)) if nome_exame == 'Exame Clinico' else False
                rows_cargo.append({
                    'GHE / Setor': nome_ghe,
                    'Cargo': cargo,
                    'Exame': nome_exame,
                    'ADM': _flag(ex_info.get('adm', True)),
                    'PER': _fmt_per(ex_info.get('per')),
                    'MRO': _flag(ex_info.get('mro', True)),
                    'RT': _flag(rt),
                    'DEM': _flag(ex_info.get('dem', False)),
                    'Justificativa': ex_info.get('motivo', ''),
                })
            linhas.extend(_ordenar_exames(rows_cargo))
    return pd.DataFrame(linhas)

def gerar_html_pcmso(df, cabecalho=None):
    if not cabecalho:
        cabecalho = {}
    razao = cabecalho.get('razao_social', 'Empresa não informada')
    cnpj = cabecalho.get('cnpj', '---')
    obra = cabecalho.get('obra', '---')
    medico = cabecalho.get('medico_rt', 'Não informado')
    vig_i = cabecalho.get('vig_ini', '---')
    vig_f = cabecalho.get('vig_fim', '---')
    tec = cabecalho.get('responsavel_tec', '---')
    ghe_grupos = {}
    for _, row in df.iterrows():
        ghe_grupos.setdefault(row['GHE / Setor'], {}).setdefault(row['Cargo'], []).append(row)
    linhas_html = ''
    for ghe_nome, cargos_dict in ghe_grupos.items():
        total_rows = sum(len(v) for v in cargos_dict.values())
        primeiro_ghe = True
        for cargo, rows in cargos_dict.items():
            primeiro_cargo = True
            for row in rows:
                def cel(val, bg='#d4edda'):
                    return f'<td style="text-align:center;background:{bg};">X</td>' if val == 'X' else '<td style="text-align:center;color:#999;">-</td>'
                per_td = f'<td style="text-align:center;font-weight:bold;">{row["PER"]}</td>' if row['PER'] != '-' else '<td style="text-align:center;color:#999;">-</td>'
                ghe_td = ''
                if primeiro_ghe:
                    ghe_td = f'<td rowspan="{total_rows}" style="background:#084D22;color:#fff;font-weight:bold;vertical-align:middle;text-align:center;padding:8px;">{ghe_nome}</td>'
                    primeiro_ghe = False
                cargo_td = ''
                if primeiro_cargo:
                    cargo_td = f'<td rowspan="{len(rows)}" style="vertical-align:middle;font-weight:bold;">{cargo}</td>'
                    primeiro_cargo = False
                linhas_html += f"<tr>{ghe_td}{cargo_td}<td>{row['Exame']}</td>{cel(row['ADM'])}{per_td}{cel(row['MRO'])}{cel(row['RT'])}{cel(row['DEM'])}<td style='font-size:11px;color:#555;'>{row['Justificativa']}</td></tr>"
    return f'''<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><style>body{{font-family:Arial,sans-serif;font-size:13px;margin:20px;}}table{{width:100%;border-collapse:collapse;margin-top:10px;}}th{{background:#1AA04B;color:#fff;padding:10px 6px;border:1px solid #084D22;font-size:12px;}}th.c{{text-align:center;}}td{{border:1px solid #ccc;padding:8px 6px;vertical-align:middle;}}tr:nth-child(even) td{{background:#F4F8F5;}}</style></head><body><table style="margin-bottom:12px;border:2px solid #084D22;"><tr style="background:#084D22;color:#fff;"><td colspan="5" style="padding:8px;font-size:12pt;font-weight:bold;text-align:center;">PROGRAMA DE CONTROLE MÉDICO DE SAÚDE OCUPACIONAL — PCMSO</td></tr><tr><td><b>Empresa:</b> {razao}</td><td><b>CNPJ:</b> {cnpj}</td><td><b>Obra:</b> {obra}</td><td><b>Vigência:</b> {vig_i} a {vig_f}</td><td><b>Emissão:</b> {datetime.now().strftime('%d/%m/%Y')}</td></tr><tr><td colspan="3"><b>Médico(a):</b> {medico}</td><td colspan="2"><b>Técnico SST:</b> {tec}</td></tr></table><table><tr><th style="width:12%">GHE</th><th style="width:14%">Função</th><th style="width:30%">Exame Solicitado</th><th class="c" style="width:5%">ADM</th><th class="c" style="width:6%">PER</th><th class="c" style="width:5%">MRO</th><th class="c" style="width:4%">RT</th><th class="c" style="width:5%">DEM</th><th style="width:19%">Justificativa</th></tr>{linhas_html}</table><p style="font-size:8pt;color:#555;margin-top:12px;">Gerado por Sistema Automação SST Seconci-GO.</p></body></html>'''


def gerar_docx_rq61(df, cabecalho=None):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
    if not cabecalho:
        cabecalho = {}
    razao = cabecalho.get('razao_social', 'Empresa não informada')
    obra = cabecalho.get('obra', '---')
    medico = cabecalho.get('medico_rt', 'Não informado')
    crm = cabecalho.get('crm', '')
    vig_i = cabecalho.get('vig_ini', '---')
    tipo = cabecalho.get('tipo_obra', 'Renovação')
    VERDE_ESC = '084D22'
    VERDE_MED = '1AA04B'
    BRANCO = RGBColor(0xFF, 0xFF, 0xFF)
    def shd(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        s = OxmlElement('w:shd')
        s.set(qn('w:val'), 'clear')
        s.set(qn('w:color'), 'auto')
        s.set(qn('w:fill'), hex_color)
        tcPr.append(s)
    def set_borders(cell, color='084D22'):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for side in ('top', 'left', 'bottom', 'right'):
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), color)
            tcBorders.append(border)
        tcPr.append(tcBorders)
    def txt(cell, text, bold=False, color=None, size=9, align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = align
        r = p.add_run(str(text))
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color
    def _fmt_exame_rq61(row):
        partes = []
        if str(row.get('ADM', '-')) == 'X':
            partes.append('ADM')
        per = str(row.get('PER', '-')).strip().upper()
        if per and per != '-':
            per_num = per.replace('M', '').strip()
            try:
                partes.append(f'PER {int(per_num)} meses')
            except Exception:
                partes.append(f'PER {per}')
        if str(row.get('MRO', '-')) == 'X':
            partes.append('MRO')
        if str(row.get('RT', '-')) == 'X':
            partes.append('RET')
        if str(row.get('DEM', '-')) == 'X':
            partes.append('DEM')
        return f"{row['Exame']} ({', '.join(partes)})" if partes else str(row['Exame'])
    doc = Document()
    for sec in doc.sections:
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(2.0)
        sec.right_margin = Cm(1.5)
    cab = doc.add_table(rows=4, cols=4)
    cab.style = 'Table Grid'
    cab.rows[0].cells[0].merge(cab.rows[0].cells[3])
    shd(cab.rows[0].cells[0], VERDE_ESC)
    txt(cab.rows[0].cells[0], 'MATRIZ FUNÇÃO – EXAMES PCMSO', bold=True, color=BRANCO, size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
    cab.rows[1].cells[0].merge(cab.rows[1].cells[1])
    cab.rows[1].cells[2].merge(cab.rows[1].cells[3])
    txt(cab.rows[1].cells[0], f'Empresa: {razao}', bold=True, size=9)
    adendo_txt = f'Obra Nova (   )   {tipo} ( X )' if str(tipo).lower() == 'renovação' else 'Obra Nova ( X )   Renovação (   )'
    txt(cab.rows[1].cells[2], adendo_txt, size=9)
    cab.rows[2].cells[0].merge(cab.rows[2].cells[1])
    cab.rows[2].cells[2].merge(cab.rows[2].cells[3])
    txt(cab.rows[2].cells[0], f'Obra: {obra}', bold=True, size=9)
    txt(cab.rows[2].cells[2], f"Data: {datetime.now().strftime('%d/%m/%Y')}", bold=True, size=9)
    cab.rows[3].cells[0].merge(cab.rows[3].cells[3])
    crm_txt = f'  CRM-GO {crm}' if crm else ''
    txt(cab.rows[3].cells[0], f'Médico(a) Coordenador(a) do PCMSO: {medico}{crm_txt}', size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    ghe_grupos = {}
    for _, row in df.iterrows():
        ghe_grupos.setdefault(row['GHE / Setor'], {}).setdefault(row['Cargo'], []).append(row)
    for ghe_nome, cargos_dict in ghe_grupos.items():
        tbl = doc.add_table(rows=0, cols=2)
        tbl.style = 'Table Grid'
        tbl.columns[0].width = Cm(5.5)
        tbl.columns[1].width = Cm(12.0)
        row_ghe = tbl.add_row()
        row_ghe.cells[0].merge(row_ghe.cells[1])
        shd(row_ghe.cells[0], VERDE_ESC)
        set_borders(row_ghe.cells[0])
        txt(row_ghe.cells[0], ghe_nome.upper(), bold=True, color=BRANCO, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
        row_h = tbl.add_row()
        shd(row_h.cells[0], VERDE_MED)
        shd(row_h.cells[1], VERDE_MED)
        txt(row_h.cells[0], 'FUNÇÃO', bold=True, color=BRANCO, size=9)
        txt(row_h.cells[1], 'EXAMES SOLICITADOS', bold=True, color=BRANCO, size=9)
        for cargo, rows_cargo in cargos_dict.items():
            exames_fmt = [_fmt_exame_rq61(r) for r in rows_cargo]
            primeira = True
            for exame_str in exames_fmt:
                row_ex = tbl.add_row()
                if primeira:
                    txt(row_ex.cells[0], cargo, bold=True, size=9)
                    primeira = False
                else:
                    row_ex.cells[0].text = ''
                set_borders(row_ex.cells[0])
                set_borders(row_ex.cells[1])
                txt(row_ex.cells[1], exame_str, size=9)
        doc.add_paragraph()
    p = doc.add_paragraph(f"Responsável pelo preenchimento: {cabecalho.get('responsavel_tec', '---')}\nMédico(a) Responsável pela validação: {medico}{(' CRM-GO ' + crm) if crm else ''}\nData do PCMAT/PGR: {vig_i}")
    p.runs[0].font.size = Pt(8)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

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
from utils.fuzzy_utils import normalizar_agente

_BANCO_V2_PATH = os.path.join("data", "banco_matrizes_v2.json")
try:
    with open(_BANCO_V2_PATH, "r", encoding="utf-8") as _f:
        _BANCO_MATRIZES_V2 = json.load(_f)
except FileNotFoundError:
    _BANCO_MATRIZES_V2 = {}

VERSAO_MODULO_PCMSO = '6.0 (Universal)'

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
    'OBRA', 'CANTEIRO', 'CONSTRUCAO', 'REFORMA', 'RESIDENCIAL',
    'EDIFICIO', 'BLOCO', 'TORRE', 'HETRIN', 'VIADUTO', 'PONTE', 'SHOPPING',
    'CONDOMINIO', 'EMPREENDIMENTO', 'MONTAGEM', 'INSTALACAO', 'CAMPO',
    'PRODUCAO', 'CREMALHEIRA', 'GRUA', 'ARMACAO', 'BETONEIRA', 'CARPINTARIA',
    'LIMPEZA', 'ELETRICA', 'PINTURA', 'GESSO', 'HIDROSSANITARIAS', 'SERRALHERIA', 
    'SINALIZACAO', 'MANUTENCAO', 'IMPERMEABILIZACAO', 'ESTRUTURA'
]

_PALAVRAS_ESCRITORIO = [
    'ESCRITORIO', 'SEDE', 'CORPORATIVO', 'ADMINISTRACAO', 'MARKETING',
    'TECNOLOGIA DA INFORMACAO', 'RECURSOS HUMANOS', 'FINANCEIRO',
    'CONTABILIDADE', 'JURIDICO', 'COMERCIAL', 'SAUDE', 'CLINICA', 'AMBULATORIO',
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
    'METILETILCETONA': 'METIL-ETIL-CETONA',
    'METILETILCETONA (MEK)': 'METIL-ETIL-CETONA',
    'MEK': 'METIL-ETIL-CETONA',
    'N-HEXANO': 'N-HEXANO',
    'TOLUENO': 'TOLUENO',
    'XILENO': 'XILENO',
    # MANTENHA A INTEGRIDADE QUÍMICA
    'TRICLOROETILENO': 'TRICLOROETILENO',
    'TRICLOROETENO': 'TRICLOROETILENO', # Tricloroeteno é sinônimo direto de Tricloroetileno
    'TRICLOROETANO': '1,1,1-TRICLOROETANO',
    'METILETILCETONA': 'METIL-ETIL-CETONA',
}

_RE_GHE = re.compile(r'(?:GHE[\s:\.\-]*\d|GRUPO\s+HOMOGENEO|LOCAL\s+DE\s+TRABALHO\s*:\s*\w|SETOR\s*:\s*\w)', re.IGNORECASE)
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
    'ESPIROMETRIA': 'Espirometria',
    'ANTI HBS HBSAG ANTI HCV': 'Anti-HBs + HBsAg + Anti-HCV',
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

def extrair_texto_pdf(uploaded_file):
    texto = []
    with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
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
        if not lc: continue
        lu = normalizar_texto(lc)
        if _is_linha_ghe(lc) and len(lc) < 120 and len(lc.strip()) >= 4 and not lc.strip().endswith('.'):
            if ghe_atual and (ghe_atual['cargos'] or ghe_atual['riscos_mapeados']):
                ghes.append(ghe_atual)
            nome_ghe_limpo = re.split(r'\s+CMO\b|\s+[–\-]\s+CMO|\s+SPE\b|\s+LTDA\b', lc, flags=re.IGNORECASE)[0].strip()
            ghe_atual = {'ghe': nome_ghe_limpo, 'cargos': [], 'riscos_mapeados': []}
            agentes_set = set()
            continue
        if ghe_atual is None: continue
        if not any(normalizar_texto(exc) in lu for exc in PALAVRAS_EXCLUIR_CARGO):
            for cargo in MAPA_CARGOS_CONHECIDOS:
                if normalizar_texto(cargo) in lu and cargo not in ghe_atual['cargos']:
                    ghe_atual['cargos'].append(cargo)
                    break
        for palavra, chave_risco in _MAPA_AGENTES.items():
            if normalizar_texto(palavra) in lu and chave_risco not in agentes_set:
                agentes_set.add(chave_risco)
                ghe_atual['riscos_mapeados'].append({'nome_agente': chave_risco, 'perigo_especifico': lc[:200]})
        # Fuzzy matching: tenta normalizar palavras da linha que não bateram no _MAPA_AGENTES
        for token in lu.split():
            if len(token) < 3:
                continue
            nome_normalizado = normalizar_agente(token)
            if nome_normalizado != token:
                chave_risco = _MAPA_AGENTES.get(nome_normalizado.upper(), nome_normalizado.upper())
                if chave_risco not in agentes_set:
                    agentes_set.add(chave_risco)
                    ghe_atual['riscos_mapeados'].append({'nome_agente': chave_risco, 'perigo_especifico': lc[:200]})
    if ghe_atual and (ghe_atual['cargos'] or ghe_atual['riscos_mapeados']):
        ghes.append(ghe_atual)
    return _deduplicar_ghes(ghes)

def _deduplicar_ghes(ghes):
    vistos = {}
    resultado = []
    for ghe in ghes:
        nome_ghe = ghe.get('ghe', '').strip().upper()
        if not nome_ghe:
            resultado.append(ghe)
            continue
            
        if nome_ghe not in vistos:
            vistos[nome_ghe] = len(resultado)
            resultado.append(ghe)
        else:
            idx = vistos[nome_ghe]
            for c in ghe.get('cargos', []):
                if c not in resultado[idx]['cargos']:
                    resultado[idx]['cargos'].append(c)
            
            riscos_existentes = {r['nome_agente'] for r in resultado[idx]['riscos_mapeados']}
            for r in ghe.get('riscos_mapeados', []):
                if r['nome_agente'] not in riscos_existentes:
                    resultado[idx]['riscos_mapeados'].append(r)
                    riscos_existentes.add(r['nome_agente'])
    return resultado

def extrair_pgr_com_fallback(texto_pgr, chave_api=None):
    local = extrair_pgr_local(texto_pgr)
    return local, 'local'

def _novo_exame(exame, adm=True, per=None, mro=True, rt=False, dem=False, obs='', motivo=''):
    return {'exame': _nome_oficial_exame(exame), 'adm': adm, 'per': per, 'mro': mro, 'rt': rt, 'dem': dem, 'obs': obs, 'motivo': motivo}

def _match_funcao_matriz(cargo_upper, funcao_matriz):
    alvo = _norm(funcao_matriz)
    cargo_n = _norm(cargo_upper)
    return alvo == cargo_n or alvo in cargo_n or cargo_n in alvo

def _forcar_regras_universais(ex, cargo_norm, riscos=None):
    """
    Garante as flags e periodicidades corretas da NR-7 e Matriz Seconci,
    agora com Inteligência Baseada em Riscos.
    """
    if riscos is None: riscos = []
    ex = deepcopy(ex)
    nome = _nome_oficial_exame(ex.get('exame', ''))
    ex['exame'] = nome
    
    cargo_upper = cargo_norm.upper()

    # Concatena todos os riscos para busca inteligente de palavras-chave
    texto_riscos = " ".join([normalizar_texto(r.get('nome_agente', '') + ' ' + r.get('perigo_especifico', '')) for r in riscos])
    tem_quimico = any(q in texto_riscos for q in ['TOLUENO', 'XILENO', 'BENZENO', 'HEXANO', 'TRICLOROETILENO', 'CETONA', 'SOLVENTE', 'TINTA', 'QUIMICO'])

    # 1. Base mínima e Regra de 6 Meses para Químicos Pesados
    if nome == 'Exame Clinico':
        ex['adm'], ex['mro'], ex['rt'], ex['dem'] = True, True, True, True
        
        # Reduz para 6 meses se houver risco químico mapeado ou função de pintura/impermeabilização
        if tem_quimico or any(c in cargo_upper for c in ['PINTOR', 'IMPERMEABILIZADOR']):
            ex['per'] = '6'
        elif not ex.get('per'): 
            ex['per'] = '12'

    # 2. Exames Pulmonares e Auditivos
    elif nome in ['Audiometria', 'Espirometria', 'RX de Tórax OIT']:
        ex['adm'], ex['mro'], ex['dem'] = True, True, True
        ex['rt'] = False  # Retorno ao trabalho é False para complementares
        
        if not ex.get('per'):
            if nome == 'Audiometria': ex['per'] = '12'
            elif nome == 'Espirometria': ex['per'] = '24'
            elif nome == 'RX de Tórax OIT':
                # Regra Inteligente: Sílica/Cimento no GHE = 12 meses. Restante = 60 meses.
                tem_poeira_pesada = any(x in texto_riscos for x in ['SILICA', 'CIMENTO', 'ASBESTO', 'POEIRA MINERAL'])
                if tem_poeira_pesada or any(x in cargo_upper for x in ['PEDREIRO', 'BETONEIRA']):
                    ex['per'] = '12'
                else:
                    ex['per'] = '60'

    # 3. Sangue e Kit Operacional
    elif nome in ['Acuidade Visual', 'ECG', 'Glicemia em Jejum', 'Hemograma', 'Hemograma Completo']:
        ex['adm'], ex['mro'] = True, True
        ex['rt'], ex['dem'] = False, False 
        
        # Exceção: Hemograma para químicos tem validade menor e exige demissional
        if nome in ['Hemograma', 'Hemograma Completo'] and (tem_quimico or any(c in cargo_upper for c in ['PINTOR', 'IMPERMEABILIZADOR'])):
            ex['per'] = '6'
            ex['dem'] = True
        elif not ex.get('per'): 
            ex['per'] = '12'
            
    # 4. Sangue e Urina Tóxicos (Restrições NR-7) - Mantendo a blindagem de 6 meses
    elif nome in {'Ácido tricloroacético na urina', 'Acetona na urina', 'Metil-Etil-Cetona', 
                  'Metiletilcetona na urina', 'Ciclohexanol na urina', 'Tetrahidrofurnano na urina', 
                  'Carboxiemoglobina', 'Ácido trans-trans mucônico', 'Ortocresol na urina', 
                  'Ác. Metil-hipúrico na urina', '2,5 Hexanodiona na Urina'}:
        ex['adm'], ex['mro'], ex['rt'], ex['dem'] = False, False, False, False
        if not ex.get('per'): ex['per'] = '6' 

    elif nome == 'Manganês sanguíneo':
        ex['adm'], ex['mro'], ex['rt'], ex['dem'] = True, True, False, False

    # 5. Exceções de Cargo
    if cargo_upper in ['OPERADOR DE GRUA', 'GRUEIRO']:
        if nome == 'Audiometria':
            ex['per'] = None
            ex['rt'] = False
            ex['dem'] = False

    return ex

def _aplicar_funcao_matriz(exames, cargo_norm, riscos):
    for funcao, lista_ex in MATRIZ_FUNCAO_EXAME.items():
        if _match_funcao_matriz(cargo_norm, funcao):
            for ex in lista_ex:
                exame = _novo_exame(
                    ex.get('exame', ''), adm=ex.get('adm', True), per=ex.get('per'),
                    mro=ex.get('mro', True), rt=ex.get('rt', False), dem=ex.get('dem', False),
                    obs=ex.get('obs', ''), motivo=f'Matriz de Função: {funcao.title()}',
                )
                # Passando 'riscos' para a inteligência de validação
                adicionar_exame_dedup(exames, _forcar_regras_universais(exame, cargo_norm, riscos))

def _aplicar_riscos_matriz(exames, riscos, cargo_norm):
    bio_real = tem_risco_biologico_real(riscos)
    for risco in riscos:
        chave_r = normalizar_texto(risco.get('nome_agente', ''))
        cas_r = str(risco.get('cas', '')).strip()
        
        # ====================================================================
        # NOVO MOTOR CAS-DRIVEN: Aciona exames com base no número exato!
        # ====================================================================
        
        # Tricloroetileno (79-01-6) e 1,1,1-Tricloroetano (71-55-6)
        if cas_r in ['79-01-6', '71-55-6'] or chave_r in ['TRICLOROETILENO', '1,1,1-TRICLOROETANO', 'TRICLOROETANO']:
            exame = _novo_exame('Ácido tricloroacético na urina', adm=False, per='6', mro=False, rt=False, dem=False, motivo=f"IBE NR-07 (CAS: {cas_r or chave_r})")
            adicionar_exame_dedup(exames, _forcar_regras_universais(exame, cargo_norm))
            continue 
            
        # n-Hexano (110-54-3)
        elif cas_r == '110-54-3' or chave_r in ['N-HEXANO', 'HEXANO']:
            exame = _novo_exame('2,5 Hexanodiona na Urina', adm=False, per='6', mro=False, rt=False, dem=False, motivo=f"IBE NR-07 (CAS: {cas_r or chave_r})")
            adicionar_exame_dedup(exames, _forcar_regras_universais(exame, cargo_norm))
            continue
            
        # Tolueno (108-88-3)
        elif cas_r == '108-88-3' or chave_r == 'TOLUENO':
            exame = _novo_exame('Ortocresol na urina', adm=False, per='6', mro=False, rt=False, dem=False, motivo=f"IBE NR-07 (CAS: {cas_r or chave_r})")
            adicionar_exame_dedup(exames, _forcar_regras_universais(exame, cargo_norm))
            continue
            
        # ====================================================================

        if chave_r in CHAVES_BIOLOGICAS_MATRIZ and not bio_real:
            continue
            
        regra = MATRIZ_RISCO_EXAME.get(chave_r)
        if not regra:
            continue
            
        exame = _novo_exame(
            regra['exame'], adm=regra.get('adm', True), per=regra.get('per'), mro=regra.get('mro', True),
            rt=regra.get('rt', False), dem=regra.get('dem', False), obs=regra.get('obs', ''),
            motivo=f"Risco Mapeado: {chave_r.title()}",
        )
        # Passando 'riscos' para a inteligência de validação
        adicionar_exame_dedup(exames, _forcar_regras_universais(exame, cargo_norm, riscos))

def _ordenar_exames(rows):
    ordem = ['Exame Clinico', 'Audiometria', 'Acuidade Visual', 'Hemograma', 'Glicemia em Jejum', 'ECG', 'Anti-HBs + HBsAg + Anti-HCV', 'Ácido tricloroacético na urina', 'Acetona na urina', 'Metil-Etil-Cetona', 'Ciclohexanol na urina', 'Tetrahidrofurnano na urina', 'Manganês sanguíneo', 'Carboxiemoglobina', 'Contagem de Reticulócitos', 'Ácido trans-trans mucônico', 'Ortocresol na urina', 'Ác. Metil-hipúrico na urina', 'Avaliação Psicossocial', 'Espirometria', 'RX de coluna lombo-sacra', 'RX de Tórax OIT']
    peso = {nome: i for i, nome in enumerate(ordem)}
    return sorted(rows, key=lambda r: (peso.get(_nome_oficial_exame(r['Exame']), 999), _norm(r['Exame'])))

# =====================================================================
# 1. COLE A FUNÇÃO NOVA AQUI (ACIMA DO PROCESSAR_PCMSO)
# =====================================================================
def enriquecer_pgr_com_fispq(dados_pgr, resultados_medicos_fispq):
    """
    Cruza os GHEs do PGR com os agentes químicos descobertos pelo Módulo de FISPQ
    e injeta os riscos químicos ocultos para gerar os exames de sangue/urina.
    """
    if not resultados_medicos_fispq:
        return dados_pgr
        
    for ghe_pgr in dados_pgr:
        nome_ghe_pgr = normalizar_texto(ghe_pgr.get('ghe', ''))
        
        # Procura os agentes da FISPQ que foram mapeados para este GHE
        agentes_para_injetar = []
        for item_fispq in resultados_medicos_fispq:
            nome_ghe_fispq = normalizar_texto(item_fispq.get('GHE', ''))
            
            # Se o GHE da FISPQ bater com o GHE do PGR
            if nome_ghe_fispq in nome_ghe_pgr or nome_ghe_pgr in nome_ghe_fispq:
                agentes_para_injetar.append(item_fispq)
        
        # Injeta os riscos no GHE do PGR
        riscos_existentes = {normalizar_texto(r['nome_agente']) for r in ghe_pgr.get('riscos_mapeados', [])}
        
        # Modificamos a captura para pegar o dicionário inteiro do Módulo de Engenharia
        for item in agentes_para_injetar:
            agente = item.get('Agente Quimico', '')
            cas = item.get('N CAS', '') # <--- AGORA ESTAMOS PEGANDO O CAS!
            agente_norm = normalizar_texto(agente)
            
            if agente_norm not in riscos_existentes:
                ghe_pgr['riscos_mapeados'].append({
                    'nome_agente': agente, 
                    'cas': cas, # <--- INJETAMOS O CAS DENTRO DO PGR
                    'perigo_especifico': f'Mapeado via FISPQ (CAS: {cas})'
                })
                riscos_existentes.add(agente_norm)
                
    return dados_pgr

def processar_pcmso(dados_pgr, tipo_ambiente='misto'):
    linhas = []

    for ghe in dados_pgr:
        nome_ghe_raw = ghe.get('ghe', 'Sem GHE')
        nome_ghe = _limpar_nome_ghe(str(nome_ghe_raw))
        cargos = ghe.get('cargos') or []
        riscos = ghe.get('riscos_mapeados') or []

        if not _ghe_valido(nome_ghe):
            continue

        if tipo_ambiente == 'canteiro':
            e_canteiro = True
        elif tipo_ambiente == 'escritorio':
            e_canteiro = False
        else:
            e_canteiro = _ghe_e_canteiro_misto(nome_ghe, riscos)

        # ANÁLISE DE RISCOS DO GHE (A mágica que limpa o excesso de exames)
        texto_riscos_ghe = " ".join([normalizar_texto(r.get('nome_agente', '') + ' ' + r.get('perigo_especifico', '')) for r in riscos])
        tem_altura_confinado = any(x in texto_riscos_ghe for x in ['ALTURA', 'CONFINADO', 'ESPACO CONFINADO'])
        tem_eletricidade = 'ELETRIC' in texto_riscos_ghe

        for cargo in cargos:
            cargo_norm = normalizar_cargo(cargo)
            exames = {}

            # 1. Base: Todo mundo ganha o Clínico Universal
            clinico = _novo_exame('Exame Clinico', motivo='Obrigatório NR-07')
            adicionar_exame_dedup(exames, _forcar_regras_universais(clinico, cargo_norm, riscos))

            e_cargo_adm = any(adm in cargo_norm for adm in ['RH', 'SUPERINTENDENTE', 'RECEPCIONISTA', 'DIRETOR', 'ADVOGADO', 'JURIDICO'])
            if 'ADMINISTRATIVO' in cargo_norm and 'OBRA' not in cargo_norm:
                e_cargo_adm = True

            # Verifica se o cargo é de operação de máquina pesada
            operador_maquina = any(x in cargo_norm for x in ['OPERADOR', 'MOTORISTA', 'GUINDASTE', 'GRUA', 'EMPILHADEIRA'])

            if e_canteiro and not e_cargo_adm:
                # Kit Básico Respiratório / Auditivo
                adicionar_exame_dedup(exames, _forcar_regras_universais(_novo_exame('Audiometria', motivo='Base Canteiro/Ruído'), cargo_norm, riscos))
                adicionar_exame_dedup(exames, _forcar_regras_universais(_novo_exame('Espirometria', motivo='Base Canteiro/Poeira'), cargo_norm, riscos))
                adicionar_exame_dedup(exames, _forcar_regras_universais(_novo_exame('RX de Tórax OIT', motivo='Base Canteiro/Poeira'), cargo_norm, riscos))
                
                # KIT OPERACIONAL (Aplicado APENAS se houver os riscos críticos ou for operador)
                if tem_altura_confinado or tem_eletricidade or operador_maquina:
                    adicionar_exame_dedup(exames, _forcar_regras_universais(_novo_exame('Acuidade Visual', motivo='Op. Máquina/Altura/Elétrica'), cargo_norm, riscos))
                    adicionar_exame_dedup(exames, _forcar_regras_universais(_novo_exame('ECG', motivo='Op. Máquina/Altura/Elétrica'), cargo_norm, riscos))
                    adicionar_exame_dedup(exames, _forcar_regras_universais(_novo_exame('Glicemia em Jejum', motivo='Op. Máquina/Altura/Elétrica'), cargo_norm, riscos))
                    adicionar_exame_dedup(exames, _forcar_regras_universais(_novo_exame('Hemograma', motivo='Op. Máquina/Altura/Elétrica'), cargo_norm, riscos))

                # AVALIAÇÃO PSICOSSOCIAL (Aplicada APENAS para Altura ou Espaço Confinado)
                if tem_altura_confinado:
                    adicionar_exame_dedup(exames, _forcar_regras_universais(_novo_exame('Avaliação Psicossocial', motivo='NR-35 / NR-33'), cargo_norm, riscos))

            # 3 e 4. Cruzamento com as Matrizes e FISPQ
            _aplicar_funcao_matriz(exames, cargo_norm, riscos)
            _aplicar_riscos_matriz(exames, riscos, cargo_norm)

            # Preparar as linhas finais
            rows_cargo = []
            for ex_info in exames.values():
                nome_exame = _nome_oficial_exame(ex_info.get('exame', ''))
                rt = bool(ex_info.get('rt', False))
                dem = bool(ex_info.get('dem', False))

                rows_cargo.append({
                    'GHE / Setor': nome_ghe,
                    'Cargo': cargo,
                    'Exame': nome_exame,
                    'ADM': _flag(ex_info.get('adm', True)),
                    'PER': _fmt_per(ex_info.get('per')),
                    'MRO': _flag(ex_info.get('mro', True)),
                    'RT': _flag(rt),
                    'DEM': _flag(dem),
                    'Justificativa': ex_info.get('motivo', ''),
                })

            linhas.extend(_ordenar_exames(rows_cargo))

    return pd.DataFrame(linhas)

def gerar_html_pcmso(df, cabecalho=None):
    if not cabecalho: cabecalho = {}
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

    if not cabecalho: cabecalho = {}
    razao = cabecalho.get('razao_social', 'Empresa não informada')
    obra = cabecalho.get('obra', '---')
    medico = cabecalho.get('medico_rt', 'Não informado')
    crm = cabecalho.get('crm', '')
    vig_i = cabecalho.get('vig_ini', '---')
    tipo = cabecalho.get('tipo_obra', 'Renovação')

    # ── CORREÇÃO 1: Paleta de cores alinhada ao template real RQ.61 Seconci-GO ──
    AZUL_GHE  = '4472C4'   # fundo linha do GHE (azul médio)
    CINZA_COL = 'D9D9D9'   # fundo cabeçalho FUNÇÃO / EXAMES SOLICITADOS
    BRANCO    = RGBColor(0xFF, 0xFF, 0xFF)
    PRETO     = RGBColor(0x00, 0x00, 0x00)

    def shd(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        s = OxmlElement('w:shd')
        s.set(qn('w:val'), 'clear')
        s.set(qn('w:color'), 'auto')
        s.set(qn('w:fill'), hex_color)
        tcPr.append(s)

    def set_borders(cell, color='000000'):
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

    # ── Cabeçalho do documento ──
    cab = doc.add_table(rows=4, cols=4)
    cab.style = 'Table Grid'
    cab.rows[0].cells[0].merge(cab.rows[0].cells[3])
    shd(cab.rows[0].cells[0], '084D22')
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

        # ── Linha do GHE: azul médio (#4472C4) ──
        row_ghe = tbl.add_row()
        row_ghe.cells[0].merge(row_ghe.cells[1])
        shd(row_ghe.cells[0], AZUL_GHE)
        set_borders(row_ghe.cells[0])
        txt(row_ghe.cells[0], ghe_nome.upper(), bold=True, color=BRANCO, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)

        # ── CORREÇÃO 2: Cabeçalho FUNÇÃO/EXAMES: cinza (#D9D9D9) com texto preto ──
        row_h = tbl.add_row()
        shd(row_h.cells[0], CINZA_COL)
        shd(row_h.cells[1], CINZA_COL)
        set_borders(row_h.cells[0])
        set_borders(row_h.cells[1])
        txt(row_h.cells[0], 'FUNÇÃO', bold=True, color=PRETO, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        txt(row_h.cells[1], 'EXAMES SOLICITADOS', bold=True, color=PRETO, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

        # ── CORREÇÃO 3: Todos os exames do cargo em uma única célula ──
        for cargo, rows_cargo in cargos_dict.items():
            exames_fmt = [_fmt_exame_rq61(r) for r in rows_cargo]
            row_ex = tbl.add_row()
            set_borders(row_ex.cells[0])
            set_borders(row_ex.cells[1])
            txt(row_ex.cells[0], cargo, bold=True, size=9)
            cell_exames = row_ex.cells[1]
            cell_exames.text = ''
            p = cell_exames.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for i, exame_str in enumerate(exames_fmt):
                if i > 0:
                    p.add_run('\n')
                run = p.add_run(exame_str)
                run.font.size = Pt(9)

        doc.add_paragraph()

    p = doc.add_paragraph(f"Responsável pelo preenchimento: {cabecalho.get('responsavel_tec', '---')}\nMédico(a) Responsável pela validação: {medico}{(' CRM-GO ' + crm) if crm else ''}\nData do PCMAT/PGR: {vig_i}")
    p.runs[0].font.size = Pt(8)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

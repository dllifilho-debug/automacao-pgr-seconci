# =============================================================================
# AGENTE MÉDICO IA - Motor universal de PCMSO
# Versão: 1.0
# Raciocínio em 4 camadas para definir exames médicos ocupacionais
# Baseado na metodologia CMO/Dra. Patrícia Montalvo / Dra. Carolini Polesso
# =============================================================================

import json
import os
import re
from copy import deepcopy

# ---------------------------------------------------------------------------
# Mapa de sinonimos de cargos → chave-mestra do banco
# Quanto mais granular, melhor o resultado
# ---------------------------------------------------------------------------
MAPA_CARGO_CHAVE = {
    # Administrativos
    'engenheiro': 'ENGENHEIRO',
    'estagiario de engenharia': 'ESTAGIARIO',
    'estagiario': 'ESTAGIARIO',
    'tecnico de seguranca': 'TECNICO_SST',
    'tecnico de seguranca do trabalho': 'TECNICO_SST',
    'tecnico sst': 'TECNICO_SST',
    'estagiario de seguranca': 'ESTAGIARIO',
    'mestre de obra': 'MESTRE_OBRA',
    'mestre obra': 'MESTRE_OBRA',
    'encarregado de pedreiro': 'ENCARREGADO_GERAL',
    'encarregado de pintor': 'ENCARREGADO_GERAL',
    'encarregado de eletricista': 'ENCARREGADO_GERAL',
    'encarregado de encanador': 'ENCARREGADO_GERAL',
    'encarregado de serralheiro': 'ENCARREGADO_GERAL',
    'encarregado de carpinteiro': 'ENCARREGADO_GERAL',
    'encarregado de armador': 'ENCARREGADO_GERAL',
    'encarregado de impermeabilizacao': 'IMPERMEABILIZADOR',
    'encarregado impermeabilizacao': 'IMPERMEABILIZADOR',
    'encarregado': 'ENCARREGADO_SUPERVISAO',
    'administrativo de obra': 'AUXILIAR_ADMINISTRATIVO',
    'auxiliar administrativo': 'AUXILIAR_ADMINISTRATIVO',
    'jovem aprendiz': 'JOVEM_APRENDIZ',
    'almoxarife': 'ALMOXARIFE',
    # Operacionais canteiro
    'carpinteiro': 'CARPINTEIRO',
    'meio oficial de carpinteiro': 'CARPINTEIRO',
    'armador': 'ARMADOR',
    'meio oficial de armador': 'ARMADOR',
    'pedreiro': 'PEDREIRO',
    'meio oficial de pedreiro': 'PEDREIRO',
    'servente': 'SERVENTE_CANTEIRO',
    'servente de obra': 'SERVENTE_CANTEIRO',
    'pintor': 'PINTOR',
    'meio oficial de pintor': 'PINTOR',
    'encarregado de pintura': 'PINTOR',
    'serralheiro': 'SERRALHEIRO',
    'meio oficial de serralheiro': 'SERRALHEIRO',
    'eletricista': 'ELETRICISTA',
    'meio oficial de eletricista': 'ELETRICISTA',
    'eletricista industrial': 'ELETRICISTA_ENERGIZADO',
    'eletricista energizado': 'ELETRICISTA_ENERGIZADO',
    'encanador': 'ENCANADOR',
    'meio oficial de encanador': 'ENCANADOR',
    'gesseiro': 'GESSEIRO',
    'meio oficial de gesseiro': 'GESSEIRO',
    'impermeabilizador': 'IMPERMEABILIZADOR',
    'aplicador de asfalto': 'IMPERMEABILIZADOR',
    'aplicador asfalto impermeabilizante': 'IMPERMEABILIZADOR',
    'operador de betoneira': 'OPERADOR_BETONEIRA',
    'operador betoneira': 'OPERADOR_BETONEIRA',
    'operador de grua': 'OPERADOR_GRUA',
    'operador grua': 'OPERADOR_GRUA',
    'operador de cremalheira': 'OPERADOR_CREMALHEIRA',
    'operador de elevador de cremalheira': 'OPERADOR_CREMALHEIRA',
    'sinaleiro': 'SINALEIRO',
    'porteiro': 'PORTEIRO_VIGIA',
    'vigia': 'PORTEIRO_VIGIA',
    'motorista': 'MOTORISTA',
    'mecanico de manutencao': 'MECANICO_MANUTENCAO',
    'mecanico manutencao': 'MECANICO_MANUTENCAO',
}

# ---------------------------------------------------------------------------
# Mapa de risco / agente → exames adicionais (Camada 3 - NR-7 IBE)
# Baseado na planilha validada Dra. Patricia + NR-7 Anexos I e II
# ---------------------------------------------------------------------------
MATRIZ_RISCO_EXAME_IA = {
    'benzeno': [
        {'nome': 'Ácido Trans-trans Mucônico', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
        {'nome': 'Hemograma Completo', 'adm': True, 'per': '6', 'mro': True, 'ret': False, 'dem': True},
        {'nome': 'Contagem de Reticulócitos', 'adm': True, 'per': '6', 'mro': True, 'ret': False, 'dem': True},
    ],
    'tolueno': [
        {'nome': 'Ortocresol na urina', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'xileno': [
        {'nome': 'Ácido Metil-hipúrico na urina', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'estireno': [
        {'nome': 'Ácido Mandélico', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
        {'nome': 'Ácido Fenilglioxílico', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'manganês': [
        {'nome': 'Manganês Sanguíneo', 'adm': True, 'per': '6', 'mro': True, 'ret': False, 'dem': False},
    ],
    'manganes': [
        {'nome': 'Manganês Sanguíneo', 'adm': True, 'per': '6', 'mro': True, 'ret': False, 'dem': False},
    ],
    'monoxido de carbono': [
        {'nome': 'Carboxiemoglobina', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'monóxido de carbono': [
        {'nome': 'Carboxiemoglobina', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'tricloroetileno': [
        {'nome': 'Ácido tricloroacético na urina', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'fluoreto': [
        {'nome': 'Fluoreto Urinário', 'adm': True, 'per': '6', 'mro': True, 'ret': True, 'dem': True},
    ],
    'fluoreto de hidrogenio': [
        {'nome': 'Fluoreto Urinário', 'adm': True, 'per': '6', 'mro': True, 'ret': True, 'dem': True},
    ],
    'acetona': [
        {'nome': 'Acetona na urina', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'metil etil cetona': [
        {'nome': 'Metil-Etil-Cetona', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'n-hexano': [
        {'nome': '2,5-Hexanodiona na urina', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'chumbo': [
        {'nome': 'Chumbo Sanguíneo', 'adm': True, 'per': '6', 'mro': True, 'ret': True, 'dem': True},
        {'nome': 'Ác. Delta Amino Levulínico na urina (ALA-U)', 'adm': True, 'per': '6', 'mro': True, 'ret': True, 'dem': True},
    ],
    'mercurio': [
        {'nome': 'Mercúrio na urina', 'adm': False, 'per': '6', 'mro': False, 'ret': False, 'dem': False},
    ],
    'ruido': [
        {'nome': 'Audiometria', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': True},
    ],
    'ruído': [
        {'nome': 'Audiometria', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': True},
    ],
    'silica': [
        {'nome': 'RX de Tórax OIT', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': True},
        {'nome': 'Espirometria', 'adm': True, 'per': '24', 'mro': True, 'ret': False, 'dem': True},
    ],
    'sílica': [
        {'nome': 'RX de Tórax OIT', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': True},
        {'nome': 'Espirometria', 'adm': True, 'per': '24', 'mro': True, 'ret': False, 'dem': True},
    ],
    'espaco confinado': [
        {'nome': 'Avaliação Psicossocial', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
    ],
    'espaço confinado': [
        {'nome': 'Avaliação Psicossocial', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
    ],
    'trabalho em altura': [
        {'nome': 'Hemograma Completo', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
        {'nome': 'Glicemia em Jejum', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
        {'nome': 'ECG', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
        {'nome': 'Acuidade Visual', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
        {'nome': 'Avaliação Psicossocial', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
    ],
    'eletricidade': [
        {'nome': 'Acuidade Visual', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
        {'nome': 'ECG', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
    ],
    'poeira mineral': [
        {'nome': 'Espirometria', 'adm': True, 'per': '24', 'mro': True, 'ret': False, 'dem': True},
        {'nome': 'RX de Tórax OIT', 'adm': True, 'per': '60', 'mro': True, 'ret': False, 'dem': True},
    ],
    'nevoas': [
        {'nome': 'Espirometria', 'adm': True, 'per': '24', 'mro': True, 'ret': False, 'dem': True},
    ],
    'névoas': [
        {'nome': 'Espirometria', 'adm': True, 'per': '24', 'mro': True, 'ret': False, 'dem': True},
    ],
    'tinta': [
        {'nome': 'Espirometria', 'adm': True, 'per': '24', 'mro': True, 'ret': False, 'dem': True},
    ],
    'impermeabilizacao': [
        {'nome': 'Espirometria', 'adm': True, 'per': '24', 'mro': True, 'ret': False, 'dem': True},
    ],
    'impermeabilização': [
        {'nome': 'Espirometria', 'adm': True, 'per': '24', 'mro': True, 'ret': False, 'dem': True},
    ],
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalizar(texto: str) -> str:
    """Normaliza texto para comparação: lowercase, sem acentos simplificados."""
    if not texto:
        return ''
    return texto.lower().strip()


def _remover_acentos_simples(texto: str) -> str:
    subs = {
        'á': 'a', 'â': 'a', 'ã': 'a', 'à': 'a',
        'é': 'e', 'ê': 'e',
        'í': 'i', 'î': 'i',
        'ó': 'o', 'ô': 'o', 'õ': 'o',
        'ú': 'u', 'û': 'u',
        'ç': 'c',
    }
    for k, v in subs.items():
        texto = texto.replace(k, v)
    return texto


def _norm(texto: str) -> str:
    return _remover_acentos_simples(_normalizar(texto))


def _exame_ja_existe(lista: list, nome: str) -> bool:
    n = _norm(nome)
    return any(_norm(e.get('nome', '')) == n for e in lista)


def _merge_exame(lista: list, novo: dict) -> list:
    """Adiciona exame se não existir. Se existir, atualiza periodicidade para o menor valor."""
    n_novo = _norm(novo['nome'])
    for e in lista:
        if _norm(e.get('nome', '')) == n_novo:
            # Mantém a periodicidade menor (mais conservadora)
            per_atual = int(e.get('per') or 99)
            per_novo = int(novo.get('per') or 99)
            if per_novo < per_atual:
                e['per'] = str(per_novo)
            # OR booleanos — se qualquer fonte manda true, fica true
            for flag in ('adm', 'mro', 'ret', 'dem'):
                e[flag] = e.get(flag, False) or novo.get(flag, False)
            return lista
    lista.append(deepcopy(novo))
    return lista


# ---------------------------------------------------------------------------
# Carregamento do banco de perfis
# ---------------------------------------------------------------------------

_BANCO_CACHE = None


def carregar_banco() -> dict:
    global _BANCO_CACHE
    if _BANCO_CACHE is not None:
        return _BANCO_CACHE
    caminho = os.path.join(os.path.dirname(__file__), '..', 'data', 'banco_matrizes_v2.json')
    caminho = os.path.normpath(caminho)
    if os.path.exists(caminho):
        with open(caminho, 'r', encoding='utf-8') as f:
            _BANCO_CACHE = json.load(f)
    else:
        _BANCO_CACHE = {}
    return _BANCO_CACHE


# ---------------------------------------------------------------------------
# Camada 1 — Resolução de chave-mestra (cargo → perfil)
# ---------------------------------------------------------------------------

def resolver_chave_mestra(cargo: str) -> str:
    """
    Mapeia o nome do cargo para uma chave-mestra do banco_matrizes_v2.
    Retorna None se não houver mapeamento.
    """
    n = _norm(cargo)
    # 1. Tentativa direta
    if n in MAPA_CARGO_CHAVE:
        return MAPA_CARGO_CHAVE[n]
    # 2. Busca parcial (substring)
    for k, v in MAPA_CARGO_CHAVE.items():
        if k in n or n in k:
            return v
    # 3. Heurística por palavras-chave fortes
    tokens = set(n.split())
    if {'engenheiro'} & tokens:
        return 'ENGENHEIRO'
    if {'tecnico', 'seguranca'} <= tokens or {'tecnico', 'sst'} <= tokens:
        return 'TECNICO_SST'
    if {'mestre'} & tokens:
        return 'MESTRE_OBRA'
    if {'almoxarife', 'almoxarifado'} & tokens:
        return 'ALMOXARIFE'
    if {'porteiro', 'vigia', 'vigilante'} & tokens:
        return 'PORTEIRO_VIGIA'
    if {'carpinteiro'} & tokens:
        return 'CARPINTEIRO'
    if {'armador'} & tokens:
        return 'ARMADOR'
    if {'pedreiro'} & tokens:
        return 'PEDREIRO'
    if {'gesseiro'} & tokens:
        return 'GESSEIRO'
    if {'servente'} & tokens:
        return 'SERVENTE_CANTEIRO'
    if {'pintor'} & tokens:
        return 'PINTOR'
    if {'serralheiro'} & tokens:
        return 'SERRALHEIRO'
    if {'impermeabilizador', 'impermeabilizacao', 'asfalto', 'manta'} & tokens:
        return 'IMPERMEABILIZADOR'
    if {'eletricista', 'eletrico'} & tokens:
        if 'energizado' in n or 'industrial' in n:
            return 'ELETRICISTA_ENERGIZADO'
        return 'ELETRICISTA'
    if {'encanador', 'hidraulico', 'hidro'} & tokens:
        return 'ENCANADOR'
    if {'grua'} & tokens:
        return 'OPERADOR_GRUA'
    if {'betoneira'} & tokens:
        return 'OPERADOR_BETONEIRA'
    if {'cremalheira', 'elevador'} & tokens:
        return 'OPERADOR_CREMALHEIRA'
    if {'sinaleiro'} & tokens:
        return 'SINALEIRO'
    if {'motorista'} & tokens:
        return 'MOTORISTA'
    if {'mecanico', 'manutencao'} & tokens:
        return 'MECANICO_MANUTENCAO'
    if {'encarregado'} & tokens:
        return 'ENCARREGADO_GERAL'
    if {'administrativo', 'auxiliar', 'aprendiz', 'escriturario', 'recepcionist'} & tokens:
        return 'ADMINISTRATIVO'
    return None


# ---------------------------------------------------------------------------
# Camada 2 — Ajustes finos por tipo de GHE / contexto de obra
# ---------------------------------------------------------------------------

def _aplicar_ajustes_contexto(exames: list, contexto: dict) -> list:
    """
    Ajusta exames com base em flags de contexto extraídos do GHE/PGR.
    contexto pode ter: altura=True, confinado=True, eletricidade=True, maquinas_pesadas=True
    """
    if contexto.get('altura') or contexto.get('maquinas_pesadas'):
        for ex_novo in [
            {'nome': 'Hemograma Completo', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
            {'nome': 'Glicemia em Jejum', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
            {'nome': 'ECG', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
            {'nome': 'Acuidade Visual', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
        ]:
            exames = _merge_exame(exames, ex_novo)
    if contexto.get('eletricidade'):
        for ex_novo in [
            {'nome': 'Acuidade Visual', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
            {'nome': 'ECG', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
            {'nome': 'Audiometria', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
        ]:
            exames = _merge_exame(exames, ex_novo)
    if contexto.get('confinado'):
        exames = _merge_exame(exames, {
            'nome': 'Avaliação Psicossocial', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False
        })
    return exames


# ---------------------------------------------------------------------------
# Camada 3 — Injeção de IBE/NR-7 por agente químico
# ---------------------------------------------------------------------------

def _aplicar_riscos_quimicos(exames: list, riscos: list) -> list:
    """
    Para cada risco/agente identificado, injeta exames NR-7/IBE.
    riscos: lista de strings com nomes de agentes/riscos do PGR.
    """
    for risco in riscos:
        risco_n = _norm(risco)
        for chave, novos_exames in MATRIZ_RISCO_EXAME_IA.items():
            if chave in risco_n or risco_n in chave:
                for ex in novos_exames:
                    exames = _merge_exame(exames, ex)
    return exames


# ---------------------------------------------------------------------------
# Camada 4 — Validação universal (regras mínimas NR-7)
# ---------------------------------------------------------------------------

def _validacao_universal(exames: list, e_canteiro: bool = True) -> list:
    """
    Garante que regras mínimas da NR-7 estejam presentes.
    - Exame Clínico é SEMPRE obrigatório
    - Canteiro: garante Audiometria se não houver
    """
    if not _exame_ja_existe(exames, 'Exame Clínico'):
        exames.insert(0, {'nome': 'Exame Clínico', 'adm': True, 'per': '12', 'mro': True, 'ret': True, 'dem': True})
    if e_canteiro and not _exame_ja_existe(exames, 'Audiometria'):
        exames = _merge_exame(exames, {
            'nome': 'Audiometria', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False
        })
    return exames


# ---------------------------------------------------------------------------
# Motor principal: processar_cargo_ia
# ---------------------------------------------------------------------------

def processar_cargo_ia(
    cargo: str,
    riscos: list = None,
    contexto: dict = None,
    e_canteiro: bool = True,
) -> dict:
    """
    Motor principal do Agente Médico IA.
    Retorna dict com:
        exames: lista de exames
        chave_mestra: chave do banco utilizada (ou None)
        fonte_regra: 'banco_perfil' | 'agente_ia_sem_perfil'
        cargo_normalizado: cargo após normalização
    """
    if riscos is None:
        riscos = []
    if contexto is None:
        contexto = {}

    banco = carregar_banco()
    chave = resolver_chave_mestra(cargo)
    exames = []
    fonte = 'agente_ia_sem_perfil'

    # --- Camada 1: Perfil base do banco ---
    if chave and chave in banco:
        perfil = banco[chave]
        exames = deepcopy(perfil.get('exames', []))
        fonte = f'banco_perfil:{chave}'
    else:
        # Sem perfil: cria base mínima de canteiro ou administrativo
        if e_canteiro:
            exames = [
                {'nome': 'Exame Clínico', 'adm': True, 'per': '12', 'mro': True, 'ret': True, 'dem': True},
                {'nome': 'Audiometria', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': True},
                {'nome': 'Acuidade Visual', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
                {'nome': 'Hemograma Completo', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
                {'nome': 'Glicemia em Jejum', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
                {'nome': 'ECG', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': False},
            ]
        else:
            exames = [
                {'nome': 'Exame Clínico', 'adm': True, 'per': '12', 'mro': True, 'ret': True, 'dem': True},
            ]
        fonte = f'heuristica_base:sem_perfil_para_{cargo}'

    # --- Camada 2: Ajustes por contexto (altura, confinado, eletricidade) ---
    exames = _aplicar_ajustes_contexto(exames, contexto)

    # --- Camada 3: Injeção de riscos químicos / IBE NR-7 ---
    exames = _aplicar_riscos_quimicos(exames, riscos)

    # --- Camada 4: Validação universal ---
    exames = _validacao_universal(exames, e_canteiro=e_canteiro)

    return {
        'cargo': cargo,
        'cargo_normalizado': _norm(cargo),
        'chave_mestra': chave,
        'fonte_regra': fonte,
        'e_canteiro': e_canteiro,
        'exames': exames,
    }


# ---------------------------------------------------------------------------
# Interface: processar_ghe_ia
# ---------------------------------------------------------------------------

def processar_ghe_ia(ghe_nome: str, cargos: list, riscos_ghe: list = None, contexto_ghe: dict = None) -> list:
    """
    Processa todos os cargos de um GHE e retorna lista de resultados.
    """
    if riscos_ghe is None:
        riscos_ghe = []
    if contexto_ghe is None:
        contexto_ghe = {}

    resultados = []
    # Detectar se é GHE de canteiro ou administrativo pelo nome
    nome_n = _norm(ghe_nome)
    e_canteiro = not any(x in nome_n for x in [
        'administrativo', 'engenharia', 'planejamento', 'escritorio', 'gerencia', 'direcao'
    ])
    # Exceção: almoxarifado ainda é canteiro
    if 'almoxarifado' in nome_n:
        e_canteiro = True

    for cargo in cargos:
        resultado = processar_cargo_ia(
            cargo=cargo,
            riscos=riscos_ghe,
            contexto=contexto_ghe,
            e_canteiro=e_canteiro,
        )
        resultado['ghe'] = ghe_nome
        resultados.append(resultado)

    return resultados

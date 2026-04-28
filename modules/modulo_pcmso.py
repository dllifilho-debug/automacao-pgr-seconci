# =============================================================================
# MÓDULO PCMSO - Motor de geração de Programa de Controle Médico de Saúde
# Versão: 7.0 (AgenteMedicoIA + Banco v2 completo)
# Integra o AgenteMedicoIA como motor principal
# =============================================================================

import json
import os
import re
import unicodedata
from copy import deepcopy

VERSAO_MODULO_PCMSO = '7.0 (AgenteMedicoIA Universal)'

# Importa o Agente Médico IA
try:
    from modules.agente_medico_ia import processar_ghe_ia, processar_cargo_ia, carregar_banco
    _AGENTE_IA_DISPONIVEL = True
except ImportError:
    try:
        from agente_medico_ia import processar_ghe_ia, processar_cargo_ia, carregar_banco
        _AGENTE_IA_DISPONIVEL = True
    except ImportError:
        _AGENTE_IA_DISPONIVEL = False

# ---------------------------------------------------------------------------
# Imports de dados legados (mantidos para compatibilidade e fallback)
# ---------------------------------------------------------------------------
try:
    from data.matriz_exames import (
        MATRIZ_RISCO_EXAME,
        MATRIZ_FUNCAO_EXAME,
        MAPA_CARGOS_CONHECIDOS,
        DICIONARIO_CARGOS,
    )
except ImportError:
    MATRIZ_RISCO_EXAME = {}
    MATRIZ_FUNCAO_EXAME = {}
    MAPA_CARGOS_CONHECIDOS = {}
    DICIONARIO_CARGOS = {}

try:
    from data.dicionario_cas import DICIONARIO_CAS
except ImportError:
    DICIONARIO_CAS = {}


# ---------------------------------------------------------------------------
# Helpers de normalização
# ---------------------------------------------------------------------------

def _normalizar_texto(texto: str) -> str:
    if not texto:
        return ''
    nfkd = unicodedata.normalize('NFKD', str(texto))
    ascii_str = nfkd.encode('ASCII', 'ignore').decode('ASCII')
    return ascii_str.lower().strip()


def _exame_existe(lista: list, nome: str) -> bool:
    n = _normalizar_texto(nome)
    return any(_normalizar_texto(e.get('nome', '')) == n for e in lista)


def _merge_exame_pcmso(lista: list, novo: dict) -> list:
    n_novo = _normalizar_texto(novo['nome'])
    for e in lista:
        if _normalizar_texto(e.get('nome', '')) == n_novo:
            per_atual = int(e.get('per') or 99)
            per_novo = int(novo.get('per') or 99)
            if per_novo < per_atual:
                e['per'] = str(per_novo)
            for flag in ('adm', 'mro', 'ret', 'dem'):
                e[flag] = e.get(flag, False) or novo.get(flag, False)
            return lista
    lista.append(deepcopy(novo))
    return lista


# ---------------------------------------------------------------------------
# Extração de contexto de risco a partir de texto do GHE
# ---------------------------------------------------------------------------

def _extrair_contexto_ghe(texto_ghe: str) -> dict:
    t = _normalizar_texto(texto_ghe)
    return {
        'altura': any(x in t for x in ['altura', 'nr-35', 'nr35', 'telhado', 'andaime', 'cremalheira', 'grua']),
        'confinado': any(x in t for x in ['confinado', 'espaco confinado', 'cisterna', 'poco']),
        'eletricidade': any(x in t for x in ['eletric', 'nr-10', 'nr10', 'energizado', 'choque']),
        'maquinas_pesadas': any(x in t for x in ['maquina', 'guindaste', 'retroescavadeira', 'betoneira', 'grua', 'cremalheira']),
    }


# ---------------------------------------------------------------------------
# Processar PCMSO via Agente Médico IA (função principal)
# ---------------------------------------------------------------------------

def processar_pcmso(ghe_list: list, texto_pgr: str = '', nome_empresa: str = '', nome_obra: str = '') -> dict:
    """
    Gera a matriz de exames do PCMSO para uma lista de GHEs.
    
    Parâmetros:
        ghe_list: lista de dicts com { 'ghe': str, 'cargos': list[str], 'riscos': list[str] }
        texto_pgr: texto bruto extraído do PGR (para contexto adicional)
        nome_empresa: nome da empresa
        nome_obra: nome da obra
    
    Retorna:
        dict com 'matriz' (lista de exames por cargo/GHE) e 'resumo' (metadados)
    """
    if not _AGENTE_IA_DISPONIVEL:
        return _processar_pcmso_fallback(ghe_list)

    matriz_final = []
    resumo = {
    'versao_modulo': VERSAO_MODULO_PCMSO,
    'total_ghe': len(ghe_list),
    'total_cargos': 0,
    'cargos_banco_perfil': [],
    'cargos_heuristica': [],
    'agente_ia_ativo': True,
    }

    for ghe_item in ghe_list:
        nome_ghe = ghe_item.get('ghe', 'GHE sem nome')
        cargos = ghe_item.get('cargos', [])
        riscos = ghe_item.get('riscos', [])

        # Contexto extraído do nome do GHE + riscos
        contexto = _extrair_contexto_ghe(nome_ghe + ' ' + ' '.join(riscos))

        # Processa via Agente Médico IA
        resultados = processar_ghe_ia(
            ghe_nome=nome_ghe,
            cargos=cargos,
            riscos_ghe=riscos,
            contexto_ghe=contexto,
        )

        for resultado in resultados:
            fonte = resultado.get('fonte_regra', '')
            cargo = resultado.get('cargo', '')
            exames = resultado.get('exames', [])

            # Enriquece com IBE por CAS se disponível
            exames = _enriquecer_por_cas(exames, riscos)

            entrada_matriz = {
                'ghe': nome_ghe,
                'cargo': cargo,
                'chave_mestra': resultado.get('chave_mestra', ''),
                'fonte_regra': fonte,
                'exames': exames,
            }
            matriz_final.append(entrada_matriz)
            resumo['total_cargos'] += 1

            if 'banco_perfil' in fonte:
                resumo['cargos_banco_perfil'].append(f"{nome_ghe} | {cargo}")
            else:
                resumo['cargos_heuristica'].append(f"{nome_ghe} | {cargo} ({fonte})")

    return {
        'matriz': matriz_final,
        'resumo': resumo,
        'empresa': nome_empresa,
        'obra': nome_obra,
    }


# ---------------------------------------------------------------------------
# Enriquecimento por número CAS (FISPQ/FDS)
# ---------------------------------------------------------------------------

def _enriquecer_por_cas(exames: list, riscos: list) -> list:
    """
    Tenta enriquecer exames com IBE por número CAS identificado nos riscos.
    """
    for risco in riscos:
        # Procura padrão CAS: XXXX-XX-X ou XXXXXXX-XX-X
        cas_matches = re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', str(risco))
        for cas in cas_matches:
            if cas in DICIONARIO_CAS:
                for ex_novo in DICIONARIO_CAS[cas].get('exames', []):
                    exames = _merge_exame_pcmso(exames, ex_novo)
    return exames


# ---------------------------------------------------------------------------
# Fallback: lógica legada (mantida para compatibilidade)
# ---------------------------------------------------------------------------

def _processar_pcmso_fallback(ghe_list: list) -> dict:
    """
    Fallback para quando o AgenteMedicoIA não está disponível.
    Usa a lógica heurística anterior.
    """
    matriz = []
    for ghe_item in ghe_list:
        nome_ghe = ghe_item.get('ghe', '')
        for cargo in ghe_item.get('cargos', []):
            matriz.append({
                'ghe': nome_ghe,
                'cargo': cargo,
                'chave_mestra': None,
                'fonte_regra': 'fallback_legado',
                'exames': [
                    {'nome': 'Exame Clínico', 'adm': True, 'per': '12', 'mro': True, 'ret': True, 'dem': True},
                    {'nome': 'Audiometria', 'adm': True, 'per': '12', 'mro': True, 'ret': False, 'dem': True},
                ],
            })
    return {
        'matriz': matriz,
        'resumo': {'versao_modulo': VERSAO_MODULO_PCMSO, 'agente_ia_ativo': False},
    }


# ---------------------------------------------------------------------------
# Utilitário: resumo de cobertura do banco
# ---------------------------------------------------------------------------

def relatorio_cobertura(resultado_pcmso: dict) -> str:
    """
    Gera um texto de diagnóstico sobre a cobertura do banco de perfis.
    Útil para debugging e melhoria contínua.
    """
    resumo = resultado_pcmso.get('resumo', {})
    total = resumo.get('total_cargos', 0)
    banco = len(resumo.get('cargos_banco_perfil', []))
    heur = len(resumo.get('cargos_heuristica', []))
    pct = round(banco / total * 100, 1) if total else 0

    linhas = [
        f"=== RELATÓRIO DE COBERTURA - {resumo.get('versao_modulo', '')} ===",
        f"Total de GHEs processados: {resumo.get('total_ghe', 0)}",
        f"Total de cargos processados: {total}",
        f"Cargos resolvidos pelo banco de perfis: {banco} ({pct}%)",
        f"Cargos em modo heurístico/sem perfil: {heur} ({round(100-pct,1)}%)",
        "",
    ]
    if resumo.get('cargos_heuristica'):
        linhas.append("⚠ Cargos sem perfil dedicado (usar para enriquecer o banco):")
        for c in resumo['cargos_heuristica']:
            linhas.append(f"  - {c}")
    else:
        linhas.append("✓ Todos os cargos resolvidos pelo banco de perfis!")
    return '\n'.join(linhas)

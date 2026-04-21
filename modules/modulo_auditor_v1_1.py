import json
import re
import unicodedata
from pathlib import Path
from collections import defaultdict


def _strip(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto or ''))
                   if unicodedata.category(c) != 'Mn')


def norm(texto):
    return re.sub(r'\s+', ' ', _strip(texto).upper().strip())


def normalizar_exame(nome):
    s = norm(nome)
    mapa = {
        'EXAME CLINICO': 'Exame Clinico',
        'EXAME CLINICO SEMESTRAL': 'Exame Clinico',
        'AUDIOMETRIA': 'Audiometria',
        'AUDIOMETRIA TONAL': 'Audiometria',
        'ACUIDADE VISUAL': 'Acuidade Visual',
        'HEMOGRAMA': 'Hemograma',
        'HEMOGRAMA COMPLETO': 'Hemograma',
        'GLICEMIA EM JEJUM': 'Glicemia em Jejum',
        'ECG': 'ECG',
        'ESPIROMETRIA': 'Espirometria',
        'ESPIROMETRIA (SOMENTE)': 'Espirometria (somente)',
        'RX DE TORAX OIT': 'RX de Tórax OIT',
        'RAIO X TORAX OIT': 'RX de Tórax OIT',
        'RX TORAX OIT': 'RX de Tórax OIT',
        'RX DE COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
        'AVALIACAO PSICOSSOCIAL': 'Avaliação Psicossocial',
        'CARBOXIEMOGLOBINA': 'Carboxiemoglobina',
        'CARBOXIHEMOGLOBINA': 'Carboxiemoglobina',
        'MANGANES SANGUINEO': 'Manganês sanguíneo',
        'CONTAGEM DE RETICULOCITOS': 'Contagem de Reticulócitos',
        'ACIDO TRANS TRANS MUCONICO': 'Ácido trans-trans mucônico',
        'ACIDO TRICLOROACETICO NA URINA': 'Ácido tricloroacético na urina',
        'ACETONA NA URINA': 'Acetona na urina',
        'ORTOCRESOL NA URINA': 'Ortocresol na urina',
        'METIL ETIL CETONA': 'Metil-Etil-Cetona',
        'ACIDO METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
        'AC. METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
        'EPF (COPROPARASITOLOGICO) + ANTI-HBS': 'EPF (Coproparasitológico) + Anti-HBs',
    }
    return mapa.get(s, nome.strip())


def normalizar_cargo(nome):
    s = norm(nome)
    aliases = {
        # Já existentes
        'SERVENTE DE ARMADOR': 'Servente',
        'SERVENTE DE CARPINTEIRO': 'Servente',
        'MEIO OFICIAL DE PEDREIRO': 'Meio Oficial de Pedreiro',
        'OPERADOR DE BETONEIRA': 'Operador de Betoneira',
        'OPERADOR DE GRUA': 'Operador de Grua',
        'OPERADOR DE CREMALHEIRA': 'Operador de Cremalheira',
        # ✅ Novos — casos identificados na matriz RQ.61
        'MESTRE DE OBRA': 'Mestre de Obra',
        'MESTRE DE OBRAS': 'Mestre de Obra',
        'ELETRICISTA INDUSTRIAL': 'Eletricista',  # GHE 06 unifica
        'SERRALHEIRO': 'Serralheiro',
        'PINTOR': 'Pintor',
        'SERVENTE': 'Servente',
        'ARMADOR': 'Armador',
    }
    return aliases.get(s, nome.strip().title())


def carregar_banco_matrizes(caminho):
    with Path(caminho).open('r', encoding='utf-8') as f:
        return json.load(f)


def _norm_bool(v):
    if isinstance(v, bool):
        return v
    return norm(str(v)) in {'X', 'SIM', 'TRUE', '1'}


def _norm_per(v):
    if v is None:
        return None
    m = re.search(r'(\d+)', str(v))
    return m.group(1) if m else None


def _exam_obj(nome, dados=None):
    d = dados or {}
    return {
        'nome': normalizar_exame(nome),
        'adm': _norm_bool(d.get('adm')),
        'per': _norm_per(d.get('per')),
        'mro': _norm_bool(d.get('mro')),
        'ret': _norm_bool(d.get('ret')),
        'dem': _norm_bool(d.get('dem')),
    }


def _lista_de_exames(payload):
    if isinstance(payload, list):
        out = []
        for item in payload:
            if isinstance(item, str):
                out.append(_exam_obj(item))
            elif isinstance(item, dict):
                nome = item.get('nome') or item.get('Exame') or item.get('exame') or ''
                out.append(_exam_obj(nome, item))
        return out
    if isinstance(payload, dict):
        return [_exam_obj(payload.get('nome', ''), payload)]
    return []


def auditar_pcmso(dados_pcmso, banco, obra_id=None):
    divergencias = []
    resumo = defaultdict(int)
    obras = banco.get('obras_referencia', {})
    if obra_id:
        obras = {obra_id: obras.get(obra_id, {})}

    for obra_key, ghes in obras.items():
        for ghe_id, ghe_ref in ghes.items():
            saida_ghe = dados_pcmso.get(ghe_id) or dados_pcmso.get(ghe_id.replace(' ', '')) or {}
            cargos_saida = {normalizar_cargo(k): v for k, v in (saida_ghe or {}).items()}

            for cargo_ref, exames_ref in ghe_ref.get('cargos', {}).items():
                cargo_n = normalizar_cargo(cargo_ref)
                payload_saida = cargos_saida.get(cargo_n)

                if payload_saida is None:
                    divergencias.append({'obra': obra_key, 'ghe': ghe_id, 'cargo': cargo_n,
                                         'tipo': 'cargo_faltando', 'detalhe': 'cargo ausente'})
                    resumo['cargo_faltando'] += 1
                    continue

                mapa_saida = {normalizar_exame(x['nome']): x for x in _lista_de_exames(payload_saida)}
                mapa_ref   = {normalizar_exame(e['nome']): e for e in exames_ref}

                for nome in sorted(set(mapa_ref) - set(mapa_saida)):
                    divergencias.append({'obra': obra_key, 'ghe': ghe_id, 'cargo': cargo_n,
                                         'tipo': 'exame_faltando', 'detalhe': nome})
                    resumo['exame_faltando'] += 1

                for nome in sorted(set(mapa_saida) - set(mapa_ref)):
                    divergencias.append({'obra': obra_key, 'ghe': ghe_id, 'cargo': cargo_n,
                                         'tipo': 'exame_excedente', 'detalhe': nome})
                    resumo['exame_excedente'] += 1

                for nome in sorted(set(mapa_ref) & set(mapa_saida)):
                    r, s = mapa_ref[nome], mapa_saida[nome]
                    for flag in ('adm', 'mro', 'ret', 'dem'):
                        if r.get(flag) != s.get(flag):
                            divergencias.append({'obra': obra_key, 'ghe': ghe_id, 'cargo': cargo_n,
                                                 'tipo': 'flag_incorreta',
                                                 'detalhe': f'{nome}::{flag} esperado={r.get(flag)} atual={s.get(flag)}'})
                            resumo['flag_incorreta'] += 1
                    if _norm_per(r.get('per')) != _norm_per(s.get('per')):
                        divergencias.append({'obra': obra_key, 'ghe': ghe_id, 'cargo': cargo_n,
                                             'tipo': 'periodicidade_incorreta',
                                             'detalhe': f'{nome}::per esperado={_norm_per(r.get("per"))} atual={_norm_per(s.get("per"))}'})
                        resumo['periodicidade_incorreta'] += 1

    return {
        'ok': not divergencias,
        'total_divergencias': len(divergencias),
        'resumo': dict(resumo),
        'divergencias': divergencias,
    }


def formatar_relatorio_auditoria(resultado):
    linhas = [f"Total de divergencias: {resultado.get('total_divergencias', 0)}"]
    for k, v in sorted(resultado.get('resumo', {}).items()):
        linhas.append(f'  {k}: {v}')
    if resultado.get('divergencias'):
        linhas.append('')
        linhas.append('Detalhes:')
        for d in resultado['divergencias']:
            linhas.append(f"  [{d['obra']}] {d['ghe']} | {d['cargo']} | {d['tipo']} | {d['detalhe']}")
    return '\n'.join(linhas)

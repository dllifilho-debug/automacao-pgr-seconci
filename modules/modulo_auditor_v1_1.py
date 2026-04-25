import json
import re
import unicodedata
from pathlib import Path
from collections import defaultdict


def _strip(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto or ''))
                   if unicodedata.category(c) != 'Mn')


def norm(texto):
    # remove acentos, hifens, pontos e multiplos espacos
    s = _strip(texto).upper().strip()
    s = re.sub(r'[\-\.\'\`]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def normalizar_exame(nome):
    s = norm(nome)
    mapa = {
        # Clinico
        'EXAME CLINICO': 'Exame Clinico',
        'EXAME CLINICO SEMESTRAL': 'Exame Clinico',
        # Audiometria
        'AUDIOMETRIA': 'Audiometria',
        'AUDIOMETRIA TONAL': 'Audiometria',
        'AUDIOMETRIA TONAL PTA': 'Audiometria',
        # Acuidade
        'ACUIDADE VISUAL': 'Acuidade Visual',
        'AVALIACAO OFTALMOLOGICA': 'Acuidade Visual',
        # Hemograma
        'HEMOGRAMA': 'Hemograma',
        'HEMOGRAMA COMPLETO': 'Hemograma',
        # Glicemia
        'GLICEMIA EM JEJUM': 'Glicemia em Jejum',
        'GLICEMIA DE JEJUM': 'Glicemia em Jejum',
        # ECG
        'ECG': 'ECG',
        'ELETROCARDIOGRAMA': 'ECG',
        'ELETROCARDIOGRAMA ECG': 'ECG',
        # Espirometria
        'ESPIROMETRIA': 'Espirometria',
        'ESPIROMETRIA (SOMENTE)': 'Espirometria',
        # RX Torax
        'RX DE TORAX OIT': 'RX de Tórax OIT',
        'RX TORAX OIT': 'RX de Tórax OIT',
        'RAIO X TORAX OIT': 'RX de Tórax OIT',
        'RX DE TORAX': 'RX de Tórax OIT',
        # RX Coluna
        'RX DE COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
        'RX COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
        'RAIO X COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
        # Psicossocial
        'AVALIACAO PSICOSSOCIAL': 'Avaliação Psicossocial',
        'AVALIACAO PSICOSSOCIAL NR 35': 'Avaliação Psicossocial',
        # Carboxiemoglobina — aceita ambas as grafias
        'CARBOXIEMOGLOBINA': 'Carboxiemoglobina',
        'CARBOXIHEMOGLOBINA': 'Carboxiemoglobina',
        'CARBOXIHEMOGLOBINA NO SANGUE': 'Carboxiemoglobina',
        'CARBOXIEMOGLOBINA NO SANGUE': 'Carboxiemoglobina',
        # Manganes
        'MANGANES SANGUINEO': 'Manganês sanguíneo',
        'MANGANES NO SANGUE': 'Manganês sanguíneo',
        # Reticulocitos
        'CONTAGEM DE RETICULOCITOS': 'Contagem de Reticulócitos',
        'RETICULOCITOS': 'Contagem de Reticulócitos',
        # Acido muconico — varias grafias
        'ACIDO TRANS TRANS MUCONICO': 'Ácido trans-trans mucônico',
        'ACIDO TRANS TRANS MUCONICO NA URINA': 'Ácido trans-trans mucônico',
        'AC TRANS TRANS MUCONICO NA URINA': 'Ácido trans-trans mucônico',
        'AC TRANS TRANS MUCONICO': 'Ácido trans-trans mucônico',
        # Acido tricloroacetico
        'ACIDO TRICLOROACETICO NA URINA': 'Ácido tricloroacético na urina',
        'AC TRICLOROACETICO NA URINA': 'Ácido tricloroacético na urina',
        'ACIDO TRICLOROACETICO': 'Ácido tricloroacético na urina',
        # Acetona
        'ACETONA NA URINA': 'Acetona na urina',
        # Ortocresol
        'ORTOCRESOL NA URINA': 'Ortocresol na urina',
        # MEK — todas as variantes com e sem hifen
        'METIL ETIL CETONA': 'Metil-Etil-Cetona',
        'METIL ETIL CETONA NA URINA': 'Metil-Etil-Cetona',
        'METILETILCETONA NA URINA': 'Metil-Etil-Cetona',
        'MEK NA URINA': 'Metil-Etil-Cetona',
        'METIL ETIL CETONA (MEK) NA URINA': 'Metil-Etil-Cetona',
        # Acido metil hipurico
        'ACIDO METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
        'AC METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
        'AC  METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
        # Ciclohexanol
        'CICLOHEXANOL NA URINA': 'Ciclohexanol na urina',
        'CICLOHEXANOL H NA URINA': 'Ciclohexanol na urina',
        # Tetrahidrofurano
        'TETRAHIDROFURNANO NA URINA': 'Tetrahidrofurnano na urina',
        'TETRAHIDROFURANO NA URINA': 'Tetrahidrofurnano na urina',
        # EPF
        'EPF (COPROPARASITOLOGICO) + ANTI HBS': 'EPF (Coproparasitológico) + Anti-HBs',
        'EPF COPROPARASITOLOGICO + ANTI HBS': 'EPF (Coproparasitológico) + Anti-HBs',
    }
    return mapa.get(s, nome.strip())


def normalizar_cargo(nome):
    """Normaliza nome de cargo para comparacao — NAO colapsa cargos distintos."""
    s = norm(nome)
    aliases = {
        # variantes de escrita para o mesmo cargo canônico
        'MESTRE DE OBRA': 'Mestre de Obra',
        'MESTRE DE OBRAS': 'Mestre de Obra',
        'OPERADOR DE BETONEIRA': 'Operador de Betoneira',
        'OPERADOR DE GRUA': 'Operador de Grua',
        'OPERADOR DE CREMALHEIRA': 'Operador de Cremalheira',
        'MEIO OFICIAL DE PEDREIRO': 'Meio Oficial de Pedreiro',
        'MEIO OFICIAL DE ARMADOR': 'Meio Oficial de Armador',
        'MEIO OFICIAL DE CARPINTEIRO': 'Meio Oficial de Carpinteiro',
        'MEIO OFICIAL DE ELETRICISTA': 'Meio Oficial de Eletricista',
        'MEIO OFICIAL DE ENCANADOR': 'Meio Oficial de Encanador',
        'MEIO OFICIAL DE SERRALHEIRO': 'Meio Oficial de Serralheiro',
        'MEIO OFICIAL DE PINTOR': 'Meio Oficial de Pintor',
        'SERVENTE DE ARMADOR': 'Servente de Armador',
        'SERVENTE DE CARPINTEIRO': 'Servente de Carpinteiro',
        'SERVENTE DE OBRA': 'Servente',
        'ELETRICISTA INDUSTRIAL': 'Eletricista Industrial',
        'ENCARREGADO DE IMPERMEABILIZACAO': 'Encarregado de Impermeabilização',
        'ENCARREGADO DE PEDREIRO': 'Encarregado de Pedreiro',
        'ENCARREGADO DE PINTOR': 'Encarregado de Pintor',
        'ENCARREGADO DE ELETRICISTA': 'Encarregado de Eletricista',
        'ENCARREGADO DE ENCANADOR': 'Encarregado de Encanador',
        'ENCARREGADO DE REJUNTE': 'Encarregado de Rejunte',
        'ESTAGIARIO DE ENGENHARIA': 'Estagiário de Engenharia',
        'ESTAGIARIO DE ENGENHARIA CIVIL': 'Estagiário de Engenharia',
        'ESTAGIARIO DE SEGURANCA DO TRABALHO': 'Estagiário de Segurança do Trabalho',
        'TECNICO DE SEGURANCA DO TRABALHO': 'Técnico de Segurança do Trabalho',
        'ADMINISTRATIVO DE OBRAS': 'Administrativo de Obras',
        'AUXILIAR ADMINISTRATIVO DE OBRAS': 'Auxiliar Administrativo de Obras',
        'JOVEM APRENDIZ': 'Jovem Aprendiz',
        'MECANICO DE MANUTENCAO': 'Mecânico de Manutenção',
        'SERRALHEIRO': 'Serralheiro',
        'PINTOR': 'Pintor',
        'SERVENTE': 'Servente',
        'ARMADOR': 'Armador',
        'CARPINTEIRO': 'Carpinteiro',
        'PEDREIRO': 'Pedreiro',
        'GESSEIRO': 'Gesseiro',
        'ENCANADOR': 'Encanador',
        'MONTADOR': 'Montador',
        'SOLDADOR': 'Soldador',
        'ELETRICISTA': 'Eletricista',
        'SINALEIRO': 'Sinaleiro',
        'ALMOXARIFE': 'Almoxarife',
        'ENGENHEIRO': 'Engenheiro',
        'ENGENHEIRO CIVIL': 'Engenheiro',
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
    """Extrai numero inteiro de periodicidade. Retorna None se ausente ou nao numerico."""
    if v is None:
        return None
    s = str(v).strip().upper()
    if s in ('', 'NONE', 'FALSE', '-', 'NULL'):
        return None
    m = re.search(r'(\d+)', s)
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


def pcmso_df_para_dict(df) -> dict:
    """
    Converte o DataFrame retornado por processar_pcmso() no formato
    esperado por auditar_pcmso().
    """
    resultado = {}
    for _, row in df.iterrows():
        ghe_raw = str(row.get('GHE / Setor', '')).strip()
        cargo   = str(row.get('Cargo', '')).strip()
        exame   = str(row.get('Exame', '')).strip()

        if not ghe_raw or not cargo or not exame:
            continue

        m = re.search(r'GHE\s*(\d{1,2})', ghe_raw, re.IGNORECASE)
        ghe_key = f"GHE {int(m.group(1)):02d}" if m else ghe_raw

        resultado.setdefault(ghe_key, {})
        resultado[ghe_key].setdefault(cargo, [])
        resultado[ghe_key][cargo].append({
            'nome': exame,
            'adm':  row.get('ADM', '-'),
            'per':  row.get('PER', '-'),
            'mro':  row.get('MRO', '-'),
            'ret':  row.get('RT',  '-'),
            'dem':  row.get('DEM', '-'),
        })

    return resultado


def auditar_pcmso(dados_pcmso, banco, obra_id=None):
    divergencias = []
    resumo = defaultdict(int)
    obras = banco.get('obras_referencia', {})
    if obra_id:
        obras = {obra_id: obras.get(obra_id, {})}

    for obra_key, ghes in obras.items():
        for ghe_id, ghe_ref in ghes.items():
            saida_ghe = dados_pcmso.get(ghe_id) or dados_pcmso.get(ghe_id.replace(' ', '')) or {}
            # normaliza chaves de cargo na saída gerada
            cargos_saida = {normalizar_cargo(k): v for k, v in (saida_ghe or {}).items()}

            for cargo_ref, exames_ref in ghe_ref.get('cargos', {}).items():
                cargo_n = normalizar_cargo(cargo_ref)
                payload_saida = cargos_saida.get(cargo_n)

                if payload_saida is None:
                    divergencias.append({'obra': obra_key, 'ghe': ghe_id, 'cargo': cargo_n,
                                         'tipo': 'cargo_faltando', 'detalhe': 'cargo ausente'})
                    resumo['cargo_faltando'] += 1
                    continue

                mapa_saida = {normalizar_exame(x['nome']): x
                              for x in _lista_de_exames(payload_saida)}
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
                        rv = _norm_bool(r.get(flag))
                        sv = _norm_bool(s.get(flag))
                        if rv != sv:
                            divergencias.append({
                                'obra': obra_key, 'ghe': ghe_id, 'cargo': cargo_n,
                                'tipo': 'flag_incorreta',
                                'detalhe': f'{nome}::{flag} esperado={rv} atual={sv}'
                            })
                            resumo['flag_incorreta'] += 1
                    rp = _norm_per(r.get('per'))
                    sp = _norm_per(s.get('per'))
                    if rp != sp:
                        divergencias.append({
                            'obra': obra_key, 'ghe': ghe_id, 'cargo': cargo_n,
                            'tipo': 'periodicidade_incorreta',
                            'detalhe': f'{nome}::per esperado={rp} atual={sp}'
                        })
                        resumo['periodicidade_incorreta'] += 1

    return {
        'ok': not divergencias,
        'total_divergencias': len(divergencias),
        'resumo': dict(resumo),
        'divergencias': divergencias,
    }


def formatar_relatorio_auditoria(resultado):
    linhas = [f"Auditoria concluída — {resultado.get('total_divergencias', 0)} divergência(s) detectada(s)."]
    linhas.append('')
    linhas.append(f"Total de divergencias: {resultado.get('total_divergencias', 0)}")
    for k, v in sorted(resultado.get('resumo', {}).items()):
        linhas.append(f'  {k}: {v}')
    if resultado.get('divergencias'):
        linhas.append('')
        linhas.append('Detalhes:')
        for d in resultado['divergencias']:
            linhas.append(f"  [{d['obra']}] {d['ghe']} | {d['cargo']} | {d['tipo']} | {d['detalhe']}")
    return '\n'.join(linhas)

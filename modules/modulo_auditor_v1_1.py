import json
import re
import unicodedata
from pathlib import Path
from collections import defaultdict

# ── rapidfuzz opcional (fuzzy matching) ────────────────────────────────────
try:
    from rapidfuzz import process as _rfprocess, fuzz as _rffuzz
    _FUZZY_OK = True
except ImportError:
    _FUZZY_OK = False

# ── Mapa CBO → cargo canônico (Camada 2) ──────────────────────────────────
_CBO_PATH = Path(__file__).parent.parent / 'data' / 'mapa_cbo_cargo.json'
try:
    with _CBO_PATH.open('r', encoding='utf-8') as _f:
        MAPA_CBO_CARGO: dict = json.load(_f)
except Exception:
    MAPA_CBO_CARGO = {}

# ── Banco de aprendizado (Camada 3) ───────────────────────────────────────
_APRENDIZADO_PATH = Path(__file__).parent.parent / 'data' / 'banco_aprendizado.json'


def _carregar_aprendizado() -> dict:
    try:
        if _APRENDIZADO_PATH.exists():
            with _APRENDIZADO_PATH.open('r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _salvar_aprendizado(dados: dict):
    try:
        _APRENDIZADO_PATH.parent.mkdir(parents=True, exist_ok=True)
        with _APRENDIZADO_PATH.open('w', encoding='utf-8') as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def registrar_aprendizado(nome_cargo: str, exames: list):
    """
    Camada 3 — chamado na aprovação do PCMSO.
    Salva o cargo + exames aprovados no banco_aprendizado.json para uso futuro.
    Não sobrescreve se já existir entrada mais completa (mais exames = mais seguro).
    """
    if not nome_cargo or not exames:
        return
    dados = _carregar_aprendizado()
    chave = norm(nome_cargo)
    existente = dados.get(chave, [])
    if len(exames) > len(existente):
        dados[chave] = exames
        _salvar_aprendizado(dados)


def _strip(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto or ''))
                   if unicodedata.category(c) != 'Mn')


def norm(texto):
    s = _strip(texto).upper().strip()
    s = re.sub(r'[\-\.\'\`]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def normalizar_exame(nome):
    s = norm(nome)
    mapa = {
        'EXAME CLINICO': 'Exame Clinico',
        'EXAME CLINICO SEMESTRAL': 'Exame Clinico',
        'AUDIOMETRIA': 'Audiometria',
        'AUDIOMETRIA TONAL': 'Audiometria',
        'AUDIOMETRIA TONAL PTA': 'Audiometria',
        'ACUIDADE VISUAL': 'Acuidade Visual',
        'AVALIACAO OFTALMOLOGICA': 'Acuidade Visual',
        'HEMOGRAMA': 'Hemograma',
        'HEMOGRAMA COMPLETO': 'Hemograma',
        'GLICEMIA EM JEJUM': 'Glicemia em Jejum',
        'GLICEMIA DE JEJUM': 'Glicemia em Jejum',
        'ECG': 'ECG',
        'ELETROCARDIOGRAMA': 'ECG',
        'ELETROCARDIOGRAMA ECG': 'ECG',
        'ESPIROMETRIA': 'Espirometria',
        'ESPIROMETRIA (SOMENTE)': 'Espirometria',
        'RX DE TORAX OIT': 'RX de Tórax OIT',
        'RX TORAX OIT': 'RX de Tórax OIT',
        'RAIO X TORAX OIT': 'RX de Tórax OIT',
        'RX DE TORAX': 'RX de Tórax OIT',
        'RX DE COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
        'RX COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
        'RAIO X COLUNA LOMBO SACRA': 'RX de coluna lombo-sacra',
        'AVALIACAO PSICOSSOCIAL': 'Avaliação Psicossocial',
        'AVALIACAO PSICOSSOCIAL NR 35': 'Avaliação Psicossocial',
        'CARBOXIEMOGLOBINA': 'Carboxiemoglobina',
        'CARBOXIHEMOGLOBINA': 'Carboxiemoglobina',
        'CARBOXIHEMOGLOBINA NO SANGUE': 'Carboxiemoglobina',
        'CARBOXIEMOGLOBINA NO SANGUE': 'Carboxiemoglobina',
        'MANGANES SANGUINEO': 'Manganês sanguíneo',
        'MANGANES NO SANGUE': 'Manganês sanguíneo',
        'CONTAGEM DE RETICULOCITOS': 'Contagem de Reticulócitos',
        'RETICULOCITOS': 'Contagem de Reticulócitos',
        'ACIDO TRANS TRANS MUCONICO': 'Ácido trans-trans mucônico',
        'ACIDO TRANS TRANS MUCONICO NA URINA': 'Ácido trans-trans mucônico',
        'AC TRANS TRANS MUCONICO NA URINA': 'Ácido trans-trans mucônico',
        'AC TRANS TRANS MUCONICO': 'Ácido trans-trans mucônico',
        'ACIDO TRICLOROACETICO NA URINA': 'Ácido tricloroacético na urina',
        'AC TRICLOROACETICO NA URINA': 'Ácido tricloroacético na urina',
        'ACIDO TRICLOROACETICO': 'Ácido tricloroacético na urina',
        'ACETONA NA URINA': 'Acetona na urina',
        'ORTOCRESOL NA URINA': 'Ortocresol na urina',
        'METIL ETIL CETONA': 'Metil-Etil-Cetona',
        'METIL ETIL CETONA NA URINA': 'Metil-Etil-Cetona',
        'METILETILCETONA NA URINA': 'Metil-Etil-Cetona',
        'MEK NA URINA': 'Metil-Etil-Cetona',
        'METIL ETIL CETONA (MEK) NA URINA': 'Metil-Etil-Cetona',
        'ACIDO METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
        'AC METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
        'AC  METIL HIPURICO NA URINA': 'Ác. Metil-hipúrico na urina',
        'CICLOHEXANOL NA URINA': 'Ciclohexanol na urina',
        'CICLOHEXANOL H NA URINA': 'Ciclohexanol na urina',
        'TETRAHIDROFURNANO NA URINA': 'Tetrahidrofurnano na urina',
        'TETRAHIDROFURANO NA URINA': 'Tetrahidrofurnano na urina',
        'EPF (COPROPARASITOLOGICO) + ANTI HBS': 'EPF (Coproparasitológico) + Anti-HBs',
        'EPF COPROPARASITOLOGICO + ANTI HBS': 'EPF (Coproparasitológico) + Anti-HBs',
    }
    return mapa.get(s, nome.strip())


def normalizar_cargo(nome):
    """Normaliza nome de cargo para comparacao."""
    s = norm(nome)
    aliases = {
        # ── canteiro ──────────────────────────────────────────────────────
        'MESTRE DE OBRA': 'Mestre de Obra',
        'MESTRE DE OBRAS': 'Mestre de Obra',
        'OPERADOR DE BETONEIRA': 'Operador de Betoneira',
        'OPERADOR DE GRUA': 'Operador de Grua',
        'OPERADOR DE CREMALHEIRA': 'Operador de Cremalheira',
        'OPERADOR DE MUNCK': 'Operador de Grua',
        'MEIO OFICIAL DE PEDREIRO': 'Meio Oficial de Pedreiro',
        'MEIO OFICIAL DE ARMADOR': 'Meio Oficial de Armador',
        'MEIO OFICIAL DE CARPINTEIRO': 'Meio Oficial de Carpinteiro',
        'MEIO OFICIAL DE ELETRICISTA': 'Meio Oficial de Eletricista',
        'MEIO OFICIAL DE ELETRICA': 'Meio Oficial de Eletricista',
        'MEIO OFICIAL DE ENCANADOR': 'Meio Oficial de Encanador',
        'MEIO OFICIAL HIDRAULICO': 'Meio Oficial de Encanador',
        'MEIO OFICIAL DE HIDRAULICA': 'Meio Oficial de Encanador',
        'MEIO OFICIAL DE SERRALHEIRO': 'Meio Oficial de Serralheiro',
        'MEIO OFICIAL DE PINTOR': 'Meio Oficial de Pintor',
        'SERVENTE DE ARMADOR': 'Servente de Armador',
        'SERVENTE DE CARPINTEIRO': 'Servente de Carpinteiro',
        'SERVENTE DE OBRA': 'Servente',
        'SERVENTE DE OBRAS': 'Servente',
        'ELETRICISTA INDUSTRIAL': 'Eletricista Industrial',
        'ENCARREGADO DE IMPERMEABILIZACAO': 'Encarregado de Impermeabilização',
        'ENCARREGADO DE PEDREIRO': 'Encarregado de Pedreiro',
        'ENCARREGADO DE PINTOR': 'Encarregado de Pintor',
        'ENCARREGADO DE ELETRICISTA': 'Encarregado de Eletricista',
        'ENCARREGADO DE ENCANADOR': 'Encarregado de Encanador',
        'ENCARREGADO DE REJUNTE': 'Encarregado de Rejunte',
        'ENCARREGADO DE OBRAS': 'Encarregado de Obras',
        'ENCARREGADO DE OBRA': 'Encarregado de Obras',
        'ENCARREGADO DE ACABAMENTO': 'Encarregado de Pedreiro',
        'ENCARREGADO DE ARMACAO': 'Encarregado de Pedreiro',
        'ENCARREGADO DE FORMA': 'Encarregado de Carpinteiro',
        'ENCARREGADO DE INSTALACOES': 'Encarregado de Eletricista',
        'ENCARREGADO DE INSTALAÇÕES': 'Encarregado de Eletricista',
        'ENCARREGADO ADMINISTRATIVO DE OBRAS': 'Auxiliar Administrativo de Obras',
        # ── estagiários ───────────────────────────────────────────────────
        'ESTAGIARIO DE ENGENHARIA': 'Estagiário de Engenharia',
        'ESTAGIARIO DE ENGENHARIA CIVIL': 'Estagiário de Engenharia',
        'ESTAGIARIO DE SEGURANCA DO TRABALHO': 'Estagiário de Segurança do Trabalho',
        'ESTAGIARIO': 'Estagiário de Engenharia',
        'ESTAGIÁRIO': 'Estagiário de Engenharia',
        # ── técnicos ──────────────────────────────────────────────────────
        'TECNICO DE SEGURANCA DO TRABALHO': 'Técnico de Segurança do Trabalho',
        'TECNICO DE SEGURANÇA DO TRABALHO': 'Técnico de Segurança do Trabalho',
        'TÉCNICO DE SEGURANÇA DO TRABALHO': 'Técnico de Segurança do Trabalho',
        'TECNICO EM EDIFICACOES': 'Técnico em Edificações',
        'TECNICO EM EDIFICAÇÕES': 'Técnico em Edificações',
        'TÉCNICO EM EDIFICAÇÕES': 'Técnico em Edificações',
        # ── administrativo ────────────────────────────────────────────────
        'ADMINISTRATIVO DE OBRAS': 'Administrativo de Obras',
        'AUXILIAR ADMINISTRATIVO DE OBRAS': 'Auxiliar Administrativo de Obras',
        'ASSISTENTE ADMINISTRATIVO': 'Assistente Administrativo',
        'AUXILIAR ADMINISTRATIVO': 'Auxiliar Administrativo de Obras',
        'JOVEM APRENDIZ': 'Jovem Aprendiz',
        'ALMOXARIFE': 'Almoxarife',
        'AUXILIAR DE ALMOXARIFE': 'Almoxarife',
        'AUXILIAR DE ALMOXARIFADO': 'Almoxarife',
        # ── serviços gerais ───────────────────────────────────────────────
        'AUXILIAR DE LIMPEZA': 'Servente',
        'AUXILIAR DE SERVICOS GERAIS': 'Servente',
        'AUXILIAR DE SERVIÇOS GERAIS': 'Servente',
        'COPEIRA': 'Servente',
        'VIGIA': 'Vigia',
        'VIGIA DIURNO': 'Vigia',
        'VIGIA NOTURNO': 'Vigia',
        # ── engenharia ────────────────────────────────────────────────────
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
        'ENGENHEIRO': 'Engenheiro',
        'ENGENHEIRO CIVIL': 'Engenheiro',
        'AUXILIAR DE ENGENHARIA': 'Estagiário de Engenharia',
    }
    if s in aliases:
        return aliases[s]
    return nome.strip().title()


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


# ───────────────────────────────────────────────────────────────────────────
# Camada 2 — resolve cargo pelo CBO quando nome nao bate
# ───────────────────────────────────────────────────────────────────────────

def _resolver_por_cbo(cbo: str) -> str | None:
    """
    Recebe CBO string (ex: '715210') e retorna nome canônico do cargo
    conforme mapa_cbo_cargo.json, ou None se não encontrar.
    """
    if not cbo:
        return None
    cbo_limpo = re.sub(r'[^\d]', '', str(cbo))
    return MAPA_CBO_CARGO.get(cbo_limpo)


# ───────────────────────────────────────────────────────────────────────────
# Camada 1b — fuzzy matching contra cargos do banco
# ───────────────────────────────────────────────────────────────────────────

def _buscar_fuzzy(nome_cargo: str, banco: dict, threshold: int = 82) -> list | None:
    """
    Tenta encontrar exames via similaridade de string (rapidfuzz).
    Retorna lista de exames do melhor match ou None.
    """
    if not _FUZZY_OK:
        return None

    # Coleta todos os nomes de cargo do banco
    candidatos = {}
    for obra in banco.get('obras_referencia', {}).values():
        for ghe in obra.values():
            for cargo_ref, exames_ref in ghe.get('cargos', {}).items():
                chave = normalizar_cargo(cargo_ref)
                if chave not in candidatos or len(exames_ref) > len(candidatos[chave]):
                    candidatos[chave] = exames_ref

    if not candidatos:
        return None

    cargo_n = normalizar_cargo(nome_cargo)
    resultado = _rfprocess.extractOne(
        cargo_n,
        list(candidatos.keys()),
        scorer=_rffuzz.token_sort_ratio,
    )
    if resultado and resultado[1] >= threshold:
        return list(candidatos[resultado[0]])
    return None


# ───────────────────────────────────────────────────────────────────────────
# REFERENCIA TECNICA POR CARGO — independente do nome do arquivo
# ───────────────────────────────────────────────────────────────────────────

def buscar_exames_por_cargo(
    nome_cargo: str,
    banco: dict,
    cbo: str = '',
) -> tuple[list | None, str]:
    """
    Busca exames em 4 camadas (retorna o primeiro match + origem):
      1. Alias exato via normalizar_cargo()  → origem 'alias'
      2. Aprendizado anterior (banco_aprendizado.json) → origem 'aprendizado'
      3. Fallback CBO via mapa_cbo_cargo.json          → origem 'cbo'
      4. Fuzzy matching via rapidfuzz                  → origem 'fuzzy'

    Retorna (lista_exames | None, origem_str)
    """
    # 1 — Alias + busca exata no banco
    cargo_norm = normalizar_cargo(nome_cargo)
    melhor = None
    for obra in banco.get('obras_referencia', {}).values():
        for ghe in obra.values():
            for cargo_ref, exames_ref in ghe.get('cargos', {}).items():
                if normalizar_cargo(cargo_ref) == cargo_norm:
                    candidato = list(exames_ref)
                    if melhor is None or len(candidato) > len(melhor):
                        melhor = candidato
    if melhor:
        return melhor, 'alias'

    # 2 — Banco de aprendizado
    aprendizado = _carregar_aprendizado()
    chave_aprendizado = norm(nome_cargo)
    if chave_aprendizado in aprendizado:
        return aprendizado[chave_aprendizado], 'aprendizado'

    # 3 — Fallback CBO
    if cbo:
        nome_por_cbo = _resolver_por_cbo(cbo)
        if nome_por_cbo:
            cargo_cbo_norm = normalizar_cargo(nome_por_cbo)
            for obra in banco.get('obras_referencia', {}).values():
                for ghe in obra.values():
                    for cargo_ref, exames_ref in ghe.get('cargos', {}).items():
                        if normalizar_cargo(cargo_ref) == cargo_cbo_norm:
                            candidato = list(exames_ref)
                            if melhor is None or len(candidato) > len(melhor):
                                melhor = candidato
            if melhor:
                return melhor, f'cbo:{cbo}'

    # 4 — Fuzzy
    fuzzy = _buscar_fuzzy(nome_cargo, banco)
    if fuzzy:
        return fuzzy, 'fuzzy'

    return None, 'nenhum'


def enriquecer_ghe_com_banco(dados_ghe: list, banco: dict) -> tuple:
    """
    Para cada cargo, tenta resolver em 4 camadas (alias → aprendizado → CBO → fuzzy).
    O campo 'cbo' pode vir no dict do GHE como 'cbo' ou embutido no nome do cargo
    via regex 'CBO[:\s]*(\\d+)'.

    Retorna (dados_ghe, relatorio) com relatorio contendo:
        'cargos_enriquecidos': [(nome, origem), ...]
        'cargos_mantidos':     [nome, ...]
        'mapa_exames_banco':   {cargo_norm: [exames]}
    """
    cargos_enriquecidos = []   # lista de (nome, origem)
    cargos_mantidos = []
    mapa_exames_banco = {}
    processados = set()

    for ghe in dados_ghe:
        # CBO pode vir como campo explícito no dict do GHE
        cbo_ghe = str(ghe.get('cbo', ''))

        for cargo in ghe.get('cargos', []):
            nome_cargo = str(cargo)
            cargo_norm = normalizar_cargo(nome_cargo)

            if cargo_norm in processados:
                continue
            processados.add(cargo_norm)

            # Tenta extrair CBO do nome do cargo se não veio como campo
            cbo = cbo_ghe
            if not cbo:
                m_cbo = re.search(r'CBO[:\s]*(\d+)', nome_cargo, re.IGNORECASE)
                cbo = m_cbo.group(1) if m_cbo else ''

            exames, origem = buscar_exames_por_cargo(nome_cargo, banco, cbo=cbo)

            if exames:
                mapa_exames_banco[cargo_norm] = exames
                cargos_enriquecidos.append((nome_cargo, origem))
            else:
                cargos_mantidos.append(nome_cargo)

    relatorio = {
        'cargos_enriquecidos': cargos_enriquecidos,
        'cargos_mantidos': cargos_mantidos,
        'mapa_exames_banco': mapa_exames_banco,
    }
    return dados_ghe, relatorio


# ───────────────────────────────────────────────────────────────────────────
# Funcoes legadas mantidas para compatibilidade
# ───────────────────────────────────────────────────────────────────────────

def pcmso_df_para_dict(df) -> dict:
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


def obra_tem_matriz(banco: dict, obra_id: str) -> bool:
    obras = banco.get("obras_referencia", {})
    return obra_id in obras and bool(obras[obra_id])

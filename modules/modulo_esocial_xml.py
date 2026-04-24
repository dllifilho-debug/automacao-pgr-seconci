"""
modulo_esocial_xml.py — Exportação XML eSocial
Passo 4: gera o leiaute S-2240 (Condições Ambientais do Trabalho)
a partir do DataFrame de PCMSO gerado pelo sistema.

Leiaute de referência: eSocial v2.5 / S-2240
Estrutura simplificada compatível com importação manual no portal eSocial.

Nenhuma dependência externa — usa apenas xml.etree.ElementTree da stdlib.
"""

import re
import unicodedata
from datetime import date, datetime
from xml.etree.ElementTree import (
    Element, SubElement, tostring, indent
)


# ── Mapa: nome do exame → código tabela 24 eSocial ────────────────────────
# Fonte: Tabela 24 - Exames Médicos (eSocial Manual de Orientação v2.5)
_COD_EXAME: dict[str, str] = {
    "Exame Clinico":                      "0001",
    "Audiometria":                        "0002",
    "Espirometria":                       "0003",
    "Acuidade Visual":                    "0004",
    "ECG":                                "0005",
    "Glicemia em Jejum":                  "0006",
    "Hemograma":                          "0007",
    "RX de Tórax OIT":                    "0009",
    "RX de coluna lombo-sacra":           "0010",
    "Ácido tricloroacético na urina":    "0011",
    "Acetona na urina":                   "0012",
    "Metil-Etil-Cetona":                  "0013",
    "Metiletilcetona na urina":           "0013",
    "Ciclohexanol na urina":              "0014",
    "Tetrahidrofurnano na urina":         "0015",
    "Manganês sanguíneo":               "0016",
    "Carboxiemoglobina":                  "0017",
    "Contagem de Reticulócitos":          "0018",
    "Ácido trans-trans mucônico":        "0019",
    "Ortocresol na urina":                "0020",
    "Ác. Metil-hipúrico na urina":       "0021",
    "2,5 Hexanodiona na Urina":           "0022",
    "Anti-HBs + HBsAg + Anti-HCV":       "0023",
    "Avaliação Psicossocial":            "0099",
}

# Mapa: período (meses) → código periodicidade eSocial
_COD_PERIODICIDADE: dict[str, str] = {
    "6M":  "6",
    "12M": "12",
    "24M": "24",
    "36M": "36",
    "60M": "60",
}

# Mapa: nome do agente → código da tabela 23 eSocial (Agentes Nocivos)
_COD_AGENTE: dict[str, str] = {
    "RUIDO":                  "01.01.001",
    "VIBRACAO CORPO INTEIRO": "01.02.001",
    "VIBRACAO MAOS BRACOS":   "01.02.002",
    "CALOR":                  "01.03.001",
    "RADIACAO IONIZANTE":     "01.05.001",
    "BENZENO":                "02.01.001",
    "TOLUENO":                "02.01.002",
    "XILENO":                 "02.01.003",
    "CHUMBO":                 "02.02.001",
    "MERCURIO":               "02.02.002",
    "MANGANES":               "02.02.003",
    "CROMO":                  "02.02.004",
    "CADMIO":                 "02.02.005",
    "ARSENICO":               "02.02.006",
    "SILICA":                 "02.03.001",
    "ASBESTO":                "02.03.002",
    "POEIRA MINERAL":         "02.03.003",
    "BENZENO (PETRO)":        "02.01.001",
    "DIESEL":                 "02.01.010",
    "MONOXIDO DE CARBONO":    "02.01.011",
    "AGENTE BIOLOGICO":       "03.01.001",
    "ESGOTO":                 "03.01.002",
    "QUEDA DE ALTURA":        "05.01.001",
    "ESPACO CONFINADO":       "05.01.002",
    "RISCO ELETRICO":         "05.01.003",
}


# ── Helpers ───────────────────────────────────────────────────────────────
def _sem_acentos(texto: str) -> str:
    return unicodedata.normalize('NFKD', str(texto)).encode('ascii', 'ignore').decode('ascii')


def _norm(texto: str) -> str:
    texto = _sem_acentos(str(texto or '')).upper().strip()
    return re.sub(r'\s+', ' ', re.sub(r'[^A-Z0-9 ]', ' ', texto)).strip()


def _cod_exame(nome_exame: str) -> str:
    """Retorna o código da tabela 24 ou '9999' (Outro) se não mapeado."""
    return _COD_EXAME.get(nome_exame.strip(), "9999")


def _cod_agente(nome_agente: str) -> str | None:
    """Retorna o código da tabela 23 ou None se não mapeado."""
    chave = _norm(nome_agente)
    return _COD_AGENTE.get(chave)


def _per_codigo(per_str: str) -> str:
    """Converte '12M' → '12', '6M' → '6', etc."""
    per = str(per_str).strip().upper()
    return _COD_PERIODICIDADE.get(per, per.replace('M', ''))


def _flag_xml(val) -> str:
    """Converte 'X' / True → 'S', '-' / False → 'N'"""
    if isinstance(val, bool):
        return 'S' if val else 'N'
    return 'S' if str(val).strip().upper() in ('X', 'S', 'SIM', 'TRUE', '1') else 'N'


def _cnpj_limpo(cnpj: str) -> str:
    return re.sub(r'\D', '', str(cnpj or ''))[:14].zfill(14)


# ── Gerador principal ──────────────────────────────────────────────────────────
def gerar_xml_s2240(df, cabecalho: dict, dados_pgr: list | None = None) -> bytes:
    """
    Gera o XML do evento S-2240 (Condições Ambientais do Trabalho).

    Parâmetros:
        df          — DataFrame retornado por processar_pcmso()
        cabecalho   — dict com razao_social, cnpj, medico_rt, vig_ini, obra etc.
        dados_pgr   — lista de GHEs do PGR (opcional, para preencher agentes nocivos)

    Retorna:
        bytes com o XML codificado em UTF-8.
    """
    cnpj = _cnpj_limpo(cabecalho.get('cnpj', ''))
    razao = cabecalho.get('razao_social', 'Empresa')
    dt_vig = cabecalho.get('vig_ini', date.today().strftime('%d/%m/%Y'))
    # Converte dd/mm/aaaa → aaaa-mm-dd
    try:
        dt_inicio = datetime.strptime(dt_vig, '%d/%m/%Y').strftime('%Y-%m-%d')
    except ValueError:
        dt_inicio = date.today().isoformat()

    # ── Raiz do documento ─────────────────────────────────────────────────
    root = Element('eSocial', xmlns='http://www.esocial.gov.br/schema/evt/evtCAT/v_S_01_02_00')

    # evtCondAmb
    evt = SubElement(root, 'evtCondAmb', Id=f'ID1{cnpj}{datetime.now().strftime("%Y%m%d%H%M%S")}0001')

    # ideEvento
    ide_evt = SubElement(evt, 'ideEvento')
    SubElement(ide_evt, 'indRetif').text = '1'          # 1 = original
    SubElement(ide_evt, 'nrRec').text = ''
    SubElement(ide_evt, 'tpAmb').text = '2'             # 2 = produção restrita
    SubElement(ide_evt, 'procEmi').text = '1'           # 1 = aplicação do empregador
    SubElement(ide_evt, 'verProc').text = '1.0'

    # ideEmpregador
    ide_emp = SubElement(evt, 'ideEmpregador')
    SubElement(ide_emp, 'tpInsc').text = '1'            # 1 = CNPJ
    SubElement(ide_emp, 'nrInsc').text = cnpj

    # infoCondAmb
    info = SubElement(evt, 'infoCondAmb')
    SubElement(info, 'dtVigencia').text = dt_inicio
    SubElement(info, 'cnpjRespReg').text = cnpj

    # ── Agrupa DataFrame por GHE e Cargo ─────────────────────────────────────
    col_ghe   = _achar_coluna(df, ('GHE / Setor', 'GHE', 'GHE_ID'))
    col_cargo = _achar_coluna(df, ('Cargo', 'CARGO', 'Funcão'))
    col_exame = _achar_coluna(df, ('Exame', 'EXAME'))
    col_adm   = _achar_coluna(df, ('ADM',))
    col_per   = _achar_coluna(df, ('PER', 'Periodicidade'))
    col_mro   = _achar_coluna(df, ('MRO',))
    col_rt    = _achar_coluna(df, ('RT',))
    col_dem   = _achar_coluna(df, ('DEM',))

    # Índice de agentes por GHE (do dados_pgr, se disponível)
    agentes_por_ghe: dict[str, list] = {}
    if dados_pgr:
        for ghe_item in dados_pgr:
            nome_ghe = str(ghe_item.get('ghe', '')).strip()
            agentes_por_ghe[nome_ghe] = ghe_item.get('riscos_mapeados', [])

    ghe_grupos: dict[str, dict[str, list]] = {}
    for _, row in df.iterrows():
        ghe   = str(row[col_ghe]).strip()   if col_ghe   else 'GHE'
        cargo = str(row[col_cargo]).strip() if col_cargo else 'Cargo'
        ghe_grupos.setdefault(ghe, {}).setdefault(cargo, []).append(row)

    # ── Gera um <ambiente> por GHE ────────────────────────────────────────────
    for seq_ghe, (nome_ghe, cargos_dict) in enumerate(ghe_grupos.items(), start=1):

        amb = SubElement(info, 'ambiente')
        SubElement(amb, 'seqAmbiente').text = str(seq_ghe)
        SubElement(amb, 'dscSetor').text    = nome_ghe[:100]
        SubElement(amb, 'dtIniCondicao').text = dt_inicio

        # Agentes nocivos do GHE
        agentes_ghe = agentes_por_ghe.get(nome_ghe, [])
        for agente in agentes_ghe:
            nome_ag = str(agente.get('nome_agente', '')).strip()
            cod_ag  = _cod_agente(nome_ag)
            if not cod_ag:
                continue
            ag_el = SubElement(amb, 'agNocivo')
            SubElement(ag_el, 'codAgNoc').text = cod_ag
            SubElement(ag_el, 'dscAgNoc').text = nome_ag[:100]
            SubElement(ag_el, 'tpAval').text   = '2'   # 2 = qualitativo

        # Cargos / funções
        for seq_cargo, (nome_cargo, rows) in enumerate(cargos_dict.items(), start=1):
            func_el = SubElement(amb, 'funcaoAtividade')
            SubElement(func_el, 'codFuncao').text = str(seq_cargo).zfill(4)
            SubElement(func_el, 'dscFuncao').text = nome_cargo[:100]

            # Exames por cargo
            for seq_exame, row in enumerate(rows, start=1):
                nome_ex = str(row[col_exame]).strip() if col_exame else ''
                if not nome_ex:
                    continue

                cod_ex = _cod_exame(nome_ex)
                per_val = str(row[col_per]).strip() if col_per else '-'

                exame_el = SubElement(func_el, 'exameOcupacional')
                SubElement(exame_el, 'seqExame').text    = str(seq_exame)
                SubElement(exame_el, 'codExame').text    = cod_ex
                SubElement(exame_el, 'dscExame').text    = nome_ex[:100]
                SubElement(exame_el, 'indAdm').text      = _flag_xml(row[col_adm])  if col_adm  else 'N'
                SubElement(exame_el, 'indPer').text      = 'S' if per_val not in ('-', '', 'None') else 'N'
                if per_val not in ('-', '', 'None'):
                    SubElement(exame_el, 'perExame').text = _per_codigo(per_val)
                SubElement(exame_el, 'indMRO').text      = _flag_xml(row[col_mro])  if col_mro  else 'N'
                SubElement(exame_el, 'indRetTrab').text  = _flag_xml(row[col_rt])   if col_rt   else 'N'
                SubElement(exame_el, 'indDem').text      = _flag_xml(row[col_dem])  if col_dem  else 'N'

    # ── Serializa com indentação ─────────────────────────────────────────────────
    indent(root, space='  ')
    xml_bytes = b'<?xml version="1.0" encoding="UTF-8"?>\n' + tostring(root, encoding='unicode').encode('utf-8')
    return xml_bytes


def _achar_coluna(df, candidatos: tuple) -> str | None:
    """Retorna o nome da primeira coluna do df que bate com algum candidato."""
    for c in candidatos:
        if c in df.columns:
            return c
    return None


# ── Render Streamlit (widget autossuficiente) ────────────────────────────────
def render_botao_xml(df, cabecalho: dict, dados_pgr: list | None = None, key_prefix: str = ""):
    """
    Renderiza o botão de exportação XML eSocial S-2240.
    Pode ser chamado de qualquer módulo após a aprovação da matriz.

    Parâmetros:
        df          — DataFrame aprovado (editado ou não)
        cabecalho   — dict com dados da empresa
        dados_pgr   — lista de GHEs com riscos (opcional, enriquece os agentes no XML)
        key_prefix  — prefixo para evitar conflito de chaves Streamlit
    """
    import streamlit as st

    st.markdown("---")
    st.markdown("### 📄 Exportar XML eSocial — S-2240")
    st.info(
        "É gerado o evento **S-2240 (Condições Ambientais do Trabalho)** "
        "com base na matriz PCMSO aprovada. "
        "O arquivo XML pode ser importado manualmente no portal eSocial ou via API."
    )

    with st.expander("ℹ️ O que é incluído no XML S-2240", expanded=False):
        st.markdown("""
        | Campo | Origem |
        |---|---|
        | `ideEmpregador` | CNPJ do cabeçalho |
        | `dtVigencia` | Data de início da vigência |
        | `dscSetor` | Nome do GHE / Setor |
        | `agNocivo` | Agentes mapeados no PGR (se disponíveis) |
        | `dscFuncao` | Cargo / Função |
        | `codExame` | Código Tabela 24 eSocial |
        | `indAdm / indPer / indMRO / indRetTrab / indDem` | Flags da matriz PCMSO |
        | `perExame` | Periodicidade em meses |
        """)

    cnpj_ok = len(re.sub(r'\D', '', cabecalho.get('cnpj', ''))) == 14
    if not cnpj_ok:
        st.warning("⚠️ CNPJ não informado ou inválido. O XML será gerado com CNPJ zerado — preencha antes de importar no eSocial.")

    nome_arq = (cabecalho.get('razao_social') or 'PCMSO').replace(' ', '_')[:30]

    if st.button(
        "📄 Gerar e Baixar XML S-2240",
        key=f"{key_prefix}_btn_xml_esocial",
        use_container_width=True,
    ):
        try:
            xml_bytes = gerar_xml_s2240(df, cabecalho, dados_pgr=dados_pgr)
            st.download_button(
                label="⬇️ Baixar XML eSocial S-2240",
                data=xml_bytes,
                file_name=f"S2240_{nome_arq}.xml",
                mime="application/xml",
                key=f"{key_prefix}_dl_xml_esocial",
                use_container_width=True,
            )
            with st.expander("👁️ Preview do XML gerado", expanded=False):
                st.code(xml_bytes.decode('utf-8')[:4000], language='xml')
            st.success("✅ XML S-2240 gerado com sucesso!")
        except Exception as e:
            import traceback
            st.error(f"❌ Erro ao gerar XML: {type(e).__name__}: {e}")
            st.code(traceback.format_exc(), language='python')

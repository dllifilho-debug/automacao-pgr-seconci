"""
modules/modulo_pcmso.py — v4
Ajustes:
  [1] Campo tipo_ambiente: canteiro / escritorio / misto
  [2] Limpeza de nome do GHE (remove lixo textual do PDF)
  [3] DOCX com layout fiel ao adendo Dinamica Engenharia
"""

import io, re, unicodedata
from datetime import datetime

import pdfplumber
import pandas as pd

from data.matriz_exames    import MATRIZ_FUNCAO_EXAME, MATRIZ_RISCO_EXAME
from utils.cargo_utils     import normalizar_cargo, normalizar_texto, MAPA_CARGOS_CONHECIDOS, PALAVRAS_EXCLUIR_CARGO
from utils.exame_utils     import adicionar_exame_dedup
from utils.biologico_utils import tem_risco_biologico_real, CHAVES_BIOLOGICAS_MATRIZ

# ── Constantes ────────────────────────────────────────────────────────────────

_INVALIDOS_GHE = [
    "QUANTIDADE","PREVISTOS","EXPOSTOS","TOTAL DE","NUMERO DE",
    "FUNCIONARIOS","TRABALHADORES","MEDIDAS DE CONTROLE",
    "FONTE GERADORA","TRAJETORIA","DESCRICAO",
    "ATIVIDADES EXERCIDAS","INFORMACOES SOBRE",
    "PAGINA DE REVISAO","IDENTIFICACAO DA EMPRESA",
]

_CARGOS_ADMIN = {
    "ADMINISTRATIVO","RECEPCIONISTA","GESTOR","ANALISTA","ASSISTENTE",
    "COORDENADOR","DIRETOR","GERENTE","ESTAGIARIO","JOVEM APRENDIZ",
    "COMPRADOR","DESIGNER","TELEFONISTA",
}

_GHES_ADMIN = {
    "ADMINISTRACAO","PLANEJAMENTO","SUPRIMENTOS","MARKETING",
    "TI","RH","RECURSOS HUMANOS","FINANCEIRO","CONTABILIDADE",
    "ESCRITORIO","JURIDICO","COMERCIAL",
}

_PALAVRAS_CANTEIRO = [
    "OBRA","CANTEIRO","CONSTRUCAO","REFORMA","HOSPITAL","RESIDENCIAL",
    "EDIFICIO","BLOCO","TORRE","HETRIN","VIADUTO","PONTE","SHOPPING",
    "CONDOMINIO","EMPREENDIMENTO","MONTAGEM","INSTALACAO","CAMPO",
]

_PALAVRAS_ESCRITORIO = [
    "ESCRITORIO","SEDE","CORPORATIVO","ADMINISTRACAO","MARKETING",
    "TECNOLOGIA DA INFORMACAO","RECURSOS HUMANOS","FINANCEIRO",
    "CONTABILIDADE","JURIDICO","COMERCIAL",
]

_RISCOS_CANTEIRO = [
    "RUIDO","VIBRACAO","POEIRA","CIMENTO","SILICA","TINTA",
    "SOLDA","ALTURA","CONFINADO","MAQUINA","INCENDIO",
]

# Pacote padrao canteiro — baseado na matriz RICCO/HETRIN validada Dra. Patricia
_PACOTE_CANTEIRO = [
    {"exame": "Audiometria Tonal (PTA)",                   "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": True,  "obs": "Canteiro de obras"},
    {"exame": "Avaliacao Oftalmologica (Acuidade Visual)",  "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    {"exame": "Eletrocardiograma (ECG)",                    "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    {"exame": "Glicemia de Jejum",                          "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    {"exame": "Hemograma Completo",                         "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    {"exame": "Espirometria",                               "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": True,  "obs": "Exposicao a poeiras/cimento"},
    {"exame": "Raio-X de Torax OIT",                        "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": True,  "obs": "Exposicao a poeiras/cimento"},
]

_LIXO_GHE = [
    r"caracteristicas e as circunstancias",
    r"atividades exercidas",
    r"descricao das atividades",
    r"informacoes sobre",
    r"pagina de revisao",
    r"digitacao de textos",
]

_MAPA_AGENTES = {
    "TOLUENO":"Tolueno","XILENO":"Xileno","BENZENO":"Benzeno",
    "ACETONA":"Acetona","THINNER":"Solventes (Thinner)",
    "SOLVENTE":"Solventes Organicos","TINTA":"Tinta / Verniz",
    "VERNIZ":"Tinta / Verniz","PRIMER":"Primer (Solventes)",
    "GRAXA":"Graxas / Lubrificantes","DIESEL":"Diesel / Combustivel",
    "QUEROSENE":"Querosene","ACIDO":"Acidos (geral)",
    "CIMENTO":"Cimento Portland","SILICA":"Silica Cristalina",
    "POEIRA":"Poeiras Minerais","AMIANTO":"Asbesto / Amianto",
    "CHUMBO":"Chumbo","RUIDO":"Ruido",
    "VIBRACAO":"Vibracao","CALOR":"Calor (IBUTG)",
    "RADIACAO":"Radiacao Ionizante","BIOLOGICO":"Agentes Biologicos",
    "ESGOTO":"Esgoto / Aguas Servidas","SANGUE":"Material Biologico",
    "ERGONO":"Fator Ergonomico","POSTURA":"Postura Inadequada",
    "LEVANTAMENTO":"Levantamento de Carga","REPETITIVO":"Movimento Repetitivo",
    "ELETRICO":"Risco Eletrico","ALTURA":"Queda de Altura",
    "CONFINADO":"Espaco Confinado","MAQUINA":"Maquinas e Equipamentos",
    "INCENDIO":"Incendio / Explosao",
}

_RE_GHE = re.compile(
    r"(?:GHE[\s:.\-]+\w|GRUPO\s+HOMOGENEO|LOCAL\s+DE\s+TRABALHO\s*:\s*\w|SETOR\s*:\s*\w)",
    re.IGNORECASE,
)

# ── Helpers ───────────────────────────────────────────────────────────────────

def _limpar_nome_ghe(nome: str) -> str:
    if len(nome) > 100:
        return nome[:100].strip() + "..."
    norm = normalizar_texto(nome)
    for lixo in _LIXO_GHE:
        if re.search(lixo, norm):
            return "GHE (revisar nome)"
    return nome.strip()


def _is_linha_ghe(linha: str) -> bool:
    if _RE_GHE.search(linha):
        return True
    lu = normalizar_texto(linha)
    if len(linha) <= 80 and "/" not in linha:
        if any(p in lu for p in ["DEPARTAMENTO","ATIVIDADE"]):
            return True
    return False


def _ghe_valido(nome_ghe: str) -> bool:
    norm = normalizar_texto(nome_ghe)
    return not any(inv in norm for inv in _INVALIDOS_GHE)


def _fallback_necessario(ghes: list) -> bool:
    for g in ghes:
        if len(normalizar_texto(g["ghe"])) <= 60 and g["cargos"]:
            return False
    return True


def _fmt_per(per: str) -> str:
    if not per:
        return "-"
    per = per.strip().upper().replace("MESES","").strip()
    try:
        return f"{int(per)}M"
    except ValueError:
        return per


def _flag(val: bool) -> str:
    return "X" if val else "-"


def _ghe_e_canteiro_misto(nome_ghe: str, riscos: list) -> bool:
    """Detecta se GHE é canteiro no modo MISTO."""
    norm = normalizar_texto(nome_ghe)
    if any(p in norm for p in _PALAVRAS_CANTEIRO):
        return True
    if any(p in norm for p in _PALAVRAS_ESCRITORIO):
        return False
    texto_r = " ".join(
        normalizar_texto(r.get("nome_agente","") + " " + r.get("perigo_especifico",""))
        for r in riscos
    )
    return any(rc in texto_r for rc in _RISCOS_CANTEIRO)


# ── Extração ──────────────────────────────────────────────────────────────────

def extrair_texto_pdf(uploaded_file) -> str:
    texto = []
    with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                texto.append(t)
    return "\n".join(texto)


def extrair_texto_pdf_path(caminho: str) -> str:
    texto = []
    with pdfplumber.open(caminho) as pdf:
        for p in pdf.pages:
            t = p.extract_text()
            if t:
                texto.append(t)
    return "\n".join(texto)


def extrair_pgr_local(texto: str) -> list:
    linhas = texto.split("\n")
    ghes, ghe_atual, agentes_set = [], None, set()
    for linha in linhas:
        lc = linha.strip()
        if not lc:
            continue
        lu = normalizar_texto(lc)
        if _is_linha_ghe(lc) and len(lc) < 120:
            if ghe_atual and (ghe_atual["cargos"] or ghe_atual["riscos_mapeados"]):
                ghes.append(ghe_atual)
            ghe_atual   = {"ghe": lc, "cargos": [], "riscos_mapeados": []}
            agentes_set = set()
            continue
        if ghe_atual is None:
            continue
        if not any(normalizar_texto(exc) in lu for exc in PALAVRAS_EXCLUIR_CARGO):
            for cargo in MAPA_CARGOS_CONHECIDOS:
                if normalizar_texto(cargo) in lu and cargo not in ghe_atual["cargos"]:
                    ghe_atual["cargos"].append(cargo)
                    break
        for palavra, agente in _MAPA_AGENTES.items():
            if palavra in lu and agente not in agentes_set:
                agentes_set.add(agente)
                ghe_atual["riscos_mapeados"].append({
                    "nome_agente": agente,
                    "perigo_especifico": lc[:200],
                })
    if ghe_atual and (ghe_atual["cargos"] or ghe_atual["riscos_mapeados"]):
        ghes.append(ghe_atual)
    return ghes


def extrair_pgr_com_fallback(texto_pgr: str, chave_api: str = None) -> tuple:
    local = extrair_pgr_local(texto_pgr)
    if not _fallback_necessario(local):
        return local, "local"
    if chave_api:
        try:
            from utils.ia_client import extrair_pgr_via_ia
            ia = extrair_pgr_via_ia(texto_pgr, chave_api)
            return (ia, "ia") if ia else (local or [], "parcial")
        except Exception as e:
            print(f"[WARN] Falha IA: {e}")
    return (local or [], "parcial")


# ── Processamento PCMSO ───────────────────────────────────────────────────────

def processar_pcmso(dados_pgr: list, tipo_ambiente: str = "escritorio") -> pd.DataFrame:
    """
    tipo_ambiente:
      "canteiro"   — todo cargo recebe pacote completo de canteiro (RICCO/HETRIN)
      "escritorio" — cargo admin recebe so Exame Clinico + Acuidade Visual
      "misto"      — detecta por GHE usando palavras-chave e riscos
    """
    linhas = []

    for ghe in dados_pgr:
        nome_ghe_raw = ghe.get("ghe", "Sem GHE")
        nome_ghe     = _limpar_nome_ghe(nome_ghe_raw)   # Ajuste 2
        nome_norm    = normalizar_texto(nome_ghe)
        cargos       = ghe.get("cargos", [])[:15]
        riscos       = ghe.get("riscos_mapeados", [])[:10]

        if not _ghe_valido(nome_ghe):
            continue

        # Ajuste 1 — determinar contexto do GHE
        if tipo_ambiente == "canteiro":
            e_canteiro = True
        elif tipo_ambiente == "escritorio":
            e_canteiro = False
        else:  # misto
            e_canteiro = _ghe_e_canteiro_misto(nome_ghe, riscos)

        for cargo in cargos:
            exames: dict = {}

            # Exame Clinico base — sempre
            adicionar_exame_dedup(exames, {
                "exame": "Exame Clinico (Anamnese / Exame Fisico)",
                "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": True,
                "motivo": "NR-07 Basico",
            })

            cargo_norm  = normalizar_cargo(cargo)
            cargo_upper = cargo_norm.upper()

            if e_canteiro:
                # Pacote completo para todos no canteiro
                for ex in _PACOTE_CANTEIRO:
                    adicionar_exame_dedup(exames, {**ex, "motivo": "Canteiro de Obras — padrao RICCO/HETRIN"})
            else:
                # Escritorio: exames por funcao da matriz
                for funcao, lista_ex in MATRIZ_FUNCAO_EXAME.items():
                    if normalizar_texto(funcao) in cargo_upper:
                        for ex in lista_ex:
                            adicionar_exame_dedup(exames, {**ex, "motivo": f"Funcao: {funcao.title()}"})

            # Exames por risco — sempre (canteiro e escritorio)
            bio_real = tem_risco_biologico_real(riscos)
            for risco in riscos:
                texto_r = normalizar_texto(
                    risco.get("nome_agente","") + " " + risco.get("perigo_especifico","")
                )
                for chave_r, regra in MATRIZ_RISCO_EXAME.items():
                    if chave_r in CHAVES_BIOLOGICAS_MATRIZ and not bio_real:
                        continue
                    if chave_r in texto_r:
                        adicionar_exame_dedup(exames, {
                            "exame":  regra["exame"],
                            "adm":    regra.get("adm", True),
                            "per":    regra.get("periodico","12 MESES"),
                            "mro":    regra.get("mro", True),
                            "rt":     regra.get("rt", False),
                            "dem":    regra.get("dem", False),
                            "motivo": f"Exposicao: {chave_r.title()} — {regra.get('obs','')}",
                        })

            for ex_info in exames.values():
                per_fmt = _fmt_per(ex_info.get("per","12 MESES"))
                linhas.append({
                    "GHE / Setor": nome_ghe,
                    "Cargo":       cargo,
                    "Exame":       ex_info["exame"],
                    "ADM":         _flag(ex_info.get("adm", True)),
                    "PER":         per_fmt,
                    "MRO":         _flag(ex_info.get("mro", True)),
                    "RT":          _flag(ex_info.get("rt",  False)),
                    "DEM":         _flag(ex_info.get("dem", False)),
                    "Justificativa": ex_info.get("motivo",""),
                })

    return pd.DataFrame(linhas)


# ── Geração HTML ──────────────────────────────────────────────────────────────

def gerar_html_pcmso(df: pd.DataFrame, cabecalho: dict = None) -> str:
    if not cabecalho:
        cabecalho = {}
    razao  = cabecalho.get("razao_social","Empresa nao informada")
    cnpj   = cabecalho.get("cnpj","---")
    obra   = cabecalho.get("obra","---")
    medico = cabecalho.get("medico_rt","Nao informado")
    vig_i  = cabecalho.get("vig_ini","---")
    vig_f  = cabecalho.get("vig_fim","---")
    tec    = cabecalho.get("responsavel_tec","---")

    ghe_grupos = {}
    for _, row in df.iterrows():
        g = row["GHE / Setor"]
        c = row["Cargo"]
        ghe_grupos.setdefault(g, {}).setdefault(c, []).append(row)

    linhas_html = ""
    for ghe_nome, cargos_dict in ghe_grupos.items():
        total_rows   = sum(len(v) for v in cargos_dict.values())
        primeiro_ghe = True
        for cargo, rows in cargos_dict.items():
            primeiro_cargo = True
            for row in rows:
                def cel(val, bg="#d4edda"):
                    if val == "X":
                        return f'<td style="text-align:center;background:{bg};">{val}</td>'
                    return '<td style="text-align:center;color:#999;">-</td>'
                per_td = (f'<td style="text-align:center;font-weight:bold;">{row["PER"]}</td>'
                          if row["PER"] != "-" else '<td style="text-align:center;color:#999;">-</td>')
                ghe_td = ""
                if primeiro_ghe:
                    ghe_td = (f'<td rowspan="{total_rows}" style="background:#084D22;color:#fff;'
                              f'font-weight:bold;vertical-align:middle;text-align:center;padding:8px;">{ghe_nome}</td>')
                    primeiro_ghe = False
                cargo_td = ""
                if primeiro_cargo:
                    cargo_td = (f'<td rowspan="{len(rows)}" style="vertical-align:middle;font-weight:bold;">{cargo}</td>')
                    primeiro_cargo = False
                linhas_html += (
                    f"<tr>{ghe_td}{cargo_td}<td>{row['Exame']}</td>"
                    f"{cel(row['ADM'])}{per_td}{cel(row['MRO'])}{cel(row['RT'])}{cel(row['DEM'])}"
                    f"<td style='font-size:11px;color:#555;'>{row['Justificativa']}</td></tr>"
                )

    return f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">
<style>
  body{{font-family:Arial,sans-serif;font-size:13px;margin:20px;}}
  table{{width:100%;border-collapse:collapse;margin-top:10px;}}
  th{{background:#1AA04B;color:#fff;padding:10px 6px;border:1px solid #084D22;font-size:12px;}}
  th.c{{text-align:center;}} td{{border:1px solid #ccc;padding:8px 6px;vertical-align:middle;}}
  tr:nth-child(even) td{{background:#F4F8F5;}}
</style></head><body>
<table style="margin-bottom:12px;border:2px solid #084D22;">
  <tr style="background:#084D22;color:#fff;">
    <td colspan="5" style="padding:8px;font-size:12pt;font-weight:bold;text-align:center;">
      PROGRAMA DE CONTROLE MÉDICO DE SAÚDE OCUPACIONAL — PCMSO</td></tr>
  <tr><td><b>Empresa:</b> {razao}</td><td><b>CNPJ:</b> {cnpj}</td>
      <td><b>Obra:</b> {obra}</td><td><b>Vigência:</b> {vig_i} a {vig_f}</td>
      <td><b>Emissão:</b> {datetime.now().strftime("%d/%m/%Y")}</td></tr>
  <tr><td colspan="3"><b>Médico(a):</b> {medico}</td>
      <td colspan="2"><b>Técnico SST:</b> {tec}</td></tr>
</table>
<table>
  <tr><th style="width:12%">GHE</th><th style="width:14%">Função</th>
      <th style="width:30%">Exame Solicitado</th>
      <th class="c" style="width:5%">ADM</th><th class="c" style="width:6%">PER</th>
      <th class="c" style="width:5%">MRO</th><th class="c" style="width:4%">RT</th>
      <th class="c" style="width:5%">DEM</th><th style="width:19%">Justificativa</th></tr>
  {linhas_html}
</table>
<p style="font-size:8pt;color:#555;margin-top:12px;">
  Gerado por Sistema Automação SST Seconci-GO | NR-07 (Port.1.031/2018), NR-09, NR-15, NR-35, Dec.3.048/99.
</p></body></html>"""


# ── Geração DOCX ──────────────────────────────────────────────────────────────

def gerar_docx_pcmso(df: pd.DataFrame, cabecalho: dict = None) -> bytes:
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    if not cabecalho:
        cabecalho = {}
    razao  = cabecalho.get("razao_social","Empresa nao informada")
    cnpj   = cabecalho.get("cnpj","---")
    obra   = cabecalho.get("obra","---")
    medico = cabecalho.get("medico_rt","Nao informado")
    vig_i  = cabecalho.get("vig_ini","---")
    vig_f  = cabecalho.get("vig_fim","---")
    tec    = cabecalho.get("responsavel_tec","---")

    VERDE_ESC = RGBColor(0x08,0x4D,0x22)
    VERDE_CLR = RGBColor(0x1A,0xA0,0x4B)
    BRANCO    = RGBColor(0xFF,0xFF,0xFF)

    def shd(cell, hex_color):
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        s = OxmlElement("w:shd")
        s.set(qn("w:val"),"clear"); s.set(qn("w:color"),"auto")
        s.set(qn("w:fill"), hex_color); tcPr.append(s)

    def txt(cell, text, bold=False, color=None, size=9, align=WD_ALIGN_PARAGRAPH.LEFT):
        cell.text = ""
        p = cell.paragraphs[0]; p.alignment = align
        r = p.add_run(text); r.bold = bold; r.font.size = Pt(size)
        if color: r.font.color.rgb = color

    doc = Document()
    for sec in doc.sections:
        sec.top_margin=sec.bottom_margin=Cm(1.5)
        sec.left_margin=sec.right_margin=Cm(1.5)

    doc.add_paragraph()
    cab = doc.add_table(rows=3, cols=5); cab.style="Table Grid"
    cab.rows[0].cells[0].merge(cab.rows[0].cells[4])
    shd(cab.rows[0].cells[0],"084D22")
    txt(cab.rows[0].cells[0],"PROGRAMA DE CONTROLE MÉDICO DE SAÚDE OCUPACIONAL — PCMSO",
        bold=True,color=BRANCO,size=11,align=WD_ALIGN_PARAGRAPH.CENTER)
    for i,t in enumerate([f"Empresa: {razao}",f"CNPJ: {cnpj}",f"Obra: {obra}",
                           f"Vigência: {vig_i} a {vig_f}",
                           f"Emissão: {datetime.now().strftime('%d/%m/%Y')}"]):
        txt(cab.rows[1].cells[i],t,size=9)
    cab.rows[2].cells[0].merge(cab.rows[2].cells[2])
    cab.rows[2].cells[3].merge(cab.rows[2].cells[4])
    txt(cab.rows[2].cells[0],f"Médico(a): {medico}",size=9)
    txt(cab.rows[2].cells[3],f"Técnico SST: {tec}",size=9)
    doc.add_paragraph()

    cols_n = ["GHE","Função","Exame Solicitado","ADM","PER","MRO","RT","DEM","Justificativa"]
    cols_w = [Cm(2.5),Cm(3.0),Cm(6.0),Cm(1.0),Cm(1.2),Cm(1.0),Cm(0.9),Cm(1.0),Cm(3.9)]
    tab = doc.add_table(rows=1,cols=len(cols_n)); tab.style="Table Grid"
    for i,(cn,cw) in enumerate(zip(cols_n,cols_w)):
        c=tab.rows[0].cells[i]; shd(c,"1AA04B"); c.width=cw
        txt(c,cn,bold=True,color=BRANCO,size=9,
            align=WD_ALIGN_PARAGRAPH.CENTER if i>=3 else WD_ALIGN_PARAGRAPH.LEFT)

    ghe_grupos={}
    for _,row in df.iterrows():
        ghe_grupos.setdefault(row["GHE / Setor"],{}).setdefault(row["Cargo"],[]).append(row)

    for ghe_nome,cargos_dict in ghe_grupos.items():
        first_ghe=None
        for cargo,rows in cargos_dict.items():
            first_cargo=None
            for row in rows:
                tr=tab.add_row()
                for i,cw in enumerate(cols_w):
                    tr.cells[i].width=cw
                    tr.cells[i].vertical_alignment=WD_ALIGN_VERTICAL.CENTER
                if first_ghe is None:
                    first_ghe=tr; shd(tr.cells[0],"084D22")
                    txt(tr.cells[0],ghe_nome,bold=True,color=BRANCO,size=8,
                        align=WD_ALIGN_PARAGRAPH.CENTER)
                else:
                    shd(tr.cells[0],"084D22"); tr.cells[0].text=""
                if first_cargo is None:
                    first_cargo=tr; txt(tr.cells[1],cargo,bold=True,size=9)
                else:
                    tr.cells[1].text=""
                txt(tr.cells[2],str(row["Exame"]),size=9)
                for idx,col in enumerate(["ADM","PER","MRO","RT","DEM"],start=3):
                    val=str(row[col]); is_x=(val=="X")
                    if is_x: shd(tr.cells[idx],"d4edda")
                    txt(tr.cells[idx],val,bold=is_x,size=9,
                        align=WD_ALIGN_PARAGRAPH.CENTER)
                txt(tr.cells[8],str(row["Justificativa"]),size=8)

    doc.add_paragraph()
    rod=doc.add_paragraph(
        "Gerado por Sistema Automação SST Seconci-GO  |  "
        "NR-07 (Port.1.031/2018), NR-09, NR-15, NR-35, Decreto 3.048/99.")
    rod.runs[0].font.size=Pt(7)
    rod.runs[0].font.color.rgb=RGBColor(0x55,0x55,0x55)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()

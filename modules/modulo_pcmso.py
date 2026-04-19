"""
modules/modulo_pcmso.py
Modulo Medicina: PGR -> PCMSO
Logica pura — sem Streamlit. UI fica no app.py.

VERSAO v3 — melhorias:
  [Bug 1] Regex seletivo para deteccao de GHE real
  [Bug 2] normalizar_texto() resolve divergencia de acentos
  [Bug 3] Criterio de fallback por qualidade, nao apenas quantidade
  [Novo]  processar_pcmso() gera colunas ADM/PER/MRO/RT/DEM por exame
  [Novo]  gerar_html_pcmso() — tabela no modelo do adendo Dinamica Engenharia
  [Novo]  gerar_docx_pcmso() — exporta DOCX editavel (python-docx)
"""

import io
import re
import unicodedata
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
}

_GHES_ADMIN = {
    "ADMINISTRACAO","PLANEJAMENTO","SUPRIMENTOS","MARKETING",
    "TI","RH","RECURSOS HUMANOS","FINANCEIRO","CONTABILIDADE","ESCRITORIO",
}

_MAPA_AGENTES = {
    "TOLUENO":"Tolueno","XILENO":"Xileno","BENZENO":"Benzeno",
    "ACETONA":"Acetona","THINNER":"Solventes (Thinner)",
    "SOLVENTE":"Solventes Organicos","TINTA":"Tinta / Verniz",
    "VERNIZ":"Tinta / Verniz","PRIMER":"Primer (Solventes)",
    "GRAXA":"Graxas / Lubrificantes","DIESEL":"Diesel / Combustivel",
    "QUEROSENE":"Querosene","ACIDO":"Acidos (geral)",
    "CIMENTO":"Cimento Portland","SILICA":"Silica Cristalina (Quartzo)",
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
_PALAVRAS_GHE_FRACAS = ["DEPARTAMENTO", "ATIVIDADE"]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _is_linha_ghe(linha: str) -> bool:
    if _RE_GHE.search(linha):
        return True
    lu = normalizar_texto(linha)
    if len(linha) <= 80 and "/" not in linha:
        if any(p in lu for p in _PALAVRAS_GHE_FRACAS):
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
    """Formata periodicidade: '12 MESES' -> '12M', '06 MESES' -> '6M', '' -> '-'"""
    if not per:
        return "-"
    per = per.strip().upper().replace("MESES", "").replace("MESES", "").strip()
    try:
        return f"{int(per)}M"
    except ValueError:
        return per


def _flag(val: bool) -> str:
    return "X" if val else "-"


# ── Extração de texto e PGR ───────────────────────────────────────────────────

def extrair_texto_pdf(uploaded_file) -> str:
    texto = []
    with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
            if t:
                texto.append(t)
    return "\n".join(texto)


def extrair_texto_pdf_path(caminho: str) -> str:
    texto = []
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
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
                    "nome_agente":       agente,
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
            print(f"[WARN] Falha na IA: {e}")
    return (local or [], "parcial")


# ── Processamento PCMSO ───────────────────────────────────────────────────────

def processar_pcmso(dados_pgr: list) -> pd.DataFrame:
    """
    Gera DataFrame com colunas:
      GHE / Setor | Cargo | Exame | ADM | PER | MRO | RT | DEM | Justificativa
    """
    linhas = []

    for ghe in dados_pgr:
        nome_ghe  = ghe.get("ghe", "Sem GHE")
        nome_norm = normalizar_texto(nome_ghe)
        cargos    = ghe.get("cargos", [])[:15]
        riscos    = ghe.get("riscos_mapeados", [])[:10]

        if not _ghe_valido(nome_ghe):
            continue

        for cargo in cargos:
            exames: dict = {}

            # Exame clínico base — sempre (NR-07 básico)
            adicionar_exame_dedup(exames, {
                "exame": "Exame Clinico (Anamnese / Exame Fisico)",
                "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": True,
                "motivo": "NR-07 Basico",
            })

            cargo_norm  = normalizar_cargo(cargo)
            cargo_upper = cargo_norm.upper()

            # Exames por função
            for funcao, lista_ex in MATRIZ_FUNCAO_EXAME.items():
                if normalizar_texto(funcao) in cargo_upper:
                    for ex in lista_ex:
                        adicionar_exame_dedup(exames, {**ex, "motivo": f"Funcao: {funcao.title()}"})

            # Exames por risco
            bio_real = tem_risco_biologico_real(riscos)
            for risco in riscos:
                texto_r = normalizar_texto(
                    risco.get("nome_agente", "") + " " + risco.get("perigo_especifico", "")
                )

                for chave_r, regra in MATRIZ_RISCO_EXAME.items():
                    if chave_r in CHAVES_BIOLOGICAS_MATRIZ and not bio_real:
                        continue
                    if chave_r in texto_r:
                        adicionar_exame_dedup(exames, {
                            "exame":   regra["exame"],
                            "adm":     regra.get("adm", True),
                            "per":     regra.get("periodico", "12 MESES"),
                            "mro":     regra.get("mro", True),
                            "rt":      regra.get("rt", False),
                            "dem":     regra.get("dem", False),
                            "motivo":  f"Exposicao: {chave_r.title()} — {regra.get('obs','')}",
                        })

                _admin = (
                    any(a in cargo_upper for a in _CARGOS_ADMIN) or
                    any(g in nome_norm   for g in _GHES_ADMIN)
                )
                if "ALTURA" in texto_r and not _admin:
                    for ex in MATRIZ_FUNCAO_EXAME.get("TRABALHO EM ALTURA", []):
                        adicionar_exame_dedup(exames, {**ex, "motivo": "Trabalho em Altura (NR-35)"})

            for ex_info in exames.values():
                per_fmt = _fmt_per(ex_info.get("per", "12 MESES"))
                linhas.append({
                    "GHE / Setor":     nome_ghe,
                    "Cargo":           cargo,
                    "Exame":           ex_info["exame"],
                    "ADM":             _flag(ex_info.get("adm", True)),
                    "PER":             per_fmt,
                    "MRO":             _flag(ex_info.get("mro", True)),
                    "RT":              _flag(ex_info.get("rt", False)),
                    "DEM":             _flag(ex_info.get("dem", False)),
                    "Justificativa":   ex_info.get("motivo", ""),
                })

    return pd.DataFrame(linhas)


# ── Geração HTML ──────────────────────────────────────────────────────────────

def gerar_html_pcmso(df: pd.DataFrame, cabecalho: dict = None) -> str:
    """HTML no modelo do adendo Dinamica Engenharia — tabela com colunas ADM/PER/MRO/RT/DEM."""
    if not cabecalho:
        cabecalho = {}

    razao  = cabecalho.get("razao_social",    "Empresa nao informada")
    cnpj   = cabecalho.get("cnpj",            "---")
    obra   = cabecalho.get("obra",            "---")
    medico = cabecalho.get("medico_rt",       "Nao informado")
    vig_i  = cabecalho.get("vig_ini",         "---")
    vig_f  = cabecalho.get("vig_fim",         "---")
    tec    = cabecalho.get("responsavel_tec", "---")

    # Agrupa por GHE
    ghe_grupos = {}
    for _, row in df.iterrows():
        ghe_key = row["GHE / Setor"]
        if ghe_key not in ghe_grupos:
            ghe_grupos[ghe_key] = {}
        cargo = row["Cargo"]
        if cargo not in ghe_grupos[ghe_key]:
            ghe_grupos[ghe_key][cargo] = []
        ghe_grupos[ghe_key][cargo].append(row)

    linhas_html = ""
    for ghe_nome, cargos_dict in ghe_grupos.items():
        total_rows = sum(len(v) for v in cargos_dict.values())
        primeiro_ghe = True
        for cargo, rows in cargos_dict.items():
            primeiro_cargo = True
            for row in rows:
                adm_cell = f'<td style="text-align:center;background:#d4edda;">{row["ADM"]}</td>' if row["ADM"] == "X" else '<td style="text-align:center;">-</td>'
                per_cell = f'<td style="text-align:center;font-weight:bold;">{row["PER"]}</td>' if row["PER"] != "-" else '<td style="text-align:center;color:#999;">-</td>'
                mro_cell = f'<td style="text-align:center;background:#d4edda;">{row["MRO"]}</td>' if row["MRO"] == "X" else '<td style="text-align:center;">-</td>'
                rt_cell  = f'<td style="text-align:center;background:#d4edda;">{row["RT"]}</td>' if row["RT"]  == "X" else '<td style="text-align:center;">-</td>'
                dem_cell = f'<td style="text-align:center;background:#d4edda;">{row["DEM"]}</td>' if row["DEM"] == "X" else '<td style="text-align:center;">-</td>'

                if primeiro_ghe:
                    ghe_cell = f'<td rowspan="{total_rows}" style="background:#084D22;color:#fff;font-weight:bold;vertical-align:middle;text-align:center;padding:8px;">{ghe_nome}</td>'
                    primeiro_ghe = False
                else:
                    ghe_cell = ""

                cargo_rows = len(rows)
                if primeiro_cargo:
                    cargo_cell = f'<td rowspan="{cargo_rows}" style="vertical-align:middle;font-weight:bold;">{cargo}</td>'
                    primeiro_cargo = False
                else:
                    cargo_cell = ""

                linhas_html += (
                    f"<tr>{ghe_cell}{cargo_cell}"
                    f"<td>{row['Exame']}</td>"
                    f"{adm_cell}{per_cell}{mro_cell}{rt_cell}{dem_cell}"
                    f"<td style='font-size:11px;color:#555;'>{row['Justificativa']}</td></tr>"
                )

    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8">
<style>
  body {{font-family:Arial,sans-serif;color:#000;margin:20px;font-size:13px;}}
  table {{width:100%;border-collapse:collapse;margin-top:10px;}}
  th {{background:#1AA04B;color:#fff;padding:10px 6px;text-align:left;border:1px solid #084D22;font-size:12px;}}
  th.center {{text-align:center;}}
  td {{border:1px solid #ccc;padding:8px 6px;vertical-align:middle;}}
  tr:nth-child(even) td {{background:#F4F8F5;}}
  .cab td {{border:1px solid #ccc;padding:5px 8px;font-size:9pt;}}
</style></head><body>

<table class="cab" style="margin-bottom:12px;border:2px solid #084D22;">
  <tr style="background:#084D22;color:#fff;">
    <td colspan="5" style="padding:8px;font-size:12pt;font-weight:bold;text-align:center;">
      PROGRAMA DE CONTROLE MÉDICO DE SAÚDE OCUPACIONAL — PCMSO
    </td>
  </tr>
  <tr>
    <td><b>Empresa:</b> {razao}</td>
    <td><b>CNPJ:</b> {cnpj}</td>
    <td><b>Obra / Unidade:</b> {obra}</td>
    <td><b>Vigência:</b> {vig_i} a {vig_f}</td>
    <td><b>Emissão:</b> {datetime.now().strftime("%d/%m/%Y")}</td>
  </tr>
  <tr>
    <td colspan="3"><b>Médico(a) Coordenador(a):</b> {medico}</td>
    <td colspan="2"><b>Técnico SST:</b> {tec}</td>
  </tr>
</table>

<table>
  <tr>
    <th style="width:12%">GHE</th>
    <th style="width:14%">Função</th>
    <th style="width:30%">Exame Solicitado</th>
    <th class="center" style="width:5%">ADM</th>
    <th class="center" style="width:6%">PER</th>
    <th class="center" style="width:5%">MRO</th>
    <th class="center" style="width:4%">RT</th>
    <th class="center" style="width:5%">DEM</th>
    <th style="width:19%">Justificativa / Agente</th>
  </tr>
  {linhas_html}
</table>

<p style="font-size:8pt;color:#555;margin-top:12px;">
  Gerado por Sistema Automação SST Seconci-GO &nbsp;|&nbsp;
  Base legal: NR-07 (Port. 1.031/2018), NR-09, NR-15, NR-35, Decreto 3.048/99.
</p>
</body></html>"""


# ── Geração DOCX ──────────────────────────────────────────────────────────────

def gerar_docx_pcmso(df: pd.DataFrame, cabecalho: dict = None) -> bytes:
    """
    Gera DOCX editável no modelo do adendo Dinamica Engenharia.
    Retorna bytes prontos para st.download_button() ou gravação em arquivo.
    Requer: pip install python-docx
    """
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    if not cabecalho:
        cabecalho = {}

    razao  = cabecalho.get("razao_social",    "Empresa nao informada")
    cnpj   = cabecalho.get("cnpj",            "---")
    obra   = cabecalho.get("obra",            "---")
    medico = cabecalho.get("medico_rt",       "Nao informado")
    vig_i  = cabecalho.get("vig_ini",         "---")
    vig_f  = cabecalho.get("vig_fim",         "---")
    tec    = cabecalho.get("responsavel_tec", "---")

    VERDE_ESCURO = RGBColor(0x08, 0x4D, 0x22)
    VERDE_CLARO  = RGBColor(0x1A, 0xA0, 0x4B)
    BRANCO       = RGBColor(0xFF, 0xFF, 0xFF)

    def set_cell_bg(cell, hex_color: str):
        tc   = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        tcPr.append(shd)

    def cell_text(cell, text: str, bold=False, color=None, size=9, align=WD_ALIGN_PARAGRAPH.LEFT):
        cell.text = ""
        p   = cell.paragraphs[0]
        p.alignment = align
        run = p.add_run(text)
        run.bold      = bold
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color

    doc = Document()

    # Margens
    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(1.5)
        section.right_margin  = Cm(1.5)

    # ── Cabeçalho ──
    doc.add_paragraph()
    cab = doc.add_table(rows=3, cols=5)
    cab.style = "Table Grid"

    # Linha 0 — título
    cab.rows[0].cells[0].merge(cab.rows[0].cells[4])
    set_cell_bg(cab.rows[0].cells[0], "084D22")
    cell_text(cab.rows[0].cells[0],
              "PROGRAMA DE CONTROLE MÉDICO DE SAÚDE OCUPACIONAL — PCMSO",
              bold=True, color=BRANCO, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)

    # Linha 1
    dados_l1 = [
        f"Empresa: {razao}", f"CNPJ: {cnpj}",
        f"Obra: {obra}", f"Vigência: {vig_i} a {vig_f}",
        f"Emissão: {datetime.now().strftime('%d/%m/%Y')}",
    ]
    for i, txt in enumerate(dados_l1):
        cell_text(cab.rows[1].cells[i], txt, size=9)

    # Linha 2
    cab.rows[2].cells[0].merge(cab.rows[2].cells[2])
    cab.rows[2].cells[3].merge(cab.rows[2].cells[4])
    cell_text(cab.rows[2].cells[0], f"Médico(a) Coordenador(a): {medico}", size=9)
    cell_text(cab.rows[2].cells[3], f"Técnico SST: {tec}", size=9)

    doc.add_paragraph()

    # ── Tabela principal ──
    colunas = ["GHE", "Função", "Exame Solicitado", "ADM", "PER", "MRO", "RT", "DEM", "Justificativa"]
    larguras = [Cm(2.5), Cm(3.0), Cm(6.0), Cm(1.0), Cm(1.2), Cm(1.0), Cm(0.9), Cm(1.0), Cm(3.9)]

    tab = doc.add_table(rows=1, cols=len(colunas))
    tab.style = "Table Grid"

    # Cabeçalho da tabela
    for i, (col, larg) in enumerate(zip(colunas, larguras)):
        c = tab.rows[0].cells[i]
        set_cell_bg(c, "1AA04B")
        cell_text(c, col, bold=True, color=BRANCO, size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER if i >= 3 else WD_ALIGN_PARAGRAPH.LEFT)
        c.width = larg

    # Agrupa por GHE e Cargo para rowspan
    ghe_grupos = {}
    for _, row in df.iterrows():
        g = row["GHE / Setor"]
        c = row["Cargo"]
        ghe_grupos.setdefault(g, {}).setdefault(c, []).append(row)

    for ghe_nome, cargos_dict in ghe_grupos.items():
        total_rows_ghe = sum(len(v) for v in cargos_dict.values())
        first_ghe_row = None

        for cargo, rows in cargos_dict.items():
            first_cargo_row = None

            for row in rows:
                tr = tab.add_row()
                for i, larg in enumerate(larguras):
                    tr.cells[i].width = larg
                    tr.cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                # GHE — só primeira célula; demais serão mergeadas depois
                if first_ghe_row is None:
                    first_ghe_row = tr
                    set_cell_bg(tr.cells[0], "084D22")
                    cell_text(tr.cells[0], ghe_nome, bold=True, color=BRANCO, size=8,
                              align=WD_ALIGN_PARAGRAPH.CENTER)
                else:
                    set_cell_bg(tr.cells[0], "084D22")
                    tr.cells[0].text = ""

                # Cargo — só primeira célula do cargo
                if first_cargo_row is None:
                    first_cargo_row = tr
                    cell_text(tr.cells[1], cargo, bold=True, size=9)
                else:
                    tr.cells[1].text = ""

                # Exame
                cell_text(tr.cells[2], str(row["Exame"]), size=9)

                # ADM/PER/MRO/RT/DEM
                for idx, col in enumerate(["ADM", "PER", "MRO", "RT", "DEM"], start=3):
                    val = str(row[col])
                    is_x = val == "X"
                    if is_x:
                        set_cell_bg(tr.cells[idx], "d4edda")
                    cell_text(tr.cells[idx], val, bold=is_x, size=9,
                              align=WD_ALIGN_PARAGRAPH.CENTER)

                # Justificativa
                cell_text(tr.cells[8], str(row["Justificativa"]), size=8)

    # Rodapé
    doc.add_paragraph()
    rod = doc.add_paragraph(
        f"Gerado por Sistema Automação SST Seconci-GO  |  "
        f"Base legal: NR-07 (Port. 1.031/2018), NR-09, NR-15, NR-35, Decreto 3.048/99."
    )
    rod.runs[0].font.size = Pt(7)
    rod.runs[0].font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

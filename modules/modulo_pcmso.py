"""
Modulo Medicina: PGR -> PCMSO
Logica pura — sem Streamlit. UI fica no app.py.

CORRECOES v2:
  [Bug 1] Trigger de GHE reescrito com regex — evita falsos positivos em
          linhas "CARGO/FUNCAO:", "DESCRICAO DAS ATIVIDADES EXERCIDAS:", etc.
  [Bug 2] normalizar_texto() aplica unicodedata antes de todas as comparacoes,
          resolvendo divergencia DESCRICAO vs DESCRIÇAO em _INVALIDOS_GHE.
  [Bug 3] Criterio de fallback agora exige GHE com nome curto e valido + cargos,
          evitando falso positivo quando so ha lixo textual extraido.
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

# ── Constantes ──────────────────────────────────────────────────────────

# [Bug 2 FIX] Todos os termos sem acento para bater com normalizar_texto()
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
    "TOLUENO":      "Tolueno",
    "XILENO":       "Xileno",
    "BENZENO":      "Benzeno",
    "ACETONA":      "Acetona",
    "THINNER":      "Solventes (Thinner)",
    "SOLVENTE":     "Solventes Organicos",
    "TINTA":        "Tinta / Verniz (Solventes)",
    "VERNIZ":       "Tinta / Verniz (Solventes)",
    "PRIMER":       "Primer (Solventes)",
    "GRAXA":        "Graxas / Lubrificantes",
    "DIESEL":       "Diesel / Combustivel",
    "QUEROSENE":    "Querosene",
    "ACIDO":        "Acidos (geral)",
    "CIMENTO":      "Cimento Portland (Poeiras)",
    "SILICA":       "Silica Cristalina (Quartzo)",
    "POEIRA":       "Poeiras Minerais",
    "AMIANTO":      "Asbesto / Amianto",
    "CHUMBO":       "Chumbo (Fumos/Poeiras)",
    "RUIDO":        "Ruido Continuo ou Intermitente",
    "VIBRACAO":     "Vibracao (VMB/VCI)",
    "CALOR":        "Calor (IBUTG)",
    "RADIACAO":     "Radiacoes Ionizantes",
    "BIOLOGICO":    "Agentes Biologicos",
    "ESGOTO":       "Esgoto / Aguas Servidas",
    "SANGUE":       "Material Biologico (Sangue/Fluidos)",
    "ERGONO":       "Fator Ergonomico",
    "POSTURA":      "Postura Inadequada",
    "LEVANTAMENTO": "Levantamento de Carga",
    "REPETITIVO":   "Movimento Repetitivo",
    "ELETRICO":     "Risco Eletrico",
    "ALTURA":       "Queda de Altura",
    "MAQUINA":      "Maquinas e Equipamentos",
    "INCENDIO":     "Incendio / Explosao",
}

# [Bug 1 FIX] Regex seletivo para linhas que realmente marcam um GHE
_RE_GHE = re.compile(
    r"(?:GHE[\s:.\-]+\w|GRUPO\s+HOMOGENEO|LOCAL\s+DE\s+TRABALHO\s*:\s*\w|SETOR\s*:\s*\w)",
    re.IGNORECASE,
)

_PALAVRAS_GHE_FRACAS = ["DEPARTAMENTO", "ATIVIDADE"]


def _is_linha_ghe(linha: str) -> bool:
    """[Bug 1 FIX] Detecta GHE real via regex forte ou palavras fracas em linhas curtas."""
    if _RE_GHE.search(linha):
        return True
    lu = normalizar_texto(linha)
    if len(linha) <= 80 and "/" not in linha:
        if any(p in lu for p in _PALAVRAS_GHE_FRACAS):
            return True
    return False


def _ghe_valido(nome_ghe: str) -> bool:
    """[Bug 2 FIX] Compara _INVALIDOS_GHE após normalização de acentos."""
    norm = normalizar_texto(nome_ghe)
    return not any(inv in norm for inv in _INVALIDOS_GHE)


def _fallback_necessario(ghes: list) -> bool:
    """
    [Bug 3 FIX] Criterio de qualidade real:
    Exige ao menos 1 GHE com nome normalizado <= 60 chars e pelo menos 1 cargo.
    """
    for g in ghes:
        nome_norm = normalizar_texto(g["ghe"])
        if len(nome_norm) <= 60 and g["cargos"]:
            return False
    return True


# ── Funções principais ──────────────────────────────────────────────────

def extrair_texto_pdf(uploaded_file) -> str:
    """Extrai texto de UploadedFile do Streamlit via pdfplumber."""
    texto = []
    with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
            if t:
                texto.append(t)
    return "\n".join(texto)


def extrair_texto_pdf_path(caminho: str) -> str:
    """Versao standalone (testes sem Streamlit) — aceita path local."""
    texto = []
    with pdfplumber.open(caminho) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
            if t:
                texto.append(t)
    return "\n".join(texto)


def extrair_pgr_local(texto: str) -> list:
    """
    Camada 1 — extracao por palavras-chave, sem IA.
    Bugs 1 e 2 corrigidos aqui.
    """
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
                cargo_norm = normalizar_texto(cargo)
                if cargo_norm in lu and cargo not in ghe_atual["cargos"]:
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
    """
    Motor cascata: Local -> IA.
    Retorna (dados, fonte) onde fonte = "local" | "ia" | "parcial".
    [Bug 3 FIX] _fallback_necessario() verifica qualidade real dos GHEs.
    """
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


def processar_pcmso(dados_pgr: list) -> pd.DataFrame:
    """Gera DataFrame com 1 linha por exame por cargo por GHE."""
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
            adicionar_exame_dedup(exames, {
                "exame":         "Exame Clinico (Anamnese / Exame Fisico)",
                "periodicidade": "12 MESES",
                "motivo":        "NR-07 Basico",
            })

            cargo_norm  = normalizar_cargo(cargo)
            cargo_upper = cargo_norm.upper()

            for funcao, lista_ex in MATRIZ_FUNCAO_EXAME.items():
                if funcao in cargo_upper:
                    for ex in lista_ex:
                        adicionar_exame_dedup(exames, {**ex, "motivo": f"Funcao: {funcao.title()}"})

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
                            "exame":         regra["exame"],
                            "periodicidade": regra["periodico"],
                            "motivo":        f"Exposicao: {chave_r.title()}",
                        })

                _admin = (
                    any(a in cargo_upper for a in _CARGOS_ADMIN) or
                    any(g in nome_norm   for g in _GHES_ADMIN)
                )
                if "ALTURA" in texto_r and not _admin:
                    for ex in MATRIZ_FUNCAO_EXAME.get("TRABALHO EM ALTURA", []):
                        adicionar_exame_dedup(exames, {**ex, "motivo": "Trabalho em Altura (NR-35)"})

            for ex_info in exames.values():
                linhas.append({
                    "GHE / Setor":                 nome_ghe,
                    "Cargo":                       cargo,
                    "Exame Clinico/Complementar":  ex_info["exame"],
                    "Periodicidade":               ex_info["periodicidade"],
                    "Justificativa Legal / Risco": ex_info["motivo"],
                })

    return pd.DataFrame(linhas)


def gerar_html_pcmso(df: pd.DataFrame, cabecalho: dict = None) -> str:
    """Gera HTML completo do PCMSO a partir do DataFrame."""
    if not cabecalho:
        cabecalho = {}

    razao  = cabecalho.get("razao_social",    "Empresa nao informada")
    cnpj   = cabecalho.get("cnpj",            "---")
    medico = cabecalho.get("medico_rt",       "Nao informado")
    vig_i  = cabecalho.get("vig_ini",         "---")
    vig_f  = cabecalho.get("vig_fim",         "---")
    tec    = cabecalho.get("responsavel_tec", "---")

    linhas_html = ""
    for _, row in df.iterrows():
        cargo_val  = str(row["Cargo"])
        pendente   = "verificar" in cargo_val.lower() or not cargo_val.strip()
        bg         = "background-color:#FFF3CD;" if pendente else ""
        cargo_disp = (
            f'ATENCAO: {cargo_val} <small style="color:#c0392b">(Cargo nao identificado)</small>'
            if pendente else cargo_val
        )
        linhas_html += (
            f'<tr style="{bg}">'
            f'<td><strong>{row["GHE / Setor"]}</strong></td>'
            f'<td>{cargo_disp}</td>'
            f'<td>{row["Exame Clinico/Complementar"]}</td>'
            f'<td>{row["Periodicidade"]}</td>'
            f'<td>{row["Justificativa Legal / Risco"]}</td></tr>'
        )

    return f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">
<style>
  body{{font-family:Arial,sans-serif;color:#000;margin:20px;}}
  table{{width:100%;border-collapse:collapse;font-size:13px;margin-top:10px;}}
  th{{background:#1AA04B;color:#fff;padding:12px 8px;text-align:left;border:1px solid #000;}}
  td{{border:1px solid #000;padding:10px 8px;vertical-align:top;}}
  tr:nth-child(even){{background:#F4F8F5;}}
  .cab td{{border:1px solid #ccc;padding:6px;font-size:9pt;}}
</style></head><body>
<table class="cab" style="width:100%;border-collapse:collapse;margin-bottom:12px;border:1px solid #084D22;">
  <tr style="background:#084D22;color:#fff">
    <td colspan="4" style="padding:8px;font-size:11pt;font-weight:bold;">
      PROGRAMA DE CONTROLE MEDICO DE SAUDE OCUPACIONAL - PCMSO
    </td>
  </tr>
  <tr>
    <td width="40%"><b>Empresa:</b> {razao}</td>
    <td width="20%"><b>CNPJ:</b> {cnpj}</td>
    <td width="20%"><b>Vigencia:</b> {vig_i} a {vig_f}</td>
    <td width="20%"><b>Emissao:</b> {datetime.now().strftime("%d/%m/%Y")}</td>
  </tr>
  <tr>
    <td colspan="2"><b>Medico RT:</b> {medico}</td>
    <td colspan="2"><b>Tecnico SST:</b> {tec}</td>
  </tr>
</table>
<table>
  <tr>
    <th>GHE / Setor</th><th>Cargo</th>
    <th>Exame Clinico / Complementar</th>
    <th>Periodicidade</th><th>Justificativa / Agente</th>
  </tr>
  {linhas_html}
</table>
<p style="font-size:8pt;color:#555;margin-top:10px;">
  Gerado por Sistema Automacao SST Seconci-GO |
  Base legal: NR-07 (Port. 1.031/2018), NR-09, Decreto 3.048/99.
</p>
</body></html>"""
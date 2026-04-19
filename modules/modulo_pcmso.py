"""
Modulo Medicina: PGR -> PCMSO
Logica pura — sem Streamlit. UI fica no app.py.
"""
import io
import pdfplumber
import pandas as pd
from datetime import datetime

import streamlit as st

from data.matriz_exames    import MATRIZ_FUNCAO_EXAME, MATRIZ_RISCO_EXAME
from utils.cargo_utils     import normalizar_cargo, MAPA_CARGOS_CONHECIDOS, PALAVRAS_EXCLUIR_CARGO
from utils.exame_utils     import adicionar_exame_dedup
from utils.biologico_utils import tem_risco_biologico_real, CHAVES_BIOLOGICAS_MATRIZ
from utils.ia_client       import extrair_pgr_via_ia

_INVALIDOS_GHE = [
    "QUANTIDADE","PREVISTOS","EXPOSTOS","TOTAL DE","NUMERO DE",
    "FUNCIONARIOS","TRABALHADORES","MEDIDAS DE CONTROLE",
    "FONTE GERADORA","TRAJETORIA","DESCRICAO",
]

_CARGOS_ADMIN = {
    "ADMINISTRATIVO","RECEPCIONISTA","GESTOR","ANALISTA","ASSISTENTE",
    "COORDENADOR","DIRETOR","GERENTE","ESTAGIARIO","JOVEM APRENDIZ",
}

_GHES_ADMIN = {
    "ADMINISTRACAO","PLANEJAMENTO","SUPRIMENTOS","MARKETING",
    "TI","RH","RECURSOS HUMANOS","FINANCEIRO","CONTABILIDADE",
}

_MAPA_AGENTES = {
    "TOLUENO":"Tolueno","XILENO":"Xileno","BENZENO":"Benzeno",
    "ACETONA":"Acetona","THINNER":"Solventes (Thinner)",
    "SOLVENTE":"Solventes Organicos","TINTA":"Tinta / Verniz (Solventes)",
    "VERNIZ":"Tinta / Verniz (Solventes)","PRIMER":"Primer (Solventes)",
    "GRAXA":"Graxas / Lubrificantes","DIESEL":"Diesel / Combustivel",
    "QUEROSENE":"Querosene","ACIDO":"Acidos (geral)",
    "CIMENTO":"Cimento Portland (Poeiras)","SILICA":"Silica Cristalina (Quartzo)",
    "POEIRA":"Poeiras Minerais","AMIANTO":"Asbesto / Amianto",
    "CHUMBO":"Chumbo (Fumos/Poeiras)",
    "RUIDO":"Ruido Continuo ou Intermitente","VIBRACAO":"Vibracao (VMB/VCI)",
    "CALOR":"Calor (IBUTG)","RADIACAO":"Radiacoes Ionizantes",
    "BIOLOGICO":"Agentes Biologicos","ESGOTO":"Esgoto / Aguas Servidas",
    "SANGUE":"Material Biologico (Sangue/Fluidos)",
    "ERGONO":"Fator Ergonomico","POSTURA":"Postura Inadequada",
    "LEVANTAMENTO":"Levantamento de Carga","REPETITIVO":"Movimento Repetitivo",
    "ELETRICO":"Risco Eletrico","ALTURA":"Queda de Altura",
    "MAQUINA":"Maquinas e Equipamentos","INCENDIO":"Incendio / Explosao",
}

_PALAVRAS_GHE = [
    "GHE","GRUPO HOMOGENEO","SETOR","DEPARTAMENTO","FUNCAO","CARGO","ATIVIDADE",
]


def extrair_texto_pdf(uploaded_file) -> str:
    """Extrai texto de UploadedFile do Streamlit via pdfplumber."""
    texto = []
    with pdfplumber.open(io.BytesIO(uploaded_file.read())) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text()
            if t:
                texto.append(t)
    return "\n".join(texto)


def extrair_pgr_local(texto: str) -> list:
    """Camada 1 — extracao por palavras-chave, sem IA."""
    linhas = texto.split("\n")
    ghes, ghe_atual, agentes_set = [], None, set()

    for linha in linhas:
        lc = linha.strip()
        if not lc:
            continue
        lu = lc.upper()

        if any(p in lu for p in _PALAVRAS_GHE) and len(lc) < 120:
            if ghe_atual and (ghe_atual["cargos"] or ghe_atual["riscos_mapeados"]):
                ghes.append(ghe_atual)
            ghe_atual   = {"ghe": lc, "cargos": [], "riscos_mapeados": []}
            agentes_set = set()
            continue

        if ghe_atual is None:
            continue

        if not any(exc in lu for exc in PALAVRAS_EXCLUIR_CARGO):
            for cargo in MAPA_CARGOS_CONHECIDOS:
                if cargo in lu and cargo not in [c.upper() for c in ghe_atual["cargos"]]:
                    ghe_atual["cargos"].append(cargo.title())
                    break

        for palavra, agente in _MAPA_AGENTES.items():
            if palavra in lu and agente not in agentes_set:
                agentes_set.add(agente)
                ghe_atual["riscos_mapeados"].append({
                    "nome_agente":        agente,
                    "perigo_especifico":  lc[:200],
                })

    if ghe_atual and (ghe_atual["cargos"] or ghe_atual["riscos_mapeados"]):
        ghes.append(ghe_atual)
    return ghes


def extrair_pgr_com_fallback(texto_pgr: str) -> tuple:
    """Motor cascata: Local -> IA. Retorna (dados, fonte)."""
    chave = str(st.secrets["CHAVE_API_GOOGLE"]).strip().replace('"', "").replace("'", "")
    local = extrair_pgr_local(texto_pgr)
    if len(local) >= 2 and any(g["cargos"] or g["riscos_mapeados"] for g in local):
        return local, "local"
    st.info("Extracao local insuficiente — acionando IA (Gemini)...")
    ia = extrair_pgr_via_ia(texto_pgr, chave)
    return (ia, "ia") if ia else (local or [], "parcial")


def processar_pcmso(dados_pgr: list) -> pd.DataFrame:
    """Gera DataFrame com 1 linha por exame por cargo por GHE."""
    linhas = []
    for ghe in dados_pgr:
        nome_ghe = ghe.get("ghe", "Sem GHE")
        cargos   = ghe.get("cargos", [])[:15]
        riscos   = ghe.get("riscos_mapeados", [])[:10]

        if any(p in nome_ghe.upper() for p in _INVALIDOS_GHE):
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
                texto_r = (
                    risco.get("nome_agente", "") + " " + risco.get("perigo_especifico", "")
                ).upper()

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
                    any(g in nome_ghe.upper() for g in _GHES_ADMIN)
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
        cabecalho = st.session_state.get("pcmso_cabecalho", {})

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
            f"ATENCAO: {cargo_val} <small style='color:#c0392b'>(Cargo nao identificado)</small>"
            if pendente else cargo_val
        )
        linhas_html += (
            f"<tr style='{bg}'>"
            f"<td><strong>{row['GHE / Setor']}</strong></td>"
            f"<td>{cargo_disp}</td>"
            f"<td>{row['Exame Clinico/Complementar']}</td>"
            f"<td>{row['Periodicidade']}</td>"
            f"<td>{row['Justificativa Legal / Risco']}</td></tr>"
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

"""
modules/modulo_pcmso.py — v4.7
Fixes:
  [1] _ghe_valido() — normaliza texto ANTES das validações
  [2] _INVALIDOS_GHE_REGEX — barras corrigidas, re.IGNORECASE em todas as buscas
  [3] Novos padrões de rejeição GHE adicionados
  [4] join/split usa \n real
  [5] extrair_pgr_matriz_aiha() — suporte a PGRs no formato Matriz AIHA (Ricco/Hetrin)
  [6] _detectar_formato_pgr() — detecção automática de formato (ghe / aiha / misto)
  [7] extrair_pgr_com_fallback() — roteamento automático por formato
"""

import io
import re
import unicodedata
from datetime import datetime

import pdfplumber
import pandas as pd

from data.matriz_exames import MATRIZ_FUNCAO_EXAME, MATRIZ_RISCO_EXAME
from utils.cargo_utils import normalizar_cargo, normalizar_texto, MAPA_CARGOS_CONHECIDOS, PALAVRAS_EXCLUIR_CARGO
from utils.exame_utils import adicionar_exame_dedup
from utils.biologico_utils import tem_risco_biologico_real, CHAVES_BIOLOGICAS_MATRIZ

# ── Constantes ────────────────────────────────────────────────────────────────

_INVALIDOS_GHE = [
    "QUANTIDADE", "PREVISTOS", "EXPOSTOS", "TOTAL DE", "NUMERO DE",
    "FUNCIONARIOS", "TRABALHADORES", "MEDIDAS DE CONTROLE",
    "FONTE GERADORA", "TRAJETORIA", "DESCRICAO",
    "ATIVIDADES EXERCIDAS", "INFORMACOES SOBRE",
    "PAGINA DE REVISAO", "IDENTIFICACAO DA EMPRESA",
    "COMUNICAR", "DESEMPENHA ATIVIDADES", "UTILIZAM-SE",
    "DIRETORES DA", "DURANTE O DESENVOLVIMENTO",
    "OCULOS DE", "NIVEIS BAIXOS", "IMPORTANCIA",
    "PERMANENTE ELEVADISSIMA", "INTERMITENTE",
    "ATIVIDADES DE -", "ATIVIDADES, UTILIZAM",
    "ATIVIDADES PERMANENTE", "DESENVOLV",
]

_INVALIDOS_GHE_REGEX = [
    r"neste\s+ghe",
    r"expostos\s+neste",
    r"quantidade\s+de\s+func",
    r"^-\s+\w",
    r"comunicar",
    r"desempenha",
    r"utilizam.se",
    r"diretores\s+da",
    r"durante\s+o\s+desenvolv",
    r"oculos\s+de",
    r"niveis\s+baixos",
    r"permanente\s+elevad",
    r"intermitente\s+niveis",
    r"em\s+fun.ao\s+das",
    r"atividades\s+de\s+-",
    r"atividades\s+desempenh",
    r"riscos\s+ocupacionais",
    r"altura,\s+em\s+fun",
    r"para\s+execu",
    r"que\s+executam",
    r"os\s+trabalhadores",
    r"expostos\s+a",
    r"conforme\s+",
    r"verificar\s+",
    r"realizar\s+",
    r"responsav",
    r"^\w\)\s+",
    r"departamento de seguranca",
    r"quantitativa",
    r"para verifica",
    r"avalia.ao quantitativa",
    r"confirma.ao da categoria",
    r"monitoramento peri",
    r"medidas de controle",
    r"grau\s+\d",
    r"avaliacao quantitativa do setor",
    r"iniciar processo",
    r"confirmacao da categoria",
    r"monitoramento periodico",
]

_PALAVRAS_CANTEIRO = [
    "OBRA", "CANTEIRO", "CONSTRUCAO", "REFORMA", "HOSPITAL", "RESIDENCIAL",
    "EDIFICIO", "BLOCO", "TORRE", "HETRIN", "VIADUTO", "PONTE", "SHOPPING",
    "CONDOMINIO", "EMPREENDIMENTO", "MONTAGEM", "INSTALACAO", "CAMPO",
]

_PALAVRAS_ESCRITORIO = [
    "ESCRITORIO", "SEDE", "CORPORATIVO", "ADMINISTRACAO", "MARKETING",
    "TECNOLOGIA DA INFORMACAO", "RECURSOS HUMANOS", "FINANCEIRO",
    "CONTABILIDADE", "JURIDICO", "COMERCIAL",
]

_RISCOS_CANTEIRO = [
    "RUIDO", "VIBRACAO", "POEIRA", "CIMENTO", "SILICA", "TINTA",
    "SOLDA", "ALTURA", "CONFINADO", "MAQUINA", "INCENDIO",
]

_PACOTE_CANTEIRO = [
    {"exame": "Audiometria Tonal (PTA)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": True, "obs": "Canteiro de obras"},
    {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    {"exame": "Eletrocardiograma (ECG)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    {"exame": "Glicemia de Jejum", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    {"exame": "Hemograma Completo", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    {"exame": "Espirometria", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": True, "obs": "Exposicao a poeiras/cimento"},
    {"exame": "Raio-X de Torax OIT", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": True, "obs": "Exposicao a poeiras/cimento"},
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
    "TOLUENO": "Tolueno", "XILENO": "Xileno", "BENZENO": "Benzeno",
    "ACETONA": "Acetona", "THINNER": "Solventes (Thinner)",
    "SOLVENTE": "Solventes Organicos", "TINTA": "Tinta / Verniz",
    "VERNIZ": "Tinta / Verniz", "PRIMER": "Primer (Solventes)",
    "GRAXA": "Graxas / Lubrificantes", "DIESEL": "Diesel / Combustivel",
    "QUEROSENE": "Querosene", "ACIDO": "Acidos (geral)",
    "CIMENTO": "Cimento Portland", "SILICA": "Silica Cristalina",
    "POEIRA": "Poeiras Minerais", "AMIANTO": "Asbesto / Amianto",
    "CHUMBO": "Chumbo", "RUIDO": "Ruido",
    "VIBRACAO": "Vibracao", "CALOR": "Calor (IBUTG)",
    "RADIOATIVO": "Radiacao Ionizante", "RAIOS X IONIZANTE": "Radiacao Ionizante",
    "ERGONO": "Fator Ergonomico", "POSTURA": "Postura Inadequada",
    "LEVANTAMENTO": "Levantamento de Carga", "REPETITIVO": "Movimento Repetitivo",
    "ELETRICO": "Risco Eletrico", "ALTURA": "Queda de Altura",
    "CONFINADO": "Espaco Confinado", "MAQUINA": "Maquinas e Equipamentos",
    "INCENDIO": "Incendio / Explosao",
}

_RE_GHE = re.compile(
    r"(?:GHE[\s:\.\-]+\d|GHE\s+\d+\s*[\-–]|GRUPO\s+HOMOGENEO|LOCAL\s+DE\s+TRABALHO\s*:\s*\w|SETOR\s*:\s*\w)",
    re.IGNORECASE,
)

_RE_TIPO_RISCO = re.compile(r"^[FQBEA]$")

_RE_CABECALHO_AIHA = re.compile(
    r"matriz de risco aiha|tipo de risco|identificacao de perigo|codigo e.?social|"
    r"avaliacao de risco|meio de propagacao|nivel de risco|"
    r"pouca importancia|probabilidade|efeito",
    re.IGNORECASE,
)

_RE_DESCRICAO_FUNCAO = re.compile(
    r"supervisiona|elabora documentacao|controla recursos|cronograma da obra|"
    r"executa atividades|responsavel por|realiza tarefas|desenvolve|presta servicos",
    re.IGNORECASE,
)

_MAPA_TIPO_RISCO = {
    "F": "Fisico",
    "Q": "Quimico",
    "B": "Biologico",
    "E": "Ergonomico",
    "A": "Acidente",
}

_PALAVRAS_CARGO_AIHA = [
    "ENCARREGADO", "PEDREIRO", "ELETRICISTA", "CARPINTEIRO", "SOLDADOR",
    "SERVENTE", "MOTORISTA", "ENGENHEIRO", "TECNICO", "MESTRE", "OPERADOR",
    "ADMINISTRATIVO", "ASSISTENTE", "AUXILIAR", "COMPRADOR", "SUPERVISOR",
    "PINTOR", "ARMADOR", "MONTADOR", "INSTALADOR", "ENCANADOR", "BOMBEIRO",
    "SERRALHEIRO", "TOPOGRAFO", "ALMOXARIFE", "VIGIA", "PORTEIRO", "ZELADOR",
    "MENOR", "APRENDIZ", "ESTAGIARIO", "COORDENADOR", "GERENTE", "DIRETOR",
    "SERVICOS GERAIS", "FISCAL", "INSPETOR", "PROJETISTA", "DESENHISTA",
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def _limpar_nome_ghe(nome: str) -> str:
    if len(nome) > 100:
        return nome[:100].strip() + "..."
    norm = normalizar_texto(nome)
    for lixo in _LIXO_GHE:
        if re.search(lixo, norm, re.IGNORECASE):
            return "GHE (revisar nome)"
    return nome.strip()


def _is_linha_ghe(linha: str) -> bool:
    lu = normalizar_texto(linha.strip())
    for pat in _INVALIDOS_GHE_REGEX:
        if re.search(pat, lu, re.IGNORECASE):
            return False
    if re.match(r"^GHE\s+\d+", linha.strip(), re.IGNORECASE):
        return True
    if _RE_GHE.search(linha):
        return True
    if len(linha.strip()) <= 50 and "/" not in linha and "," not in linha:
        if "DEPARTAMENTO" in lu:
            return True
    return False


def _ghe_valido(nome_ghe: str) -> bool:
    norm = normalizar_texto(nome_ghe)
    if len(nome_ghe.strip()) > 60:
        return False
    if len(norm.strip()) < 4:
        return False
    if any(re.search(pat, norm, re.IGNORECASE) for pat in _INVALIDOS_GHE_REGEX):
        return False
    return not any(inv in norm for inv in _INVALIDOS_GHE)


def _fallback_necessario(ghes: list) -> bool:
    for g in ghes:
        if len(normalizar_texto(g["ghe"])) <= 60 and g["cargos"]:
            return False
    return True


def _fmt_per(per) -> str:
    if per is None or per is False:
        return "-"
    per = str(per).strip().upper().replace("MESES", "").strip()
    if not per or per in ("TRUE", "FALSE", "NONE", ""):
        return "-"
    try:
        return f"{int(per)}M"
    except ValueError:
        return per if per else "-"


def _flag(val) -> str:
    if isinstance(val, bool):
        return "X" if val else "-"
    return "X" if str(val).strip().upper() in ("X", "TRUE", "1", "SIM") else "-"


def _ghe_e_canteiro_misto(nome_ghe: str, riscos: list) -> bool:
    norm = normalizar_texto(nome_ghe)
    if any(p in norm for p in _PALAVRAS_CANTEIRO):
        return True
    if any(p in norm for p in _PALAVRAS_ESCRITORIO):
        return False
    texto_r = " ".join(
        normalizar_texto(r.get("nome_agente", "") + " " + r.get("perigo_especifico", ""))
        for r in riscos
    )
    return any(rc in texto_r for rc in _RISCOS_CANTEIRO)


def _is_nome_funcao_aiha(linha: str) -> bool:
    lstrip = linha.strip()
    lu = normalizar_texto(lstrip)
    if not lstrip or len(lstrip) > 60:
        return False
    if _RE_CABECALHO_AIHA.search(lu):
        return False
    if _RE_DESCRICAO_FUNCAO.search(lu):
        return False
    if _RE_TIPO_RISCO.match(lstrip):
        return False
    if lstrip.startswith("-"):
        return False
    if re.match(r"^\d{2}\.\d{2}\.\d{3}$", lstrip):
        return False
    palavras = lstrip.split()
    if len(palavras) < 2:
        return False
    return any(p in lu for p in _PALAVRAS_CARGO_AIHA)


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

        if (
           _is_linha_ghe(lc)
    and len(lc) < 120          # era 80, aumentar para pegar "GHE 01 - BETONEIRA CMO - RESIDENCIAL..."
    and len(lc.strip()) >= 4
    and not lc.strip().endswith(".")
        ):
            if ghe_atual and (ghe_atual["cargos"] or ghe_atual["riscos_mapeados"]):
                ghes.append(ghe_atual)
            ghe_atual = {"ghe": lc, "cargos": [], "riscos_mapeados": []}
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

    ghes = _deduplicar_ghes(ghes)
    return ghes


def extrair_pgr_matriz_aiha(texto: str) -> list:
    linhas = texto.split("\n")
    ghes = []
    funcao_atual = None
    tipo_risco_atual = None
    agentes_set = set()

    i = 0
    while i < len(linhas):
        lc = linhas[i].strip()
        lu = normalizar_texto(lc)

        if not lc or _RE_CABECALHO_AIHA.search(lu):
            i += 1
            continue

        if _RE_TIPO_RISCO.match(lc):
            tipo_risco_atual = lc.strip()
            i += 1
            continue

        if lc.startswith("-") and funcao_atual is not None and tipo_risco_atual:
            agente_texto = lc.lstrip("- ").split("(")[0].split("\u2013")[0].split("–")[0].strip()
            agente_texto = agente_texto[:120]
            agente_norm = normalizar_texto(agente_texto)
            agente_mapeado = None
            for palavra, agente in _MAPA_AGENTES.items():
                if palavra in agente_norm:
                    agente_mapeado = agente
                    break
            if not agente_mapeado:
                agente_mapeado = agente_texto[:80]

            if agente_mapeado not in agentes_set:
                agentes_set.add(agente_mapeado)
                funcao_atual["riscos_mapeados"].append({
                    "nome_agente": agente_mapeado,
                    "perigo_especifico": lc[:200],
                    "tipo_risco": _MAPA_TIPO_RISCO.get(tipo_risco_atual, tipo_risco_atual),
                })
            i += 1
            continue

        if _is_nome_funcao_aiha(lc):
            nome_completo = lc
            if i + 1 < len(linhas):
                proxima = linhas[i + 1].strip()
                if (
                    proxima
                    and len(proxima) <= 40
                    and not _RE_CABECALHO_AIHA.search(normalizar_texto(proxima))
                    and not _RE_TIPO_RISCO.match(proxima)
                    and not proxima.startswith("-")
                    and not re.match(r"^\d{2}\.\d{2}\.\d{3}$", proxima)
                    and not _RE_DESCRICAO_FUNCAO.search(normalizar_texto(proxima))
                ):
                    nome_completo = f"{lc} {proxima}"
                    i += 1

            if funcao_atual and (funcao_atual["cargos"] or funcao_atual["riscos_mapeados"]):
                ghes.append(funcao_atual)

            funcao_atual = {
                "ghe": nome_completo,
                "cargos": [nome_completo],
                "riscos_mapeados": [],
            }
            agentes_set = set()
            tipo_risco_atual = None
            i += 1
            continue

        i += 1

    if funcao_atual and (funcao_atual["cargos"] or funcao_atual["riscos_mapeados"]):
        ghes.append(funcao_atual)

    return ghes


def _detectar_formato_pgr(texto: str) -> str:
    norm = normalizar_texto(texto)
    tem_aiha = "MATRIZ DE RISCO AIHA" in norm
    tem_ghe = bool(re.search(r"GHE\s*[\d:\-]", texto, re.IGNORECASE))
    if tem_aiha and tem_ghe:
        return "misto"
    if tem_aiha:
        return "aiha"
    return "ghe"


def _deduplicar_ghes(ghes: list) -> list:
    vistos: dict = {}
    resultado = []

    for ghe in ghes:
        chave = frozenset(ghe.get("cargos", []))
        if not chave:
            resultado.append(ghe)
            continue
        if chave not in vistos:
            vistos[chave] = len(resultado)
            resultado.append(ghe)
        else:
            idx = vistos[chave]
            if len(ghe["ghe"]) < len(resultado[idx]["ghe"]):
                resultado[idx]["ghe"] = ghe["ghe"]
            riscos_existentes = {r["nome_agente"] for r in resultado[idx]["riscos_mapeados"]}
            for r in ghe.get("riscos_mapeados", []):
                if r["nome_agente"] not in riscos_existentes:
                    resultado[idx]["riscos_mapeados"].append(r)
                    riscos_existentes.add(r["nome_agente"])

    return resultado


def extrair_pgr_com_fallback(texto_pgr: str, chave_api: str = None) -> tuple:
    formato = _detectar_formato_pgr(texto_pgr)

    if formato == "aiha":
        resultado = extrair_pgr_matriz_aiha(texto_pgr)
        return (resultado, "aiha") if resultado else ([], "parcial")

    if formato == "misto":
        local = extrair_pgr_local(texto_pgr)
        aiha = extrair_pgr_matriz_aiha(texto_pgr)
        nomes_local = {x["ghe"] for x in local}
        merged = local + [g for g in aiha if g["ghe"] not in nomes_local]
        return (merged, "misto") if merged else ([], "parcial")

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
    linhas = []

    for ghe in dados_pgr:
        nome_ghe_raw = ghe.get("ghe", "Sem GHE")
        nome_ghe = _limpar_nome_ghe(str(nome_ghe_raw))
        cargos = ghe.get("cargos") or []
        riscos = ghe.get("riscos_mapeados") or []
        cargos = cargos[:15]
        riscos = riscos[:10]

        if not _ghe_valido(nome_ghe):
            continue

        if tipo_ambiente == "canteiro":
            e_canteiro = True
        elif tipo_ambiente == "escritorio":
            e_canteiro = False
        else:
            e_canteiro = _ghe_e_canteiro_misto(nome_ghe, riscos)

        for cargo in cargos:
            exames: dict = {}

            adicionar_exame_dedup(exames, {
                "exame": "Exame Clinico (Anamnese / Exame Fisico)",
                "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": True,
                "motivo": "NR-07 Basico",
            })

            cargo_norm = normalizar_cargo(cargo)
            cargo_upper = cargo_norm.upper()

            if e_canteiro:
                for ex in _PACOTE_CANTEIRO:
                    adicionar_exame_dedup(exames, {
                        "exame": ex["exame"],
                        "adm": ex["adm"],
                        "per": ex["per"],
                        "mro": ex["mro"],
                        "rt": ex["rt"],
                        "dem": ex["dem"],
                        "motivo": f"Canteiro de Obras — {ex['obs']}" if ex.get("obs") else "Canteiro de Obras — padrao RICCO/HETRIN",
                    })
            else:
                for funcao, lista_ex in MATRIZ_FUNCAO_EXAME.items():
                    if normalizar_texto(funcao) in cargo_upper:
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
                            "exame": regra["exame"],
                            "adm": regra.get("adm", True),
                            "per": regra.get("periodico", "12 MESES"),
                            "mro": regra.get("mro", True),
                            "rt": regra.get("rt", False),
                            "dem": regra.get("dem", False),
                            "motivo": f"Exposicao: {chave_r.title()} — {regra.get('obs', '')}",
                        })

            for ex_info in exames.values():
                per_fmt = _fmt_per(ex_info.get("per", "12 MESES"))
                linhas.append({
                    "GHE / Setor": nome_ghe,
                    "Cargo": cargo,
                    "Exame": ex_info["exame"],
                    "ADM": _flag(ex_info.get("adm", True)),
                    "PER": per_fmt,
                    "MRO": _flag(ex_info.get("mro", True)),
                    "RT": _flag(ex_info.get("rt", False)),
                    "DEM": _flag(ex_info.get("dem", False)),
                    "Justificativa": ex_info.get("motivo", ""),
                })

    return pd.DataFrame(linhas)


# ── Geração HTML ──────────────────────────────────────────────────────────────

def gerar_html_pcmso(df: pd.DataFrame, cabecalho: dict = None) -> str:
    if not cabecalho:
        cabecalho = {}
    razao = cabecalho.get("razao_social", "Empresa nao informada")
    cnpj = cabecalho.get("cnpj", "---")
    obra = cabecalho.get("obra", "---")
    medico = cabecalho.get("medico_rt", "Nao informado")
    vig_i = cabecalho.get("vig_ini", "---")
    vig_f = cabecalho.get("vig_fim", "---")
    tec = cabecalho.get("responsavel_tec", "---")

    ghe_grupos = {}
    for _, row in df.iterrows():
        g = row["GHE / Setor"]
        c = row["Cargo"]
        ghe_grupos.setdefault(g, {}).setdefault(c, []).append(row)

    linhas_html = ""
    for ghe_nome, cargos_dict in ghe_grupos.items():
        total_rows = sum(len(v) for v in cargos_dict.values())
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
  Gerado por Sistema Automacao SST Seconci-GO | NR-07 (Port.1.031/2018), NR-09, NR-15, NR-35, Dec.3.048/99.
</p></body></html>"""


# ── Geração DOCX ──────────────────────────────────────────────────────────────


def gerar_docx_rq61(df: pd.DataFrame, cabecalho: dict = None) -> bytes:
    """
    Gera DOCX no formato RQ.61 Seconci-GO:
    - Cabeçalho com empresa, obra, médico, data
    - Por GHE: linha verde com nome do GHE (colspan)
    - Tabela: FUNÇÃO | EXAMES SOLICITADOS
    - Exames formatados como "Nome (ADM, PER X meses, MRO, DEM)"
    """
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    if not cabecalho:
        cabecalho = {}
    razao  = cabecalho.get("razao_social", "Empresa não informada")
    cnpj   = cabecalho.get("cnpj", "---")
    obra   = cabecalho.get("obra", "---")
    medico = cabecalho.get("medico_rt", "Não informado")
    crm    = cabecalho.get("crm", "")
    vig_i  = cabecalho.get("vig_ini", "---")
    vig_f  = cabecalho.get("vig_fim", "---")
    tipo   = cabecalho.get("tipo_obra", "Renovação")

    VERDE_ESC = "084D22"
    VERDE_MED = "1AA04B"
    BRANCO    = RGBColor(0xFF, 0xFF, 0xFF)
    PRETO     = RGBColor(0x00, 0x00, 0x00)

    def shd(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        s = OxmlElement("w:shd")
        s.set(qn("w:val"), "clear")
        s.set(qn("w:color"), "auto")
        s.set(qn("w:fill"), hex_color)
        tcPr.append(s)

    def set_borders(cell, color="084D22"):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right"):
            border = OxmlElement(f"w:{side}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "4")
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), color)
            tcBorders.append(border)
        tcPr.append(tcBorders)

    def txt(cell, text, bold=False, color=None, size=9,
            align=WD_ALIGN_PARAGRAPH.LEFT, italic=False):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = align
        r = p.add_run(str(text))
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color

    def _fmt_exame_rq61(row) -> str:
        """Converte as colunas ADM/PER/MRO/RT/DEM para formato '(ADM, PER X meses, MRO, DEM)'"""
        partes = []
        if str(row.get("ADM", "-")) == "X":
            partes.append("ADM")
        per = str(row.get("PER", "-")).strip().upper()
        if per and per != "-":
            # normalizar: "12M" -> "PER 12 meses", "6M" -> "PER 6 meses"
            per_num = per.replace("M", "").strip()
            try:
                n = int(per_num)
                partes.append(f"PER {n} meses" if n != 12 else "PER")
            except:
                partes.append(f"PER {per}")
        if str(row.get("MRO", "-")) == "X":
            partes.append("MRO")
        if str(row.get("RT", "-")) == "X":
            partes.append("RET")
        if str(row.get("DEM", "-")) == "X":
            partes.append("DEM")
        sufixo = f" ({', '.join(partes)})" if partes else ""
        return str(row["Exame"]) + sufixo

    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin   = Cm(2.0)
        sec.right_margin  = Cm(1.5)

    # ── Cabeçalho institucional ────────────────────────────────────────────
    cab = doc.add_table(rows=4, cols=4)
    cab.style = "Table Grid"

    # Linha 0: título
    cab.rows[0].cells[0].merge(cab.rows[0].cells[3])
    shd(cab.rows[0].cells[0], VERDE_ESC)
    txt(cab.rows[0].cells[0],
        "MATRIZ FUNÇÃO – EXAMES PCMSO",
        bold=True, color=BRANCO, size=13,
        align=WD_ALIGN_PARAGRAPH.CENTER)

    # Linha 1: empresa
    cab.rows[1].cells[0].merge(cab.rows[1].cells[1])
    cab.rows[1].cells[2].merge(cab.rows[1].cells[3])
    txt(cab.rows[1].cells[0], f"Empresa: {razao}", bold=True, size=9)
    adendo_txt = f"Obra Nova (   )   {tipo} ( X )" if tipo.lower() == "renovação" else f"Obra Nova ( X )   Renovação (   )"
    txt(cab.rows[1].cells[2], adendo_txt, size=9)

    # Linha 2: obra + data
    cab.rows[2].cells[0].merge(cab.rows[2].cells[1])
    cab.rows[2].cells[2].merge(cab.rows[2].cells[3])
    txt(cab.rows[2].cells[0], f"Obra: {obra}", bold=True, size=9)
    txt(cab.rows[2].cells[2], f"Data: {datetime.now().strftime('%d/%m/%Y')}", bold=True, size=9)

    # Linha 3: médico
    cab.rows[3].cells[0].merge(cab.rows[3].cells[3])
    crm_txt = f"  CRM-GO {crm}" if crm else ""
    txt(cab.rows[3].cells[0],
        f"Médico(a) Coordenador(a) do PCMSO: {medico}{crm_txt}",
        size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    # ── Agrupar por GHE → Cargo → lista de exames ────────────────────────
    ghe_grupos = {}
    for _, row in df.iterrows():
        g = row["GHE / Setor"]
        c = row["Cargo"]
        ghe_grupos.setdefault(g, {}).setdefault(c, []).append(row)

    for ghe_nome, cargos_dict in ghe_grupos.items():
        # Tabela por GHE: 2 colunas (FUNÇÃO | EXAMES SOLICITADOS)
        tbl = doc.add_table(rows=0, cols=2)
        tbl.style = "Table Grid"
        tbl.columns[0].width = Cm(5.5)
        tbl.columns[1].width = Cm(12.0)

        # Linha cabeçalho do GHE (verde escuro, largura total)
        row_ghe = tbl.add_row()
        row_ghe.cells[0].merge(row_ghe.cells[1])
        shd(row_ghe.cells[0], VERDE_ESC)
        set_borders(row_ghe.cells[0])
        txt(row_ghe.cells[0], ghe_nome.upper(),
            bold=True, color=BRANCO, size=10,
            align=WD_ALIGN_PARAGRAPH.CENTER)

        # Linha cabeçalho da tabela (verde médio)
        row_h = tbl.add_row()
        shd(row_h.cells[0], VERDE_MED)
        shd(row_h.cells[1], VERDE_MED)
        txt(row_h.cells[0], "FUNÇÃO",          bold=True, color=BRANCO, size=9)
        txt(row_h.cells[1], "EXAMES SOLICITADOS", bold=True, color=BRANCO, size=9)

        for cargo, rows_cargo in cargos_dict.items():
            exames_fmt = [_fmt_exame_rq61(r) for r in rows_cargo]
            # primeira linha do cargo: célula FUNÇÃO com rowspan simulado
            primeira = True
            for exame_str in exames_fmt:
                row_ex = tbl.add_row()
                if primeira:
                    txt(row_ex.cells[0], cargo, bold=True, size=9)
                    primeira = False
                else:
                    row_ex.cells[0].text = ""
                set_borders(row_ex.cells[0])
                set_borders(row_ex.cells[1])
                txt(row_ex.cells[1], exame_str, size=9)

        doc.add_paragraph()

    # Rodapé
    p = doc.add_paragraph(
        f"Responsável pelo preenchimento: {cabecalho.get('responsavel_tec', '---')}\n"
        f"Médico(a) Responsável pela validação: {medico}{(' CRM-GO ' + crm) if crm else ''}\n"
        f"Data do PCMAT/PGR: {vig_i}"
    )
    p.runs[0].font.size = Pt(8)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


"""
Automacao SST - Seconci GO
app.py v5.17 — feat: integra parser_pgr v2 (suporte a GHE e CARGO/CBO)
               badge de formato detectado no Passo 1
"""
import json
import os
import traceback
from datetime import date

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from config.db import (
    get_supabase,
    salvar_historico,
    carregar_historico,
    carregar_html_historico,
)
from modules.modulo_pcmso import (
    extrair_texto_pdf,
    extrair_pgr_com_fallback,
    enriquecer_pgr_com_fispq,
    processar_pcmso,
    gerar_html_pcmso,
    gerar_docx_rq61,
)

st.set_page_config(
    page_title="Automacao SST - Seconci GO",
    layout="wide",
    page_icon=":shield:",
)

# ── CSS CUSTOMIZADO ────────────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .block-container{padding-top:1.4rem;padding-bottom:2rem;max-width:1400px;}
  #MainMenu {visibility: hidden;}
  footer    {visibility: hidden;}
  header    {visibility: hidden;}
  [data-testid="stAppViewContainer"]{background:#F0F2F5;}
  [data-testid="stSidebar"]{background:#F7F9FB;border-right:1px solid #E4E8EE;}
  [data-testid="stSidebar"] *{color:#1A1D23;}
  [data-testid="stFileUploadDropzone"]{border:2px dashed #1AA04B;border-radius:16px;background:#FBFFFC;padding:1rem;}
  .stButton>button{background:linear-gradient(135deg,#084D22,#0E6B31);color:white;border-radius:10px;border:none;box-shadow:0 8px 18px rgba(8,77,34,.16);transition:all .2s ease;font-weight:700;padding:.62rem 1rem;}
  .stButton>button:hover{background:linear-gradient(135deg,#0E6B31,#1AA04B);transform:translateY(-1px);}
  .stDownloadButton>button{border-radius:10px;font-weight:700;}
  h1,h2,h3{color:#084D22!important;letter-spacing:-0.02em;}
  .stAlert{border-radius:12px;border:1px solid #E6EBF1;box-shadow:0 4px 10px rgba(15,23,42,.04);}
  .kaiju-card{background:#FFFFFF;border:1px solid #E6EBF1;border-radius:18px;padding:1.2rem 1.2rem;box-shadow:0 10px 30px rgba(15,23,42,.06);margin-bottom:1rem;}
  .kaiju-card-title{font-size:.92rem;font-weight:700;color:#5B6472;text-transform:uppercase;letter-spacing:.04em;margin-bottom:.35rem;}
  .kaiju-hero{background:linear-gradient(135deg,#084D22 0%,#1AA04B 100%);border-radius:22px;padding:1.3rem 1.4rem;color:white;box-shadow:0 18px 40px rgba(8,77,34,.20);margin-bottom:1rem;}
  .kaiju-hero h2{color:white!important;margin:0 0 .25rem 0;}
  .kaiju-hero p{margin:0;opacity:.96;font-size:0.98rem;}
  div[data-testid="metric-container"] {
    background-color: #ffffff;
    border: 1px solid #e0e0e0;
    padding: 15px;
    border-radius: 8px;
    box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
    border-left: 5px solid #084D22;
  }
  div[data-testid="metric-container"] label{font-weight:700;color:#5B6472;}
  div[data-testid="stDataFrame"] {
    border-radius: 8px;
    box-shadow: 2px 2px 10px rgba(0,0,0,0.05);
  }
  .audit-panel{background:#FFFFFF;border:1px solid #E6EBF1;border-radius:18px;padding:1rem 1.1rem;box-shadow:0 10px 28px rgba(15,23,42,.05);}
  .stTabs [data-baseweb="tab-list"]{gap:.5rem;}
  .stTabs [data-baseweb="tab"]{background:#EAF3ED;border-radius:10px;padding:.55rem .95rem;font-weight:700;}
  .stTabs [aria-selected="true"]{background:#084D22!important;color:white!important;}
</style>
""", unsafe_allow_html=True)


def card_inicio(titulo: str, texto: str):
    st.markdown(
        f"""
        <div class="kaiju-card">
            <div class="kaiju-card-title">{titulo}</div>
            <div>{texto}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_auditoria_metrics(resultado_auditoria: dict):
    total = int(resultado_auditoria.get("total_divergencias", 0))
    faltando = len(resultado_auditoria.get("exames_faltando", []))
    excedente = len(resultado_auditoria.get("exames_excedentes", []))
    cargo_faltando = len(resultado_auditoria.get("cargo_faltando", []))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Divergencias", total)
    c2.metric("Exames faltando", faltando)
    c3.metric("Exames excedentes", excedente)
    c4.metric("Cargos sem match", cargo_faltando)


# ── Autenticacao ──────────────────────────────────────────────────────────────────────────────
def check_password():
    def validar_login():
        usr_digitado = st.session_state["username_input"]
        pwd_digitada = st.session_state["password_input"]
        usr_correto = st.secrets.get("USUARIO_SISTEMA", "diovanni")
        pwd_correta = st.secrets.get("SENHA_SISTEMA", "seconci123")
        if usr_digitado == usr_correto and pwd_digitada == pwd_correta:
            st.session_state["autenticado"] = True
            del st.session_state["password_input"]
            del st.session_state["username_input"]
        else:
            st.session_state["autenticado"] = False

    if st.session_state.get("autenticado", False):
        return True

    st.markdown("""
        <h2 style='text-align:center;color:#084D22;margin-top:80px;'>
            🔒 Acesso Restrito — Seconci GO
        </h2>
        <p style='text-align:center;color:#6C757D;margin-bottom:30px;'>
            Sistema de Automação SST
        </p>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("""
            <div style='border:1px solid #DDE6E0;padding:30px;
                        border-radius:18px;background-color:#FFFFFF;
                        box-shadow:0 10px 24px rgba(8,77,34,.08);'>
        """, unsafe_allow_html=True)
        st.text_input("Usuário", key="username_input")
        st.text_input("Senha", type="password", key="password_input")
        st.button("Entrar no Sistema", on_click=validar_login, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
        if "autenticado" in st.session_state and not st.session_state["autenticado"]:
            st.error("Usuário ou senha incorretos.")
    return False


if not check_password():
    st.stop()


# ── Sidebar ──────────────────────────────────────────────────────────────────────────────
for logo in ("logo.png", "logo.jpg"):
    if os.path.exists(logo):
        st.sidebar.image(logo, width=220)
        break
else:
    st.sidebar.markdown(
        "<h2 style='text-align:center;color:#084D22;'>SECONCI-GO</h2>",
        unsafe_allow_html=True,
    )

st.sidebar.markdown("---")
st.sidebar.title("Modulos do Sistema")
modulo = st.sidebar.radio(
    "Selecione a funcionalidade:",
    [
        "Dashboard",
        "Engenharia: FISPQ / FDS - PGR",
        "Medicina: PGR - PCMSO",
        "Construtor Visual de GHEs",
    ],
)

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Sair do Sistema", use_container_width=True):
    st.session_state["autenticado"] = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.title("Historico de Laudos")
historico = carregar_historico()
historico_html = None

if historico:
    opcoes = ["Selecione um projeto salvo..."] + [
        f"{r['id']} - {r['nome_projeto']} ({r['data_salvamento']})"
        for r in historico
    ]
    sel = st.sidebar.selectbox("Carregar projeto:", opcoes)
    if sel != "Selecione um projeto salvo...":
        id_sel = int(sel.split(" - ")[0])
        historico_html = carregar_html_historico(id_sel)
        if historico_html:
            st.sidebar.success("Projeto carregado.")
else:
    st.sidebar.write("Nenhum projeto salvo ainda.")


# ── Banco de matrizes ─────────────────────────────────────────────────────────────────────────────────
banco_path = os.path.join("data", "banco_matrizes_v1_1.json")
banco_matrizes = {}

if os.path.exists(banco_path):
    try:
        with open(banco_path, "r", encoding="utf-8") as f:
            banco_matrizes = json.load(f)
    except Exception as e:
        st.sidebar.warning(f"Banco de matrizes nao carregado: {e}")

_banco_ativo = bool(banco_matrizes)


# ── Roteamento ─────────────────────────────────────────────────────────────────────────────────
if historico_html:
    st.title("Laudo Carregado do Historico")
    components.html(historico_html, height=700, scrolling=True)

elif modulo == "Dashboard":
    st.title("Dashboard - Sistema Integrado SST")
    try:
        sb = get_supabase()
        total     = sb.table("historico_laudos").select("id", count="exact").execute().count or 0
        total_cas = sb.table("dicionario_dinamico").select("cas", count="exact").execute().count or 0
    except Exception:
        total, total_cas = 0, 0
    c1, c2, c3 = st.columns(3)
    c1.metric("Laudos Gerados", total)
    c2.metric("CAS no Banco Dinamico", total_cas)
    c3.metric("Modulos Ativos", 4)
    st.info("Use o menu lateral para acessar os modulos.")

elif modulo == "Engenharia: FISPQ / FDS - PGR":
    try:
        from modules.modulo_engenharia import render_engenharia
        render_engenharia()
    except ImportError:
        st.warning("modulo_engenharia.py nao encontrado.")

elif modulo == "Construtor Visual de GHEs":
    try:
        from modules.modulo_construtor_visual import render_construtor_visual
        render_construtor_visual()
    except Exception as e:
        st.error(f"❌ Erro no Construtor Visual: {type(e).__name__}: {e}")
        st.code(traceback.format_exc(), language="python")

elif modulo == "Medicina: PGR - PCMSO":
    st.markdown("""
        <div class="kaiju-hero">
            <h2>🩺 Módulo Médico — Importador de PGR e Gerador de PCMSO</h2>
            <p>Fluxo guiado em 3 etapas: upload e extração, auditoria clínica e aprovação final.</p>
        </div>
    """, unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs([
        "📂 Passo 1: Upload e Extração",
        "🔍 Passo 2: Conferência de Cargos",
        "✅ Passo 3: Aprovação e Download",
    ])

    with tab1:
        card_inicio(
            "Motor de extração",
            "parser_pgr v2: detecta automaticamente formato GHE ou CARGO/CBO. "
            "Fallback para IA Gemini somente se necessario.",
        )

        fispq_carregados = st.session_state.get("fispq_resultados_medicos", [])
        if fispq_carregados:
            st.success(
                f"🧪 {len(fispq_carregados)} agente(s) da FISPQ em memoria — serão injetados automaticamente no PCMSO apos a extracao."
            )

        if _banco_ativo:
            st.info(
                "📚 Banco de matrizes tecnicas ativo — os exames de cada cargo serao preenchidos "
                "automaticamente com base no padrao validado, **independente do nome do arquivo**."
            )

        with st.container():
            st.markdown("<div class='kaiju-card'>", unsafe_allow_html=True)
            st.markdown("### Dados de Identificacao do PCMSO")
            st.caption("NR-07 item 7.5.19.1")
            col1, col2 = st.columns(2)
            cab = st.session_state.get("pcmso_cabecalho", {})
            with col1:
                razao_social = st.text_input("Razao Social da Empresa *", value=cab.get("razao_social", ""))
                cnpj = st.text_input("CNPJ *", value=cab.get("cnpj", ""))
                medico_rt = st.text_input("Medico Responsavel RT (Nome + CRM) *", value=cab.get("medico_rt", ""))
            with col2:
                vig_ini = st.date_input("Vigencia - Inicio", value=date.today())
                vig_fim = st.date_input("Vigencia - Fim", value=date.today())
                resp_tec = st.text_input("Tecnico SST Responsavel (opcional)", value=cab.get("responsavel_tec", ""))
                obra = st.text_input("Obra / Unidade (opcional)", value=cab.get("obra", ""))
            st.markdown("---")
            st.markdown("**Tipo de Ambiente da Obra** *(define o pacote de exames)*")
            opcoes_amb = {
                "🏗️ Canteiro de Obras / Obra": "canteiro",
                "🏢 Escritório Corporativo": "escritorio",
                "🔀 Misto (Canteiro + Escritório no mesmo PGR)": "misto",
            }
            label_amb = st.radio(
                "Selecione o tipo:",
                list(opcoes_amb.keys()),
                index=1,
                horizontal=True,
            )
            tipo_ambiente = opcoes_amb[label_amb]
            st.markdown("</div>", unsafe_allow_html=True)

        st.session_state["pcmso_cabecalho"] = {
            "razao_social": razao_social,
            "cnpj": cnpj,
            "medico_rt": medico_rt,
            "vig_ini": vig_ini.strftime("%d/%m/%Y"),
            "vig_fim": vig_fim.strftime("%d/%m/%Y"),
            "responsavel_tec": resp_tec,
            "obra": obra,
        }
        st.session_state["tipo_ambiente"] = tipo_ambiente

        st.markdown("<div class='kaiju-card'>", unsafe_allow_html=True)
        pdf_file = st.file_uploader("Arraste o PDF do PGR aqui", type=["pdf"])
        if pdf_file:
            st.session_state["nome_pdf_atual"] = pdf_file.name

        def _marcar_extracao():
            st.session_state["_executar_extracao"] = True

        st.button(
            "Extrair Riscos e Gerar PCMSO",
            use_container_width=True,
            on_click=_marcar_extracao,
        )
        st.markdown("</div>", unsafe_allow_html=True)

        if st.session_state.pop("_executar_extracao", False):
            if not pdf_file:
                st.error("Faca upload do PDF do PGR antes de continuar.")
                st.stop()

            with st.spinner("Extraindo texto do PDF..."):
                texto_pgr = extrair_texto_pdf(pdf_file)
            st.success(f"Texto extraido: {len(texto_pgr):,} caracteres em {pdf_file.name}")

            with st.expander("DEBUG: Primeiras 100 linhas do PDF"):
                for i, linha in enumerate(texto_pgr.split("\n")[:100], 1):
                    st.text(f"{i:3}: {linha}")

            # ───────────────────────────────────────────────────────────────────────────
            # PARSER v2: GHE ou CARGO/CBO automatico
            # ───────────────────────────────────────────────────────────────────────────
            with st.spinner("Identificando GHEs / Cargos e riscos..."):
                _resultado_pgr = None
                try:
                    from parser_pgr import parsear_pgr as _parsear_pgr
                    with open("regras_pcmso.json", encoding="utf-8") as _f:
                        _regras = json.load(_f)
                    # Re-leitura dos bytes (file_uploader pode ter cursor no fim)
                    pdf_file.seek(0)
                    _resultado_pgr = _parsear_pgr(pdf_file.read(), _regras)
                except Exception as _e_parser:
                    st.warning(
                        f"⚠️ parser_pgr nao disponivel ({type(_e_parser).__name__}: {_e_parser}) — "
                        "usando fallback Gemini."
                    )

            if _resultado_pgr:
                _fmt     = _resultado_pgr["formato"]
                _n_sec   = len(_resultado_pgr["ghe_blocos"])
                _metodo  = _resultado_pgr["metodo_extracao"]
                st.success(
                    f"✅ Formato detectado: **{_fmt}** | "
                    f"**{_n_sec}** seções encontradas | "
                    f"Extração: **{_metodo}**"
                )
                if _resultado_pgr.get("aviso"):
                    st.warning(_resultado_pgr["aviso"])

                # Converte ghe_blocos → dados_ghe (contrato do restante do app)
                dados_ghe = {}
                for _nome_sec, _info in _resultado_pgr["ghe_blocos"].items():
                    dados_ghe[_nome_sec] = {
                        "cargo":  _nome_sec,
                        "riscos": _info["riscos_identificados"],
                        "exames": [_e["exame"] for _e in _info["exames_gerados"]],
                    }
                fonte = "local"
            else:
                # Fallback: pipeline antigo
                dados_ghe, fonte = extrair_pgr_com_fallback(texto_pgr)

            # ───────────────────────────────────────────────────────────────────────────
            st.session_state["dados_ghe_processados"] = dados_ghe

            if fonte == "local":
                st.success("Dados extraidos localmente — sem consumo de IA!")
            elif fonte == "ia":
                st.info("Dados extraidos via IA (Gemini).")
            else:
                st.warning("Extracao parcial — revise os resultados.")

            if not dados_ghe:
                st.error("Nenhum GHE / Cargo identificado. Verifique se o PDF e um PGR valido.")
                st.stop()

            resultados_fispq = st.session_state.get("fispq_resultados_medicos", [])
            if resultados_fispq:
                st.success(f"🧪 {len(resultados_fispq)} Agente(s) Químico(s) da FISPQ detectados! Injetando no PGR...")
                dados_ghe = enriquecer_pgr_com_fispq(dados_ghe, resultados_fispq)

            if _banco_ativo:
                from modules.modulo_auditor_v1_1 import enriquecer_ghe_com_banco
                with st.spinner("Aplicando padrao tecnico de exames por cargo..."):
                    dados_ghe, rel_banco = enriquecer_ghe_com_banco(dados_ghe, banco_matrizes)

                n_enr = len(rel_banco['cargos_enriquecidos'])
                n_man = len(rel_banco['cargos_mantidos'])
                if n_enr:
                    st.success(
                        f"✅ {n_enr} cargo(s) preenchidos com padrao tecnico do banco: "
                        f"{', '.join(dict.fromkeys(rel_banco['cargos_enriquecidos']))}"
                    )
                if n_man:
                    st.warning(
                        f"⚠️ {n_man} cargo(s) nao encontrados no banco — exames extraidos do PGR mantidos: "
                        f"{', '.join(dict.fromkeys(rel_banco['cargos_mantidos']))}"
                    )

            tipo_amb = st.session_state.get("tipo_ambiente", "escritorio")
            with st.spinner(f"Gerando matriz PCMSO ({tipo_amb})..."):
                try:
                    df_pcmso = processar_pcmso(dados_ghe, tipo_ambiente=tipo_amb)
                except Exception as e:
                    st.error(f"❌ Erro em processar_pcmso(): {type(e).__name__}: {e}")
                    st.code(traceback.format_exc(), language="python")
                    st.stop()

            if df_pcmso.empty:
                st.warning("PCMSO gerado vazio — nenhum cargo/exame identificado.")
                st.stop()

            st.success(f"PCMSO gerado com {len(df_pcmso)} linhas de exames. Revise no Passo 2.")
            st.session_state["df_pcmso_gerado"] = df_pcmso
            st.session_state["relatorio_banco"] = rel_banco if _banco_ativo else None
            # Limpa XML anterior ao gerar novo PCMSO
            st.session_state.pop("med_xml_s2240_bytes", None)

    with tab2:
        st.markdown("<div class='audit-panel'>", unsafe_allow_html=True)
        st.markdown("### Conferência de Cargos e Exames")

        rel_banco = st.session_state.get("relatorio_banco")
        if rel_banco:
            enr = list(dict.fromkeys(rel_banco.get('cargos_enriquecidos', [])))
            man = list(dict.fromkeys(rel_banco.get('cargos_mantidos', [])))
            c1, c2 = st.columns(2)
            c1.metric("✅ Cargos com padrao tecnico aplicado", len(enr))
            c2.metric("⚠️ Cargos mantidos do PGR", len(man))
            if enr:
                with st.expander("Cargos preenchidos pelo banco de matrizes", expanded=True):
                    for c in enr:
                        st.markdown(f"- ✅ {c}")
            if man:
                with st.expander("Cargos nao encontrados no banco (mantidos do PGR)"):
                    for c in man:
                        st.markdown(f"- ⚠️ {c} — considere adicionar ao banco de matrizes")
        elif "df_pcmso_gerado" in st.session_state:
            st.info("Matriz gerada sem banco de matrizes ativo. Revise os exames manualmente no Passo 3.")
        else:
            st.info("Extraia um PGR no Passo 1 para preencher este painel.")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab3:
        st.markdown("<div class='kaiju-card'>", unsafe_allow_html=True)
        st.markdown("### Aprovacao e Download")
        if "df_pcmso_gerado" in st.session_state:
            st.info(
                "Revise a matriz abaixo. Voce pode corrigir nomes, alterar periodicidades, "
                "marcar ou desmarcar ADM/PER/DEM, excluir linhas e adicionar novos exames."
            )

            df_atual = st.session_state["df_pcmso_gerado"].copy()
            total_linhas = len(df_atual)
            colunas_df = list(df_atual.columns)

            with st.expander("➕ Inserir nova linha em posicao especifica", expanded=False):
                st.caption(f"Matriz tem {total_linhas} linha(s). A nova linha sera inserida ANTES da posicao escolhida (1 = inicio, {total_linhas + 1} = final).")
                col_pos, col_ghe_n, col_cargo_n, col_exame_n = st.columns([1, 2, 2, 3])
                with col_pos:
                    pos_inserir = st.number_input(
                        "Posicao",
                        min_value=1,
                        max_value=total_linhas + 1,
                        value=total_linhas + 1,
                        step=1,
                        key="ins_posicao",
                    )
                with col_ghe_n:
                    ghe_opcoes = list(df_atual["GHE / Setor"].unique()) if "GHE / Setor" in df_atual.columns else []
                    ghe_novo = st.selectbox("GHE / Setor", options=ghe_opcoes + ["-- Novo --"], key="ins_ghe")
                    if ghe_novo == "-- Novo --":
                        ghe_novo = st.text_input("Nome do GHE", key="ins_ghe_custom")
                with col_cargo_n:
                    cargo_opcoes = list(df_atual["Cargo"].unique()) if "Cargo" in df_atual.columns else []
                    cargo_novo = st.selectbox("Cargo", options=cargo_opcoes + ["-- Novo --"], key="ins_cargo")
                    if cargo_novo == "-- Novo --":
                        cargo_novo = st.text_input("Nome do Cargo", key="ins_cargo_custom")
                with col_exame_n:
                    exame_novo = st.text_input("Exame", key="ins_exame")

                col_adm_n, col_per_n, col_mro_n, col_rt_n, col_dem_n, col_just_n = st.columns(6)
                adm_novo  = col_adm_n.selectbox("ADM",  ["-", "X"], key="ins_adm")
                per_novo  = col_per_n.text_input("PER",  value="12M", key="ins_per")
                mro_novo  = col_mro_n.selectbox("MRO",  ["X", "-"], key="ins_mro")
                rt_novo   = col_rt_n.selectbox("RT",   ["-", "X"], key="ins_rt")
                dem_novo  = col_dem_n.selectbox("DEM",  ["-", "X"], key="ins_dem")
                just_novo = col_just_n.text_input("Justificativa", value="Inserido manualmente", key="ins_just")

                if st.button("➕ Confirmar insercao", key="btn_inserir_linha", use_container_width=True):
                    if not exame_novo.strip():
                        st.warning("Informe o nome do exame antes de inserir.")
                    else:
                        nova_linha = {
                            "GHE / Setor": ghe_novo or "",
                            "Cargo": cargo_novo or "",
                            "Exame": exame_novo.strip(),
                            "ADM": adm_novo,
                            "PER": per_novo.upper().strip(),
                            "MRO": mro_novo,
                            "RT": rt_novo,
                            "DEM": dem_novo,
                            "Justificativa": just_novo,
                        }
                        for col in colunas_df:
                            if col not in nova_linha:
                                nova_linha[col] = ""

                        idx = int(pos_inserir) - 1
                        parte_antes = df_atual.iloc[:idx]
                        parte_depois = df_atual.iloc[idx:]
                        df_inserida = pd.concat(
                            [parte_antes, pd.DataFrame([nova_linha]), parte_depois],
                            ignore_index=True,
                        )
                        st.session_state["df_pcmso_gerado"] = df_inserida
                        st.success(f"✅ Linha inserida na posicao {int(pos_inserir)}: {exame_novo} — {cargo_novo}")
                        st.rerun()

            df_editado = st.data_editor(
                st.session_state["df_pcmso_gerado"],
                num_rows="dynamic",
                use_container_width=True,
                key="editor_matriz_pcmso",
                height=500,
            )

            if not df_editado.equals(st.session_state["df_pcmso_gerado"]):
                st.session_state["df_pcmso_gerado"] = df_editado

            st.markdown("---")

            if st.button("✅ Aprovar Matriz e Gerar Documentos", type="primary", use_container_width=True):
                cabecalho_atual = st.session_state["pcmso_cabecalho"]
                razao_social_ap = cabecalho_atual.get("razao_social", "")
                medico_rt_ap    = cabecalho_atual.get("medico_rt", "")
                with st.spinner("Consolidando correcoes e gerando laudos oficiais..."):
                    try:
                        html_pcmso = gerar_html_pcmso(df_editado, cabecalho=cabecalho_atual)
                        bytes_docx = gerar_docx_rq61(df_editado, cabecalho=cabecalho_atual)
                    except Exception as e:
                        st.error(f"❌ Erro na geracao dos documentos: {type(e).__name__}: {e}")
                        st.code(traceback.format_exc(), language="python")
                        st.stop()

                st.session_state["html_pcmso_aprovado"] = html_pcmso
                st.session_state["bytes_docx_aprovado"] = bytes_docx
                st.session_state["nome_arq_aprovado"] = (
                    razao_social_ap.replace(" ", "_")[:30] if razao_social_ap else "PCMSO"
                )

                if razao_social_ap and medico_rt_ap:
                    nome_proj = f"PCMSO - {razao_social_ap[:40]} ({date.today().strftime('%d/%m/%Y')})"
                    salvar_historico(nome_proj, html_pcmso)
                    st.success("✅ Laudo aprovado e salvo no historico com sucesso!")
                else:
                    st.warning("Preencha Razao Social e Medico RT para salvar no historico.")

            if "html_pcmso_aprovado" in st.session_state:
                nome_arq = st.session_state.get("nome_arq_aprovado", "PCMSO")
                st.markdown("### ⬇️ Documentos Prontos para Download")
                col_html, col_docx = st.columns(2)
                with col_html:
                    st.download_button(
                        label="📄 Baixar PCMSO (.html)",
                        data=st.session_state["html_pcmso_aprovado"].encode("utf-8"),
                        file_name=f"PCMSO_{nome_arq}.html",
                        mime="text/html",
                        use_container_width=True,
                    )
                with col_docx:
                    st.download_button(
                        label="📝 Baixar PCMSO (.docx)",
                        data=st.session_state["bytes_docx_aprovado"],
                        file_name=f"PCMSO_{nome_arq}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                with st.expander("👁️ Preview do PCMSO gerado", expanded=False):
                    components.html(st.session_state["html_pcmso_aprovado"], height=600, scrolling=True)

            from modules.modulo_esocial_xml import render_botao_xml
            render_botao_xml(
                df_editado,
                st.session_state.get("pcmso_cabecalho", {}),
                dados_pgr=st.session_state.get("dados_ghe_processados"),
                key_prefix="med",
            )

        else:
            st.info("A matriz aprovada aparecera aqui apos a extracao do Passo 1.")
        st.markdown("</div>", unsafe_allow_html=True)

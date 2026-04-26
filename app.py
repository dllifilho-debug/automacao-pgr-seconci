"""
Automacao SST - Seconci GO
app.py v5.13 — fix: usa flag auxiliar _executar_extracao para evitar StreamlitAPIException
               ao tentar setar session_state["btn_extrair_pcmso"] = False
"""
import json
import os
import traceback
from datetime import date

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

st.markdown("""<style>
  .block-container{padding-top:1.4rem;padding-bottom:2rem;max-width:1400px;}
  #MainMenu, footer, header {visibility:hidden;}
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
  div[data-testid="metric-container"]{background:#FFFFFF;border:1px solid #E6EBF1;padding:1rem;border-radius:16px;box-shadow:0 8px 24px rgba(15,23,42,.05);}
  div[data-testid="metric-container"] label{font-weight:700;color:#5B6472;}
  .audit-panel{background:#FFFFFF;border:1px solid #E6EBF1;border-radius:18px;padding:1rem 1.1rem;box-shadow:0 10px 28px rgba(15,23,42,.05);}
  .stTabs [data-baseweb="tab-list"]{gap:.5rem;}
  .stTabs [data-baseweb="tab"]{background:#EAF3ED;border-radius:10px;padding:.55rem .95rem;font-weight:700;}
  .stTabs [aria-selected="true"]{background:#084D22!important;color:white!important;}
</style>""", unsafe_allow_html=True)


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
    c1.metric("Divergências", total)
    c2.metric("Exames faltando", faltando)
    c3.metric("Exames excedentes", excedente)
    c4.metric("Cargos sem match", cargo_faltando)


# ── Autenticação ──────────────────────────────────────────────────────────────
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


# ── Sidebar ───────────────────────────────────────────────────────────────────
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


# ── Banco de matrizes ───────────────────────────────────────────────────────────
banco_path = os.path.join("data", "banco_matrizes_v1_1.json")
banco_matrizes = {}
bases_disponiveis = []

if os.path.exists(banco_path):
    try:
        with open(banco_path, "r", encoding="utf-8") as f:
            banco_matrizes = json.load(f)
        bases_disponiveis = list(banco_matrizes.get("obras_referencia", {}).keys())
    except Exception as e:
        st.sidebar.warning(f"Banco de matrizes não carregado: {e}")


def selecionar_base_automatica(nome_arquivo: str) -> str | None:
    if not bases_disponiveis:
        return None
    import unicodedata
    def strip_acc(s):
        return ''.join(
            c for c in unicodedata.normalize('NFD', s)
            if unicodedata.category(c) != 'Mn'
        )
    nome_norm = strip_acc(nome_arquivo).upper().replace("-", " ").replace("_", " ")
    for base in bases_disponiveis:
        base_norm = strip_acc(base).upper().replace("_", " ")
        for palavra in base_norm.split():
            if len(palavra) > 4 and palavra in nome_norm:
                return base
    # Fallback silencioso — exibe warning fora desta função para evitar
    # chamadas duplicadas que geram warnings repetidos
    return bases_disponiveis[0]


# ── Roteamento ──────────────────────────────────────────────────────────────────
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
        "🔍 Passo 2: Auditoria Clínica",
        "✅ Passo 3: Aprovação e Download",
    ])

    with tab1:
        card_inicio(
            "Motor de extração",
            "Etapa 1 — Extração Local (gratuita e instantânea). Etapa 2 — IA Gemini (fallback somente se necessário).",
        )

        fispq_carregados = st.session_state.get("fispq_resultados_medicos", [])
        if fispq_carregados:
            st.success(
                f"🧪 {len(fispq_carregados)} agente(s) da FISPQ em memória — serão injetados automaticamente no PCMSO após a extração."
            )

        with st.container():
            st.markdown("<div class='kaiju-card'>", unsafe_allow_html=True)
            st.markdown("### Dados de Identificação do PCMSO")
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
                help=(
                    "Canteiro → todo cargo recebe pacote completo. Escritório → cargo admin recebe Exame Clínico + Acuidade Visual. Misto → sistema detecta por GHE automaticamente."
                ),
            )
            tipo_ambiente = opcoes_amb[label_amb]
            st.markdown("---")
            if bases_disponiveis:
                st.caption(
                    f"🔍 Auditoria automática ativa — base detectada pelo nome do PDF. Bases disponíveis: `{'` · `'.join(bases_disponiveis)}`"
                )
            else:
                st.caption("⚠️ Nenhuma base de auditoria encontrada — `banco_matrizes_v1_1.json` não localizado.")
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

        # FIX: usa on_click para setar flag auxiliar — nunca escreve direto na chave do widget
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

            nome_pdf = st.session_state.get("nome_pdf_atual", "")
            base_sel = selecionar_base_automatica(nome_pdf)

            # Exibe aviso de fallback aqui, fora de selecionar_base_automatica
            palavras_base = set()
            for base in bases_disponiveis:
                import unicodedata
                def strip_acc(s):
                    return ''.join(
                        c for c in unicodedata.normalize('NFD', s)
                        if unicodedata.category(c) != 'Mn'
                    )
                for p in strip_acc(base).upper().replace("_", " ").split():
                    if len(p) > 4:
                        palavras_base.add(p)
            nome_norm_check = ''.join(
                c for c in unicodedata.normalize('NFD', nome_pdf)
                if unicodedata.category(c) != 'Mn'
            ).upper().replace("-", " ").replace("_", " ")
            detectado_exato = any(p in nome_norm_check for p in palavras_base)

            if base_sel and detectado_exato:
                st.info(f"🔍 Base de auditoria detectada automaticamente: **{base_sel}**")
            elif base_sel and not detectado_exato:
                st.warning(
                    f"⚠️ Nenhuma base detectada no nome '{nome_pdf}'. "
                    f"Usando base padrão: **{base_sel}**. "
                    "Renomeie o PDF com o nome da obra para detecção automática."
                )
            else:
                st.caption("ℹ️ Nenhuma base de auditoria disponível — PCMSO gerado sem validação.")

            with st.spinner("Extraindo texto do PDF..."):
                texto_pgr = extrair_texto_pdf(pdf_file)
            st.success(f"Texto extraido: {len(texto_pgr):,} caracteres em {pdf_file.name}")

            with st.expander("DEBUG: Primeiras 100 linhas do PDF"):
                for i, linha in enumerate(texto_pgr.split("\n")[:100], 1):
                    st.text(f"{i:3}: {linha}")

            with st.spinner("Processando GHEs..."):
                dados_ghe, fonte = extrair_pgr_com_fallback(texto_pgr)

            st.session_state["dados_ghe_processados"] = dados_ghe

            if fonte == "local":
                st.success("Dados extraidos localmente — sem consumo de IA!")
            elif fonte == "ia":
                st.info("Dados extraidos via IA (Gemini).")
            else:
                st.warning("Extracao parcial — revise os resultados.")

            if not dados_ghe:
                st.error("Nenhum GHE identificado. Verifique se o PDF e um PGR valido.")
                st.stop()

            resultados_fispq = st.session_state.get("fispq_resultados_medicos", [])
            if resultados_fispq:
                st.success(f"🧪 {len(resultados_fispq)} Agente(s) Químico(s) da FISPQ detectados! Injetando no PGR...")
                dados_ghe_enriquecido = enriquecer_pgr_com_fispq(dados_ghe, resultados_fispq)
            else:
                dados_ghe_enriquecido = dados_ghe

            tipo_amb = st.session_state.get("tipo_ambiente", "escritorio")
            with st.spinner(f"Gerando matriz PCMSO ({tipo_amb})..."):
                try:
                    df_pcmso = processar_pcmso(dados_ghe_enriquecido, tipo_ambiente=tipo_amb)
                except Exception as e:
                    st.error(f"❌ Erro em processar_pcmso(): {type(e).__name__}: {e}")
                    st.code(traceback.format_exc(), language="python")
                    st.stop()

            if df_pcmso.empty:
                st.warning("PCMSO gerado vazio — nenhum cargo/exame identificado.")
                st.stop()

            st.success(f"PCMSO gerado preliminarmente com {len(df_pcmso)} linhas de exames.")
            st.session_state["df_pcmso_gerado"] = df_pcmso
            st.session_state["base_auditoria_atual"] = base_sel
            st.session_state["resultado_auditoria"] = None
            st.session_state["relatorio_auditoria_txt"] = None

            if base_sel and banco_matrizes:
                try:
                    from modules.modulo_auditor_v1_1 import (
                        auditar_pcmso, pcmso_df_para_dict,
                        formatar_relatorio_auditoria, obra_tem_matriz
                    )

                    tem_historico = obra_tem_matriz(banco_matrizes, base_sel)

                    if not tem_historico:
                        st.warning(
                            f"⚠️ Obra **{base_sel}** sem matriz validada no histórico. "
                            "Matriz gerada pela IA — encaminhe para revisão médica antes de emitir."
                        )
                    else:
                        dados_para_auditoria = pcmso_df_para_dict(df_pcmso)
                        resultado_auditoria = auditar_pcmso(dados_para_auditoria, banco_matrizes, obra_id=base_sel)
                        relatorio_txt = formatar_relatorio_auditoria(resultado_auditoria)
                        st.session_state["resultado_auditoria"] = resultado_auditoria
                        st.session_state["relatorio_auditoria_txt"] = relatorio_txt
                except Exception as e:
                    st.error(f"❌ Erro na auditoria: {type(e).__name__}: {e}")
                    st.code(traceback.format_exc(), language="python")

    with tab2:
        st.markdown("<div class='audit-panel'>", unsafe_allow_html=True)
        st.markdown("### Painel de Auditoria Clínica")
        resultado_auditoria = st.session_state.get("resultado_auditoria")
        relatorio_txt = st.session_state.get("relatorio_auditoria_txt")

        if resultado_auditoria:
            if resultado_auditoria.get("ok"):
                st.success("✅ Auditoria concluída — nenhuma divergência encontrada!")
            else:
                st.warning(f"⚠️ {resultado_auditoria.get('total_divergencias', 0)} divergência(s) detectada(s) — revise antes de aprovar.")
            render_auditoria_metrics(resultado_auditoria)
            with st.expander("📋 Relatório da Auditoria", expanded=True):
                st.code(relatorio_txt or "Sem relatório disponível.", language="text")
        elif "df_pcmso_gerado" in st.session_state:
            st.info("Matriz gerada, mas sem auditoria validada no histórico para esta obra.")
        else:
            st.info("Extraia um PGR no Passo 1 para preencher este painel.")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab3:
        st.markdown("<div class='kaiju-card'>", unsafe_allow_html=True)
        st.markdown("### Aprovação e Download")
        if "df_pcmso_gerado" in st.session_state:
            st.info(
                "Revise a matriz abaixo. Você pode corrigir nomes, alterar periodicidades, marcar ou desmarcar ADM/PER/DEM, excluir linhas e adicionar novos exames."
            )

            df_editado = st.data_editor(
                st.session_state["df_pcmso_gerado"],
                num_rows="dynamic",
                use_container_width=True,
                key="editor_matriz_pcmso",
                height=500,
            )

            st.markdown("---")

            if st.button("✅ Aprovar Matriz e Gerar Documentos", type="primary", use_container_width=True):
                cabecalho_atual = st.session_state["pcmso_cabecalho"]
                razao_social = cabecalho_atual.get("razao_social", "")
                medico_rt = cabecalho_atual.get("medico_rt", "")
                with st.spinner("Consolidando correções e gerando laudos oficiais..."):
                    try:
                        html_pcmso = gerar_html_pcmso(df_editado, cabecalho=cabecalho_atual)
                        bytes_docx = gerar_docx_rq61(df_editado, cabecalho=cabecalho_atual)
                    except Exception as e:
                        st.error(f"❌ Erro na geração dos documentos: {type(e).__name__}: {e}")
                        st.code(traceback.format_exc(), language="python")
                        st.stop()

                nome_arquivo = razao_social.replace(" ", "_")[:30] if razao_social else "PCMSO"
                st.markdown("### ⬇️ Documentos Prontos para Download")
                col_html, col_docx = st.columns(2)
                with col_html:
                    st.download_button(
                        label="📄 Baixar PCMSO (.html)",
                        data=html_pcmso.encode("utf-8"),
                        file_name=f"PCMSO_{nome_arquivo}.html",
                        mime="text/html",
                        use_container_width=True,
                    )
                with col_docx:
                    st.download_button(
                        label="📝 Baixar PCMSO (.docx)",
                        data=bytes_docx,
                        file_name=f"PCMSO_{nome_arquivo}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )

                with st.expander("👁️ Preview do PCMSO gerado", expanded=False):
                    components.html(html_pcmso, height=600, scrolling=True)

                if razao_social and medico_rt:
                    nome_proj = f"PCMSO - {razao_social[:40]} ({date.today().strftime('%d/%m/%Y')})"
                    salvar_historico(nome_proj, html_pcmso)
                    st.success("Laudo aprovado e salvo no histórico com sucesso!")
                else:
                    st.warning("Preencha Razão Social e Médico RT para salvar no histórico.")

                from modules.modulo_esocial_xml import render_botao_xml
                render_botao_xml(
                    df_editado,
                    cabecalho_atual,
                    dados_pgr=st.session_state.get("dados_ghe_processados"),
                    key_prefix="med",
                )
        else:
            st.info("A matriz aprovada aparecerá aqui após a extração do Passo 1.")
        st.markdown("</div>", unsafe_allow_html=True)

"""
Automacao SST - Seconci GO
app.py v5.4 — Login/Logout integrado
"""
import streamlit as st
import streamlit.components.v1 as components
import traceback
import os
from datetime import date

from config.db            import (get_supabase, salvar_historico,
                                   carregar_historico, carregar_html_historico)
from modules.modulo_pcmso import (extrair_texto_pdf, extrair_pgr_com_fallback,
                                   processar_pcmso, gerar_html_pcmso,
                                   gerar_docx_rq61)

st.set_page_config(page_title="Automacao SST - Seconci GO",
                   layout="wide", page_icon=":shield:")

st.markdown("""<style>
  .block-container{padding-top:2rem;padding-bottom:2rem;}
  .stButton>button{background-color:#084D22;color:white;border-radius:8px;
    border:none;box-shadow:0 4px 6px rgba(0,0,0,.1);transition:all .3s;
    font-weight:600;padding:.5rem 1rem;}
  .stButton>button:hover{background-color:#1AA04B;transform:translateY(-2px);}
  h1,h2,h3{color:#084D22!important;}
  [data-testid="stSidebar"]{background:#F4F8F5;border-right:1px solid #E0ECE4;}
  [data-testid="stFileUploadDropzone"]{border:2px dashed #1AA04B;
    border-radius:12px;background:#FAFFFA;}
  .stAlert{border-radius:8px;border-left:5px solid #084D22;}
</style>""", unsafe_allow_html=True)

# ── Sistema de Login ──────────────────────────────────────────────
def check_password():
    def validar_login():
        usr_digitado = st.session_state["username_input"]
        pwd_digitada = st.session_state["password_input"]
        usr_correto  = st.secrets.get("USUARIO_SISTEMA", "diovanni")
        pwd_correta  = st.secrets.get("SENHA_SISTEMA",   "seconci123")
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
            <div style='border:1px solid #1AA04B;padding:30px;
                        border-radius:10px;background-color:#F8F9FA;
                        box-shadow:0 4px 12px rgba(8,77,34,.1);'>
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

# ── Sidebar (só renderiza após login) ────────────────────────────
for logo in ("logo.png", "logo.jpg"):
    if os.path.exists(logo):
        st.sidebar.image(logo, width=220)
        break
else:
    st.sidebar.markdown(
        "<h2 style='text-align:center;color:#084D22;'>SECONCI-GO</h2>",
        unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.title("Modulos do Sistema")
modulo = st.sidebar.radio("Selecione a funcionalidade:", [
    "Dashboard",
    "Engenharia: FISPQ / FDS - PGR",
    "Medicina: PGR - PCMSO",
])

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

# ── Roteador ─────────────────────────────────────────────────────
if historico_html:
    st.title("Laudo Carregado do Historico")
    components.html(historico_html, height=700, scrolling=True)

elif modulo == "Dashboard":
    st.title("Dashboard - Sistema Integrado SST")
    try:
        sb        = get_supabase()
        total     = sb.table("historico_laudos").select("id", count="exact").execute().count or 0
        total_cas = sb.table("dicionario_dinamico").select("cas", count="exact").execute().count or 0
    except Exception:
        total, total_cas = 0, 0
    c1, c2, c3 = st.columns(3)
    c1.metric("Laudos Gerados",        total)
    c2.metric("CAS no Banco Dinamico", total_cas)
    c3.metric("Modulos Ativos",        2)
    st.info("Use o menu lateral para acessar os modulos.")

elif modulo == "Engenharia: FISPQ / FDS - PGR":
    try:
        from modules.modulo_engenharia import render_engenharia
        render_engenharia()
    except ImportError:
        st.warning("modulo_engenharia.py nao encontrado.")

elif modulo == "Medicina: PGR - PCMSO":
    st.title("Modulo Medico: Importador de PGR e Gerador de PCMSO")
    st.info(
        "Motor de extracao em 2 etapas: "
        "Etapa 1 — Extracao Local (gratuita, instantanea). "
        "Etapa 2 — IA Gemini (fallback — acionada so se necessario)."
    )

    with st.expander("Dados de Identificacao do PCMSO (NR-07 item 7.5.19.1)", expanded=True):
        col1, col2 = st.columns(2)
        cab = st.session_state.get("pcmso_cabecalho", {})

        with col1:
            razao_social = st.text_input("Razao Social da Empresa *",           value=cab.get("razao_social",""))
            cnpj         = st.text_input("CNPJ *",                               value=cab.get("cnpj",""))
            medico_rt    = st.text_input("Medico Responsavel RT (Nome + CRM) *", value=cab.get("medico_rt",""))

        with col2:
            vig_ini  = st.date_input("Vigencia - Inicio", value=date.today())
            vig_fim  = st.date_input("Vigencia - Fim",    value=date.today())
            resp_tec = st.text_input("Tecnico SST Responsavel (opcional)",  value=cab.get("responsavel_tec",""))
            obra     = st.text_input("Obra / Unidade (opcional)",           value=cab.get("obra",""))
              
        st.markdown("**Base de Auditoria** *(opcional — compara o PCMSO gerado com uma matriz validada)*")
        import json, os as _os
        _bases = ["Não auditar agora"]
        if _os.path.exists("banco_matrizes_v1_1.json"):
            with open("banco_matrizes_v1_1.json", "r", encoding="utf-8") as _f:
                _bases += list(json.load(_f).keys())
        base_auditoria = st.selectbox("Selecione a base:", _bases)
        st.session_state["base_auditoria"] = None if base_auditoria == "Não auditar agora" else base_auditoria

        st.markdown("**Tipo de Ambiente da Obra** *(define o pacote de exames)*")
        _opcoes_amb = {
            "🏗️ Canteiro de Obras / Obra": "canteiro",
            "🏢 Escritório Corporativo":   "escritorio",
            "🔀 Misto (Canteiro + Escritório no mesmo PGR)": "misto",
        }
        _label_amb = st.radio(
            "Selecione o tipo:",
            list(_opcoes_amb.keys()),
            index=1,
            help=(
                "Canteiro → todo cargo recebe pacote completo (Audiometria, ECG, Espirometria, RX).\n"
                "Escritório → cargo admin recebe só Exame Clínico + Acuidade Visual.\n"
                "Misto → sistema detecta por GHE automaticamente."
            ),
            horizontal=True,
        )
        tipo_ambiente = _opcoes_amb[_label_amb]

        st.session_state["pcmso_cabecalho"] = {
            "razao_social":    razao_social,
            "cnpj":            cnpj,
            "medico_rt":       medico_rt,
            "vig_ini":         vig_ini.strftime("%d/%m/%Y"),
            "vig_fim":         vig_fim.strftime("%d/%m/%Y"),
            "responsavel_tec": resp_tec,
            "obra":            obra,
        }
        st.session_state["tipo_ambiente"] = tipo_ambiente

    pdf_file = st.file_uploader("Arraste o PDF do PGR aqui", type=["pdf"])

    if st.button("Extrair Riscos e Gerar PCMSO", use_container_width=True):
        if not pdf_file:
            st.error("Faca upload do PDF do PGR antes de continuar.")
            st.stop()

        with st.spinner("Extraindo texto do PDF..."):
            texto_pgr = extrair_texto_pdf(pdf_file)

        st.success(f"Texto extraido: {len(texto_pgr):,} caracteres em {pdf_file.name}")

        with st.expander("DEBUG: Primeiras 100 linhas do PDF"):
            for i, linha in enumerate(texto_pgr.split("\n")[:100], 1):
                st.text(f"{i:3}: {linha}")

        with st.spinner("Processando GHEs..."):
            dados_ghe, fonte = extrair_pgr_com_fallback(texto_pgr)

        if fonte == "local":
            st.success("Dados extraidos localmente — sem consumo de IA!")
        elif fonte == "ia":
            st.info("Dados extraidos via IA (Gemini).")
        else:
            st.warning("Extracao parcial — revise os resultados.")

        if not dados_ghe:
            st.error("Nenhum GHE identificado. Verifique se o PDF e um PGR valido.")
            st.stop()

        with st.expander("DEBUG: Estrutura dos GHEs extraídos (primeiros 3)"):
            for g in dados_ghe[:3]:
                st.json(g)

        tipo_amb = st.session_state.get("tipo_ambiente", "escritorio")
        with st.spinner(f"Gerando matriz PCMSO ({tipo_amb})..."):
            try:
                df_pcmso = processar_pcmso(dados_ghe, tipo_ambiente=tipo_amb)
            except Exception as e:
                st.error(f"❌ Erro em processar_pcmso(): {type(e).__name__}: {e}")
                st.code(traceback.format_exc(), language="python")
                st.stop()

        if df_pcmso.empty:
            st.warning("PCMSO gerado vazio — nenhum cargo/exame identificado. Verifique os GHEs extraídos.")
            st.stop()

        st.success(f"PCMSO gerado com {len(df_pcmso)} linhas de exames.")
        st.dataframe(df_pcmso, use_container_width=True)

              base_sel = st.session_state.get("base_auditoria")
        if base_sel:
            from modules.modulo_auditor_v1_1 import auditar_pcmso
            with st.spinner("Auditando PCMSO..."):
                relatorio_auditoria = auditar_pcmso(df_pcmso, base_sel)
            with st.expander("📋 Relatório da Auditoria", expanded=True):
                components.html(relatorio_auditoria, height=500, scrolling=True)

        cabecalho_atual = st.session_state["pcmso_cabecalho"]

        try:
            html_pcmso = gerar_html_pcmso(df_pcmso, cabecalho=cabecalho_atual)
        except Exception as e:
            st.error(f"❌ Erro em gerar_html_pcmso(): {type(e).__name__}: {e}")
            st.code(traceback.format_exc(), language="python")
            st.stop()

        with st.spinner("Gerando DOCX..."):
            try:
                bytes_docx = gerar_docx_rq61(df_pcmso, cabecalho=cabecalho_atual)
            except Exception as e:
                st.error(f"❌ Erro em gerar_docx_pcmso(): {type(e).__name__}: {e}")
                st.code(traceback.format_exc(), language="python")
                st.stop()

        nome_arquivo = razao_social.replace(" ", "_")[:30] if razao_social else "PCMSO"

        st.markdown("### ⬇️ Baixar PCMSO")
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
            st.success("Laudo salvo no historico!")
        else:
            st.warning("Preencha Razao Social e Medico RT para salvar no historico.")

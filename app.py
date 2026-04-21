"""
Automacao SST - Seconci GO
app.py v5.2 — campo tipo_ambiente no formulario PCMSO
+ auditoria automatica por matriz validada
+ pos-processamento corretivo v5.2
"""
import os
from pathlib import Path
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
    processar_pcmso,
    gerar_html_pcmso,
    gerar_docx_pcmso,
)

from modules.modulo_pcmso_v5_2 import aplicar_posprocessamento_v52
from modules.modulo_auditor_v1_1 import (
    carregar_banco_matrizes,
    auditar_pcmso,
    formatar_relatorio_auditoria,
)


st.set_page_config(
    page_title="Automacao SST - Seconci GO",
    layout="wide",
    page_icon=":shield:"
)

BANCO_MATRIZES = carregar_banco_matrizes(
    Path("data/banco_matrizes_v1_1.json")
)

MAPA_AUDITORIA = {
    "Não auditar agora": None,
    "Vistamérica 2025": "vistamerica_2025",
    "CMO Construtora ADM 2025": "cmo_construtora_adm_2025",
    "Viverde Areião 2025": "viverde_areiao_2025",
}


def df_pcmso_para_dict(df):
    dados = {}

    for _, row in df.iterrows():
        ghe = str(row.get("GHE / Setor", "")).strip()
        cargo = str(row.get("Cargo", "")).strip()
        exame = str(row.get("Exame", "")).strip()

        per = str(row.get("PER", "-")).strip()
        if per in ("-", "", "None"):
            per = None
        else:
            per = (
                per.replace("MESES", "")
                .replace("M", "")
                .strip()
            )

        item = {
            "nome": exame,
            "adm": str(row.get("ADM", "-")).strip().upper() == "X",
            "per": per,
            "mro": str(row.get("MRO", "-")).strip().upper() == "X",
            "ret": str(row.get("RT", "-")).strip().upper() == "X",
            "dem": str(row.get("DEM", "-")).strip().upper() == "X",
            "motivo": str(row.get("Justificativa", "")).strip(),
        }

        dados.setdefault(ghe, {}).setdefault(cargo, []).append(item)

    return dados


def dict_pcmso_para_df(dados):
    linhas = []

    for ghe, cargos in dados.items():
        for cargo, exames in cargos.items():
            for ex in exames:
                per = ex.get("per")
                if per in (None, "", "-", False):
                    per_fmt = "-"
                else:
                    per_str = str(per).strip()
                    per_fmt = f"{per_str}M" if per_str.isdigit() else per_str

                linhas.append({
                    "GHE / Setor": ghe,
                    "Cargo": cargo,
                    "Exame": ex.get("nome", ""),
                    "ADM": "X" if ex.get("adm") else "-",
                    "PER": per_fmt,
                    "MRO": "X" if ex.get("mro") else "-",
                    "RT": "X" if ex.get("ret") else "-",
                    "DEM": "X" if ex.get("dem") else "-",
                    "Justificativa": ex.get("motivo", ""),
                })

    df = pd.DataFrame(linhas)

    if not df.empty:
        colunas = [
            "GHE / Setor", "Cargo", "Exame",
            "ADM", "PER", "MRO", "RT", "DEM",
            "Justificativa"
        ]
        df = df[colunas]

    return df


st.markdown("""
<style>
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
</style>
""", unsafe_allow_html=True)


# ── Sidebar ───────────────────────────────────────────────────────
for logo in ("logo.png", "logo.jpg"):
    if os.path.exists(logo):
        st.sidebar.image(logo, width=220)
        break
else:
    st.sidebar.markdown(
        "<h2 style='text-align:center;color:#084D22;'>SECONCI-GO</h2>",
        unsafe_allow_html=True
    )

st.sidebar.markdown("---")
st.sidebar.title("Modulos do Sistema")
modulo = st.sidebar.radio(
    "Selecione a funcionalidade:",
    [
        "Dashboard",
        "Engenharia: FISPQ / FDS - PGR",
        "Medicina: PGR - PCMSO",
    ]
)

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
        sb = get_supabase()
        total = sb.table("historico_laudos").select("id", count="exact").execute().count or 0
        total_cas = sb.table("dicionario_dinamico").select("cas", count="exact").execute().count or 0
    except Exception:
        total, total_cas = 0, 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Laudos Gerados", total)
    c2.metric("CAS no Banco Dinamico", total_cas)
    c3.metric("Modulos Ativos", 2)
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
            razao_social = st.text_input(
                "Razao Social da Empresa *",
                value=cab.get("razao_social", "")
            )
            cnpj = st.text_input(
                "CNPJ *",
                value=cab.get("cnpj", "")
            )
            medico_rt = st.text_input(
                "Medico Responsavel RT (Nome + CRM) *",
                value=cab.get("medico_rt", "")
            )

        with col2:
            vig_ini = st.date_input("Vigencia - Inicio", value=date.today())
            vig_fim = st.date_input("Vigencia - Fim", value=date.today())
            resp_tec = st.text_input(
                "Tecnico SST Responsavel (opcional)",
                value=cab.get("responsavel_tec", "")
            )
            obra = st.text_input(
                "Obra / Unidade (opcional)",
                value=cab.get("obra", "")
            )

        st.markdown("**Tipo de Ambiente da Obra** *(define o pacote de exames)*")
        _opcoes_amb = {
            "🏗️ Canteiro de Obras / Obra": "canteiro",
            "🏢 Escritório Corporativo": "escritorio",
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

        st.markdown("**Base de Auditoria (opcional)**")
        auditoria_label = st.selectbox(
            "Selecione a matriz validada para comparar:",
            list(MAPA_AUDITORIA.keys()),
            index=0,
            help="Use apenas quando esta obra corresponder a uma matriz já validada."
        )
        obra_id_auditoria = MAPA_AUDITORIA[auditoria_label]

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
        st.session_state["obra_id_auditoria"] = obra_id_auditoria

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

        tipo_amb = st.session_state.get("tipo_ambiente", "escritorio")
        obra_id_auditoria = st.session_state.get("obra_id_auditoria")

        with st.spinner(f"Gerando matriz PCMSO ({tipo_amb})..."):
            df_pcmso_bruto = processar_pcmso(dados_ghe, tipo_ambiente=tipo_amb)
            dados_pcmso = df_pcmso_para_dict(df_pcmso_bruto)
            dados_pcmso_corrigidos = aplicar_posprocessamento_v52(dados_pcmso)
            df_pcmso = dict_pcmso_para_df(dados_pcmso_corrigidos)

        st.success(f"PCMSO gerado com {len(df_pcmso)} linhas de exames.")
        st.dataframe(df_pcmso, use_container_width=True)

        resultado_auditoria = None
        if obra_id_auditoria:
            resultado_auditoria = auditar_pcmso(
                dados_pcmso_corrigidos,
                BANCO_MATRIZES,
                obra_id=obra_id_auditoria
            )

            if resultado_auditoria["total_divergencias"] > 0:
                st.warning(
                    f"Auditoria encontrou {resultado_auditoria['total_divergencias']} divergencias."
                )
            else:
                st.success("Auditoria concluida: nenhuma divergencia encontrada.")

            with st.expander("📋 Relatorio da Auditoria", expanded=False):
                st.code(formatar_relatorio_auditoria(resultado_auditoria))

        cabecalho_atual = st.session_state["pcmso_cabecalho"]
        html_pcmso = gerar_html_pcmso(df_pcmso, cabecalho=cabecalho_atual)

        with st.spinner("Gerando DOCX..."):
            bytes_docx = gerar_docx_pcmso(df_pcmso, cabecalho=cabecalho_atual)

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
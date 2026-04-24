"""
modulo_construtor_visual.py — Construtor Visual de GHEs
Passo 3: permite montar a estrutura de GHEs manualmente,
sem necessidade de PDF, usando interface Streamlit interativa.
"""

import json
import os

import streamlit as st

from modules.modulo_pcmso import (
    processar_pcmso,
    gerar_html_pcmso,
    gerar_docx_rq61,
)
from utils.cargo_utils import MAPA_CARGOS_CONHECIDOS

# ── Listas de referência ──────────────────────────────────────────────────────
_RISCOS_FISICOS = [
    "RUIDO", "VIBRACAO CORPO INTEIRO", "VIBRACAO MAOS E BRACOS",
    "CALOR", "FRIO", "RADIACAO NAO IONIZANTE", "RADIACAO IONIZANTE",
    "PRESSAO ANORMAL", "UMIDADE",
]

_RISCOS_QUIMICOS = [
    "BENZENO", "TOLUENO", "XILENO", "ACETONA", "METIL-ETIL-CETONA",
    "N-HEXANO", "DICLOROMETANO", "TRICLOROETILENO", "ESTIRENO",
    "METANOL", "FENOL", "CHUMBO", "MANGANES", "MERCURIO", "CROMO",
    "CADMIO", "ARSENICO", "COBALTO", "FLUOR", "SILICA", "CIMENTO",
    "ASBESTO", "POEIRA MINERAL", "FUMOS METALICOS", "DIESEL",
    "GASOLINA", "TINTA", "IMPERMEABILIZACAO", "MADEIRA", "SOLDA",
    "MONOXIDO DE CARBONO",
]

_RISCOS_BIOLOGICOS = [
    "AGENTE BIOLOGICO", "ESGOTO", "EFLUENTE",
    "BACTERIA", "VIRUS", "FUNGO",
]

_RISCOS_ERGONOMICOS = [
    "ERGONOMICO", "ESFORCO REPETITIVO", "POSTURA INADEQUADA",
    "LEVANTAMENTO DE PESO", "TRABALHO EM PE",
]

_RISCOS_ACIDENTE = [
    "QUEDA DE ALTURA", "ESPACO CONFINADO", "RISCO ELETRICO",
    "MAQUINAS E EQUIPAMENTOS", "INCENDIO", "EXPLOSAO",
    "ANIMAIS PECONHENTOS", "FERRAMENTAS MANUAIS",
]

_TODOS_RISCOS = (
    [(r, "Físico") for r in _RISCOS_FISICOS]
    + [(r, "Químico") for r in _RISCOS_QUIMICOS]
    + [(r, "Biológico") for r in _RISCOS_BIOLOGICOS]
    + [(r, "Ergonômico") for r in _RISCOS_ERGONOMICOS]
    + [(r, "Acidente") for r in _RISCOS_ACIDENTE]
)

_TODOS_RISCOS_NOMES = [r for r, _ in _TODOS_RISCOS]
_MAPA_RISCO_TIPO = {r: t for r, t in _TODOS_RISCOS}

_CARGOS_LISTA = sorted(set(MAPA_CARGOS_CONHECIDOS))


# ── Helpers de estado ─────────────────────────────────────────────────────────
def _init_state():
    if "cv_ghes" not in st.session_state:
        st.session_state["cv_ghes"] = []  # lista de dicts {ghe, cargos, riscos_mapeados}
    if "cv_ghe_editando" not in st.session_state:
        st.session_state["cv_ghe_editando"] = None


def _ghe_vazio(nome: str) -> dict:
    return {"ghe": nome, "cargos": [], "riscos_mapeados": []}


def _indice_ghe(nome: str) -> int | None:
    for i, g in enumerate(st.session_state["cv_ghes"]):
        if g["ghe"] == nome:
            return i
    return None


# ── Render principal ──────────────────────────────────────────────────────────
def render_construtor_visual():
    _init_state()

    st.title("🏗️ Construtor Visual de GHEs")
    st.info(
        "Monte a estrutura de GHEs manualmente, sem precisar de um PDF. "
        "Adicione quantos GHEs quiser, defina cargos e riscos para cada um, "
        "e gere o PCMSO ao final."
    )

    # ── Painel esquerdo: lista de GHEs ────────────────────────────────────────
    col_lista, col_editor = st.columns([1, 2])

    with col_lista:
        st.markdown("### 📋 GHEs Criados")

        if not st.session_state["cv_ghes"]:
            st.caption("Nenhum GHE criado ainda.")

        for i, ghe in enumerate(st.session_state["cv_ghes"]):
            cols = st.columns([3, 1])
            label = f"GHE {i+1:02d} — {ghe['ghe'][:30]}"
            n_cargos = len(ghe["cargos"])
            n_riscos = len(ghe["riscos_mapeados"])
            if cols[0].button(
                f"{label}\n{n_cargos} cargo(s) · {n_riscos} risco(s)",
                key=f"btn_ghe_{i}",
                use_container_width=True,
            ):
                st.session_state["cv_ghe_editando"] = ghe["ghe"]

            if cols[1].button("🗑️", key=f"del_ghe_{i}", help="Excluir GHE"):
                st.session_state["cv_ghes"].pop(i)
                if st.session_state["cv_ghe_editando"] == ghe["ghe"]:
                    st.session_state["cv_ghe_editando"] = None
                st.rerun()

        st.markdown("---")
        st.markdown("#### ➕ Novo GHE")
        novo_nome = st.text_input(
            "Nome do GHE",
            placeholder="Ex: GHE 01 - Canteiro de Estrutura",
            key="cv_novo_nome",
        )
        if st.button("Adicionar GHE", use_container_width=True, key="cv_btn_add"):
            nome = novo_nome.strip()
            if not nome:
                st.warning("Digite um nome para o GHE.")
            elif _indice_ghe(nome) is not None:
                st.warning("Já existe um GHE com esse nome.")
            else:
                st.session_state["cv_ghes"].append(_ghe_vazio(nome))
                st.session_state["cv_ghe_editando"] = nome
                st.rerun()

        # Importar GHEs de sessão anterior (PCMSO já processado)
        st.markdown("---")
        st.markdown("#### 📥 Importar GHEs do PGR")
        st.caption("Se você já processou um PGR nesta sessão, pode importar os GHEs aqui.")
        if st.button("Importar da sessão atual", key="cv_btn_importar", use_container_width=True):
            dados_sessao = st.session_state.get("dados_ghe_processados", [])
            if not dados_sessao:
                st.warning("Nenhum PGR processado nesta sessão. Processe um PDF no módulo Medicina primeiro.")
            else:
                importados = 0
                for g in dados_sessao:
                    nome_g = g.get("ghe", "").strip()
                    if nome_g and _indice_ghe(nome_g) is None:
                        st.session_state["cv_ghes"].append({
                            "ghe": nome_g,
                            "cargos": list(g.get("cargos", [])),
                            "riscos_mapeados": list(g.get("riscos_mapeados", [])),
                        })
                        importados += 1
                st.success(f"{importados} GHE(s) importado(s) com sucesso!")
                st.rerun()

    # ── Painel direito: editor do GHE selecionado ─────────────────────────────
    with col_editor:
        ghe_nome_editando = st.session_state.get("cv_ghe_editando")
        idx = _indice_ghe(ghe_nome_editando) if ghe_nome_editando else None

        if idx is None:
            st.markdown("### ✏️ Editor de GHE")
            st.caption("Selecione um GHE na lista à esquerda ou crie um novo para começar a editar.")
        else:
            ghe = st.session_state["cv_ghes"][idx]
            st.markdown(f"### ✏️ Editando: **{ghe['ghe']}**")

            tab_cargos, tab_riscos, tab_resumo = st.tabs(["👷 Cargos", "⚠️ Riscos", "📊 Resumo"])

            # ── Tab Cargos ────────────────────────────────────────────────────
            with tab_cargos:
                st.markdown("**Cargos atribuídos a este GHE:**")

                # Remove cargo
                for ci, cargo in enumerate(list(ghe["cargos"])):
                    c1, c2 = st.columns([4, 1])
                    c1.markdown(f"• {cargo}")
                    if c2.button("❌", key=f"rm_cargo_{idx}_{ci}", help="Remover"):
                        ghe["cargos"].remove(cargo)
                        st.rerun()

                if not ghe["cargos"]:
                    st.caption("Nenhum cargo adicionado.")

                st.markdown("---")
                st.markdown("**Adicionar cargo:**")
                c_sel, c_outro = st.columns([2, 2])

                with c_sel:
                    cargo_sel = st.selectbox(
                        "Selecionar da lista",
                        ["— selecione —"] + _CARGOS_LISTA,
                        key=f"cv_sel_cargo_{idx}",
                    )
                    if st.button("+ Adicionar selecionado", key=f"add_cargo_sel_{idx}"):
                        if cargo_sel != "— selecione —":
                            if cargo_sel not in ghe["cargos"]:
                                ghe["cargos"].append(cargo_sel)
                                st.rerun()
                            else:
                                st.warning("Cargo já adicionado.")

                with c_outro:
                    cargo_livre = st.text_input(
                        "Ou digitar cargo personalizado",
                        placeholder="Ex: Soldador de Estrutura",
                        key=f"cv_cargo_livre_{idx}",
                    )
                    if st.button("+ Adicionar digitado", key=f"add_cargo_livre_{idx}"):
                        nome_c = cargo_livre.strip().upper()
                        if nome_c:
                            if nome_c not in ghe["cargos"]:
                                ghe["cargos"].append(nome_c)
                                st.rerun()
                            else:
                                st.warning("Cargo já adicionado.")

            # ── Tab Riscos ────────────────────────────────────────────────────
            with tab_riscos:
                st.markdown("**Riscos mapeados neste GHE:**")

                riscos_por_tipo: dict[str, list] = {}
                for r in ghe["riscos_mapeados"]:
                    tipo = _MAPA_RISCO_TIPO.get(r["nome_agente"].upper(), "Outro")
                    riscos_por_tipo.setdefault(tipo, []).append(r)

                for tipo, lista in riscos_por_tipo.items():
                    st.markdown(f"**{tipo}**")
                    for ri, risco in enumerate(list(lista)):
                        r1, r2 = st.columns([4, 1])
                        r1.markdown(f"• {risco['nome_agente']}")
                        chave_btn = f"rm_risco_{idx}_{risco['nome_agente'].replace(' ','_')}"
                        if r2.button("❌", key=chave_btn, help="Remover risco"):
                            ghe["riscos_mapeados"] = [
                                r for r in ghe["riscos_mapeados"]
                                if r["nome_agente"] != risco["nome_agente"]
                            ]
                            st.rerun()

                if not ghe["riscos_mapeados"]:
                    st.caption("Nenhum risco adicionado.")

                st.markdown("---")
                st.markdown("**Adicionar risco:**")

                r_col1, r_col2 = st.columns([3, 1])
                with r_col1:
                    tipo_filtro = st.selectbox(
                        "Filtrar por tipo",
                        ["Todos", "Físico", "Químico", "Biológico", "Ergonômico", "Acidente"],
                        key=f"cv_tipo_risco_{idx}",
                    )
                    lista_filtrada = (
                        _TODOS_RISCOS_NOMES
                        if tipo_filtro == "Todos"
                        else [r for r, t in _TODOS_RISCOS if t == tipo_filtro]
                    )
                    risco_sel = st.selectbox(
                        "Selecionar risco",
                        ["— selecione —"] + lista_filtrada,
                        key=f"cv_sel_risco_{idx}",
                    )

                with r_col2:
                    st.markdown("&nbsp;", unsafe_allow_html=True)
                    st.markdown("&nbsp;", unsafe_allow_html=True)
                    if st.button("+ Adicionar", key=f"add_risco_sel_{idx}", use_container_width=True):
                        if risco_sel != "— selecione —":
                            nomes_existentes = [r["nome_agente"].upper() for r in ghe["riscos_mapeados"]]
                            if risco_sel.upper() not in nomes_existentes:
                                ghe["riscos_mapeados"].append({
                                    "nome_agente": risco_sel,
                                    "perigo_especifico": f"Mapeado manualmente via Construtor Visual",
                                })
                                st.rerun()
                            else:
                                st.warning("Risco já adicionado.")

                # Risco personalizado
                st.markdown("**Ou adicionar risco personalizado:**")
                risco_livre = st.text_input(
                    "Nome do agente",
                    placeholder="Ex: Acrílico, Isocianato, Solvente de Borracha",
                    key=f"cv_risco_livre_{idx}",
                )
                if st.button("+ Adicionar agente personalizado", key=f"add_risco_livre_{idx}"):
                    nome_r = risco_livre.strip().upper()
                    if nome_r:
                        nomes_existentes = [r["nome_agente"].upper() for r in ghe["riscos_mapeados"]]
                        if nome_r not in nomes_existentes:
                            ghe["riscos_mapeados"].append({
                                "nome_agente": nome_r,
                                "perigo_especifico": "Adicionado manualmente — Construtor Visual",
                            })
                            st.rerun()
                        else:
                            st.warning("Agente já adicionado.")

            # ── Tab Resumo ────────────────────────────────────────────────────
            with tab_resumo:
                st.markdown(f"**GHE:** {ghe['ghe']}")
                st.markdown(f"**Cargos ({len(ghe['cargos'])}):** {', '.join(ghe['cargos']) or '—'}")
                st.markdown(f"**Riscos ({len(ghe['riscos_mapeados'])}):**")
                for r in ghe["riscos_mapeados"]:
                    tipo = _MAPA_RISCO_TIPO.get(r["nome_agente"].upper(), "Outro")
                    st.markdown(f"  • `{r['nome_agente']}` _{tipo}_")

    # ── Rodapé: gerar PCMSO ───────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### ⚙️ Gerar PCMSO a partir do Construtor Visual")

    if not st.session_state["cv_ghes"]:
        st.warning("Adicione pelo menos um GHE antes de gerar o PCMSO.")
        return

    ghes_validos = [
        g for g in st.session_state["cv_ghes"]
        if g["cargos"] or g["riscos_mapeados"]
    ]
    st.caption(
        f"{len(ghes_validos)} de {len(st.session_state['cv_ghes'])} GHE(s) "
        "com cargos/riscos preenchidos serão processados."
    )

    # Tipo de ambiente
    opcoes_amb = {
        "🏗️ Canteiro de Obras": "canteiro",
        "🏢 Escritório": "escritorio",
        "🔀 Misto": "misto",
    }
    label_amb = st.radio(
        "Tipo de ambiente:",
        list(opcoes_amb.keys()),
        horizontal=True,
        key="cv_tipo_ambiente",
    )
    tipo_amb = opcoes_amb[label_amb]

    # Cabeçalho rápido
    with st.expander("📝 Cabeçalho do PCMSO (opcional)", expanded=False):
        cab_col1, cab_col2 = st.columns(2)
        with cab_col1:
            cv_razao = st.text_input("Razão Social", key="cv_razao")
            cv_cnpj = st.text_input("CNPJ", key="cv_cnpj")
            cv_medico = st.text_input("Médico RT", key="cv_medico")
        with cab_col2:
            cv_obra = st.text_input("Obra / Unidade", key="cv_obra")
            cv_tec = st.text_input("Técnico SST", key="cv_tec")
            cv_crm = st.text_input("CRM", key="cv_crm")

    if st.button(
        "🚀 Gerar PCMSO do Construtor Visual",
        type="primary",
        use_container_width=True,
        key="cv_btn_gerar",
    ):
        if not ghes_validos:
            st.error("Nenhum GHE tem cargos ou riscos preenchidos. Adicione ao menos um cargo ou risco.")
            return

        import traceback
        import streamlit.components.v1 as components
        from datetime import date

        cabecalho_cv = {
            "razao_social": st.session_state.get("cv_razao", ""),
            "cnpj": st.session_state.get("cv_cnpj", ""),
            "medico_rt": st.session_state.get("cv_medico", ""),
            "obra": st.session_state.get("cv_obra", ""),
            "responsavel_tec": st.session_state.get("cv_tec", ""),
            "crm": st.session_state.get("cv_crm", ""),
            "vig_ini": date.today().strftime("%d/%m/%Y"),
            "vig_fim": date.today().strftime("%d/%m/%Y"),
        }

        with st.spinner("Processando PCMSO..."):
            try:
                df_cv = processar_pcmso(ghes_validos, tipo_ambiente=tipo_amb)
            except Exception as e:
                st.error(f"❌ Erro em processar_pcmso(): {type(e).__name__}: {e}")
                st.code(traceback.format_exc(), language="python")
                return

        if df_cv.empty:
            st.warning("PCMSO vazio — verifique se os GHEs têm cargos e riscos configurados.")
            return

        st.success(f"✅ PCMSO gerado com {len(df_cv)} linhas de exames.")

        # Staging area — triagem médica
        st.markdown("#### 👩‍⚕️ Triagem Médica")
        df_cv_editado = st.data_editor(
            df_cv,
            num_rows="dynamic",
            use_container_width=True,
            key="cv_editor_matriz",
            height=450,
        )

        if st.button("✅ Aprovar e Baixar Documentos", type="primary", use_container_width=True, key="cv_btn_aprovar"):
            try:
                html_cv = gerar_html_pcmso(df_cv_editado, cabecalho=cabecalho_cv)
                docx_cv = gerar_docx_rq61(df_cv_editado, cabecalho=cabecalho_cv)
            except Exception as e:
                st.error(f"❌ Erro na geração: {type(e).__name__}: {e}")
                st.code(traceback.format_exc(), language="python")
                return

            nome_arq = (cabecalho_cv.get("razao_social") or "PCMSO_Visual").replace(" ", "_")[:30]
            dl1, dl2 = st.columns(2)
            with dl1:
                st.download_button(
                    "📄 Baixar HTML",
                    data=html_cv.encode("utf-8"),
                    file_name=f"PCMSO_{nome_arq}.html",
                    mime="text/html",
                    use_container_width=True,
                )
            with dl2:
                st.download_button(
                    "📝 Baixar DOCX",
                    data=docx_cv,
                    file_name=f"PCMSO_{nome_arq}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

            with st.expander("👁️ Preview", expanded=False):
                components.html(html_cv, height=600, scrolling=True)

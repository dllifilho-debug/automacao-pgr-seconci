"""
modules/modulo_engenharia.py
Modulo de Engenharia: FISPQ / FDS -> PGR
Extraido e refatorado do app.py original (Automacao_PGR local).
"""
import streamlit as st
import streamlit.components.v1 as components
import pdfplumber
import re
import pandas as pd
import sqlite3
import requests
from datetime import datetime

# ── Dicionarios ──────────────────────────────────────────────────

DICIONARIO_H = {
    "H315": {"desc": "Provoca irritacao a pele", "sev": 1, "epi": "Luvas de protecao e vestimenta"},
    "H319": {"desc": "Provoca irritacao ocular grave", "sev": 1, "epi": "Oculos de protecao contra respingos"},
    "H336": {"desc": "Pode provocar sonolencia ou vertigem", "sev": 1, "epi": "Local ventilado; mascara se necessario"},
    "H317": {"desc": "Reacoes alergicas na pele", "sev": 2, "epi": "Luvas (nitrilica/PVC) e manga longa"},
    "H335": {"desc": "Irritacao das vias respiratorias", "sev": 2, "epi": "Protecao respiratoria (Filtro especifico)"},
    "H302": {"desc": "Nocivo em caso de ingestao", "sev": 2, "epi": "Higiene rigorosa; luvas adequadas"},
    "H312": {"desc": "Nocivo em contato com a pele", "sev": 2, "epi": "Luvas impermeaveis e avental"},
    "H332": {"desc": "Nocivo se inalado", "sev": 2, "epi": "Mascara respiratoria adequada (PFF2/VO)"},
    "H314": {"desc": "Queimadura severa a pele e dano", "sev": 4, "epi": "Traje quimico, luvas longas e botas"},
    "H318": {"desc": "Lesoes oculares graves", "sev": 3, "epi": "Oculos ampla visao / protetor facial"},
    "H301": {"desc": "Toxico em caso de ingestao", "sev": 4, "epi": "Higiene rigorosa; luvas"},
    "H311": {"desc": "Toxico em contato com a pele", "sev": 4, "epi": "Luvas, avental impermeavel e botas"},
    "H331": {"desc": "Toxico se inalado", "sev": 4, "epi": "Respirador facial inteiro"},
    "H334": {"desc": "Sintomas de asma ou dificuldades respiratorias", "sev": 4, "epi": "Respirador facial inteiro (Filtro P3)"},
    "H372": {"desc": "Danos aos orgaos (exposicao prolongada/repetida)", "sev": 4, "epi": "Protecao respiratoria e dermica estrita"},
    "H373": {"desc": "Pode provocar danos aos orgaos (exp. repetida)", "sev": 4, "epi": "Avaliar via; EPIs combinados obrigatorios"},
    "H300": {"desc": "Fatal em caso de ingestao", "sev": 5, "epi": "Isolamento total; Higiene extrema"},
    "H310": {"desc": "Fatal em contato com a pele", "sev": 5, "epi": "Traje encapsulado nivel A/B"},
    "H330": {"desc": "Fatal se inalado", "sev": 5, "epi": "Equipamento de Respiracao Autonoma (EPR)"},
    "H340": {"desc": "Pode provocar defeitos geneticos", "sev": 5, "epi": "Isolamento; Traje quimico e EPR"},
    "H350": {"desc": "Pode provocar cancer", "sev": 5, "epi": "Isolamento; Traje quimico e EPR completo"},
    "H351": {"desc": "Suspeito de provocar cancer", "sev": 4, "epi": "Protecao respiratoria e dermica maxima"},
    "H360": {"desc": "Pode prejudicar a fertilidade ou o feto", "sev": 5, "epi": "Afastamento de gestantes; EPI maximo"},
    "H370": {"desc": "Provoca danos aos orgaos", "sev": 5, "epi": "EPI maximo conforme via de exposicao"},
}

MATRIZ_OFICIAL = {
    (1,1): "TRIVIAL",      (1,2): "TRIVIAL",      (1,3): "TOLERAVEL",    (1,4): "TOLERAVEL",    (1,5): "MODERADO",
    (2,1): "TRIVIAL",      (2,2): "TOLERAVEL",    (2,3): "MODERADO",     (2,4): "MODERADO",     (2,5): "SUBSTANCIAL",
    (3,1): "TOLERAVEL",    (3,2): "TOLERAVEL",    (3,3): "MODERADO",     (3,4): "SUBSTANCIAL",  (3,5): "SUBSTANCIAL",
    (4,1): "TOLERAVEL",    (4,2): "MODERADO",     (4,3): "SUBSTANCIAL",  (4,4): "INTOLERAVEL",  (4,5): "INTOLERAVEL",
    (5,1): "MODERADO",     (5,2): "MODERADO",     (5,3): "SUBSTANCIAL",  (5,4): "INTOLERAVEL",  (5,5): "INTOLERAVEL",
}

ACOES_REQUERIDAS = {
    "TRIVIAL":      "Manter controles existentes; monitoramento periodico.",
    "TOLERAVEL":    "Manter controles. Considerar melhorias.",
    "MODERADO":     "Implantar controles. EPI e monitoramento PCMSO.",
    "SUBSTANCIAL":  "Trabalho nao deve iniciar sem reducao do risco.",
    "INTOLERAVEL":  "TRABALHO PROIBIDO. Risco grave e iminente.",
}

TEXTO_SEV = {1: "1-LEVE", 2: "2-BAIXA", 3: "3-MODERADA", 4: "4-ALTA", 5: "5-EXTREMA"}

DICIONARIO_CAS = {
    "108-88-3":   {"agente": "Tolueno",                  "nr15_lt": "78 ppm ou 290 mg/m3",           "nr09_acao": "39 ppm ou 145 mg/m3",          "nr07_ibe": "o-Cresol ou Acido Hipurico",         "dec_3048": "25 anos (Linha 1.0.19)", "esocial_24": "01.19.036"},
    "1330-20-7":  {"agente": "Xileno",                   "nr15_lt": "78 ppm ou 340 mg/m3",           "nr09_acao": "39 ppm ou 170 mg/m3",          "nr07_ibe": "Acidos Metilhipuricos",              "dec_3048": "25 anos (Linha 1.0.19)", "esocial_24": "01.19.036"},
    "71-43-2":    {"agente": "Benzeno",                  "nr15_lt": "VRT-MPT (Cancerigeno)",          "nr09_acao": "Avaliacao Qualitativa",         "nr07_ibe": "Acido trans,trans-muconico",         "dec_3048": "25 anos (Linha 1.0.3)",  "esocial_24": "01.01.006"},
    "67-64-1":    {"agente": "Acetona",                  "nr15_lt": "780 ppm ou 1870 mg/m3",         "nr09_acao": "390 ppm ou 935 mg/m3",         "nr07_ibe": "Avaliacao Clinica",                  "dec_3048": "Nao Enquadrado",         "esocial_24": "09.01.001"},
    "64-17-5":    {"agente": "Etanol (Alcool Etilico)",  "nr15_lt": "780 ppm ou 1480 mg/m3",         "nr09_acao": "390 ppm ou 740 mg/m3",         "nr07_ibe": "Avaliacao Clinica",                  "dec_3048": "Nao Enquadrado",         "esocial_24": "09.01.001"},
    "78-93-3":    {"agente": "Metiletilcetona (MEK)",    "nr15_lt": "155 ppm ou 460 mg/m3",          "nr09_acao": "77.5 ppm ou 230 mg/m3",        "nr07_ibe": "MEK na Urina",                       "dec_3048": "Nao Enquadrado",         "esocial_24": "09.01.001"},
    "110-54-3":   {"agente": "n-Hexano",                 "nr15_lt": "50 ppm ou 176 mg/m3",           "nr09_acao": "25 ppm ou 88 mg/m3",           "nr07_ibe": "2,5-Hexanodiona",                    "dec_3048": "25 anos (Linha 1.0.19)", "esocial_24": "01.19.014"},
    "14808-60-7": {"agente": "Silica Cristalina (Quartzo)", "nr15_lt": "Anexo 12",                   "nr09_acao": "50% do L.T.",                  "nr07_ibe": "Raio-X (OIT) e Espirometria",        "dec_3048": "25 anos (Linha 1.0.18)", "esocial_24": "01.18.001"},
    "1332-21-4":  {"agente": "Asbesto / Amianto",        "nr15_lt": "0,1 f/cm3",                     "nr09_acao": "0,05 f/cm3",                    "nr07_ibe": "Raio-X (OIT) e Espirometria",        "dec_3048": "20 anos (Linha 1.0.2)",  "esocial_24": "01.02.001"},
    "7439-92-1":  {"agente": "Chumbo (Fumos)",           "nr15_lt": "0,1 mg/m3",                     "nr09_acao": "0,05 mg/m3",                    "nr07_ibe": "Chumbo no sangue e ALA-U",           "dec_3048": "25 anos (Linha 1.0.8)",  "esocial_24": "01.08.001"},
    "141-78-6":   {"agente": "Acetato de Etila",         "nr15_lt": "310 ppm ou 1090 mg/m3",         "nr09_acao": "155 ppm ou 545 mg/m3",         "nr07_ibe": "Avaliacao Clinica",                  "dec_3048": "Nao Enquadrado",         "esocial_24": "09.01.001"},
    "13463-67-7": {"agente": "Dioxido de Titanio",       "nr15_lt": "Ausente NR-15 (ACGIH: 10 mg/m3)", "nr09_acao": "5 mg/m3 (Ref. ACGIH)",       "nr07_ibe": "Avaliacao Clinica / Raio-X",         "dec_3048": "Nao Enquadrado",         "esocial_24": "09.01.001"},
    "1333-86-4":  {"agente": "Negro de Fumo (Carbon Black)", "nr15_lt": "Ausente NR-15 (ACGIH: 3 mg/m3)", "nr09_acao": "1,5 mg/m3 (Ref. ACGIH)", "nr07_ibe": "Avaliacao Clinica / Espirometria",   "dec_3048": "Avaliar Anexo IV",       "esocial_24": "01.01.000"},
    "8052-41-3":  {"agente": "Aguarras Mineral",         "nr15_lt": "Ausente NR-15 (ACGIH: 100 ppm)", "nr09_acao": "50 ppm (Ref. ACGIH)",         "nr07_ibe": "Avaliacao Clinica",                  "dec_3048": "Nao Enquadrado",         "esocial_24": "09.01.001"},
    "1309-37-1":  {"agente": "Oxido de Ferro",           "nr15_lt": "Ausente NR-15 (ACGIH: 5 mg/m3)", "nr09_acao": "2,5 mg/m3 (Ref. ACGIH)",     "nr07_ibe": "Raio-X (OIT)",                       "dec_3048": "Nao Enquadrado",         "esocial_24": "09.01.001"},
    "136-51-6":   {"agente": "Octoato de Calcio",        "nr15_lt": "Nao Estabelecido",               "nr09_acao": "Nao Estabelecido",             "nr07_ibe": "Avaliacao Clinica",                  "dec_3048": "Nao Enquadrado",         "esocial_24": "09.01.001"},
}

DICIONARIO_FIS_BIO = {
    "Ruido Continuo/Intermitente": {
        "agente": "Ruido Continuo ou Intermitente", "nr15_lt": "85 dB(A)", "nr09_acao": "80 dB(A)",
        "nr07_ibe": "Audiometria", "dec_3048": "25 anos (Linha 2.0.1)", "esocial_24": "02.01.001",
        "perigo": "Exposicao a niveis elevados de pressao sonora", "sev": 3, "epi": "Protetor Auditivo (Atenuacao adequada)",
    },
    "Vibracao de Maos e Bracos (VMB)": {
        "agente": "Vibracao de Maos e Bracos (VMB)", "nr15_lt": "5,0 m/s2", "nr09_acao": "2,5 m/s2",
        "nr07_ibe": "Avaliacao Clinica", "dec_3048": "25 anos (Linha 2.0.2)", "esocial_24": "02.01.002",
        "perigo": "Transmissao de energia mecanica para o sistema mao-braco", "sev": 3, "epi": "Luvas antivibracao / Revezamento",
    },
    "Vibracao de Corpo Inteiro (VCI)": {
        "agente": "Vibracao de Corpo Inteiro (VCI)", "nr15_lt": "1,1 m/s2 ou 21,0 m/s1.75", "nr09_acao": "0,5 m/s2 ou 9,1 m/s1.75",
        "nr07_ibe": "Avaliacao Clinica", "dec_3048": "25 anos (Linha 2.0.2)", "esocial_24": "02.01.003",
        "perigo": "Transmissao de energia mecanica para o corpo inteiro", "sev": 3, "epi": "Assentos com amortecimento / Revezamento",
    },
    "Biologico: Esgoto / Fossas": {
        "agente": "Microorganismos - Esgoto / Fossas", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo",
        "nr07_ibe": "Exames Clinicos / Vacinas", "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.005",
        "perigo": "Exposicao a agentes biologicos infectocontagiosos", "sev": 4, "epi": "Luvas, Botas de PVC, Protecao facial",
    },
    "Biologico: Lixo Urbano": {
        "agente": "Microorganismos - Lixo Urbano", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo",
        "nr07_ibe": "Exames Clinicos / Vacinas", "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.007",
        "perigo": "Contato com residuos e agentes biologicos", "sev": 4, "epi": "Luvas anticorte, Botas, Uniforme impermeavel",
    },
    "Biologico: Estab. Saude": {
        "agente": "Microorganismos - Area da Saude", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo",
        "nr07_ibe": "Exames Clinicos / Vacinas", "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.001",
        "perigo": "Exposicao a patogenos em ambiente de saude", "sev": 4, "epi": "Luvas de procedimento, Mascara, Avental",
    },
}

# ── HTML ──────────────────────────────────────────────────────────

def gerar_html_anexo(resultados_pgr, resultados_medicos):
    css = """<style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
      :root{--g800:#084D22;--g700:#0F823B;--g600:#1AA04B;--g100:#D9F585;--g900:#042B13;
            --white:#fff;--gray50:#F8F9FA;--gray200:#E9ECEF;--gray600:#6C757D;--gray900:#212529;}
      body{font-family:'Inter',sans-serif;font-size:10pt;color:var(--gray900);background:var(--white);padding:20px;}
      .header{background:var(--g800);color:var(--white);padding:14px 20px;font-size:13pt;font-weight:700;
              margin-bottom:20px;border-radius:2px;text-align:center;text-transform:uppercase;border-bottom:4px solid var(--g100);}
      .aviso{background:var(--g100);color:var(--g900);padding:10px;margin-bottom:15px;border-radius:4px;
             font-size:9pt;border:1px solid var(--g600);}
      .card{border:1px solid var(--g800);border-radius:4px;margin-bottom:20px;page-break-inside:avoid;}
      .card-head{background:var(--g800);padding:10px 16px;font-weight:700;color:var(--white);font-size:10pt;
                 border-bottom:2px solid var(--g100);}
      .card-body{padding:12px 16px;}
      table{width:100%;border-collapse:collapse;font-size:8.5pt;margin:8px 0;}
      th{background:var(--g700);color:var(--white);padding:8px;text-align:left;}
      td{padding:5px 10px;border:1px solid var(--gray200);}
      td:first-child{font-weight:600;color:var(--g800);width:120px;background:var(--gray50);}
      .ri{background:#FEE2E2!important;color:#991B1B;font-weight:bold;}
      .rs{background:#FFF3E0!important;color:#E8590C;font-weight:bold;}
      .ra{background:#FEF08A!important;color:#854D0E;font-weight:bold;text-align:center;}
      h4{color:var(--g800);margin:15px 0 5px;font-size:9.5pt;text-transform:uppercase;}
    </style>"""

    html = f"<!DOCTYPE html><html lang='pt-BR'><head><meta charset='UTF-8'>{css}</head><body>"
    html += "<div class='header'>ANEXO I - INVENTARIO DE RISCOS E ENQUADRAMENTO PREVIDENCIARIO</div>"
    html += "<div class='aviso'><strong>Atencao Equipe Tecnica:</strong> Documento gerado automaticamente. Consolide FISPQs e avaliacoes de campo.</div>"

    df_pgr = pd.DataFrame(resultados_pgr)
    df_med = pd.DataFrame(resultados_medicos)
    ghes_pgr = df_pgr["GHE"].unique().tolist() if not df_pgr.empty else []
    ghes_med = df_med["GHE"].unique().tolist() if not df_med.empty else []
    ghes = sorted(set(ghes_pgr + ghes_med))

    for ghe in ghes:
        html += f"<div class='card'><div class='card-head' contenteditable='true'>{ghe}</div><div class='card-body'>"

        pgr_ghe = df_pgr[df_pgr["GHE"] == ghe] if not df_pgr.empty else pd.DataFrame()
        if not pgr_ghe.empty:
            html += "<h4>Inventario de Risco (NR-01)</h4><table><thead><tr><th>Origem/FISPQ</th><th>Perigo</th><th>Sev.</th><th>Prob.</th><th>Nivel de Risco</th><th>EPI (NR-06)</th></tr></thead><tbody>"
            for _, r in pgr_ghe.iterrows():
                cl = "ri" if r["NIVEL DE RISCO"] == "INTOLERAVEL" else "rs" if r["NIVEL DE RISCO"] == "SUBSTANCIAL" else ""
                html += (f"<tr><td contenteditable='true'>{r['Arquivo Origem']}</td>"
                         f"<td contenteditable='true'>{r['Codigo GHS']} {r['Perigo Identificado']}</td>"
                         f"<td contenteditable='true'>{r['Severidade']}</td>"
                         f"<td contenteditable='true'>{r['Probabilidade']}</td>"
                         f"<td contenteditable='true' class='{cl}'>{r['NIVEL DE RISCO']}</td>"
                         f"<td contenteditable='true'>{r['EPI (NR-06)']}</td></tr>")
            html += "</tbody></table>"

        med_ghe = df_med[df_med["GHE"] == ghe] if not df_med.empty else pd.DataFrame()
        if not med_ghe.empty:
            html += "<h4>Diretrizes Medicas e Previdenciarias (NR-07/NR-09/NR-15/Dec 3.048/eSocial)</h4>"
            html += "<table><thead><tr><th>CAS</th><th>Agente</th><th>Lim. Tol. (NR-15)</th><th>Nivel Acao (NR-09)</th><th>IBE (NR-07)</th><th>Dec 3.048</th><th>eSocial (Tab 24)</th></tr></thead><tbody>"
            for _, r in med_ghe.iterrows():
                cl = "ra" if "NAO MAPEADO" in str(r["Agente Quimico"]) else ""
                html += (f"<tr><td contenteditable='true'>{r['N CAS']}</td>"
                         f"<td contenteditable='true' class='{cl}'>{r['Agente Quimico']}</td>"
                         f"<td contenteditable='true' class='{cl}'>{r['Lim. Tolerancia (NR-15)']}</td>"
                         f"<td contenteditable='true' class='{cl}'>{r['Nivel de Acao (NR-09)']}</td>"
                         f"<td contenteditable='true' class='{cl}'>{r['IBE (NR-07)']}</td>"
                         f"<td contenteditable='true' class='{cl}'>{r['Dec 3048']}</td>"
                         f"<td contenteditable='true' class='{cl}'>{r['eSocial']}</td></tr>")
            html += "</tbody></table>"

        html += "</div></div>"

    html += "</body></html>"
    return html


# ── Render principal ─────────────────────────────────────────────

def render_engenharia():
    try:
        CHAVE_API = st.secrets["CHAVE_GOOGLE"]
    except Exception:
        CHAVE_API = None

    st.title("Modulo de Engenharia: FISPQ / FDS -> PGR")
    st.info("Versao do Motor: 26.0 (Cerebro Google Gemini - Auto-Discovery Engine)")

    # ── Upload FISPQs ────────────────────────────────────────────
    st.markdown("### 1. Upload das FISPQs (Agentes Quimicos)")
    arquivos_fispq = st.file_uploader("Insira as FISPQs em PDF", type=["pdf"], accept_multiple_files=True)
    textos_pdfs = {}

    df_editado     = pd.DataFrame()
    ghe_opcoes     = ["Nenhum GHE definido"]

    if arquivos_fispq:
        with st.spinner("Lendo o conteudo dos PDFs..."):
            for arq in arquivos_fispq:
                with pdfplumber.open(arq) as pdf:
                    texto = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                    textos_pdfs[arq.name] = texto

        st.markdown("---")
        st.markdown("### 2. Definicao de GHEs Quimicos")

        nomes = [a.name for a in arquivos_fispq]
        dados_iniciais = [{"GHE": "GHE 01 - Digite a Funcao", "Arquivo FISPQ": n, "Probabilidade": 3} for n in nomes]

        df_editado = st.data_editor(
            pd.DataFrame(dados_iniciais), num_rows="dynamic",
            column_config={
                "GHE":           st.column_config.TextColumn("Nome do GHE", required=True),
                "Arquivo FISPQ": st.column_config.SelectboxColumn("Arquivo (FISPQ)", options=nomes, required=True),
                "Probabilidade": st.column_config.NumberColumn("Prob.", min_value=1, max_value=5, required=True),
            },
            use_container_width=True,
        )
        ghe_opcoes = df_editado["GHE"].unique().tolist() if not df_editado.empty else ["Nenhum GHE definido"]

    # ── Agentes Fisicos e Biologicos ─────────────────────────────
    st.markdown("---")
    st.markdown("### 3. Avaliacoes de Campo: Agentes Fisicos e Biologicos (Opcional)")

    agentes_opcoes     = list(DICIONARIO_FIS_BIO.keys())
    df_fis_bio_inicial = pd.DataFrame([{"GHE": ghe_opcoes[0], "Agente": agentes_opcoes[0], "Probabilidade": 3}])

    df_fis_bio = st.data_editor(
        df_fis_bio_inicial, num_rows="dynamic",
        column_config={
            "GHE":           st.column_config.SelectboxColumn("GHE de Destino",         options=ghe_opcoes,    required=True),
            "Agente":        st.column_config.SelectboxColumn("Agente Fisico/Biologico", options=agentes_opcoes, required=True),
            "Probabilidade": st.column_config.NumberColumn("Prob.", min_value=1, max_value=5, required=True),
        },
        use_container_width=True,
    )

    # ── Processar ────────────────────────────────────────────────
    st.markdown("---")
    if st.button("Processar GHEs e Gerar Relatorio Completo", use_container_width=True):
        with st.spinner("Consolidando avaliacoes..."):
            resultados_pgr      = []
            resultados_medicos  = []

            # Quimico
            if not df_editado.empty:
                for _, row in df_editado.iterrows():
                    nome_ghe = row["GHE"]
                    nome_arq = row["Arquivo FISPQ"]
                    v_prob   = int(row["Probabilidade"])

                    if nome_arq in textos_pdfs:
                        texto = textos_pdfs[nome_arq]

                        for cas in set(re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', texto)):
                            d = DICIONARIO_CAS.get(cas, {
                                "agente": "AGENTE NAO MAPEADO", "nr15_lt": "REVISAO DA ENGENHARIA",
                                "nr09_acao": "REVISAO DA ENGENHARIA", "nr07_ibe": "REVISAO DA ENGENHARIA",
                                "dec_3048": "REVISAO DA ENGENHARIA", "esocial_24": "REVISAO DA ENGENHARIA",
                            })
                            resultados_medicos.append({
                                "GHE": nome_ghe, "Arquivo Origem": nome_arq, "N CAS": cas,
                                "Agente Quimico": d["agente"], "Lim. Tolerancia (NR-15)": d["nr15_lt"],
                                "Nivel de Acao (NR-09)": d["nr09_acao"], "IBE (NR-07)": d.get("nr07_ibe", "N/A"),
                                "Dec 3048": d.get("dec_3048", "Nao Enquadrado"), "eSocial": d.get("esocial_24", "09.01.001"),
                            })

                        for cod in set(re.findall(r'H\d{3}', texto)):
                            if cod in DICIONARIO_H:
                                d         = DICIONARIO_H[cod]
                                s_val     = d["sev"]
                                nivel     = MATRIZ_OFICIAL.get((s_val, v_prob), "N/A")
                                resultados_pgr.append({
                                    "GHE": nome_ghe, "Arquivo Origem": nome_arq, "Codigo GHS": cod,
                                    "Perigo Identificado": d["desc"], "Severidade": TEXTO_SEV.get(s_val, str(s_val)),
                                    "Probabilidade": str(v_prob), "NIVEL DE RISCO": nivel,
                                    "Acao Requerida": ACOES_REQUERIDAS.get(nivel, "Manual"), "EPI (NR-06)": d["epi"],
                                })

            # Fisico / Biologico
            if not df_fis_bio.empty and df_fis_bio["GHE"].iloc[0] != "Nenhum GHE definido":
                for _, row in df_fis_bio.iterrows():
                    nome_ghe  = row["GHE"]
                    nome_ag   = row["Agente"]
                    v_prob    = int(row["Probabilidade"])
                    if nome_ag in DICIONARIO_FIS_BIO:
                        d     = DICIONARIO_FIS_BIO[nome_ag]
                        nivel = MATRIZ_OFICIAL.get((d["sev"], v_prob), "N/A")
                        resultados_medicos.append({
                            "GHE": nome_ghe, "Arquivo Origem": "Campo", "N CAS": "-",
                            "Agente Quimico": d["agente"], "Lim. Tolerancia (NR-15)": d["nr15_lt"],
                            "Nivel de Acao (NR-09)": d["nr09_acao"], "IBE (NR-07)": d["nr07_ibe"],
                            "Dec 3048": d["dec_3048"], "eSocial": d["esocial_24"],
                        })
                        resultados_pgr.append({
                            "GHE": nome_ghe, "Arquivo Origem": "Campo", "Codigo GHS": "-",
                            "Perigo Identificado": d["perigo"], "Severidade": TEXTO_SEV.get(d["sev"], str(d["sev"])),
                            "Probabilidade": str(v_prob), "NIVEL DE RISCO": nivel,
                            "Acao Requerida": ACOES_REQUERIDAS.get(nivel, "Manual"), "EPI (NR-06)": d["epi"],
                        })

            if resultados_pgr or resultados_medicos:
                html_final = gerar_html_anexo(resultados_pgr, resultados_medicos)
                st.success("Relatorio Consolidado Gerado!")
                st.session_state["ultimo_html_eng"] = html_final
            else:
                st.warning("Nenhum dado processado. Verifique os PDFs e GHEs.")

    # ── Salvar e Download ─────────────────────────────────────────
    if "ultimo_html_eng" in st.session_state:
        st.markdown("### Salvar Projeto no Banco de Dados")
        col1, col2 = st.columns([3, 1])
        with col1:
            nome_proj = st.text_input("Nome da Empresa / Projeto:")
        with col2:
            st.write("")
            st.write("")
            if st.button("Gravar no Historico", use_container_width=True):
                if nome_proj:
                    try:
                        from config.db import salvar_historico
                        salvar_historico(nome_proj, st.session_state["ultimo_html_eng"])
                        st.success("Projeto salvo com sucesso!")
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")
                else:
                    st.warning("Digite um nome para o projeto.")

        aba1, aba2 = st.tabs(["Pre-visualizar Documento", "Baixar em Word (.doc)"])
        with aba1:
            components.html(st.session_state["ultimo_html_eng"], height=700, scrolling=True)
        with aba2:
            st.download_button(
                label="Baixar Relatorio em Word",
                data=st.session_state["ultimo_html_eng"].encode("utf-8"),
                file_name="Relatorio_Seconci.doc",
                mime="application/msword",
                use_container_width=True,
            )

    # ── Consultoria IA ────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### Consultoria de Higiene Ocupacional (Google Gemini)")

    if not textos_pdfs:
        st.info("Faca o upload de uma FISPQ no Passo 1 para ativar o chat de consultoria da IA.")
    elif not CHAVE_API:
        st.warning("Chave de API nao configurada em st.secrets['CHAVE_GOOGLE'].")
    else:
        st.success("Conectado ao Google Gemini!")
        fispq_sel   = st.selectbox("Selecione a FISPQ para analise:", list(textos_pdfs.keys()))
        pergunta_ia = st.text_input("Sua pergunta tecnica:")

        if st.button("Perguntar ao Gemini"):
            if pergunta_ia:
                with st.spinner("O Gemini esta analisando..."):
                    try:
                        resumo  = textos_pdfs[fispq_sel][:15000]
                        prompt  = (
                            "Voce e um Engenheiro de Seguranca do Trabalho Senior do Seconci-GO. "
                            "Responda com precisao baseando-se EXCLUSIVAMENTE nas NRs brasileiras e no texto abaixo. "
                            "NUNCA invente informacoes.\n\n"
                            f"Documento (FISPQ): {resumo}\n\nPergunta: {pergunta_ia}"
                        )

                        url_lista   = f"https://generativelanguage.googleapis.com/v1beta/models?key={CHAVE_API}"
                        resp_lista  = requests.get(url_lista)

                        if resp_lista.status_code == 200:
                            modelos        = resp_lista.json().get("models", [])
                            modelos_texto  = [m["name"] for m in modelos if "generateContent" in m.get("supportedGenerationMethods", [])]
                            modelo         = None
                            for pref in ["models/gemini-1.5-flash", "models/gemini-1.5-pro", "models/gemini-pro", "models/gemini-1.0-pro"]:
                                if pref in modelos_texto:
                                    modelo = pref
                                    break
                            if not modelo and modelos_texto:
                                modelo = modelos_texto[0]

                            if modelo:
                                url_gen  = f"https://generativelanguage.googleapis.com/v1beta/{modelo}:generateContent?key={CHAVE_API}"
                                payload  = {"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"temperature": 0.1}}
                                resp_gen = requests.post(url_gen, headers={"Content-Type": "application/json"}, json=payload)
                                if resp_gen.status_code == 200:
                                    txt = resp_gen.json()["candidates"][0]["content"]["parts"][0]["text"]
                                    st.success(f"Analise concluida! (Motor: {modelo.split('/')[-1]})")
                                    st.write(txt)
                                else:
                                    st.error(f"Erro na geracao: {resp_gen.text}")
                            else:
                                st.error("Nenhum modelo disponivel para esta chave de API.")
                        else:
                            st.error(f"Chave de API bloqueada. Status: {resp_lista.status_code}")
                    except Exception as e:
                        st.error(f"Falha na conexao: {e}")
            else:
                st.warning("Digite uma pergunta para a IA.")

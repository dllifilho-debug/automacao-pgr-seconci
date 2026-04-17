import streamlit as st
import streamlit.components.v1 as components
import pdfplumber
import re
import pandas as pd
import sqlite3
import io
from datetime import datetime
import os
import requests
import json
import base64

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E CSS (UI/UX)
# ==========================================
st.set_page_config(
    page_title="Automação SST - Seconci GO",
    layout="wide",
    page_icon="🛡"
)

css_personalizado = """
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .stButton > button {
        background-color: #084D22; color: white; border-radius: 8px;
        border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        transition: all 0.3s ease; font-weight: 600; padding: 0.5rem 1rem;
    }
    .stButton > button:hover {
        background-color: #1AA04B; color: white;
        box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        transform: translateY(-2px); border-color: #1AA04B;
    }
    h1, h2, h3 { color: #084D22 !important; }
    [data-testid="stSidebar"] {
        background-color: #F4F8F5; border-right: 1px solid #E0ECE4;
    }
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #1AA04B; border-radius: 12px;
        background-color: #FAFFFA; transition: all 0.3s ease;
    }
    [data-testid="stFileUploadDropzone"]:hover {
        background-color: #F0FDF4; border-color: #084D22;
    }
    .stAlert { border-radius: 8px; border-left: 5px solid #084D22; }
    [data-testid="stDataFrame"] {
        border: 1px solid #E0ECE4; border-radius: 8px; overflow: hidden;
    }
</style>
"""
st.markdown(css_personalizado, unsafe_allow_html=True)

# ==========================================
# CONFIGURAÇÃO DA IA (CONEXÃO DIRETA REST)
# ==========================================
CHAVE_API_GOOGLE = str(st.secrets["CHAVE_API_GOOGLE"]).strip().replace('"', '').replace("'", "")

# ==========================================
# INICIALIZAÇÃO DO BANCO DE DADOS (SQLITE)
# ==========================================
def init_db():
    conn = sqlite3.connect('seconci_banco_dados.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS historico_laudos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_projeto TEXT,
            data_salvamento TEXT,
            html_relatorio TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS dicionario_dinamico (
            cas TEXT PRIMARY KEY,
            agente TEXT,
            nr15_lt TEXT,
            nr09_acao TEXT,
            nr07_ibe TEXT,
            dec_3048 TEXT,
            esocial_24 TEXT,
            data_aprendizado TEXT,
            fonte TEXT
        )
    ''')
    conn.commit()
    conn.close()

init_db()

# ==========================================
# BARRA LATERAL (MENU E HISTÓRICO)
# ==========================================
if os.path.exists("logo.png"):
    st.sidebar.image("logo.png", use_column_width=True)
elif os.path.exists("logo.jpg"):
    st.sidebar.image("logo.jpg", use_column_width=True)
else:
    st.sidebar.markdown(
        "<h2 style='text-align: center; color: #084D22;'>SECONCI-GO</h2>",
        unsafe_allow_html=True
    )

st.sidebar.markdown("---")
st.sidebar.title("🧩 Módulos do Sistema")
modulo_selecionado = st.sidebar.radio("Selecione a funcionalidade:", [
    "🏠 Dashboard",
    "1️⃣ Engenharia: FISPQ / FDS ➡ PGR",
    "2️⃣ Medicina: PGR ➡ PCMSO"
])

st.sidebar.markdown("---")
st.sidebar.title("📂 Histórico de Laudos")
conn = sqlite3.connect('seconci_banco_dados.db')
df_historico = pd.read_sql_query(
    "SELECT id, nome_projeto, data_salvamento FROM historico_laudos ORDER BY id DESC",
    conn
)
conn.close()

historico_selecionado = None
if not df_historico.empty:
    opcoes_historico = ["Selecione um projeto salvo..."] + [
        f"{row['id']} - {row['nome_projeto']} ({row['data_salvamento']})"
        for _, row in df_historico.iterrows()
    ]
    selecao = st.sidebar.selectbox("Carregar projeto antigo:", opcoes_historico)
    if selecao != "Selecione um projeto salvo...":
        id_selecionado = int(selecao.split(" - ")[0])
        conn = sqlite3.connect('seconci_banco_dados.db')
        cursor = conn.cursor()
        cursor.execute(
            "SELECT html_relatorio FROM historico_laudos WHERE id = ?",
            (id_selecionado,)
        )
        historico_selecionado = cursor.fetchone()[0]
        conn.close()
        st.sidebar.success("✅ Projeto carregado.")
else:
    st.sidebar.write("Nenhum projeto salvo ainda.")

# ==========================================
# DICIONÁRIOS FASE 1 (ENGENHARIA E QUÍMICA)
# ==========================================
dicionario_h = {
    "H315": {"desc": "Provoca irritação à pele",
             "sev": 1, "epi": "Luvas de proteção e vestimenta"},
    "H319": {"desc": "Provoca irritação ocular grave",
             "sev": 1, "epi": "Óculos de proteção contra respingos"},
    "H336": {"desc": "Pode provocar sonolência ou vertigem",
             "sev": 1, "epi": "Local ventilado; máscara se necessário"},
    "H317": {"desc": "Reações alérgicas na pele",
             "sev": 2, "epi": "Luvas (nitrílica/PVC) e manga longa"},
    "H335": {"desc": "Irritação das vias respiratórias",
             "sev": 2, "epi": "Proteção respiratória (Filtro específico)"},
    "H302": {"desc": "Nocivo em caso de ingestão",
             "sev": 2, "epi": "Higiene rigorosa; luvas adequadas"},
    "H312": {"desc": "Nocivo em contato com a pele",
             "sev": 2, "epi": "Luvas impermeáveis e avental"},
    "H332": {"desc": "Nocivo se inalado",
             "sev": 2, "epi": "Máscara respiratória adequada (PFF2/VO)"},
    "H314": {"desc": "Queimadura severa à pele e dano ocular",
             "sev": 4, "epi": "Traje químico, luvas longas e botas"},
    "H318": {"desc": "Lesões oculares graves",
             "sev": 3, "epi": "Óculos ampla visão / protetor facial"},
    "H301": {"desc": "Tóxico em caso de ingestão",
             "sev": 4, "epi": "Higiene rigorosa; luvas"},
    "H311": {"desc": "Tóxico em contato com a pele",
             "sev": 4, "epi": "Luvas, avental impermeável e botas"},
    "H331": {"desc": "Tóxico se inalado",
             "sev": 4, "epi": "Respirador facial inteiro"},
    "H334": {"desc": "Sintomas de asma ou dificuldades respiratórias",
             "sev": 4, "epi": "Respirador facial inteiro (Filtro P3)"},
    "H372": {"desc": "Danos aos órgãos (exposição prolongada/repetida)",
             "sev": 4, "epi": "Proteção respiratória e dérmica estrita"},
    "H373": {"desc": "Pode provocar danos aos órgãos (exp. repetida)",
             "sev": 4, "epi": "Avaliar via; EPIs combinados obrigatórios"},
    "H300": {"desc": "Fatal em caso de ingestão",
             "sev": 5, "epi": "Isolamento total; Higiene extrema"},
    "H310": {"desc": "Fatal em contato com a pele",
             "sev": 5, "epi": "Traje encapsulado nível A/B"},
    "H330": {"desc": "Fatal se inalado",
             "sev": 5, "epi": "Equipamento de Respiração Autônoma (EPR)"},
    "H340": {"desc": "Pode provocar defeitos genéticos",
             "sev": 5, "epi": "Isolamento; Traje químico e EPR"},
    "H350": {"desc": "Pode provocar câncer",
             "sev": 5, "epi": "Isolamento; Traje químico e EPR completo"},
    "H351": {"desc": "Suspeito de provocar câncer",
             "sev": 4, "epi": "Proteção respiratória e dérmica máxima"},
    "H360": {"desc": "Pode prejudicar a fertilidade ou o feto",
             "sev": 5, "epi": "Afastamento de gestantes; EPI máximo"},
    "H370": {"desc": "Provoca danos aos órgãos",
             "sev": 5, "epi": "EPI máximo conforme via de exposição"},
    "H341": {"desc": "Suspeito de provocar defeitos genéticos",
             "sev": 4, "epi": "Proteção respiratória e dérmica máxima"},
    "H361": {"desc": "Suspeito de prejudicar a fertilidade ou o feto",
             "sev": 4, "epi": "Afastamento de gestantes; EPI máximo"},
}

matriz_oficial = {
    (1,1): "TRIVIAL",     (1,2): "TRIVIAL",     (1,3): "TOLERÁVEL",
    (1,4): "TOLERÁVEL",   (1,5): "MODERADO",
    (2,1): "TRIVIAL",     (2,2): "TOLERÁVEL",   (2,3): "MODERADO",
    (2,4): "MODERADO",    (2,5): "SUBSTANCIAL",
    (3,1): "TOLERÁVEL",   (3,2): "TOLERÁVEL",   (3,3): "MODERADO",
    (3,4): "SUBSTANCIAL", (3,5): "SUBSTANCIAL",
    (4,1): "TOLERÁVEL",   (4,2): "MODERADO",    (4,3): "SUBSTANCIAL",
    (4,4): "INTOLERÁVEL", (4,5): "INTOLERÁVEL",
    (5,1): "MODERADO",    (5,2): "MODERADO",    (5,3): "SUBSTANCIAL",
    (5,4): "INTOLERÁVEL", (5,5): "INTOLERÁVEL",
}

acoes_requeridas = {
    "TRIVIAL":      "Manter controles existentes; monitoramento periódico.",
    "TOLERÁVEL":    "Manter controles. Considerar melhorias.",
    "MODERADO":     "Implantar controles. EPI e monitoramento PCMSO.",
    "SUBSTANCIAL":  "Trabalho não deve iniciar sem redução do risco.",
    "INTOLERÁVEL":  "TRABALHO PROIBIDO. Risco grave e iminente.",
}

texto_sev = {1: "1-LEVE", 2: "2-BAIXA", 3: "3-MODERADA", 4: "4-ALTA", 5: "5-EXTREMA"}

# ==========================================
# CORREÇÃO 3: Conjuntos de códigos H críticos
# ==========================================
CODIGOS_CARCINOGENICOS = {"H340", "H350", "H360"}
CODIGOS_SUSPEITOS      = {"H341", "H351", "H361"}

CAS_CARCINOGENICOS_CONHECIDOS = {
    "71-43-2",    # Benzeno — Grupo 1 IARC
    "79-01-6",    # Tricloroetileno — Grupo 1 IARC
    "1332-21-4",  # Asbesto/Amianto — Grupo 1 IARC
    "136-52-7",   # Octoato de Cobalto — Cat. 1B
    "96-29-7",    # MEKO — Cat. 2 UE
}

# ==========================================
# BANCO LOCAL — Agentes comuns da construção civil (Camada 1)
# ==========================================
dicionario_cas = {
    "108-88-3": {
        "agente": "Tolueno",
        "nr15_lt": "78 ppm ou 290 mg/m³",
        "nr09_acao": "39 ppm ou 145 mg/m³",
        "nr07_ibe": "o-Cresol ou Ácido Hipúrico",
        "dec_3048": "25 anos (Linha 1.0.19)",
        "esocial_24": "01.19.036",
    },
    "1330-20-7": {
        "agente": "Xileno",
        "nr15_lt": "78 ppm ou 340 mg/m³",
        "nr09_acao": "39 ppm ou 170 mg/m³",
        "nr07_ibe": "Ácidos Metilhipúricos",
        "dec_3048": "25 anos (Linha 1.0.19)",
        "esocial_24": "01.19.036",
    },
    "71-43-2": {
        "agente": "Benzeno",
        "nr15_lt": "VRT-MPT (Cancerígeno)",
        "nr09_acao": "Avaliação Qualitativa",
        "nr07_ibe": "Ácido trans,trans-mucônico",
        "dec_3048": "25 anos (Linha 1.0.3)",
        "esocial_24": "01.01.006",
    },
    "67-64-1": {
        "agente": "Acetona",
        "nr15_lt": "780 ppm ou 1870 mg/m³",
        "nr09_acao": "390 ppm ou 935 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "64-17-5": {
        "agente": "Etanol (Álcool Etílico)",
        "nr15_lt": "780 ppm ou 1480 mg/m³",
        "nr09_acao": "390 ppm ou 740 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "78-93-3": {
        "agente": "Metiletilcetona (MEK)",
        "nr15_lt": "155 ppm ou 460 mg/m³",
        "nr09_acao": "77.5 ppm ou 230 mg/m³",
        "nr07_ibe": "MEK na Urina",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "110-54-3": {
        "agente": "n-Hexano",
        "nr15_lt": "50 ppm ou 176 mg/m³",
        "nr09_acao": "25 ppm ou 88 mg/m³",
        "nr07_ibe": "2,5-Hexanodiona",
        "dec_3048": "25 anos (Linha 1.0.19)",
        "esocial_24": "01.19.014",
    },
    "14808-60-7": {
        "agente": "Sílica Cristalina (Quartzo)",
        "nr15_lt": "Anexo 12",
        "nr09_acao": "50% do L.T.",
        "nr07_ibe": "Raio-X (OIT) e Espirometria",
        "dec_3048": "25 anos (Linha 1.0.18)",
        "esocial_24": "01.18.001",
    },
    "1332-21-4": {
        "agente": "Asbesto / Amianto",
        "nr15_lt": "0,1 f/cm³",
        "nr09_acao": "0,05 f/cm³",
        "nr07_ibe": "Raio-X (OIT) e Espirometria",
        "dec_3048": "20 anos (Linha 1.0.2)",
        "esocial_24": "01.02.001",
    },
    "7439-92-1": {
        "agente": "Chumbo (Fumos)",
        "nr15_lt": "0,1 mg/m³",
        "nr09_acao": "0,05 mg/m³",
        "nr07_ibe": "Chumbo no sangue e ALA-U",
        "dec_3048": "25 anos (Linha 1.0.8)",
        "esocial_24": "01.08.001",
    },
    "65997-15-1": {
        "agente": "Cimento Portland",
        "nr15_lt": "Anexo 12 (Poeiras)",
        "nr09_acao": "50% do L.T.",
        "nr07_ibe": "Raio-X (OIT) e Espirometria",
        "dec_3048": "25 anos (Linha 1.0.18)",
        "esocial_24": "01.18.001",
    },
    "1305-62-0": {
        "agente": "Hidróxido de Cálcio",
        "nr15_lt": "5 mg/m³",
        "nr09_acao": "2,5 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "7664-38-2": {
        "agente": "Ácido Fosfórico",
        "nr15_lt": "1 mg/m³",
        "nr09_acao": "0,5 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "7647-01-0": {
        "agente": "Ácido Clorídrico",
        "nr15_lt": "5 ppm ou 7 mg/m³",
        "nr09_acao": "2,5 ppm ou 3,5 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "7664-93-9": {
        "agente": "Ácido Sulfúrico (Névoas)",
        "nr15_lt": "0,2 mg/m³ (névoas ácidas)",
        "nr09_acao": "0,1 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "01.01.006",
    },
    "12042-78-3": {
        "agente": "Aluminato de Cálcio",
        "nr15_lt": "Anexo 12 (Poeiras Insolúveis)",
        "nr09_acao": "50% do L.T.",
        "nr07_ibe": "Raio-X (OIT) e Espirometria",
        "dec_3048": "25 anos (Linha 1.0.18)",
        "esocial_24": "01.18.001",
    },
    "1309-48-4": {
        "agente": "Óxido de Magnésio (Fumos)",
        "nr15_lt": "10 mg/m³ (fumos)",
        "nr09_acao": "5 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "1305-78-8": {
        "agente": "Óxido de Cálcio (Cal Virgem)",
        "nr15_lt": "2 mg/m³",
        "nr09_acao": "1 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "12168-85-3": {
        "agente": "Silicato de Cálcio",
        "nr15_lt": "Anexo 12 (Poeiras)",
        "nr09_acao": "50% do L.T.",
        "nr07_ibe": "Raio-X (OIT) e Espirometria",
        "dec_3048": "25 anos (Linha 1.0.18)",
        "esocial_24": "01.18.001",
    },
    "1317-65-3": {
        "agente": "Carbonato de Cálcio",
        "nr15_lt": "10 mg/m³ (poeira total)",
        "nr09_acao": "5 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "12068-35-8": {
        "agente": "Silicato Tricálcico (C3S)",
        "nr15_lt": "Anexo 12 (Poeiras)",
        "nr09_acao": "50% do L.T.",
        "nr07_ibe": "Raio-X (OIT) e Espirometria",
        "dec_3048": "25 anos (Linha 1.0.18)",
        "esocial_24": "01.18.001",
    },
    "10034-77-2": {
        "agente": "Sulfato de Cálcio Hemihidratado (Gesso)",
        "nr15_lt": "10 mg/m³",
        "nr09_acao": "5 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "68334-30-5": {
        "agente": "Asfalto / Betume de Petróleo",
        "nr15_lt": "5 mg/m³ (névoas)",
        "nr09_acao": "2,5 mg/m³",
        "nr07_ibe": "Avaliação Clínica / Dermatológica",
        "dec_3048": "25 anos (Linha 1.0.19)",
        "esocial_24": "01.19.036",
    },
    "112-80-1": {
        "agente": "Ácido Oleico",
        "nr15_lt": "Não estabelecido (NR-15)",
        "nr09_acao": "Avaliar NR-09",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "55965-84-9": {
        "agente": "Mistura MIT/CMIT (Biocida)",
        "nr15_lt": "Avaliação Qualitativa",
        "nr09_acao": "Avaliação Qualitativa",
        "nr07_ibe": "Avaliação Dermatológica / Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "52-51-7": {
        "agente": "Bronopol (Conservante)",
        "nr15_lt": "Avaliação Qualitativa",
        "nr09_acao": "Avaliação Qualitativa",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "533-74-4": {
        "agente": "Dazomet (Fungicida)",
        "nr15_lt": "Avaliação Qualitativa",
        "nr09_acao": "Avaliação Qualitativa",
        "nr07_ibe": "Avaliação Clínica / Neurológica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    # ── CORREÇÃO 2: CAS críticos adicionados na auditoria ──────────
    "79-01-6": {
        "agente": "Tricloroetileno",
        "nr15_lt": "Anexo 13-A (Cancerígeno Ocupacional)",
        "nr09_acao": "Avaliação Qualitativa (Cancerígeno)",
        "nr07_ibe": "Ácido Tricloroacético (TCA) e Tricloroetanol na urina",
        "dec_3048": "25 anos (Linha 1.0.19)",
        "esocial_24": "01.19.036",
    },
    "136-52-7": {
        "agente": "Octoato de Cobalto (Naftenato de Cobalto)",
        "nr15_lt": "0,02 mg/m³ como Co (Cancerígeno Cat. 1B)",
        "nr09_acao": "0,01 mg/m³ como Co",
        "nr07_ibe": "Cobalto na urina (pós-jornada)",
        "dec_3048": "25 anos (Linha 1.0.7)",
        "esocial_24": "01.07.001",
    },
    "96-29-7": {
        "agente": "Metiletilcetoxima (MEKO / Butanona oxima)",
        "nr15_lt": "Avaliação Qualitativa (Cancerígeno Cat. 2 UE / Suspeito IARC)",
        "nr09_acao": "Avaliação Qualitativa",
        "nr07_ibe": "Avaliação Clínica e Dermatológica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "64742-47-8": {
        "agente": "Nafta Aromática Pesada (Hidrocarbonetos C9-C12)",
        "nr15_lt": "Avaliar fração aromática — risco Benzeno residual (NR-15 Anexo 13-A)",
        "nr09_acao": "Avaliação Qualitativa (fração cancerígena)",
        "nr07_ibe": "Ácido trans,trans-mucônico (Benzeno residual) / Avaliação Clínica",
        "dec_3048": "25 anos (Linha 1.0.19) — avaliar teor aromático",
        "esocial_24": "01.19.036",
    },
    "64742-95-6": {
        "agente": "Nafta Aromática Leve (Hidrocarbonetos C8-C10)",
        "nr15_lt": "Avaliar fração aromática — risco Benzeno residual (NR-15 Anexo 13-A)",
        "nr09_acao": "Avaliação Qualitativa (fração cancerígena)",
        "nr07_ibe": "Ácido trans,trans-mucônico (Benzeno residual) / Avaliação Clínica",
        "dec_3048": "25 anos (Linha 1.0.19) — avaliar teor aromático",
        "esocial_24": "01.19.036",
    },
    "8052-42-4": {
        "agente": "Asfalto de Petróleo / Alcatrão (CAS alternativo)",
        "nr15_lt": "5 mg/m³ (névoas/fumos)",
        "nr09_acao": "2,5 mg/m³",
        "nr07_ibe": "Avaliação Clínica / Dermatológica",
        "dec_3048": "25 anos (Linha 1.0.19)",
        "esocial_24": "01.19.036",
    },
    "8006-64-2": {
        "agente": "Aguarrás Mineral / Terebentina",
        "nr15_lt": "100 ppm ou 560 mg/m³",
        "nr09_acao": "50 ppm ou 280 mg/m³",
        "nr07_ibe": "Avaliação Clínica e Dermatológica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "8008-20-6": {
        "agente": "Querosene / Combustível de Aviação",
        "nr15_lt": "200 mg/m³ (névoas)",
        "nr09_acao": "100 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "95-63-6": {
        "agente": "1,2,4-Trimetilbenzeno (Pseudocumeno)",
        "nr15_lt": "25 ppm ou 125 mg/m³",
        "nr09_acao": "12,5 ppm ou 62,5 mg/m³",
        "nr07_ibe": "3,4-Dimetilhipúrico na urina",
        "dec_3048": "25 anos (Linha 1.0.19)",
        "esocial_24": "01.19.036",
    },
    "98-82-8": {
        "agente": "Cumeno (Isopropilbenzeno)",
        "nr15_lt": "50 ppm ou 245 mg/m³",
        "nr09_acao": "25 ppm ou 122,5 mg/m³",
        "nr07_ibe": "2-Fenilpropanol na urina (pós-jornada)",
        "dec_3048": "25 anos (Linha 1.0.19)",
        "esocial_24": "01.19.036",
    },
    "78-92-2": {
        "agente": "2-Butanol (Álcool sec-Butílico)",
        "nr15_lt": "100 ppm ou 305 mg/m³",
        "nr09_acao": "50 ppm ou 152,5 mg/m³",
        "nr07_ibe": "MEK na urina (metabólito)",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "141-78-6": {
        "agente": "Acetato de Etila (Éster Etílico)",
        "nr15_lt": "400 ppm ou 1400 mg/m³",
        "nr09_acao": "200 ppm ou 700 mg/m³",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "22464-99-9": {
        "agente": "Zircônio bis(2-etilhexanoato) / Octoato de Zircônio",
        "nr15_lt": "5 mg/m³ como Zr (poeira/fumos)",
        "nr09_acao": "2,5 mg/m³ como Zr",
        "nr07_ibe": "Avaliação Clínica Pulmonar",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
    "6107-56-8": {
        "agente": "Neodecanoato de Zircônio (Secante de Tinta)",
        "nr15_lt": "5 mg/m³ como Zr",
        "nr09_acao": "2,5 mg/m³ como Zr",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado",
        "esocial_24": "09.01.001",
    },
}

# ==========================================
# DICIONÁRIO DE CAMPO (Físico, Biológico, Ergonômico, Acidente)
# ==========================================
dicionario_campo = {
    "Físico: Ruído Contínuo/Intermitente": {
        "agente": "Ruído Contínuo ou Intermitente",
        "nr15_lt": "85 dB(A)", "nr09_acao": "80 dB(A)",
        "nr07_ibe": "Audiometria",
        "dec_3048": "25 anos (Linha 2.0.1)", "esocial_24": "02.01.001",
        "perigo": "Exposição a níveis elevados de pressão sonora",
        "sev": 3, "epi": "Protetor Auditivo (Atenuação adequada)",
    },
    "Físico: Vibração de Mãos e Braços (VMB)": {
        "agente": "Vibração de Mãos e Braços (VMB)",
        "nr15_lt": "5,0 m/s²", "nr09_acao": "2,5 m/s²",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "25 anos (Linha 2.0.2)", "esocial_24": "02.01.002",
        "perigo": "Transmissão de energia mecânica para o sistema mão-braço",
        "sev": 3, "epi": "Luvas antivibração / Revezamento",
    },
    "Físico: Vibração de Corpo Inteiro (VCI)": {
        "agente": "Vibração de Corpo Inteiro (VCI)",
        "nr15_lt": "1,1 m/s² ou 21,0 m/s¹.75",
        "nr09_acao": "0,5 m/s² ou 9,1 m/s¹.75",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "25 anos (Linha 2.0.2)", "esocial_24": "02.01.003",
        "perigo": "Transmissão de energia mecânica para o corpo inteiro",
        "sev": 3, "epi": "Assentos com amortecimento / Revezamento",
    },
    "Físico: Calor": {
        "agente": "Calor (IBUTG)",
        "nr15_lt": "Anexo 3 (IBUTG)", "nr09_acao": "50% do L.T.",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "25 anos (Linha 2.0.3)", "esocial_24": "02.01.004",
        "perigo": "Exposição ao calor acima dos limites de tolerância",
        "sev": 3, "epi": "Roupas leves / Pausas obrigatórias",
    },
    "Físico: Radiações Ionizantes": {
        "agente": "Radiações Ionizantes",
        "nr15_lt": "Anexo 5", "nr09_acao": "50% do L.T.",
        "nr07_ibe": "Hemograma Completo / Avaliação Clínica",
        "dec_3048": "20 anos (Linha 2.0.4)", "esocial_24": "02.01.005",
        "perigo": "Exposição a radiações ionizantes (Raio-X, Gama)",
        "sev": 5, "epi": "Dosímetro / Avental plumbífero",
    },
    "Biológico: Esgoto / Fossas": {
        "agente": "Microorganismos - Esgoto / Fossas",
        "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo",
        "nr07_ibe": "Exames Clínicos / Vacinas",
        "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.005",
        "perigo": "Exposição a agentes biológicos infectocontagiosos",
        "sev": 4, "epi": "Luvas, Botas de PVC, Proteção facial",
    },
    "Biológico: Lixo Urbano": {
        "agente": "Microorganismos - Lixo Urbano",
        "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo",
        "nr07_ibe": "Exames Clínicos / Vacinas",
        "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.007",
        "perigo": "Contato com resíduos e agentes biológicos",
        "sev": 4, "epi": "Luvas anticorte, Botas, Uniforme impermeável",
    },
    "Biológico: Estab. Saúde": {
        "agente": "Microorganismos - Área da Saúde",
        "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo",
        "nr07_ibe": "Exames Clínicos / Vacinas",
        "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.001",
        "perigo": "Exposição a patógenos em ambiente de saúde",
        "sev": 4, "epi": "Luvas de procedimento, Máscara, Avental",
    },
    "Ergonômico: Postura Inadequada": {
        "agente": "Fator Ergonômico - Postura",
        "nr15_lt": "N/A (NR-17)", "nr09_acao": "Avaliação AEP/AET",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Exigência de postura inadequada ou prolongada",
        "sev": 2, "epi": "Medidas Administrativas / Mobiliário Adequado",
    },
    "Ergonômico: Levantamento/Transporte de Peso": {
        "agente": "Fator Ergonômico - Levantamento",
        "nr15_lt": "N/A (NR-17)", "nr09_acao": "Avaliação AEP/AET",
        "nr07_ibe": "Avaliação Clínica / Osteomuscular",
        "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Esforço físico intenso e levantamento manual de cargas",
        "sev": 3, "epi": "Auxílio Mecânico / Treinamento",
    },
    "Acidente: Queda de Altura": {
        "agente": "Risco de Acidente - Altura",
        "nr15_lt": "N/A (NR-35)", "nr09_acao": "N/A",
        "nr07_ibe": "Protocolo Trabalho em Altura",
        "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Trabalho executado acima de 2 metros do nível inferior",
        "sev": 4, "epi": "Cinturão de Segurança, Talabarte, Capacete com Jugular",
    },
    "Acidente: Choque Elétrico": {
        "agente": "Risco de Acidente - Eletricidade",
        "nr15_lt": "N/A (NR-10)", "nr09_acao": "N/A",
        "nr07_ibe": "Avaliação Clínica / ECG",
        "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Contato direto ou indireto com partes energizadas",
        "sev": 5, "epi": "Luvas Isolantes, Vestimenta ATPV, Capacete Classe B",
    },
    "Acidente: Máquinas e Equipamentos": {
        "agente": "Risco de Acidente - Partes Móveis",
        "nr15_lt": "N/A (NR-12)", "nr09_acao": "N/A",
        "nr07_ibe": "Avaliação Clínica",
        "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Operação de máquinas com risco de corte ou esmagamento",
        "sev": 4, "epi": "Luvas de Proteção, Óculos, Botas de Segurança",
    },
}

# ==========================================
# DICIONÁRIOS FASE 2 (CLÍNICA E PCMSO)
# ==========================================
matriz_risco_exame = {
    "TOLUENO":   {"exame": "Ortocresol na Urina",                        "periodico": "6 MESES"},
    "RUÍDO":     {"exame": "Audiometria",                                 "periodico": "12 MESES"},
    "SÍLICA":    {"exame": "Raio-X de Tórax (OIT) + Espirometria",       "periodico": "12 a 24 MESES"},
    "VIBRAÇÃO":  {"exame": "Avaliação Clínica e Osteomuscular",           "periodico": "12 MESES"},
    "POEIRA":    {"exame": "Raio-X de Tórax (OIT)",                      "periodico": "12 MESES"},
    "BIOLÓGIC":  {"exame": "HBsAg / Anti-HBs / Anti-HCV",               "periodico": "12 a 24 MESES"},
    "SANGUE":    {"exame": "HBsAg / Anti-HBs / Anti-HCV",               "periodico": "12 a 24 MESES"},
    "VÍRUS":     {"exame": "HBsAg / Anti-HBs / Anti-HCV",               "periodico": "12 a 24 MESES"},
    "BACTÉRIA":  {"exame": "HBsAg / Anti-HBs / Anti-HCV",               "periodico": "12 a 24 MESES"},
    "QUÍMIC":    {"exame": "Hemograma Completo / Avaliação Hepática",    "periodico": "12 MESES"},
    "BENZENO":   {"exame": "Ácido trans,trans-mucônico / Hemograma",     "periodico": "6 MESES"},
    "CHUMBO":    {"exame": "Chumbo no Sangue (PbS) + ALA-U",            "periodico": "6 MESES"},
    "AMIANTO":   {"exame": "Raio-X (OIT) + Espirometria + TC de Tórax", "periodico": "12 MESES"},
}

matriz_funcao_exame = {
    "TRABALHO EM ALTURA": [
        {"exame": "Hemograma",              "periodicidade": "12 MESES"},
        {"exame": "Glicemia de Jejum",      "periodicidade": "12 MESES"},
        {"exame": "Audiometria",            "periodicidade": "12 MESES"},
        {"exame": "ECG",                    "periodicidade": "12 MESES"},
        {"exame": "Acuidade Visual",        "periodicidade": "12 MESES"},
        {"exame": "Avaliação Psicossocial", "periodicidade": "12 MESES"},
    ],
    "ENCANADOR": [
        {"exame": "Exame Clínico",      "periodicidade": "6 MESES"},
        {"exame": "Audiometria",        "periodicidade": "12 MESES"},
        {"exame": "Acuidade Visual",    "periodicidade": "12 MESES"},
        {"exame": "ECG",                "periodicidade": "12 MESES"},
        {"exame": "Glicemia de Jejum",  "periodicidade": "12 MESES"},
        {"exame": "Hemograma",          "periodicidade": "12 MESES"},
        {"exame": "Espirometria",       "periodicidade": "24 MESES"},
        {"exame": "RX Tórax OIT",       "periodicidade": "12 MESES"},
    ],
    "PEDREIRO": [
        {"exame": "Exame Clínico",  "periodicidade": "6 MESES"},
        {"exame": "RX Tórax OIT",  "periodicidade": "12 MESES"},
        {"exame": "Espirometria",   "periodicidade": "24 MESES"},
        {"exame": "Audiometria",    "periodicidade": "12 MESES"},
    ],
    "ELETRICISTA": [
        {"exame": "Exame Clínico",   "periodicidade": "6 MESES"},
        {"exame": "ECG",             "periodicidade": "12 MESES"},
        {"exame": "Audiometria",     "periodicidade": "12 MESES"},
        {"exame": "Acuidade Visual", "periodicidade": "12 MESES"},
    ],
    "ADMINISTRATIVO": [
        {"exame": "Exame Clínico",   "periodicidade": "12 MESES"},
        {"exame": "Acuidade Visual", "periodicidade": "12 MESES"},
    ],
    "RECEPCIONISTA": [{"exame": "Exame Clínico", "periodicidade": "12 MESES"}],
    "GESTOR":        [{"exame": "Exame Clínico", "periodicidade": "12 MESES"}],
    "BIOMÉDICO": [
        {"exame": "HBsAg",    "periodicidade": "12 MESES"},
        {"exame": "Anti-HBs", "periodicidade": "24 MESES"},
        {"exame": "Anti-HCV", "periodicidade": "12 MESES"},
    ],
    "LABORATÓRIO": [
        {"exame": "HBsAg",    "periodicidade": "12 MESES"},
        {"exame": "Anti-HBs", "periodicidade": "24 MESES"},
        {"exame": "Anti-HCV", "periodicidade": "12 MESES"},
    ],
    "ENFERMAGEM": [
        {"exame": "HBsAg",    "periodicidade": "12 MESES"},
        {"exame": "Anti-HBs", "periodicidade": "24 MESES"},
        {"exame": "Anti-HCV", "periodicidade": "12 MESES"},
    ],
}

# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def limpar_json_ia(texto_bruto):
    texto = texto_bruto.strip()
    texto = re.sub(r'^```json\s*', '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'^```\s*',     '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'\s*```$',     '', texto)
    return texto.strip()


# ==========================================
# CORREÇÃO 1: Validação robusta de números CAS
# ==========================================
def validar_digito_cas(cas: str) -> bool:
    """
    Valida o dígito verificador oficial do número CAS Registry.
    Rejeita CAS com zero à esquerda no primeiro segmento.
    """
    partes = cas.split("-")
    if len(partes) != 3:
        return False
    if partes[0].startswith("0"):
        return False
    if not all(p.isdigit() for p in partes):
        return False
    if not (2 <= len(partes[0]) <= 7 and len(partes[1]) == 2 and len(partes[2]) == 1):
        return False
    digitos     = partes[0] + partes[1]
    verificador = int(partes[2])
    soma = sum(int(d) * (i + 1) for i, d in enumerate(reversed(digitos)))
    return soma % 10 == verificador


def extrair_cas_validos(texto: str) -> list:
    """
    Extrai e valida números CAS de um texto.
    Rejeita falsos positivos (lotes, refs internas, etc).
    """
    candidatos = re.findall(r'\b([1-9]\d{1,6}-\d{2}-\d)\b', texto)
    return [cas for cas in set(candidatos) if validar_digito_cas(cas)]


# ==========================================
# CORREÇÃO 3: Override para cancerígenos confirmados
# ==========================================
def aplicar_override_carcinogenico(codigo_h: str, nivel_risco: str, cas: str = "") -> tuple:
    """
    Aplica override de nível de risco para agentes cancerígenos.
    Base legal: NR-15 Anexo 13-A, NR-01 item 1.5.4, Princípio ALARA.
    Retorna: (nivel_risco_final, acao_final, foi_sobrescrito)
    """
    if codigo_h in CODIGOS_CARCINOGENICOS or cas in CAS_CARCINOGENICOS_CONHECIDOS:
        return (
            "INTOLERÁVEL",
            "TRABALHO PROIBIDO. Agente cancerígeno confirmado — NR-15 Anexo 13-A. "
            "Substituição do agente ou controle de engenharia obrigatório. "
            "Princípio ALARA aplicável.",
            True,
        )
    if codigo_h in CODIGOS_SUSPEITOS:
        niveis_ordem = ["TRIVIAL", "TOLERÁVEL", "MODERADO", "SUBSTANCIAL", "INTOLERÁVEL"]
        idx_atual  = niveis_ordem.index(nivel_risco) if nivel_risco in niveis_ordem else 0
        idx_minimo = niveis_ordem.index("SUBSTANCIAL")
        if idx_atual < idx_minimo:
            return (
                "SUBSTANCIAL",
                "Trabalho não deve iniciar sem redução do risco. "
                "Agente com suspeita de cancerigenicidade — monitoramento PCMSO obrigatório.",
                True,
            )
    return (nivel_risco, acoes_requeridas.get(nivel_risco, "Manual"), False)


# ==========================================
# PATCH #2: Consultar e salvar no dicionário dinâmico (SQLite)
# ==========================================
def consultar_dicionario_dinamico(cas):
    conn = sqlite3.connect('seconci_banco_dados.db')
    c    = conn.cursor()
    c.execute(
        "SELECT agente, nr15_lt, nr09_acao, nr07_ibe, dec_3048, esocial_24 "
        "FROM dicionario_dinamico WHERE cas = ?",
        (cas,)
    )
    row = c.fetchone()
    conn.close()
    if row:
        return {
            "agente":     row[0], "nr15_lt":    row[1],
            "nr09_acao":  row[2], "nr07_ibe":   row[3],
            "dec_3048":   row[4], "esocial_24": row[5],
        }
    return None


def salvar_dicionario_dinamico(cas, dados):
    conn = sqlite3.connect('seconci_banco_dados.db')
    c    = conn.cursor()
    c.execute("""
        INSERT OR REPLACE INTO dicionario_dinamico
        (cas, agente, nr15_lt, nr09_acao, nr07_ibe, dec_3048, esocial_24,
         data_aprendizado, fonte)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        cas,
        dados.get("agente",     "Não identificado"),
        dados.get("nr15_lt",    "Avaliar NR-15"),
        dados.get("nr09_acao",  "Avaliar NR-09"),
        dados.get("nr07_ibe",   "Avaliar NR-07"),
        dados.get("dec_3048",   "Avaliar Anexo IV"),
        dados.get("esocial_24", "Avaliar Tabela 24"),
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        "IA Gemini (automático)",
    ))
    conn.commit()
    conn.close()


# ==========================================
# CORREÇÃO 4A: Prompt aprimorado para IA
# ==========================================
def buscar_dados_completos_ia(cas, texto_fispq):
    """
    MOTOR CASCATA — Camada 3 (versão corrigida):
    Usa gemini-2.5-flash como modelo principal com fallback automático.
    """
    prompt = f"""
Você é um Higienista Ocupacional sênior, especialista em legislação brasileira
de SST e banco de dados químicos internacionais (ECHA, IARC, ACGIH, NIOSH).

Analise o texto da FISPQ/FDS abaixo e identifique o agente químico com
número CAS: {cas}

CONTEXTO IMPORTANTE para sua análise:
- Este documento é de uma FISPQ brasileira de produto usado na construção civil
- Produtos comuns: tintas, solventes, graxas, thinner, primers, vernizes
- Frações petrolíferas (CAS 64742-XX-X): identificar a fração específica e
  verificar teor de benzeno/aromáticos
- Secantes metálicos (cobalto, zircônio, manganês): usar limites do metal
- Cetonas e álcoois: verificar NR-15 Anexo 11
- Para CAS de misturas (UVCB): descrever a mistura e usar LT mais restritivo
  dos componentes principais

Retorne APENAS um objeto JSON puro (sem markdown, sem texto extra) com estes campos:

{{
  "agente": "Nome técnico completo — seja específico",
  "nr15_lt": "Limite de Tolerância conforme Anexos da NR-15. Se cancerígeno: 'Anexo 13-A (Cancerígeno Ocupacional — sem LT)'.",
  "nr09_acao": "Nível de Ação conforme NR-09. Se cancerígeno: 'Avaliação Qualitativa (Princípio ALARA)'.",
  "nr07_ibe": "IBE conforme NR-07/ACGIH. Para metais: metal no sangue/urina. Para solventes: metabólito urinário e momento da coleta.",
  "dec_3048": "Enquadramento no Decreto 3048/99. Se cancerígeno IARC Grupo 1: indicar Linha do Anexo IV. Se não: 'Não Enquadrado'.",
  "esocial_24": "Código exato da Tabela 24 do eSocial. Se não houver específico: '09.01.001'.",
  "classificacao_iarc": "Grupo IARC se disponível (1, 2A, 2B, 3 ou N/A)",
  "cancerigenico": true ou false
}}

Texto da FISPQ (Seções 1, 2 e 3):
{texto_fispq[:15000]}
"""
    # Modelos em ordem de preferência — mesmo padrão do Módulo 2
    MODELOS_CASCATA = [
        "models/gemini-2.5-flash",
        "models/gemini-2.5-pro",
        "models/gemini-2.5-flash-lite",
        "models/gemini-2.0-flash-001",
        "models/gemini-2.0-flash",
        "models/gemini-flash-latest",
    ]

    for modelo in MODELOS_CASCATA:
        try:
            url = (
                f"https://generativelanguage.googleapis.com/v1beta/"
                f"{modelo}:generateContent?key={CHAVE_API_GOOGLE}"
            )
            payload = {
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {
                    "temperature": 0.0,
                    "responseMimeType": "application/json",
                    "maxOutputTokens": 1024,
                },
            }
            resposta = requests.post(
                url,
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=45,
            )

            # Cota excedida ou modelo não encontrado → tenta o próximo
            if resposta.status_code in [429, 404]:
                st.warning(
                    f"⚠ Modelo `{modelo}` indisponível "
                    f"(HTTP {resposta.status_code}). Tentando próximo..."
                )
                continue

            if resposta.status_code == 200:
                texto_bruto = (
                    resposta.json()["candidates"][0]["content"]["parts"][0]["text"]
                )
                texto_limpo = limpar_json_ia(texto_bruto)

                try:
                    dados = json.loads(texto_limpo)
                except json.JSONDecodeError:
                    match = re.search(r'\{.*\}', texto_limpo, re.DOTALL)
                    if match:
                        dados = json.loads(match.group(0))
                    else:
                        continue  # JSON inválido → tenta próximo modelo

                if dados.get("agente") and dados["agente"] not in [
                    "", "Não identificado", "null", "Desconhecido"
                ]:
                    # Sinaliza cancerígenos identificados pela IA
                    if dados.get("cancerigenico") is True:
                        if "Anexo 13-A" not in dados.get("nr15_lt", ""):
                            dados["nr15_lt"] = (
                                dados.get("nr15_lt", "") +
                                " ⚠ Verificar Anexo 13-A (possível cancerígeno)"
                            )
                    salvar_dicionario_dinamico(cas, dados)
                    return dados

        except requests.exceptions.Timeout:
            st.warning(f"⏱ Timeout no modelo `{modelo}` para CAS {cas}. Tentando próximo...")
            continue
        except Exception as e:
            st.warning(f"⚠ Erro no modelo `{modelo}` para CAS {cas}: {e}")
            continue

    # Todos os modelos falharam
    st.warning(f"⚠ Nenhum modelo disponível conseguiu identificar CAS {cas}.")
    return None


# ==========================================
# MOTOR CASCATA — Função principal de resolução de CAS
# ==========================================
def resolver_cas(cas, texto_fispq):
    """
    Camada 1 → Banco Local     — instantâneo, 100% preciso
    Camada 2 → SQLite Dinâmico — instantâneo, aprendido da IA
    Camada 3 → IA Gemini       — online, salva no SQLite ao retornar
    """
    if cas in dicionario_cas:
        return dicionario_cas[cas], "local"

    dados_dinamicos = consultar_dicionario_dinamico(cas)
    if dados_dinamicos:
        return dados_dinamicos, "cache"

    dados_ia = buscar_dados_completos_ia(cas, texto_fispq)
    if dados_ia:
        return dados_ia, "ia"

    return {
        "agente":     "Produto Químico (Revisão Manual Necessária)",
        "nr15_lt":    "Avaliar Anexo 11/12",
        "nr09_acao":  "Avaliar NR-09",
        "nr07_ibe":   "Avaliar NR-07",
        "dec_3048":   "Avaliar Anexo IV",
        "esocial_24": "Avaliar Tabela 24",
    }, "fallback"


# ==========================================
# CORREÇÃO 4B: Pré-processamento e seção de revisão
# ==========================================
def pre_processar_resultados_medicos(resultados_medicos: list) -> tuple:
    resolvidos = [i for i in resultados_medicos if i.get("Fonte") != "fallback"]
    pendentes  = [i for i in resultados_medicos if i.get("Fonte") == "fallback"]
    return resolvidos, pendentes


def gerar_html_secao_revisao(pendentes: list) -> str:
    if not pendentes:
        return ""
    html = f"""
    <div style='border:2px solid #B71C1C; border-radius:6px;
                margin-top:20px; margin-bottom:10px;'>
      <div style='background:#B71C1C; color:#fff; padding:8px 12px;
                  font-weight:bold; font-size:9pt;'>
        ⚠ ATENÇÃO — {len(pendentes)} Agente(s) Requerem Revisão Manual
      </div>
      <div style='padding:10px; font-size:8.5pt; color:#333;'>
        <p style='margin:0 0 8px 0;'>
          Os números CAS abaixo foram encontrados nas FISPQs mas
          <strong>não puderam ser identificados automaticamente</strong>.
          Um técnico deve revisar cada um antes da emissão do laudo final.
        </p>
        <table style='width:100%; border-collapse:collapse; font-size:8pt;'>
          <thead>
            <tr style='background:#F5F5F5;'>
              <th style='border:1px solid #ccc; padding:5px;'>Nº CAS</th>
              <th style='border:1px solid #ccc; padding:5px;'>FISPQ de Origem</th>
              <th style='border:1px solid #ccc; padding:5px;'>GHE</th>
              <th style='border:1px solid #ccc; padding:5px;'>Ação Requerida</th>
            </tr>
          </thead>
          <tbody>
    """
    for item in pendentes:
        html += f"""
            <tr>
              <td style='border:1px solid #ccc; padding:5px;
                         font-family:monospace; font-weight:bold;'>
                {item['Nº CAS']}
              </td>
              <td style='border:1px solid #ccc; padding:5px;'>
                {item['Arquivo Origem']}
              </td>
              <td style='border:1px solid #ccc; padding:5px;'>
                {item['GHE']}
              </td>
              <td style='border:1px solid #ccc; padding:5px; color:#B71C1C;'>
                Pesquisar no ECHA (echa.europa.eu) ou NIOSH Pocket Guide
              </td>
            </tr>
        """
    html += """
          </tbody>
        </table>
        <p style='margin:8px 0 0 0; font-size:7.5pt; color:#666;'>
          Recursos: ECHA | NIOSH NPG | ABHO
        </p>
      </div>
    </div>
    """
    return html


# ==========================================
# MÓDULO 2 — EXTRAÇÃO LOCAL SEM IA
# ==========================================

PALAVRAS_GHE = [
    "GHE", "GRUPO HOMOGÊNEO", "SETOR", "DEPARTAMENTO",
    "FUNÇÃO", "CARGO", "ATIVIDADE",
]

MAPA_AGENTES_TEXTO = {
    # Químicos
    "TOLUENO":            "Tolueno",
    "XILENO":             "Xileno",
    "BENZENO":            "Benzeno",
    "ACETONA":            "Acetona",
    "THINNER":            "Solventes (Thinner)",
    "SOLVENTE":           "Solventes Orgânicos",
    "TINTA":              "Tinta / Verniz (Solventes)",
    "VERNIZ":             "Tinta / Verniz (Solventes)",
    "PRIMER":             "Primer (Solventes)",
    "FUNDO":              "Primer / Fundo (Solventes)",
    "GRAXA":              "Graxa / Lubrificante",
    "DIESEL":             "Diesel / Combustível",
    "QUEROSENE":          "Querosene",
    "ÁCIDO":              "Ácidos (geral)",
    "CIMENTO":            "Cimento Portland (Poeiras)",
    "SÍLICA":             "Sílica Cristalina (Quartzo)",
    "POEIRA":             "Poeiras Minerais",
    "AMIANTO":            "Asbesto / Amianto",
    "CHUMBO":             "Chumbo (Fumos/Poeiras)",
    # Físicos
    "RUÍDO":              "Ruído Contínuo ou Intermitente",
    "VIBRAÇÃO":           "Vibração (VMB/VCI)",
    "CALOR":              "Calor (IBUTG)",
    "FRIO":               "Frio",
    "RADIAÇÃO":           "Radiações Ionizantes",
    "ELETROMAGN":         "Radiações Não Ionizantes",
    "PRESSÃO":            "Pressão Sonora / Atmosférica",
    # Biológicos
    "BIOLÓGICO":          "Agentes Biológicos",
    "VÍRUS":              "Vírus",
    "BACTÉRIA":           "Bactérias",
    "FUNGO":              "Fungos",
    "ESGOTO":             "Esgoto / Águas Servidas",
    "LIXO":               "Resíduos Sólidos / Lixo",
    "SANGUE":             "Material Biológico (Sangue/Fluidos)",
    # Ergonômicos
    "ERGONÔM":            "Fator Ergonômico",
    "POSTURA":            "Postura Inadequada",
    "LEVANTAMENTO":       "Levantamento de Carga",
    "REPETITIV":          "Movimento Repetitivo",
    "TRABALHO EM ALTURA": "Trabalho em Altura",
    # Acidentes
    "ELÉTRICO":           "Risco Elétrico",
    "ELETRICIDADE":       "Risco Elétrico",
    "ALTURA":             "Queda de Altura",
    "MÁQUINA":            "Máquinas e Equipamentos",
    "INCÊNDIO":           "Incêndio / Explosão",
    "QUEDA":              "Queda de Altura",
}

MAPA_CARGOS_CONHECIDOS = [
    "PEDREIRO", "SERVENTE", "ELETRICISTA", "ENCANADOR", "PINTOR",
    "SOLDADOR", "CARPINTEIRO", "ARMADOR", "OPERADOR", "MOTORISTA",
    "TÉCNICO", "ENGENHEIRO", "MESTRE", "ENCARREGADO", "ALMOXARIFE",
    "ADMINISTRATIVO", "AUXILIAR", "ASSISTENTE", "GERENTE", "DIRETOR",
    "RECEPCIONISTA", "SEGURANÇA", "VIGILANTE", "LIMPEZA", "ZELADOR",
    "MÉDICO", "ENFERMEIRO", "TÉCNICO DE ENFERMAGEM", "BIOMÉDICO",
    "LABORATORISTA", "FARMACÊUTICO",
]


def extrair_ghe_texto(texto_pgr: str) -> list:
    """
    Extrai GHEs, cargos e riscos do texto do PGR sem usar IA.
    """
    linhas    = texto_pgr.split("\n")
    blocos    = []
    ghe_atual     = None
    cargos_atual  = set()
    riscos_atual  = []

    for linha in linhas:
        linha_upper = linha.strip().upper()
        if not linha_upper:
            continue

        # Detecta início de novo GHE
        eh_ghe = any(palavra in linha_upper for palavra in PALAVRAS_GHE)
        if eh_ghe and len(linha.strip()) < 80:
            if ghe_atual and (cargos_atual or riscos_atual):
                blocos.append({
                    "ghe":             ghe_atual,
                    "cargos":          list(cargos_atual) or ["Cargo não identificado"],
                    "riscos_mapeados": riscos_atual,
                })
            ghe_atual    = linha.strip()
            cargos_atual = set()
            riscos_atual = []
            continue

        # Detecta cargos
        for cargo in MAPA_CARGOS_CONHECIDOS:
            if cargo in linha_upper:
                cargos_atual.add(linha.strip()[:60])
                break

        # Detecta agentes/riscos
        for palavra_chave, nome_agente in MAPA_AGENTES_TEXTO.items():
            if palavra_chave in linha_upper:
                agentes_existentes = [r["nome_agente"] for r in riscos_atual]
                if nome_agente not in agentes_existentes:
                    riscos_atual.append({
                        "nome_agente":       nome_agente,
                        "perigo_especifico": linha.strip()[:120],
                    })

    # Salva o último bloco
    if ghe_atual and (cargos_atual or riscos_atual):
        blocos.append({
            "ghe":             ghe_atual,
            "cargos":          list(cargos_atual) or ["Cargo não identificado"],
            "riscos_mapeados": riscos_atual,
        })

    # Fallback: bloco geral se não encontrou estrutura
    if not blocos:
        todos_cargos = set()
        todos_riscos = []
        for linha in linhas:
            linha_upper = linha.strip().upper()
            for cargo in MAPA_CARGOS_CONHECIDOS:
                if cargo in linha_upper:
                    todos_cargos.add(linha.strip()[:60])
            for palavra_chave, nome_agente in MAPA_AGENTES_TEXTO.items():
                if palavra_chave in linha_upper:
                    agentes_existentes = [r["nome_agente"] for r in todos_riscos]
                    if nome_agente not in agentes_existentes:
                        todos_riscos.append({
                            "nome_agente":       nome_agente,
                            "perigo_especifico": linha.strip()[:120],
                        })
        if todos_cargos or todos_riscos:
            blocos.append({
                "ghe":             "Geral (extraído automaticamente)",
                "cargos":          list(todos_cargos) or ["Verificar manualmente"],
                "riscos_mapeados": todos_riscos,
            })

    return blocos


def extrair_pgr_com_fallback_ia(texto_pgr: str, pdf_b64: str) -> tuple:
    """
    Etapa 1: extração local por texto (gratuito)
    Etapa 2: IA como fallback se local falhar
    Retorna: (json_pgr, metodo_usado)
    """
    # Etapa 1: local
    resultado_local = extrair_ghe_texto(texto_pgr)
    tem_riscos = any(
        len(bloco.get("riscos_mapeados", [])) > 0
        for bloco in resultado_local
    )
    if resultado_local and tem_riscos:
        return resultado_local, "local"

    # Etapa 2: IA fallback
    MODELOS_FALLBACK = [
        "models/gemini-2.0-flash-lite",
        "models/gemini-2.0-flash-lite-001",
        "models/gemini-flash-lite-latest",
        "models/gemini-2.0-flash",
        "models/gemini-2.5-flash-lite",
        "models/gemini-2.5-flash",
    ]
    prompt_extracao = """
Você é um médico do trabalho. Analise este PDF (Inventário de Riscos).
Identifique GHEs/Setores, Cargos e Agentes Nocivos.
Retorne APENAS JSON puro neste formato:
[{"ghe": "Nome", "cargos": ["Cargo"], "riscos_mapeados": [{"nome_agente": "Agente", "perigo_especifico": "Descrição"}]}]
"""
    for modelo in MODELOS_FALLBACK:
        try:
            url = (
                f"https://generativelanguage.googleapis.com/v1beta/"
                f"{modelo}:generateContent?key={CHAVE_API_GOOGLE}"
            )
            payload = {
                "contents": [{
                    "parts": [
                        {"text": prompt_extracao},
                        {"inlineData": {"mimeType": "application/pdf", "data": pdf_b64}},
                    ]
                }],
                "generationConfig": {
                    "temperature": 0.0,
                    "responseMimeType": "application/json",
                },
            }
            resp = requests.post(
                url,
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=120,
            )
            if resp.status_code == 200:
                texto = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                limpo = limpar_json_ia(texto)
                try:
                    dados = json.loads(limpo)
                    if dados:
                        return dados, "ia"
                except Exception:
                    match = re.search(r'\[.*\]', limpo, re.DOTALL)
                    if match:
                        dados = json.loads(match.group(0))
                        if dados:
                            return dados, "ia"
            elif resp.status_code in [429, 404]:
                continue
        except Exception:
            continue

    return resultado_local, "local_parcial"


# ==========================================
# PROCESSAMENTO PCMSO
# ==========================================
def processar_pcmso(dados_pgr_json):
    tabela_pcmso = []
    for ghe in dados_pgr_json:
        nome_ghe = ghe.get("ghe", "Sem GHE")
        cargos   = ghe.get("cargos", [])
        riscos   = ghe.get("riscos_mapeados", [])

        for cargo in cargos:
            exames_do_cargo = [{
                "exame": "Exame Clínico (Anamnese/Físico)",
                "periodicidade": "12 MESES",
                "motivo": "NR-07 Básico",
            }]
            cargo_upper = cargo.upper()

            for funcao_chave, exames in matriz_funcao_exame.items():
                if funcao_chave in cargo_upper:
                    exames_do_cargo.extend(exames)

            for risco in riscos:
                agente = risco.get("nome_agente", "").upper()
                perigo = risco.get("perigo_especifico", "").upper()
                texto_risco = agente + " " + perigo

                for agente_chave, regra in matriz_risco_exame.items():
                    if agente_chave in texto_risco:
                        exames_do_cargo.append({
                            "exame":         regra["exame"],
                            "periodicidade": regra["periodico"],
                            "motivo":        f"Exposição a {agente_chave}",
                        })

                if "ALTURA" in texto_risco:
                    exames_do_cargo.extend(matriz_funcao_exame["TRABALHO EM ALTURA"])

            exames_unicos = list({v["exame"]: v for v in exames_do_cargo}.values())
            for ex in exames_unicos:
                tabela_pcmso.append({
                    "GHE / Setor":                  nome_ghe,
                    "Cargo":                        cargo,
                    "Exame Clínico/Complementar":   ex["exame"],
                    "Periodicidade":                ex.get("periodicidade", ex.get("periodico", "12 MESES")),
                    "Justificativa Legal / Risco":  ex.get("motivo", "Protocolo Função"),
                })
    return pd.DataFrame(tabela_pcmso)


# ==========================================
# GERADORES DE HTML
# ==========================================
def cor_risco(nivel):
    return {
        "TRIVIAL":     "#FFFFFF",
        "TOLERÁVEL":   "#92D050",
        "MODERADO":    "#FFFF00",
        "SUBSTANCIAL": "#FF6600",
        "INTOLERÁVEL": "#FF0000",
    }.get(nivel, "#FFFFFF")


def badge_fonte(fonte):
    badges = {
        "local":    "<span style='background:#084D22;color:#fff;border-radius:4px;padding:1px 5px;font-size:8pt;'>LOCAL</span>",
        "cache":    "<span style='background:#1565C0;color:#fff;border-radius:4px;padding:1px 5px;font-size:8pt;'>CACHE</span>",
        "ia":       "<span style='background:#6A1B9A;color:#fff;border-radius:4px;padding:1px 5px;font-size:8pt;'>IA</span>",
        "fallback": "<span style='background:#B71C1C;color:#fff;border-radius:4px;padding:1px 5px;font-size:8pt;'>REVISAR</span>",
    }
    return badges.get(fonte, "")


def gerar_html_anexo(resultados_pgr, resultados_medicos, pendentes=None):
    if pendentes is None:
        pendentes = []

    html_content = """<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset="utf-8">
    <style>
      body { font-family: 'Arial', sans-serif; font-size: 10pt; color: #000; }
      .anexo-header { background-color: #084D22; color: #FFF; padding: 14px 20px;
        font-size: 13pt; font-weight: bold; margin-bottom: 20px; text-align: center; }
      .funcao-card { border: 1px solid #084D22; margin-bottom: 20px; }
      .funcao-card-header { background-color: #084D22; padding: 10px;
        font-weight: bold; color: #FFF; font-size: 10pt; }
      .funcao-mini-table { width: 100%; border-collapse: collapse;
        font-size: 9pt; margin: 8px 0; }
      .funcao-mini-table th { background-color: #0F823B; color: #FFF;
        padding: 8px; text-align: left; border: 1px solid #000; }
      .funcao-mini-table td { padding: 5px; border: 1px solid #000; vertical-align: top; }
      h4 { color: #084D22; margin: 15px 0 5px 0; font-size: 10pt; padding: 5px; }
    </style></head><body>
    <div class='anexo-header'>
      ANEXO I - INVENTÁRIO DE RISCOS E ENQUADRAMENTO PREVIDENCIÁRIO
    </div>
    """

    df_pgr = pd.DataFrame(resultados_pgr)
    df_med = pd.DataFrame(resultados_medicos)
    df_pen = pd.DataFrame(pendentes) if pendentes else pd.DataFrame()

    ghes_pgr = df_pgr["GHE"].unique().tolist() if not df_pgr.empty else []
    ghes_med = df_med["GHE"].unique().tolist() if not df_med.empty else []
    ghes_pen = df_pen["GHE"].unique().tolist() if not df_pen.empty else []
    ghes = sorted(set(ghes_pgr + ghes_med + ghes_pen))

    for ghe in ghes:
        html_content += (
            f"<div class='funcao-card'>"
            f"<div class='funcao-card-header'>GHE: {ghe}</div>"
            f"<div style='padding:10px;'>"
        )

        # ── Inventário de Risco ────────────────────────────────────
        pgr_ghe = df_pgr[df_pgr["GHE"] == ghe] if not df_pgr.empty else pd.DataFrame()
        if not pgr_ghe.empty:
            html_content += (
                "<h4>📋 Inventário de Risco (NR-01)</h4>"
                "<table class='funcao-mini-table'><thead><tr>"
                "<th>Origem / FISPQ / FDS</th><th>Perigo Identificado</th>"
                "<th>Sev.</th><th>Prob.</th><th>Nível de Risco</th>"
                "<th>Ação Requerida</th><th>EPI Recomendado (NR-06)</th>"
                "</tr></thead><tbody>"
            )
            for _, row in pgr_ghe.iterrows():
                cor         = cor_risco(row["NÍVEL DE RISCO"])
                linha_style = "background-color:#FFF3E0;" if "OVERRIDE" in str(row.get("Perigo Identificado", "")) else ""
                html_content += (
                    f"<tr style='{linha_style}'>"
                    f"<td>{row['Arquivo Origem']}</td>"
                    f"<td><b>{row['Código GHS']}</b> {row['Perigo Identificado']}</td>"
                    f"<td>{row['Severidade']}</td>"
                    f"<td>{row['Probabilidade']}</td>"
                    f"<td style='background-color:{cor};font-weight:bold;text-align:center;'>"
                    f"{row['NÍVEL DE RISCO']}</td>"
                    f"<td>{row['Ação Requerida']}</td>"
                    f"<td>{row['EPI (NR-06)']}</td>"
                    f"</tr>"
                )
            html_content += "</tbody></table>"

        # ── Diretrizes Médicas (apenas resolvidos) ─────────────────
        med_ghe = df_med[df_med["GHE"] == ghe].copy() if not df_med.empty else pd.DataFrame()
        if not med_ghe.empty:
            med_ghe = med_ghe.drop_duplicates(subset=["Nº CAS"], keep="first")
            html_content += (
                "<h4>⚕ Diretrizes Médicas e Previdenciárias</h4>"
                "<table class='funcao-mini-table'><thead><tr>"
                "<th>Fonte</th><th>Cód / CAS</th><th>Agente</th>"
                "<th>Lim. Tol. (NR-15)</th><th>Nível Ação (NR-09)</th>"
                "<th>Exame/IBE (NR-07)</th><th>Dec 3048</th><th>eSocial</th>"
                "</tr></thead><tbody>"
            )
            for _, row in med_ghe.iterrows():
                fonte_badge = badge_fonte(row.get("Fonte", "local"))
                lt          = str(row.get("Lim. Tolerância (NR-15)", ""))
                cell_style  = "background-color:#FFF3E0; font-weight:bold;" if "Anexo 13-A" in lt or "Cancerígeno" in lt else ""
                html_content += (
                    f"<tr style='{cell_style}'>"
                    f"<td>{fonte_badge}</td>"
                    f"<td>{row['Nº CAS']}</td>"
                    f"<td>{row['Agente Químico']}</td>"
                    f"<td>{row['Lim. Tolerância (NR-15)']}</td>"
                    f"<td>{row['Nível de Ação (NR-09)']}</td>"
                    f"<td>{row['IBE (NR-07)']}</td>"
                    f"<td>{row['Dec 3048']}</td>"
                    f"<td>{row['eSocial']}</td>"
                    f"</tr>"
                )
            html_content += "</tbody></table>"

        # ── Seção de revisão manual (pendentes deste GHE) ─────────
        pen_ghe = [p for p in pendentes if p.get("GHE") == ghe]
        if pen_ghe:
            html_content += gerar_html_secao_revisao(pen_ghe)

        html_content += "</div></div>"

    html_content += "</body></html>"
    return html_content


def gerar_html_pcmso(df_pcmso):
    html = """<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset="utf-8">
    <style>
      body { font-family: 'Arial', sans-serif; color: #000000; }
      .header { background-color: #084D22; color: #FFFFFF; padding: 14px;
        text-align: center; font-weight: bold; font-size: 16px; margin-bottom: 20px; }
      table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 13px; }
      th { background-color: #1AA04B; color: #FFFFFF; padding: 12px 8px;
        text-align: left; border: 1px solid #000000; }
      td { border: 1px solid #000000; padding: 10px 8px; vertical-align: top; }
      tr:nth-child(even) { background-color: #F4F8F5; }
    </style></head><body>
    <div class='header'>MATRIZ DE EXAMES - PCMSO (GERADO VIA IA)</div>
    <table><tr>
      <th>GHE / Setor</th><th>Cargo</th>
      <th>Exame Clínico / Complementar</th>
      <th>Periodicidade</th><th>Justificativa / Agente</th>
    </tr>
    """
    for _, row in df_pcmso.iterrows():
        html += (
            f"<tr><td><strong>{row['GHE / Setor']}</strong></td>"
            f"<td>{row['Cargo']}</td>"
            f"<td>{row['Exame Clínico/Complementar']}</td>"
            f"<td>{row['Periodicidade']}</td>"
            f"<td>{row['Justificativa Legal / Risco']}</td></tr>"
        )
    html += "</table></body></html>"
    return html


# ==========================================
# HELPER: SALVAR NO BANCO DE LAUDOS
# ==========================================
def salvar_historico(nome_projeto, html_relatorio):
    conn = sqlite3.connect('seconci_banco_dados.db')
    c    = conn.cursor()
    c.execute(
        "INSERT INTO historico_laudos (nome_projeto, data_salvamento, html_relatorio) VALUES (?, ?, ?)",
        (nome_projeto, datetime.now().strftime("%d/%m/%Y %H:%M"), html_relatorio)
    )
    conn.commit()
    conn.close()


# ==========================================
# CORPO PRINCIPAL
# ==========================================
st.title("Sistema Integrado SST - Seconci GO 🚀")

# ── VISUALIZAR HISTÓRICO ──────────────────────────────────────────────────────
if historico_selecionado:
    st.markdown("### 🗄 Visualizando Relatório do Histórico")
    aba_preview, aba_download = st.tabs(["👁 Pré-visualizar", "📄 Baixar em Word (.doc)"])
    with aba_preview:
        components.html(historico_selecionado, height=700, scrolling=True)
    with aba_download:
        st.download_button(
            "⬇ Baixar Relatório",
            data=historico_selecionado.encode("utf-8"),
            file_name="Relatorio_Historico.doc",
            mime="application/msword",
        )

# ── DASHBOARD ─────────────────────────────────────────────────────────────────
elif "🏠" in modulo_selecionado:
    st.header("🏠 Dashboard — Visão Geral")
    conn    = sqlite3.connect('seconci_banco_dados.db')
    df_dash = pd.read_sql_query("SELECT * FROM historico_laudos ORDER BY id DESC", conn)
    df_din  = pd.read_sql_query("SELECT cas, agente, data_aprendizado FROM dicionario_dinamico ORDER BY rowid DESC", conn)
    conn.close()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📁 Total de Laudos", len(df_dash))
    col2.metric("⚙ Relatórios PGR",   len(df_dash[df_dash["nome_projeto"].str.contains("PGR",   na=False)]))
    col3.metric("🩺 Relatórios PCMSO", len(df_dash[df_dash["nome_projeto"].str.contains("PCMSO", na=False)]))
    col4.metric("🧠 Agentes Aprendidos (IA)", len(df_din))

    st.markdown("---")
    if not df_dash.empty:
        st.markdown("### 📋 Últimos Projetos")
        st.dataframe(
            df_dash[["id", "nome_projeto", "data_salvamento"]].rename(columns={
                "id": "ID", "nome_projeto": "Projeto", "data_salvamento": "Data"
            }),
            use_container_width=True, hide_index=True,
        )
    if not df_din.empty:
        st.markdown("### 🧠 Base de Conhecimento Dinâmica (aprendido pela IA)")
        st.caption("Estes agentes foram identificados automaticamente pelo Gemini e já estão disponíveis sem precisar chamar a IA novamente.")
        st.dataframe(
            df_din.rename(columns={
                "cas": "Nº CAS", "agente": "Agente Identificado", "data_aprendizado": "Aprendido em"
            }),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Nenhum laudo salvo ainda. Use os módulos para gerar e salvar relatórios.")

# ── MÓDULO 1: ENGENHARIA ──────────────────────────────────────────────────────
elif "1️⃣" in modulo_selecionado:
    st.header("⚙ Módulo de Engenharia: Extrator de FISPQs / FDS (com IA Integrada)")
    st.info("""
    **Motor Híbrido em 3 camadas ativo:**
    🟢 **Camada 1** — Banco local (instantâneo): agentes da construção civil já cadastrados.
    🔵 **Camada 2** — Cache inteligente (instantâneo): agentes já aprendidos pela IA em consultas anteriores.
    🟣 **Camada 3** — IA Gemini (online): qualquer agente desconhecido é pesquisado e **salvo automaticamente** para a próxima vez.
    """)

    arquivos_fispq = st.file_uploader(
        "Insira as FISPQs / FDS em PDF", type=["pdf"], accept_multiple_files=True
    )

    textos_pdfs = {}
    df_editado  = pd.DataFrame()
    ghe_opcoes  = ["Nenhum GHE definido"]

    if arquivos_fispq:
        with st.spinner("Lendo o conteúdo dos PDFs..."):
            for arquivo in arquivos_fispq:
                pdf_bytes = io.BytesIO(arquivo.getvalue())
                try:
                    with pdfplumber.open(pdf_bytes) as pdf:
                        texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
                    textos_pdfs[arquivo.name] = texto
                except Exception as e:
                    st.error(f"❌ Erro ao ler '{arquivo.name}': {e}")
                    textos_pdfs[arquivo.name] = ""

        st.success(f"✅ {len(textos_pdfs)} PDF(s) carregado(s) com sucesso.")
        st.markdown("---")
        st.markdown("### 2️⃣ Definição de GHEs Químicos")

        nomes_arquivos = [arq.name for arq in arquivos_fispq]
        df_mapeamento  = pd.DataFrame([
            {"GHE": "GHE 01 - Digite a Função", "Arquivo FISPQ/FDS": nome, "Probabilidade": 3}
            for nome in nomes_arquivos
        ])
        df_editado = st.data_editor(
            df_mapeamento, num_rows="dynamic",
            column_config={
                "GHE": st.column_config.TextColumn("Nome do GHE", required=True),
                "Arquivo FISPQ/FDS": st.column_config.SelectboxColumn(
                    "Arquivo (FISPQ/FDS)", options=nomes_arquivos, required=True
                ),
                "Probabilidade": st.column_config.NumberColumn(
                    "Prob. (1-5)", min_value=1, max_value=5, required=True
                ),
            },
            use_container_width=True, key="editor_ghe",
        )
        ghe_opcoes = df_editado["GHE"].unique().tolist() if not df_editado.empty else ["Nenhum GHE definido"]

        st.markdown("### 3️⃣ Avaliações de Campo: Físicos, Biológicos, Ergonômicos e Acidentes")
        agentes_opcoes   = list(dicionario_campo.keys())
        df_fis_bio_inicial = pd.DataFrame([{
            "GHE": ghe_opcoes[0], "Agente": agentes_opcoes[0], "Probabilidade": 3
        }])
        df_fis_bio_editado = st.data_editor(
            df_fis_bio_inicial, num_rows="dynamic",
            column_config={
                "GHE": st.column_config.SelectboxColumn(
                    "GHE de Destino", options=ghe_opcoes, required=True
                ),
                "Agente": st.column_config.SelectboxColumn(
                    "Agente / Fator de Risco", options=agentes_opcoes, required=True
                ),
                "Probabilidade": st.column_config.NumberColumn(
                    "Prob. (1-5)", min_value=1, max_value=5, required=True
                ),
            },
            use_container_width=True, key="editor_campo",
        )

        if st.button("🪄 Processar GHEs e Gerar Relatório", use_container_width=True, type="primary"):
            with st.spinner("Processando com Motor Híbrido (Local → Cache → IA)..."):
                resultados_pgr, resultados_medicos = [], []
                contadores     = {"local": 0, "cache": 0, "ia": 0, "fallback": 0}
                vistos_por_ghe = {}

                # ── PROCESSAR FISPQs ──────────────────────────────────────
                if not df_editado.empty:
                    for _, row in df_editado.iterrows():
                        nome_ghe = row["GHE"]
                        nome_arq = row["Arquivo FISPQ/FDS"]
                        v_prob   = int(row["Probabilidade"])

                        if nome_arq not in textos_pdfs or not textos_pdfs[nome_arq]:
                            continue

                        texto_completo = textos_pdfs[nome_arq]

                        # CORREÇÃO 1: extração com validação de dígito verificador
                        cas_encontrados = extrair_cas_validos(texto_completo)

                        for cas in cas_encontrados:
                            chave_unica = f"{nome_ghe}|{cas}"
                            if chave_unica in vistos_por_ghe:
                                continue
                            vistos_por_ghe[chave_unica] = True

                            dados_med, fonte = resolver_cas(cas, texto_completo)
                            contadores[fonte] += 1

                            resultados_medicos.append({
                                "GHE":                    nome_ghe,
                                "Arquivo Origem":         nome_arq,
                                "Nº CAS":                 cas,
                                "Agente Químico":         dados_med["agente"],
                                "Lim. Tolerância (NR-15)": dados_med["nr15_lt"],
                                "Nível de Ação (NR-09)":  dados_med["nr09_acao"],
                                "IBE (NR-07)":            dados_med.get("nr07_ibe",   "N/A"),
                                "Dec 3048":               dados_med.get("dec_3048",   "Não Enquadrado"),
                                "eSocial":                dados_med.get("esocial_24", "09.01.001"),
                                "Fonte":                  fonte,
                            })

                        # CORREÇÃO 3: override cancerígenos nos códigos H
                        h_encontradas = list(set(re.findall(r'\bH\d{3}\b', texto_completo)))
                        for codigo in h_encontradas:
                            if codigo in dicionario_h:
                                dados_h = dicionario_h[codigo]
                                nivel_risco_base = matriz_oficial.get(
                                    (dados_h["sev"], v_prob), "N/A"
                                )
                                nivel_risco_final, acao_final, foi_sobrescrito = aplicar_override_carcinogenico(
                                    codigo_h=codigo,
                                    nivel_risco=nivel_risco_base,
                                )
                                perigo_texto = dados_h["desc"]
                                if foi_sobrescrito:
                                    perigo_texto += " ⚠ [OVERRIDE NORMATIVO APLICADO]"

                                resultados_pgr.append({
                                    "GHE":               nome_ghe,
                                    "Arquivo Origem":    nome_arq,
                                    "Código GHS":        codigo,
                                    "Perigo Identificado": perigo_texto,
                                    "Severidade":        texto_sev.get(dados_h["sev"], str(dados_h["sev"])),
                                    "Probabilidade":     str(v_prob),
                                    "NÍVEL DE RISCO":    nivel_risco_final,
                                    "Ação Requerida":    acao_final,
                                    "EPI (NR-06)":       dados_h["epi"],
                                })

                # ── PROCESSAR AGENTES DE CAMPO ────────────────────────────
                if (not df_fis_bio_editado.empty and
                        df_fis_bio_editado["GHE"].iloc[0] != "Nenhum GHE definido"):
                    for _, row in df_fis_bio_editado.iterrows():
                        nome_ghe   = row["GHE"]
                        nome_agente = row["Agente"]
                        v_prob     = int(row["Probabilidade"])

                        if nome_agente in dicionario_campo:
                            dados_fis   = dicionario_campo[nome_agente]
                            nivel_risco = matriz_oficial.get((dados_fis["sev"], v_prob), "N/A")

                            resultados_medicos.append({
                                "GHE":                    nome_ghe,
                                "Arquivo Origem":         "Avaliação de Campo",
                                "Nº CAS":                 "-",
                                "Agente Químico":         dados_fis["agente"],
                                "Lim. Tolerância (NR-15)": dados_fis["nr15_lt"],
                                "Nível de Ação (NR-09)":  dados_fis["nr09_acao"],
                                "IBE (NR-07)":            dados_fis["nr07_ibe"],
                                "Dec 3048":               dados_fis["dec_3048"],
                                "eSocial":                dados_fis["esocial_24"],
                                "Fonte":                  "local",
                            })
                            resultados_pgr.append({
                                "GHE":               nome_ghe,
                                "Arquivo Origem":    "Avaliação de Campo",
                                "Código GHS":        "-",
                                "Perigo Identificado": dados_fis["perigo"],
                                "Severidade":        texto_sev.get(dados_fis["sev"], str(dados_fis["sev"])),
                                "Probabilidade":     str(v_prob),
                                "NÍVEL DE RISCO":    nivel_risco,
                                "Ação Requerida":    acoes_requeridas.get(nivel_risco, "Manual"),
                                "EPI (NR-06)":       dados_fis["epi"],
                            })

                if resultados_pgr or resultados_medicos:
                    # CORREÇÃO 4B: separa resolvidos dos pendentes
                    med_resolvidos, med_pendentes = pre_processar_resultados_medicos(resultados_medicos)

                    html_final = gerar_html_anexo(resultados_pgr, med_resolvidos, med_pendentes)
                    st.session_state["ultimo_html_eng"]   = html_final
                    st.session_state["dados_pgr_brutos"]  = resultados_pgr
                    st.session_state["dados_med_brutos"]  = resultados_medicos

                    total = sum(contadores.values())
                    st.success("✅ Relatório Consolidado Gerado com Motor Híbrido Inteligente!")

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("🟢 Banco Local",       contadores["local"],    help="Encontrado no dicionário fixo")
                    c2.metric("🔵 Cache SQLite",      contadores["cache"],    help="Aprendido em consulta anterior")
                    c3.metric("🟣 Consultado por IA", contadores["ia"],       help="Novo agente identificado pelo Gemini e salvo")
                    c4.metric("🔴 Revisão Manual",    contadores["fallback"], help="IA não conseguiu identificar — requer revisão")

                    if med_pendentes:
                        pct = len(med_pendentes) / max(total, 1) * 100
                        if pct > 30:
                            st.error(
                                f"⚠ {len(med_pendentes)} agentes ({pct:.0f}%) precisam de revisão manual. "
                                f"Verifique a qualidade do texto das FISPQs."
                            )
                        else:
                            st.warning(
                                f"⚠ {len(med_pendentes)} agente(s) aguardam revisão manual "
                                f"(listados no final do relatório)."
                            )
                else:
                    st.warning("⚠ Nenhum dado encontrado. Verifique se os PDFs contêm texto extraível e os GHEs estão corretamente mapeados.")

    # ── EXIBIR RESULTADO ──────────────────────────────────────────────────────
    if "ultimo_html_eng" in st.session_state:
        st.markdown("---")
        col1, col2 = st.columns([3, 1])
        with col1:
            nome_projeto_eng = st.text_input("Nome da Empresa / Projeto (para salvar):", key="nome_eng")
        with col2:
            st.write(""); st.write("")
            if st.button("💾 Gravar no Banco de Dados", key="btn_salvar_eng", use_container_width=True):
                if nome_projeto_eng:
                    salvar_historico(nome_projeto_eng + " (PGR)", st.session_state["ultimo_html_eng"])
                    st.success("✅ Salvo com sucesso!")
                else:
                    st.error("Digite o nome do projeto antes de salvar.")

        aba_preview, aba_download = st.tabs(["👁 Pré-visualizar", "📄 Baixar (.doc)"])
        with aba_preview:
            components.html(st.session_state["ultimo_html_eng"], height=600, scrolling=True)
        with aba_download:
            st.download_button(
                "⬇ Baixar Relatório PGR",
                data=st.session_state["ultimo_html_eng"].encode("utf-8"),
                file_name="PGR_Inventario_Riscos.doc",
                mime="application/msword",
                use_container_width=True,
            )

# ── MÓDULO 2: MEDICINA ────────────────────────────────────────────────────────
elif "2️⃣" in modulo_selecionado:
    st.header("🩺 Módulo Médico: Importador de PGR e Gerador de PCMSO")
    st.info("Faça o upload do Inventário de Riscos (PGR). A IA fará a leitura e o cruzamento com as matrizes da NR-07.")

    arquivo_pgr = st.file_uploader("Arraste o PDF do PGR aqui", type=["pdf"])

    if arquivo_pgr:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 Extrair Riscos e Gerar PCMSO", type="primary", use_container_width=True):
            with st.spinner("Motor IA analisando o documento... Aguarde."):
                pdf_bytes = arquivo_pgr.getvalue()
                pdf_b64   = base64.b64encode(pdf_bytes).decode("utf-8")

                prompt_extracao = """
Você é um médico do trabalho e engenheiro de segurança. Analise este documento
PDF (Inventário de Riscos Ocupacionais).
Sua missão é CAÇAR qualquer relação entre Funções/Cargos e Agentes Nocivos
(Físicos, Químicos, Biológicos).
CRÍTICO: Para cada risco encontrado, classifique o "nome_agente" de forma clara
(ex: "Risco Biológico", "Ruído", "Produtos Químicos").
Se não houver a palavra "GHE", agrupe pelo nome do "Setor" ou da "Função". NUNCA retorne vazio.
Retorne APENAS JSON puro (sem markdown) neste formato exato:
[
  {
    "ghe": "Nome do Setor, GHE ou Função",
    "cargos": ["Nome do Cargo 1", "Nome do Cargo 2"],
    "riscos_mapeados": [
      {"nome_agente": "Ex: Risco Biológico", "perigo_especifico": "Ex: Exposição a vírus"}
    ]
  }
]
"""
                # ── Descobre os modelos disponíveis na conta ───────────────
                modelos_disponiveis = []
                try:
                    url_lista  = (
                        "https://generativelanguage.googleapis.com/v1beta/models?key="
                        + CHAVE_API_GOOGLE
                    )
                    resp_lista = requests.get(url_lista, timeout=15)
                    if resp_lista.status_code == 200:
                        todos = resp_lista.json().get("models", [])
                        modelos_disponiveis = [
                            m["name"] for m in todos
                            if "generateContent" in m.get("supportedGenerationMethods", [])
                        ]
                except Exception as e_lista:
                    st.warning(f"⚠ Não foi possível listar modelos: {e_lista}")

                # ── Ordem de preferência ───────────────────────────────────
                # Começa pelos mais leves (free tier) → mais pesados (pago)
                ORDEM_PREFERENCIA = [
                    "models/gemini-2.0-flash-lite",
                    "models/gemini-2.0-flash-lite-001",
                    "models/gemini-flash-lite-latest",
                    "models/gemini-2.0-flash",
                    "models/gemini-2.0-flash-001",
                    "models/gemini-flash-latest",
                    "models/gemini-2.5-flash-lite",
                    "models/gemini-2.5-flash",
                    "models/gemini-2.5-pro",
                    "models/gemini-pro-latest",
                ]

                # Filtra apenas os que estão disponíveis na conta
                candidatos = [m for m in ORDEM_PREFERENCIA if m in modelos_disponiveis]

                # Se a lista filtrada estiver vazia, tenta todos da preferência
                if not candidatos:
                    candidatos = ORDEM_PREFERENCIA

                modelo_escolhido = None
                resposta         = None

                # ── Tenta cada modelo até um funcionar ────────────────────
                for modelo in candidatos:
                    st.caption(f"🔄 Tentando modelo: `{modelo}`")

                    url_google = (
                        f"https://generativelanguage.googleapis.com/v1beta/"
                        f"{modelo}:generateContent?key={CHAVE_API_GOOGLE}"
                    )
                    payload = {
                        "contents": [{
                            "parts": [
                                {"text": prompt_extracao},
                                {"inlineData": {"mimeType": "application/pdf", "data": pdf_b64}},
                            ]
                        }],
                        "generationConfig": {
                            "temperature": 0.0,
                            "responseMimeType": "application/json",
                        },
                    }

                    try:
                        resposta = requests.post(
                            url_google,
                            headers={"Content-Type": "application/json"},
                            json=payload,
                            timeout=120,
                        )

                        if resposta.status_code == 200:
                            modelo_escolhido = modelo
                            st.caption(f"✅ Modelo funcionando: `{modelo}`")
                            break

                        elif resposta.status_code == 429:
                            # Extrai a mensagem de limite para diagnóstico
                            erro_json = resposta.json().get("error", {})
                            msg_erro  = erro_json.get("message", "")

                            if "limit: 0" in msg_erro:
                                st.caption(
                                    f"🔒 `{modelo}` — cota zero no plano atual. "
                                    f"Tentando próximo..."
                                )
                            else:
                                st.caption(
                                    f"⏳ `{modelo}` — limite de requisições atingido. "
                                    f"Tentando próximo..."
                                )
                            continue

                        elif resposta.status_code == 404:
                            st.caption(f"❌ `{modelo}` — não encontrado. Tentando próximo...")
                            continue

                        else:
                            st.caption(
                                f"⚠ `{modelo}` retornou HTTP {resposta.status_code}. "
                                f"Tentando próximo..."
                            )
                            continue

                    except requests.exceptions.Timeout:
                        st.caption(f"⏱ `{modelo}` — timeout. Tentando próximo...")
                        continue
                    except Exception as e_req:
                        st.caption(f"⚠ `{modelo}` — erro: {e_req}. Tentando próximo...")
                        continue

                # ── Nenhum modelo funcionou ────────────────────────────────
                if not modelo_escolhido or resposta is None or resposta.status_code != 200:
                    st.error("""
❌ **Nenhum modelo de IA disponível conseguiu processar o documento.**

**Causa mais provável:** Sua chave de API está no plano **gratuito (Free Tier)**,
que tem cota zero (`limit: 0`) para os modelos mais recentes.

**Como resolver:**
1. Acesse [Google AI Studio](https://aistudio.google.com) → verifique seu plano
2. Acesse [Google Cloud Console](https://console.cloud.google.com) → ative o faturamento
3. Ou acesse [ai.dev/rate-limit](https://ai.dev/rate-limit) para ver seus limites atuais

**Alternativa sem custo:** Reduza o tamanho do PDF (menos páginas) e tente novamente.
O modelo `gemini-2.0-flash-lite` tem a cota gratuita mais generosa.
                    """)
                    st.stop()

                # ── Processa a resposta do modelo que funcionou ────────────
                try:
                    resultado_texto = (
                        resposta.json()["candidates"][0]["content"]["parts"][0]["text"]
                    )
                    resultado_limpo = limpar_json_ia(resultado_texto)

                    try:
                        json_pgr = json.loads(resultado_limpo)
                    except json.JSONDecodeError:
                        match = re.search(r'\[.*\]', resultado_limpo, re.DOTALL)
                        if match:
                            json_pgr = json.loads(match.group(0))
                        else:
                            st.error("❌ A IA retornou um formato inesperado. Tente novamente.")
                            st.code(resultado_limpo[:500])
                            st.stop()

                    df_pcmso = processar_pcmso(json_pgr)

                    if not df_pcmso.empty:
                        html_pcmso = gerar_html_pcmso(df_pcmso)
                        st.session_state["ultimo_html_pcmso"] = html_pcmso
                        st.session_state["df_pcmso_preview"]  = df_pcmso
                        st.session_state["modelo_usado_pcmso"] = modelo_escolhido
                        st.success(
                            f"✅ PCMSO gerado com {len(df_pcmso)} linhas de exames! "
                            f"(modelo: `{modelo_escolhido}`)"
                        )
                    else:
                        st.warning(
                            "⚠ A IA não encontrou riscos no documento. "
                            "Verifique se é um PGR/Inventário de Riscos válido."
                        )

                except Exception as e_parse:
                    st.error(f"❌ Erro ao processar resposta da IA: {e_parse}")

    # ── Exibe resultado se já gerado ───────────────────────────────────────
    if "ultimo_html_pcmso" in st.session_state:
        st.markdown("---")
        st.markdown("### 📊 Prévia da Matriz de Exames")
        st.dataframe(
            st.session_state["df_pcmso_preview"],
            use_container_width=True, hide_index=True,
        )

        col1, col2 = st.columns([3, 1])
        with col1:
            nome_projeto_pcmso = st.text_input(
                "Nome da Empresa / Projeto (para salvar):", key="nome_pcmso"
            )
        with col2:
            st.write(""); st.write("")
            if st.button(
                "💾 Gravar no Banco de Dados",
                key="btn_salvar_pcmso",
                use_container_width=True
            ):
                if nome_projeto_pcmso:
                    salvar_historico(
                        nome_projeto_pcmso + " (PCMSO)",
                        st.session_state["ultimo_html_pcmso"]
                    )
                    st.success("✅ Salvo com sucesso!")
                else:
                    st.error("Digite o nome do projeto antes de salvar.")

        aba_preview, aba_download = st.tabs(["👁 Pré-visualizar", "📄 Baixar (.doc)"])
        with aba_preview:
            components.html(
                st.session_state["ultimo_html_pcmso"],
                height=600, scrolling=True
            )
        with aba_download:
            st.download_button(
                "⬇ Baixar Matriz PCMSO",
                data=st.session_state["ultimo_html_pcmso"].encode("utf-8"),
                file_name="PCMSO_Matriz_Exames.doc",
                mime="application/msword",
                use_container_width=True,
            )

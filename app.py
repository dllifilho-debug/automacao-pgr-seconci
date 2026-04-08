import streamlit as st
import streamlit.components.v1 as components
import pdfplumber
import re
import pandas as pd
import sqlite3
from datetime import datetime
import os
import requests
import json
import base64

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E CSS (UI/UX)
# ==========================================
st.set_page_config(page_title="Automação SST - Seconci GO", layout="wide", page_icon="🛡️")

css_personalizado = """
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .stButton > button { background-color: #084D22; color: white; border-radius: 8px; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: all 0.3s ease; font-weight: 600; padding: 0.5rem 1rem; }
    .stButton > button:hover { background-color: #1AA04B; color: white; box-shadow: 0 6px 12px rgba(0,0,0,0.15); transform: translateY(-2px); border-color: #1AA04B; }
    h1, h2, h3 { color: #084D22 !important; }
    [data-testid="stSidebar"] { background-color: #F4F8F5; border-right: 1px solid #E0ECE4; }
    [data-testid="stFileUploadDropzone"] { border: 2px dashed #1AA04B; border-radius: 12px; background-color: #FAFFFA; transition: all 0.3s ease; }
    [data-testid="stFileUploadDropzone"]:hover { background-color: #F0FDF4; border-color: #084D22; }
    .stAlert { border-radius: 8px; border-left: 5px solid #084D22; }
    [data-testid="stDataFrame"] { border: 1px solid #E0ECE4; border-radius: 8px; overflow: hidden; }
</style>
"""
st.markdown(css_personalizado, unsafe_allow_html=True)

# ==========================================
# CONFIGURAÇÃO DA IA (CONEXÃO DIRETA REST)
# ==========================================
# Blindagem: Removemos espaços, quebras de linha ou aspas acidentais que possam vir do cofre.
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
    conn.commit()
    conn.close()

init_db()

# ==========================================
# BARRA LATERAL (MENU E HISTÓRICO)
# ==========================================
if os.path.exists("logo.png"): st.sidebar.image("logo.png", width="stretch")
elif os.path.exists("logo.jpg"): st.sidebar.image("logo.jpg", width="stretch")
else: st.sidebar.markdown("<h2 style='text-align: center; color: #084D22;'>SECONCI-GO</h2>", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.title("🧩 Módulos do Sistema")
modulo_selecionado = st.sidebar.radio("Selecione a funcionalidade:", ["1️⃣ Engenharia: FISPQ / FDS ➡️ PGR", "2️⃣ Medicina: PGR ➡️ PCMSO"])

st.sidebar.markdown("---")
st.sidebar.title("📂 Histórico de Laudos")
conn = sqlite3.connect('seconci_banco_dados.db')
df_historico = pd.read_sql_query("SELECT id, nome_projeto, data_salvamento FROM historico_laudos ORDER BY id DESC", conn)
conn.close()

historico_selecionado = None
if not df_historico.empty:
    opcoes_historico = ["Selecione um projeto salvo..."] + [f"{row['id']} - {row['nome_projeto']} ({row['data_salvamento']})" for _, row in df_historico.iterrows()]
    selecao = st.sidebar.selectbox("Carregar projeto antigo:", opcoes_historico)
    if selecao != "Selecione um projeto salvo...":
        id_selecionado = int(selecao.split(" - ")[0])
        conn = sqlite3.connect('seconci_banco_dados.db')
        cursor = conn.cursor()
        cursor.execute("SELECT html_relatorio FROM historico_laudos WHERE id = ?", (id_selecionado,))
        historico_selecionado = cursor.fetchone()[0]
        conn.close()
        st.sidebar.success("✅ Projeto carregado.")
else: st.sidebar.write("Nenhum projeto salvo ainda.")

# ==========================================
# DICIONÁRIOS FASE 1 (ENGENHARIA E QUÍMICA)
# ==========================================
dicionario_h = {
    "H315": {"desc": "Provoca irritação à pele", "sev": 1, "epi": "Luvas de proteção e vestimenta"},
    "H319": {"desc": "Provoca irritação ocular grave", "sev": 1, "epi": "Óculos de proteção contra respingos"},
    "H336": {"desc": "Pode provocar sonolência ou vertigem", "sev": 1, "epi": "Local ventilado; máscara se necessário"},
    "H317": {"desc": "Reações alérgicas na pele", "sev": 2, "epi": "Luvas (nitrílica/PVC) e manga longa"},
    "H335": {"desc": "Irritação das vias respiratórias", "sev": 2, "epi": "Proteção respiratória (Filtro específico)"},
    "H302": {"desc": "Nocivo em caso de ingestão", "sev": 2, "epi": "Higiene rigorosa; luvas adequadas"},
    "H312": {"desc": "Nocivo em contato com a pele", "sev": 2, "epi": "Luvas impermeáveis e avental"},
    "H332": {"desc": "Nocivo se inalado", "sev": 2, "epi": "Máscara respiratória adequada (PFF2/VO)"},
    "H314": {"desc": "Queimadura severa à pele e dano", "sev": 4, "epi": "Traje químico, luvas longas e botas"},
    "H318": {"desc": "Lesões oculares graves", "sev": 3, "epi": "Óculos ampla visão / protetor facial"},
    "H301": {"desc": "Tóxico em caso de ingestão", "sev": 4, "epi": "Higiene rigorosa; luvas"},
    "H311": {"desc": "Tóxico em contato com a pele", "sev": 4, "epi": "Luvas, avental impermeável e botas"},
    "H331": {"desc": "Tóxico se inalado", "sev": 4, "epi": "Respirador facial inteiro"},
    "H334": {"desc": "Sintomas de asma ou dificuldades respiratórias", "sev": 4, "epi": "Respirador facial inteiro (Filtro P3)"},
    "H372": {"desc": "Danos aos órgãos (exposição prolongada/repetida)", "sev": 4, "epi": "Proteção respiratória e dérmica estrita"},
    "H373": {"desc": "Pode provocar danos aos órgãos (exp. repetida)", "sev": 4, "epi": "Avaliar via; EPIs combinados obrigatórios"},
    "H300": {"desc": "Fatal em caso de ingestão", "sev": 5, "epi": "Isolamento total; Higiene extrema"},
    "H310": {"desc": "Fatal em contato com a pele", "sev": 5, "epi": "Traje encapsulado nível A/B"},
    "H330": {"desc": "Fatal se inalado", "sev": 5, "epi": "Equipamento de Respiração Autônoma (EPR)"},
    "H340": {"desc": "Pode provocar defeitos genéticos", "sev": 5, "epi": "Isolamento; Traje químico e EPR"},
    "H350": {"desc": "Pode provocar câncer", "sev": 5, "epi": "Isolamento; Traje químico e EPR completo"},
    "H351": {"desc": "Suspeito de provocar câncer", "sev": 4, "epi": "Proteção respiratória e dérmica máxima"},
    "H360": {"desc": "Pode prejudicar a fertilidade ou o feto", "sev": 5, "epi": "Afastamento de gestantes; EPI máximo"},
    "H370": {"desc": "Provoca danos aos órgãos", "sev": 5, "epi": "EPI máximo conforme via de exposição"}
}

matriz_oficial = {
    (1,1): "TRIVIAL", (1,2): "TRIVIAL", (1,3): "TOLERÁVEL", (1,4): "TOLERÁVEL", (1,5): "MODERADO",
    (2,1): "TRIVIAL", (2,2): "TOLERÁVEL", (2,3): "MODERADO", (2,4): "MODERADO", (2,5): "SUBSTANCIAL",
    (3,1): "TOLERÁVEL", (3,2): "TOLERÁVEL", (3,3): "MODERADO", (3,4): "SUBSTANCIAL", (3,5): "SUBSTANCIAL",
    (4,1): "TOLERÁVEL", (4,2): "MODERADO", (4,3): "SUBSTANCIAL", (4,4): "INTOLERÁVEL", (4,5): "INTOLERÁVEL",
    (5,1): "MODERADO", (5,2): "MODERADO", (5,3): "SUBSTANCIAL", (5,4): "INTOLERÁVEL", (5,5): "INTOLERÁVEL"
}

acoes_requeridas = {
    "TRIVIAL": "Manter controles existentes; monitoramento periódico.",
    "TOLERÁVEL": "Manter controles. Considerar melhorias.",
    "MODERADO": "Implantar controles. EPI e monitoramento PCMSO.",
    "SUBSTANCIAL": "Trabalho não deve iniciar sem redução do risco.",
    "INTOLERÁVEL": "TRABALHO PROIBIDO. Risco grave e iminente."
}

texto_sev = {1: "1-LEVE", 2: "2-BAIXA", 3: "3-MODERADA", 4: "4-ALTA", 5: "5-EXTREMA"}

dicionario_cas = {
    "108-88-3": {"agente": "Tolueno", "nr15_lt": "78 ppm ou 290 mg/m³", "nr09_acao": "39 ppm ou 145 mg/m³", "nr07_ibe": "o-Cresol ou Ácido Hipúrico", "dec_3048": "25 anos (Linha 1.0.19)", "esocial_24": "01.19.036"},
    "1330-20-7": {"agente": "Xileno", "nr15_lt": "78 ppm ou 340 mg/m³", "nr09_acao": "39 ppm ou 170 mg/m³", "nr07_ibe": "Ácidos Metilhipúricos", "dec_3048": "25 anos (Linha 1.0.19)", "esocial_24": "01.19.036"},
    "71-43-2": {"agente": "Benzeno", "nr15_lt": "VRT-MPT (Cancerígeno)", "nr09_acao": "Avaliação Qualitativa", "nr07_ibe": "Ácido trans,trans-mucônico", "dec_3048": "25 anos (Linha 1.0.3)", "esocial_24": "01.01.006"},
    "67-64-1": {"agente": "Acetona", "nr15_lt": "780 ppm ou 1870 mg/m³", "nr09_acao": "390 ppm ou 935 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "64-17-5": {"agente": "Etanol (Álcool Etílico)", "nr15_lt": "780 ppm ou 1480 mg/m³", "nr09_acao": "390 ppm ou 740 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "78-93-3": {"agente": "Metiletilcetona (MEK)", "nr15_lt": "155 ppm ou 460 mg/m³", "nr09_acao": "77.5 ppm ou 230 mg/m³", "nr07_ibe": "MEK na Urina", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "110-54-3": {"agente": "n-Hexano", "nr15_lt": "50 ppm ou 176 mg/m³", "nr09_acao": "25 ppm ou 88 mg/m³", "nr07_ibe": "2,5-Hexanodiona", "dec_3048": "25 anos (Linha 1.0.19)", "esocial_24": "01.19.014"},
    "14808-60-7": {"agente": "Sílica Cristalina (Quartzo)", "nr15_lt": "Anexo 12", "nr09_acao": "50% do L.T.", "nr07_ibe": "Raio-X (OIT) e Espirometria", "dec_3048": "25 anos (Linha 1.0.18)", "esocial_24": "01.18.001"},
    "1332-21-4": {"agente": "Asbesto / Amianto", "nr15_lt": "0,1 f/cm³", "nr09_acao": "0,05 f/cm³", "nr07_ibe": "Raio-X (OIT) e Espirometria", "dec_3048": "20 anos (Linha 1.0.2)", "esocial_24": "01.02.001"},
    "7439-92-1": {"agente": "Chumbo (Fumos)", "nr15_lt": "0,1 mg/m³", "nr09_acao": "0,05 mg/m³", "nr07_ibe": "Chumbo no sangue e ALA-U", "dec_3048": "25 anos (Linha 1.0.8)", "esocial_24": "01.08.001"},
    "141-78-6": {"agente": "Acetato de Etila", "nr15_lt": "310 ppm ou 1090 mg/m³", "nr09_acao": "155 ppm ou 545 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "13463-67-7": {"agente": "Dióxido de Titânio", "nr15_lt": "Ausente NR-15 (ACGIH: 10 mg/m³)", "nr09_acao": "5 mg/m³ (Ref. ACGIH)", "nr07_ibe": "Avaliação Clínica / Raio-X", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "1333-86-4": {"agente": "Negro de Fumo (Carbon Black)", "nr15_lt": "Ausente NR-15 (ACGIH: 3 mg/m³)", "nr09_acao": "1,5 mg/m³ (Ref. ACGIH)", "nr07_ibe": "Avaliação Clínica / Espirometria", "dec_3048": "Avaliar Anexo IV (Hidrocarbonetos)", "esocial_24": "01.01.000"},
    "8052-41-3": {"agente": "Aguarrás Mineral (Stoddard Solvent)", "nr15_lt": "Ausente NR-15 (ACGIH: 100 ppm)", "nr09_acao": "50 ppm (Ref. ACGIH)", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "1309-37-1": {"agente": "Óxido de Ferro", "nr15_lt": "Ausente NR-15 (ACGIH: 5 mg/m³)", "nr09_acao": "2,5 mg/m³ (Ref. ACGIH)", "nr07_ibe": "Raio-X (OIT)", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "136-51-6": {"agente": "Octoato de Cálcio", "nr15_lt": "Não Estabelecido", "nr09_acao": "Não Estabelecido", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"}
}

# Expandido com Físicos, Biológicos, Ergonômicos e Acidentes
dicionario_campo = {
    # Riscos Físicos
    "Físico: Ruído Contínuo/Intermitente": {
        "agente": "Ruído Contínuo ou Intermitente", "nr15_lt": "85 dB(A)", "nr09_acao": "80 dB(A)", "nr07_ibe": "Audiometria", "dec_3048": "25 anos (Linha 2.0.1)", "esocial_24": "02.01.001",
        "perigo": "Exposição a níveis elevados de pressão sonora", "sev": 3, "epi": "Protetor Auditivo (Atenuação adequada)"
    },
    "Físico: Vibração de Mãos e Braços (VMB)": {
        "agente": "Vibração de Mãos e Braços (VMB)", "nr15_lt": "5,0 m/s²", "nr09_acao": "2,5 m/s²", "nr07_ibe": "Avaliação Clínica", "dec_3048": "25 anos (Linha 2.0.2)", "esocial_24": "02.01.002",
        "perigo": "Transmissão de energia mecânica para o sistema mão-braço", "sev": 3, "epi": "Luvas antivibração / Revezamento"
    },
    "Físico: Vibração de Corpo Inteiro (VCI)": {
        "agente": "Vibração de Corpo Inteiro (VCI)", "nr15_lt": "1,1 m/s² ou 21,0 m/s¹.75", "nr09_acao": "0,5 m/s² ou 9,1 m/s¹.75", "nr07_ibe": "Avaliação Clínica", "dec_3048": "25 anos (Linha 2.0.2)", "esocial_24": "02.01.003",
        "perigo": "Transmissão de energia mecânica para o corpo inteiro", "sev": 3, "epi": "Assentos com amortecimento / Revezamento"
    },
    # Riscos Biológicos
    "Biológico: Esgoto / Fossas": {
        "agente": "Microorganismos - Esgoto / Fossas", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.005",
        "perigo": "Exposição a agentes biológicos infectocontagiosos", "sev": 4, "epi": "Luvas, Botas de PVC, Proteção facial"
    },
    "Biológico: Lixo Urbano": {
        "agente": "Microorganismos - Lixo Urbano", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.007",
        "perigo": "Contato com resíduos e agentes biológicos", "sev": 4, "epi": "Luvas anticorte, Botas, Uniforme impermeável"
    },
    "Biológico: Estab. Saúde": {
        "agente": "Microorganismos - Área da Saúde", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos (Linha 3.0.1)", "esocial_24": "03.01.001",
        "perigo": "Exposição a patógenos em ambiente de saúde", "sev": 4, "epi": "Luvas de procedimento, Máscara, Avental"
    },
    # Riscos Ergonômicos (PGR Obrigatório, eSocial Dispensado)
    "Ergonômico: Postura Inadequada": {
        "agente": "Fator Ergonômico - Postura", "nr15_lt": "N/A (NR-17)", "nr09_acao": "Avaliação AEP/AET", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Exigência de postura inadequada ou prolongada", "sev": 2, "epi": "Medidas Administrativas / Mobiliário Adequado"
    },
    "Ergonômico: Levantamento/Transporte de Peso": {
        "agente": "Fator Ergonômico - Levantamento de Peso", "nr15_lt": "N/A (NR-17)", "nr09_acao": "Avaliação AEP/AET", "nr07_ibe": "Avaliação Clínica / Osteomuscular", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Esforço físico intenso e levantamento manual de cargas", "sev": 3, "epi": "Auxílio Mecânico / Treinamento"
    },
    # Riscos de Acidentes / Mecânicos (PGR Obrigatório, eSocial Dispensado)
    "Acidente: Queda de Altura": {
        "agente": "Risco de Acidente - Altura", "nr15_lt": "N/A (NR-35)", "nr09_acao": "N/A", "nr07_ibe": "Protocolo Trabalho em Altura", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Trabalho executado acima de 2 metros do nível inferior", "sev": 4, "epi": "Cinturão de Segurança, Talabarte, Capacete com Jugular"
    },
    "Acidente: Choque Elétrico": {
        "agente": "Risco de Acidente - Eletricidade", "nr15_lt": "N/A (NR-10)", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica / ECG", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Contato direto ou indireto com partes energizadas", "sev": 5, "epi": "Luvas Isolantes, Vestimenta ATPV, Capacete Classe B"
    },
    "Acidente: Máquinas e Equipamentos": {
        "agente": "Risco de Acidente - Partes Móveis", "nr15_lt": "N/A (NR-12)", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)",
        "perigo": "Operação de máquinas com risco de corte ou esmagamento", "sev": 4, "epi": "Luvas de Proteção, Óculos, Botas de Segurança"
    }
}

# ==========================================
# DICIONÁRIOS FASE 2 (CLÍNICA E PCMSO)
# ATUALIZADO COM MATRIZES DA DRA. PATRÍCIA (LABORATÓRIO E CONSTRUÇÃO)
# ==========================================
matriz_risco_exame = {
    "TOLUENO": {"exame": "Ortocresol na Urina", "periodico": "6 MESES"},
    "RUÍDO": {"exame": "Audiometria", "periodico": "12 MESES"},
    "SÍLICA": {"exame": "Raio-X de Tórax (OIT) + Espirometria", "periodico": "12 a 24 MESES"},
    "VIBRAÇÃO": {"exame": "Avaliação Clínica e Osteomuscular", "periodico": "12 MESES"},
    "POEIRA": {"exame": "Raio-X de Tórax (OIT)", "periodico": "12 MESES"},
    # Matriz Laboratorial / Infecciosa
    "BIOLÓGIC": {"exame": "HBsAg / Anti-HBs / Anti-HCV", "periodico": "12 a 24 MESES"},
    "SANGUE": {"exame": "HBsAg / Anti-HBs / Anti-HCV", "periodico": "12 a 24 MESES"},
    "VÍRUS": {"exame": "HBsAg / Anti-HBs / Anti-HCV", "periodico": "12 a 24 MESES"},
    "BACTÉRIA": {"exame": "HBsAg / Anti-HBs / Anti-HCV", "periodico": "12 a 24 MESES"},
    "QUÍMIC": {"exame": "Hemograma Completo / Avaliação Hepática", "periodico": "12 MESES"}
}

matriz_funcao_exame = {
    "TRABALHO EM ALTURA": [
        {"exame": "Hemograma", "periodicidade": "12 MESES"},
        {"exame": "Glicemia de Jejum", "periodicidade": "12 MESES"},
        {"exame": "Audiometria", "periodicidade": "12 MESES"},
        {"exame": "ECG", "periodicidade": "12 MESES"},
        {"exame": "Acuidade Visual", "periodicidade": "12 MESES"},
        {"exame": "Avaliação Psicossocial", "periodicidade": "12 MESES"}
    ],
    "ENCANADOR": [
        {"exame": "Exame Clínico", "periodicidade": "6 MESES"},
        {"exame": "Audiometria", "periodicidade": "12 MESES"},
        {"exame": "Acuidade Visual", "periodicidade": "12 MESES"},
        {"exame": "ECG", "periodicidade": "12 MESES"},
        {"exame": "Glicemia de Jejum", "periodicidade": "12 MESES"},
        {"exame": "Hemograma", "periodicidade": "12 MESES"},
        {"exame": "Espirometria", "periodicidade": "24 MESES"},
        {"exame": "RX Tórax OIT", "periodicidade": "12 MESES"}
    ],
    "ADMINISTRATIVO": [
        {"exame": "Exame Clínico", "periodicidade": "12 MESES"},
        {"exame": "Acuidade Visual", "periodicidade": "12 MESES"}
    ],
    "RECEPCIONISTA": [
        {"exame": "Exame Clínico", "periodicidade": "12 MESES"}
    ],
    "GESTOR": [
        {"exame": "Exame Clínico", "periodicidade": "12 MESES"}
    ],
    # Matriz Laboratorial (Dra. Patrícia)
    "BIOMÉDICO": [
        {"exame": "HBsAg", "periodicidade": "12 MESES"},
        {"exame": "Anti-HBs", "periodicidade": "24 MESES"},
        {"exame": "Anti-HCV", "periodicidade": "12 MESES"}
    ],
    "LABORATÓRIO": [
        {"exame": "HBsAg", "periodicidade": "12 MESES"},
        {"exame": "Anti-HBs", "periodicidade": "24 MESES"},
        {"exame": "Anti-HCV", "periodicidade": "12 MESES"}
    ],
    "ENFERMAGEM": [
        {"exame": "HBsAg", "periodicidade": "12 MESES"},
        {"exame": "Anti-HBs", "periodicidade": "24 MESES"},
        {"exame": "Anti-HCV", "periodicidade": "12 MESES"}
    ]
}

def processar_pcmso(dados_pgr_json):
    tabela_pcmso = []
    for ghe in dados_pgr_json:
        nome_ghe = ghe.get("ghe", "Sem GHE")
        cargos = ghe.get("cargos", [])
        riscos = ghe.get("riscos_mapeados", [])
        
        for cargo in cargos:
            exames_do_cargo = [{"exame": "Exame Clínico (Anamnese/Físico)", "periodicidade": "12 MESES", "motivo": "NR-07 Básico"}]
            cargo_upper = cargo.upper()
            
            # Cruzamento 1: Por Função/Cargo
            for funcao_chave, exames in matriz_funcao_exame.items():
                if funcao_chave in cargo_upper:
                    exames_do_cargo.extend(exames)
            
            # Cruzamento 2: Por Agente ou Perigo Especifico
            for risco in riscos:
                agente = risco.get("nome_agente", "").upper()
                perigo = risco.get("perigo_especifico", "").upper()
                texto_risco_completo = agente + " " + perigo
                
                for agente_chave, regra in matriz_risco_exame.items():
                    if agente_chave in texto_risco_completo:
                        exames_do_cargo.append({
                            "exame": regra["exame"],
                            "periodicidade": regra["periodico"],
                            "motivo": f"Exposição a {agente_chave}"
                        })
                
                if "ALTURA" in texto_risco_completo:
                     exames_do_cargo.extend(matriz_funcao_exame["TRABALHO EM ALTURA"])
            
            # Remove duplicatas baseadas no nome do exame
            exames_unicos = {v['exame']:v for v in exames_do_cargo}.values()
            
            for ex in exames_unicos:
                tabela_pcmso.append({
                    "GHE / Setor": nome_ghe,
                    "Cargo": cargo,
                    "Exame Clínico/Complementar": ex["exame"],
                    "Periodicidade": ex["periodicidade"],
                    "Justificativa Legal / Risco": ex.get("motivo", f"Protocolo Função")
                })
                
    return pd.DataFrame(tabela_pcmso)

# ==========================================
# GERADORES DE HTML (FORMATADOS PARA MICROSOFT WORD)
# ==========================================
def gerar_html_anexo(resultados_pgr, resultados_medicos):
    html_content = """<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset="utf-8">
    <style>
      body { font-family: 'Arial', sans-serif; font-size: 10pt; color: #000; }
      .anexo-header { background-color: #084D22; color: #FFF; padding: 14px 20px; font-size: 13pt; font-weight: bold; margin-bottom: 20px; text-align: center; }
      .funcao-card { border: 1px solid #084D22; margin-bottom: 20px; }
      .funcao-card-header { background-color: #084D22; padding: 10px; font-weight: bold; color: #FFF; font-size: 10pt; }
      .funcao-mini-table { width: 100%; border-collapse: collapse; font-size: 9pt; margin: 8px 0; }
      .funcao-mini-table th { background-color: #0F823B; color: #FFF; padding: 8px; text-align: left; border: 1px solid #000; }
      .funcao-mini-table td { padding: 5px; border: 1px solid #000; }
      h4 { color: #084D22; margin: 15px 0 5px 0; font-size: 10pt; }
    </style></head><body>
    <div class='anexo-header'>ANEXO I - INVENTÁRIO DE RISCOS E ENQUADRAMENTO PREVIDENCIÁRIO</div>
    """
    df_pgr = pd.DataFrame(resultados_pgr)
    df_med = pd.DataFrame(resultados_medicos)
    ghes = set(df_pgr['GHE'].unique().tolist() + df_med['GHE'].unique().tolist() if not df_med.empty else [])
    
    for ghe in sorted(ghes):
        html_content += f"<div class='funcao-card'><div class='funcao-card-header'>GHE: {ghe}</div><div>"
        
        pgr_ghe = df_pgr[df_pgr['GHE'] == ghe] if not df_pgr.empty else pd.DataFrame()
        if not pgr_ghe.empty:
            html_content += "<h4>Inventário de Risco (NR-01)</h4><table class='funcao-mini-table'><thead><tr><th>Origem / FISPQ / FDS</th><th>Perigo Identificado</th><th>Sev.</th><th>Prob.</th><th>Nível de Risco</th><th>EPI Recomendado (NR-06)</th></tr></thead><tbody>"
            for _, row in pgr_ghe.iterrows():
                html_content += f"<tr><td>{row['Arquivo Origem']}</td><td>{row['Código GHS']} {row['Perigo Identificado']}</td><td>{row['Severidade']}</td><td>{row['Probabilidade']}</td><td>{row['NÍVEL DE RISCO']}</td><td>{row['EPI (NR-06)']}</td></tr>"
            html_content += "</tbody></table>"
            
        med_ghe = df_med[df_med['GHE'] == ghe] if not df_med.empty else pd.DataFrame()
        if not med_ghe.empty:
            html_content += "<h4>Diretrizes Médicas e Previdenciárias</h4><table class='funcao-mini-table'><thead><tr><th>Cód / CAS</th><th>Agente</th><th>Lim. Tol. (NR-15)</th><th>Nível Ação (NR-09)</th><th>Exame/IBE (NR-07)</th><th>Dec 3048</th><th>eSocial</th></tr></thead><tbody>"
            for _, row in med_ghe.iterrows():
                html_content += f"<tr><td>{row['Nº CAS']}</td><td>{row['Agente Químico']}</td><td>{row['Lim. Tolerância (NR-15)']}</td><td>{row['Nível de Ação (NR-09)']}</td><td>{row['IBE (NR-07)']}</td><td>{row['Dec 3048']}</td><td>{row['eSocial']}</td></tr>"
            html_content += "</tbody></table>"
            
        html_content += "</div></div>"
    html_content += "</body></html>"
    return html_content

def gerar_html_pcmso(df_pcmso):
    html = """<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset="utf-8">
    <style>
      body { font-family: 'Arial', sans-serif; color: #000000; }
      .header { background-color: #084D22; color: #FFFFFF; padding: 14px; text-align: center; font-weight: bold; font-size: 16px; margin-bottom: 20px;}
      table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 13px; }
      th { background-color: #1AA04B; color: #FFFFFF; padding: 12px 8px; text-align: left; border: 1px solid #000000; }
      td { border: 1px solid #000000; padding: 10px 8px; }
    </style></head><body>
    <div class='header'>MATRIZ DE EXAMES - PCMSO (GERADO VIA IA)</div>
    <table><tr><th>GHE / Setor</th><th>Cargo</th><th>Exame Clínico / Complementar</th><th>Periodicidade</th><th>Justificativa / Agente</th></tr>
    """
    for _, row in df_pcmso.iterrows():
        html += f"<tr><td><strong>{row['GHE / Setor']}</strong></td><td>{row['Cargo']}</td><td>{row['Exame Clínico/Complementar']}</td><td>{row['Periodicidade']}</td><td>{row['Justificativa Legal / Risco']}</td></tr>"
    html += "</table></body></html>"
    return html

# ==========================================
# CORPO PRINCIPAL
# ==========================================
st.title("Sistema Integrado SST - Seconci GO 🚀")

if historico_selecionado:
    st.markdown("### 🗄️ Visualizando Relatório do Histórico")
    aba_preview, aba_download = st.tabs(["👁️ Pré-visualizar", "📄 Baixar em Word (.doc)"])
    with aba_preview: components.html(historico_selecionado, height=700, scrolling=True)
    with aba_download: st.download_button("Baixar Relatório", data=historico_selecionado.encode('utf-8'), file_name="Relatorio_Historico.doc", mime="application/msword")

# ==========================================
# MÓDULO 1: ENGENHARIA (FISPQS / FDS -> PGR)
# ==========================================
elif "1️⃣" in modulo_selecionado:
    st.header("Módulo de Engenharia: Extrator de FISPQs / FDS")
    
    arquivos_fispq = st.file_uploader("Insira as FISPQs / FDS em PDF", type=["pdf"], accept_multiple_files=True)
    textos_pdfs = {}
    df_editado = pd.DataFrame()
    ghe_opcoes = ["Nenhum GHE definido"]

    if arquivos_fispq:
        with st.spinner("Lendo o conteúdo dos PDFs..."):
            for arquivo in arquivos_fispq:
                with pdfplumber.open(arquivo) as pdf:
                    texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
                    textos_pdfs[arquivo.name] = texto

        st.markdown("### 2️⃣ Definição de GHEs Químicos")
        nomes_arquivos = [arq.name for arq in arquivos_fispq]
        dados_iniciais = [{"GHE": "GHE 01 - Digite a Função", "Arquivo FISPQ/FDS": nome, "Probabilidade": 3} for nome in nomes_arquivos]
        df_mapeamento = pd.DataFrame(dados_iniciais)
        
        df_editado = st.data_editor(
            df_mapeamento, num_rows="dynamic",
            column_config={
                "GHE": st.column_config.TextColumn("Nome do GHE", required=True),
                "Arquivo FISPQ/FDS": st.column_config.SelectboxColumn("Arquivo (FISPQ/FDS)", options=nomes_arquivos, required=True),
                "Probabilidade": st.column_config.NumberColumn("Prob.", min_value=1, max_value=5, required=True)
            }, width="stretch"
        )
        ghe_opcoes = df_editado["GHE"].unique().tolist() if not df_editado.empty else ["Nenhum GHE definido"]

    st.markdown("### 3️⃣ Avaliações de Campo: Físicos, Biológicos, Ergonômicos e Acidentes")
    agentes_opcoes = list(dicionario_campo.keys())
    df_fis_bio_inicial = pd.DataFrame([{"GHE": ghe_opcoes[0], "Agente": agentes_opcoes[0], "Probabilidade": 3}])

    df_fis_bio_editado = st.data_editor(
        df_fis_bio_inicial, num_rows="dynamic",
        column_config={
            "GHE": st.column_config.SelectboxColumn("GHE de Destino", options=ghe_opcoes, required=True),
            "Agente": st.column_config.SelectboxColumn("Agente / Fator de Risco", options=agentes_opcoes, required=True),
            "Probabilidade": st.column_config.NumberColumn("Prob.", min_value=1, max_value=5, required=True)
        }, width="stretch"
    )

    if st.button("🪄 Processar GHEs e Gerar Relatório", width="stretch", type="primary"):
        with st.spinner("Consolidando avaliações..."):
            resultados_pgr, resultados_medicos = [], []
            
            if not df_editado.empty:
                for index, row in df_editado.iterrows():
                    nome_ghe, nome_arq, v_prob = row["GHE"], row["Arquivo FISPQ/FDS"], int(row["Probabilidade"])
                    if nome_arq in textos_pdfs:
                        texto_completo = textos_pdfs[nome_arq]
                        cas_encontrados_linha = list(set(re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', texto_completo)))
                        for cas in cas_encontrados_linha:
                            dados_med = dicionario_cas.get(cas, {
                                "agente": "⚠️ NÃO MAPEADO", "nr15_lt": "REVISÃO", "nr09_acao": "REVISÃO", 
                                "nr07_ibe": "REVISÃO", "dec_3048": "REVISÃO", "esocial_24": "REVISÃO"
                            })
                            resultados_medicos.append({
                                "GHE": nome_ghe, "Arquivo Origem": nome_arq, "Nº CAS": cas,
                                "Agente Químico": dados_med["agente"], "Lim. Tolerância (NR-15)": dados_med["nr15_lt"],
                                "Nível de Ação (NR-09)": dados_med["nr09_acao"], "IBE (NR-07)": dados_med.get("nr07_ibe", "N/A"),
                                "Dec 3048": dados_med.get("dec_3048", "Não Enquadrado"), "eSocial": dados_med.get("esocial_24", "09.01.001")
                            })
                            
                        h_encontradas_linha = list(set(re.findall(r'H\d{3}', texto_completo)))
                        for codigo in h_encontradas_linha:
                            if codigo in dicionario_h:
                                dados_h = dicionario_h[codigo]
                                nivel_risco = matriz_oficial.get((dados_h["sev"], v_prob), "N/A")
                                resultados_pgr.append({
                                    "GHE": nome_ghe, "Arquivo Origem": nome_arq, "Código GHS": codigo,
                                    "Perigo Identificado": dados_h["desc"], "Severidade": texto_sev.get(dados_h["sev"], str(dados_h["sev"])),
                                    "Probabilidade": str(v_prob), "NÍVEL DE RISCO": nivel_risco,
                                    "Ação Requerida": acoes_requeridas.get(nivel_risco, "Manual"), "EPI (NR-06)": dados_h["epi"]
                                })

            if not df_fis_bio_editado.empty and df_fis_bio_editado["GHE"].iloc[0] != "Nenhum GHE definido":
                for index, row in df_fis_bio_editado.iterrows():
                    nome_ghe, nome_agente, v_prob = row["GHE"], row["Agente"], int(row["Probabilidade"])
                    if nome_agente in dicionario_campo:
                        dados_fis = dicionario_campo[nome_agente]
                        resultados_medicos.append({
                            "GHE": nome_ghe, "Arquivo Origem": "Avaliação de Campo", "Nº CAS": "-",
                            "Agente Químico": dados_fis["agente"], "Lim. Tolerância (NR-15)": dados_fis["nr15_lt"],
                            "Nível de Ação (NR-09)": dados_fis["nr09_acao"], "IBE (NR-07)": dados_fis["nr07_ibe"],
                            "Dec 3048": dados_fis["dec_3048"], "eSocial": dados_fis["esocial_24"]
                        })
                        nivel_risco = matriz_oficial.get((dados_fis["sev"], v_prob), "N/A")
                        resultados_pgr.append({
                            "GHE": nome_ghe, "Arquivo Origem": "Avaliação de Campo", "Código GHS": "-",
                            "Perigo Identificado": dados_fis["perigo"], "Severidade": texto_sev.get(dados_fis["sev"], str(dados_fis["sev"])),
                            "Probabilidade": str(v_prob), "NÍVEL DE RISCO": nivel_risco,
                            "Ação Requerida": acoes_requeridas.get(nivel_risco, "Manual"), "EPI (NR-06)": dados_fis["epi"]
                        })

            if resultados_pgr or resultados_medicos:
                html_final = gerar_html_anexo(resultados_pgr, resultados_medicos)
                st.session_state['ultimo_html'] = html_final
                st.success("✅ Relatório Consolidado Gerado!")

    if 'ultimo_html' in st.session_state:
        col1, col2 = st.columns([3, 1])
        with col1: nome_projeto = st.text_input("Nome da Empresa:")
        with col2:
            st.write(""); st.write("")
            if st.button("Gravar no Banco de Dados", width="stretch") and nome_projeto:
                conn = sqlite3.connect('seconci_banco_dados.db')
                c = conn.cursor()
                c.execute("INSERT INTO historico_laudos (nome_projeto, data_salvamento, html_relatorio) VALUES (?, ?, ?)", 
                          (nome_projeto, datetime.now().strftime("%d/%m/%Y %H:%M"), st.session_state['ultimo_html']))
                conn.commit()
                conn.close()
                st.success("Salvo com sucesso!")

        aba_preview, aba_download = st.tabs(["👁️ Pré-visualizar", "📄 Baixar (.doc)"])
        with aba_preview: components.html(st.session_state['ultimo_html'], height=500, scrolling=True)
        with aba_download: st.download_button("Baixar Relatório", st.session_state['ultimo_html'].encode('utf-8'), "PGR_Fase1.doc")

# ==========================================
# MÓDULO 2: MEDICINA (PGR -> PCMSO)
# ==========================================
elif "2️⃣" in modulo_selecionado:
    st.header("🩺 Módulo Médico: Importador de PGR e Gerador de PCMSO")
    st.info("Faça o upload do Inventário de Riscos (PGR) de terceiros. A IA fará a leitura visual e o cruzamento com as matrizes da NR-07 (Laboratorial e Construção).")
    
    with st.container():
        arquivo_pgr = st.file_uploader("Arraste o PDF do PGR aqui", type=["pdf"])
        
        if arquivo_pgr:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚀 Extrair Riscos e Gerar PCMSO", type="primary", use_container_width=True):
                with st.spinner("Motor IA analisando as imagens do documento e cruzando protocolos médicos... Isso pode levar até 1 minuto."):
                    
                    # 1. EMPACOTA O PDF (Base64)
                    pdf_bytes = arquivo_pgr.getvalue()
                    pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
                    
                    # 2. PROMPT AGRESSIVO (Instrui a IA a padronizar os nomes dos riscos)
                    prompt_extracao = """
                    Você é um médico do trabalho e engenheiro de segurança. Analise este documento PDF (Inventário de Riscos Ocupacionais).
                    Sua missão é CAÇAR qualquer relação entre Funções/Cargos e Agentes Nocivos (Físicos, Químicos, Biológicos).
                    
                    CRÍTICO: Para cada risco encontrado, tente classificar o "nome_agente" de forma clara (ex: "Risco Biológico", "Vírus e Bactérias", "Produtos Químicos", "Ruído").
                    
                    Se não houver a palavra "GHE", agrupe pelo nome do "Setor" ou da "Função". NUNCA retorne vazio.
                    
                    Retorne EXATAMENTE este formato JSON (uma lista de objetos):
                    [
                      {
                        "ghe": "Nome do Setor, GHE ou Função",
                        "cargos": ["Nome do Cargo 1", "Nome do Cargo 2"],
                        "riscos_mapeados": [
                          {"nome_agente": "Ex: Risco Biológico / Produtos Químicos", "perigo_especifico": "Ex: Exposição a vírus / Manuseio de formol"}
                        ]
                      }
                    ]
                    """
                    
                    try:
                        # 3. AUTO-DISCOVERY DO MODELO
                        url_lista = "https://generativelanguage.googleapis.com/v1beta/models?key=" + CHAVE_API_GOOGLE
                        resp_lista = requests.get(url_lista)
                        
                        if resp_lista.status_code == 200:
                            modelos = resp_lista.json().get('models', [])
                            modelos_texto = [m['name'] for m in modelos if 'generateContent' in m.get('supportedGenerationMethods', [])]
                            
                            modelo_escolhido = None
                            for pref in ['models/gemini-1.5-pro-latest', 'models/gemini-1.5-pro', 'models/gemini-1.5-flash-latest', 'models/gemini-1.5-flash']:
                                if pref in modelos_texto:
                                    modelo_escolhido = pref
                                    break
                            
                            if not modelo_escolhido and modelos_texto:
                                modelo_escolhido = modelos_texto[0]
                                
                            # 4. REQUISIÇÃO MULTIMODAL
                            url_google = "https://generativelanguage.googleapis.com/v1beta/" + modelo_escolhido + ":generateContent?key=" + CHAVE_API_GOOGLE
                            payload = {
                                "contents": [
                                    {
                                        "parts": [
                                            {"text": prompt_extracao},
                                            {
                                                "inlineData": {
                                                    "mimeType": "application/pdf",
                                                    "data": pdf_b64
                                                }
                                            }
                                        ]
                                    }
                                ],
                                "generationConfig": {
                                    "temperature": 0.0,
                                    "responseMimeType": "application/json"
                                }
                            }
                            
                            resposta = requests.post(url_google, headers={'Content-Type': 'application/json'}, json=payload)
                            
                            if resposta.status_code == 200:
                                resultado_texto = resposta.json()['candidates'][0]['content']['parts'][0]['text']
                                resultado_texto = resultado_texto.replace('```json', '').replace('```', '').strip()
                                
                                try:
                                    json_pgr = json.loads(resultado_texto)
                                    
                                    if not json_pgr or len(json_pgr) == 0:
                                        st.error("⚠️ A IA analisou o documento, mas não conseguiu estruturar as tabelas.")
                                        with st.expander("🔍 Ver o que a IA devolveu (Modo Raio-X)"):
                                            st.code(resultado_texto)
                                    else:
                                        st.success(f"✅ Protocolos Médicos Cruzados com Sucesso! (Motor: {modelo_escolhido.split('/')[-1]})")
                                        
                                        # Processa os dados com a nova inteligência da Dra. Patrícia
                                        df_pcmso_gerado = processar_pcmso(json_pgr)
                                        html_final = gerar_html_pcmso(df_pcmso_gerado)
                                        
                                        # INTERFACE COM ABAS (TABS)
                                        aba_dados, aba_preview, aba_download = st.tabs(["📊 Dados Extraídos", "👁️ Pré-visualizar Documento", "📄 Baixar (.doc)"])
                                        
                                        with aba_dados:
                                            st.dataframe(df_pcmso_gerado, use_container_width=True)
                                            
                                        with aba_preview:
                                            components.html(html_final, height=600, scrolling=True)
                                            
                                        with aba_download:
                                            st.info("Clique no botão abaixo para baixar a Matriz PCMSO pronta para o Microsoft Word.")
                                            st.download_button(
                                                label="📄 Baixar Matriz PCMSO em Word (.doc)",
                                                data=html_final.encode('utf-8'), 
                                                file_name="PCMSO_Gerado_Seconci.doc", 
                                                mime="application/msword", 
                                                use_container_width=True
                                            )
                                    
                                except json.JSONDecodeError:
                                    st.error("A IA leu o documento, mas quebrou a formatação dos dados.")
                                    with st.expander("🔍 Ver resposta bruta da IA"):
                                        st.code(resultado_texto)
                            else:
                                 st.error(f"Erro na geração de conteúdo. Detalhes: {resposta.text}")
                        else:
                            st.error(f"Falha ao listar modelos do Google. Erro: {resp_lista.text}")
                            
                    except Exception as e:
                        st.error(f"Falha na comunicação de rede ou processamento: {e}")

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
import io

# ==========================================
# 1. CONFIGURAÇÃO E INTERFACE SaaS (Premium)
# ==========================================
st.set_page_config(page_title="Sistema SST - Seconci GO", layout="wide", page_icon="🛡️")

css_premium = """
<style>
    .block-container { padding-top: 1.5rem; }
    h1, h2, h3, h4 { color: #0f4c23 !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stButton > button { background-color: #0f4c23; color: white; border-radius: 6px; font-weight: bold; border: none; transition: 0.2s; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .stButton > button:hover { background-color: #1a803b; transform: translateY(-1px); box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
    [data-testid="stSidebar"] { background-color: #f8f9fa; border-right: 1px solid #e0e0e0; }
    div[data-testid="metric-container"] { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    [data-testid="stDataFrame"] { border-radius: 8px; border: 1px solid #e0e0e0; }
</style>
"""
st.markdown(css_premium, unsafe_allow_html=True)

try:
    CHAVE_API = str(st.secrets["CHAVE_API_GOOGLE"]).strip().replace('"', '').replace("'", "")
except Exception:
    CHAVE_API = ""
    st.sidebar.error("⚠️ Chave de API não encontrada em st.secrets.")

# ==========================================
# 2. BANCO DE DADOS E DICIONÁRIOS (100% COMPLETOS)
# ==========================================
def init_db():
    conn = sqlite3.connect('seconci_banco_dados.db')
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS historico_laudos (id INTEGER PRIMARY KEY AUTOINCREMENT, nome_projeto TEXT, modulo TEXT, data_salvamento TEXT, html_relatorio TEXT)')
    conn.commit()
    conn.close()

init_db()

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

texto_sev = {1: "1-LEVE", 2: "2-BAIXA", 3: "3-MODERADA", 4: "4-ALTA", 5: "5-EXTREMA"}

dicionario_cas = {
    "108-88-3": {"agente": "Tolueno", "nr15_lt": "78 ppm ou 290 mg/m³", "nr09_acao": "39 ppm", "nr07_ibe": "o-Cresol", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "1330-20-7": {"agente": "Xileno", "nr15_lt": "78 ppm ou 340 mg/m³", "nr09_acao": "39 ppm", "nr07_ibe": "Ác. Metilhipúricos", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "71-43-2": {"agente": "Benzeno", "nr15_lt": "VRT-MPT (Cancerígeno)", "nr09_acao": "Qualitativo", "nr07_ibe": "Ácido trans,trans-mucônico", "dec_3048": "25 anos", "esocial_24": "01.01.006"},
    "67-64-1": {"agente": "Acetona", "nr15_lt": "780 ppm ou 1870 mg/m³", "nr09_acao": "390 ppm", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "64-17-5": {"agente": "Etanol (Álcool Etílico)", "nr15_lt": "780 ppm ou 1480 mg/m³", "nr09_acao": "390 ppm", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "78-93-3": {"agente": "Metiletilcetona (MEK)", "nr15_lt": "155 ppm ou 460 mg/m³", "nr09_acao": "77.5 ppm", "nr07_ibe": "MEK na Urina", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "110-54-3": {"agente": "n-Hexano", "nr15_lt": "50 ppm ou 176 mg/m³", "nr09_acao": "25 ppm", "nr07_ibe": "2,5-Hexanodiona", "dec_3048": "25 anos", "esocial_24": "01.19.014"},
    "14808-60-7": {"agente": "Sílica Cristalina (Quartzo)", "nr15_lt": "Anexo 12", "nr09_acao": "50% do L.T.", "nr07_ibe": "RX OIT / Espirometria", "dec_3048": "25 anos", "esocial_24": "01.18.001"},
    "1332-21-4": {"agente": "Asbesto / Amianto", "nr15_lt": "0,1 f/cm³", "nr09_acao": "0,05 f/cm³", "nr07_ibe": "RX OIT / Espirometria", "dec_3048": "20 anos", "esocial_24": "01.02.001"},
    "7439-92-1": {"agente": "Chumbo (Fumos)", "nr15_lt": "0,1 mg/m³", "nr09_acao": "0,05 mg/m³", "nr07_ibe": "Chumbo no sangue e ALA-U", "dec_3048": "25 anos", "esocial_24": "01.08.001"},
    "65997-15-1": {"agente": "Cimento Portland", "nr15_lt": "10 mg/m³ (Poeira)", "nr09_acao": "5 mg/m³", "nr07_ibe": "RX OIT e Espirometria", "dec_3048": "Não Enquadrado", "esocial_24": "01.18.001"},
    "1317-65-3": {"agente": "Carbonato de Cálcio", "nr15_lt": "10 mg/m³ (Poeira)", "nr09_acao": "5 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "1305-78-8": {"agente": "Óxido de Cálcio", "nr15_lt": "2 mg/m³", "nr09_acao": "1 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12168-85-3": {"agente": "Silicato Tricálcico", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "RX OIT", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "10034-77-2": {"agente": "Silicato Dicálcico", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "RX OIT", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12042-78-3": {"agente": "Aluminato de Cálcio", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "RX OIT", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "1309-48-4": {"agente": "Óxido de Magnésio", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "RX OIT", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "68334-30-5": {"agente": "Óleo Diesel", "nr15_lt": "Qualitativo", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Avaliar Hidrocarbonetos", "esocial_24": "01.01.026"},
    "112-80-1": {"agente": "Ácido Oleico", "nr15_lt": "N/A", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "52-51-7": {"agente": "Bronopol", "nr15_lt": "N/A", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12068-35-8": {"agente": "Silicato (Misto)", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Aval. Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"}
}

dicionario_campo = {
    "Físico: Ruído Contínuo/Intermitente": {"agente": "Ruído Contínuo ou Intermitente", "nr15_lt": "85 dB(A)", "nr09_acao": "80 dB(A)", "nr07_ibe": "Audiometria", "dec_3048": "25 anos", "esocial_24": "02.01.001", "perigo": "Exposição a níveis elevados de pressão sonora", "sev": 3, "epi": "Protetor Auditivo (Atenuação adequada)"},
    "Físico: Vibração de Mãos e Braços (VMB)": {"agente": "Vibração de Mãos e Braços (VMB)", "nr15_lt": "5,0 m/s²", "nr09_acao": "2,5 m/s²", "nr07_ibe": "Avaliação Clínica e Osteomuscular", "dec_3048": "25 anos", "esocial_24": "02.01.002", "perigo": "Transmissão de energia mecânica para o sistema mão-braço", "sev": 3, "epi": "Luvas antivibração / Revezamento"},
    "Físico: Vibração de Corpo Inteiro (VCI)": {"agente": "Vibração de Corpo Inteiro (VCI)", "nr15_lt": "1,1 m/s² ou 21,0 m/s¹.75", "nr09_acao": "0,5 m/s² ou 9,1 m/s¹.75", "nr07_ibe": "Avaliação Clínica e Osteomuscular", "dec_3048": "25 anos", "esocial_24": "02.01.003", "perigo": "Transmissão de energia mecânica para o corpo inteiro", "sev": 3, "epi": "Assentos com amortecimento / Revezamento"},
    "Biológico: Esgoto / Fossas": {"agente": "Microorganismos - Esgoto / Fossas", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos", "esocial_24": "03.01.005", "perigo": "Exposição a agentes biológicos infectocontagiosos", "sev": 4, "epi": "Luvas, Botas de PVC, Proteção facial"},
    "Biológico: Lixo Urbano": {"agente": "Microorganismos - Lixo Urbano", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos", "esocial_24": "03.01.007", "perigo": "Contato com resíduos e agentes biológicos", "sev": 4, "epi": "Luvas anticorte, Botas, Uniforme impermeável"},
    "Biológico: Estab. Saúde": {"agente": "Microorganismos - Área da Saúde", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos", "esocial_24": "03.01.001", "perigo": "Exposição a patógenos em ambiente de saúde", "sev": 4, "epi": "Luvas de procedimento, Máscara, Avental"},
    "Ergonômico: Postura Inadequada": {"agente": "Fator Ergonômico - Postura", "nr15_lt": "N/A (NR-17)", "nr09_acao": "Avaliação AEP/AET", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)", "perigo": "Exigência de postura inadequada ou prolongada", "sev": 2, "epi": "Medidas Administrativas / Mobiliário Adequado"},
    "Ergonômico: Levantamento/Transporte de Peso": {"agente": "Fator Ergonômico - Levantamento de Peso", "nr15_lt": "N/A (NR-17)", "nr09_acao": "Avaliação AEP/AET", "nr07_ibe": "Avaliação Clínica / Osteomuscular", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)", "perigo": "Esforço físico intenso e levantamento manual de cargas", "sev": 3, "epi": "Auxílio Mecânico / Treinamento"},
    "Acidente: Queda de Altura": {"agente": "Risco de Acidente - Altura", "nr15_lt": "N/A (NR-35)", "nr09_acao": "N/A", "nr07_ibe": "Protocolo Trabalho em Altura", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)", "perigo": "Trabalho executado acima de 2 metros do nível inferior", "sev": 4, "epi": "Cinturão de Segurança, Talabarte, Capacete com Jugular"},
    "Acidente: Choque Elétrico": {"agente": "Risco de Acidente - Eletricidade", "nr15_lt": "N/A (NR-10)", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica / ECG", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)", "perigo": "Contato direto ou indireto com partes energizadas", "sev": 5, "epi": "Luvas Isolantes, Vestimenta ATPV, Capacete Classe B"},
    "Acidente: Máquinas e Equipamentos": {"agente": "Risco de Acidente - Partes Móveis", "nr15_lt": "N/A (NR-12)", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (Apenas PGR)", "perigo": "Operação de máquinas com risco de corte ou esmagamento", "sev": 4, "epi": "Luvas de Proteção, Óculos, Botas de Segurança"}
}

# ==========================================
# 3. MOTOR HÍBRIDO DA IA EM CASCATA
# ==========================================
def chamar_api_gemini(prompt, pdf_b64=None):
    if not CHAVE_API: return None
    parts = [{"text": prompt}]
    if pdf_b64: parts.append({"inlineData": {"mimeType": "application/pdf", "data": pdf_b64}})
        
    payload = {"contents": [{"parts": parts}], "generationConfig": {"temperature": 0.0}}
    
    for modelo in ["gemini-1.5-pro-latest", "gemini-1.5-flash-latest", "gemini-1.5-pro", "gemini-1.5-flash"]:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent?key={CHAVE_API}"
        try:
            resp = requests.post(url, headers={'Content-Type': 'application/json'}, json=payload, timeout=20)
            if resp.status_code == 200:
                return resp.json().get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
        except: continue
    return None

def buscar_cas_na_ia(lista_cas):
    if not lista_cas: return {}
    cas_str = ", ".join(lista_cas)
    prompt = f"Atue como Higienista Ocupacional Brasileiro. Identifique limites legais para CAS: {cas_str}. Retorne EXATAMENTE UM JSON, onde a chave é o CAS. Exemplo: {{\"{lista_cas[0]}\": {{\"agente\": \"Nome\", \"nr15_lt\": \"Limite\", \"nr09_acao\": \"Ação\", \"nr07_ibe\": \"Exame\", \"dec_3048\": \"Tempo\", \"esocial_24\": \"Codigo\"}}}}"
    
    resposta = chamar_api_gemini(prompt)
    if resposta:
        try: return json.loads(resposta.replace('```json', '').replace('```', '').strip())
        except: pass
    return {}

# ==========================================
# 4. GERADOR DE HTML (RELATÓRIO)
# ==========================================
def gerar_html_anexo(df_pgr_list, df_med_list):
    df_pgr = pd.DataFrame(df_pgr_list)
    df_med = pd.DataFrame(df_med_list)
    
    html = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'><head><style>body{font-family:Arial;font-size:10pt;} .header{background-color:#0f4c23;color:#FFF;padding:15px;text-align:center;font-weight:bold;} table{width:100%;border-collapse:collapse;margin-bottom:20px;} th{background-color:#1a803b;color:#FFF;padding:8px;border:1px solid #000;} td{padding:5px;border:1px solid #000;}</style></head><body>"
    html += "<div class='header'>INVENTÁRIO DE RISCOS E ENQUADRAMENTO LEGAL</div>"
    
    ghes = set()
    if not df_pgr.empty: ghes.update(df_pgr['GHE'].unique())
    if not df_med.empty: ghes.update(df_med['GHE'].unique())
    
    for ghe in sorted(list(ghes)):
        html += f"<h3 style='color:#0f4c23; border-bottom: 2px solid #0f4c23;'>GHE: {ghe}</h3>"
        
        p_ghe = df_pgr[df_pgr['GHE'] == ghe] if not df_pgr.empty else pd.DataFrame()
        if not p_ghe.empty:
            html += "<h4>Inventário de Perigos (NR-01)</h4><table><tr><th>Origem</th><th>Perigo</th><th>Sev</th><th>Prob</th><th>Risco</th><th>EPI</th></tr>"
            for _, r in p_ghe.iterrows(): html += f"<tr><td>{r['Arquivo Origem']}</td><td>{r['Código']} {r['Desc']}</td><td>{r['Sev']}</td><td>{r['Prob']}</td><td>{r['Risco']}</td><td>{r['EPI']}</td></tr>"
            html += "</table>"
            
        m_ghe = df_med[df_med['GHE'] == ghe] if not df_med.empty else pd.DataFrame()
        if not m_ghe.empty:
            html += "<h4>Enquadramento (NR-15 / eSocial)</h4><table><tr><th>CAS</th><th>Agente</th><th>NR-15</th><th>NR-09</th><th>Exame</th><th>Dec 3048</th><th>eSocial</th></tr>"
            for _, r in m_ghe.iterrows(): html += f"<tr><td>{r['CAS']}</td><td>{r['Agente']}</td><td>{r['NR15']}</td><td>{r['NR09']}</td><td>{r['NR07']}</td><td>{r['Dec3048']}</td><td>{r['eSocial']}</td></tr>"
            html += "</table>"
            
    return html + "</body></html>"

# ==========================================
# 5. ESTRUTURA SAAS - MENU LATERAL
# ==========================================
if os.path.exists("logo.png"): st.sidebar.image("logo.png", use_container_width=True)
elif os.path.exists("logo.jpg"): st.sidebar.image("logo.jpg", use_container_width=True)

st.sidebar.markdown("### 🧭 Navegação")
menu = st.sidebar.radio("", ["📊 Dashboard Inicial", "⚙️ Extrator FISPQ", "📂 Central de Arquivos"])
st.sidebar.markdown("---")
st.sidebar.info("Motor Híbrido Ativado.\nIA em Cascata Operante.")

# ==========================================
# TELA 1: DASHBOARD
# ==========================================
if "Dashboard" in menu:
    st.title("📊 Painel de Controle SST")
    st.markdown("Bem-vindo ao sistema de automação.")
    
    col1, col2, col3 = st.columns(3)
    conn = sqlite3.connect('seconci_banco_dados.db')
    cursor = conn.cursor()
    cursor.execute("SELECT count(name) FROM sqlite_master WHERE type='table' AND name='historico_laudos'")
    total_laudos = pd.read_sql_query("SELECT COUNT(*) FROM historico_laudos", conn).iloc[0,0] if cursor.fetchone()[0] == 1 else 0
    conn.close()
    
    col1.metric("Projetos Processados", f"{total_laudos}")
    col2.metric("Motor IA", "Online / Seguro")
    col3.metric("Banco Local", f"{len(dicionario_cas)} Substâncias")

# ==========================================
# TELA 2: EXTRATOR FISPQ (MOTOR SEGURO)
# ==========================================
elif "Extrator" in menu:
    st.title("⚙️ Extrator Dinâmico de FISPQ")
    
    arquivos_up = st.file_uploader("Arraste as FISPQs aqui (PDF)", type=["pdf"], accept_multiple_files=True)
    ghe_opcoes = ["GHE 01 - Produção"]
    df_editado = pd.DataFrame()
    
    if arquivos_up:
        nomes_arq = [a.name for a in arquivos_up]
        df_editado = st.data_editor(
            pd.DataFrame([{"GHE": "GHE 01 - Produção", "Arquivo": n, "Probabilidade": 3} for n in nomes_arq]), 
            use_container_width=True,
            column_config={"Arquivo": st.column_config.SelectboxColumn(options=nomes_arq)}
        )
        ghe_opcoes = df_editado["GHE"].unique().tolist()

    st.markdown("### 2. Riscos Adicionais (Avaliação de Campo)")
    agentes_campo = list(dicionario_campo.keys())
    df_fis_bio = st.data_editor(
        pd.DataFrame([{"GHE": ghe_opcoes[0], "Agente": agentes_campo[0], "Probabilidade": 3}]), 
        use_container_width=True, num_rows="dynamic",
        column_config={"Agente": st.column_config.SelectboxColumn(options=agentes_campo)}
    )

    if st.button("🚀 Iniciar Análise e Cruzamento", type="primary", use_container_width=True):
        res_pgr, res_med = [], []
        my_bar = st.progress(0, text="Processando...")
        
        if arquivos_up and not df_editado.empty:
            for index, row in df_editado.iterrows():
                nome_ghe, arq_nome, prob = row["GHE"], row["Arquivo"], int(row["Probabilidade"])
                my_bar.progress((index + 1) / (len(df_editado) + 1), text=f"Analisando: {arq_nome}...")
                
                arq_obj = next((f for f in arquivos_up if f.name == arq_nome), None)
                if arq_obj:
                    pdf_bytes = arq_obj.getvalue()
                    texto_pdf = ""
                    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                        for page in pdf.pages: texto_pdf += (page.extract_text() or "") + "\n"
                    
                    cas_encontrados = list(set(re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', texto_pdf)))
                    h_encontrados = list(set(re.findall(r'H\d{3}', texto_pdf)))
                    
                    if not cas_encontrados and not h_encontrados and CHAVE_API:
                        pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
                        txt_visao = chamar_api_gemini("Leia o PDF e liste CAS e códigos GHS. Responda: CAS: X\nH: Y", pdf_b64) or ""
                        cas_encontrados = list(set(re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', txt_visao)))
                        h_encontrados = list(set(re.findall(r'H\d{3}', txt_visao)))

                    cas_desc = [c for c in cas_encontrados if c not in dicionario_cas]
                    ia_dados = buscar_cas_na_ia(cas_desc) if cas_desc else {}
                    
                    for cas in cas_encontrados:
                        c_limpo = cas.strip()
                        d = dicionario_cas.get(c_limpo, ia_dados.get(c_limpo, {"agente": f"Não Mapeado ({c_limpo})", "nr15_lt": "Avaliar NR-15", "nr09_acao": "Avaliar NR-09", "nr07_ibe": "Aval. Clínica", "dec_3048": "Verificar", "esocial_24": "09.01.001"}))
                        res_med.append({"GHE": nome_ghe, "Arquivo Origem": arq_nome, "CAS": c_limpo, "Agente": d["agente"], "NR15": d["nr15_lt"], "NR09": d["nr09_acao"], "NR07": d["nr07_ibe"], "Dec3048": d["dec_3048"], "eSocial": d["esocial_24"]})
                    
                    for h in h_encontrados:
                        if h in dicionario_h:
                            dh = dicionario_h[h]
                            res_pgr.append({"GHE": nome_ghe, "Arquivo Origem": arq_nome, "Código": h, "Desc": dh["desc"], "Sev": texto_sev.get(dh["sev"], str(dh["sev"])), "Prob": str(prob), "Risco": matriz_oficial.get((dh["sev"], prob), "MODERADO"), "EPI": dh["epi"]})

        if not df_fis_bio.empty:
            for _, r in df_fis_bio.iterrows():
                if r["Agente"] in dicionario_campo:
                    dfc = dicionario_campo[r["Agente"]]
                    res_med.append({"GHE": r["GHE"], "Arquivo Origem": "Campo", "CAS": "-", "Agente": dfc["agente"], "NR15": dfc["nr15_lt"], "NR09": dfc["nr09_acao"], "NR07": dfc["nr07_ibe"], "Dec3048": dfc["dec_3048"], "eSocial": dfc["esocial_24"]})
                    res_pgr.append({"GHE": r["GHE"], "Arquivo Origem": "Campo", "Código": "-", "Desc": dfc["perigo"], "Sev": texto_sev.get(dfc["sev"], str(dfc["sev"])), "Prob": str(r["Probabilidade"]), "Risco": matriz_oficial.get((dfc["sev"], int(r["Probabilidade"])), "MODERADO"), "EPI": dfc["epi"]})

        my_bar.progress(100, text="Processamento finalizado!")
        
        if res_pgr or res_med:
            st.session_state['html_pgr'] = gerar_html_anexo(res_pgr, res_med)
            st.success("✅ Matriz gerada com sucesso!")
        else:
            st.warning("⚠️ Nenhum agente químico (CAS) ou perigo (Código H) foi identificado nos PDFs enviados.")

    if 'html_pgr' in st.session_state:
        st.markdown("### 📄 Resultado")
        col_name, col_btn = st.columns([3, 1])
        with col_name: projeto = st.text_input("Nome do Projeto para Salvar:")
        with col_btn:
            st.write("")
            if st.button("💾 Salvar no Banco", use_container_width=True) and projeto:
                conn = sqlite3.connect('seconci_banco_dados.db')
                c = conn.cursor()
                c.execute("INSERT INTO historico_laudos (nome_projeto, modulo, data_salvamento, html_relatorio) VALUES (?, ?, ?, ?)", (projeto, "PGR", datetime.now().strftime("%d/%m/%Y"), st.session_state['html_pgr']))
                conn.commit(); conn.close()
                st.toast("Projeto Salvo!", icon="✅")

        aba1, aba2 = st.tabs(["👁️ Pré-visualizar Documento", "⬇️ Exportar (.doc)"])
        with aba1: components.html(st.session_state['html_pgr'], height=600, scrolling=True)
        with aba2: st.download_button("Baixar Word", st.session_state['html_pgr'].encode('utf-8'), "PGR_Automatizado.doc", use_container_width=True)

# ==========================================
# TELA 3: CENTRAL DE ARQUIVOS
# ==========================================
elif "Central" in menu:
    st.title("📂 Central de Relatórios Salvos")
    conn = sqlite3.connect('seconci_banco_dados.db')
    
    cursor = conn.cursor()
    cursor.execute("SELECT count(name) FROM sqlite_master WHERE type='table' AND name='historico_laudos'")
    if cursor.fetchone()[0] == 1:
        df_hist = pd.read_sql_query("SELECT id, nome_projeto, modulo, data_salvamento FROM historico_laudos ORDER BY id DESC", conn)
        if not df_hist.empty:
            st.dataframe(df_hist, use_container_width=True, hide_index=True)
            selecao_id = st.selectbox("Selecione o ID do projeto para carregar:", df_hist['id'].tolist())
            if st.button("Carregar Documento", use_container_width=True):
                html_salvo = conn.cursor().execute("SELECT html_relatorio FROM historico_laudos WHERE id = ?", (selecao_id,)).fetchone()[0]
                st.components.v1.html(html_salvo, height=600, scrolling=True)
        else:
            st.info("Nenhum projeto salvo no banco de dados ainda.")
    else:
        st.info("Nenhum projeto salvo no banco de dados ainda.")
    conn.close()

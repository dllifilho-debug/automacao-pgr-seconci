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
# CONFIGURAÇÃO DA IA E BANCO DE DADOS
# ==========================================
CHAVE_API_GOOGLE = str(st.secrets["CHAVE_API_GOOGLE"]).strip().replace('"', '').replace("'", "")

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
# DICIONÁRIOS BASE (100% COMPLETOS)
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
    "108-88-3": {"agente": "Tolueno", "nr15_lt": "78 ppm ou 290 mg/m³", "nr09_acao": "39 ppm", "nr07_ibe": "o-Cresol ou Ácido Hipúrico", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "1330-20-7": {"agente": "Xileno", "nr15_lt": "78 ppm ou 340 mg/m³", "nr09_acao": "39 ppm", "nr07_ibe": "Ácidos Metilhipúricos", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "71-43-2": {"agente": "Benzeno", "nr15_lt": "VRT-MPT (Cancerígeno)", "nr09_acao": "Qualitativo", "nr07_ibe": "Ácido trans,trans-mucônico", "dec_3048": "25 anos", "esocial_24": "01.01.006"},
    "67-64-1": {"agente": "Acetona", "nr15_lt": "780 ppm ou 1870 mg/m³", "nr09_acao": "390 ppm", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "64-17-5": {"agente": "Etanol (Álcool Etílico)", "nr15_lt": "780 ppm ou 1480 mg/m³", "nr09_acao": "390 ppm", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "78-93-3": {"agente": "Metiletilcetona (MEK)", "nr15_lt": "155 ppm ou 460 mg/m³", "nr09_acao": "77.5 ppm", "nr07_ibe": "MEK na Urina", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "110-54-3": {"agente": "n-Hexano", "nr15_lt": "50 ppm ou 176 mg/m³", "nr09_acao": "25 ppm", "nr07_ibe": "2,5-Hexanodiona", "dec_3048": "25 anos", "esocial_24": "01.19.014"},
    "14808-60-7": {"agente": "Sílica Cristalina (Quartzo)", "nr15_lt": "Anexo 12", "nr09_acao": "50% do L.T.", "nr07_ibe": "Raio-X (OIT) e Espirometria", "dec_3048": "25 anos", "esocial_24": "01.18.001"},
    "1332-21-4": {"agente": "Asbesto / Amianto", "nr15_lt": "0,1 f/cm³", "nr09_acao": "0,05 f/cm³", "nr07_ibe": "Raio-X (OIT) e Espirometria", "dec_3048": "20 anos", "esocial_24": "01.02.001"},
    "7439-92-1": {"agente": "Chumbo (Fumos)", "nr15_lt": "0,1 mg/m³", "nr09_acao": "0,05 mg/m³", "nr07_ibe": "Chumbo no sangue e ALA-U", "dec_3048": "25 anos", "esocial_24": "01.08.001"},
    "65997-15-1": {"agente": "Cimento Portland", "nr15_lt": "10 mg/m³ (Poeira)", "nr09_acao": "5 mg/m³", "nr07_ibe": "Raio-X (OIT) e Espirometria", "dec_3048": "Não Enquadrado", "esocial_24": "01.18.001"},
    "1317-65-3": {"agente": "Carbonato de Cálcio", "nr15_lt": "10 mg/m³ (Poeira)", "nr09_acao": "5 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "1305-78-8": {"agente": "Óxido de Cálcio", "nr15_lt": "2 mg/m³", "nr09_acao": "1 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12168-85-3": {"agente": "Silicato Tricálcico", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Raio-X (OIT)", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "10034-77-2": {"agente": "Silicato Dicálcico", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Raio-X (OIT)", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12042-78-3": {"agente": "Aluminato de Cálcio", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Raio-X (OIT)", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "1309-48-4": {"agente": "Óxido de Magnésio", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Raio-X (OIT)", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "68334-30-5": {"agente": "Óleo Diesel", "nr15_lt": "Qualitativo", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Avaliar Hidrocarbonetos", "esocial_24": "01.01.026"},
    "112-80-1": {"agente": "Ácido Oleico", "nr15_lt": "N/A", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "52-51-7": {"agente": "Bronopol", "nr15_lt": "N/A", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"}
}

dicionario_campo = {
    "Físico: Ruído Contínuo/Intermitente": {"agente": "Ruído Contínuo ou Intermitente", "nr15_lt": "85 dB(A)", "nr09_acao": "80 dB(A)", "nr07_ibe": "Audiometria", "dec_3048": "25 anos", "esocial_24": "02.01.001", "perigo": "Exposição a níveis elevados de pressão sonora", "sev": 3, "epi": "Protetor Auditivo"},
    "Físico: Vibração de Mãos e Braços (VMB)": {"agente": "Vibração de Mãos e Braços (VMB)", "nr15_lt": "5,0 m/s²", "nr09_acao": "2,5 m/s²", "nr07_ibe": "Avaliação Clínica e Osteomuscular", "dec_3048": "25 anos", "esocial_24": "02.01.002", "perigo": "Transmissão de energia mecânica para o sistema mão-braço", "sev": 3, "epi": "Luvas antivibração / Revezamento"},
    "Físico: Vibração de Corpo Inteiro (VCI)": {"agente": "Vibração de Corpo Inteiro (VCI)", "nr15_lt": "1,1 m/s² ou 21,0 m/s¹.75", "nr09_acao": "0,5 m/s² ou 9,1 m/s¹.75", "nr07_ibe": "Avaliação Clínica e Osteomuscular", "dec_3048": "25 anos", "esocial_24": "02.01.003", "perigo": "Transmissão de energia mecânica para o corpo inteiro", "sev": 3, "epi": "Assentos com amortecimento / Revezamento"},
    "Biológico: Esgoto / Fossas": {"agente": "Microorganismos - Esgoto / Fossas", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos", "esocial_24": "03.01.005", "perigo": "Exposição a agentes biológicos infectocontagiosos", "sev": 4, "epi": "Luvas, Botas de PVC, Proteção facial"},
    "Biológico: Lixo Urbano": {"agente": "Microorganismos - Lixo Urbano", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos", "esocial_24": "03.01.007", "perigo": "Contato com resíduos e agentes biológicos", "sev": 4, "epi": "Luvas anticorte, Botas, Uniforme impermeável"},
    "Biológico: Estab. Saúde": {"agente": "Microorganismos - Área da Saúde", "nr15_lt": "Qualitativo (Anexo 14)", "nr09_acao": "Qualitativo", "nr07_ibe": "Exames Clínicos / Vacinas", "dec_3048": "25 anos", "esocial_24": "03.01.001", "perigo": "Exposição a patógenos em ambiente de saúde", "sev": 4, "epi": "Luvas de procedimento, Máscara, Avental"},
    "Ergonômico: Postura Inadequada": {"agente": "Fator Ergonômico - Postura", "nr15_lt": "N/A (NR-17)", "nr09_acao": "Avaliação AEP/AET", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente", "perigo": "Exigência de postura inadequada ou prolongada", "sev": 2, "epi": "Medidas Administrativas"},
    "Ergonômico: Levantamento/Transporte de Peso": {"agente": "Fator Ergonômico - Levantamento de Peso", "nr15_lt": "N/A (NR-17)", "nr09_acao": "Avaliação AEP/AET", "nr07_ibe": "Avaliação Clínica / Osteomuscular", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente", "perigo": "Esforço físico intenso e levantamento manual", "sev": 3, "epi": "Auxílio Mecânico"},
    "Acidente: Queda de Altura": {"agente": "Risco de Acidente - Altura", "nr15_lt": "N/A (NR-35)", "nr09_acao": "N/A", "nr07_ibe": "Protocolo Trabalho em Altura", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente", "perigo": "Trabalho executado acima de 2 metros", "sev": 4, "epi": "Cinturão de Segurança, Talabarte, Capacete"},
    "Acidente: Choque Elétrico": {"agente": "Risco de Acidente - Eletricidade", "nr15_lt": "N/A (NR-10)", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica / ECG", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente", "perigo": "Contato direto ou indireto com partes energizadas", "sev": 5, "epi": "Luvas Isolantes, Vestimenta ATPV, Capacete Classe B"},
    "Acidente: Máquinas e Equipamentos": {"agente": "Risco de Acidente - Partes Móveis", "nr15_lt": "N/A (NR-12)", "nr09_acao": "N/A", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente", "perigo": "Operação de máquinas com risco de corte/esmagamento", "sev": 4, "epi": "Luvas, Óculos, Botas de Segurança"}
}

matriz_risco_exame = {
    "TOLUENO": {"exame": "Ortocresol na Urina", "periodico": "6 MESES"},
    "RUÍDO": {"exame": "Audiometria", "periodico": "12 MESES"},
    "SÍLICA": {"exame": "Raio-X de Tórax (OIT) + Espirometria", "periodico": "12 a 24 MESES"},
    "VIBRAÇÃO": {"exame": "Avaliação Clínica e Osteomuscular", "periodico": "12 MESES"},
    "POEIRA": {"exame": "Raio-X de Tórax (OIT)", "periodico": "12 MESES"},
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
    ]
}

# ==========================================
# MOTOR IA - SISTEMA DE CASCATA BLINDADO
# ==========================================
def chamar_api_gemini_cascata(prompt, pdf_b64=None):
    parts = [{"text": prompt}]
    if pdf_b64:
        parts.append({"inlineData": {"mimeType": "application/pdf", "data": pdf_b64}})
        
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {"temperature": 0.0},
        "safetySettings": [
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_ONLY_HIGH"}
        ]
    }
    
    # Array de servidores robusto para nunca falhar
    modelos_para_testar = [
        "gemini-1.5-pro-latest",
        "gemini-1.5-flash-latest",
        "gemini-1.5-pro",
        "gemini-1.5-flash",
        "gemini-pro" 
    ]
    
    for modelo in modelos_para_testar:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent?key={CHAVE_API_GOOGLE}"
        try:
            resp = requests.post(url, headers={'Content-Type': 'application/json'}, json=payload, timeout=25)
            if resp.status_code == 200:
                return resp.json().get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
            elif resp.status_code == 400 or resp.status_code == 404:
                continue 
        except:
            continue

    st.error("🚨 API do Google inoperante. Usando banco local.")
    return None

def extrair_json_seguro(texto_ia):
    try:
        texto_limpo = texto_ia.replace('```json', '').replace('```', '').strip()
        inicio = texto_limpo.find('{')
        if inicio != -1 and texto_limpo.startswith('['):
            inicio = texto_limpo.find('[') # Tratamento para listas JSON
        if inicio != -1: 
            texto_limpo = texto_limpo[inicio:]
        return json.loads(texto_limpo)
    except json.JSONDecodeError:
        return {}

def buscar_dados_completos_ia(cas_faltantes):
    if not cas_faltantes: return {}
    lista_cas_str = ", ".join([str(c).strip() for c in cas_faltantes])
    
    prompt = f"""
    Você é um Higienista Ocupacional Brasileiro. Para os CAS: {lista_cas_str}, forneça os limites.
    Retorne APENAS um JSON válido. As chaves devem ser os números CAS. Exemplo de estrutura:
    {{
        "1317-65-3": {{
            "agente": "Nome Químico em Português", 
            "nr15_lt": "Ex: 10 mg/m³", 
            "nr09_acao": "Ex: 5 mg/m³",
            "nr07_ibe": "Ex: Avaliação Clínica", 
            "dec_3048": "Ex: Não Enquadrado", 
            "esocial_24": "Ex: 09.01.001"
        }}
    }}
    """
    
    texto_ia = chamar_api_gemini_cascata(prompt)
    if texto_ia: return extrair_json_seguro(texto_ia)
    return {}

def processar_pcmso(dados_pgr_json):
    tabela_pcmso = []
    
    if isinstance(dados_pgr_json, dict):
        dados_pgr_json = [dados_pgr_json]
        
    for ghe in dados_pgr_json:
        if not isinstance(ghe, dict): continue
        nome_ghe = ghe.get("ghe", "Sem GHE")
        cargos = ghe.get("cargos", [])
        riscos = ghe.get("riscos_mapeados", [])
        
        for cargo in cargos:
            exames_cargo = [{"exame": "Exame Clínico (Anamnese/Físico)", "periodicidade": "12 MESES", "motivo": "NR-07 Básico"}]
            cargo_upper = str(cargo).upper()
            
            for funcao, exames in matriz_funcao_exame.items():
                if funcao in cargo_upper: exames_cargo.extend(exames)
            
            for risco in riscos:
                if not isinstance(risco, dict): continue
                txt_risco = (str(risco.get("nome_agente", "")) + " " + str(risco.get("perigo_especifico", ""))).upper()
                for agente, regra in matriz_risco_exame.items():
                    if agente in txt_risco:
                        exames_cargo.append({"exame": regra["exame"], "periodicidade": regra["periodico"], "motivo": f"Exposição a {agente}"})
                if "ALTURA" in txt_risco: exames_cargo.extend(matriz_funcao_exame["TRABALHO EM ALTURA"])
            
            unicos = {v['exame']:v for v in exames_cargo}.values()
            
            for ex in unicos:
                tabela_pcmso.append({
                    "GHE / Setor": nome_ghe, 
                    "Cargo": cargo, 
                    "Exame Clínico/Complementar": ex["exame"], 
                    "Periodicidade": ex["periodicidade"], 
                    "Justificativa Legal / Risco": ex.get("motivo", "Protocolo Função")
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
            html_content += "<h4>Diretrizes Médicas e Previdenciárias (Automação IA)</h4><table class='funcao-mini-table'><thead><tr><th>Cód / CAS</th><th>Agente</th><th>Lim. Tol. (NR-15)</th><th>Nível Ação (NR-09)</th><th>Exame/IBE (NR-07)</th><th>Dec 3048</th><th>eSocial</th></tr></thead><tbody>"
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
    <div class='header'>MATRIZ DE EXAMES - PCMSO</div>
    <table><tr><th>GHE / Setor</th><th>Cargo</th><th>Exame Clínico / Complementar</th><th>Periodicidade</th><th>Justificativa / Agente</th></tr>
    """
    for _, row in df_pcmso.iterrows():
        html += f"<tr><td><strong>{row['GHE / Setor']}</strong></td><td>{row['Cargo']}</td><td>{row['Exame Clínico/Complementar']}</td><td>{row['Periodicidade']}</td><td>{row['Justificativa Legal / Risco']}</td></tr>"
    html += "</table></body></html>"
    return html

# ==========================================
# BARRA LATERAL (SIDEBAR)
# ==========================================
if os.path.exists("logo.png"): st.sidebar.image("logo.png", width="stretch")
elif os.path.exists("logo.jpg"): st.sidebar.image("logo.jpg", width="stretch")
else: st.sidebar.markdown("<h2 style='text-align: center; color: #084D22;'>SECONCI-GO</h2>", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.title("🧩 Módulos do Sistema")
modulo_selecionado = st.sidebar.radio(
    "Selecione a funcionalidade:", 
    ["1️⃣ Engenharia: FISPQ / FDS ➡️ PGR", "2️⃣ Medicina: PGR ➡️ PCMSO"],
    key="menu_lateral" # CHAVE ÚNICA PARA EVITAR O DUPLICATE ERROR
)

st.sidebar.markdown("---")
st.sidebar.title("📂 Histórico de Laudos")
conn = sqlite3.connect('seconci_banco_dados.db')
df_historico = pd.read_sql_query("SELECT id, nome_projeto, data_salvamento FROM historico_laudos ORDER BY id DESC", conn)
conn.close()

historico_selecionado = None
if not df_historico.empty:
    opcoes_historico = ["Selecione um projeto salvo..."] + [f"{row['id']} - {row['nome_projeto']} ({row['data_salvamento']})" for _, row in df_historico.iterrows()]
    selecao = st.sidebar.selectbox("Carregar projeto antigo:", opcoes_historico, key="select_historico")
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
# INTERFACE PRINCIPAL E LÓGICA
# ==========================================
st.title("Sistema Integrado SST - Seconci GO 🚀")

if historico_selecionado:
    st.markdown("### 🗄️ Visualizando Relatório")
    aba_preview_hist, aba_download_hist = st.tabs(["👁️ Pré-visualizar", "📄 Baixar em Word (.doc)"])
    with aba_preview_hist: components.html(historico_selecionado, height=700, scrolling=True)
    with aba_download_hist: st.download_button("Baixar Relatório", data=historico_selecionado.encode('utf-8'), file_name="Relatorio_Historico.doc", mime="application/msword", key="btn_download_hist")

elif "1️⃣" in modulo_selecionado:
    st.header("Módulo de Engenharia: Extrator de FISPQs")
    st.info("A API está configurada em cascata para evitar erros. O banco local está com capacidade máxima ativada.")
    
    arquivos_fispq = st.file_uploader("Insira as FISPQs em PDF", type=["pdf"], accept_multiple_files=True, key="uploader_eng")
    textos_pdfs = {}
    df_editado = pd.DataFrame()
    ghe_opcoes = ["Nenhum GHE definido"]

    if arquivos_fispq:
        with st.spinner("Lendo PDFs..."):
            for arquivo in arquivos_fispq:
                with pdfplumber.open(arquivo) as pdf:
                    textos_pdfs[arquivo.name] = "\n".join([p.extract_text() or "" for p in pdf.pages])

        st.markdown("### 2️⃣ GHEs Químicos")
        nomes_arquivos = [arq.name for arq in arquivos_fispq]
        dados_iniciais = [{"GHE": "GHE 01 - Função", "Arquivo FISPQ/FDS": nome, "Probabilidade": 3} for nome in nomes_arquivos]
        df_editado = st.data_editor(
            pd.DataFrame(dados_iniciais), num_rows="dynamic",
            column_config={
                "GHE": st.column_config.TextColumn("GHE", required=True),
                "Arquivo FISPQ/FDS": st.column_config.SelectboxColumn("Arquivo", options=nomes_arquivos, required=True),
                "Probabilidade": st.column_config.NumberColumn("Prob.", min_value=1, max_value=5, required=True)
            }, width="stretch", key="editor_quimico"
        )
        ghe_opcoes = df_editado["GHE"].unique().tolist() if not df_editado.empty else ["Nenhum GHE definido"]

    st.markdown("### 3️⃣ Riscos Físicos, Biológicos e de Acidentes")
    agentes_opcoes = list(dicionario_campo.keys())
    df_fis_bio_editado = st.data_editor(
        pd.DataFrame([{"GHE": ghe_opcoes[0], "Agente": agentes_opcoes[0], "Probabilidade": 3}]), num_rows="dynamic",
        column_config={
            "GHE": st.column_config.SelectboxColumn("GHE", options=ghe_opcoes, required=True),
            "Agente": st.column_config.SelectboxColumn("Agente", options=agentes_opcoes, required=True),
            "Probabilidade": st.column_config.NumberColumn("Prob.", min_value=1, max_value=5, required=True)
        }, width="stretch", key="editor_fisico"
    )

    if st.button("🪄 Processar Relatório e Buscar Legislação", width="stretch", type="primary", key="btn_processar_eng"):
        with st.spinner("Processando dados e verificando banco normativo/IA..."):
            resultados_pgr, resultados_medicos = [], []
            
            if not df_editado.empty:
                for index, row in df_editado.iterrows():
                    nome_ghe, nome_arq, v_prob = row["GHE"], row["Arquivo FISPQ/FDS"], int(row["Probabilidade"])
                    if nome_arq in textos_pdfs:
                        texto_completo = textos_pdfs[nome_arq]
                        cas_encontrados_linha = list(set(re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', texto_completo)))
                        cas_desconhecidos = [c for c in cas_encontrados_linha if c not in dicionario_cas]
                        
                        dados_dinamicos_ia = {}
                        if cas_desconhecidos:
                            st.toast(f"Analisando CAS não mapeados na IA: {', '.join(cas_desconhecidos)}", icon="🧠")
                            dados_dinamicos_ia = buscar_dados_completos_ia(cas_desconhecidos)
                        
                        for cas in cas_encontrados_linha:
                            cas_limpo = cas.strip()
                            if cas_limpo in dicionario_cas:
                                dados_med = dicionario_cas[cas_limpo]
                            else:
                                ia_info = dados_dinamicos_ia.get(cas_limpo, {})
                                dados_med = {
                                    "agente": ia_info.get("agente", f"Produto Desconhecido ({cas_limpo})"),
                                    "nr15_lt": ia_info.get("nr15_lt", "Consultar NR-15"),
                                    "nr09_acao": ia_info.get("nr09_acao", "Consultar NR-09"),
                                    "nr07_ibe": ia_info.get("nr07_ibe", "Avaliação Clínica"),
                                    "dec_3048": ia_info.get("dec_3048", "Não Enquadrado"),
                                    "esocial_24": ia_info.get("esocial_24", "09.01.001")
                                }
                                
                            resultados_medicos.append({
                                "GHE": nome_ghe, "Arquivo Origem": nome_arq, "Nº CAS": cas,
                                "Agente Químico": dados_med["agente"], "Lim. Tolerância (NR-15)": dados_med["nr15_lt"],
                                "Nível de Ação (NR-09)": dados_med["nr09_acao"], "IBE (NR-07)": dados_med["nr07_ibe"],
                                "Dec 3048": dados_med["dec_3048"], "eSocial": dados_med["esocial_24"]
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
                            "GHE": nome_ghe, "Arquivo Origem": "Campo", "Nº CAS": "-",
                            "Agente Químico": dados_fis["agente"], "Lim. Tolerância (NR-15)": dados_fis["nr15_lt"],
                            "Nível de Ação (NR-09)": dados_fis["nr09_acao"], "IBE (NR-07)": dados_fis["nr07_ibe"],
                            "Dec 3048": dados_fis["dec_3048"], "eSocial": dados_fis["esocial_24"]
                        })
                        nivel_risco = matriz_oficial.get((dados_fis["sev"], v_prob), "N/A")
                        resultados_pgr.append({
                            "GHE": nome_ghe, "Arquivo Origem": "Campo", "Código GHS": "-",
                            "Perigo Identificado": dados_fis["perigo"], "Severidade": texto_sev.get(dados_fis["sev"], str(dados_fis["sev"])),
                            "Probabilidade": str(v_prob), "NÍVEL DE RISCO": nivel_risco,
                            "Ação Requerida": acoes_requeridas.get(nivel_risco, "Manual"), "EPI (NR-06)": dados_fis["epi"]
                        })

            if resultados_pgr or resultados_medicos:
                html_final = gerar_html_anexo(resultados_pgr, resultados_medicos)
                st.session_state['ultimo_html_eng'] = html_final
                st.success("✅ Relatório Técnico e Enquadramentos Completados com Sucesso.")

    if 'ultimo_html_eng' in st.session_state:
        col1, col2 = st.columns([3, 1])
        with col1: nome_projeto_eng = st.text_input("Nome da Empresa / Projeto:", key="input_nome_proj")
        with col2:
            st.write(""); st.write("")
            if st.button("Gravar Banco", key="btn_salvar_eng", width="stretch") and nome_projeto_eng:
                conn = sqlite3.connect('seconci_banco_dados.db')
                c = conn.cursor()
                c.execute("INSERT INTO historico_laudos (nome_projeto, data_salvamento, html_relatorio) VALUES (?, ?, ?)", 
                          (nome_projeto_eng + " (PGR)", datetime.now().strftime("%d/%m/%Y %H:%M"), st.session_state['ultimo_html_eng']))
                conn.commit()
                conn.close()
                st.success("Salvo com sucesso!")

        aba_preview_eng, aba_download_eng = st.tabs(["👁️ Pré-visualizar", "📄 Baixar (.doc)"])
        with aba_preview_eng: components.html(st.session_state['ultimo_html_eng'], height=500, scrolling=True)
        with aba_download_eng: st.download_button("Baixar Word", st.session_state['ultimo_html_eng'].encode('utf-8'), "PGR_Fase1.doc", key="btn_down_word_eng")

elif "2️⃣" in modulo_selecionado:
    st.header("🩺 Módulo Médico: Importador de PGR e Gerador de PCMSO")
    
    with st.container():
        arquivo_pgr = st.file_uploader("Arraste o PDF do PGR aqui", type=["pdf"], key="uploader_med")
        if arquivo_pgr:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚀 Extrair Riscos e Gerar PCMSO", type="primary", use_container_width=True, key="btn_processar_med"):
                with st.spinner("Analisando matrizes e cruzando protocolos NR-07..."):
                    pdf_bytes = arquivo_pgr.getvalue()
                    pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
                    
                    prompt_extracao = """
                    Analise o texto deste documento de PGR/Inventário de Riscos.
                    Retorne EXATAMENTE UM JSON válido (uma lista de objetos) contendo a relação de GHE, cargos e riscos.
                    Sem introdução, sem marcadores markdown.
                    
                    Modelo esperado:
                    [
                      {
                        "ghe": "Nome do Setor ou GHE",
                        "cargos": ["Nome do Cargo"],
                        "riscos_mapeados": [
                          {"nome_agente": "Ex: Risco Biológico", "perigo_especifico": "Vírus"}
                        ]
                      }
                    ]
                    """
                    
                    texto_ia = chamar_api_gemini_cascata(prompt_extracao, pdf_b64)
                    
                    if texto_ia:
                        try:
                            json_pgr = extrair_json_seguro(texto_ia)
                            df_pcmso_gerado = processar_pcmso(json_pgr)
                            st.session_state['ultimo_html_med'] = gerar_html_pcmso(df_pcmso_gerado)
                            st.session_state['df_pcmso_gerado'] = df_pcmso_gerado
                            st.success("✅ Matriz Cruzada e Processada!")
                        except Exception as e:
                            st.error(f"Erro ao montar a tabela médica. Detalhe: {e}")

        if 'ultimo_html_med' in st.session_state:
            aba_dados, aba_preview_med, aba_download_med = st.tabs(["📊 Dados", "👁️ Visão", "📄 Word"])
            with aba_dados: 
                if 'df_pcmso_gerado' in st.session_state and not st.session_state['df_pcmso_gerado'].empty:
                    st.dataframe(st.session_state['df_pcmso_gerado'], use_container_width=True)
            with aba_preview_med: components.html(st.session_state['ultimo_html_med'], height=600, scrolling=True)
            with aba_download_med: st.download_button("Baixar PCMSO", st.session_state['ultimo_html_med'].encode('utf-8'), "PCMSO_Seconci.doc", "application/msword", key="btn_down_word_med")

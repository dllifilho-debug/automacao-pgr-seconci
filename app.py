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
# BARRA LATERAL
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
# DICIONÁRIOS BASE (ENGENHARIA E QUÍMICA)
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
    "TRIVIAL": "Manter controles existentes.",
    "TOLERÁVEL": "Manter controles. Considerar melhorias.",
    "MODERADO": "Implantar controles. EPI e monitoramento.",
    "SUBSTANCIAL": "Trabalho não deve iniciar sem redução do risco.",
    "INTOLERÁVEL": "TRABALHO PROIBIDO. Risco grave e iminente."
}

texto_sev = {1: "1-LEVE", 2: "2-BAIXA", 3: "3-MODERADA", 4: "4-ALTA", 5: "5-EXTREMA"}

dicionario_cas = {
    "108-88-3": {"agente": "Tolueno", "nr15_lt": "78 ppm ou 290 mg/m³", "nr09_acao": "39 ppm", "nr07_ibe": "o-Cresol / Ácido Hipúrico", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "1330-20-7": {"agente": "Xileno", "nr15_lt": "78 ppm ou 340 mg/m³", "nr09_acao": "39 ppm", "nr07_ibe": "Ácidos Metilhipúricos", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "71-43-2": {"agente": "Benzeno", "nr15_lt": "VRT-MPT (Cancerígeno)", "nr09_acao": "Qualitativo", "nr07_ibe": "Ácido trans,trans-mucônico", "dec_3048": "25 anos", "esocial_24": "01.01.006"},
    "14808-60-7": {"agente": "Sílica Cristalina (Quartzo)", "nr15_lt": "Avaliar Anexo 12", "nr09_acao": "50% do L.T.", "nr07_ibe": "Raio-X (OIT) e Espirometria", "dec_3048": "25 anos", "esocial_24": "01.18.001"},
    "65997-15-1": {"agente": "Cimento Portland", "nr15_lt": "10 mg/m³ (Poeira)", "nr09_acao": "5 mg/m³", "nr07_ibe": "Raio-X (OIT) e Espirometria", "dec_3048": "Não Enquadrado", "esocial_24": "01.18.001"},
    "1317-65-3": {"agente": "Carbonato de Cálcio", "nr15_lt": "10 mg/m³ (Poeira)", "nr09_acao": "5 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "1305-78-8": {"agente": "Óxido de Cálcio", "nr15_lt": "2 mg/m³", "nr09_acao": "1 mg/m³", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12042-78-3": {"agente": "Aluminato de Cálcio", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Raio-X (OIT)", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"}
}

dicionario_campo = {
    "Físico: Ruído Contínuo/Intermitente": {"agente": "Ruído Contínuo/Intermitente", "nr15_lt": "85 dB(A)", "nr09_acao": "80 dB(A)", "nr07_ibe": "Audiometria", "dec_3048": "25 anos", "esocial_24": "02.01.001", "perigo": "Pressão sonora elevada", "sev": 3, "epi": "Protetor Auditivo"},
    "Físico: Vibração (VMB)": {"agente": "Vibração (VMB)", "nr15_lt": "5,0 m/s²", "nr09_acao": "2,5 m/s²", "nr07_ibe": "Avaliação Clínica", "dec_3048": "25 anos", "esocial_24": "02.01.002", "perigo": "Energia mecânica", "sev": 3, "epi": "Luvas antivibração"},
    "Biológico: Esgoto / Fossas": {"agente": "Microorganismos - Esgoto", "nr15_lt": "Qualitativo", "nr09_acao": "Qualitativo", "nr07_ibe": "Clínico/Vacinas", "dec_3048": "25 anos", "esocial_24": "03.01.005", "perigo": "Agentes biológicos", "sev": 4, "epi": "Luvas/Botas PVC"},
    "Acidente: Queda de Altura": {"agente": "Risco de Acidente - Altura", "nr15_lt": "NR-35", "nr09_acao": "N/A", "nr07_ibe": "Protocolo Altura", "dec_3048": "Não Enquadrado", "esocial_24": "Ausente (PGR)", "perigo": "Trabalho > 2m", "sev": 4, "epi": "Cinturão/Talabarte"}
}

matriz_risco_exame = {
    "TOLUENO": {"exame": "Ortocresol na Urina", "periodico": "6 MESES"},
    "RUÍDO": {"exame": "Audiometria", "periodico": "12 MESES"},
    "SÍLICA": {"exame": "Raio-X de Tórax (OIT) + Espirometria", "periodico": "12 a 24 MESES"},
    "POEIRA": {"exame": "Raio-X de Tórax (OIT)", "periodico": "12 MESES"}
}

matriz_funcao_exame = {
    "ALTURA": [
        {"exame": "Glicemia de Jejum", "periodicidade": "12 MESES"},
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
# MOTOR IA CORE (BLINDAGEM CONTRA 400 E 404)
# ==========================================
def chamar_api_gemini(prompt, pdf_b64=None):
    parts = [{"text": prompt}]
    if pdf_b64:
        parts.append({"inlineData": {"mimeType": "application/pdf", "data": pdf_b64}})
        
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {"temperature": 0.1},
        # BLOCK_ONLY_HIGH evita erro 400 (INVALID_ARGUMENT) em chaves standard
        "safetySettings": [
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_ONLY_HIGH"}
        ]
    }
    
    headers = {'Content-Type': 'application/json'}
    
    try:
        # Tentativa 1: Modelo PRO (Mais adequado para leitura estrutural de tabelas)
        url_pro = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent?key={CHAVE_API_GOOGLE}"
        resp = requests.post(url_pro, headers=headers, json=payload, timeout=50)
        
        if resp.status_code == 200:
            return resp.json().get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
            
        elif resp.status_code == 404:
            # Tentativa 2 (Fallback Invisível): Se o Pro estiver offline na região da chave, puxa o Flash
            url_flash = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key={CHAVE_API_GOOGLE}"
            resp_fb = requests.post(url_flash, headers=headers, json=payload, timeout=50)
            if resp_fb.status_code == 200:
                return resp_fb.json().get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
            else:
                st.error(f"🚨 ERRO HTTP DA API DO GOOGLE (Fallback): {resp_fb.status_code} - {resp_fb.text}")
                return None
        else:
            st.error(f"🚨 ERRO HTTP DA API DO GOOGLE: {resp.status_code} - {resp.text}")
            return None
            
    except Exception as e:
        st.error(f"🚨 FALHA DE CONEXÃO ESTRUTURAL COM O GOOGLE: {e}")
        return None

def buscar_dados_completos_ia(cas_faltantes):
    if not cas_faltantes: return {}
    
    cas_limpos = [str(c).strip() for c in cas_faltantes]
    lista_cas_str = ", ".join(cas_limpos)
    
    prompt = f"""
    Você é um Engenheiro de Segurança do Trabalho e Especialista em Higiene Ocupacional (NR-15, NR-09, NR-07 e eSocial).
    Para os seguintes números CAS: {lista_cas_str}, forneça os parâmetros legais vigentes.
    
    CRÍTICO: Retorne APENAS um objeto JSON válido. Sem introdução, sem marcadores markdown (como ```json).
    As chaves do JSON DEVEM SER exatamente os números CAS solicitados.
    
    Exemplo de estruturação obrigatória:
    {{
        "{cas_limpos[0] if cas_limpos else '1317-65-3'}": {{
            "agente": "Nome Químico Oficial em Português",
            "nr15_lt": "Ex: 10 mg/m³, 50 ppm ou Avaliar Anexo 12",
            "nr09_acao": "Ex: 5 mg/m³ ou N/A",
            "nr07_ibe": "Ex: Raio-X OIT ou Avaliação Clínica",
            "dec_3048": "Ex: 25 anos (Linha 1.0.19) ou Não Enquadrado",
            "esocial_24": "Ex: 01.18.001 ou 09.01.001"
        }}
    }}
    """
    
    texto_ia = chamar_api_gemini(prompt)
    if texto_ia:
        try:
            # Remoção bruta de markdown residual para evitar JSONDecodeError
            texto_limpo = texto_ia.replace('```json', '').replace('```', '').strip()
            return json.loads(texto_limpo)
        except json.JSONDecodeError as e:
            st.error(f"🚨 A IA retornou texto fora do padrão JSON. Erro de Parsing: {e}")
            return {}
    return {}

def processar_pcmso(dados_pgr_json):
    tabela_pcmso = []
    for ghe in dados_pgr_json:
        nome_ghe = ghe.get("ghe", "Sem GHE")
        cargos = ghe.get("cargos", [])
        riscos = ghe.get("riscos_mapeados", [])
        
        for cargo in cargos:
            exames_do_cargo = [{"exame": "Exame Clínico (Anamnese/Físico)", "periodicidade": "12 MESES", "motivo": "NR-07 Básico"}]
            cargo_upper = cargo.upper()
            
            for funcao_chave, exames in matriz_funcao_exame.items():
                if funcao_chave in cargo_upper: exames_do_cargo.extend(exames)
            
            for risco in riscos:
                agente = risco.get("nome_agente", "").upper()
                perigo = risco.get("perigo_especifico", "").upper()
                texto_risco_completo = agente + " " + perigo
                
                for agente_chave, regra in matriz_risco_exame.items():
                    if agente_chave in texto_risco_completo:
                        exames_do_cargo.append({
                            "exame": regra["exame"], "periodicidade": regra["periodico"], "motivo": f"Exposição a {agente_chave}"
                        })
                
                if "ALTURA" in texto_risco_completo: exames_do_cargo.extend(matriz_funcao_exame["ALTURA"])
            
            exames_unicos = {v['exame']:v for v in exames_do_cargo}.values()
            
            for ex in exames_unicos:
                tabela_pcmso.append({
                    "GHE / Setor": nome_ghe, "Cargo": cargo, "Exame Clínico/Complementar": ex["exame"],
                    "Periodicidade": ex["periodicidade"], "Justificativa Legal / Risco": ex.get("motivo", f"Protocolo Função")
                })
    return pd.DataFrame(tabela_pcmso)

# ==========================================
# GERADORES HTML WORD
# ==========================================
def gerar_html_anexo(resultados_pgr, resultados_medicos):
    html_content = """<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="[http://www.w3.org/TR/REC-html40](http://www.w3.org/TR/REC-html40)">
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
    html = """<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="[http://www.w3.org/TR/REC-html40](http://www.w3.org/TR/REC-html40)">
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
# CORPO PRINCIPAL
# ==========================================
st.title("Sistema Integrado SST - Seconci GO 🚀")

if historico_selecionado:
    st.markdown("### 🗄️ Visualizando Relatório")
    aba_preview, aba_download = st.tabs(["👁️ Pré-visualizar", "📄 Baixar em Word (.doc)"])
    with aba_preview: components.html(historico_selecionado, height=700, scrolling=True)
    with aba_download: st.download_button("Baixar Relatório", data=historico_selecionado.encode('utf-8'), file_name="Relatorio_Historico.doc", mime="application/msword")

# ==========================================
# MÓDULO 1: ENGENHARIA
# ==========================================
elif "1️⃣" in modulo_selecionado:
    st.header("Módulo de Engenharia: Extrator de FISPQs (Automação Especialista)")
    st.info("O sistema agora atuará como Especialista em Higiene Ocupacional, buscando dinamicamente os Limites de Tolerância e Códigos eSocial.")
    
    arquivos_fispq = st.file_uploader("Insira as FISPQs em PDF", type=["pdf"], accept_multiple_files=True)
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
            }, width="stretch"
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
        }, width="stretch"
    )

    if st.button("🪄 Processar Relatório e Buscar Legislação", width="stretch", type="primary"):
        with st.spinner("Estruturando enquadramentos de eSocial/NRs com Inteligência Artificial..."):
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
                            st.toast(f"Analisando matriz legal (NR-15/eSocial) para os CAS: {', '.join(cas_desconhecidos)}", icon="🧠")
                            dados_dinamicos_ia = buscar_dados_completos_ia(cas_desconhecidos)
                        
                        for cas in cas_encontrados_linha:
                            cas_limpo = cas.strip()
                            if cas_limpo in dicionario_cas:
                                dados_med = dicionario_cas[cas_limpo]
                            else:
                                ia_info = dados_dinamicos_ia.get(cas_limpo, {})
                                dados_med = {
                                    "agente": ia_info.get("agente", "Produto Não Localizado/Erro API"),
                                    "nr15_lt": ia_info.get("nr15_lt", "Sem Limite Estabelecido"),
                                    "nr09_acao": ia_info.get("nr09_acao", "N/A"),
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
        with col1: nome_projeto_eng = st.text_input("Nome da Empresa / Projeto:")
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

        aba_preview, aba_download = st.tabs(["👁️ Pré-visualizar", "📄 Baixar (.doc)"])
        with aba_preview: components.html(st.session_state['ultimo_html_eng'], height=500, scrolling=True)
        with aba_download: st.download_button("Baixar Word", st.session_state['ultimo_html_eng'].encode('utf-8'), "PGR_Fase1.doc")

# ==========================================
# MÓDULO 2: MEDICINA (PGR -> PCMSO)
# ==========================================
elif "2️⃣" in modulo_selecionado:
    st.header("🩺 Módulo Médico: Importador de PGR e Gerador de PCMSO")
    
    with st.container():
        arquivo_pgr = st.file_uploader("Arraste o PDF do PGR aqui", type=["pdf"])
        if arquivo_pgr:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚀 Extrair Riscos e Gerar PCMSO", type="primary", use_container_width=True):
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
                    
                    texto_ia = chamar_api_gemini(prompt_extracao, pdf_b64)
                    
                    if texto_ia:
                        try:
                            # Limpeza de markdown garantida
                            texto_limpo = texto_ia.replace('```json', '').replace('```', '').strip()
                            json_pgr = json.loads(texto_limpo)
                            df_pcmso_gerado = processar_pcmso(json_pgr)
                            st.session_state['ultimo_html_med'] = gerar_html_pcmso(df_pcmso_gerado)
                            st.session_state['df_pcmso_gerado'] = df_pcmso_gerado
                            st.success("✅ Matriz Cruzada e Processada!")
                        except json.JSONDecodeError as e:
                            st.error(f"Erro ao montar a tabela médica. Resposta da IA fora de padrão. Detalhe: {e}")

        if 'ultimo_html_med' in st.session_state:
            aba_dados, aba_preview, aba_download = st.tabs(["📊 Dados", "👁️ Visão", "📄 Word"])
            with aba_dados: st.dataframe(st.session_state['df_pcmso_gerado'], use_container_width=True)
            with aba_preview: components.html(st.session_state['ultimo_html_med'], height=600, scrolling=True)
            with aba_download: st.download_button("Baixar PCMSO", st.session_state['ultimo_html_med'].encode('utf-8'), "PCMSO_Seconci.doc", "application/msword")

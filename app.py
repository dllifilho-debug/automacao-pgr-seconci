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

# Configuração da página
st.set_page_config(page_title="Automação PGR - Seconci GO", layout="wide")

# ==========================================
# CONFIGURAÇÃO DA IA (CONEXÃO DIRETA REST)
# ==========================================
CHAVE_API_GOOGLE = st.secrets["CHAVE_GOOGLE"]

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
if os.path.exists("logo.png"):
    st.sidebar.image("logo.png", width="stretch")
elif os.path.exists("logo.jpg"):
    st.sidebar.image("logo.jpg", width="stretch")
else:
    st.sidebar.markdown("<h2 style='text-align: center; color: #084D22;'>SECONCI-GO</h2>", unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.title("📂 Histórico de Laudos")
st.sidebar.info("Acesse relatórios salvos no banco de dados para consulta ou envio ao eSocial.")

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
        st.sidebar.success("✅ Projeto carregado na tela principal.")
else:
    st.sidebar.write("Nenhum projeto salvo ainda.")

# ==========================================
# CORPO PRINCIPAL DO SISTEMA
# ==========================================
st.title("Sistema de Automação de PGR, FISPQ e eSocial 🚀")
st.info("Versão do Motor: 26.0 (Cérebro Google Gemini - Auto-Discovery Engine & Trava de Alucinação)")

# ==========================================
# MEGA BANCO DE DADOS LEGAL E DICIONÁRIOS
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

dicionario_fis_bio = {
    "Ruído Contínuo/Intermitente": {
        "agente": "Ruído Contínuo ou Intermitente", "nr15_lt": "85 dB(A)", "nr09_acao": "80 dB(A)", "nr07_ibe": "Audiometria", "dec_3048": "25 anos (Linha 2.0.1)", "esocial_24": "02.01.001",
        "perigo": "Exposição a níveis elevados de pressão sonora", "sev": 3, "epi": "Protetor Auditivo (Atenuação adequada)"
    },
    "Vibração de Mãos e Braços (VMB)": {
        "agente": "Vibração de Mãos e Braços (VMB)", "nr15_lt": "5,0 m/s²", "nr09_acao": "2,5 m/s²", "nr07_ibe": "Avaliação Clínica", "dec_3048": "25 anos (Linha 2.0.2)", "esocial_24": "02.01.002",
        "perigo": "Transmissão de energia mecânica para o sistema mão-braço", "sev": 3, "epi": "Luvas antivibração / Revezamento"
    },
    "Vibração de Corpo Inteiro (VCI)": {
        "agente": "Vibração de Corpo Inteiro (VCI)", "nr15_lt": "1,1 m/s² ou 21,0 m/s¹.75", "nr09_acao": "0,5 m/s² ou 9,1 m/s¹.75", "nr07_ibe": "Avaliação Clínica", "dec_3048": "25 anos (Linha 2.0.2)", "esocial_24": "02.01.003",
        "perigo": "Transmissão de energia mecânica para o corpo inteiro", "sev": 3, "epi": "Assentos com amortecimento / Revezamento"
    },
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
    }
}

# ==========================================
# GERADOR DE HTML (CONSOLIDADO - CORES SECONCI)
# ==========================================
def gerar_html_anexo(resultados_pgr, resultados_medicos):
    css_base = """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
      :root {
        --green-900: #042B13; 
        --green-800: #084D22; 
        --green-700: #0F823B; 
        --green-600: #1AA04B; 
        --green-500: #28C45D; 
        --green-200: #C7F092; 
        --green-100: #D9F585; 
        --white: #FFFFFF; --gray-50: #F8F9FA;
        --gray-200: #E9ECEF; --gray-600: #6C757D; --gray-900: #212529;
      }
      body { font-family: 'Inter', sans-serif; font-size: 10pt; color: var(--gray-900); background: var(--white); padding: 20px; }
      .anexo-header { background: var(--green-800); color: var(--white); padding: 14px 20px; font-size: 13pt; font-weight: 700; margin-bottom: 20px; border-radius: 2px; text-align: center; text-transform: uppercase; border-bottom: 4px solid var(--green-100); }
      .funcao-card { border: 1px solid var(--green-800); border-radius: 4px; margin-bottom: 20px; page-break-inside: avoid; }
      .funcao-card-header { background: var(--green-800); padding: 10px 16px; font-weight: 700; color: var(--white); font-size: 10pt; border-bottom: 2px solid var(--green-100); }
      .funcao-card-body { padding: 12px 16px; }
      .funcao-mini-table { width: 100%; border-collapse: collapse; font-size: 8.5pt; margin: 8px 0; }
      .funcao-mini-table th { background: var(--green-700); color: var(--white); padding: 8px; text-align: left; }
      .funcao-mini-table td { padding: 5px 10px; border: 1px solid var(--gray-200); transition: background 0.2s; }
      .funcao-mini-table td:hover { background-color: #f1f8f5; cursor: text; }
      .funcao-mini-table td:first-child { font-weight: 600; color: var(--green-800); width: 120px; background: var(--gray-50); }
      .risco-intoleravel { background-color: #FEE2E2 !important; color: #991B1B; font-weight: bold; }
      .risco-substancial { background-color: #FFF3E0 !important; color: #E8590C; font-weight: bold; }
      .alerta-engenharia { background-color: #FEF08A !important; color: #854D0E; font-weight: bold; text-align: center;}
      h4 { color: var(--green-800); margin: 15px 0 5px 0; font-size: 9.5pt; text-transform: uppercase; }
      .aviso-editavel { background-color: var(--green-100); color: var(--green-900); padding: 10px; margin-bottom: 15px; border-radius: 4px; font-size: 9pt; border: 1px solid var(--green-500); }
    </style>
    """
    
    html_content = f"<!DOCTYPE html><html lang='pt-BR'><head><meta charset='UTF-8'>{css_base}</head><body>"
    html_content += "<div class='anexo-header'>ANEXO I - INVENTÁRIO DE RISCOS E ENQUADRAMENTO PREVIDENCIÁRIO</div>"
    html_content += "<div class='aviso-editavel'>💡 <strong>Atenção Equipe Técnica:</strong> Documento gerado automaticamente. Consolidação de FISPQs e avaliações de campo.</div>"
    
    df_pgr = pd.DataFrame(resultados_pgr)
    df_med = pd.DataFrame(resultados_medicos)
    ghes = set(df_pgr['GHE'].unique().tolist() + df_med['GHE'].unique().tolist() if not df_med.empty else [])
    
    for ghe in sorted(ghes):
        html_content += f"<div class='funcao-card'><div class='funcao-card-header' contenteditable='true'>📁 {ghe}</div><div class='funcao-card-body'>"
        
        pgr_ghe = df_pgr[df_pgr['GHE'] == ghe] if not df_pgr.empty else pd.DataFrame()
        if not pgr_ghe.empty:
            html_content += "<h4>⚙️ Inventário de Risco (NR-01)</h4><table class='funcao-mini-table'><thead><tr><th>Origem / FISPQ</th><th>Perigo Identificado</th><th>Sev.</th><th>Prob.</th><th>Nível de Risco</th><th>EPI Recomendado (NR-06)</th></tr></thead><tbody>"
            for _, row in pgr_ghe.iterrows():
                classe_risco = "risco-intoleravel" if row['NÍVEL DE RISCO'] == "INTOLERÁVEL" else "risco-substancial" if row['NÍVEL DE RISCO'] == "SUBSTANCIAL" else ""
                html_content += f"<tr><td contenteditable='true'>{row['Arquivo Origem']}</td><td contenteditable='true'>{row['Código GHS']} {row['Perigo Identificado']}</td><td contenteditable='true'>{row['Severidade']}</td><td contenteditable='true'>{row['Probabilidade']}</td><td contenteditable='true' class='{classe_risco}'>{row['NÍVEL DE RISCO']}</td><td contenteditable='true'>{row['EPI (NR-06)']}</td></tr>"
            html_content += "</tbody></table>"
            
        med_ghe = df_med[df_med['GHE'] == ghe] if not df_med.empty else pd.DataFrame()
        if not med_ghe.empty:
            html_content += "<h4>🩺 Diretrizes Médicas e Previdenciárias (NR-07 / NR-09 / NR-15 / Dec 3.048 / eSocial)</h4><table class='funcao-mini-table'><thead><tr><th>Cód / CAS</th><th>Agente (Físico/Químico/Bio)</th><th>Lim. Tol. (NR-15)</th><th>Nível Ação (NR-09)</th><th>Exame/IBE (NR-07)</th><th>Dec 3.048 (Aposent.)</th><th>Cód. eSocial (Tab 24)</th></tr></thead><tbody>"
            for _, row in med_ghe.iterrows():
                classe_alerta = "alerta-engenharia" if "NÃO MAPEADO" in row['Agente Químico'] else ""
                html_content += f"<tr><td contenteditable='true'>{row['Nº CAS']}</td><td contenteditable='true' class='{classe_alerta}'>{row['Agente Químico']}</td><td contenteditable='true' class='{classe_alerta}'>{row['Lim. Tolerância (NR-15)']}</td><td contenteditable='true' class='{classe_alerta}'>{row['Nível de Ação (NR-09)']}</td><td contenteditable='true' class='{classe_alerta}'>{row['IBE (NR-07)']}</td><td contenteditable='true' class='{classe_alerta}'>{row['Dec 3048']}</td><td contenteditable='true' class='{classe_alerta}'>{row['eSocial']}</td></tr>"
            html_content += "</tbody></table>"
            
        html_content += "</div></div>"
    
    html_content += "</body></html>"
    return html_content

# ==========================================
# SEÇÃO DE HISTÓRICO 
# ==========================================
if historico_selecionado:
    st.markdown("### 🗄️ Visualizando Relatório do Histórico")
    st.info("Este documento foi puxado do Banco de Dados SQLite. Pronto para auditoria ou envio ao eSocial.")
    
    aba_preview, aba_download_word = st.tabs(["👁️ Pré-visualizar Documento", "📄 Baixar em Word (.doc)"])
    
    with aba_preview:
        components.html(historico_selecionado, height=700, scrolling=True)
        
    with aba_download_word:
        st.download_button(
            label="📄 Baixar Relatório Histórico em Word",
            data=historico_selecionado.encode('utf-8'), 
            file_name="Relatorio_Historico_Seconci.doc", 
            mime="application/msword", 
            width="stretch"
        )

# ==========================================
# SEÇÃO DE CRIAÇÃO (ACESSO DIRETO)
# ==========================================
else:
    st.markdown("### 1️⃣ Upload das FISPQs (Agentes Químicos)")
    arquivos_fispq = st.file_uploader("Insira as FISPQs em PDF", type=["pdf"], accept_multiple_files=True)
    textos_pdfs = {}

    df_editado = pd.DataFrame()
    ghe_opcoes = ["Nenhum GHE definido"]

    if arquivos_fispq:
        with st.spinner("Lendo o conteúdo dos PDFs..."):
            for arquivo in arquivos_fispq:
                with pdfplumber.open(arquivo) as pdf:
                    texto = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
                    textos_pdfs[arquivo.name] = texto

        st.markdown("---")
        st.markdown("### 2️⃣ Definição de GHEs Químicos")
        
        nomes_arquivos = [arq.name for arq in arquivos_fispq]
        dados_iniciais = [{"GHE": "GHE 01 - Digite a Função", "Arquivo FISPQ": nome, "Probabilidade": 3} for nome in nomes_arquivos]
        df_mapeamento = pd.DataFrame(dados_iniciais)
        
        df_editado = st.data_editor(
            df_mapeamento, num_rows="dynamic",
            column_config={
                "GHE": st.column_config.TextColumn("Nome do GHE", required=True),
                "Arquivo FISPQ": st.column_config.SelectboxColumn("Arquivo (FISPQ)", options=nomes_arquivos, required=True),
                "Probabilidade": st.column_config.NumberColumn("Prob.", min_value=1, max_value=5, required=True)
            },
            width="stretch"
        )
        ghe_opcoes = df_editado["GHE"].unique().tolist() if not df_editado.empty else ["Nenhum GHE definido"]

    st.markdown("---")
    st.markdown("### 3️⃣ Avaliações de Campo: Agentes Físicos e Biológicos (Opcional)")
    
    agentes_opcoes = list(dicionario_fis_bio.keys())
    df_fis_bio_inicial = pd.DataFrame([{"GHE": ghe_opcoes[0], "Agente": agentes_opcoes[0], "Probabilidade": 3}])

    df_fis_bio_editado = st.data_editor(
        df_fis_bio_inicial, num_rows="dynamic",
        column_config={
            "GHE": st.column_config.SelectboxColumn("GHE de Destino", options=ghe_opcoes, required=True),
            "Agente": st.column_config.SelectboxColumn("Agente Físico/Biológico", options=agentes_opcoes, required=True),
            "Probabilidade": st.column_config.NumberColumn("Prob.", min_value=1, max_value=5, required=True)
        },
        width="stretch"
    )

    st.markdown("---")
    if st.button("🪄 Processar GHEs e Gerar Relatório Completo", width="stretch"):
        with st.spinner("Consolidando avaliações..."):
            resultados_pgr = []
            resultados_medicos = []
            
            # PROCESSAMENTO QUÍMICO
            if not df_editado.empty:
                for index, row in df_editado.iterrows():
                    nome_ghe = row["GHE"]
                    nome_arq = row["Arquivo FISPQ"]
                    v_prob = int(row["Probabilidade"])
                    
                    if nome_arq in textos_pdfs:
                        texto_completo = textos_pdfs[nome_arq]
                        
                        cas_encontrados_linha = list(set(re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', texto_completo)))
                        for cas in cas_encontrados_linha:
                            dados_med = dicionario_cas.get(cas, {
                                "agente": "⚠️ AGENTE NÃO MAPEADO", "nr15_lt": "REVISÃO DA ENGENHARIA", 
                                "nr09_acao": "REVISÃO DA ENGENHARIA", "nr07_ibe": "REVISÃO DA ENGENHARIA",
                                "dec_3048": "REVISÃO DA ENGENHARIA", "esocial_24": "REVISÃO DA ENGENHARIA"
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
                                s_val = dados_h["sev"]
                                nivel_risco = matriz_oficial.get((s_val, v_prob), "N/A")
                                resultados_pgr.append({
                                    "GHE": nome_ghe, "Arquivo Origem": nome_arq, "Código GHS": codigo,
                                    "Perigo Identificado": dados_h["desc"], "Severidade": texto_sev.get(s_val, str(s_val)),
                                    "Probabilidade": str(v_prob), "NÍVEL DE RISCO": nivel_risco,
                                    "Ação Requerida": acoes_requeridas.get(nivel_risco, "Manual"), "EPI (NR-06)": dados_h["epi"]
                                })

            # PROCESSAMENTO FÍSICO E BIOLÓGICO
            if not df_fis_bio_editado.empty and df_fis_bio_editado["GHE"].iloc[0] != "Nenhum GHE definido":
                for index, row in df_fis_bio_editado.iterrows():
                    nome_ghe = row["GHE"]
                    nome_agente = row["Agente"]
                    v_prob = int(row["Probabilidade"])
                    
                    if nome_agente in dicionario_fis_bio:
                        dados_fis = dicionario_fis_bio[nome_agente]
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
                st.success("✅ Relatório Consolidado Gerado!")
                st.session_state['ultimo_html'] = html_final

    # SALVAR NO BANCO
    if 'ultimo_html' in st.session_state:
        st.markdown("### 💾 Salvar Projeto no Banco de Dados")
        col1, col2 = st.columns([3, 1])
        with col1:
            nome_projeto = st.text_input("Nome da Empresa / Projeto:")
        with col2:
            st.write("")
            st.write("")
            if st.button("Gravar no Histórico", width="stretch"):
                if nome_projeto:
                    conn = sqlite3.connect('seconci_banco_dados.db')
                    c = conn.cursor()
                    data_atual = datetime.now().strftime("%d/%m/%Y %H:%M")
                    c.execute("INSERT INTO historico_laudos (nome_projeto, data_salvamento, html_relatorio) VALUES (?, ?, ?)", 
                              (nome_projeto, data_atual, st.session_state['ultimo_html']))
                    conn.commit()
                    conn.close()
                    st.success(f"Projeto salvo com sucesso!")
                else:
                    st.warning("Digite um nome para o projeto.")

        aba_preview, aba_download_word = st.tabs(["👁️ Pré-visualizar Documento", "📄 Baixar em Word (.doc)"])
        with aba_preview:
            components.html(st.session_state['ultimo_html'], height=700, scrolling=True)
        with aba_download_word:
            st.download_button(
                label="📄 Baixar Relatório em Word",
                data=st.session_state['ultimo_html'].encode('utf-8'), 
                file_name="Relatorio_Seconci.doc", 
                mime="application/msword", 
                width="stretch"
            )

# ==========================================
# 4. MÓDULO DE IA PARA HIGIENE OCUPACIONAL (GOOGLE GEMINI - REST API)
# ==========================================
if not historico_selecionado:
    st.markdown("---")
    st.markdown("### 🧠 Consultoria de Higiene Ocupacional (Cérebro Gemini 1.5 - REST)")
    
    if 'textos_pdfs' not in locals() or not textos_pdfs:
        st.info("Faça o upload de uma FISPQ no Passo 1 para ativar o chat de consultoria da IA.")
    else:
        st.success("✅ Protocolo de Conexão Direta Ativado. Conectado ao Google Gemini!")
        fispq_selecionada = st.selectbox("Selecione a FISPQ para análise da IA:", list(textos_pdfs.keys()))
        pergunta_ia = st.text_input("Sua pergunta técnica (Ex: Quais os EPIs indicados para este produto?)")
        
        if st.button("🤖 Perguntar ao Gemini"):
            if pergunta_ia:
                with st.spinner("O Gemini está analisando as Normas e o Documento via REST..."):
                    try:
                        texto_resumo = textos_pdfs[fispq_selecionada][:15000]
                        prompt_seguranca = f"""
                        Você é um Engenheiro de Segurança do Trabalho Sênior do Seconci-GO. 
                        Responda com precisão cirúrgica, baseando-se EXCLUSIVAMENTE nas Normas Regulamentadoras (NRs) brasileiras e no texto do documento abaixo. NUNCA invente informações.
                        
                        Documento (FISPQ): {texto_resumo}
                        
                        Pergunta do analista: {pergunta_ia}
                        """
                        
                        # 1. AUTO-DISCOVERY: Pergunta ao Google quais modelos a sua chave tem acesso hoje
                        url_lista = f"https://generativelanguage.googleapis.com/v1beta/models?key={CHAVE_API_GOOGLE}"
                        resp_lista = requests.get(url_lista)
                        
                        if resp_lista.status_code == 200:
                            modelos = resp_lista.json().get('models', [])
                            modelos_texto = [m['name'] for m in modelos if 'generateContent' in m.get('supportedGenerationMethods', [])]
                            
                            # Define a ordem de preferência (do mais moderno para o mais antigo)
                            modelo_escolhido = None
                            for pref in ['models/gemini-1.5-flash', 'models/gemini-1.5-pro', 'models/gemini-pro', 'models/gemini-1.0-pro']:
                                if pref in modelos_texto:
                                    modelo_escolhido = pref
                                    break
                                    
                            if not modelo_escolhido and modelos_texto:
                                modelo_escolhido = modelos_texto[0] # Fallback de segurança
                                
                            if modelo_escolhido:
                                # 2. Faz a chamada usando o modelo garantido que existe na sua chave
                                url_google = f"https://generativelanguage.googleapis.com/v1beta/{modelo_escolhido}:generateContent?key={CHAVE_API_GOOGLE}"
                                headers = {'Content-Type': 'application/json'}
                                payload = {
                                    "contents": [{"parts": [{"text": prompt_seguranca}]}],
                                    "generationConfig": {"temperature": 0.1} # Trava de alucinação no máximo!
                                }
                                
                                resposta = requests.post(url_google, headers=headers, json=payload)
                                
                                if resposta.status_code == 200:
                                    resultado_texto = resposta.json()['candidates'][0]['content']['parts'][0]['text']
                                    st.success(f"Análise concluída com precisão! (Motor utilizado: {modelo_escolhido.split('/')[-1]})")
                                    st.write(resultado_texto)
                                else:
                                    st.error(f"Erro na geração da resposta: {resposta.text}")
                            else:
                                st.error("A sua chave de API é válida, mas o Google ainda não habilitou modelos de texto para ela.")
                        else:
                            st.error(f"Sua chave de API está bloqueada pelo Google. Retorno: {resp_lista.status_code}")
                            
                    except Exception as e:
                        st.error(f"Falha na conexão de rede: {e}")
            else:
                st.warning("Digite uma pergunta para a IA.")
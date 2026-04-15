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
# CONFIGURAÇÃO DA PÁGINA E CSS
# ==========================================
st.set_page_config(page_title="Automação SST - Seconci GO", layout="wide", page_icon="🛡️")

css_personalizado = """
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .stButton > button { background-color: #084D22; color: white; border-radius: 8px; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1); font-weight: 600; padding: 0.5rem 1rem; }
    .stButton > button:hover { background-color: #1AA04B; color: white; }
    h1, h2, h3 { color: #084D22 !important; }
    [data-testid="stSidebar"] { background-color: #F4F8F5; }
    .stAlert { border-left: 5px solid #084D22; }
</style>
"""
st.markdown(css_personalizado, unsafe_allow_html=True)

CHAVE_API_GOOGLE = str(st.secrets["CHAVE_API_GOOGLE"]).strip().replace('"', '').replace("'", "")

# ==========================================
# MOTOR IA - AUTO-DISCOVERY (PREVINE ERRO 404)
# ==========================================
@st.cache_data(ttl=3600) # Faz cache por 1 hora para não gastar requisições à toa
def descobrir_modelo_valido(chave_api):
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={chave_api}"
    try:
        resp = requests.get(url, timeout=15)
        if resp.status_code == 200:
            dados = resp.json()
            # Extrai apenas modelos que suportam geração de texto
            modelos_disponiveis = [
                m['name'] for m in dados.get('models', []) 
                if 'generateContent' in m.get('supportedGenerationMethods', [])
            ]
            
            # Tenta pegar o melhor modelo na ordem do que estiver liberado na sua chave
            ordem_preferencia = ['models/gemini-1.5-flash', 'models/gemini-1.5-pro', 'models/gemini-pro', 'models/gemini-1.0-pro']
            for pref in ordem_preferencia:
                if pref in modelos_disponiveis:
                    return pref
            
            # Se não achar os preferidos, pega absolutamente qualquer um que suporte texto
            if modelos_disponiveis:
                return modelos_disponiveis[0]
                
            return None
        else:
            st.sidebar.error(f"Erro na verificação de modelos: {resp.status_code}")
            return None
    except Exception as e:
        st.sidebar.error(f"Falha de rede ao buscar modelos: {e}")
        return None

def chamar_api_gemini(prompt, pdf_b64=None):
    modelo_exato = descobrir_modelo_valido(CHAVE_API_GOOGLE)
    
    if not modelo_exato:
        st.error("🚨 ERRO FATAL: Nenhum modelo do Google suportado foi encontrado para a sua chave de API.")
        return None
        
    url_geracao = f"https://generativelanguage.googleapis.com/v1beta/{modelo_exato}:generateContent?key={CHAVE_API_GOOGLE}"
    
    parts = [{"text": prompt}]
    if pdf_b64:
        parts.append({"inlineData": {"mimeType": "application/pdf", "data": pdf_b64}})
        
    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {"temperature": 0.1},
        "safetySettings": [
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_ONLY_HIGH"}
        ]
    }
    
    try:
        resp = requests.post(url_geracao, headers={'Content-Type': 'application/json'}, json=payload, timeout=50)
        if resp.status_code == 200:
            return resp.json().get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
        else:
            st.error(f"🚨 ERRO HTTP DA API DO GOOGLE ({resp.status_code}): Endpoint {modelo_exato} recusou a carga. Mensagem: {resp.text}")
            return None
    except Exception as e:
        st.error(f"🚨 FALHA DE CONEXÃO COM O GOOGLE: {e}")
        return None

def buscar_dados_completos_ia(cas_faltantes):
    if not cas_faltantes: return {}
    
    cas_limpos = [str(c).strip() for c in cas_faltantes]
    lista_cas_str = ", ".join(cas_limpos)
    
    prompt = f"""
    Você é um Engenheiro de Segurança do Trabalho e Especialista em Higiene Ocupacional no Brasil.
    Para os seguintes números CAS: {lista_cas_str}, forneça os parâmetros legais vigentes.
    
    CRÍTICO: Retorne APENAS um objeto JSON válido. Sem introdução, sem marcadores markdown. As chaves do JSON DEVEM SER os números CAS.
    
    Exemplo:
    {{
        "{cas_limpos[0] if cas_limpos else '1317-65-3'}": {{
            "agente": "Nome Químico",
            "nr15_lt": "10 mg/m³ ou Avaliar Anexo 12",
            "nr09_acao": "5 mg/m³ ou N/A",
            "nr07_ibe": "Raio-X OIT ou Avaliação Clínica",
            "dec_3048": "25 anos ou Não Enquadrado",
            "esocial_24": "01.18.001 ou 09.01.001"
        }}
    }}
    """
    
    texto_ia = chamar_api_gemini(prompt)
    if texto_ia:
        try:
            texto_limpo = texto_ia.replace('```json', '').replace('```', '').strip()
            return json.loads(texto_limpo)
        except json.JSONDecodeError as e:
            st.error(f"🚨 A IA não retornou um JSON estruturado. Erro de Parsing. Texto recebido: {texto_limpo}")
            return {}
    return {}

# ==========================================
# BANCO DE DADOS E DICIONÁRIOS LOCAIS
# ==========================================
def init_db():
    conn = sqlite3.connect('seconci_banco_dados.db')
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS historico_laudos (id INTEGER PRIMARY KEY AUTOINCREMENT, nome_projeto TEXT, data_salvamento TEXT, html_relatorio TEXT)')
    conn.commit()
    conn.close()
init_db()

dicionario_h = {
    "H315": {"desc": "Provoca irritação à pele", "sev": 1, "epi": "Luvas e vestimenta"},
    "H319": {"desc": "Provoca irritação ocular grave", "sev": 1, "epi": "Óculos de proteção"},
    "H317": {"desc": "Reações alérgicas na pele", "sev": 2, "epi": "Luvas longas"},
    "H351": {"desc": "Suspeito de provocar câncer", "sev": 4, "epi": "Proteção respiratória e dérmica máxima"},
    "H373": {"desc": "Danos aos órgãos (exp. repetida)", "sev": 4, "epi": "EPIs combinados obrigatórios"},
    "H302": {"desc": "Nocivo em caso de ingestão", "sev": 2, "epi": "Higiene rigorosa; luvas"}
}

matriz_oficial = {
    (1,3): "TOLERÁVEL", (2,3): "MODERADO", (3,3): "MODERADO", (4,3): "SUBSTANCIAL", (5,3): "SUBSTANCIAL",
    (1,4): "TOLERÁVEL", (2,4): "MODERADO", (3,4): "SUBSTANCIAL", (4,4): "INTOLERÁVEL", (5,4): "INTOLERÁVEL"
}

texto_sev = {1: "1-LEVE", 2: "2-BAIXA", 3: "3-MODERADA", 4: "4-ALTA", 5: "5-EXTREMA"}

dicionario_cas = {
    "108-88-3": {"agente": "Tolueno", "nr15_lt": "78 ppm", "nr09_acao": "39 ppm", "nr07_ibe": "o-Cresol", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "65997-15-1": {"agente": "Cimento Portland", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Raio-X e Espirometria", "dec_3048": "Não Enquadrado", "esocial_24": "01.18.001"}
}

dicionario_campo = {
    "Físico: Ruído Contínuo/Intermitente": {"agente": "Ruído", "nr15_lt": "85 dB(A)", "nr09_acao": "80 dB(A)", "nr07_ibe": "Audiometria", "dec_3048": "25 anos", "esocial_24": "02.01.001", "perigo": "Pressão sonora", "sev": 3, "epi": "Protetor Auditivo"}
}

matriz_risco_exame = {
    "TOLUENO": {"exame": "Ortocresol na Urina", "periodico": "6 MESES"},
    "RUÍDO": {"exame": "Audiometria", "periodico": "12 MESES"}
}

matriz_funcao_exame = {
    "ALTURA": [{"exame": "ECG", "periodicidade": "12 MESES"}, {"exame": "Glicemia", "periodicidade": "12 MESES"}],
    "ENCANADOR": [{"exame": "Audiometria", "periodicidade": "12 MESES"}]
}

def processar_pcmso(dados_pgr_json):
    tabela_pcmso = []
    for ghe in dados_pgr_json:
        nome_ghe = ghe.get("ghe", "Sem GHE")
        cargos = ghe.get("cargos", [])
        riscos = ghe.get("riscos_mapeados", [])
        
        for cargo in cargos:
            exames_cargo = [{"exame": "Exame Clínico (Anamnese/Físico)", "periodicidade": "12 MESES", "motivo": "NR-07 Básico"}]
            cargo_upper = cargo.upper()
            
            for funcao, exames in matriz_funcao_exame.items():
                if funcao in cargo_upper: exames_cargo.extend(exames)
            
            for risco in riscos:
                txt_risco = (risco.get("nome_agente", "") + " " + risco.get("perigo_especifico", "")).upper()
                for agente, regra in matriz_risco_exame.items():
                    if agente in txt_risco:
                        exames_cargo.append({"exame": regra["exame"], "periodicidade": regra["periodico"], "motivo": f"Exposição a {agente}"})
                if "ALTURA" in txt_risco: exames_cargo.extend(matriz_funcao_exame["ALTURA"])
            
            unicos = {v['exame']:v for v in exames_cargo}.values()
            for ex in unicos:
                tabela_pcmso.append({"GHE / Setor": nome_ghe, "Cargo": cargo, "Exame Clínico/Complementar": ex["exame"], "Periodicidade": ex["periodicidade"], "Justificativa Legal / Risco": ex.get("motivo", "Protocolo Função")})
    return pd.DataFrame(tabela_pcmso)

# ==========================================
# GERADORES HTML
# ==========================================
def gerar_html_anexo(resultados_pgr, resultados_medicos):
    html = """<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><style>body { font-family: 'Arial', sans-serif; font-size: 10pt; color: #000; } .anexo-header { background-color: #084D22; color: #FFF; padding: 14px 20px; font-weight: bold; margin-bottom: 20px; text-align: center; } .funcao-card { border: 1px solid #084D22; margin-bottom: 20px; } .funcao-card-header { background-color: #084D22; padding: 10px; font-weight: bold; color: #FFF; } .funcao-mini-table { width: 100%; border-collapse: collapse; font-size: 9pt; margin: 8px 0; } .funcao-mini-table th { background-color: #0F823B; color: #FFF; padding: 8px; text-align: left; border: 1px solid #000; } .funcao-mini-table td { padding: 5px; border: 1px solid #000; } h4 { color: #084D22; margin: 15px 0 5px 0; }</style></head><body><div class='anexo-header'>ANEXO I - INVENTÁRIO DE RISCOS E ENQUADRAMENTO PREVIDENCIÁRIO</div>"""
    df_pgr, df_med = pd.DataFrame(resultados_pgr), pd.DataFrame(resultados_medicos)
    ghes = set(df_pgr['GHE'].unique().tolist() + df_med['GHE'].unique().tolist() if not df_med.empty else [])
    
    for ghe in sorted(ghes):
        html += f"<div class='funcao-card'><div class='funcao-card-header'>GHE: {ghe}</div><div>"
        pgr_ghe = df_pgr[df_pgr['GHE'] == ghe] if not df_pgr.empty else pd.DataFrame()
        if not pgr_ghe.empty:
            html += "<h4>Inventário de Risco (NR-01)</h4><table class='funcao-mini-table'><thead><tr><th>Origem / FISPQ</th><th>Perigo Identificado</th><th>Sev.</th><th>Prob.</th><th>Nível de Risco</th><th>EPI Recomendado (NR-06)</th></tr></thead><tbody>"
            for _, r in pgr_ghe.iterrows(): html += f"<tr><td>{r['Arquivo Origem']}</td><td>{r['Código GHS']} {r['Perigo Identificado']}</td><td>{r['Severidade']}</td><td>{r['Probabilidade']}</td><td>{r['NÍVEL DE RISCO']}</td><td>{r['EPI (NR-06)']}</td></tr>"
            html += "</tbody></table>"
            
        med_ghe = df_med[df_med['GHE'] == ghe] if not df_med.empty else pd.DataFrame()
        if not med_ghe.empty:
            html += "<h4>Diretrizes Médicas e Previdenciárias (Automação IA)</h4><table class='funcao-mini-table'><thead><tr><th>Cód / CAS</th><th>Agente</th><th>Lim. Tol. (NR-15)</th><th>Nível Ação (NR-09)</th><th>Exame/IBE (NR-07)</th><th>Dec 3048</th><th>eSocial</th></tr></thead><tbody>"
            for _, r in med_ghe.iterrows(): html += f"<tr><td>{r['Nº CAS']}</td><td>{r['Agente Químico']}</td><td>{r['Lim. Tolerância (NR-15)']}</td><td>{r['Nível de Ação (NR-09)']}</td><td>{r['IBE (NR-07)']}</td><td>{r['Dec 3048']}</td><td>{r['eSocial']}</td></tr>"
            html += "</tbody></table>"
        html += "</div></div>"
    return html + "</body></html>"

def gerar_html_pcmso(df_pcmso):
    html = """<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><style>body { font-family: 'Arial'; color: #000; } .header { background-color: #084D22; color: #FFF; padding: 14px; text-align: center; font-weight: bold; margin-bottom: 20px;} table { width: 100%; border-collapse: collapse; font-size: 13px; } th { background-color: #1AA04B; color: #FFF; padding: 12px 8px; text-align: left; border: 1px solid #000; } td { border: 1px solid #000; padding: 10px 8px; }</style></head><body><div class='header'>MATRIZ DE EXAMES - PCMSO</div><table><tr><th>GHE / Setor</th><th>Cargo</th><th>Exame Clínico / Complementar</th><th>Periodicidade</th><th>Justificativa / Agente</th></tr>"""
    for _, r in df_pcmso.iterrows(): html += f"<tr><td><strong>{r['GHE / Setor']}</strong></td><td>{r['Cargo']}</td><td>{r['Exame Clínico/Complementar']}</td><td>{r['Periodicidade']}</td><td>{r['Justificativa Legal / Risco']}</td></tr>"
    return html + "</table></body></html>"

# ==========================================
# INTERFACE PRINCIPAL
# ==========================================
st.title("Sistema Integrado SST - Seconci GO 🚀")
modulo_selecionado = st.sidebar.radio("Selecione:", ["1️⃣ Engenharia: FISPQ ➡️ PGR", "2️⃣ Medicina: PGR ➡️ PCMSO"])

if "1️⃣" in modulo_selecionado:
    st.header("Módulo de Engenharia: Extrator de FISPQs (Automação Dinâmica)")
    st.info("A API está configurada para mapear e conectar-se dinamicamente ao modelo liberado pela sua chave.")
    
    arquivos = st.file_uploader("Insira as FISPQs (PDF)", type=["pdf"], accept_multiple_files=True)
    textos_pdfs = {}
    ghe_opcoes = ["GHE 01 - Função"]
    
    if arquivos:
        for arq in arquivos:
            with pdfplumber.open(arq) as pdf:
                textos_pdfs[arq.name] = "\n".join([p.extract_text() or "" for p in pdf.pages])
                
        df_editado = st.data_editor(pd.DataFrame([{"GHE": "GHE 01", "Arquivo": a.name, "Prob.": 3} for a in arquivos]), num_rows="dynamic", width="stretch")
        ghe_opcoes = df_editado["GHE"].unique().tolist() if not df_editado.empty else ghe_opcoes

    df_fis = st.data_editor(pd.DataFrame([{"GHE": ghe_opcoes[0], "Agente": list(dicionario_campo.keys())[0], "Prob.": 3}]), num_rows="dynamic", width="stretch")

    if st.button("Processar Relatório", type="primary", use_container_width=True):
        with st.spinner("Conectando aos Servidores do Google e Estruturando Tabela..."):
            res_pgr, res_med = [], []
            if arquivos and not df_editado.empty:
                for _, row in df_editado.iterrows():
                    nome_ghe, arq, prob = row["GHE"], row["Arquivo"], int(row["Prob."])
                    texto = textos_pdfs.get(arq, "")
                    cas_found = list(set(re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', texto)))
                    cas_missing = [c for c in cas_found if c not in dicionario_cas]
                    
                    ia_data = buscar_dados_completos_ia(cas_missing) if cas_missing else {}
                    
                    for cas in cas_found:
                        c_limpo = cas.strip()
                        dados = dicionario_cas.get(c_limpo, ia_data.get(c_limpo, {"agente": "Produto Químico", "nr15_lt": "Avaliar Anexo 11/12", "nr09_acao": "Avaliar NR-09", "nr07_ibe": "Avaliação Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"}))
                        res_med.append({"GHE": nome_ghe, "Arquivo Origem": arq, "Nº CAS": cas, "Agente Químico": dados["agente"], "Lim. Tolerância (NR-15)": dados.get("nr15_lt", "N/A"), "Nível de Ação (NR-09)": dados.get("nr09_acao", "N/A"), "IBE (NR-07)": dados.get("nr07_ibe", "N/A"), "Dec 3048": dados.get("dec_3048", "N/A"), "eSocial": dados.get("esocial_24", "N/A")})
                    
                    for h_cod in list(set(re.findall(r'H\d{3}', texto))):
                        if h_cod in dicionario_h:
                            d_h = dicionario_h[h_cod]
                            n_risco = matriz_oficial.get((d_h["sev"], prob), "MODERADO")
                            res_pgr.append({"GHE": nome_ghe, "Arquivo Origem": arq, "Código GHS": h_cod, "Perigo Identificado": d_h["desc"], "Severidade": texto_sev.get(d_h["sev"], str(d_h["sev"])), "Probabilidade": str(prob), "NÍVEL DE RISCO": n_risco, "EPI (NR-06)": d_h["epi"]})

            if res_pgr or res_med:
                html_out = gerar_html_anexo(res_pgr, res_med)
                st.session_state['html_pgr'] = html_out
                st.success("Relatório gerado!")

    if 'html_pgr' in st.session_state:
        st.download_button("Baixar PGR em Word", st.session_state['html_pgr'].encode('utf-8'), "PGR.doc", "application/msword", use_container_width=True)
        components.html(st.session_state['html_pgr'], height=500, scrolling=True)

elif "2️⃣" in modulo_selecionado:
    st.header("Módulo Médico: Importador PGR -> PCMSO")
    arq_pgr = st.file_uploader("Upload PDF PGR", type=["pdf"])
    if arq_pgr and st.button("Extrair Riscos e Gerar Matriz PCMSO", type="primary", use_container_width=True):
        with st.spinner("Analisando matrizes via IA..."):
            pdf_b64 = base64.b64encode(arq_pgr.getvalue()).decode('utf-8')
            prompt = """Analise o texto deste PGR. Retorne EXATAMENTE UM JSON contendo a relação de GHE, cargos e riscos mapeados:
            [{"ghe": "Nome do GHE", "cargos": ["Cargo 1"], "riscos_mapeados": [{"nome_agente": "Ex: Ruído", "perigo_especifico": "Físico"}]}]"""
            texto_ia = chamar_api_gemini(prompt, pdf_b64)
            if texto_ia:
                try:
                    df = processar_pcmso(json.loads(texto_ia.replace('```json', '').replace('```', '').strip()))
                    st.session_state['html_pcmso'] = gerar_html_pcmso(df)
                    st.success("Matriz extraída!")
                except Exception as e:
                    st.error(f"Falha de formatação dos dados da IA: {e}")
    if 'html_pcmso' in st.session_state:
        st.download_button("Baixar PCMSO em Word", st.session_state['html_pcmso'].encode('utf-8'), "PCMSO.doc", "application/msword", use_container_width=True)
        components.html(st.session_state['html_pcmso'], height=600, scrolling=True)

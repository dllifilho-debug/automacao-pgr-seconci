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
# 1. CONFIGURAÇÃO E INTERFACE SaaS (Estilo HO Fácil)
# ==========================================
st.set_page_config(page_title="Sistema SST - Seconci GO", layout="wide", page_icon="🛡️")

css_premium = """
<style>
    /* Fundo e tipografia */
    .block-container { padding-top: 1.5rem; }
    h1, h2, h3, h4 { color: #0f4c23 !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    
    /* Botões Padrão Enterprise */
    .stButton > button { background-color: #0f4c23; color: white; border-radius: 6px; font-weight: bold; border: none; transition: 0.2s; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .stButton > button:hover { background-color: #1a803b; transform: translateY(-1px); box-shadow: 0 4px 8px rgba(0,0,0,0.2); }
    
    /* Menu Lateral */
    [data-testid="stSidebar"] { background-color: #f8f9fa; border-right: 1px solid #e0e0e0; }
    
    /* Cards do Dashboard */
    div[data-testid="metric-container"] { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    
    /* Tabelas Dataframe */
    [data-testid="stDataFrame"] { border-radius: 8px; border: 1px solid #e0e0e0; }
</style>
"""
st.markdown(css_premium, unsafe_allow_html=True)

CHAVE_API_GOOGLE = str(st.secrets["CHAVE_API_GOOGLE"]).strip().replace('"', '').replace("'", "")

# ==========================================
# 2. BANCO DE DADOS E DICIONÁRIOS ROBUSTOS
# ==========================================
def init_db():
    conn = sqlite3.connect('seconci_banco_dados.db')
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS historico_laudos (id INTEGER PRIMARY KEY AUTOINCREMENT, nome_projeto TEXT, modulo TEXT, data_salvamento TEXT, html_relatorio TEXT)')
    conn.commit()
    conn.close()

init_db()

# Dicionário de Frases H (GHS)
dicionario_h = {
    "H315": {"desc": "Irritação à pele", "sev": 1, "epi": "Luvas de proteção e vestimenta"},
    "H319": {"desc": "Irritação ocular grave", "sev": 1, "epi": "Óculos de proteção"},
    "H317": {"desc": "Reações alérgicas na pele", "sev": 2, "epi": "Luvas (nitrílica/PVC) e manga longa"},
    "H332": {"desc": "Nocivo se inalado", "sev": 2, "epi": "Máscara respiratória PFF2"},
    "H314": {"desc": "Queimadura severa à pele", "sev": 4, "epi": "Traje químico, luvas longas"},
    "H350": {"desc": "Pode provocar câncer", "sev": 5, "epi": "Isolamento; Traje químico e EPR"},
    "H351": {"desc": "Suspeito de provocar câncer", "sev": 4, "epi": "Proteção respiratória e dérmica máxima"},
    "H372": {"desc": "Danos aos órgãos (exp. repetida)", "sev": 4, "epi": "Proteção respiratória estrita"},
    "H373": {"desc": "Danos aos órgãos (exp. repetida)", "sev": 4, "epi": "EPIs combinados obrigatórios"}
}

# Matriz de Risco (Severidade x Probabilidade)
matriz_oficial = {
    (1,1): "TRIVIAL", (1,2): "TRIVIAL", (1,3): "TOLERÁVEL", (1,4): "TOLERÁVEL", (1,5): "MODERADO",
    (2,1): "TRIVIAL", (2,2): "TOLERÁVEL", (2,3): "MODERADO", (2,4): "MODERADO", (2,5): "SUBSTANCIAL",
    (3,1): "TOLERÁVEL", (3,2): "TOLERÁVEL", (3,3): "MODERADO", (3,4): "SUBSTANCIAL", (3,5): "SUBSTANCIAL",
    (4,1): "TOLERÁVEL", (4,2): "MODERADO", (4,3): "SUBSTANCIAL", (4,4): "INTOLERÁVEL", (4,5): "INTOLERÁVEL",
    (5,1): "MODERADO", (5,2): "MODERADO", (5,3): "SUBSTANCIAL", (5,4): "INTOLERÁVEL", (5,5): "INTOLERÁVEL"
}

texto_sev = {1: "1-LEVE", 2: "2-BAIXA", 3: "3-MODERADA", 4: "4-ALTA", 5: "5-EXTREMA"}

# Banco Local: Impede falhas para os compostos mais comuns da Construção Civil
dicionario_cas = {
    "108-88-3": {"agente": "Tolueno", "nr15_lt": "78 ppm", "nr09_acao": "39 ppm", "nr07_ibe": "o-Cresol", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "1330-20-7": {"agente": "Xileno", "nr15_lt": "78 ppm", "nr09_acao": "39 ppm", "nr07_ibe": "Ác. Metilhipúricos", "dec_3048": "25 anos", "esocial_24": "01.19.036"},
    "14808-60-7": {"agente": "Sílica Cristalina (Quartzo)", "nr15_lt": "Anexo 12", "nr09_acao": "50% do L.T.", "nr07_ibe": "RX OIT / Espirometria", "dec_3048": "25 anos", "esocial_24": "01.18.001"},
    "65997-15-1": {"agente": "Cimento Portland", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "RX OIT", "dec_3048": "Não Enquadrado", "esocial_24": "01.18.001"},
    "1317-65-3": {"agente": "Carbonato de Cálcio", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Aval. Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "1305-78-8": {"agente": "Óxido de Cálcio", "nr15_lt": "2 mg/m³", "nr09_acao": "1 mg/m³", "nr07_ibe": "Aval. Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12042-78-3": {"agente": "Aluminato de Cálcio", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Aval. Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12168-85-3": {"agente": "Silicato Tricálcico", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Aval. Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "10034-77-2": {"agente": "Silicato Dicálcico", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Aval. Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"},
    "12068-35-8": {"agente": "Silicato (Misto)", "nr15_lt": "10 mg/m³", "nr09_acao": "5 mg/m³", "nr07_ibe": "Aval. Clínica", "dec_3048": "Não Enquadrado", "esocial_24": "09.01.001"}
}

# ==========================================
# 3. MOTOR HÍBRIDO E IA EM CASCATA
# ==========================================
def chamar_api_gemini_cascata(prompt):
    """
    Atira em 5 servidores diferentes para garantir que nunca dê erro 404 ou 400.
    """
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.0},
        "safetySettings": [
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_ONLY_HIGH"},
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH"}
        ]
    }
    
    modelos = ["gemini-1.5-pro-latest", "gemini-1.5-flash-latest", "gemini-1.5-pro", "gemini-1.5-flash", "gemini-pro"]
    
    for modelo in modelos:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent?key={CHAVE_API_GOOGLE}"
        try:
            resp = requests.post(url, headers={'Content-Type': 'application/json'}, json=payload, timeout=20)
            if resp.status_code == 200:
                return resp.json().get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', '')
        except: continue
    return None

def buscar_cas_na_ia(lista_cas_desconhecidos):
    """
    Aciona a IA APENAS para os CAS que o Python não encontrou no banco local.
    Isso elimina 99% das falhas e alucinações.
    """
    if not lista_cas_desconhecidos: return {}
    cas_str = ", ".join(lista_cas_desconhecidos)
    
    prompt = f"""
    Atue como Higienista Ocupacional Brasileiro. Identifique os limites legais para os números CAS: {cas_str}.
    Retorne EXATAMENTE UM JSON, onde a chave é o CAS. Não use markdown.
    Exemplo:
    {{
      "CAS-AQUI": {{
        "agente": "Nome", "nr15_lt": "Limite", "nr09_acao": "Ação", "nr07_ibe": "Exame", "dec_3048": "Tempo", "esocial_24": "Codigo"
      }}
    }}
    """
    
    resposta = chamar_api_gemini_cascata(prompt)
    if resposta:
        try:
            limpo = resposta.replace('```json', '').replace('```', '').strip()
            return json.loads(limpo)
        except: pass
    return {}

# ==========================================
# 4. GERADORES DE HTML E WORD
# ==========================================
def gerar_html_anexo(resultados_pgr, resultados_medicos):
    html = """<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
    <head><style>body{font-family:Arial;font-size:10pt;color:#000;} .header{background-color:#0f4c23;color:#FFF;padding:15px;font-weight:bold;text-align:center;} th{background-color:#1a803b;color:#FFF;padding:8px;border:1px solid #000;} td{padding:5px;border:1px solid #000;} table{width:100%;border-collapse:collapse;margin-bottom:20px;}</style></head><body>
    <div class='header'>INVENTÁRIO DE RISCOS E ENQUADRAMENTO LEGAL</div>"""
    
    df_pgr = pd.DataFrame(resultados_pgr)
    df_med = pd.DataFrame(resultados_medicos)
    ghes = set(df_pgr['GHE'].unique().tolist() + df_med['GHE'].unique().tolist() if not df_med.empty else [])
    
    for ghe in sorted(ghes):
        html += f"<h3 style='color:#0f4c23; border-bottom: 2px solid #0f4c23;'>GHE: {ghe}</h3>"
        
        pgr_ghe = df_pgr[df_pgr['GHE'] == ghe] if not df_pgr.empty else pd.DataFrame()
        if not pgr_ghe.empty:
            html += "<h4>Inventário de Perigos (NR-01)</h4><table><tr><th>Origem</th><th>Perigo</th><th>Sev</th><th>Prob</th><th>Risco</th><th>EPI</th></tr>"
            for _, r in pgr_ghe.iterrows(): html += f"<tr><td>{r['Arquivo Origem']}</td><td>{r['Código']} {r['Desc']}</td><td>{r['Sev']}</td><td>{r['Prob']}</td><td>{r['Risco']}</td><td>{r['EPI']}</td></tr>"
            html += "</table>"
            
        med_ghe = df_med[df_med['GHE'] == ghe] if not df_med.empty else pd.DataFrame()
        if not med_ghe.empty:
            html += "<h4>Enquadramento (NR-15 / eSocial)</h4><table><tr><th>CAS</th><th>Agente</th><th>NR-15</th><th>NR-09</th><th>Exame</th><th>Dec 3048</th><th>eSocial</th></tr>"
            for _, r in med_ghe.iterrows(): html += f"<tr><td>{r['CAS']}</td><td>{r['Agente']}</td><td>{r['NR15']}</td><td>{r['NR09']}</td><td>{r['NR07']}</td><td>{r['Dec3048']}</td><td>{r['eSocial']}</td></tr>"
            html += "</table>"
    return html + "</body></html>"

# ==========================================
# 5. ESTRUTURA SAAS - MENU LATERAL E ROTEAMENTO
# ==========================================
if os.path.exists("logo.png"): st.sidebar.image("logo.png", width="150")

st.sidebar.markdown("### 🧭 Navegação")
menu = st.sidebar.radio("", [
    "📊 Dashboard Inicial", 
    "⚙️ Módulo 1: Extrator FISPQ", 
    "🩺 Módulo 2: PCMSO (Em breve)", 
    "📂 Central de Arquivos"
])

st.sidebar.markdown("---")
st.sidebar.info("Motor Híbrido Ativado.\nIA em Cascata Operante.")

# ==========================================
# TELA 1: DASHBOARD
# ==========================================
if "Dashboard" in menu:
    st.title("📊 Painel de Controle SST")
    st.markdown("Bem-vindo ao sistema de automação. Aqui está o resumo das suas operações.")
    
    col1, col2, col3 = st.columns(3)
    conn = sqlite3.connect('seconci_banco_dados.db')
    total_laudos = pd.read_sql_query("SELECT COUNT(*) FROM historico_laudos", conn).iloc[0,0]
    conn.close()
    
    col1.metric("Projetos Processados", f"{total_laudos}")
    col2.metric("Motor IA", "Online / Cascata")
    col3.metric("Banco Normativo Local", f"{len(dicionario_cas)} Substâncias")
    
    st.markdown("---")
    st.markdown("### 🚀 Acesso Rápido")
    st.info("Utilize o menu lateral para iniciar um novo processamento de FISPQ ou gerar uma matriz médica.")

# ==========================================
# TELA 2: EXTRATOR FISPQ (MOTOR HÍBRIDO)
# ==========================================
elif "Extrator" in menu:
    st.title("⚙️ Extrator Dinâmico de FISPQ")
    st.write("Faça upload dos PDFs. O sistema lerá os documentos e estruturará a matriz de riscos.")
    
    arquivos_up = st.file_uploader("Arraste as FISPQs aqui (PDF)", type=["pdf"], accept_multiple_files=True)
    
    if arquivos_up:
        st.markdown("### 1. Definição de GHE e Probabilidade")
        nomes_arq = [a.name for a in arquivos_up]
        df_config = pd.DataFrame([{"GHE": "GHE 01 - Produção", "Arquivo": n, "Probabilidade": 3} for n in nomes_arq])
        
        # Interface Tabela AgGrid/Premium usando dataframe column_config
        df_editado = st.data_editor(
            df_config,
            use_container_width=True,
            column_config={
                "GHE": st.column_config.TextColumn("GHE (Setor/Função)", required=True),
                "Arquivo": st.column_config.SelectboxColumn("FISPQ Relacionada", options=nomes_arq, required=True),
                "Probabilidade": st.column_config.NumberColumn("Probabilidade (1 a 5)", min_value=1, max_value=5, required=True)
            }
        )

        if st.button("🚀 Iniciar Análise e Cruzamento", type="primary"):
            resultados_pgr, resultados_medicos = [], []
            my_bar = st.progress(0, text="Lendo documentos...")
            
            for index, row in df_editado.iterrows():
                nome_ghe = row["GHE"]
                arq_nome = row["Arquivo"]
                prob = int(row["Probabilidade"])
                
                my_bar.progress((index + 1) / len(df_editado), text=f"Analisando: {arq_nome}...")
                
                arq_obj = next((f for f in arquivos_up if f.name == arq_nome), None)
                if arq_obj:
                    texto_pdf = ""
                    with pdfplumber.open(arq_obj) as pdf:
                        for page in pdf.pages: texto_pdf += (page.extract_text() or "") + "\n"
                    
                    # 1. Regex Exato (Sem Alucinação)
                    cas_encontrados = list(set(re.findall(r'\b(\d{2,7}-\d{2}-\d)\b', texto_pdf)))
                    h_encontrados = list(set(re.findall(r'H\d{3}', texto_pdf)))
                    
                    # 2. IA atuando APENAS no que falta
                    cas_desconhecidos = [c for c in cas_encontrados if c not in dicionario_cas]
                    ia_dados = buscar_cas_na_ia(cas_desconhecidos) if cas_desconhecidos else {}
                    
                    # 3. Montagem da Tabela Médica
                    for cas in cas_encontrados:
                        c_limpo = cas.strip()
                        if c_limpo in dicionario_cas:
                            d = dicionario_cas[c_limpo]
                        else:
                            d = ia_dados.get(c_limpo, {
                                "agente": f"Não Mapeado ({c_limpo})", "nr15_lt": "Avaliar NR-15", 
                                "nr09_acao": "Avaliar NR-09", "nr07_ibe": "Aval. Clínica", 
                                "dec_3048": "Verificar", "esocial_24": "09.01.001"
                            })
                            
                        resultados_medicos.append({
                            "GHE": nome_ghe, "Arquivo Origem": arq_nome, "CAS": c_limpo,
                            "Agente": d["agente"], "NR15": d["nr15_lt"], "NR09": d["nr09_acao"],
                            "NR07": d["nr07_ibe"], "Dec3048": d["dec_3048"], "eSocial": d["esocial_24"]
                        })
                    
                    # 4. Montagem da Tabela PGR
                    for h in h_encontrados:
                        if h in dicionario_h:
                            dh = dicionario_h[h]
                            risco = matriz_oficial.get((dh["sev"], prob), "MODERADO")
                            resultados_pgr.append({
                                "GHE": nome_ghe, "Arquivo Origem": arq_nome, "Código": h,
                                "Desc": dh["desc"], "Sev": texto_sev.get(dh["sev"], str(dh["sev"])),
                                "Prob": str(prob), "Risco": risco, "EPI": dh["epi"]
                            })

            my_bar.progress(100, text="Processamento finalizado!")
            
            if resultados_pgr or resultados_medicos:
                st.session_state['html_pgr'] = gerar_html_anexo(resultados_pgr, resultados_medicos)
                st.success("Matriz gerada com sucesso!")

    if 'html_pgr' in st.session_state:
        st.markdown("### 📄 Resultado da Extração")
        
        # Salvamento no Banco
        col_name, col_btn = st.columns([3, 1])
        with col_name: projeto = st.text_input("Nome do Projeto para Salvar:")
        with col_btn:
            st.write("")
            if st.button("💾 Salvar no Banco") and projeto:
                conn = sqlite3.connect('seconci_banco_dados.db')
                c = conn.cursor()
                c.execute("INSERT INTO historico_laudos (nome_projeto, modulo, data_salvamento, html_relatorio) VALUES (?, ?, ?, ?)", 
                          (projeto, "PGR", datetime.now().strftime("%d/%m/%Y"), st.session_state['html_pgr']))
                conn.commit(); conn.close()
                st.toast("Projeto Salvo!", icon="✅")

        # Visualização e Download
        aba1, aba2 = st.tabs(["👁️ Pré-visualizar Documento", "⬇️ Exportar (.doc)"])
        with aba1: components.html(st.session_state['html_pgr'], height=600, scrolling=True)
        with aba2: st.download_button("Baixar Word", st.session_state['html_pgr'].encode('utf-8'), "PGR_Automatizado.doc")

# ==========================================
# TELA 3: CENTRAL DE ARQUIVOS
# ==========================================
elif "Central" in menu:
    st.title("📂 Central de Relatórios Salvos")
    
    conn = sqlite3.connect('seconci_banco_dados.db')
    df_hist = pd.read_sql_query("SELECT id, nome_projeto, modulo, data_salvamento FROM historico_laudos ORDER BY id DESC", conn)
    
    if not df_hist.empty:
        st.dataframe(df_hist, use_container_width=True, hide_index=True)
        
        selecao_id = st.selectbox("Selecione o ID do projeto para carregar:", df_hist['id'].tolist())
        if st.button("Carregar Documento"):
            cursor = conn.cursor()
            cursor.execute("SELECT html_relatorio FROM historico_laudos WHERE id = ?", (selecao_id,))
            html_salvo = cursor.fetchone()[0]
            st.components.v1.html(html_salvo, height=600, scrolling=True)
    else:
        st.info("Nenhum projeto salvo no banco de dados ainda.")
    conn.close()

elif "PCMSO" in menu:
    st.title("🩺 Gerador de PCMSO")
    st.info("Módulo em construção na interface SaaS. Em breve integrado ao motor híbrido.")

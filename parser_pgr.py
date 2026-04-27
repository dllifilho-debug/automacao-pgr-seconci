import re
import json
import unicodedata
import io
from pathlib import Path


# ──────────────────────────────────────────────────────────────────────────────
# MAPEAMENTO: termos livres do PGR → chave do regras_pcmso.json
# ──────────────────────────────────────────────────────────────────────────────
TERMOS_PARA_CHAVE = {
    r"ru[ii]do": "RUIDO",
    r"vibra.{0,10}corpo.*inteiro": "VIBRACAO_CORPO_INTEIRO",
    r"trabalho em altura": "TRABALHO_EM_ALTURA_ESPACO_CONFINADO_MOTORISTA",
    r"espaco confinado": "ESPACO_CONFINADO",
    r"eletricidade|energia eletrica|choque eletrico": "PORTEIRO_ELETRICIDADE_ALTURA_MOTORISTA",
    r"risco psicossocial|psicossocial": "TRABALHO_EM_ALTURA_MAQUINAS_PESADAS_PSICOSSOCIAL",
    r"poeira.*madeira|madeira.*poeira": "POEIRA_PNOS_GESSO_MADEIRA_METALICA",
    r"poeira.*gesso|gesso.*poeira": "POEIRA_PNOS_GESSO_MADEIRA_METALICA",
    r"poeira.*metalica|fumos metalicos|fumo metalico": "POEIRA_PNOS_GESSO_MADEIRA_METALICA",
    r"pnos|fibra.*vidro|talco\b": "POEIRA_PNOS_GESSO_MADEIRA_METALICA",
    r"silica|quartzo|abesto|amianto": "POEIRA_MINERAL_SILICA_QUARTZO_OPERADOR_BETONEIRA",
    r"operador.*betoneira|betoneira": "POEIRA_MINERAL_SILICA_QUARTZO_OPERADOR_BETONEIRA",
    r"azulej": "POEIRA_MINERAL_SILICA_QUARTZO_OPERADOR_BETONEIRA",
    r"poeira.*mineral": "POEIRA_MINERAL_SILICA_QUARTZO_OPERADOR_BETONEIRA",
    r"tinta[s]?\b|nevoa|neblina": "NEVOAS_TINTAS_COLAS_IMPERMEABILIZACAO",
    r"cola[s]?\b|adesivo\b": "NEVOAS_TINTAS_COLAS_IMPERMEABILIZACAO",
    r"impermeabiliz": "NEVOAS_TINTAS_COLAS_IMPERMEABILIZACAO",
    r"cimento.*sem.*silica": "CONTATO_QUIMICOS_AGRESSORES_PULMONARES",
    r"mascara.*respiratoria.*epi": "USO_MASCARA_EPI_SEM_RISCO_QUIMICO",
    r"tricloroetileno|tricloroetano": "TRICLOROETILENO",
    r"benzeno\b": "BENZENO",
    r"tolueno\b": "TOLUENO",
    r"xileno\b": "XILENO",
    r"estireno\b": "ESTIRENO",
    r"fenol\b": "FENOL",
    r"monoxido.*carbono": "MONOXIDO_DE_CARBONO",
    r"manganes\b": "MANGANES",
    r"cromo.*hexavalente|cromo.*vi\b": "CROMO_HEXAVALENTE",
    r"fluoreto|acido fluoridrico": "FLUOR_ACIDO_FLUORIDRICO_FLUORETOS",
    r"metil.etil.cetona|mek\b": "METIL_ETIL_CETONA",
    r"acetona\b": "ACETONA",
    r"tetrahidrofurano|thf\b": "TETRAHIDROFURANO",
    r"cicloexanona|ciclohexanona": "CICLOEXANONA",
    r"policorte|corte.*plasma|plasma.*corte|solda\b": "POLICORTE_SOLDA",
    r"trabalhad.*saude|saude.*trabalhad": "TRABALHADORES_DA_SAUDE",
    r"manipul.*alimento|alimento.*manipul": "MANIPULAR_ALIMENTOS",
}

# Químicos que também são ototóxicos — gera SUBSTANCIA_OTOTOXICA automaticamente
OTOTOXICOS_COM_BIOLOGICO = {
    "TOLUENO", "XILENO", "ESTIRENO",
    "TRICLOROETILENO", "MONOXIDO_DE_CARBONO", "MANGANES",
}


# ──────────────────────────────────────────────────────────────────────────────
# EXTRAÇÃO DE TEXTO (pdfplumber → OCR automático)
# ──────────────────────────────────────────────────────────────────────────────
def _eh_pdf_bloqueado(paginas_texto: list, limite_chars_media: int = 30) -> bool:
    """Detecta PDF assinado/digitalizado pela baixa densidade de texto."""
    chars = sum(len(t.strip()) for t in paginas_texto)
    media = chars / max(len(paginas_texto), 1)
    return media < limite_chars_media


def _texto_via_pdfplumber(pdf_bytes: bytes) -> tuple:
    """Tenta extrair texto com pdfplumber. Retorna (texto, sucesso)."""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            paginas = [pg.extract_text() or "" for pg in pdf.pages]
        bloqueado = _eh_pdf_bloqueado(paginas)
        return "\n".join(paginas), not bloqueado
    except Exception:
        return "", False


def _texto_via_ocr(pdf_bytes: bytes, lang: str = "por") -> str:
    """Fallback OCR usando pdf2image + pytesseract."""
    try:
        import pytesseract
        from pdf2image import convert_from_bytes
        imagens = convert_from_bytes(pdf_bytes, dpi=200)
        return "\n".join(pytesseract.image_to_string(img, lang=lang) for img in imagens)
    except ImportError as e:
        raise RuntimeError(
            f"OCR indisponivel: {e}\n"
            "Instale: pip install pytesseract pdf2image\n"
            "Tesseract: https://github.com/UB-Mannheim/tesseract/wiki"
        )


def extrair_texto_pgr(pdf_bytes: bytes, forcar_ocr: bool = False) -> dict:
    """
    Extrai texto do PGR com detecção automática de PDF bloqueado.

    Retorna:
        {
            "texto": str,         # texto completo extraído
            "metodo": str,        # "pdfplumber" | "ocr"
            "bloqueado": bool,    # True se PDF estava bloqueado
            "aviso": str | None,  # mensagem de aviso para o usuário
        }
    """
    aviso = None

    if not forcar_ocr:
        texto, sucesso = _texto_via_pdfplumber(pdf_bytes)
        if sucesso and texto.strip():
            return {"texto": texto, "metodo": "pdfplumber", "bloqueado": False, "aviso": None}
        aviso = (
            "⚠️ PDF com texto bloqueado ou digitalizado detectado. "
            "Usando OCR automatico (pode ser mais lento)..."
        )
    else:
        aviso = "OCR forcado pelo usuario."

    texto = _texto_via_ocr(pdf_bytes)
    return {"texto": texto, "metodo": "ocr", "bloqueado": True, "aviso": aviso}


# ──────────────────────────────────────────────────────────────────────────────
# PARSING DE GHEs E RISCOS
# ──────────────────────────────────────────────────────────────────────────────
def _normalizar(texto: str) -> str:
    texto = texto.lower()
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def extrair_ghe_blocos(texto_pgr: str) -> dict:
    """Divide o texto do PGR em blocos por GHE."""
    padrao = re.compile(
        r"(GHE\s*\d+\s*[-:\u2013\u2014]+\s*[^\n]{3,80})",
        re.IGNORECASE
    )
    partes = padrao.split(texto_pgr)
    blocos = {}
    for i in range(1, len(partes), 2):
        nome = partes[i].strip()
        conteudo = partes[i + 1] if i + 1 < len(partes) else ""
        blocos[nome] = conteudo
    return blocos


def identificar_riscos(texto_bloco: str) -> list:
    """Identifica riscos pelo TERMOS_PARA_CHAVE e retorna lista de chaves."""
    texto_norm = _normalizar(texto_bloco)
    chaves = set()
    for padrao, chave in TERMOS_PARA_CHAVE.items():
        if re.search(padrao, texto_norm):
            chaves.add(chave)
            if chave in OTOTOXICOS_COM_BIOLOGICO:
                chaves.add("SUBSTANCIA_OTOTOXICA")
    return sorted(chaves)


def gerar_exames_por_riscos(lista_riscos: list, regras: dict) -> list:
    """Cruza lista de chaves de risco com o banco de regras e retorna exames."""
    vistos = set()
    lista = []
    for risco in lista_riscos:
        if risco not in regras:
            continue
        r = regras[risco]
        itens = r["exames"] if "exames" in r else [r]
        for item in itens:
            nome = item.get("exame") or r.get("exame", "")
            if nome and nome not in vistos:
                vistos.add(nome)
                lista.append({
                    "exame": nome,
                    "periodicidade_meses": item.get("periodicidade_meses", r.get("periodicidade_meses")),
                    "momentos": item.get("momentos", r.get("momentos", [])),
                    "origem_risco": risco,
                })
    return lista


# ──────────────────────────────────────────────────────────────────────────────
# FUNÇÃO PRINCIPAL
# ──────────────────────────────────────────────────────────────────────────────
def parsear_pgr(fonte, regras: dict, forcar_ocr: bool = False) -> dict:
    """
    Parseia um PGR e retorna GHEs com riscos e exames.

    Aceita:
      - bytes do PDF
      - caminho str/Path para o arquivo PDF
      - texto puro str (com newlines)

    Retorna:
        {
            "metodo_extracao": str,
            "bloqueado": bool,
            "aviso": str | None,
            "ghe_blocos": {nome_ghe: {"riscos_identificados": [...], "exames_gerados": [...]}}
        }
    """
    # Texto puro detectado pela presença de newlines
    if isinstance(fonte, str) and ("\n" in fonte or len(fonte) > 300):
        texto = fonte
        metodo, bloqueado, aviso = "texto_direto", False, None

    elif isinstance(fonte, (bytes, bytearray)):
        r = extrair_texto_pgr(bytes(fonte), forcar_ocr)
        texto, metodo, bloqueado, aviso = r["texto"], r["metodo"], r["bloqueado"], r["aviso"]

    elif isinstance(fonte, (str, Path)):
        with open(fonte, "rb") as f:
            pdf_bytes = f.read()
        r = extrair_texto_pgr(pdf_bytes, forcar_ocr)
        texto, metodo, bloqueado, aviso = r["texto"], r["metodo"], r["bloqueado"], r["aviso"]

    else:
        raise ValueError("fonte deve ser bytes, caminho de arquivo ou texto string.")

    blocos = extrair_ghe_blocos(texto)
    ghe_resultado = {}
    for ghe, conteudo in blocos.items():
        riscos = identificar_riscos(conteudo)
        exames = gerar_exames_por_riscos(riscos, regras)
        ghe_resultado[ghe] = {"riscos_identificados": riscos, "exames_gerados": exames}

    return {
        "metodo_extracao": metodo,
        "bloqueado": bloqueado,
        "aviso": aviso,
        "ghe_blocos": ghe_resultado,
    }

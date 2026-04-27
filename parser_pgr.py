import re
import io
import unicodedata
from pathlib import Path


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

OTOTOXICOS_COM_BIOLOGICO = {
    "TOLUENO", "XILENO", "ESTIRENO",
    "TRICLOROETILENO", "MONOXIDO_DE_CARBONO", "MANGANES",
}

# Codigos eSocial → chave interna (PGRs gerados pelo SistemaEso e similares)
ESOCIAL_PARA_CHAVE = {
    "02.01.001": "RUIDO",
    "02.01.002": "RUIDO",
    "02.03.001": "VIBRACAO_CORPO_INTEIRO",
    "02.03.002": "VIBRACAO_CORPO_INTEIRO",
    "01.18.001": "POEIRA_MINERAL_SILICA_QUARTZO_OPERADOR_BETONEIRA",
    "01.02.001": "NEVOAS_TINTAS_COLAS_IMPERMEABILIZACAO",
    "01.06.001": "BENZENO",
    "01.06.002": "TOLUENO",
    "01.06.003": "XILENO",
    "01.06.004": "ESTIRENO",
    "01.06.011": "MONOXIDO_DE_CARBONO",
    "01.06.019": "MANGANES",
    "01.06.020": "CROMO_HEXAVALENTE",
    "01.06.031": "TRICLOROETILENO",
    "01.06.040": "FENOL",
    "01.06.048": "FLUOR_ACIDO_FLUORIDRICO_FLUORETOS",
}


# ──────────────────────────────────────────────────────────────────────────────
# EXTRAÇÃO DE TEXTO
# ──────────────────────────────────────────────────────────────────────────────
def _eh_pdf_bloqueado(paginas_texto: list, limite_chars_media: int = 30) -> bool:
    chars = sum(len(t.strip()) for t in paginas_texto)
    return (chars / max(len(paginas_texto), 1)) < limite_chars_media


def _texto_via_pdfplumber(pdf_bytes: bytes) -> tuple:
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            paginas = [pg.extract_text() or "" for pg in pdf.pages]
        bloqueado = _eh_pdf_bloqueado(paginas)
        return "\n".join(paginas), not bloqueado
    except Exception:
        return "", False


def _texto_via_ocr(pdf_bytes: bytes, lang: str = "por") -> str:
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
    Extrai texto do PGR com detecção automatica de PDF bloqueado.
    Retorna: {texto, metodo, bloqueado, aviso}
    """
    if not forcar_ocr:
        texto, sucesso = _texto_via_pdfplumber(pdf_bytes)
        if sucesso and texto.strip():
            return {"texto": texto, "metodo": "pdfplumber", "bloqueado": False, "aviso": None}
        aviso = (
            "PDF com texto bloqueado ou digitalizado detectado. "
            "Usando OCR automatico (pode ser mais lento)..."
        )
    else:
        aviso = "OCR forcado pelo usuario."
    texto = _texto_via_ocr(pdf_bytes)
    return {"texto": texto, "metodo": "ocr", "bloqueado": True, "aviso": aviso}


# ──────────────────────────────────────────────────────────────────────────────
# DETECCAO DE FORMATO
# ──────────────────────────────────────────────────────────────────────────────
def detectar_formato(texto: str) -> str:
    """Retorna 'GHE' ou 'CARGO' conforme o padrao dominante no documento."""
    n_ghe   = len(re.findall(r"\bGHE\s*\d+", texto, re.IGNORECASE))
    n_cargo = len(re.findall(r"\bCARGO\s+[A-Z]{2}", texto, re.IGNORECASE))
    return "CARGO" if n_cargo >= n_ghe else "GHE"


# ──────────────────────────────────────────────────────────────────────────────
# NORMALIZACAO
# ──────────────────────────────────────────────────────────────────────────────
def _normalizar(texto: str) -> str:
    nfkd = unicodedata.normalize("NFKD", texto.lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# ──────────────────────────────────────────────────────────────────────────────
# EXTRATORES DE BLOCOS
# ──────────────────────────────────────────────────────────────────────────────
def extrair_blocos_ghe(texto: str) -> dict:
    """Divide o texto em blocos iniciados por GHE NN - Nome."""
    padrao = re.compile(
        r"(GHE\s*\d+\s*[-:\u2013\u2014]+\s*[^\n]{3,80})",
        re.IGNORECASE
    )
    partes = padrao.split(texto)
    blocos = {}
    for i in range(1, len(partes), 2):
        nome = partes[i].strip()
        conteudo = partes[i + 1] if i + 1 < len(partes) else ""
        blocos[nome] = conteudo
    return blocos


def extrair_blocos_cargo(texto: str) -> dict:
    """
    Divide o texto em blocos iniciados por:
      CARGO <NOME> - CBO: XXXXXX
    ou simplesmente CARGO <NOME> (sem CBO).
    Captura o nome completo ate o fim da linha.
    """
    padrao = re.compile(
        r"(CARGO[ \t]+[A-Z\xC0-\xFF][A-Z\xC0-\xFF \t\/\-]*?(?:[ \t]*-[ \t]*CBO[: \t]*\d{6})?)[ \t]*\n",
        re.IGNORECASE
    )
    partes = padrao.split(texto)
    blocos = {}
    for i in range(1, len(partes), 2):
        nome = re.sub(r"\s+", " ", partes[i]).strip()
        conteudo = partes[i + 1] if i + 1 < len(partes) else ""
        blocos[nome] = conteudo
    return blocos


# ──────────────────────────────────────────────────────────────────────────────
# IDENTIFICACAO DE RISCOS
# ──────────────────────────────────────────────────────────────────────────────
def identificar_riscos(texto_bloco: str) -> list:
    texto_norm = _normalizar(texto_bloco)
    chaves = set()

    # 1. Codigos eSocial explícitos
    for cod, chave in ESOCIAL_PARA_CHAVE.items():
        if cod in texto_bloco:
            chaves.add(chave)
            if chave in OTOTOXICOS_COM_BIOLOGICO:
                chaves.add("SUBSTANCIA_OTOTOXICA")

    # 2. Termos livres (narrativo)
    for padrao, chave in TERMOS_PARA_CHAVE.items():
        if re.search(padrao, texto_norm):
            chaves.add(chave)
            if chave in OTOTOXICOS_COM_BIOLOGICO:
                chaves.add("SUBSTANCIA_OTOTOXICA")

    return sorted(chaves)


# ──────────────────────────────────────────────────────────────────────────────
# GERACAO DE EXAMES
# ──────────────────────────────────────────────────────────────────────────────
def gerar_exames_por_riscos(lista_riscos: list, regras: dict) -> list:
    vistos = set()
    lista = []
    for risco in lista_riscos:
        if risco not in regras:
            continue
        r = regras[risco]
        itens = r.get("exames", [r])
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
# FUNCAO PRINCIPAL
# ──────────────────────────────────────────────────────────────────────────────
def parsear_pgr(fonte, regras: dict, forcar_ocr: bool = False) -> dict:
    """
    Parseia um PGR e retorna blocos com riscos e exames.

    Aceita:
      - bytes do PDF
      - caminho str/Path para o arquivo
      - texto puro str

    Detecta automaticamente o formato: GHE ou CARGO/CBO.

    Retorna:
        {
            "metodo_extracao": str,
            "bloqueado":       bool,
            "aviso":           str | None,
            "formato":         str,   # "GHE" | "CARGO"
            "ghe_blocos":      dict,  # chave = nome do GHE ou CARGO
        }
    """
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

    formato = detectar_formato(texto)
    blocos  = extrair_blocos_cargo(texto) if formato == "CARGO" else extrair_blocos_ghe(texto)

    ghe_resultado = {}
    for nome, conteudo in blocos.items():
        riscos = identificar_riscos(conteudo)
        exames = gerar_exames_por_riscos(riscos, regras)
        ghe_resultado[nome] = {"riscos_identificados": riscos, "exames_gerados": exames}

    return {
        "metodo_extracao": metodo,
        "bloqueado":       bloqueado,
        "aviso":           aviso,
        "formato":         formato,
        "ghe_blocos":      ghe_resultado,
    }

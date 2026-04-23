"""
data/matriz_exames.py
Banco de dados de exames validado pela Dra. Patrícia — versão 06/2025
Gerado automaticamente a partir da planilha:
  Matriz-funcao-risco-exames-validado-Dra.-Patricia-06.2025-1.xlsx
"""

MATRIZ_RISCO_EXAME = {
    # ── FÍSICOS ───────────────────────────────────────────────────────────────
    "RUIDO": {
        "exame": "Audiometria",
        "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 Anexo I — Ruído. Demissional se último exame > 120 dias.",
    },
    "VIBRACAO CORPO INTEIRO": {
        "exame": "RX de coluna lombo-sacra",
        "adm": True, "per": None, "mro": True, "rt": False, "dem": False,
        "obs": "Vibração de corpo inteiro — avaliação radiológica admissional.",
    },

    # ── QUÍMICOS — Benzeno e derivados ───────────────────────────────────────
    "BENZENO": {
        "exame": "Ácido trans-trans mucônico",
        "adm": True, "per": "6 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 Anexo II — Benzeno. Inclui hemograma e reticulócitos.",
    },
    "TOLUENO": {
        "exame": "Ortocresol na urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Tolueno (solvente orgânico).",
    },
    "XILENO": {
        "exame": "Ác. Metil-hipúrico na urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Xileno.",
    },
    "ACETONA": {
        "exame": "Acetona na urina",
        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "NR-7 — Acetona / 2-Propanol.",
    },
    "METIL-ETIL-CETONA": {
        "exame": "Metil-Etil-Cetona",
        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "NR-7 — Metil-Etil-Cetona.",
    },
    "TETRAHIDROFURANO": {
        "exame": "Tetrahidrofurnano na urina",
        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "NR-7 — Tetrahidrofurano.",
    },
    "DICLOROMETANO": {
        "exame": "Diclorometano na urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Diclorometano.",
    },
    "TRICLOROETILENO": {
        "exame": "Ácido tricloroacético na urina",
        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "NR-7 — Tricloroetileno.",
    },
    "ESTIRENO": {
        "exame": "Soma dos Ácidos Mandélico e Fenilglioxílico na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Estireno.",
    },
    "N-HEXANO": {
        "exame": "2,5 Hexanodiona na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — N-Hexano.",
    },
    "FENOL": {
        "exame": "Fenol na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Fenol.",
    },
    "MERCURIO": {
        "exame": "Mercúrio na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Mercúrio metálico.",
    },
    "METANOL": {
        "exame": "Metanol na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Metanol.",
    },
    "CICLOHEXANONA": {
        "exame": "Ciclohexanol na urina",
        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "NR-7 — Ciclohexanona.",
    },

    # ── QUÍMICOS — Metais ─────────────────────────────────────────────────────
    "CHUMBO": {
        "exame": "Chumbo no Sangue + ALA-U",
        "adm": True, "per": "6 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 Anexo II — Chumbo.",
    },
    "MANGANES": {
        "exame": "Manganês sanguíneo",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Eletrodo de solda que libera manganês.",
    },
    "CROMO": {
        "exame": "Cromo na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Cromo hexavalente.",
    },
    "CADMIO": {
        "exame": "Cádmio na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 Anexo II — Cádmio.",
    },
    "ARSENICO": {
        "exame": "Arsênio Inorgânico + Metabólitos Metilados na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Arsênico elementar.",
    },
    "COBALTO": {
        "exame": "Cobalto na Urina",
        "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Cobalto e compostos inorgânicos.",
    },
    "FLUOR": {
        "exame": "Fluoreto Urinário",
        "adm": True, "per": "6 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 — Flúor e fluoretos.",
    },

    # ── FÍSICO-QUÍMICO — Fumos / Solda / Combustão ───────────────────────────
    "SOLDA": {
        "exame": "Carboxiemoglobina",
        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "Policorte / Solda — CO liberado.",
    },
    "MONOXIDO DE CARBONO": {
        "exame": "Carboxiemoglobina",
        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "Monóxido de Carbono.",
    },
    "POLICORTE": {
        "exame": "Carboxiemoglobina",
        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "Policorte — CO liberado.",
    },
    "COMBUSTIVEL": {
        "exame": "Hemograma",
        "adm": True, "per": "6 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "Exposto a combustíveis / benzeno.",
    },

    # ── FÍSICO — Poeiras / Fibras / Pulmão ───────────────────────────────────
    "SILICA": {
        "exame": "RX de Tórax OIT",
        "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "Sílica / Quartzo / Poeira mineral — NR-7.",
    },
    "POEIRA MINERAL": {
        "exame": "RX de Tórax OIT",
        "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "Poeira mineral / sílica / carvão.",
    },
    "CIMENTO": {
        "exame": "RX de Tórax OIT",
        "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "Cimento / betoneira.",
    },
    "ASBESTO": {
        "exame": "RX de Tórax OIT",
        "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "Asbesto / Amianto.",
    },
    "FUMOS METALICOS": {
        "exame": "RX de Tórax OIT",
        "adm": True, "per": "60 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "PNOS — Fumos metálicos / poeira metálica.",
    },
    "MADEIRA": {
        "exame": "RX de Tórax OIT",
        "adm": True, "per": "60 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "PNOS — Poeira de madeira.",
    },
    "TINTA": {
        "exame": "Espirometria",
        "adm": True, "per": "24 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "Névoas/neblinas/tintas/colas — agressor pulmonar.",
    },
    "IMPERMEABILIZACAO": {
        "exame": "Espirometria",
        "adm": True, "per": "24 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "Impermeabilização — agressor pulmonar.",
    },
    "MASCARA RESPIRATORIA": {
        "exame": "Espirometria",
        "adm": True, "per": "24 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "Uso de máscara de proteção respiratória.",
    },

    # ── RISCO DE ACIDENTE — Altura / Confinado / Eletricidade ────────────────
    "QUEDA DE ALTURA": {
        "exame": "Avaliação Psicossocial",
        "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Trabalho em altura — NR-35.",
    },
    "ESPACO CONFINADO": {
        "exame": "Avaliação Psicossocial",
        "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Espaço confinado — NR-33.",
    },
    "RISCO ELETRICO": {
        "exame": "Acuidade Visual",
        "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Eletricidade — NR-10.",
    },

    # ── BIOLÓGICO ─────────────────────────────────────────────────────────────
    "AGENTE BIOLOGICO": {
        "exame": "Anti-HBs + HBsAg",
        "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Trabalhadores da saúde / risco biológico.",
    },
    "ESGOTO": {
        "exame": "EPF (Coproparasitológico) + Anti-HBs",
        "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Contato com esgoto / efluentes.",
    },

    # ── MOTORISTA / OPERADOR DE MÁQUINAS PESADAS ──────────────────────────────
    "MOTORISTA": {
        "exame": "Acuidade Visual",
        "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Motorista veículos leves.",
    },
}

MATRIZ_FUNCAO_EXAME = {
    "SOLDADOR": [
        {"exame": "Manganês sanguíneo", "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False, "obs": "Eletrodo libera manganês"},
        {"exame": "Carboxiemoglobina",  "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "CO — policorte/solda"},
        {"exame": "Acuidade Visual",    "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Solda — NR-7"},
    ],
    "SERRALHEIRO": [
        {"exame": "Manganês sanguíneo", "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False, "obs": "Solda — eletrodo manganês"},
        {"exame": "Carboxiemoglobina",  "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "CO — solda/policorte"},
    ],
    "PINTOR": [
        {"exame": "Ácido trans-trans mucônico", "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "Tinta — benzeno/solventes"},
        {"exame": "Contagem de Reticulócitos",  "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False, "obs": "Exposto benzeno/combustíveis"},
        {"exame": "Ortocresol na urina",        "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "Tolueno — tintas"},
        {"exame": "Avaliação Psicossocial",     "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Risco psicossocial — altura/supervisão"},
    ],
    "IMPERMEABILIZADOR": [
        {"exame": "Ácido trans-trans mucônico", "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "Benzeno — impermeabilizante"},
        {"exame": "Contagem de Reticulócitos",  "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False, "obs": "Benzeno — impermeabilização"},
        {"exame": "Carboxiemoglobina",          "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "CO — impermeabilização"},
        {"exame": "Avaliação Psicossocial",     "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Risco altura/confinado"},
    ],
    "ENCARREGADO": [
        {"exame": "EPF (Coproparasitológico) + Anti-HBs", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "GHE 13 — supervisão biológico"},
        {"exame": "Avaliação Psicossocial",               "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Supervisor de equipes — NR-35"},
    ],
    "MESTRE DE OBRA": [
        {"exame": "EPF (Coproparasitológico) + Anti-HBs", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "GHE 13 — supervisão"},
        {"exame": "Avaliação Psicossocial",               "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Supervisor — NR-35"},
    ],
    "ELETRICISTA": [
        {"exame": "Avaliação Psicossocial", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Eletricista — risco altura/eletricidade"},
    ],
    "ELETRICISTA INDUSTRIAL": [
        {"exame": "Avaliação Psicossocial",   "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Eletricista industrial — NR-35"},
        {"exame": "RX de coluna lombo-sacra", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False, "obs": "Carga/vibração — eletricista industrial"},
    ],
    "OPERADOR DE CREMALHEIRA": [
        {"exame": "RX de coluna lombo-sacra", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False, "obs": "Vibração — operador cremalheira"},
    ],
    "SINALEIRO": [
        {"exame": "Avaliação Psicossocial", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Sinaleiro grua — NR-35"},
    ],
    "TECNICO DE SEGURANCA DO TRABALHO": [
        {"exame": "Avaliação Psicossocial", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Técnico SST — circula em toda a obra"},
    ],
    "ESTAGIARIO DE SEGURANCA DO TRABALHO": [
        {"exame": "Avaliação Psicossocial", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Estagiário SST — circula em obra"},
    ],
    "ALMOXARIFE": [
        {"exame": "Avaliação Psicossocial", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Almoxarife — área de obra"},
    ],
    "ENGENHEIRO": [
        {"exame": "Avaliação Psicossocial", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Engenheiro — fiscalização em obra/altura"},
    ],
    "MECANICO DE MANUTENCAO": [
        {"exame": "Avaliação Psicossocial",   "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Manutenção — risco altura/confinado"},
        {"exame": "RX de coluna lombo-sacra", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False, "obs": "Vibração — manutenção"},
    ],
    "MONTADOR": [
        {"exame": "Avaliação Psicossocial",   "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Montador — risco altura"},
        {"exame": "RX de coluna lombo-sacra", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False, "obs": "Vibração — montador"},
    ],
    "MOTORISTA": [
        {"exame": "Audiometria",       "adm": True, "per": "12 MESES", "mro": True, "rt": True,  "dem": True,  "obs": "Motorista — risco de batida"},
        {"exame": "Acuidade Visual",   "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista"},
        {"exame": "Hemograma",         "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista/op. máq. pesadas"},
        {"exame": "Glicemia em Jejum", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista/op. máq. pesadas"},
        {"exame": "ECG",               "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista/op. máq. pesadas"},
    ],
    "ENCANADOR": [
        {"exame": "Acetona na urina",           "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "Cetonaster — risco solvente encanador"},
        {"exame": "Metil-Etil-Cetona",          "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "Cetonaster — risco solvente encanador"},
        {"exame": "Ciclohexanol na urina",      "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "Cetonaster — risco solvente encanador"},
        {"exame": "Tetrahidrofurnano na urina", "adm": False, "per": "6 MESES", "mro": False, "rt": False, "dem": False, "obs": "Cetonaster — risco solvente encanador"},
    ],
}

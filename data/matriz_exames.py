"""
data/matriz_exames.py
Banco de dados de exames validado pela Dra. Patrícia — versão 06/2025
Gerado automaticamente a partir da planilha:
  Matriz-funcao-risco-exames-validado-Dra.-Patricia-06.2025-1.xlsx

Estrutura MATRIZ_RISCO_EXAME:
  chave      → palavra-chave normalizada (uppercase sem acento) buscada no texto do PGR
  exame      → nome oficial do exame a solicitar
  adm        → realizar no admissional
  per        → periodicidade ("6 MESES", "12 MESES", "24 MESES", "60 MESES", None)
  mro        → realizar no periódico (MRO = Monitoramento Regular Ocupacional)
  rt         → realizar no retorno ao trabalho
  dem        → realizar no demissional
  obs        → observação / fundamentação legal
"""

# ─────────────────────────────────────────────────────────────────────────────
# MATRIZ POR RISCO (chave no texto do PGR → exame(s))
# Fonte: Planilha validada Dra. Patrícia | NR-7 Anexos I e II | NR-33 | NR-35
# ─────────────────────────────────────────────────────────────────────────────

MATRIZ_RISCO_EXAME = {

    # ── FÍSICOS ───────────────────────────────────────────────────────────────
    "RUIDO": {
        "exame": "Audiometria Tonal (PTA)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 Anexo I — Ruído. Demissional se último exame > 120 dias.",
    },
    "VIBRACAO CORPO INTEIRO": {
        "exame": "Raio-X Coluna Lombo-Sacra",
        "adm": True, "periodico": None, "mro": False, "rt": False, "dem": False,
        "obs": "Vibração de corpo inteiro — avaliação radiológica admissional.",
    },
    "VIBRACAO": {
        "exame": "Avaliação Psicossocial (NR-35)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Vibração — risco psicossocial associado ao trabalho em altura/maquinário.",
    },

    # ── QUÍMICOS — Benzeno e derivados ───────────────────────────────────────
    "BENZENO": {
        "exame": "Ácido Trans-Trans Mucônico na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 Anexo II — Benzeno. Inclui hemograma e reticulócitos.",
    },
    "TOLUENO": {
        "exame": "Ortocresol na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Tolueno (solvente orgânico).",
    },
    "XILENO": {
        "exame": "Ác. Metil-Hipúrico na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Xileno.",
    },
    "ACETONA": {
        "exame": "Acetona na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Acetona / 2-Propanol.",
    },
    "METIL-ETIL-CETONA": {
        "exame": "Metil-Etil-Cetona (MEK) na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Metil-Etil-Cetona.",
    },
    "TETRAHIDROFURANO": {
        "exame": "Tetrahidrofurnano na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Tetrahidrofurano.",
    },
    "DICLOROMETANO": {
        "exame": "Diclorometano na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Diclorometano.",
    },
    "TRICLOROETILENO": {
        "exame": "Ác. Tricloroacético na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Tricloroetileno.",
    },
    "ESTIRENO": {
        "exame": "Soma dos Ácidos Mandélico e Fenilglioxílico na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Estireno.",
    },
    "N-HEXANO": {
        "exame": "2,5 Hexanodiona na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — N-Hexano.",
    },
    "FENOL": {
        "exame": "Fenol na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Fenol.",
    },
    "MERCURIO": {
        "exame": "Mercúrio na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Mercúrio metálico.",
    },
    "METANOL": {
        "exame": "Metanol na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Metanol.",
    },
    "CICLOHEXANONA": {
        "exame": "Ciclohexanol (H) na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Ciclohexanona.",
    },

    # ── QUÍMICOS — Metais ─────────────────────────────────────────────────────
    "CHUMBO": {
        "exame": "Chumbo no Sangue + ALA-U",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 Anexo II — Chumbo. Reaproveitável no DEM se < 6 meses.",
    },
    "MANGANES": {
        "exame": "Manganês no Sangue",
        "adm": True, "periodico": "6 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "Eletrodo de solda que libera manganês.",
    },
    "CROMO": {
        "exame": "Cromo na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Cromo hexavalente (compostos solúveis).",
    },
    "CADMIO": {
        "exame": "Cádmio na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 Anexo II — Cádmio. Reaproveitável no DEM se < 6 meses.",
    },
    "ARSENICO": {
        "exame": "Arsênio Inorgânico + Metabólitos Metilados na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Arsênico elementar e compostos inorgânicos.",
    },
    "COBALTO": {
        "exame": "Cobalto na Urina",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "NR-7 — Cobalto e compostos inorgânicos.",
    },
    "FLUOR": {
        "exame": "Fluoreto Urinário",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "NR-7 — Flúor e fluoretos inorgânicos.",
    },

    # ── FÍSICO-QUÍMICO — Fumos / Solda / Combustão ───────────────────────────
    "SOLDA": {
        "exame": "Carboxihemoglobina no Sangue",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Policorte / Solda — CO liberado.",
    },
    "MONOXIDO DE CARBONO": {
        "exame": "Carboxihemoglobina no Sangue",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Monóxido de Carbono.",
    },
    "POLICORTE": {
        "exame": "Carboxihemoglobina no Sangue",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Policorte — CO liberado.",
    },
    "COMBUSTIVEL": {
        "exame": "Hemograma Completo",
        "adm": True, "periodico": "6 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "Exposto a combustíveis / benzeno / radiação ionizante.",
    },

    # ── FÍSICO — Poeiras / Fibras / Pulmão ───────────────────────────────────
    "SILICA": {
        "exame": "Raio-X de Tórax OIT",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "Sílica / Quartzo / Poeira mineral — NR-7.",
    },
    "POEIRA MINERAL": {
        "exame": "Raio-X de Tórax OIT",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "Poeira mineral / sílica / carvão.",
    },
    "CIMENTO": {
        "exame": "Raio-X de Tórax OIT",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "Cimento / betoneira / azulejista — NR-7.",
    },
    "ASBESTO": {
        "exame": "Raio-X de Tórax OIT",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "Asbesto / Amianto.",
    },
    "FUMOS METALICOS": {
        "exame": "Raio-X de Tórax OIT",
        "adm": True, "periodico": "60 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "PNOS — Fumos metálicos / poeira metálica.",
    },
    "MADEIRA": {
        "exame": "Raio-X de Tórax OIT",
        "adm": True, "periodico": "60 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "PNOS — Poeira de madeira.",
    },
    "TINTA": {
        "exame": "Espirometria (somente)",
        "adm": True, "periodico": "24 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "Névoas/neblinas/tintas/colas — agressor pulmonar.",
    },
    "IMPERMEABILIZACAO": {
        "exame": "Espirometria (somente)",
        "adm": True, "periodico": "24 MESES", "mro": True, "rt": True, "dem": False,
        "obs": "Impermeabilização — agressor pulmonar.",
    },
    "MASCARA RESPIRATORIA": {
        "exame": "Espirometria",
        "adm": True, "periodico": "24 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Uso de máscara de proteção respiratória sem risco químico específico no PGR.",
    },

    # ── RISCO DE ACIDENTE — Altura / Confinado / Eletricidade ────────────────
    "QUEDA DE ALTURA": {
        "exame": "Avaliação Psicossocial (NR-35)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Trabalho em altura — NR-35 obrigatória.",
    },
    "ESPACO CONFINADO": {
        "exame": "Avaliação Psicossocial (NR-35)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Espaço confinado — NR-33 obrigatória.",
    },
    "RISCO ELETRICO": {
        "exame": "Acuidade Visual (Avaliação Oftalmológica)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Eletricidade — NR-10.",
    },

    # ── BIOLÓGICO ─────────────────────────────────────────────────────────────
    "AGENTE BIOLOGICO": {
        "exame": "Anti-HBs + HBsAg",
        "adm": True, "periodico": "24 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Trabalhadores da saúde / risco biológico.",
    },
    "ESGOTO": {
        "exame": "EPF (Coproparasitológico) + Anti-HBs",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Contato com esgoto / efluentes.",
    },

    # ── MOTORISTA / OPERADOR DE MÁQUINAS PESADAS ──────────────────────────────
    "MOTORISTA": {
        "exame": "Acuidade Visual (Avaliação Oftalmológica)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Motorista veículos leves.",
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# MATRIZ POR FUNÇÃO (cargo específico → exames adicionais)
# Fonte: Aba "Exames x função" — Planilha Dra. Patrícia
# ─────────────────────────────────────────────────────────────────────────────

MATRIZ_FUNCAO_EXAME = {

    # Toda função de canteiro recebe pacote base (definido em modulo_pcmso.py)
    # Aqui ficam apenas exames EXTRAS por função específica

    "SOLDADOR": [
        {"exame": "Manganês no Sangue",           "adm": True, "per": "6 MESES",  "mro": False, "rt": False, "dem": False, "obs": "Eletrodo libera manganês"},
        {"exame": "Carboxihemoglobina no Sangue", "adm": True, "per": "6 MESES",  "mro": True,  "rt": False, "dem": False, "obs": "CO — policorte/solda"},
        {"exame": "Acuidade Visual (Avaliação Oftalmológica)", "adm": True, "per": "12 MESES", "mro": True, "rt": True, "dem": False, "obs": "Solda — NR-7"},
    ],
    "SERRALHEIRO": [
        {"exame": "Manganês no Sangue",           "adm": True, "per": "6 MESES",  "mro": False, "rt": False, "dem": False, "obs": "Solda — eletrodo manganês"},
        {"exame": "Carboxihemoglobina no Sangue", "adm": True, "per": "6 MESES",  "mro": True,  "rt": False, "dem": False, "obs": "CO — solda/policorte"},
    ],
    "PINTOR": [
        {"exame": "Ácido Trans-Trans Mucônico na Urina", "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False, "obs": "Tinta — benzeno/solventes"},
        {"exame": "Contagem de Reticulócitos",    "adm": True, "per": "6 MESES",  "mro": True,  "rt": False, "dem": False, "obs": "Exposto benzeno/combustíveis"},
        {"exame": "Ortocresol na Urina",          "adm": True, "per": "6 MESES",  "mro": True,  "rt": False, "dem": False, "obs": "Tolueno — tintas"},
        {"exame": "EPF (Coproparasitológico) + Anti-HBs", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "GHE 13 Supervisão — biológico"},
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Risco psicossocial — altura/supervisão"},
    ],
    "IMPERMEABILIZADOR": [
        {"exame": "Ácido Trans-Trans Mucônico na Urina", "adm": True, "per": "6 MESES", "mro": True, "rt": False, "dem": False, "obs": "Benzeno — impermeabilizante"},
        {"exame": "Contagem de Reticulócitos",    "adm": True, "per": "6 MESES",  "mro": True,  "rt": False, "dem": False, "obs": "Benzeno — impermeabilização"},
        {"exame": "Carboxihemoglobina no Sangue", "adm": True, "per": "6 MESES",  "mro": True,  "rt": False, "dem": False, "obs": "CO — impermeabilização"},
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Risco altura/confinado"},
    ],
    "ENCARREGADO": [
        {"exame": "EPF (Coproparasitológico) + Anti-HBs", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "GHE 13 — supervisão biológico"},
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Supervisor de equipes — NR-35"},
    ],
    "MESTRE DE OBRA": [
        {"exame": "EPF (Coproparasitológico) + Anti-HBs", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "GHE 13 — supervisão"},
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Supervisor — NR-35"},
    ],
    "ELETRICISTA": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Eletricista — risco altura/eletricidade"},
        {"exame": "EPF (Coproparasitológico) + Anti-HBs", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "GHE 13 — supervisão"},
    ],
    "ELETRICISTA INDUSTRIAL": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Eletricista industrial — NR-35"},
        {"exame": "Raio-X Coluna Lombo-Sacra", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False, "obs": "Carga/vibração — eletricista industrial"},
    ],
    "OPERADOR DE CREMALHEIRA": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Op. cremalheira — risco altura"},
        {"exame": "Raio-X Coluna Lombo-Sacra", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False, "obs": "Vibração — operador cremalheira"},
    ],
    "SINALEIRO": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Sinaleiro grua — NR-35"},
    ],
    "TECNICO DE SEGURANCA DO TRABALHO": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Técnico SST — circula em toda a obra"},
    ],
    "ESTAGIARIO DE SEGURANCA DO TRABALHO": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Estagiário SST — circula em obra"},
    ],
    "ALMOXARIFE": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Almoxarife — área de obra"},
    ],
    "ENGENHEIRO": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Engenheiro — fiscalização em obra/altura"},
    ],
    "MECANICO DE MANUTENCAO": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Manutenção — risco altura/confinado"},
        {"exame": "Raio-X Coluna Lombo-Sacra", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False, "obs": "Vibração — manutenção"},
    ],
    "MONTADOR": [
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Montador — risco altura"},
        {"exame": "Raio-X Coluna Lombo-Sacra", "adm": True, "per": "24 MESES", "mro": True, "rt": False, "dem": False, "obs": "Vibração — montador"},
    ],
    "MOTORISTA": [
        {"exame": "Audiometria Tonal (PTA)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista — risco de batida"},
        {"exame": "Acuidade Visual (Avaliação Oftalmológica)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista"},
        {"exame": "Hemograma Completo", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista/op. máq. pesadas"},
        {"exame": "Glicemia de Jejum", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista/op. máq. pesadas"},
        {"exame": "Eletrocardiograma (ECG)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Motorista/op. máq. pesadas"},
    ],
    "ENCANADOR": [
        {"exame": "EPF (Coproparasitológico) + Anti-HBs", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Encanador — risco biológico/esgoto"},
        {"exame": "Avaliação Psicossocial (NR-35)", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": "Risco altura/confinado"},
    ],
}

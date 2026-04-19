"""
data/matriz_exames.py
Matriz Função x Exames e Risco x Exames — validada Dra. Patrícia Montalvo (06/2025)

Estrutura de cada exame:
  {
    "exame":       str,   # nome completo do exame
    "adm":         bool,  # realizado na admissão
    "per":         str,   # periodicidade (ex: "12 MESES", "06 MESES", "24 MESES", "" = não realizado)
    "mro":         bool,  # mudança de risco ocupacional
    "rt":          bool,  # retorno ao trabalho
    "dem":         bool,  # demissional
    "obs":         str,   # observação (opcional)
  }
"""

# ── MATRIZ FUNÇÃO → EXAMES ───────────────────────────────────────────────────
MATRIZ_FUNCAO_EXAME = {

    # ── Administrativos puros ─────────────────────────────────────────────
    "ADMINISTRATIVO": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "ASSISTENTE": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "AUXILIAR": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "ANALISTA": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "COORDENADOR": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "GESTOR": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "GERENTE": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                  "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "DIRETOR": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                  "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "RECEPCIONISTA": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "TELEFONISTA": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "COMPRADOR": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "ESTAGIARIO": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "JOVEM APRENDIZ": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": "Incluir no PCMSO somente se >= 18 anos"},
    ],
    "DESIGNER": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],

    # ── Engenharia / Técnicos ─────────────────────────────────────────────
    "ENGENHEIRO": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Dem. obrigatorio se ultima > 120 dias"},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "TECNICO": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "SUPERVISOR": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "ENCARREGADO": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],

    # ── Trabalho em Altura / Espaço Confinado / Motorista ────────────────
    "TRABALHO EM ALTURA": [
        {"exame": "Avaliacao Psicossocial (NR-35)",            "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": "NR-35"},
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "ESPACO CONFINADO": [
        {"exame": "Avaliacao Psicossocial (NR-33)",            "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": "NR-33"},
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "MOTORISTA": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "OPERADOR DE MAQUINAS": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],

    # ── Porteiro / Eletricidade ───────────────────────────────────────────
    "PORTEIRO": [
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "ELETRICISTA": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": "NR-10"},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],

    # ── Obra / Construção ─────────────────────────────────────────────────
    "PEDREIRO": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Espirometria",                              "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Tolueno 2,6-diisocianato no PGR"},
        {"exame": "Raio-X de Torax OIT",                       "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Tolueno 2,6-diisocianato no PGR"},
    ],
    "SERVENTE": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Espirometria",                              "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Raio-X de Torax OIT",                       "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
    ],
    "CARPINTEIRO": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Raio-X de Torax OIT",                       "adm": True,  "per": "60 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Poeira de madeira"},
    ],
    "PINTOR": [
        {"exame": "Exame Clinico (Anamnese / Exame Fisico)",   "adm": True,  "per": "06 MESES", "mro": True,  "rt": True,  "dem": True,  "obs": "Periodicidade semestral pelo risco quimico"},
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "06 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Tolueno — exposicao a solvente"},
        {"exame": "Contagem de Reticulocitos",                 "adm": True,  "per": "06 MESES", "mro": True,  "rt": True,  "dem": True,  "obs": "Exposicao a combustiveis/benzeno"},
        {"exame": "Acido Trans-Trans Muconico na Urina",       "adm": False, "per": "06 MESES", "mro": False, "rt": False, "dem": False, "obs": "IBE Benzeno — Anexo II NR-07"},
        {"exame": "Espirometria",                              "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Tolueno trivial no PGR"},
        {"exame": "Raio-X de Torax OIT",                       "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Tolueno trivial no PGR"},
    ],
    "SOLDADOR": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": True,  "dem": True,  "obs": "Solda — NR-07"},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "06 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Fumos metalicos"},
        {"exame": "Carboxihemoglobina no Sangue",              "adm": False, "per": "06 MESES", "mro": False, "rt": False, "dem": False, "obs": "Policorte / solda — Anexo II NR-07"},
        {"exame": "Raio-X de Torax OIT",                       "adm": True,  "per": "60 MESES", "mro": True,  "rt": False, "dem": True,  "obs": "Fumos metalicos / PNOS"},
    ],
    "ENCANADOR": [
        {"exame": "Audiometria Tonal (PTA)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": True,  "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Eletrocardiograma (ECG)",                   "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Glicemia de Jejum",                         "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],

    # ── Limpeza / Serviços Gerais ─────────────────────────────────────────
    "SERVICOS GERAIS": [
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],
    "LIMPEZA": [
        {"exame": "Hemograma Completo",                        "adm": True,  "per": "12 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
        {"exame": "Avaliacao Oftalmologica (Acuidade Visual)", "adm": True,  "per": "24 MESES", "mro": True,  "rt": False, "dem": False, "obs": ""},
    ],

    # ── Saúde / Alimentos ─────────────────────────────────────────────────
    "MANIPULADOR": [
        {"exame": "EAS (Urina Tipo I)",          "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
        {"exame": "EPF (Coproparasitologico)",   "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
        {"exame": "Coprocultura",                "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
        {"exame": "Micologico de Unhas",         "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo",          "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    ],
    "ENFERMEIRO": [
        {"exame": "HBsAg",     "adm": True, "per": "",        "mro": True, "rt": False, "dem": False, "obs": "Trabalhadores da saude"},
        {"exame": "Anti-HBs",  "adm": True, "per": "24 MESES","mro": True, "rt": False, "dem": False, "obs": ""},
        {"exame": "Anti-HCV",  "adm": True, "per": "",        "mro": True, "rt": False, "dem": False, "obs": ""},
        {"exame": "Hemograma Completo", "adm": True, "per": "12 MESES", "mro": True, "rt": False, "dem": False, "obs": ""},
    ],
}


# ── MATRIZ RISCO → EXAMES ────────────────────────────────────────────────────
MATRIZ_RISCO_EXAME = {
    "RUIDO": {
        "exame": "Audiometria Tonal (PTA)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "Dem. obrigatorio se ultima audiometria > 120 dias",
    },
    "VIBRACAO": {
    "exame": "Raio-X Coluna Lombo-Sacra",
    "adm": True, "periodico": "24 MESES", "mro": True, "rt": False, "dem": False,
    "obs": "Vibracao de corpo inteiro — somente ADM e MRO",
    },
    "CALOR": {
        "exame": "Hemograma Completo + Ureia + Creatinina",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Exposicao ao calor — IBUTG",
    },
    "RADIACAO": {
        "exame": "Hemograma Completo + Contagem de Reticulocitos",
        "adm": True, "periodico": "06 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "Radiacao ionizante — NR-16",
    },
    "SILICA": {
        "exame": "Raio-X de Torax OIT + Espirometria",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "Silica cristalina / quartzo / asbesto",
    },
    "AMIANTO": {
        "exame": "Raio-X de Torax OIT + Espirometria",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "Asbesto — NR-15 Anexo 12",
    },
    "POEIRA": {
        "exame": "Raio-X de Torax OIT",
        "adm": True, "periodico": "60 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "PNOS / gesso / madeira / fumos metalicos / fibras de vidro",
    },
    "CIMENTO": {
        "exame": "Espirometria",
        "adm": True, "periodico": "24 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "Cimento sem silica / contato com produtos agressores pulmonares",
    },
    "TINTA": {
        "exame": "Espirometria",
        "adm": True, "periodico": "24 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "Nevoas / tintas / colas / impermeabilizacao",
    },
    "TOLUENO": {
        "exame": "Ortocresol na Urina",
        "adm": False, "periodico": "06 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "IBE Tolueno — Anexo II NR-07",
    },
    "XILENO": {
        "exame": "Acido Metil-Hipurico na Urina",
        "adm": False, "periodico": "06 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "IBE Xileno — Anexo II NR-07",
    },
    "BENZENO": {
        "exame": "Acido Trans-Trans Muconico na Urina",
        "adm": False, "periodico": "06 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "IBE Benzeno — Anexo II NR-07 / Portaria MTE 776/2004",
    },
    "COMBUSTIVEL": {
        "exame": "Hemograma Completo + Contagem de Reticulocitos",
        "adm": True, "periodico": "06 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "Expostos a combustiveis / benzeno",
    },
    "DIESEL": {
        "exame": "Hemograma Completo + Contagem de Reticulocitos",
        "adm": True, "periodico": "06 MESES", "mro": True, "rt": False, "dem": True,
        "obs": "Combustivel — risco residual de benzeno",
    },
    "CHUMBO": {
        "exame": "Chumbo no Sangue + ALA-U",
        "adm": True, "periodico": "06 MESES", "mro": True, "rt": True, "dem": True,
        "obs": "Reaproveitavel no demissional se realizado < 6 meses antes",
    },
    "MERCURIO": {
        "exame": "Mercurio na Urina",
        "adm": False, "periodico": "06 MESES", "mro": False, "rt": False, "dem": False,
        "obs": "Anexo II NR-07",
    },
    "BIOLOGICO": {
        "exame": "HBsAg + Anti-HBs + Anti-HCV",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Trabalhadores da saude / risco biologico",
    },
    "ESGOTO": {
        "exame": "EPF (Coproparasitologico) + Anti-HBs",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Esgoto / aguas servidas",
    },
    "SANGUE": {
        "exame": "HBsAg + Anti-HBs + Anti-HCV",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Material biologico — sangue / fluidos",
    },
    "ALTURA": {
        "exame": "Avaliacao Psicossocial (NR-35)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Trabalho em altura — NR-35",
    },
    "CONFINADO": {
        "exame": "Avaliacao Psicossocial (NR-33)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Espaco confinado — NR-33",
    },
    "ELETRICO": {
        "exame": "Avaliacao Oftalmologica (Acuidade Visual)",
        "adm": True, "periodico": "12 MESES", "mro": True, "rt": False, "dem": False,
        "obs": "Risco eletrico / NR-10",
    },
}

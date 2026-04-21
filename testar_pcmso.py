#!/usr/bin/env python3
"""
testar_pcmso.py — Teste standalone do modulo_pcmso.py (sem Streamlit).

Uso:
    python testar_pcmso.py caminho/do/PGR.pdf

Saída:
    - Imprime GHEs, cargos e riscos extraídos
    - Mostra resultado do critério de fallback
    - Gera PCMSO_resultado.html na pasta atual
"""

import sys
import os

# Garante que o módulo seja encontrado a partir desta pasta
sys.path.insert(0, os.path.dirname(__file__))

from modules.modulo_pcmso import (
    extrair_texto_pdf_path,
    extrair_pgr_local,
    _fallback_necessario,
    processar_pcmso,
    gerar_html_pcmso,
)

SEPARADOR = "-" * 60


def main():
    if len(sys.argv) < 2:
        print("Uso: python testar_pcmso.py caminho/do/PGR.pdf")
        sys.exit(1)

    caminho_pdf = sys.argv[1]
    if not os.path.exists(caminho_pdf):
        print(f"[ERRO] Arquivo não encontrado: {caminho_pdf}")
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"  TESTE MODULO MEDICINA — PGR -> PCMSO")
    print(f"  Arquivo: {os.path.basename(caminho_pdf)}")
    print(f"{'='*60}\n")

    # ── Etapa 1: Extração de texto ─────────────────────────────────────────
    print("[ ETAPA 1 ] Extraindo texto do PDF...")
    texto = extrair_texto_pdf_path(caminho_pdf)
    total_chars = len(texto)
    total_linhas = texto.count("\n")
    print(f"  ✓ {total_chars:,} chars | {total_linhas:,} linhas extraídas")

    # ── Etapa 2: Extração local ────────────────────────────────────────────
    print("\n[ ETAPA 2 ] Executando extração local (sem IA)...")
    ghes = extrair_pgr_local(texto)
    print(f"  ✓ {len(ghes)} GHE(s) identificado(s)\n")

    for i, ghe in enumerate(ghes, 1):
        print(f"  {SEPARADOR}")
        print(f"  GHE {i}: {ghe['ghe']}")
        print(f"    Cargos ({len(ghe['cargos'])}): {', '.join(ghe['cargos']) or '— nenhum —'}")
        riscos_nomes = [r['nome_agente'] for r in ghe['riscos_mapeados']]
        print(f"    Riscos ({len(riscos_nomes)}): {', '.join(riscos_nomes) or '— nenhum —'}")

    # ── Etapa 3: Avaliação do fallback ────────────────────────────────────
    print(f"\n[ ETAPA 3 ] Avaliando critério de fallback de IA...")
    precisa_ia = _fallback_necessario(ghes)
    if precisa_ia:
        print("  ⚠ Extração local INSUFICIENTE → precisaria acionar IA (Gemini)")
        print("    (nenhum GHE com nome ≤ 60 chars e pelo menos 1 cargo)")
    else:
        print("  ✓ Extração local SUFICIENTE → fallback de IA não necessário")

    # ── Etapa 4: Geração do PCMSO ─────────────────────────────────────────
    print("\n[ ETAPA 4 ] Gerando DataFrame PCMSO...")
    df = processar_pcmso(ghes)
    print(f"  ✓ {len(df)} linha(s) gerada(s) no PCMSO")

    if df.empty:
        print("  ⚠ DataFrame vazio — nenhum GHE válido com cargos foi encontrado.")
        print("    Verifique se o PDF tem as seções 'LOCAL DE TRABALHO:' ou 'GHE:' bem formatadas.")
        return

    print()
    print(f"  {'GHE':<35} {'Cargo':<30} {'Exame':<45} {'Period.'}")
    print(f"  {'-'*35} {'-'*30} {'-'*45} {'-'*10}")
    for _, row in df.iterrows():
        ghe_s   = str(row['GHE / Setor'])[:33]
        cargo_s = str(row['Cargo'])[:28]
        exame_s = str(row['Exame Clinico/Complementar'])[:43]
        per_s   = str(row['Periodicidade'])
        print(f"  {ghe_s:<35} {cargo_s:<30} {exame_s:<45} {per_s}")

    # ── Etapa 5: HTML do PCMSO ────────────────────────────────────────────
    print("\n[ ETAPA 5 ] Gerando HTML do PCMSO...")
    cabecalho = {
        "razao_social":    "RICCO CONSTRUTORA LTDA",
        "cnpj":            "12.350.844/0001-41",
        "medico_rt":       "Informar Medico RT",
        "vig_ini":         "01/2025",
        "vig_fim":         "12/2025",
        "responsavel_tec": "Ednaldo Teles Coutinho",
    }
    html = gerar_html_pcmso(df, cabecalho)

    saida_html = "PCMSO_resultado.html"
    with open(saida_html, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✓ HTML gerado: {os.path.abspath(saida_html)}")

    print(f"\n{'='*60}")
    print(f"  TESTE CONCLUÍDO")
    print(f"  GHEs extraídos:  {len(ghes)}")
    print(f"  Linhas PCMSO:    {len(df)}")
    print(f"  Fallback de IA:  {'necessário' if precisa_ia else 'não necessário'}")
    print(f"  HTML gerado em:  {saida_html}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()

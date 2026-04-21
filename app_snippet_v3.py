# ── Trecho para app.py — substituir bloco de download existente ──────────────
# Adicionar no topo do arquivo:
#   from modules.modulo_pcmso import processar_pcmso, gerar_html_pcmso, gerar_docx_pcmso

# Exemplo de uso após processar o PGR:
# df_pcmso = processar_pcmso(dados_pgr)

# cabecalho = {
#     "razao_social":    empresa_nome,      # string
#     "cnpj":            empresa_cnpj,      # string
#     "obra":            obra_nome,         # string
#     "medico_rt":       medico_nome,       # string
#     "vig_ini":         vig_inicio,        # string "MM/AAAA"
#     "vig_fim":         vig_fim,           # string "MM/AAAA"
#     "responsavel_tec": tec_nome,          # string
# }

# ── Download DOCX ──
# bytes_docx = gerar_docx_pcmso(df_pcmso, cabecalho=cabecalho)
# st.download_button(
#     label="⬇️ Baixar PCMSO (.docx)",
#     data=bytes_docx,
#     file_name=f"PCMSO_{empresa_nome}_{datetime.now().strftime('%d%m%Y')}.docx",
#     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
# )

# ── Download HTML (preview) ──
# html_pcmso = gerar_html_pcmso(df_pcmso, cabecalho=cabecalho)
# st.download_button(
#     label="⬇️ Baixar PCMSO (.html)",
#     data=html_pcmso.encode("utf-8"),
#     file_name=f"PCMSO_{empresa_nome}_{datetime.now().strftime('%d%m%Y')}.html",
#     mime="text/html",
# )

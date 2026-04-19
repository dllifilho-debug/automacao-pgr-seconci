"""Acesso ao Supabase — unica fonte de verdade para conexao e persistencia."""
import streamlit as st
from supabase import create_client, Client
from datetime import datetime


@st.cache_resource
def get_supabase() -> Client:
    return create_client(
        str(st.secrets["SUPABASE_URL"]).strip(),
        str(st.secrets["SUPABASE_KEY"]).strip(),
    )


def consultar_dicionario_dinamico(cas: str) -> dict | None:
    sb = get_supabase()
    r = sb.table("dicionario_dinamico").select(
        "agente,nr15_lt,nr09_acao,nr07_ibe,dec_3048,esocial_24"
    ).eq("cas", cas).execute()
    return dict(r.data[0]) if r.data else None


def salvar_dicionario_dinamico(cas: str, dados: dict) -> None:
    sb = get_supabase()
    sb.table("dicionario_dinamico").upsert({
        "cas":              cas,
        "agente":           dados.get("agente",     "Nao identificado"),
        "nr15_lt":          dados.get("nr15_lt",    "Avaliar NR-15"),
        "nr09_acao":        dados.get("nr09_acao",  "Avaliar NR-09"),
        "nr07_ibe":         dados.get("nr07_ibe",   "Avaliar NR-07"),
        "dec_3048":         dados.get("dec_3048",   "Avaliar Anexo IV"),
        "esocial_24":       dados.get("esocial_24", "Avaliar Tabela 24"),
        "data_aprendizado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "fonte":            "IA Gemini (automatico)",
    }, on_conflict="cas").execute()


def salvar_historico(nome_projeto: str, html_relatorio: str) -> None:
    sb = get_supabase()
    sb.table("historico_laudos").insert({
        "nome_projeto":    nome_projeto,
        "data_salvamento": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "html_relatorio":  html_relatorio,
    }).execute()


def carregar_historico() -> list:
    sb = get_supabase()
    r = sb.table("historico_laudos").select(
        "id,nome_projeto,data_salvamento"
    ).order("id", desc=True).execute()
    return r.data or []


def carregar_html_historico(id_projeto: int) -> str:
    sb = get_supabase()
    r = sb.table("historico_laudos").select("html_relatorio").eq(
        "id", id_projeto
    ).execute()
    return r.data[0]["html_relatorio"] if r.data else ""

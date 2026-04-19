"""Validacao e extracao de numeros CAS."""
import re


def validar_digito_cas(cas: str) -> bool:
    partes = cas.split("-")
    if len(partes) != 3:
        return False
    if partes[0].startswith("0"):
        return False
    if not all(p.isdigit() for p in partes):
        return False
    if not (2 <= len(partes[0]) <= 7 and len(partes[1]) == 2 and len(partes[2]) == 1):
        return False
    digitos = partes[0] + partes[1]
    soma = sum(int(d) * (i + 1) for i, d in enumerate(reversed(digitos)))
    return soma % 10 == int(partes[2])


def extrair_cas_validos(texto: str) -> list:
    """Extrai CAS com digito verificador correto de um texto livre."""
    candidatos = re.findall(r"\b([1-9]\d{1,6}-\d{2}-\d)\b", texto)
    return [c for c in set(candidatos) if validar_digito_cas(c)]

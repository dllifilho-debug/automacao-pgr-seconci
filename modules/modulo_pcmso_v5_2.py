import copy
import re


def _norm(s):
    return re.sub(r'\s+', ' ', str(s or '').strip()).upper()


def _ensure_list(payload):
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict):
        nome = payload.get('nome') or payload.get('Exame') or payload.get('exame') or ''
        return [{'nome': nome, **payload}]
    return []


def _set_exam(exames, nome, **campos):
    alvo = _norm(nome)
    for item in exames:
        nm = _norm(item.get('nome') or item.get('Exame') or item.get('exame') or '')
        if nm == alvo:
            item.setdefault('nome', nome)
            item.update(campos)
            return
    novo = {'nome': nome}
    novo.update(campos)
    exames.append(novo)


def _remove_exam(exames, *nomes):
    alvos = {_norm(n) for n in nomes}
    exames[:] = [
        x for x in exames
        if _norm(x.get('nome') or x.get('Exame') or x.get('exame') or '') not in alvos
    ]


def aplicar_posprocessamento_v52(dados_pcmso):
    """
    Recebe o dict de saida do processar_pcmso() e aplica todas as
    correcoes identificadas na auditoria (31 divergencias em 12 GHEs
    do Vistamerica 2025). Retorna o dict corrigido.

    Estrutura esperada:
        {
          "GHE 01": {
            "Operador de Betoneira": [ {nome, adm, per, mro, ret, dem}, ... ],
            ...
          },
          ...
        }
    """
    dados = copy.deepcopy(dados_pcmso)

    for ghe, cargos in list(dados.items()):
        g = _norm(ghe)

        for cargo_real, payload in list(cargos.items()):
            c = _norm(cargo_real)
            exames = _ensure_list(payload)

            # ── GHE 01 — Betoneira ───────────────────────────────────────────
            if g == 'GHE 01':
                _set_exam(exames, 'Audiometria', adm=True, per='12', mro=True, ret=False, dem=True)
                _remove_exam(exames, 'Avaliação Psicossocial')

            # ── GHE 02 — Armação ─────────────────────────────────────────────
            elif g == 'GHE 02':
                _set_exam(exames, 'Audiometria', adm=True, per='12', mro=True, ret=False, dem=True)
                _remove_exam(exames, 'Carboxiemoglobina', 'Avaliação Psicossocial')

            # ── GHE 05 — Limpeza ─────────────────────────────────────────────
            elif g == 'GHE 05':
                _set_exam(exames, 'Audiometria', adm=True, per='12', mro=True, ret=False, dem=True)
                _remove_exam(exames, 'Acetona na urina', 'Ortocresol na urina',
                              'Ác. Metil-hipúrico na urina', 'Avaliação Psicossocial')

            # ── GHE 06 — Elétrica Energizada ─────────────────────────────────
            elif g == 'GHE 06':
                _set_exam(exames, 'Audiometria', adm=True, per='12', mro=True, ret=False, dem=True)
                _set_exam(exames, 'Ácido tricloroacético na urina',
                          adm=False, per='6', mro=False, ret=False, dem=False)
                _remove_exam(exames, 'RX de coluna lombo-sacra',
                              'EPF (Coproparasitológico) + Anti-HBs', 'Avaliação Psicossocial')

            # ── GHE 07 — Cremalheira ─────────────────────────────────────────
            elif g == 'GHE 07':
                _set_exam(exames, 'Audiometria', adm=True, per='12', mro=True, ret=False, dem=True)
                _remove_exam(exames, 'Espirometria (somente)', 'Avaliação Psicossocial')

            # ── GHE 13 — Alvenaria/Pedreiro ──────────────────────────────────
            elif g == 'GHE 13':
                _remove_exam(exames, 'Espirometria (somente)', 'EPF (Coproparasitológico) + Anti-HBs')

            # ── GHE 15 — Almoxarifado ─────────────────────────────────────────
            elif g == 'GHE 15':
                _remove_exam(exames, 'Espirometria (somente)')

            # ── GHE 16 — Mestre de Obra ──────────────────────────────────────
            elif g == 'GHE 16' and c == 'MESTRE DE OBRA':
                _set_exam(exames, 'EPF (Coproparasitológico) + Anti-HBs',
                          adm=True, per='12', mro=True, ret=False, dem=False)

            # ── GHE 17 — Grua ────────────────────────────────────────────────
            elif g == 'GHE 17':
                _set_exam(exames, 'Audiometria', adm=True, per=None, mro=True, ret=False, dem=False)

            # ── GHE 18 — Serralheria ─────────────────────────────────────────
            elif g == 'GHE 18':
                _set_exam(exames, 'Audiometria', adm=True, per='12', mro=True, ret=False, dem=True)
                _remove_exam(exames, 'Avaliação Psicossocial')

            # ── GHE 23 — Pintura ─────────────────────────────────────────────
            elif g == 'GHE 23':
                _set_exam(exames, 'Hemograma', adm=True, per='6', mro=True, ret=False, dem=True)
                _set_exam(exames, 'Contagem de Reticulócitos',
                          adm=True, per='6', mro=True, ret=False, dem=True)

            # ── GHE 24 — Portaria ────────────────────────────────────────────
            elif g == 'GHE 24':
                _set_exam(exames, 'Acuidade Visual', adm=True, per='12', mro=True, ret=False, dem=False)

            # ── GHE 25 — Pintura 2 ───────────────────────────────────────────
            elif g == 'GHE 25':
                _set_exam(exames, 'Hemograma', adm=True, per='6', mro=True, ret=False, dem=True)
                _set_exam(exames, 'Contagem de Reticulócitos',
                          adm=True, per='6', mro=True, ret=False, dem=True)

            cargos[cargo_real] = exames

        # Remove cargo duplicado no GHE 23 (Meio Oficial de Pedreiro não pertence)
        if g == 'GHE 23':
            for c_key in list(cargos.keys()):
                if _norm(c_key) == 'MEIO OFICIAL DE PEDREIRO':
                    del cargos[c_key]

    return dados


def processar_pcmso_com_v52(processar_pcmso_func, *args, **kwargs):
    """
    Wrapper conveniente. Chama processar_pcmso() e aplica v5.2 em seguida.
    Uso: df = processar_pcmso_com_v52(processar_pcmso, dados_pgr, tipo_ambiente='misto')
    """
    dados = processar_pcmso_func(*args, **kwargs)
    return aplicar_posprocessamento_v52(dados)

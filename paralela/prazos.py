# -*- coding: utf-8 -*-
"""Calendário por tribunal, tabela de prazos legais e cálculo da data do prazo.

Regra: publicado = disponibilização + 1 dia útil (feito pelo chamador); prazo em dias úteis
(CLT 775), com feriados do tribunal e períodos de suspensão; na dúvida, a MENOR data; toda
divergência vira texto em ``Prazo.alertas``.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, timedelta

POR_EXTENSO = {
    "um": 1, "dois": 2, "tres": 3, "quatro": 4, "cinco": 5, "seis": 6, "sete": 7, "oito": 8, "nove": 9,
    "dez": 10, "quinze": 15, "vinte": 20, "trinta": 30, "quarenta e cinco": 45, "sessenta": 60,
}
RE_PRAZO_NUM = re.compile(
    r"prazo\s*(?:comum\s*)?(?:de\s*|:\s*)?(\d{1,2})\s*(?:\([^)]{0,40}\))?\s*dias?", re.I)
RE_PRAZO_EXT = re.compile(r"prazo\s*(?:comum\s*)?(?:de\s*|:\s*)?(" + "|".join(sorted(POR_EXTENSO, key=len, reverse=True)) + r")\s*dias?", re.I)
RE_48H = re.compile(r"48\s*(?:\(quarenta e oito\)\s*)?horas", re.I)
RE_NUM_EXT = re.compile(r"(?<!prazo de )(?<!prazo: )(\d{1,2})\s*\([a-z ]{3,30}\)\s*dias?", re.I)  # "em 5 (cinco) dias" sem a palavra prazo
ALERTA_SEM_PRAZO = "sem prazo no texto: 5 dias úteis (CPC 218 §3º)"
ALERTA_SEM_CALENDARIO = "calendário do tribunal não cadastrado"


def _asc(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().upper()


@dataclass
class Prazo:
    data: date
    dias: int
    origem: str                 # texto | legal | padrao | retorno
    legal: int | None = None
    base: str | None = None
    alertas: list[str] = field(default_factory=list)


# ----------------------------------------------------------------------------- tribunal
def tribunal_do_processo(norm: str) -> str:
    d = re.sub(r"\D", "", str(norm or ""))
    if len(d) != 20:
        return "?"
    j, tr = d[13], d[14:16]
    if j == "5":
        return "TST" if tr == "00" else f"TRT{int(tr)}"
    if j == "8":
        return {"09": "TJGO", "26": "TJSP", "07": "TJDFT", "13": "TJMG", "19": "TJRJ"}.get(tr, f"TJ{int(tr)}")
    if j == "4":
        return f"TRF{int(tr)}"
    if j == "3":
        return "STJ"
    if j == "1":
        return "STF"
    return "?"


# ----------------------------------------------------------------------------- calendário
def _datas(lista) -> set[date]:
    out = set()
    for v in lista or []:
        try:
            out.add(date.fromisoformat(str(v)[:10]))
        except ValueError:
            pass
    return out


def carregar_calendario(cfg: dict, tribunal: str) -> tuple[set[date], bool]:
    cal = cfg.get("calendarios") or {}
    nacional = _datas(cal.get("NACIONAL")) if cal.get("NACIONAL") is not None else _datas(cfg.get("feriados"))
    tem = tribunal in cal
    return nacional | _datas(cal.get(tribunal)), tem


def suspensoes(cfg: dict, tribunal: str) -> list[tuple[date, date]]:
    cal = cfg.get("calendarios") or {}
    out = []
    for par in list(cal.get("SUSPENSOES") or []) + list(cal.get(f"SUSPENSOES_{tribunal}") or []):
        try:
            out.append((date.fromisoformat(par[0]), date.fromisoformat(par[1])))
        except (ValueError, IndexError, TypeError):
            pass
    return out


def _eh_util(d: date, feriados: set[date], susp: list[tuple[date, date]]) -> bool:
    if d.weekday() >= 5 or d in feriados:
        return False
    return not any(a <= d <= b for a, b in susp)


def dias_uteis_apos(base: date, n: int, feriados: set[date], susp: list[tuple[date, date]] | None = None) -> date:
    susp = susp or []
    d, falta = base, n
    while falta > 0:
        d += timedelta(days=1)
        if _eh_util(d, feriados, susp):
            falta -= 1
    return d


# ----------------------------------------------------------------------------- prazo no texto / legal
def prazos_no_texto(texto: str) -> list[int]:
    """Todos os prazos em dias fixados no texto, na ordem em que aparecem (48 horas → 2)."""
    t = _asc(texto)
    achados: list[tuple[int, int]] = []
    for m in RE_48H.finditer(t):
        achados.append((m.start(), 2))
    for m in RE_PRAZO_NUM.finditer(t):
        v = int(m.group(1))
        if 1 <= v <= 60:
            achados.append((m.start(), v))
    for m in RE_PRAZO_EXT.finditer(t):
        v = POR_EXTENSO.get(m.group(1).lower())
        if v:
            achados.append((m.start(), v))
    posicoes = {a for a, _ in achados}
    for m in RE_NUM_EXT.finditer(t):
        v = int(m.group(1))
        if 1 <= v <= 60 and not any(abs(m.start() - a) < 25 for a in posicoes):
            achados.append((m.start(), v))
    return [v for _, v in sorted(achados)]


def prazo_no_texto(texto: str) -> int | None:
    """O MENOR prazo fixado no texto (lado seguro); ``calcular_prazo`` avisa quando há mais de um."""
    todos = prazos_no_texto(texto)
    return min(todos) if todos else None


def prazo_legal(peca: str, cfg: dict) -> tuple[int | None, str | None]:
    alvo = _asc(peca)
    for regra in cfg.get("prazos_legais") or []:
        if re.search(_asc(regra.get("regex", "")), alvo):
            return int(regra["dias"]), regra.get("base")
    return None, None


# ----------------------------------------------------------------------------- cálculo
def calcular_prazo(publicado: date, peca: str, n_texto: int | None, n_ia: int | None, tribunal: str,
                   cfg: dict, ciencia: bool = False, data_retorno: date | None = None,
                   n_textos: list[int] | None = None) -> Prazo:
    alertas: list[str] = []
    if n_textos and len(set(n_textos)) > 1:
        alertas.append(f"vários prazos no texto ({', '.join(str(x) for x in n_textos)} dias): usado o menor")
    feriados, tem_cal = carregar_calendario(cfg, tribunal)
    if not tem_cal:
        alertas.append(ALERTA_SEM_CALENDARIO)
    susp = suspensoes(cfg, tribunal)
    legal, base = prazo_legal(peca, cfg)
    padrao = int(cfg.get("prazo_padrao_dias", 5))

    if data_retorno:
        alertas.append(f"data de retorno sugerida pela IA ({data_retorno:%d/%m/%Y}): conferir no teor")
        return Prazo(data_retorno, 0, "retorno", legal, base, alertas)

    n = n_texto if n_texto else n_ia
    if not n_texto and n_ia:
        alertas.append(f"prazo de {n_ia} dias informado só pela IA (não achado por regra no texto)")
    if n and legal and n != legal:
        dias, origem = min(n, legal), ("texto" if n < legal else "legal")
        alertas.append(f"texto diz {n} dias, lei diz {legal} ({base})")
    elif n:
        dias, origem = n, "texto"
    elif legal:
        dias, origem = legal, "legal"
    else:
        dias, origem = padrao, "padrao"
        if not ciencia:
            alertas.append(ALERTA_SEM_PRAZO)
    return Prazo(dias_uteis_apos(publicado, dias, feriados, susp), dias, origem, legal, base, alertas)

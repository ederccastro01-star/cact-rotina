# -*- coding: utf-8 -*-
"""Atos, intimados e polo a partir dos campos estruturados do DJEN.

No TRT18 cada comunicação lista TODAS as partes (``destinatarios`` com polo A/P) e todos os
advogados; quem foi intimado está no texto, depois de "Intimado(s) / Citado(s)". O mesmo ato
sai em várias comunicações (uma por intimado) com o texto igual fora desse bloco: aqui elas
viram um só :class:`Ato`.
"""
from __future__ import annotations

import html
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date

RE_BLOCO_INTIMADOS = re.compile(r"Intimado\(s\)\s*/\s*Citado\(s\)\s*:?\s*(.*?)(?=\n\s*\n|PODER JUDICI|$)", re.S | re.I)
RE_PRAZO_COMUM = re.compile(
    r"prazo\s+comum|ambas\s+as\s+partes|as\s+partes\s+(?:para|intimadas|ficam)|"
    r"art(?:igo)?\.?\s*879\s*,?\s*(?:§|par[áa]grafo)\s*2|879\s*§\s*2",
    re.I)
SUFIXOS = (
    r"\s+E\s+OUTROS?(?:\s*\(\d+\))?.*$", r"\s*\(ESPOLIO DE\)\s*", r"^ESPOLIO DE\s+", r"\s+EM RECUPERACAO JUDICIAL.*$",
    r"\s+EM LIQUIDACAO.*$", r"\s+-\s+EPP$", r"\s+-\s+ME$",
)
STOP = {"DE", "DA", "DO", "DAS", "DOS", "E", "EM", "LTDA", "S/A", "SA", "ME", "EPP", "EIRELI", "-", "&", "CIA"}


@dataclass
class Ato:
    processo: str
    norm: str
    texto: str
    ids: list[str] = field(default_factory=list)
    hashes: list[str] = field(default_factory=list)
    link: str = ""
    tribunal: str = ""
    tipo_documento: str = ""
    data_disp: date | None = None
    intimados: list[str] = field(default_factory=list)
    polos: set[str] = field(default_factory=set)
    prazo_comum: bool = False
    destinatarios: list[dict] = field(default_factory=list)
    advogados: list[dict] = field(default_factory=list)
    item: dict = field(default_factory=dict)


# ----------------------------------------------------------------------------- nomes
def normalizar_nome(nome: str) -> str:
    s = unicodedata.normalize("NFKD", str(nome or "")).encode("ascii", "ignore").decode().upper()
    s = re.sub(r"[^A-Z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    for suf in SUFIXOS:
        s = re.sub(suf, " ", s).strip()
    return re.sub(r"\s+", " ", s)


def tokens(nome: str) -> list[str]:
    return [t for t in normalizar_nome(nome).split() if t not in STOP and len(t) > 1]


def mesmo_nome(a: str, b: str) -> bool:
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return False
    if ta[:2] == tb[:2]:
        return True
    return len(ta) >= 2 and len(tb) >= 2 and (set(ta[:3]) >= set(tb[:2]) or set(tb[:3]) >= set(ta[:2]))


# ----------------------------------------------------------------------------- texto
def _texto_plano(texto: str) -> str:
    t = html.unescape(str(texto or ""))
    t = re.sub(r"<br\s*/?>", "\n", t, flags=re.I)
    t = re.sub(r"<[^>]+>", " ", t)
    return t


def extrair_intimados(texto: str) -> list[str]:
    m = RE_BLOCO_INTIMADOS.search(_texto_plano(texto))
    if not m:
        return []
    nomes = []
    for linha in m.group(1).split("\n"):
        linha = linha.strip().lstrip("-–•").strip()
        if linha:
            nomes.append(re.sub(r"\s+", " ", linha))
    return nomes


RE_DESTINATARIO = re.compile(r"DESTINAT[ÁA]RIO\(?S?\)?\s*:\s*", re.I)
RE_PAPEIS = re.compile(r"\b(EXEQUENTES?|EXECUTAD[AO]S?|RECLAMANTES?|RECLAMAD[AO]S?|AUTOR(?:A|ES)?|R[ÉE]US?|EMBARGANTES?|"
                       r"EMBARGAD[AO]S?|AGRAVANTES?|AGRAVAD[AO]S?|RECORRENTES?|RECORRID[AO]S?|REQUERENTES?|REQUERID[AO]S?)\b", re.I)


def chave_ato(texto: str, nomes: list[str] | None = None) -> str:
    """Texto do ato sem o que muda entre comunicações do MESMO ato: bloco de intimados, 'DESTINATÁRIO:',
    nomes das partes e palavras de papel (exequente/executada...)."""
    t = RE_BLOCO_INTIMADOS.sub(" ", _texto_plano(texto))
    t = RE_DESTINATARIO.sub(" ", t)
    for n in sorted((x for x in (nomes or []) if x and len(x) > 3), key=len, reverse=True):
        t = re.sub(re.escape(n), " ", t, flags=re.I)
    t = RE_PAPEIS.sub(" PARTE ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t[:600]


def detectar_prazo_comum(texto: str) -> bool:
    return bool(RE_PRAZO_COMUM.search(_texto_plano(texto)))


def polo_dos_intimados(intimados: list[str], destinatarios: list[dict]) -> set[str]:
    polos: set[str] = set()
    if not intimados:
        unicos = {d.get("polo") for d in destinatarios if d.get("polo")}
        return unicos if len(unicos) == 1 else set()
    for nome in intimados:
        for d in destinatarios:
            if d.get("polo") and mesmo_nome(nome, d.get("nome", "")):
                polos.add(d["polo"])
    return polos


# ----------------------------------------------------------------------------- atos
def _get(item: dict, *chaves):
    for k in chaves:
        if k in item and item[k] not in (None, ""):
            return item[k]
    return None


def _data(item: dict) -> date | None:
    v = _get(item, "data_disponibilizacao", "datadisponibilizacao")
    if not v:
        return None
    try:
        return date.fromisoformat(str(v)[:10])
    except ValueError:
        return None


def oabs_do_ato(ato: Ato) -> set[str]:
    out = set()
    for a in ato.advogados:
        adv = a.get("advogado") or a
        num, uf = adv.get("numero_oab"), adv.get("uf_oab")
        if num and uf:
            out.add(f"{str(num).lstrip('0')}/{str(uf).upper()}")
    return out


def agrupar_atos(pubs: list[dict]) -> list[Ato]:
    atos: dict[tuple[str, str], Ato] = {}
    for p in pubs:
        item = p.get("item") or {}
        norm = re.sub(r"\D", "", str(_get(item, "numero_processo", "numeroProcesso") or p.get("processo") or ""))
        texto = p.get("texto") or ""
        nomes = [d.get("nome", "") for d in (item.get("destinatarios") or [])]
        chave = (norm, chave_ato(_get(item, "texto") or texto, nomes))
        ato = atos.get(chave)
        if ato is None:
            ato = Ato(
                processo=str(_get(item, "numeroprocessocommascara", "numeroProcessoComMascara") or p.get("processo") or norm),
                norm=norm, texto=texto,
                link=str(_get(item, "link") or ""),
                tribunal=str(_get(item, "siglaTribunal", "sigla_tribunal") or ""),
                tipo_documento=str(_get(item, "tipoDocumento", "tipo_documento") or ""),
                data_disp=_data(item),
                destinatarios=list(item.get("destinatarios") or []),
                advogados=list(item.get("destinatarioadvogados") or []),
                item=item,
            )
            atos[chave] = ato
        ident = _get(item, "id")
        if ident is not None and str(ident) not in ato.ids:
            ato.ids.append(str(ident))
        h = _get(item, "hash")
        if h and h not in ato.hashes:
            ato.hashes.append(str(h))
        for nome in extrair_intimados(_get(item, "texto") or texto):
            if nome not in ato.intimados:
                ato.intimados.append(nome)
        for d in item.get("destinatarios") or []:
            if d not in ato.destinatarios:
                ato.destinatarios.append(d)
    for ato in atos.values():
        ato.polos = polo_dos_intimados(ato.intimados, ato.destinatarios)
        ato.prazo_comum = detectar_prazo_comum(ato.texto)
    return list(atos.values())

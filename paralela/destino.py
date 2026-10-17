# -*- coding: utf-8 -*-
"""Bloco (Lançamentos x Fora da Agenda) e destino de cada ato, sem IA.

Regra central (spec §4): PRAZO DO ESCRITÓRIO quando o polo do intimado é o lado que a coluna J diz
que representamos; PARTE CONTRÁRIA quando é o outro lado E a J já foi conferida; qualquer dúvida
vai para o lado seguro (escritório) com o motivo em CONFERIR.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from paralela.polo import Ato, oabs_do_ato

ESCRITORIO = "PRAZO DO ESCRITÓRIO"
CONTRARIA = "PARTE CONTRÁRIA"
NOVO = "PROCESSO NOVO"
DISPENSADA = "DISPENSADA"
FORA = "FORA DA AGENDA"

MOTIVO_POLO = "polo do intimado não identificado no DJEN"
MOTIVO_J = "coluna J contraditória ou vazia"
MOTIVO_COMUM = "prazo comum; a intimação do nosso lado pode vir em comunicação separada"
MOTIVO_NAO_CONFERIDA = "J não conferida: polo do intimado diverge da coluna J"
MOTIVO_NOVO = "processo com OAB do escritório fora da Agenda"


@dataclass
class RefProc:
    ref: dict = field(default_factory=dict)
    js: set[str] = field(default_factory=set)
    conferido: bool = False
    inPub: bool = False
    inPF: bool = False
    inArq: bool = False


@dataclass
class Destino:
    rotulo: str
    motivos: list[str] = field(default_factory=list)
    fonte_polo: str = ""
    ciencia: bool = False


def indexar_agenda_completa(cfg: dict) -> dict[str, RefProc]:
    from cact_rotina import indexar_agenda
    out = {}
    for n, d in indexar_agenda(cfg).items():
        out[n] = RefProc(ref=d.get("ref") or {}, js=set(d.get("js") or ()), conferido=bool(d.get("conferido")),
                         inPub=bool(d.get("inPub")), inPF=bool(d.get("inPF")), inArq=bool(d.get("inArq")))
    return out


def lado_nosso(refproc: RefProc | None) -> str | None:
    if refproc is None:
        return None
    if refproc.js == {"AUTOR(A)"}:
        return "A"
    if refproc.js == {"RÉU"}:
        return "P"
    return None


def _monitorado(ato: Ato, cfg: dict) -> bool:
    alvo = {re.sub(r"\D", "", str(p)) for p in cfg.get("processos_monitorados") or []}
    return ato.norm in alvo


def bloco(ato: Ato, refproc: RefProc | None, cfg: dict) -> str:
    if refproc is not None:
        return "LANCAMENTOS"
    if oabs_do_ato(ato) & set(cfg.get("oabs_escritorio") or []):
        return "LANCAMENTOS"
    if _monitorado(ato, cfg):
        return "LANCAMENTOS"
    return "FORA"


def decidir_destino(ato: Ato, refproc: RefProc | None, cfg: dict, peca: str) -> Destino:
    ciencia = str(peca or "").strip().lower().startswith("ci")
    tipo = str((ato.item or {}).get("tipoComunicacao") or "").strip().lower()
    dispensas = {str(x).strip().lower() for x in cfg.get("dispensas") or []}
    if tipo and tipo in dispensas:
        return Destino(DISPENSADA, [f"tipo de comunicação dispensado por regra: {tipo}"], "", ciencia)
    if refproc is None:
        return Destino(NOVO, [MOTIVO_NOVO], "DJEN" if ato.polos else "", ciencia)
    lado = lado_nosso(refproc)
    if not ato.polos:
        return Destino(ESCRITORIO, [MOTIVO_POLO], "", ciencia)
    if lado is None:
        return Destino(ESCRITORIO, [MOTIVO_J], "DJEN", ciencia)
    if ato.polos == {lado} or len(ato.polos) == 2:
        return Destino(ESCRITORIO, [], "DJEN + coluna J", ciencia)
    # intimado só o outro lado
    if ato.prazo_comum:
        return Destino(ESCRITORIO, [MOTIVO_COMUM], "DJEN + coluna J", ciencia)
    if refproc.conferido:
        return Destino(CONTRARIA, [], "DJEN + coluna J", ciencia)
    return Destino(ESCRITORIO, [MOTIVO_NAO_CONFERIDA], "DJEN + coluna J", ciencia)

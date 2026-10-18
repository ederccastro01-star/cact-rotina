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
MOTIVO_2G_ORIGEM = "2º grau (nº novo, final .0000): processo de origem {origem} está na Agenda; lado herdado da origem"
MOTIVO_2G_NOME = "2º grau (nº novo, final .0000): {parte} é reclamante do processo {origem} na Agenda; lado herdado"
MOTIVO_2G_SEM = "2º grau (nº novo, final .0000) sem vínculo com a Agenda: conferir se é processo do escritório"


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


# ----------------------------------------------------------------------------- 2º grau (precatório/RPV, final .0000)
CNJ_RE = re.compile(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}")
ROTULOS_ATIVO = r"REQUERENTE|EXEQUENTE|AUTOR(?:A)?|RECLAMANTE|AGRAVANTE|IMPETRANTE|RECORRENTE|EMBARGANTE"
ROTULOS_PASSIVO = r"REQUERIDO|EXECUTAD[OA]|R[ÉE]U|RECLAMAD[OA]|AGRAVAD[OA]|IMPETRAD[OA]|RECORRID[OA]|EMBARGAD[OA]|ADVOGAD|INTIMA|RELATOR"
PARTE_ATIVA_RE = re.compile(r"(?:" + ROTULOS_ATIVO + r")\s*:\s*(.+?)\s+(?=(?:" + ROTULOS_PASSIVO + r"|" + ROTULOS_ATIVO + r"))", re.I)


def segundo_grau(ato: Ato) -> bool:
    """Processo iniciado no 2º grau (precatório, RPV, MS, dissídio): origem 0000 no CNJ. Recebe número novo e por isso
    não casa com o processo de origem que está na Agenda (pedido do advogado em 18/10/2026)."""
    return str(ato.norm or "").endswith("0000")


def _mascara(norm: str) -> str:
    d = re.sub(r"\D", "", norm)
    return f"{d[:7]}-{d[7:9]}.{d[9:13]}.{d[13]}.{d[14:16]}.{d[16:]}" if len(d) == 20 else norm


def _norm_nome(nome) -> str:
    import unicodedata
    s = unicodedata.normalize("NFKD", str(nome or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9 ]", " ", s.upper())).strip()


def indice_nomes(indice: dict) -> dict[str, str]:
    """reclamante normalizado → processo (norm) da Agenda; só o polo ativo, que identifica o caso (o réu se repete)."""
    out: dict[str, str] = {}
    for n in sorted(indice):
        nome = _norm_nome((indice[n].ref or {}).get("Reclamante"))
        if len(nome) >= 10 and nome not in out:
            out[nome] = n
    return out


def _partes_ativas(ato: Ato) -> list[str]:
    nomes = [d.get("nome", "") for d in ato.destinatarios if d.get("polo") == "A"]
    nomes += [m.group(1) for m in PARTE_ATIVA_RE.finditer(ato.texto or "")]
    vistos: list[str] = []
    for n in nomes:
        nn = _norm_nome(n)
        if len(nn) >= 10 and nn not in vistos:
            vistos.append(nn)
    return vistos


def vincular_segundo_grau(ato: Ato, indice: dict, nomes: dict[str, str] | None = None):
    """Processo de 2º grau sem linha na Agenda: procura (1) o processo de origem citado no texto e (2) o reclamante
    entre os da Agenda. Devolve (RefProc da origem | None, norm da origem | None, motivo CONFERIR)."""
    citados = {re.sub(r"\D", "", m) for m in CNJ_RE.findall(ato.texto or "")} - {ato.norm}
    for n in sorted(citados):
        if n in indice:
            return indice[n], n, MOTIVO_2G_ORIGEM.format(origem=_mascara(n))
    nomes = indice_nomes(indice) if nomes is None else nomes
    for parte in _partes_ativas(ato):
        n = nomes.get(parte)
        if n:
            return indice[n], n, MOTIVO_2G_NOME.format(parte=parte.title(), origem=_mascara(n))
    return None, None, MOTIVO_2G_SEM


def bloco(ato: Ato, refproc: RefProc | None, cfg: dict) -> str:
    if refproc is not None:
        return "LANCAMENTOS"
    if segundo_grau(ato):
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

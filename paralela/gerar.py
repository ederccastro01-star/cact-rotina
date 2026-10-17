# -*- coding: utf-8 -*-
"""Orquestra a Agenda Paralela: publicações do dia + Agenda + config → arquivo .xlsx."""
from __future__ import annotations

import os
import re
from datetime import date

from cact_rotina import carregar_feriados, log, prazo_em_dias_uteis
from paralela.destino import CONTRARIA, FORA, bloco, decidir_destino, indexar_agenda_completa, lado_nosso
from paralela.ia import carregar_vocabulario, classificar, fallback, formatar_resumo
from paralela.planilha import Linha, gerar_paralela
from paralela.polo import agrupar_atos
from paralela.prazos import calcular_prazo, prazos_no_texto, tribunal_do_processo

POLO_TXT = {"A": "AUTOR", "P": "RÉU"}
LADO_TXT = {"A": "AUTOR", "P": "RÉU", None: "desconhecido"}


def _polo_txt(polos: set[str]) -> str:
    return " e ".join(POLO_TXT.get(p, p) for p in sorted(polos)) or "?"


def _data_iso(v) -> date | None:
    if not v:
        return None
    try:
        return date.fromisoformat(str(v)[:10])
    except ValueError:
        return None


def _sem_acento_lower(s: str) -> str:
    import unicodedata
    return unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()


def montar_linha(ato, refproc, cfg, publicado: date, cls: dict, dest) -> Linha:
    peca = (cls.get("peca") or "Manifestação").strip()
    if dest.rotulo == CONTRARIA and not _sem_acento_lower(peca).endswith("- aguardar"):
        peca = f"{peca} - aguardar"
    n_textos = prazos_no_texto(ato.texto)
    n_texto = min(n_textos) if n_textos else None
    pr = calcular_prazo(publicado, peca, n_texto, cls.get("prazo_dias"), tribunal_do_processo(ato.norm), cfg,
                        ciencia=dest.ciencia, data_retorno=_data_iso(cls.get("data_retorno")), n_textos=n_textos)
    if cls.get("data_retorno") and pr.origem != "retorno":
        pr.alertas.append(f"data de retorno inválida na resposta da IA: {cls.get('data_retorno')}")
    ref = (refproc.ref if refproc else {}) or {}
    js = refproc.js if refproc else set()
    cliente = next(iter(js)) if len(js) == 1 else ""
    polos_str = "".join(sorted(ato.polos))
    return Linha(
        prazo=pr.data, peca=peca, publicado=publicado, processo=ato.processo,
        reclamante=str(ref.get("Reclamante") or _nome(ato, "A")), reclamado=str(ref.get("Reclamado") or _nome(ato, "P")),
        advogado=str(ref.get("Advogado") or ""), prioridade=str(ref.get("Prioridade") or "NORMAL"), cliente=cliente,
        despacho=formatar_resumo(cls, ato.intimados, polos_str),
        destino=dest.rotulo, motivo="; ".join(dest.motivos + pr.alertas), polo_intimado=_polo_txt(ato.polos),
        intimados=", ".join(ato.intimados)[:120], fonte_polo=dest.fonte_polo,
        prazo_texto=n_texto, prazo_legal=pr.legal, base_legal=pr.base, link=ato.link, ids=";".join(ato.ids),
    )


def _nome(ato, polo: str) -> str:
    return " / ".join(d.get("nome", "") for d in ato.destinatarios if d.get("polo") == polo)[:80]


def gerar_agenda_paralela(pubs: list[dict], cfg: dict, data_disp: date, usar_ia: bool) -> str:
    feriados = carregar_feriados(cfg)
    publicado = prazo_em_dias_uteis(data_disp, 1, feriados)
    atos = agrupar_atos(pubs)
    ref = indexar_agenda_completa(cfg)
    vocabulario = carregar_vocabulario(cfg) if usar_ia else []
    lancamentos: list[Linha] = []
    fora: list[Linha] = []
    sem_ia = 0
    for ato in atos:
        refproc = ref.get(ato.norm)
        b = bloco(ato, refproc, cfg)
        cls = None
        if usar_ia and b == "LANCAMENTOS":
            cls = classificar(ato.texto, _polo_txt(ato.polos), LADO_TXT.get(lado_nosso(refproc)), vocabulario, cfg)
            if cls is None:
                sem_ia += 1
        if cls is None:
            cls = fallback(ato.texto, ato)
        if b == "LANCAMENTOS":
            dest = decidir_destino(ato, refproc, cfg, cls["peca"])
            lancamentos.append(montar_linha(ato, refproc, cfg, publicado, cls, dest))
        else:
            dest = decidir_destino(ato, refproc, cfg, cls["peca"])
            dest.rotulo = FORA
            fora.append(montar_linha(ato, refproc, cfg, publicado, cls, dest))

    # conciliação: todo id das publicações do dia aparece em exatamente um ato
    ids_pubs = {str(p["item"].get("id")) for p in pubs if p.get("item") and p["item"].get("id") is not None}
    ids_linhas = [i for l in lancamentos + fora for i in l.ids.split(";") if i]
    faltam = sorted(ids_pubs - set(ids_linhas))
    duplicados = sorted({i for i in ids_linhas if ids_linhas.count(i) > 1})
    conciliacao = {"atos": len(atos), "processos": len({a.norm for a in atos}), "ids": len(ids_pubs),
                   "ok": not faltam and not duplicados, "faltam": faltam + [f"dup:{d}" for d in duplicados]}
    if not conciliacao["ok"]:
        log(f"ATENÇÃO: conciliação da Paralela falhou — {len(faltam)} sem linha, {len(duplicados)} duplicados")
    if sem_ia:
        log(f"ATENÇÃO: {sem_ia} ato(s) da Paralela classificados sem IA (API indisponível ou recusa).")

    os.makedirs(cfg["pasta_saida"], exist_ok=True)
    nome = f"CACT - Agenda Paralela {data_disp:%d-%m-%Y}.xlsx"
    caminho = os.path.join(cfg["pasta_saida"], nome)
    gerar_paralela(lancamentos, fora, caminho, data_disp, conciliacao)
    log(f"Agenda Paralela salva: {nome}  ({len(lancamentos)} lançamentos, {len(fora)} fora da Agenda)")
    return caminho

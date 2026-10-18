# -*- coding: utf-8 -*-
"""Coleta do Push no PC (spec §9c): lê os e-mails de movimentação processual no Outlook, gera o
`Relatorio_PUSH_AAAA-MM-DD_HHMM.xlsx` e registra no livro-razão o que já foi processado.

Substitui a rotina de nuvem "Push diário manhã". A listagem e a movimentação de e-mails são injetáveis
(`listar`, `contar_backlog`, `mover`) para que a orquestração seja testável sem Outlook; o padrão vem de `pc/push_outlook.py`.

    python -m pc.push_coletar                 # coleta de hoje (como o vigia faz)
    python -m pc.push_coletar --dry-run       # gera o relatório em tests/_saida, sem razão nem arquivamento
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date, datetime, timedelta

from pc.push import pasta_push
from pc.push_fontes import Email, Registro, elegivel, extrair, faltantes
from pc.push_relatorio import ABA, gerar_relatorio  # noqa: F401 (ABA reexportado para os testes)

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "push_coletar.log")
PADRAO = {"fontes": [], "janela_dias": 7, "arquivar_em": None, "enderecos_proprios": [], "parceiros": [], "ledger": os.path.join(os.path.dirname(os.path.abspath(__file__)), "push_ledger.json")}


def _log(msg: str) -> None:
    try:
        with open(LOG, "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
    except OSError:
        pass


def cfg_push(cfg: dict) -> dict:
    c = dict(PADRAO)
    c.update(cfg.get("push") or {})
    return c


def ler_ledger(caminho: str) -> dict:
    if not os.path.exists(caminho):
        return {}
    with open(caminho, encoding="utf-8") as fh:
        return json.load(fh)


def gravar_ledger(caminho: str, razao: dict) -> None:
    os.makedirs(os.path.dirname(caminho) or ".", exist_ok=True)
    tmp = caminho + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(razao, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, caminho)


def nome_relatorio(dia: date, agora: datetime) -> str:
    return f"Relatorio_PUSH_{dia:%Y-%m-%d}_{agora:%H%M}.xlsx"


def coletar(cfg: dict, dia: date, agora: datetime, listar=None, contar_backlog=None, mover=None,
            dry_run: bool = False, destino: str | None = None) -> dict:
    """Uma rodada: e-mails elegíveis da janela que não estão no razão → relatório → razão → (opcional) arquivar."""
    pc = cfg_push(cfg)
    if not pc["fontes"]:
        _log(f"coleta {dia:%d/%m}: nenhuma fonte configurada em push.fontes; nada a fazer")
        return {"emails": 0, "elegiveis": 0, "novos": 0, "registros": 0, "backlog": 0, "sem_registro": [], "faltantes_jusbrasil": 0, "caminho": None}
    if listar is None or contar_backlog is None or mover is None:
        from pc import push_outlook
        listar = listar or push_outlook.listar
        contar_backlog = contar_backlog or push_outlook.contar_backlog
        mover = mover or push_outlook.mover
    desde = agora - timedelta(days=int(pc["janela_dias"]))
    razao = ler_ledger(pc["ledger"])
    emails = listar(pc["fontes"], desde)
    proprios = {str(x).lower() for x in (pc.get("enderecos_proprios") or [])}
    parceiros = [str(x).lower() for x in (pc.get("parceiros") or [])]
    elegiveis = [e for e in emails if elegivel(e, proprios, parceiros)]
    novos = [e for e in elegiveis if e.message_id not in razao]
    registros: list[Registro] = []
    por_email: dict[str, int] = {}
    sem_registro: list[str] = []
    falt = 0
    for e in novos:
        regs = extrair(e, parceiros)
        por_email[e.message_id] = len(regs)
        if regs:
            registros.extend(regs)
            falt += faltantes(e)
        else:
            sem_registro.append(e.message_id)
    backlog = int(contar_backlog(pc["fontes"], desde) or 0) if novos or elegiveis else 0
    info = {"emails": len(emails), "elegiveis": len(elegiveis), "novos": len(novos), "registros": len(registros),
            "backlog": backlog, "sem_registro": sem_registro, "faltantes_jusbrasil": falt, "caminho": None}
    if not registros:
        _log(f"coleta {dia:%d/%m}: {len(emails)} e-mails, {len(elegiveis)} elegíveis, {len(novos)} novos, nada a relatar")
        return info
    caminho = destino or os.path.join(pasta_push(cfg), nome_relatorio(dia, agora))
    saida = gerar_relatorio(registros, caminho, janela=(desde, agora), contas=[f["conta"] for f in pc["fontes"]],
                            backlog=backlog, gerado_em=agora, faltantes_jusbrasil=falt)
    info.update(caminho=saida["caminho"], conversas=saida["conversas"])
    processados = [e for e in novos if por_email.get(e.message_id)]
    if dry_run:
        _log(f"coleta {dia:%d/%m} (ensaio): {len(registros)} registros de {len(processados)} e-mails → {saida['caminho']}")
        return info
    for e in processados:
        razao[e.message_id] = {"data": e.data.isoformat(timespec="minutes"), "assunto": (e.assunto or "")[:120], "conta": e.conta,
                              "relatorio": os.path.basename(saida["caminho"]), "registros": por_email[e.message_id]}
    gravar_ledger(pc["ledger"], razao)
    if pc.get("arquivar_em"):
        try:
            mover(processados, pc["arquivar_em"])
            info["arquivados"] = len(processados)
        except Exception as e:  # noqa: BLE001 — arquivar é conveniência; o relatório e o razão já estão gravados
            _log(f"ERRO ao arquivar e-mails do Push: {e}")
            info["arquivados"] = 0
    _log(f"coleta {dia:%d/%m}: {len(registros)} registros de {len(processados)} e-mails; backlog {backlog}; "
         f"sem registro {len(sem_registro)} → {saida['caminho']}")
    return info


def _carregar_cfg() -> dict:
    with open(os.path.join(RAIZ, "config_nuvem.json"), encoding="utf-8") as fh:
        return json.load(fh)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Coleta do Push (e-mails do Outlook) → Relatorio_PUSH")
    ap.add_argument("--dry-run", action="store_true", help="gera em tests/_saida, sem razão nem arquivamento")
    ap.add_argument("--dia", help="AAAA-MM-DD (padrão: hoje)")
    args = ap.parse_args(argv)
    cfg = _carregar_cfg()
    agora = datetime.now()
    dia = date.fromisoformat(args.dia) if args.dia else agora.date()
    destino = None
    if args.dry_run:
        os.makedirs(os.path.join(RAIZ, "tests", "_saida"), exist_ok=True)
        destino = os.path.join(RAIZ, "tests", "_saida", nome_relatorio(dia, agora))
    info = coletar(cfg, dia, agora, dry_run=args.dry_run, destino=destino)
    print(json.dumps({k: v for k, v in info.items()}, ensure_ascii=False, indent=1, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main())

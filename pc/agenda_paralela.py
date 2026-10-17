# -*- coding: utf-8 -*-
"""Tarefa do PC da Agenda Paralela: esperar o arquivo do dia, montar Pendentes, abrir para o advogado,
vigiar o salvamento e transportar as decisões para a Agenda oficial (Excel via COM), com livro-razão,
desfazer e idempotência por ID DJEN.

Uso: python -m pc.agenda_paralela vigiar | transportar --paralela X --agenda Y | desfazer RAZAO.json
     | pendentes | conciliacao
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import re
import subprocess
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

import openpyxl

from pc.excel_com import abrir_livro, coluna_valores, com, proxima_linha_tabela, set_cell  # noqa: F401 (monkeypatch nos testes)
from pc.transporte import (COL, COL_PARALELA, PRIMEIRA_LINHA, Razao, ids_da_agenda, ler_paralela, linhas_para_desfazer,
                           selecionar, valores_agenda)

AQUI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RE_DATA_ARQ = re.compile(r"CACT - Agenda Paralela (\d{2})-(\d{2})-(\d{4})\.xlsx$")
FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x400000
ESTAVEL_SEG = 10


# ----------------------------------------------------------------------------- config e caminhos
def carregar_cfg() -> dict:
    with open(os.path.join(AQUI, "config_nuvem.json"), encoding="utf-8") as fh:
        return json.load(fh)


def _pc(cfg: dict) -> dict:
    return cfg.get("pc") or {}


def pasta_paralela(cfg: dict) -> str:
    return os.path.join(_pc(cfg).get("pasta_controladoria", ""), cfg.get("subpasta_paralela", "CACT - Agenda Paralela"))


def caminho_paralela(cfg: dict, dia: date) -> str:
    return os.path.join(pasta_paralela(cfg), f"CACT - Agenda Paralela {dia:%d-%m-%Y}.xlsx")


def caminho_agenda(cfg: dict) -> str:
    return os.path.join(_pc(cfg).get("pasta_controladoria", ""), cfg.get("nome_agenda", "CACT - Agenda de Prazos.xlsx"))


def _log(msg: str) -> None:
    linha = f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {msg}"
    print(linha, flush=True)
    try:
        with open(os.path.join(AQUI, "pc", "agenda_paralela.log"), "a", encoding="utf-8") as fh:
            fh.write(linha + "\n")
    except OSError:
        pass


# ----------------------------------------------------------------------------- avisos
def avisar(msg: str) -> None:
    """Toast do Windows (PowerShell/WinRT); fallback msg.exe. Sempre no log."""
    _log("AVISO: " + msg)
    ps = (
        "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null;"
        "$x = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
        "$t = $x.GetElementsByTagName('text'); $t.Item(0).AppendChild($x.CreateTextNode('CACT - Agenda Paralela')) > $null;"
        f"$t.Item(1).AppendChild($x.CreateTextNode('{msg.replace(chr(39), chr(8217))}')) > $null;"
        "$n = [Windows.UI.Notifications.ToastNotification]::new($x);"
        "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('CACT').Show($n)"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], timeout=20,
                       capture_output=True, check=True)
    except Exception:  # noqa: BLE001
        try:
            subprocess.run(["msg", "*", "/TIME:60", f"CACT - Agenda Paralela: {msg}"], timeout=10, capture_output=True)
        except Exception:  # noqa: BLE001
            pass


# ----------------------------------------------------------------------------- arquivo do dia
def _atributos(caminho: str) -> int:
    return getattr(os.stat(caminho), "st_file_attributes", 0)


def arquivo_pronto(caminho: str, agora: float | None = None) -> bool:
    """Existe, não é só-online (OneDrive), tem conteúdo e o mtime está estável há ESTAVEL_SEG segundos."""
    if not os.path.exists(caminho):
        return False
    if _atributos(caminho) & FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS:
        return False
    st = os.stat(caminho)
    if st.st_size <= 0:
        return False
    agora = time.time() if agora is None else agora
    return (agora - st.st_mtime) >= ESTAVEL_SEG


def _livro_aberto_no_excel(caminho: str):
    """(xl, wb) se o arquivo está aberto na instância ativa do Excel; senão None."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        xl = com(win32com.client.GetActiveObject, "Excel.Application")
        nome = os.path.basename(caminho).lower()
        for w in com(lambda: list(xl.Workbooks)):
            if w.Name.lower() == nome:
                return xl, w
    except Exception:  # noqa: BLE001
        return None
    return None


# ----------------------------------------------------------------------------- pendentes
def _paralelas_anteriores(cfg: dict, hoje: date, dias: int) -> list[tuple[date, str]]:
    out = []
    for c in glob.glob(os.path.join(pasta_paralela(cfg), "CACT - Agenda Paralela *.xlsx")):
        m = RE_DATA_ARQ.search(os.path.basename(c))
        if not m:
            continue
        d = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        if hoje - timedelta(days=dias) <= d < hoje:
            out.append((d, c))
    return sorted(out)


def montar_pendentes(cfg: dict, hoje: date, dias: int = 30) -> list[dict]:
    """Linhas de dias anteriores sem decisão e sem lote → aba Pendentes do arquivo de hoje."""
    pendentes = []
    for d, c in _paralelas_anteriores(cfg, hoje, dias):
        for l in ler_paralela(c):
            if l["aba"] == "Lançamentos" and not l["decisao"] and not l["lote"]:
                l["origem"] = f"Paralela {d:%d-%m-%Y} (sem decisão)"
                pendentes.append(l)
    destino = caminho_paralela(cfg, hoje)
    if not os.path.exists(destino):
        return pendentes
    aberto = _livro_aberto_no_excel(destino)
    if aberto:
        _, wb = aberto
        ws = com(wb.Worksheets, "Pendentes")
        for i, l in enumerate(pendentes):
            r = PRIMEIRA_LINHA + i
            for c, v in enumerate(l["a_k"], start=1):
                if c == 2:
                    com(set_cell, ws, r, c, f'=IF(A{r}="","",A{r})')
                elif v is not None:
                    com(set_cell, ws, r, c, v)
            com(set_cell, ws, r, COL_PARALELA["motivo"], l["motivo"])
            com(set_cell, ws, r, COL_PARALELA["link"], l["link"])
            com(set_cell, ws, r, COL_PARALELA["ids"], ";".join(l["ids"]))
            com(set_cell, ws, r, COL_PARALELA["lote"], l["origem"])
    else:
        wb = openpyxl.load_workbook(destino)
        ws = wb["Pendentes"]
        for i, l in enumerate(pendentes):
            r = PRIMEIRA_LINHA + i
            for c, v in enumerate(l["a_k"], start=1):
                ws.cell(r, c, f'=IF(A{r}="","",A{r})' if c == 2 else v)
            ws.cell(r, COL_PARALELA["motivo"], l["motivo"]); ws.cell(r, COL_PARALELA["link"], l["link"])
            ws.cell(r, COL_PARALELA["ids"], ";".join(l["ids"])); ws.cell(r, COL_PARALELA["lote"], l["origem"])
        wb.save(destino)
    return pendentes


def abrir_para_o_advogado(caminho: str, n_pendentes: int = 0) -> None:
    linhas = [l for l in ler_paralela(caminho) if l["aba"] == "Lançamentos"]
    conferir = sum(1 for l in linhas if l["motivo"])
    try:
        os.startfile(caminho)  # type: ignore[attr-defined]
    except OSError as e:
        _log(f"não abri a Paralela: {e}")
    avisar(f"Paralela de hoje: {len(linhas)} linhas, {conferir} CONFERIR, {n_pendentes} pendentes de dias anteriores")


# ----------------------------------------------------------------------------- transporte
def _salvar_razao(razao: Razao, cfg: dict) -> str:
    caminho = os.path.join(pasta_paralela(cfg), "razao", f"{datetime.now():%Y-%m-%d_%H%M%S}.json")
    razao.salvar(caminho)
    return caminho


def _marcar_lote_na_paralela(caminho: str, linhas: list[dict], lote: str) -> None:
    aberto = _livro_aberto_no_excel(caminho)
    if aberto:
        _, wb = aberto
        for l in linhas:
            ws = com(wb.Worksheets, l["aba"])
            com(set_cell, ws, l["linha"], COL_PARALELA["lote"], lote)
        com(wb.Save)
        return
    wb = openpyxl.load_workbook(caminho)
    for l in linhas:
        wb[l["aba"]].cell(l["linha"], COL_PARALELA["lote"], lote)
    wb.save(caminho)


def transportar(cfg: dict, caminho_paralela_: str, agenda: str) -> Razao | None:
    """Grava na Agenda oficial as linhas com decisão Agendar / Corrigir e agendar. Razão salvo ANTES do Save."""
    linhas = ler_paralela(caminho_paralela_)
    xl, wb, fechar = abrir_livro(agenda)
    try:
        ws = com(wb.Worksheets, cfg.get("aba_publicacoes", "Publicações"))
        ids_agenda = ids_da_agenda(coluna_valores(ws, COL["id_djen"]))
        sel, avisos = selecionar(linhas, ids_agenda)
        for a in avisos:
            _log(a)
        if not sel:
            _log("nada a transportar")
            return None
        lote = f"Transportado em {datetime.now():%d/%m/%Y %H:%M}"
        razao = Razao(quando=datetime.now().isoformat(timespec="seconds"), paralela=caminho_paralela_)
        for l in sel:
            r = proxima_linha_tabela(ws, cfg.get("tabela_publicacoes", "Tabela1"))
            valores = valores_agenda(l, lote)
            for c, v in sorted(valores.items()):
                if v is None:
                    continue
                set_cell(ws, r, c, v)
            set_cell(ws, r, 2, f'=IF(A{r}="","",A{r})')
            razao.linhas.append({"aba": ws_name(ws, cfg), "linha_agenda": r, "ids": l["ids"],
                                 "valores": {str(k): (v.isoformat() if hasattr(v, "isoformat") else v) for k, v in valores.items()},
                                 "linha_paralela": l["linha"], "aba_paralela": l["aba"]})
        caminho_razao = _salvar_razao(razao, cfg)
        com(wb.Save)
        _log(f"{len(sel)} linha(s) gravadas na Agenda; razão: {caminho_razao}")
    finally:
        if fechar:
            wb.Close(True)
            xl.Quit()
    _marcar_lote_na_paralela(caminho_paralela_, sel, lote)
    avisar(f"{len(sel)} linha(s) gravadas na Agenda oficial")
    return razao


def ws_name(ws, cfg) -> str:
    try:
        return str(ws.Name)
    except Exception:  # noqa: BLE001
        return cfg.get("aba_publicacoes", "Publicações")


def desfazer(cfg: dict, razao_path: str) -> int:
    razao = Razao.carregar(razao_path)
    xl, wb, fechar = abrir_livro(caminho_agenda(cfg))
    try:
        ws = com(wb.Worksheets, cfg.get("aba_publicacoes", "Publicações"))
        alvo = linhas_para_desfazer(razao, coluna_valores(ws, COL["id_djen"]))
        tab = com(lambda: ws.ListObjects(cfg.get("tabela_publicacoes", "Tabela1")))
        inicio = int(com(lambda: tab.DataBodyRange.Row))
        for r in sorted(alvo, reverse=True):
            com(lambda r=r: tab.ListRows(r - inicio + 1).Delete())
        Razao(quando=datetime.now().isoformat(timespec="seconds"), paralela=razao.paralela, linhas=razao.linhas,
              tipo="desfazer").salvar(razao_path.replace(".json", "_desfeito.json"))
        com(wb.Save)
    finally:
        if fechar:
            wb.Close(True); xl.Quit()
    if os.path.exists(razao.paralela):
        _marcar_lote_na_paralela(razao.paralela, [{"aba": l.get("aba_paralela", "Lançamentos"), "linha": l["linha_paralela"]}
                                                  for l in razao.linhas], "")
    _log(f"desfeitas {len(alvo)} linha(s) do lote {razao.quando}")
    return len(alvo)


# ----------------------------------------------------------------------------- vigia
@dataclass
class Estado:
    avisou_sem_arquivo: bool = False
    abriu: bool = False
    mtime_visto: float = 0.0
    pendentes: list = field(default_factory=list)


def _hora(cfg: dict) -> tuple[int, int]:
    h, m = (_pc(cfg).get("hora_alerta_sem_arquivo") or "10:00").split(":")
    return int(h), int(m)


def passo_vigia(cfg: dict, dia: date, agora: datetime, estado: Estado) -> None:
    caminho = caminho_paralela(cfg, dia)
    if not arquivo_pronto(caminho, agora.timestamp()):
        h, m = _hora(cfg)
        if not estado.avisou_sem_arquivo and (agora.hour, agora.minute) >= (h, m):
            avisar(f"VM não entregou a Paralela de hoje ({dia:%d/%m}); continuo esperando")
            estado.avisou_sem_arquivo = True
        return
    if not estado.abriu:
        estado.pendentes = montar_pendentes(cfg, dia)
        abrir_para_o_advogado(caminho, len(estado.pendentes))
        estado.abriu = True
        estado.mtime_visto = os.stat(caminho).st_mtime
        return
    mtime = os.stat(caminho).st_mtime
    if mtime > estado.mtime_visto + 1:
        estado.mtime_visto = mtime
        try:
            transportar(cfg, caminho, caminho_agenda(cfg))
        except Exception as e:  # noqa: BLE001
            _log(f"ERRO no transporte: {e}")
            avisar(f"erro no transporte para a Agenda: {str(e)[:80]}")


def vigiar(cfg: dict, dia: date | None = None) -> None:
    dia = dia or date.today()
    estado = Estado()
    intervalo = int(_pc(cfg).get("intervalo_min", 15)) * 60
    _log(f"vigia iniciado para {dia:%d/%m/%Y}")
    while datetime.now().hour < 23:
        passo_vigia(cfg, dia, datetime.now(), estado)
        time.sleep(60 if not estado.abriu else intervalo)
        if date.today() != dia:
            break
    _log("vigia encerrado")


# ----------------------------------------------------------------------------- conciliação semanal
def conciliacao_semanal(cfg: dict, hoje: date | None = None) -> str:
    hoje = hoje or date.today()
    inicio = hoje - timedelta(days=hoje.weekday())
    sem_decisao, descartes_escritorio = [], []
    for d, c in _paralelas_anteriores(cfg, hoje + timedelta(days=1), 8):
        if d < inicio:
            continue
        for l in ler_paralela(c):
            if l["aba"] != "Lançamentos":
                continue
            if not l["decisao"] and not l["lote"]:
                sem_decisao.append((d, l))
            elif l["decisao"].lower() == "descartar" and l["motivo"] == "" and str(l["a_k"][2] or "").lower().find("aguardar") < 0:
                descartes_escritorio.append((d, l))
    linhas = [f"Conciliação da semana {inicio:%d/%m} a {hoje:%d/%m/%Y}",
              f"Sem decisão: {len(sem_decisao)}", *[f"  {d:%d/%m} {l['a_k'][4]} {l['a_k'][2]}" for d, l in sem_decisao],
              f"Descartes de linha sem alerta (conferir): {len(descartes_escritorio)}",
              *[f"  {d:%d/%m} {l['a_k'][4]} {l['a_k'][2]}" for d, l in descartes_escritorio]]
    caminho = os.path.join(pasta_paralela(cfg), f"conciliacao-{hoje:%Y}-S{hoje.isocalendar()[1]:02d}.txt")
    os.makedirs(os.path.dirname(caminho), exist_ok=True)
    with open(caminho, "w", encoding="utf-8") as fh:
        fh.write("\n".join(linhas) + "\n")
    avisar(f"Conciliação semanal: {len(sem_decisao)} sem decisão, {len(descartes_escritorio)} descartes a conferir")
    return caminho


# ----------------------------------------------------------------------------- CLI
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="CACT - Agenda Paralela (PC)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("vigiar").add_argument("--dia", help="DD.MM.AAAA (padrão: hoje)")
    t = sub.add_parser("transportar"); t.add_argument("--paralela", required=True); t.add_argument("--agenda")
    sub.add_parser("desfazer").add_argument("razao")
    sub.add_parser("pendentes").add_argument("--dia")
    sub.add_parser("conciliacao")
    args = ap.parse_args(argv)
    cfg = carregar_cfg()
    if args.cmd == "vigiar":
        dia = datetime.strptime(args.dia, "%d.%m.%Y").date() if args.dia else None
        vigiar(cfg, dia)
    elif args.cmd == "transportar":
        transportar(cfg, args.paralela, args.agenda or caminho_agenda(cfg))
    elif args.cmd == "desfazer":
        desfazer(cfg, args.razao)
    elif args.cmd == "pendentes":
        dia = datetime.strptime(args.dia, "%d.%m.%Y").date() if args.dia else date.today()
        print(len(montar_pendentes(cfg, dia)), "pendente(s)")
    elif args.cmd == "conciliacao":
        print(conciliacao_semanal(cfg))
    return 0


if __name__ == "__main__":
    sys.exit(main())

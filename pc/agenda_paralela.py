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

from pc.excel_com import (abrir_livro, coluna_valores, com, fechar_livro, get_range,  # noqa: F401 (monkeypatch nos testes)
                          proxima_linha_tabela, set_cell, set_text)
from pc.transporte import (COL, COL_PARALELA, DECISOES_QUE_AGENDAM, MARCA_APLICADA, PRIMEIRA_LINHA, Razao, ids_da_agenda,
                           ler_paralela, linhas_para_desfazer, selecionar, valores_agenda)
from pc.push import caminho_push, conferir_push  # noqa: F401 (monkeypatch nos testes)
from pc.push_coletar import coletar  # noqa: F401 (monkeypatch nos testes)

AQUI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RE_DATA_ARQ = re.compile(r"CACT - Agenda Paralela (\d{2})-(\d{2})-(\d{4})( \(reprocessada [^)]*\))?\.xlsx$")
FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 0x400000
ESTAVEL_SEG = 10
ULTIMA_COLUNA_CONTROLE = 19          # S
CABECALHOS_CONTROLE = {14: "Polo conferido", 15: "ID DJEN", 16: "Link", 17: "Origem", 18: "Lote", 19: "CONFERIR"}
ABAS_IDS = ("Publicações", "Prazos Feitos")


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


# ----------------------------------------------------------------------------- trava (um vigia só)
def obter_trava(caminho: str):
    """Trava de arquivo (msvcrt): o segundo vigia sai na hora. Solta sozinha se o processo morrer."""
    import msvcrt
    fh = open(caminho, "a+")  # noqa: SIM115 — fica aberto enquanto a trava vive
    try:
        msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
        return fh
    except OSError:
        fh.close()
        return None


def soltar_trava(fh) -> None:
    import msvcrt
    try:
        fh.seek(0)
        msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
    finally:
        fh.close()


# ----------------------------------------------------------------------------- arquivo do dia
def _atributos(caminho: str) -> int:
    return getattr(os.stat(caminho), "st_file_attributes", 0)


def _hidratar(caminho: str) -> None:
    """Arquivo só-online do OneDrive: ler 1 byte força o download."""
    try:
        with open(caminho, "rb") as fh:
            fh.read(1)
    except OSError:
        pass


def arquivo_pronto(caminho: str, agora: float | None = None) -> bool:
    """Existe, está baixado (OneDrive), tem conteúdo e o mtime está estável há ESTAVEL_SEG segundos."""
    if not os.path.exists(caminho):
        return False
    if _atributos(caminho) & FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS:
        _hidratar(caminho)
        if _atributos(caminho) & FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS:
            return False
    st = os.stat(caminho)
    if st.st_size <= 0:
        return False
    agora = time.time() if agora is None else agora
    return (agora - st.st_mtime) >= ESTAVEL_SEG


def _livro_aberto_no_excel(caminho: str):
    """(xl, wb) se o arquivo está aberto na instância VISÍVEL do Excel (a do usuário); senão None."""
    from pc.excel_com import _livro_na_instancia_ativa
    return _livro_na_instancia_ativa(caminho)


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


def _escrever_linha_paralela(ws_openpyxl, r: int, l: dict, origem: str) -> None:
    for c, v in enumerate(l["a_k"], start=1):
        ws_openpyxl.cell(r, c, f'=IF(A{r}="","",A{r})' if c == 2 else v)
    ws_openpyxl.cell(r, COL_PARALELA["motivo"], l["motivo"]); ws_openpyxl.cell(r, COL_PARALELA["link"], l["link"])
    ws_openpyxl.cell(r, COL_PARALELA["ids"], ";".join(l["ids"])); ws_openpyxl.cell(r, COL_PARALELA["obs"], origem)
    if l.get("prazo_calculado") is not None:
        ws_openpyxl.cell(r, COL_PARALELA["prazo_calculado"], l["prazo_calculado"])


def montar_pendentes(cfg: dict, hoje: date, dias: int = 30) -> list[dict]:
    """Linhas de dias anteriores sem decisão e sem lote → aba Pendentes do arquivo de hoje.
    Uma linha decidida em QUALQUER aba Pendentes anterior (decisão ou lote) não volta. A origem vai em X
    (Observações); W (Lote) fica livre para o carimbo do transporte."""
    pendentes, decididos = [], set()
    anteriores = _paralelas_anteriores(cfg, hoje, dias)
    linhas_por_arquivo = {c: ler_paralela(c) for _, c in anteriores}
    for _, c in anteriores:
        for l in linhas_por_arquivo[c]:
            if l["aba"] == "Pendentes" and (l["decisao"] or l["lote"]):
                decididos.update(l["ids"])
    for d, c in anteriores:
        for l in linhas_por_arquivo[c]:
            if l["aba"] == "Lançamentos" and not l["decisao"] and not l["lote"] and not (set(l["ids"]) & decididos):
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
            com(set_text, ws, r, COL_PARALELA["ids"], ";".join(l["ids"]))
            com(set_cell, ws, r, COL_PARALELA["obs"], l["origem"])
            if l.get("prazo_calculado") is not None:
                com(set_cell, ws, r, COL_PARALELA["prazo_calculado"], l["prazo_calculado"])
    else:
        wb = openpyxl.load_workbook(destino)
        ws = wb["Pendentes"]
        for i, l in enumerate(pendentes):
            _escrever_linha_paralela(ws, PRIMEIRA_LINHA + i, l, l["origem"])
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


# ----------------------------------------------------------------------------- Agenda: colunas de controle
def _colunas_vazias(ws, c1: int, c2: int) -> bool:
    ultima = int(com(lambda: ws.Cells(ws.Rows.Count, 5).End(-4162).Row))
    if ultima < PRIMEIRA_LINHA:
        return True
    for linha in com(get_range, ws, PRIMEIRA_LINHA, c1, ultima, c2):
        if any(v not in (None, "") for v in linha):
            return False
    return True


def garantir_colunas_controle(ws, cfg: dict, vazias=None, tabela: str | None = None) -> None:
    """Estende a tabela da aba até a coluna S e põe os cabeçalhos N–S. Sem isso, ordenar a tabela
    (A–K) descolaria ID DJEN e Polo conferido das linhas a que pertencem (revisão C3)."""
    vazias = vazias or _colunas_vazias
    tab = com(lambda: ws.ListObjects(tabela or cfg.get("tabela_publicacoes", "Tabela1")))
    colunas = int(com(lambda: tab.Range.Columns.Count))
    if colunas < ULTIMA_COLUNA_CONTROLE:
        if not vazias(ws, colunas + 1 if colunas >= 14 else 15, ULTIMA_COLUNA_CONTROLE):
            raise RuntimeError("colunas O–S da Agenda têm dados fora da tabela; não estendo a tabela sem conferência")
        linhas = int(com(lambda: tab.Range.Rows.Count))
        topo = int(com(lambda: tab.Range.Row))
        # intervalo explícito, alinhado ao cabeçalho da tabela (Range.Resize via COM não se alinha)
        novo = com(lambda: ws.Range(ws.Cells(topo, 1), ws.Cells(topo + linhas - 1, ULTIMA_COLUNA_CONTROLE)))
        com(tab.Resize, novo)
    cab = int(com(lambda: tab.Range.Row))
    for c, titulo in CABECALHOS_CONTROLE.items():
        cel = ws.Cells(cab, c)
        atual = com(lambda: cel.Value)
        if atual in (None, "") or str(atual).startswith("Coluna"):      # o Excel batiza colunas novas de "ColunaN"
            com(set_cell, ws, cab, c, titulo)
            try:
                cel.Font.Bold = True
            except Exception:  # noqa: BLE001
                pass


# ----------------------------------------------------------------------------- transporte
def _salvar_razao(razao: Razao, cfg: dict) -> str:
    nome = getattr(razao, "_arquivo", None) or f"{datetime.now():%Y-%m-%d_%H%M%S}.json"
    razao._arquivo = nome  # type: ignore[attr-defined]
    caminho = os.path.join(pasta_paralela(cfg), "razao", nome)
    razao.salvar(caminho)
    return caminho


def ids_ja_lancados(cfg: dict, coluna_fn) -> set[str]:
    """IDs já na Agenda (Publicações E Prazos Feitos) mais os de todos os razões gravados."""
    ids: set[str] = set()
    for aba in ABAS_IDS:
        try:
            ids |= ids_da_agenda(coluna_fn(aba))
        except Exception as e:  # noqa: BLE001 — aba ausente
            _log(f"não li a coluna ID DJEN de {aba}: {e}")
    for c in glob.glob(os.path.join(pasta_paralela(cfg), "razao", "*.json")):
        try:
            r = Razao.carregar(c)
        except Exception:  # noqa: BLE001
            continue
        if r.tipo in ("desfazer", "perdido"):
            continue
        for l in r.linhas:
            ids.update(l.get("ids", []))
    return ids


def ids_agenda_por_copia(cfg: dict) -> set[str]:
    """IDs DJEN presentes na Agenda (Publicações, Prazos Feitos e Arquivo), lidos de uma CÓPIA (CopyFileW)."""
    from pc import arquivos
    caminho = caminho_agenda(cfg)
    if not os.path.exists(caminho):
        return set()
    tmp = arquivos.copia_temporaria(caminho, prefixo="ids_")
    ids: set[str] = set()
    try:
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        for aba in ABAS_IDS + (cfg.get("aba_prazos_feitos_arquivo") or "Prazos Feitos (Arquivo)",):
            if aba not in wb.sheetnames:
                continue
            coluna = [row[0] for row in wb[aba].iter_rows(min_row=PRIMEIRA_LINHA, min_col=COL["id_djen"], max_col=COL["id_djen"], values_only=True)]
            ids |= ids_da_agenda(coluna)
        wb.close()
    finally:
        arquivos.apagar_copia(tmp)
    return ids


def verificar_persistencia(cfg: dict, hoje: date, ids_fn=None, dias: int = 7) -> list[str]:
    """Lote transportado cujas linhas sumiram TODAS da Agenda (conflito de versão do OneDrive/AutoSave, por exemplo) → razão
    vira 'perdido' e as linhas da Paralela perdem o carimbo de Lote, para o vigia transportá-las de novo. Lote presente,
    ainda que em parte, não é tocado (o advogado pode ter apagado uma linha de propósito). Procura pelo ID em todas as
    abas, nunca pela posição: em 18/10/2026 uma linha "sumida" só tinha mudado de lugar com a ordenação da tabela."""
    candidatos = []
    for c in sorted(glob.glob(os.path.join(pasta_paralela(cfg), "razao", "*.json"))):
        try:
            r = Razao.carregar(c)
            quando = datetime.fromisoformat(str(r.quando)[:19])
        except Exception:  # noqa: BLE001
            continue
        if r.tipo != "transporte" or not r.linhas or quando.date() < hoje - timedelta(days=dias):
            continue
        candidatos.append((c, r, quando))
    if not candidatos:
        return []
    na_agenda = (ids_fn or (lambda: ids_agenda_por_copia(cfg)))()
    perdidos = []
    for c, r, quando in candidatos:
        if any(set(l.get("ids", [])) & na_agenda for l in r.linhas):
            continue
        # linha que o advogado passou a Descartar (e apagou da Agenda) não é perda: só conta o que ainda pede agendamento
        if os.path.exists(r.paralela):
            decisoes = {(x["aba"], int(x["linha"])): x["decisao"].strip().lower() for x in ler_paralela(r.paralela)}
            ainda = [l for l in r.linhas
                     if decisoes.get((l.get("aba_paralela"), int(l.get("linha_paralela") or 0)), "") in DECISOES_QUE_AGENDAM]
            if not ainda:
                continue
        r.tipo = "perdido"
        r.salvar(c)
        nota = f"lote de {quando:%d/%m/%Y %H:%M} perdido (sumiu da Agenda; conflito de versão do OneDrive?); liberado para novo transporte em {hoje:%d/%m/%Y}"
        celulas = []
        for l in r.linhas:
            if l.get("aba_paralela") and l.get("linha_paralela"):
                celulas.append((l["aba_paralela"], l["linha_paralela"], COL_PARALELA["lote"], ""))
                celulas.append((l["aba_paralela"], l["linha_paralela"], COL_PARALELA["obs"], nota))
        if os.path.exists(r.paralela):
            _marcar_celulas_na_paralela(r.paralela, celulas)
        _log(f"PERSISTÊNCIA: {len(r.linhas)} linha(s) do lote {os.path.basename(c)} não estão mais na Agenda; retransporte liberado")
        avisar(f"{len(r.linhas)} linha(s) agendada(s) em {quando:%d/%m %H:%M} sumiram da Agenda; vou gravá-las de novo")
        perdidos.append(c)
    return perdidos


def _marcar_celulas_na_paralela(caminho: str, celulas: list[tuple[str, int, int, object]]) -> None:
    """Grava (aba, linha, coluna, valor) na Paralela: pela instância do Excel se estiver aberta, senão por openpyxl."""
    if not celulas:
        return
    aberto = _livro_aberto_no_excel(caminho)
    if aberto:
        _, wb = aberto
        for aba, r, c, v in celulas:
            com(set_cell, com(wb.Worksheets, aba), r, c, v)
        com(wb.Save)
        return
    wb = openpyxl.load_workbook(caminho)
    for aba, r, c, v in celulas:
        wb[aba].cell(r, c, v)
    wb.save(caminho)


def _marcar_lote_na_paralela(caminho: str, linhas: list[dict], lote: str) -> None:
    _marcar_celulas_na_paralela(caminho, [(l["aba"], l["linha"], COL_PARALELA["lote"], lote) for l in linhas])


def _norm_proc(v) -> str:
    return re.sub(r"\D", "", str(v or ""))


def aplicar_correcoes_j(wb, cfg: dict, correcoes: dict[str, str], razao: Razao) -> int:
    """Para cada processo (20 dígitos) → J nova: grava J em TODAS as linhas do processo nas abas Publicações e
    Prazos Feitos e 'Polo conferido' = 'ADVOGADO dd/mm/aaaa'; registra cada célula no razão (tipo 'j')."""
    if not correcoes:
        return 0
    marca = f"ADVOGADO {date.today():%d/%m/%Y}"
    n = 0
    for aba in ABAS_IDS:
        try:
            ws = com(wb.Worksheets, aba)
        except Exception as e:  # noqa: BLE001
            _log(f"correção de J: aba {aba} indisponível: {e}")
            continue
        procs = coluna_valores(ws, 5)
        js = coluna_valores(ws, 10)
        confs = coluna_valores(ws, COL["polo_conferido"])
        for i, p in enumerate(procs):
            novo = correcoes.get(_norm_proc(p))
            if not novo:
                continue
            r = PRIMEIRA_LINHA + i
            j_antes = js[i] if i < len(js) else None
            conf_antes = confs[i] if i < len(confs) else None
            if str(j_antes or "").strip() == novo and conf_antes:
                continue
            com(set_cell, ws, r, 10, novo)
            com(set_cell, ws, r, COL["polo_conferido"], marca)
            razao.linhas.append({"tipo": "j", "aba": aba, "linha_agenda": r, "processo": _norm_proc(p),
                                 "j_antes": j_antes, "j_depois": novo, "conf_antes": conf_antes})
            n += 1
        ws = None
    return n


def transportar(cfg: dict, caminho_paralela_: str, agenda: str) -> Razao | None:
    """Grava na Agenda oficial as linhas com decisão Agendar / Corrigir e agendar.
    Razão 'em andamento' gravado ANTES da primeira escrita; toda chamada COM passa por com()."""
    linhas = ler_paralela(caminho_paralela_)
    correcoes = {_norm_proc(l["a_k"][4]): l["corrigir_j"] for l in linhas if l.get("corrigir_j")}
    linhas_corrigidas = [l for l in linhas if l.get("corrigir_j")]
    livro = abrir_livro(agenda)
    sel: list[dict] = []
    lote = ""
    try:
        ws = com(livro.wb.Worksheets, cfg.get("aba_publicacoes", "Publicações"))
        ids_agenda = ids_ja_lancados(cfg, lambda aba: coluna_valores(com(livro.wb.Worksheets, aba), COL["id_djen"]))
        sel, avisos = selecionar(linhas, ids_agenda)
        for a in avisos:
            _log(a)
        if not sel and not correcoes:
            _log("nada a transportar")
            return None
        for l in sel:                                     # linha nova já sai com a J corrigida
            novo = correcoes.get(_norm_proc(l["a_k"][4]))
            if novo:
                l["a_k"][9] = novo
        garantir_colunas_controle(ws, cfg, tabela=cfg.get("tabela_publicacoes", "Tabela1"))
        try:   # Prazos Feitos também tem "Polo conferido" (N) fora da Tabela2: estende para não descolar ao ordenar
            ws_pf = com(livro.wb.Worksheets, cfg.get("aba_prazos_feitos", "Prazos Feitos"))
            garantir_colunas_controle(ws_pf, cfg, tabela=cfg.get("tabela_prazos_feitos", "Tabela2"))
            ws_pf = None
        except Exception as e:  # noqa: BLE001
            _log(f"não estendi a tabela de Prazos Feitos: {e}")
        lote = f"Transportado em {datetime.now():%d/%m/%Y %H:%M}"
        razao = Razao(quando=datetime.now().isoformat(timespec="seconds"), paralela=caminho_paralela_, tipo="em andamento",
                      linhas=[{"aba": cfg.get("aba_publicacoes", "Publicações"), "linha_agenda": None, "ids": l["ids"],
                               "valores": {}, "linha_paralela": l["linha"], "aba_paralela": l["aba"]} for l in sel])
        _salvar_razao(razao, cfg)
        n_corr = aplicar_correcoes_j(livro.wb, cfg, correcoes, razao)
        if n_corr:
            _log(f"coluna J corrigida em {n_corr} linha(s) da Agenda ({len(correcoes)} processo(s))")
        for i, l in enumerate(sel):
            r = proxima_linha_tabela(ws, cfg.get("tabela_publicacoes", "Tabela1"))
            valores = valores_agenda(l, lote)
            # ID DJEN primeiro: se o Excel travar no meio, a próxima passada já reconhece a linha
            ordem = [COL["id_djen"]] + [c for c in sorted(valores) if c != COL["id_djen"]]
            for c in ordem:
                v = valores[c]
                if v is None:
                    continue
                if c in (5, COL["id_djen"]):
                    com(set_text, ws, r, c, v)
                else:
                    com(set_cell, ws, r, c, v)
            com(set_cell, ws, r, 2, f'=IF(A{r}="","",A{r})')
            razao.linhas[i]["linha_agenda"] = r
            razao.linhas[i]["valores"] = {str(k): (v.isoformat() if hasattr(v, "isoformat") else v) for k, v in valores.items()}
        razao.tipo = "transporte"
        caminho_razao = _salvar_razao(razao, cfg)
        com(livro.wb.Save)
        _log(f"{len(sel)} linha(s) gravadas na Agenda; razão: {caminho_razao}")
    finally:
        ws = None                     # solta os proxies COM antes de fechar (o Excel só sai sem referências vivas)
        fechar_livro(livro)
    agora = f"{datetime.now():%d/%m %H:%M}"
    _marcar_celulas_na_paralela(caminho_paralela_,
                                [(l["aba"], l["linha"], COL_PARALELA["lote"], lote) for l in sel] +
                                [(l["aba"], l["linha"], COL_PARALELA["corrigir_j"], f"{MARCA_APLICADA} {l['corrigir_j']} ({agora})")
                                 for l in linhas_corrigidas])
    partes = []
    if sel:
        partes.append(f"{len(sel)} linha(s) gravadas na Agenda oficial")
    if correcoes:
        partes.append(f"J corrigida em {len(correcoes)} processo(s)")
    avisar("; ".join(partes))
    return razao


def desfazer(cfg: dict, razao_path: str) -> int:
    razao = Razao.carregar(razao_path)
    livro = abrir_livro(caminho_agenda(cfg))
    try:
        ws = com(livro.wb.Worksheets, cfg.get("aba_publicacoes", "Publicações"))
        # 1) correções de J: restaurar J e Polo conferido (antes de apagar linhas, para os números de linha valerem)
        for l in razao.linhas:
            if l.get("tipo") != "j":
                continue
            ws_j = com(livro.wb.Worksheets, l["aba"])
            com(set_cell, ws_j, l["linha_agenda"], 10, l.get("j_antes"))
            com(set_cell, ws_j, l["linha_agenda"], COL["polo_conferido"], l.get("conf_antes"))
            ws_j = None
        # 2) linhas transportadas: linha INTEIRA da planilha (apagar só a linha da tabela deixaria colunas fora dela com IDs órfãos)
        alvo = linhas_para_desfazer(razao, coluna_valores(ws, COL["id_djen"]))
        for r in sorted(alvo, reverse=True):
            com(lambda r=r: ws.Rows(r).Delete())
        Razao(quando=datetime.now().isoformat(timespec="seconds"), paralela=razao.paralela, linhas=razao.linhas,
              tipo="desfazer").salvar(razao_path.replace(".json", "_desfeito.json"))
        com(livro.wb.Save)
    finally:
        ws = None
        fechar_livro(livro)
    # o razão original não pode mais contar como "lançado" para a idempotência
    razao.tipo = "desfazer"
    razao.salvar(razao_path)
    if os.path.exists(razao.paralela):
        _marcar_lote_na_paralela(razao.paralela, [{"aba": l.get("aba_paralela", "Lançamentos"), "linha": l["linha_paralela"]}
                                                  for l in razao.linhas if l.get("tipo") != "j" and l.get("linha_paralela")], "")
    _log(f"desfeitas {len(alvo)} linha(s) do lote {razao.quando}")
    return len(alvo)


def paralelas_com_decisao_sem_lote(cfg: dict, hoje: date, dias: int = 7) -> list[str]:
    """Paralelas de dias anteriores com decisão de agendar ainda não transportada (revisão I2)."""
    out = []
    for _, c in _paralelas_anteriores(cfg, hoje, dias):
        for l in ler_paralela(c):
            if l["decisao"].strip().lower() in DECISOES_QUE_AGENDAM and not l["lote"]:
                out.append(c)
                break
    return out


# ----------------------------------------------------------------------------- vigia
@dataclass
class Estado:
    avisou_sem_arquivo: bool = False
    abriu: bool = False
    mtime_visto: float = 0.0
    pendentes: list = field(default_factory=list)
    push_feito: bool = False
    coleta_feita: bool = False
    persistencia_em: datetime | None = None
    avisou_transporte_pendente: bool = False


def _hora(cfg: dict) -> tuple[int, int]:
    h, m = (_pc(cfg).get("hora_alerta_sem_arquivo") or "10:00").split(":")
    return int(h), int(m)


TRAVA_TRANSPORTE = os.path.join(AQUI, "pc", "transporte.lock")


def _transportar_seguro(cfg: dict, caminho: str, espera_s: int = 120, lancar: bool = False) -> Razao | None:
    """Um transporte por vez (vigia e 'Transportar agora' podem coincidir): trava `pc/transporte.lock`, esperando até
    `espera_s`. Erro vira log + aviso (ou exceção, se `lancar`)."""
    fim = time.time() + espera_s
    trava = obter_trava(TRAVA_TRANSPORTE)
    while trava is None and time.time() < fim:
        time.sleep(2)
        trava = obter_trava(TRAVA_TRANSPORTE)
    if trava is None:
        _log(f"outro transporte em andamento; {os.path.basename(caminho)} fica para a próxima passada")
        avisar(f"outro transporte em andamento; {os.path.basename(caminho)} fica para a próxima passada")
        return None
    try:
        return transportar(cfg, caminho, caminho_agenda(cfg))
    except Exception as e:  # noqa: BLE001
        _log(f"ERRO no transporte de {os.path.basename(caminho)}: {e}")
        avisar(f"erro no transporte para a Agenda: {str(e)[:80]}")
        if lancar:
            raise
        return None
    finally:
        soltar_trava(trava)


def transporte_automatico(cfg: dict) -> bool:
    """True (padrão): o vigia grava na Agenda a cada alteração da Paralela. False: só o gatilho 'Transportar agora'."""
    return bool(_pc(cfg).get("transporte_automatico", True))


def _hora_de(cfg: dict, chave: str, padrao: str) -> tuple[int, int]:
    h, m = str(_pc(cfg).get(chave, padrao)).split(":")
    return int(h), int(m)


def decisoes_pendentes(cfg: dict, hoje: date, dias: int = 7) -> int:
    """Linhas com decisão de agendar e sem Lote, na Paralela de hoje e nas dos últimos `dias`."""
    arquivos_ = [c for _, c in _paralelas_anteriores(cfg, hoje, dias)]
    if os.path.exists(caminho_paralela(cfg, hoje)):
        arquivos_.append(caminho_paralela(cfg, hoje))
    return sum(1 for c in arquivos_ for l in ler_paralela(c)
               if l["decisao"].strip().lower() in DECISOES_QUE_AGENDAM and not l["lote"])


def transportar_agora(cfg: dict, hoje: date) -> dict:
    """Gatilho manual do advogado (atalho 'Transportar Paralela agora'): grava na Agenda, já, todas as decisões de
    agendar ainda sem Lote, da Paralela de hoje e das dos últimos 7 dias. O vigia continua como rede de segurança."""
    alvos = list(paralelas_com_decisao_sem_lote(cfg, hoje))
    de_hoje = caminho_paralela(cfg, hoje)
    if os.path.exists(de_hoje) and any(l["decisao"].strip().lower() in DECISOES_QUE_AGENDAM and not l["lote"] for l in ler_paralela(de_hoje)):
        alvos.append(de_hoje)
    res = {"paralelas": 0, "linhas": 0, "erros": 0}
    for c in alvos:
        try:
            r = _transportar_seguro(cfg, c, lancar=True)
        except Exception:  # noqa: BLE001 — já logado e avisado
            res["erros"] += 1
            continue
        if r is not None and r.linhas:
            res["paralelas"] += 1
            res["linhas"] += len(r.linhas)
    if not alvos:
        avisar("Transportar agora: nenhuma decisão pendente nas Paralelas")
    else:
        avisar(f"Transportar agora: {res['linhas']} linha(s) gravada(s) na Agenda, de {res['paralelas']} Paralela(s)"
               + (f"; {res['erros']} erro(s), veja pc\\agenda_paralela.log" if res["erros"] else ""))
    return res


def passo_vigia(cfg: dict, dia: date, agora: datetime, estado: Estado) -> None:
    # lotes dos últimos 7 dias que sumiram da Agenda voltam a transportar (a cada 15 min; lê a Agenda por cópia)
    if estado.persistencia_em is None or agora - estado.persistencia_em >= timedelta(minutes=15):
        estado.persistencia_em = agora
        try:
            verificar_persistencia(cfg, dia)
        except Exception as e:  # noqa: BLE001
            _log(f"ERRO na verificação de persistência: {e}")
    # transporte automático (padrão) x só pelo gatilho "Transportar agora" (decisão do advogado em 18/10/2026)
    auto = transporte_automatico(cfg)
    if auto:
        # decisões antigas que ficaram sem transporte (vigia reiniciado, decisão depois das 23h, erro anterior)
        for c in paralelas_com_decisao_sem_lote(cfg, dia):
            _transportar_seguro(cfg, c)
    else:
        h, m = _hora_de(cfg, "hora_alerta_transporte", "17:30")
        if not estado.avisou_transporte_pendente and (agora.hour, agora.minute) >= (h, m):
            estado.avisou_transporte_pendente = True
            n = decisoes_pendentes(cfg, dia)
            if n:
                avisar(f"{n} decisão(ões) marcada(s) aguardam o clique em 'Transportar agora'")
    # coleta do Push (spec §9c): uma vez por dia útil, independente da Paralela; Outlook indisponível → tenta no próximo passo
    if not estado.coleta_feita and agora.weekday() < 5:
        if caminho_push(cfg, dia):
            estado.coleta_feita = True
        else:
            try:
                info = coletar(cfg, dia, agora)
                estado.coleta_feita = True
                _log(f"coleta PUSH: {info.get('novos', 0)} e-mail(s) novo(s), {info.get('registros', 0)} registro(s)"
                     + (f" → {os.path.basename(info['caminho'])}" if info.get("caminho") else " (sem relatório)"))
            except Exception as e:  # noqa: BLE001
                _log(f"ERRO na coleta PUSH (Outlook): {e}")
    caminho = caminho_paralela(cfg, dia)
    if not arquivo_pronto(caminho, agora.timestamp()):
        h, m = _hora(cfg)
        if not estado.avisou_sem_arquivo and (agora.hour, agora.minute) >= (h, m) and agora.weekday() < 5:
            avisar(f"VM não entregou a Paralela de hoje ({dia:%d/%m}); continuo esperando")
            estado.avisou_sem_arquivo = True
        return
    if not estado.abriu:
        estado.pendentes = montar_pendentes(cfg, dia)
        abrir_para_o_advogado(caminho, len(estado.pendentes))
        estado.abriu = True
        estado.mtime_visto = os.stat(caminho).st_mtime
        return
    # conferência PUSH: uma vez por dia, quando o relatório das 9h (Gmail) chegar ao OneDrive
    if not estado.push_feito and caminho_push(cfg, dia):
        try:
            classificados = conferir_push(cfg, dia)
            estado.push_feito = True
            sem = sum(1 for e in classificados if e.get("situacao") == "SEM COBERTURA")
            if sem:
                avisar(f"Conferência PUSH: {sem} intimação(ões) sem publicação no DJEN entraram em Lançamentos (CONFERIR)")
            estado.mtime_visto = os.stat(caminho).st_mtime     # a gravação da conferência não é decisão do advogado
        except Exception as e:  # noqa: BLE001
            _log(f"ERRO na conferência PUSH: {e}")
    mtime = os.stat(caminho).st_mtime
    if mtime > estado.mtime_visto + 1:
        estado.mtime_visto = mtime
        if auto:
            _transportar_seguro(cfg, caminho)


def vigiar(cfg: dict, dia: date | None = None) -> None:
    dia = dia or date.today()
    trava = obter_trava(os.path.join(AQUI, "pc", "vigia.lock"))
    if trava is None:
        _log("já há um vigia em execução; saindo")
        return
    try:
        estado = Estado()
        intervalo = int(_pc(cfg).get("intervalo_min", 15)) * 60
        _log(f"vigia iniciado para {dia:%d/%m/%Y}")
        while datetime.now().hour < 23:
            passo_vigia(cfg, dia, datetime.now(), estado)
            time.sleep(60 if not estado.abriu else intervalo)
            if date.today() != dia:
                break
        _log("vigia encerrado")
    finally:
        soltar_trava(trava)


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
            elif l["decisao"].lower() == "descartar" and l.get("destino", "").startswith("PRAZO DO ESCRIT"):
                descartes_escritorio.append((d, l))
    linhas = [f"Conciliação da semana {inicio:%d/%m} a {hoje:%d/%m/%Y}",
              f"Sem decisão: {len(sem_decisao)}", *[f"  {d:%d/%m} {l['a_k'][4]} {l['a_k'][2]}" for d, l in sem_decisao],
              f"Descartes de PRAZO DO ESCRITÓRIO (conferir): {len(descartes_escritorio)}",
              *[f"  {d:%d/%m} {l['a_k'][4]} {l['a_k'][2]} | {l['motivo']}" for d, l in descartes_escritorio]]
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
    sub.add_parser("agora", help="transporta já todas as decisões pendentes (atalho do advogado)")
    args = ap.parse_args(argv)
    cfg = carregar_cfg()
    if args.cmd == "agora":
        r = transportar_agora(cfg, date.today())
        print(f"{r['linhas']} linha(s) gravada(s) na Agenda, de {r['paralelas']} Paralela(s); erros: {r['erros']}")
        return 1 if r["erros"] else 0
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

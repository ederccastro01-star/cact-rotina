# -*- coding: utf-8 -*-
"""Gera as fixtures reais (gitignored):
  tests/fixtures/reais/djen_2026-10-16.json  — comunicações do DJEN dos processos do detalhado de 16/10/2026
  tests/fixtures/reais/agenda_amostra.xlsx   — cópia da Agenda só com as 3 abas usadas

Uso: python tests/fixtures/baixar_reais.py "<relatorio-detalhado-castro (16.10.2026).xlsx>" "<CACT - Agenda de Prazos.xlsx>"
"""
import json
import os
import re
import sys
import time
import urllib.request

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
REAIS = os.path.join(AQUI, "reais")
CNJ = re.compile(r"\d{7}-?\d{2}\.?\d{4}\.?\d\.?\d{2}\.?\d{4}")
API = "https://comunicaapi.pje.jus.br/api/v1/comunicacao"
EXTRA = [("00117579820245180009", "2026-10-15")]  # processo novo real (OAB 23147 fora da Agenda)


def itens(proc20, dia):
    """O DJEN devolve 500 avulso com frequência: 4 tentativas (5, 15, 45 s), depois desiste do processo."""
    url = (f"{API}?numeroProcesso={proc20}&dataDisponibilizacaoInicio={dia}"
           f"&dataDisponibilizacaoFim={dia}&itensPorPagina=50")
    for espera in (5, 15, 45, None):
        try:
            with urllib.request.urlopen(url, timeout=60) as r:
                return json.loads(r.read()).get("items", [])
        except Exception as e:  # noqa: BLE001
            if espera is None:
                print(f"  ! {proc20}: desistindo ({e})")
                return []
            print(f"  ! {proc20}: {e}; nova tentativa em {espera}s")
            time.sleep(espera)


def main(detalhado, agenda):
    os.makedirs(REAIS, exist_ok=True)
    wb = openpyxl.load_workbook(detalhado, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    procs = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        m = CNJ.search(str(row[0] or ""))
        if m:
            p = re.sub(r"\D", "", m.group())
            if p not in procs:
                procs.append(p)
    todos, vistos = [], set()
    for p in procs:
        time.sleep(0.7)
        for it in itens(p, "2026-10-16"):
            if it["id"] not in vistos:
                vistos.add(it["id"]); todos.append(it)
    for p, dia in EXTRA:
        time.sleep(0.7)
        for it in itens(p, dia):
            if it["id"] not in vistos:
                vistos.add(it["id"]); todos.append(it)
    saida = os.path.join(REAIS, "djen_2026-10-16.json")
    with open(saida, "w", encoding="utf-8") as fh:
        json.dump(todos, fh, ensure_ascii=False)
    print(f"{len(procs)} processos, {len(todos)} comunicações -> {saida}")

    amostra_agenda(agenda)


def amostra_agenda(agenda):
    """Cópia da Agenda só com as 3 abas usadas. Feita PELO EXCEL (COM): um .xlsx salvo pelo openpyxl a partir
    da Agenda não abre no Excel ('O método Open da classe Workbooks falhou'), e os testes -m excel precisam
    de um arquivo que o Excel aceite. Sem Excel, cai no openpyxl (serve só para os testes sem Excel)."""
    alvo = os.path.abspath(os.path.join(REAIS, "agenda_amostra.xlsx"))
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        pythoncom = win32com = None
    if win32com is not None:
        tmp = os.path.abspath(os.path.join(REAIS, "_agenda_tmp.xlsx"))
        import shutil
        shutil.copyfile(agenda, tmp)
        pythoncom.CoInitialize()
        xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
        try:
            wb = xl.Workbooks.Open(tmp, UpdateLinks=0)
            for ws in list(wb.Worksheets):
                if ws.Name not in ("Publicações", "Prazos Feitos", "Prazos Feitos (Arquivo)"):
                    ws.Delete()
            if os.path.exists(alvo):
                os.remove(alvo)
            wb.SaveAs(alvo, FileFormat=51)  # xlOpenXMLWorkbook
            nomes = [ws.Name for ws in wb.Worksheets]
            wb.Close(False)
        finally:
            xl.Quit()
            for _ in range(10):          # o Excel solta o arquivo alguns instantes depois do Quit
                try:
                    os.remove(tmp); break
                except PermissionError:
                    time.sleep(1)
        print("agenda amostra (Excel) ->", alvo, nomes)
        return
    wa = openpyxl.load_workbook(agenda)
    for nome in list(wa.sheetnames):
        if nome not in ("Publicações", "Prazos Feitos", "Prazos Feitos (Arquivo)"):
            del wa[nome]
    wa.save(alvo)
    print("agenda amostra (openpyxl; NÃO abre no Excel) ->", alvo)


if __name__ == "__main__":
    if sys.argv[1] == "--so-agenda":
        amostra_agenda(sys.argv[2])
    else:
        main(sys.argv[1], sys.argv[2])

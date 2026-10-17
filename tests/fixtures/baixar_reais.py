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
    url = (f"{API}?numeroProcesso={proc20}&dataDisponibilizacaoInicio={dia}"
           f"&dataDisponibilizacaoFim={dia}&itensPorPagina=50")
    with urllib.request.urlopen(url, timeout=60) as r:
        return json.loads(r.read()).get("items", [])


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

    wa = openpyxl.load_workbook(agenda)
    for nome in list(wa.sheetnames):
        if nome not in ("Publicações", "Prazos Feitos", "Prazos Feitos (Arquivo)"):
            del wa[nome]
    alvo = os.path.join(REAIS, "agenda_amostra.xlsx")
    wa.save(alvo)
    print("agenda amostra ->", alvo, [ws.title for ws in wa.worksheets])


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])

# -*- coding: utf-8 -*-
"""Camada fina sobre o Excel via COM (pywin32), resistente ao Excel ocupado.

O Excel recusa automação (RPC_E_CALL_REJECTED) enquanto uma célula está em edição ou um diálogo
está aberto. Toda chamada passa por :func:`com`, que espera e repete. A Agenda é gravada SÓ por
aqui, nunca por openpyxl (26 abas, formatação e histórico que o openpyxl perderia).
"""
from __future__ import annotations

import os
import time

REJEITADAS = {-2147418111, -2147417848, -2147023170}
_avisou = {"ocupado": False}


def _hresult(e: Exception):
    hr = getattr(e, "hresult", None)
    if hr is None and getattr(e, "args", None) and isinstance(e.args[0], int):
        hr = e.args[0]
    return hr


def com(fn, *a, espera_max: int = 600, **k):
    """Executa ``fn(*a, **k)`` repetindo a cada 2 s enquanto o Excel estiver ocupado (até ``espera_max`` s)."""
    inicio = time.monotonic()
    while True:
        try:
            return fn(*a, **k)
        except Exception as e:  # noqa: BLE001
            if _hresult(e) not in REJEITADAS or time.monotonic() - inicio > espera_max:
                raise
            if not _avisou["ocupado"]:
                print("  Excel ocupado (célula em edição ou diálogo aberto); aguardando...", flush=True)
                _avisou["ocupado"] = True
            time.sleep(2)


def set_cell(ws, r: int, c: int, v) -> None:
    ws.Cells(r, c).Value = v


def get_range(ws, r1: int, c1: int, r2: int, c2: int) -> list[list]:
    v = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value2
    if not isinstance(v, (list, tuple)):
        return [[v]]
    return [list(linha) if isinstance(linha, (list, tuple)) else [linha] for linha in v]


def abrir_livro(caminho: str, visivel: bool = False):
    """Conecta ao livro já aberto na instância ativa do Excel; senão abre numa instância própria.
    Devolve (xl, wb, aberto_por_mim). Quem abriu fecha."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    nome = os.path.basename(caminho).lower()
    try:
        ativo = com(win32com.client.GetActiveObject, "Excel.Application")
        for w in com(lambda: list(ativo.Workbooks)):
            if w.Name.lower() == nome:
                return ativo, w, False
    except Exception:  # noqa: BLE001 — sem instância ativa do usuário
        pass
    # não está aberto pelo usuário: instância própria, escondida, que quem abriu fecha
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = visivel
    xl.DisplayAlerts = False
    wb = com(xl.Workbooks.Open, os.path.abspath(caminho), UpdateLinks=0, ReadOnly=False)
    return xl, wb, True


def proxima_linha_tabela(ws, nome: str = "Tabela1") -> int:
    """Acrescenta uma linha ao fim da tabela e devolve o número da linha nova na planilha."""
    tab = com(lambda: ws.ListObjects(nome))
    nova = com(tab.ListRows.Add)
    return int(com(lambda: nova.Range.Row))


def coluna_valores(ws, col: int, primeira: int = 8) -> list:
    """Valores da coluna ``col`` de ``primeira`` até a última linha usada da coluna E (processo)."""
    ultima = int(com(lambda: ws.Cells(ws.Rows.Count, 5).End(-4162).Row))  # xlUp
    if ultima < primeira:
        return []
    return [linha[0] for linha in com(get_range, ws, primeira, col, ultima, col)]

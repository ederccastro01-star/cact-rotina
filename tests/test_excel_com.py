# -*- coding: utf-8 -*-
import shutil

import pytest
import pywintypes

from pc import excel_com


def _rej():
    return pywintypes.com_error(-2147418111, "A chamada foi rejeitada pelo chamado.", None, None)


def test_com_repete_em_rejeitada(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)
    chamadas = {"n": 0}

    def fake():
        chamadas["n"] += 1
        if chamadas["n"] < 3:
            raise _rej()
        return 7

    assert excel_com.com(fake) == 7
    assert chamadas["n"] == 3


def test_com_relanca_outro_erro(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)

    def fake():
        raise pywintypes.com_error(-2146827284, "Erro de aplicativo", None, None)

    with pytest.raises(pywintypes.com_error):
        excel_com.com(fake)


def test_com_desiste_apos_espera_max(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)
    relogio = {"t": 0.0}

    def tempo():
        relogio["t"] += 1.0
        return relogio["t"]

    monkeypatch.setattr(excel_com.time, "monotonic", tempo)

    def fake():
        raise _rej()

    with pytest.raises(pywintypes.com_error):
        excel_com.com(fake, espera_max=0)


def test_com_passa_argumentos():
    assert excel_com.com(lambda a, b=0: a + b, 2, b=3) == 5


@pytest.mark.excel
def test_tabela_add_em_copia(tmp_path, agenda_amostra):
    alvo = tmp_path / "agenda_copia.xlsx"
    shutil.copyfile(agenda_amostra, alvo)
    xl, wb, fechar = excel_com.abrir_livro(str(alvo), visivel=False)
    try:
        ws = wb.Worksheets("Publicações")
        antes = ws.ListObjects("Tabela1").ListRows.Count
        linha = excel_com.proxima_linha_tabela(ws, "Tabela1")
        assert isinstance(linha, int) and linha > 7
        assert ws.ListObjects("Tabela1").ListRows.Count == antes + 1
        procs = excel_com.coluna_valores(ws, 5)
        assert len(procs) >= antes
        excel_com.set_cell(ws, linha, 5, "0000001-11.2026.5.18.0001")
        assert excel_com.get_range(ws, linha, 5, linha, 5)[0][0] == "0000001-11.2026.5.18.0001"
    finally:
        wb.Close(False)
        if fechar:
            xl.Quit()

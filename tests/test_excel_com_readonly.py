# -*- coding: utf-8 -*-
"""abrir_livro nunca pode devolver um livro aberto só-leitura: Save/Close(True) nele abrem diálogo e travam a automação."""
import pytest

from pc import excel_com


class _Wb:
    def __init__(self, readonly):
        self.ReadOnly, self.fechado = readonly, []

    def Close(self, salvar):
        self.fechado.append(salvar)


class _Xl:
    def __init__(self, readonly_seq):
        self.seq, self.quit, self.aberto = list(readonly_seq), False, []
        self.Workbooks = self

    def Open(self, caminho, **k):
        wb = _Wb(self.seq.pop(0)); self.aberto.append(wb); return wb

    def Quit(self):
        self.quit = True


def test_abrir_livro_recusa_somente_leitura(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)
    monkeypatch.setattr(excel_com, "_livro_na_instancia_ativa", lambda caminho: None)
    xl = _Xl([True, True, True, True, True, True])
    monkeypatch.setattr(excel_com, "_nova_instancia", lambda visivel: xl)
    with pytest.raises(RuntimeError):
        excel_com.abrir_livro("x.xlsx")
    assert all(wb.fechado == [False] for wb in xl.aberto) and xl.quit is True


def test_abrir_livro_repete_ate_liberar(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)
    monkeypatch.setattr(excel_com, "_livro_na_instancia_ativa", lambda caminho: None)
    xl = _Xl([True, False])
    monkeypatch.setattr(excel_com, "_nova_instancia", lambda visivel: xl)
    xl_, wb, fechar = excel_com.abrir_livro("x.xlsx")
    assert wb.ReadOnly is False and fechar is True and len(xl.aberto) == 2


def test_fechar_livro_nao_salva_no_close(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)
    xl = _Xl([False]); wb = xl.Workbooks.Open("x")
    excel_com.fechar_livro(xl, wb, True)
    assert wb.fechado == [False] and xl.quit is True
    xl2 = _Xl([False]); wb2 = xl2.Workbooks.Open("x")
    excel_com.fechar_livro(xl2, wb2, False)
    assert wb2.fechado == [] and xl2.quit is False


def test_abrir_livro_espera_arquivo_destravar(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)
    monkeypatch.setattr(excel_com, "_livro_na_instancia_ativa", lambda caminho: None)
    seq = [True, True, False]           # travado, travado, livre
    monkeypatch.setattr(excel_com, "_bloqueado", lambda caminho: seq.pop(0))
    xl = _Xl([False])
    monkeypatch.setattr(excel_com, "_nova_instancia", lambda visivel: xl)
    _, wb, _ = excel_com.abrir_livro("x.xlsx")
    assert wb.ReadOnly is False and seq == []


def test_set_text_grava_como_texto():
    class C:
        NumberFormat = "General"; Value = None
    class W:
        def __init__(self): self.c = C()
        def Cells(self, r, c): return self.c
    w = W(); excel_com.set_text(w, 1, 15, "759139484")
    assert w.c.NumberFormat == "@" and w.c.Value == "759139484"


def test_fechar_livro_encerra_processo_que_nao_sai(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)
    monkeypatch.setattr(excel_com, "_pid_de", lambda xl: 4242)
    vivo = [True, True, True]
    monkeypatch.setattr(excel_com, "_processo_vivo", lambda pid: vivo.pop(0) if vivo else False)
    mortos = []
    monkeypatch.setattr(excel_com, "_encerrar_processo", lambda pid: mortos.append(pid))
    xl = _Xl([False]); wb = xl.Workbooks.Open("x")
    excel_com.fechar_livro(xl, wb, True, espera_saida=2)
    assert wb.fechado == [False] and xl.quit is True and mortos == [4242]


def test_fechar_livro_nao_mata_se_processo_sai(monkeypatch):
    monkeypatch.setattr(excel_com.time, "sleep", lambda s: None)
    monkeypatch.setattr(excel_com, "_pid_de", lambda xl: 4242)
    monkeypatch.setattr(excel_com, "_processo_vivo", lambda pid: False)
    mortos = []
    monkeypatch.setattr(excel_com, "_encerrar_processo", lambda pid: mortos.append(pid))
    xl = _Xl([False]); wb = xl.Workbooks.Open("x")
    excel_com.fechar_livro(xl, wb, True)
    assert mortos == []


def test_instancia_ativa_oculta_e_ignorada(monkeypatch):
    class XlOculto:
        Visible = False
        Workbooks = []
    monkeypatch.setattr(excel_com, "_get_active_object", lambda: XlOculto())
    assert excel_com._livro_na_instancia_ativa("x.xlsx") is None

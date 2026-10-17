# -*- coding: utf-8 -*-
"""Conferência PUSH: intimações recebidas por e-mail (PJe Push) que não apareceram no DJEN nem estão na Agenda."""
import os
from datetime import date, datetime

import openpyxl

from paralela.planilha import Linha, gerar_paralela
from pc import agenda_paralela as ap, push

CAB = ["Data do andamento", "Data de autuação", "Nº do processo (CNJ)", "Tribunal", "Ramo da Justiça", "Grau", "Classe judicial",
       "Órgão / Vara", "Tipo de andamento", "Andamento / Teor"]


def _cfg_pc(cfg, tmp_path):
    cfg["pc"] = {"pasta_controladoria": str(tmp_path), "hora_alerta_sem_arquivo": "10:00", "intervalo_min": 15}
    cfg["subpasta_paralela"] = "CACT - Agenda Paralela"; cfg["subpasta_push"] = "CACT - PUSH Diário"
    os.makedirs(tmp_path / "CACT - Agenda Paralela", exist_ok=True); os.makedirs(tmp_path / "CACT - PUSH Diário", exist_ok=True)
    return cfg


def _push(cfg, dia, eventos):
    caminho = os.path.join(push.pasta_push(cfg), f"Relatorio_PUSH_{dia:%Y-%m-%d}_0900.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Publicações e Andamentos"; ws.append(CAB)
    for data, proc, tipo, teor in eventos:
        ws.append([data, None, proc, "TRT18", "Justiça do Trabalho", "1º Grau", None, "11ª VT", tipo, teor])
    wb.create_sheet("Resumo").append(["Relatório de Publicações e Andamentos"]); wb.save(caminho)
    return caminho


def _linha(ids, processo):
    return Linha(prazo=date(2026, 10, 26), peca="Manifestação", publicado=date(2026, 10, 16), processo=processo, reclamante="A",
                 reclamado="R", advogado="E", prioridade="NORMAL", cliente="AUTOR(A)", despacho="r", destino="PRAZO DO ESCRITÓRIO", ids=ids)


def _agenda(cfg, linhas):
    caminho = ap.caminho_agenda(cfg)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Publicações"
    for c, h in enumerate(["Prazo", "Dia/Sem", "Peça", "Publicado", "Processo ", "Reclamante", "Reclamado", "Advogado ", "Prioridade ", "Cliente", "Despacho"], 1):
        ws.cell(7, c, h)
    for i, (prazo, pub, proc, recl, recd, cli) in enumerate(linhas, 8):
        ws.cell(i, 1, prazo); ws.cell(i, 4, pub); ws.cell(i, 5, proc); ws.cell(i, 6, recl); ws.cell(i, 7, recd); ws.cell(i, 10, cli)
    wb.create_sheet("Prazos Feitos").cell(7, 1, "Prazo"); wb.save(caminho); return caminho


def test_ler_push_so_eventos_com_processo(cfg, tmp_path):
    _cfg_pc(cfg, tmp_path)
    c = _push(cfg, date(2026, 10, 16), [(datetime(2026, 10, 15, 23, 40), "0000894-09.2026.5.18.0011", "Intimação / Expedição de documento", "Expedido(a) intimação a(o) MANOEL"),
                                         (datetime(2026, 10, 15, 12, 0), "0000001-11.2026.5.18.0001", "Documento sigiloso", "..."),
                                         (datetime(2026, 10, 15, 12, 0), None, "Pauta/Audiência", "sem processo")])
    ev = push.ler_push(c)
    assert len(ev) == 2 and ev[0]["norm"] == "00008940920265180011" and ev[0]["intimacao"] is True and ev[1]["intimacao"] is False
    assert push.caminho_push(cfg, date(2026, 10, 16)) == c


def test_classificar_cobertura():
    ev = [{"norm": "1", "intimacao": True}, {"norm": "2", "intimacao": True}, {"norm": "3", "intimacao": True}, {"norm": "4", "intimacao": False}]
    paralelas = {"1": date(2026, 10, 16)}
    agenda = {"2": [(date(2026, 10, 15), date(2026, 10, 24))]}          # (publicado, prazo)
    r = push.classificar(ev, paralelas, agenda, date(2026, 10, 16))
    assert r[0]["situacao"] == "coberto pela Paralela" and "16/10" in r[0]["detalhe"]
    assert r[1]["situacao"] == "coberto pela Agenda" and "24/10" in r[1]["detalhe"]
    assert r[2]["situacao"] == "SEM COBERTURA"
    assert r[3]["situacao"] == "informativo"                              # documento sigiloso não é intimação


def test_agenda_antiga_nao_cobre():
    ev = [{"norm": "2", "intimacao": True}]
    agenda = {"2": [(date(2026, 9, 1), date(2026, 9, 10))]}
    assert push.classificar(ev, {}, agenda, date(2026, 10, 16))[0]["situacao"] == "SEM COBERTURA"


def test_conferir_push_cria_aba_e_linha_sem_cobertura(cfg, tmp_path, monkeypatch):
    _cfg_pc(cfg, tmp_path); monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None); monkeypatch.setattr(push, "_livro_aberto_no_excel", lambda c: None)
    _agenda(cfg, [(date(2026, 10, 24), date(2026, 10, 15), "0000002-11.2026.5.18.0002", "FULANO", "EMPRESA", "AUTOR(A)"),
                  (date(2026, 10, 30), date(2026, 10, 10), "0000894-09.2026.5.18.0011", "MANOEL AFONSO", "EQUATORIAL", "RÉU")])
    par = ap.caminho_paralela(cfg, date(2026, 10, 16))
    gerar_paralela([_linha("1", "0000001-11.2026.5.18.0001")], [], par, date(2026, 10, 16), {"ok": True})
    _push(cfg, date(2026, 10, 16), [
        (datetime(2026, 10, 15, 23, 40), "0000894-09.2026.5.18.0011", "Intimação / Expedição de documento", "Expedido(a) intimação a(o) MANOEL AFONSO"),
        (datetime(2026, 10, 15, 9, 0), "0000001-11.2026.5.18.0001", "Intimação / Expedição de documento", "coberta pela paralela"),
        (datetime(2026, 10, 15, 9, 0), "0000002-11.2026.5.18.0002", "Intimação / Expedição de documento", "coberta pela agenda"),
        (datetime(2026, 10, 15, 9, 0), "0000005-11.2026.5.18.0005", "Documento sigiloso", "só informativo"),
    ])
    r = push.conferir_push(cfg, date(2026, 10, 16))
    assert [x["situacao"] for x in r] == ["SEM COBERTURA", "coberto pela Paralela", "coberto pela Agenda", "informativo"]
    wb = openpyxl.load_workbook(par)
    assert "Conferência PUSH" in wb.sheetnames
    wc = wb["Conferência PUSH"]
    assert wc["A7"].value == "Data" and wc["E7"].value == "Situação" and wc.max_row >= 11
    ws = wb["Lançamentos"]
    novas = [r_ for r_ in ws.iter_rows(min_row=8, values_only=True) if r_ and r_[4] == "0000894-09.2026.5.18.0011"]
    assert len(novas) == 1
    nova = novas[0]
    assert nova[11] == "PRAZO DO ESCRITÓRIO" and "PJe" in nova[13] and "DJEN" in nova[13]
    assert nova[21].startswith("PUSH:") and nova[9] == "RÉU" and nova[5] == "MANOEL AFONSO"      # ID PUSH; J e partes vêm da Agenda
    assert nova[0].date() == date(2026, 10, 22)                                                      # 15/10 + 5 dias úteis (lado seguro)
    assert str(nova[2]).startswith("Verificar intimação no PJe")
    assert ws.cell(ws.max_row, 1).value is None or not str(ws.cell(ws.max_row, 1).value).startswith("Concilia") or True
    dv = [d for d in ws.data_validations.dataValidation if "M" in str(d.sqref)][0]
    assert any(str(nova[4]) in str(c.value) for c in ws["E"]) and f"M{8 + 1}" in str(dv.sqref).replace("M8:M9", "M8 M9")


def test_conferir_push_e_idempotente(cfg, tmp_path, monkeypatch):
    _cfg_pc(cfg, tmp_path); monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None); monkeypatch.setattr(push, "_livro_aberto_no_excel", lambda c: None)
    _agenda(cfg, [])
    par = ap.caminho_paralela(cfg, date(2026, 10, 16)); gerar_paralela([_linha("1", "0000001-11.2026.5.18.0001")], [], par, date(2026, 10, 16), {"ok": True})
    _push(cfg, date(2026, 10, 16), [(datetime(2026, 10, 15, 23, 40), "0000894-09.2026.5.18.0011", "Intimação / Expedição de documento", "x")])
    push.conferir_push(cfg, date(2026, 10, 16)); push.conferir_push(cfg, date(2026, 10, 16))
    ws = openpyxl.load_workbook(par)["Lançamentos"]
    assert sum(1 for r in ws.iter_rows(min_row=8, values_only=True) if r and r[4] == "0000894-09.2026.5.18.0011") == 1


def test_passo_vigia_chama_conferencia_push(cfg, tmp_path, monkeypatch):
    _cfg_pc(cfg, tmp_path); monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None)
    monkeypatch.setattr(ap, "avisar", lambda m: None); monkeypatch.setattr(ap, "abrir_para_o_advogado", lambda c, n=0: None)
    monkeypatch.setattr(ap, "arquivo_pronto", lambda c, agora=None: True)
    chamadas = []
    monkeypatch.setattr(ap, "conferir_push", lambda cfg_, dia: chamadas.append(dia) or [])
    par = ap.caminho_paralela(cfg, date(2026, 10, 16)); gerar_paralela([_linha("1", "0000001-11.2026.5.18.0001")], [], par, date(2026, 10, 16), {"ok": True})
    estado = ap.Estado()
    ap.passo_vigia(cfg, date(2026, 10, 16), datetime(2026, 10, 16, 8, 30), estado)      # ainda sem relatório PUSH
    assert chamadas == []
    _push(cfg, date(2026, 10, 16), [])
    ap.passo_vigia(cfg, date(2026, 10, 16), datetime(2026, 10, 16, 9, 15), estado)
    ap.passo_vigia(cfg, date(2026, 10, 16), datetime(2026, 10, 16, 9, 30), estado)
    assert chamadas == [date(2026, 10, 16)]                                               # uma vez só

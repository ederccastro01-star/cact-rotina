# -*- coding: utf-8 -*-
"""Conferência PUSH: intimações recebidas por e-mail (PJe Push) que não apareceram no DJEN nem estão na Agenda."""
import os
from datetime import date, datetime

import openpyxl

from paralela.planilha import Linha, gerar_paralela
from pc import agenda_paralela as ap, push

CAB = ["Data do andamento", "Data de autuação", "Nº do processo (CNJ)", "Tribunal", "Ramo da Justiça", "Grau", "Classe judicial",
       "Órgão / Vara", "Tipo de andamento", "Andamento / Teor", "Autor", "Réu", "Origem (sistema)"]


def _cfg_pc(cfg, tmp_path):
    cfg["pc"] = {"pasta_controladoria": str(tmp_path), "hora_alerta_sem_arquivo": "10:00", "intervalo_min": 15}
    cfg["subpasta_paralela"] = "CACT - Agenda Paralela"; cfg["subpasta_push"] = "CACT - PUSH Diário"
    os.makedirs(tmp_path / "CACT - Agenda Paralela", exist_ok=True); os.makedirs(tmp_path / "CACT - PUSH Diário", exist_ok=True)
    return cfg


def _push(cfg, dia, eventos):
    caminho = os.path.join(push.pasta_push(cfg), f"Relatorio_PUSH_{dia:%Y-%m-%d}_0900.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Publicações e Andamentos"; ws.append(CAB)
    for ev in eventos:
        data, proc, tipo, teor = ev[:4]
        autor, reu = (ev[4], ev[5]) if len(ev) > 5 else (None, None)
        ws.append([data, None, proc, "TRT18", "Justiça do Trabalho", "1º Grau", None, "11ª VT", tipo, teor, autor, reu, "PJe Push"])
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


def test_pauta_publicada_nao_e_intimacao(cfg, tmp_path):
    _cfg_pc(cfg, tmp_path)
    c = _push(cfg, date(2026, 10, 16), [(datetime(2026, 10, 15, 9, 0), "0011341-45.2015.5.18.0010", "Pauta/Audiência / Publicação/Disponibilização", "Disponibilizado(a) o(a) Pauta de Julgamento")])
    assert push.ler_push(c)[0]["intimacao"] is False
    assert push.caminho_push(cfg, date(2026, 10, 16)) == c


def test_classificar_cobertura_relativa_a_data_do_evento():
    """Cobertura pela Paralela: o processo aparece numa Paralela datada entre ev-1 e ev+4 (a publicação do DJEN sai no
    dia da intimação ou nos seguintes), não só na de hoje/ontem (18/10/2026: 8 falsos SEM COBERTURA numa rodada de 29)."""
    ev = [{"norm": "1", "intimacao": True, "data": date(2026, 10, 14)}, {"norm": "2", "intimacao": True, "data": date(2026, 10, 16)},
          {"norm": "3", "intimacao": True, "data": date(2026, 10, 16)}, {"norm": "4", "intimacao": False, "data": date(2026, 10, 16)},
          {"norm": "5", "intimacao": True, "data": date(2026, 10, 16)}]
    paralelas = {"1": [date(2026, 10, 15)], "5": [date(2026, 10, 10)]}                        # 5: Paralela antiga demais
    agenda = {"2": [(date(2026, 10, 21), date(2026, 10, 30))]}                               # publicado DEPOIS do evento (anotação manual)
    r = push.classificar(ev, paralelas, agenda, date(2026, 10, 18), primeira_paralela=date(2026, 10, 15))
    assert r[0]["situacao"] == "coberto pela Paralela" and "15/10" in r[0]["detalhe"]
    assert r[1]["situacao"] == "coberto pela Agenda" and "30/10" in r[1]["detalhe"]
    assert r[2]["situacao"] == "SEM COBERTURA"
    assert r[3]["situacao"] == "informativo"
    assert r[4]["situacao"] == "SEM COBERTURA"


def test_evento_anterior_ao_inicio_das_paralelas_nao_vira_linha():
    """Rodada inicial lê 7 dias de e-mail; intimação de antes da primeira Paralela foi tratada à mão na época."""
    ev = [{"norm": "9", "intimacao": True, "data": date(2026, 10, 13)}]
    r = push.classificar(ev, {}, {}, date(2026, 10, 18), primeira_paralela=date(2026, 10, 15))
    assert r[0]["situacao"] == "anterior à Paralela"
    r = push.classificar(ev, {}, {}, date(2026, 10, 18), primeira_paralela=None)             # nenhuma Paralela ainda: não dá para conferir
    assert r[0]["situacao"] == "anterior à Paralela"


def test_agenda_antiga_nao_cobre():
    ev = [{"norm": "2", "intimacao": True, "data": date(2026, 10, 16)}]
    agenda = {"2": [(date(2026, 9, 1), date(2026, 9, 10))]}
    assert push.classificar(ev, {}, agenda, date(2026, 10, 18), primeira_paralela=date(2026, 10, 15))[0]["situacao"] == "SEM COBERTURA"


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


def test_processo_desconhecido_na_agenda_usa_partes_do_email(cfg, tmp_path, monkeypatch):
    """Encaminhamento de parceiro: processo que nunca entrou na Agenda chega em Lançamentos com Autor/Réu do e-mail do PJe."""
    _cfg_pc(cfg, tmp_path); monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None); monkeypatch.setattr(push, "_livro_aberto_no_excel", lambda c: None)
    _agenda(cfg, [])
    par = ap.caminho_paralela(cfg, date(2026, 10, 16)); gerar_paralela([_linha("1", "0000001-11.2026.5.18.0001")], [], par, date(2026, 10, 16), {"ok": True})
    c = _push(cfg, date(2026, 10, 16), [(datetime(2026, 10, 15, 23, 40), "0000777-77.2026.5.18.0007", "Intimação / Expedição de documento",
                                          "Expedido(a) intimação a(o) JOSE MARIA", "JOSE MARIA DE SA", "COMPANHIA EXEMPLO - COMURG")])
    assert push.ler_push(c)[0]["autor"] == "JOSE MARIA DE SA" and push.ler_push(c)[0]["reu"] == "COMPANHIA EXEMPLO - COMURG"
    push.conferir_push(cfg, date(2026, 10, 16))
    ws = openpyxl.load_workbook(par)["Lançamentos"]
    nova = [r for r in ws.iter_rows(min_row=8, values_only=True) if r and r[4] == "0000777-77.2026.5.18.0007"][0]
    assert nova[5] == "JOSE MARIA DE SA" and nova[6] == "COMPANHIA EXEMPLO - COMURG" and nova[9] is None      # J fica para o advogado


def test_sem_cobertura_colapsa_por_processo_e_dia_e_nao_duplica(cfg, tmp_path, monkeypatch):
    """TRT-Push manda 'Publicado' e 'Disponibilizado' da mesma intimação e o relatório repete e-mails: uma linha CONFERIR
    por processo e dia basta para o advogado conferir no PJe; repetir a conferência não cria linha nova."""
    _cfg_pc(cfg, tmp_path); monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None); monkeypatch.setattr(push, "_livro_aberto_no_excel", lambda c: None)
    _agenda(cfg, [])
    par = ap.caminho_paralela(cfg, date(2026, 10, 16)); gerar_paralela([_linha("1", "0000001-11.2026.5.18.0001")], [], par, date(2026, 10, 16), {"ok": True})
    gerar_paralela([], [], ap.caminho_paralela(cfg, date(2026, 10, 15)), date(2026, 10, 15), {"ok": True})   # publicação de 15/10 confere-se na Paralela de 14 ou 15
    _push(cfg, date(2026, 10, 16), [
        (datetime(2026, 10, 15, 5, 49), "0000777-77.2026.5.18.0007", "Intimação / Publicação/Disponibilização", "Publicado(a) o(a) intimação em 15/10/2026"),
        (datetime(2026, 10, 15, 5, 49), "0000777-77.2026.5.18.0007", "Intimação / Publicação/Disponibilização", "Disponibilizado(a) o(a) intimação no Diário"),
        (datetime(2026, 10, 15, 5, 49), "0000777-77.2026.5.18.0007", "Intimação / Publicação/Disponibilização", "Publicado(a) o(a) intimação em 15/10/2026"),   # e-mail repetido
        (datetime(2026, 10, 14, 9, 0), "0000777-77.2026.5.18.0007", "Intimação / Expedição de documento", "Expedido(a) intimação a(o) X"),                    # anterior à 1ª Paralela (16/10): sem linha
    ])
    r = push.conferir_push(cfg, date(2026, 10, 16))
    assert sum(1 for e in r if e["situacao"] == "SEM COBERTURA") == 4                       # a aba de conferência continua por evento
    for _ in range(2):                                                                      # segunda rodada: idempotente
        ws = openpyxl.load_workbook(par)["Lançamentos"]
        novas = [x for x in ws.iter_rows(min_row=8, values_only=True) if x and x[4] == "0000777-77.2026.5.18.0007"]
        assert len(novas) == 2                                                              # 15/10 (3 eventos → 1 linha) e 14/10
        linha15 = [n for n in novas if n[3].date() == date(2026, 10, 15)][0]
        assert "+2 outro" in linha15[10] and linha15[21].startswith("PUSH:")
        push.conferir_push(cfg, date(2026, 10, 16))


def test_fim_de_semana_sem_paralela_cria_paralela_so_com_push(cfg, tmp_path, monkeypatch):
    """Sábado/domingo não têm Paralela da VM. Intimação do Push SEM COBERTURA cria a Paralela do dia só com as linhas
    do Push, para o vigia abrir; sem intimação nova, nenhum arquivo é criado."""
    _cfg_pc(cfg, tmp_path); monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None); monkeypatch.setattr(push, "_livro_aberto_no_excel", lambda c: None)
    _agenda(cfg, [])
    sab = date(2026, 10, 17)
    gerar_paralela([], [], ap.caminho_paralela(cfg, date(2026, 10, 16)), date(2026, 10, 16), {"ok": True})     # já existe Paralela (sexta)
    _push(cfg, sab, [(datetime(2026, 10, 16, 23, 40), "0000005-11.2026.5.18.0005", "Documento sigiloso", "só informativo")])
    assert push.conferir_push(cfg, sab) and not os.path.exists(ap.caminho_paralela(cfg, sab))
    _push(cfg, sab, [(datetime(2026, 10, 17, 0, 35), "0000567-95.2025.5.18.0012", "Intimação / Expedição de documento", "Expedido(a) intimação a(o) AFONSO", "AFONSO SOARES", "COMURG")])
    r = push.conferir_push(cfg, sab)
    par = ap.caminho_paralela(cfg, sab)
    assert os.path.exists(par) and [e["situacao"] for e in r] == ["SEM COBERTURA"]
    wb = openpyxl.load_workbook(par)
    assert "Conferência PUSH" in wb.sheetnames
    novas = [x for x in wb["Lançamentos"].iter_rows(min_row=8, values_only=True) if x and x[4] == "0000567-95.2025.5.18.0012"]
    assert len(novas) == 1 and novas[0][5] == "AFONSO SOARES" and novas[0][21].startswith("PUSH:")


def test_publicacao_sem_paralela_do_dia_nao_e_conferivel():
    """Evento que É publicação (PJe 'Publicado/Disponibilizado' ou Jusbrasil) só se confere contra a Paralela do dia da
    disponibilização (ev-1) ou do próprio dia; sem esse arquivo, não dá para saber → 'sem Paralela para conferir', sem linha.
    Intimação expedida pelo sistema (sem DJEN) continua conferida contra Paralelas de ev-1 a ev+4."""
    ev = [{"norm": "1", "intimacao": True, "data": date(2026, 10, 14), "tipo": "Decisão / Intimação", "origem": "Jusbrasil"},
          {"norm": "2", "intimacao": True, "data": date(2026, 10, 14), "tipo": "Intimação / Publicação/Disponibilização", "origem": "TRT-Push (legado)"},
          {"norm": "3", "intimacao": True, "data": date(2026, 10, 14), "tipo": "Intimação / Expedição de documento", "origem": "PJe Push"},
          {"norm": "4", "intimacao": True, "data": date(2026, 10, 17), "tipo": "Intimação / Publicação/Disponibilização", "origem": "PJe Push"}]
    r = push.classificar(ev, {}, {}, date(2026, 10, 18), primeira_paralela=date(2026, 10, 15), datas_paralelas={date(2026, 10, 15), date(2026, 10, 16)})
    assert [e["situacao"] for e in r] == ["sem Paralela para conferir", "sem Paralela para conferir", "SEM COBERTURA", "SEM COBERTURA"]
    r = push.classificar(ev[3:], {"4": [date(2026, 10, 16)]}, {}, date(2026, 10, 18), primeira_paralela=date(2026, 10, 15), datas_paralelas={date(2026, 10, 16)})
    assert r[0]["situacao"] == "coberto pela Paralela"


def test_ler_push_traz_origem(cfg, tmp_path):
    _cfg_pc(cfg, tmp_path)
    c = _push(cfg, date(2026, 10, 16), [(datetime(2026, 10, 15, 9, 0), "0000001-11.2026.5.18.0001", "Intimação", "x")])
    assert push.ler_push(c)[0]["origem"] == "PJe Push"


def test_mensagem_de_terceiro_sem_cobertura_vira_linha_com_prazo_do_texto(cfg, tmp_path, monkeypatch):
    """E-mail de parceiro sobre processo que não está na Agenda (16/10/2026, Nicoli): entra como CONFERIR com o prazo que o
    próprio e-mail informa ('encerra-se em 28/10/2026'); sem data no texto, 5 dias úteis."""
    _cfg_pc(cfg, tmp_path); monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None); monkeypatch.setattr(push, "_livro_aberto_no_excel", lambda c: None)
    _agenda(cfg, [])
    par = ap.caminho_paralela(cfg, date(2026, 10, 16)); gerar_paralela([_linha("1", "0000001-11.2026.5.18.0001")], [], par, date(2026, 10, 16), {"ok": True})
    caminho = _push(cfg, date(2026, 10, 16), [])
    wb = openpyxl.load_workbook(caminho); ws = wb["Publicações e Andamentos"]
    ws.append([datetime(2026, 10, 16, 16, 24), None, "0000477-40.2026.5.18.0081", "TRT18", "Justiça do Trabalho", "", None, "", "Mensagem de terceiro / Prazo",
               "Bianka x Cooperativa - 0000477-40.2026.5.18.0081 — encaminho os autos para orçamento. Nosso prazo processual encerra-se em 28/10/2026.",
               "Bianka Vitória da Silva", "Cooperativa Central", "E-mail de terceiro", datetime(2026, 10, 16, 16, 24), "Bianka x Cooperativa", "lucas@nicoli.adv.br", "abc", "Não"])
    wb.save(caminho)
    r = push.conferir_push(cfg, date(2026, 10, 16))
    assert r[0]["intimacao"] is True and r[0]["situacao"] == "SEM COBERTURA"
    ws = openpyxl.load_workbook(par)["Lançamentos"]
    nova = [x for x in ws.iter_rows(min_row=8, values_only=True) if x and x[4] == "0000477-40.2026.5.18.0081"][0]
    assert nova[0].date() == date(2026, 10, 28) and nova[2].startswith("Verificar e-mail de terceiro") and nova[5] == "Bianka Vitória da Silva"
    assert "lucas@nicoli.adv.br" in nova[10] and "28/10/2026" in nova[10]


def test_email_de_terceiro_so_e_novidade_se_o_processo_for_desconhecido():
    """Processo que a Agenda já conhece (em qualquer data) ou que está em Lançamentos de uma Paralela: correspondência de
    rotina, informativo. Processo só em 'Fora da Agenda' ou em lugar nenhum: SEM COBERTURA (caso Nicoli, 16/10/2026).
    E-mail de terceiro não é publicação: a regra 'anterior à Paralela' não se aplica."""
    base = {"intimacao": True, "data": date(2026, 10, 13), "tipo": "Mensagem de terceiro / Prazo", "origem": "E-mail de terceiro"}
    ev = [dict(base, norm="1"), dict(base, norm="2"), dict(base, norm="3"), dict(base, norm="4")]
    paralelas = {"2": [date(2026, 10, 15)]}                                 # Lançamentos
    fora = {"3": [date(2026, 10, 15)]}                                      # só Fora da Agenda
    agenda = {"1": [(date(2025, 3, 1), date(2025, 3, 10))]}                 # conhecido, ainda que antigo
    r = push.classificar(ev, paralelas, agenda, date(2026, 10, 18), primeira_paralela=date(2026, 10, 15), datas_paralelas={date(2026, 10, 15)}, fora=fora)
    assert [e["situacao"] for e in r] == ["coberto pela Agenda", "coberto pela Paralela", "SEM COBERTURA", "SEM COBERTURA"]
    assert "conhecido" in r[0]["detalhe"]


def test_email_de_parceiro_sem_numero_vira_linha_sem_processo(cfg, tmp_path, monkeypatch):
    _cfg_pc(cfg, tmp_path); monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None); monkeypatch.setattr(push, "_livro_aberto_no_excel", lambda c: None)
    _agenda(cfg, [])
    par = ap.caminho_paralela(cfg, date(2026, 10, 16)); gerar_paralela([_linha("1", "0000001-11.2026.5.18.0001")], [], par, date(2026, 10, 16), {"ok": True})
    caminho = _push(cfg, date(2026, 10, 16), [])
    wb = openpyxl.load_workbook(caminho); ws = wb["Publicações e Andamentos"]
    ws.append([datetime(2026, 10, 16, 9, 0), None, None, "", "", "", None, "", "Mensagem de terceiro (sem nº de processo)",
               "Pedido — Pode me passar um orçamento para a liquidação de um acordo?", "", "", "E-mail de terceiro", datetime(2026, 10, 16, 9, 0), "Pedido", "welitonadvogado@gmail.com", "k1", "Não"])
    wb.save(caminho)
    ev = push.ler_push(caminho)
    assert len(ev) == 1 and ev[0]["norm"] == "" and ev[0]["intimacao"] is True
    r = push.conferir_push(cfg, date(2026, 10, 16))
    assert r[0]["situacao"] == "SEM COBERTURA"
    ws = openpyxl.load_workbook(par)["Lançamentos"]
    nova = [x for x in ws.iter_rows(min_row=8, values_only=True) if x and x[4] == "(sem número no e-mail)"]
    assert len(nova) == 1 and "welitonadvogado@gmail.com" in nova[0][10] and nova[0][0].date() == date(2026, 10, 23)

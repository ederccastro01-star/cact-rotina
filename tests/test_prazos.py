# -*- coding: utf-8 -*-
from datetime import date

import pytest

from paralela.prazos import calcular_prazo, prazo_legal, prazo_no_texto, tribunal_do_processo

PUB = date(2026, 10, 19)


def test_oito_dias_uteis(cfg):
    assert calcular_prazo(PUB, "Contrarrazões", 8, None, "TRT18", cfg).data == date(2026, 10, 29)


def test_ed_cinco(cfg):
    p = calcular_prazo(PUB, "Embargos de Declaração", None, None, "TRT18", cfg)
    assert p.data == date(2026, 10, 26)
    assert p.origem == "legal"
    assert p.dias == 5


def test_menor_entre_texto_e_tabela(cfg):
    p = calcular_prazo(PUB, "Embargos de Declaração", 30, None, "TRT18", cfg)
    assert p.dias == 5
    assert p.alertas == ["texto diz 30 dias, lei diz 5 (CLT 897-A)"]


def test_texto_maior_que_legal_recurso(cfg):
    p = calcular_prazo(PUB, "Contrarrazões", 15, None, "TRT18", cfg)
    assert p.dias == 8
    assert "15" in p.alertas[0] and "8" in p.alertas[0]


def test_feriado_02_11(cfg):
    assert calcular_prazo(PUB, "Manifestação", 10, None, "TRT18", cfg).data == date(2026, 11, 3)


def test_48_horas(cfg):
    assert prazo_no_texto("comprovar o depósito no prazo de 48 horas, sob pena") == 2
    assert calcular_prazo(PUB, "Pagamento", 2, None, "TRT18", cfg).data == date(2026, 10, 21)


def test_sem_prazo_padrao_5_com_alerta(cfg):
    p = calcular_prazo(PUB, "Manifestação", None, None, "TRT18", cfg)
    assert p.dias == 5 and p.origem == "padrao"
    assert p.alertas == ["sem prazo no texto: 5 dias úteis (CPC 218 §3º)"]


def test_ciencia_sem_alerta(cfg):
    p = calcular_prazo(PUB, "Ciência (ofícios expedidos)", None, None, "TRT18", cfg, ciencia=True)
    assert p.dias == 5 and p.alertas == []


def test_ia_quando_regex_nao_acha(cfg):
    p = calcular_prazo(PUB, "Manifestação", None, 10, "TRT18", cfg)
    assert p.dias == 10 and p.origem == "texto"


def test_recesso(cfg):
    p = calcular_prazo(date(2026, 12, 18), "Embargos de Declaração", None, None, "TRT18", cfg)
    assert p.data == date(2027, 1, 27)


def test_data_retorno(cfg):
    p = calcular_prazo(PUB, "Acompanhamento: sobrestado", None, None, "TRT1", cfg, data_retorno=date(2027, 2, 16))
    assert p.data == date(2027, 2, 16) and p.origem == "retorno"


def test_tribunal_sem_calendario(cfg):
    p = calcular_prazo(PUB, "Contrarrazões", 8, None, "TRT10", cfg)
    assert "calendário do tribunal não cadastrado" in p.alertas
    assert p.data == date(2026, 10, 29)


def test_tribunal_do_processo():
    assert tribunal_do_processo("00117579820245180009") == "TRT18"
    assert tribunal_do_processo("57297757520258090051") == "TJGO"
    assert tribunal_do_processo("10000446020245020040") == "TRT2"
    assert tribunal_do_processo("10005022320255000000") == "TST"
    assert tribunal_do_processo("10021952620264013504") == "TRF1"
    assert tribunal_do_processo("abc") == "?"


def test_por_extenso():
    assert prazo_no_texto("no prazo de cinco dias") == 5
    assert prazo_no_texto("Prazo: 05 (cinco) dias.") == 5
    assert prazo_no_texto("prazo comum de 8 (oito) dias") == 8
    assert prazo_no_texto("no prazo de trinta dias, conforme requerido") == 30
    assert prazo_no_texto("Prazo e fins legais.") is None
    assert prazo_no_texto("180 dias") is None


def test_prazo_legal_tabela(cfg):
    assert prazo_legal("Contraminuta ao Agravo de Petição", cfg) == (8, "CLT 895/897/896")
    assert prazo_legal("Impugnação aos cálculos (art. 879 §2º)", cfg) == (8, "CLT 879 §2º")
    assert prazo_legal("Manifestação s/ Embargos à Execução", cfg) == (5, "CLT 884")
    assert prazo_legal("Informar dados bancários", cfg) == (None, None)


def test_varios_prazos_no_texto_vale_o_menor_com_alerta(cfg):
    from paralela.prazos import prazos_no_texto
    txt = "intimem-se as partes no prazo de 15 dias ... impugnação em 5 (cinco) dias ... manifestar em 05 (cinco) dias"
    assert prazos_no_texto(txt) == [15, 5, 5]
    p = calcular_prazo(PUB, "Manifestação", prazo_no_texto(txt), None, "TRT18", cfg, n_textos=prazos_no_texto(txt))
    assert p.dias == 5 and any("vários prazos no texto" in a for a in p.alertas)


def test_prazo_so_da_ia_gera_alerta(cfg):
    p = calcular_prazo(PUB, "Manifestação", None, 10, "TRT18", cfg)
    assert p.dias == 10 and any("só pela IA" in a for a in p.alertas)


def test_data_retorno_gera_alerta(cfg):
    p = calcular_prazo(PUB, "Acompanhamento: sobrestado", None, None, "TRT1", cfg, data_retorno=date(2027, 2, 16))
    assert p.origem == "retorno" and any("retorno sugerida pela IA" in a for a in p.alertas)

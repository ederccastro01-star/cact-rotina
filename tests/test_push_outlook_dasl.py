# -*- coding: utf-8 -*-
"""Filtro DASL do Outlook: a data tem de ir em ISO. Em 18/10/2026 o formato mm/dd/yyyy ('10/11/2026') foi lido pelo
Windows em pt-BR como 10 de novembro e a coleta devolveu 0 e-mails (na véspera, '10/10/2026' era ambíguo e passou)."""
from datetime import datetime

from pc.push_outlook import _dasl


def test_dasl_usa_data_iso_e_padroes_de_assunto_e_remetente():
    q = _dasl(datetime(2026, 10, 11, 13, 12))
    assert ">= '2026-10-11 13:12'" in q and "10/11/2026" not in q
    assert q.startswith("@SQL=") and "LIKE '%PUSH%'" in q and "LIKE '%jus.br%'" in q and "LIKE '%jusbrasil%'" in q


def test_dasl_com_limite_superior_tambem_em_iso():
    q = _dasl(datetime(2025, 10, 11, 0, 0), datetime(2026, 10, 11, 13, 12))
    assert ">= '2025-10-11 00:00'" in q and "< '2026-10-11 13:12'" in q

# -*- coding: utf-8 -*-
import json
import os
import sys

import pytest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

FIXTURES = os.path.join(RAIZ, "tests", "fixtures")
REAIS = os.path.join(FIXTURES, "reais")


@pytest.fixture
def sintetico():
    import cact_rotina as cr
    return cr.carregar_de_json(os.path.join(FIXTURES, "sintetico.json"))


@pytest.fixture
def reais():
    caminho = os.path.join(REAIS, "djen_2026-10-16.json")
    if not os.path.exists(caminho):
        pytest.skip("fixture real ausente (tests/fixtures/baixar_reais.py)")
    import cact_rotina as cr
    return cr.carregar_de_json(caminho)


@pytest.fixture
def agenda_amostra():
    caminho = os.path.join(REAIS, "agenda_amostra.xlsx")
    if not os.path.exists(caminho):
        pytest.skip("agenda_amostra.xlsx ausente (tests/fixtures/baixar_reais.py)")
    return caminho


@pytest.fixture
def cfg(tmp_path):
    with open(os.path.join(RAIZ, "config_nuvem.json"), encoding="utf-8") as fh:
        base = json.load(fh)
    c = dict(base)
    c["anthropic_api_key"] = ""
    c["oabs_escritorio"] = ["23147/GO", "51509/GO", "26734/GO", "8060/GO"]
    c.setdefault("processos_monitorados", [])
    # nunca tocar no Outlook nem no razão real durante os testes (um teste do vigia disparou coleta real em 17/10)
    c["push"] = {"fontes": [], "janela_dias": 7, "arquivar_em": None, "ledger": str(tmp_path / "push_ledger.json")}
    return c

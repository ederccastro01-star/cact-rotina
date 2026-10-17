# -*- coding: utf-8 -*-
import os

import cact_rotina as cr

AQUI = os.path.dirname(os.path.abspath(__file__))


def test_carregar_de_json_sintetico():
    pubs = cr.carregar_de_json(os.path.join(AQUI, "fixtures", "sintetico.json"))
    assert len(pubs) == 6
    assert pubs[0]["item"]["destinatarios"][0]["polo"] == "A"
    assert "Prazo: 05" in pubs[0]["texto"]
    assert pubs[0]["processo"] == "0000001-11.2026.5.18.0001"

# -*- coding: utf-8 -*-
import os
from datetime import datetime

import pytest

from pc import transporte as tr

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _linha(decisao, ids, lote="", linha=8, aba="Lançamentos"):
    return {"linha": linha, "aba": aba, "a_k": [datetime(2026, 10, 29), None, "Contrarrazões", datetime(2026, 10, 19),
            "0000001-11.2026.5.18.0001", "AUTOR", "REU", "EDER", "NORMAL", "AUTOR(A)", "resumo"],
            "decisao": decisao, "motivo": "", "link": "https://x/1", "ids": ids, "lote": lote}


def test_colunas():
    assert tr.COL == {"polo_conferido": 14, "id_djen": 15, "link": 16, "origem": 17, "lote": 18, "conferir": 19}
    assert tr.COL_PARALELA == {"destino": 12, "decisao": 13, "motivo": 14, "link": 21, "ids": 22, "lote": 23, "obs": 24, "prazo_calculado": 25, "corrigir_j": 26}


def test_selecionar_so_decisoes_de_agendar():
    linhas = [_linha("Agendar", ["1"]), _linha("Descartar", ["2"], linha=9), _linha("Corrigir e agendar", ["3"], linha=10),
              _linha("", ["4"], linha=11), _linha("agendar", ["5"], linha=12)]
    sel, avisos = tr.selecionar(linhas, set())
    assert [l["linha"] for l in sel] == [8, 10, 12] and avisos == []


def test_selecionar_ignora_ja_transportada():
    sel, avisos = tr.selecionar([_linha("Agendar", ["1"], lote="Transportado em 17/10/2026 09:10")], set())
    assert sel == [] and "já transportada" in avisos[0]


def test_selecionar_ignora_id_ja_na_agenda():
    sel, avisos = tr.selecionar([_linha("Agendar", ["1", "2"])], {"2"})
    assert sel == [] and "já existe" in avisos[0]


def test_reexecucao_nao_duplica():
    linhas = [_linha("Agendar", ["1"]), _linha("Agendar", ["2"], linha=9)]
    sel1, _ = tr.selecionar(linhas, set())
    ids = {i for l in sel1 for i in l["ids"]}
    sel2, _ = tr.selecionar(linhas, ids)
    assert len(sel1) == 2 and sel2 == []


def test_valores_agenda_colunas():
    v = tr.valores_agenda(_linha("Agendar", ["1", "2"]), "Transportado em 17/10/2026 09:10")
    assert set(v) == {1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 15, 16, 17, 18, 19}
    assert v[15] == "1;2" and v[17] == "AUTO" and v[18] == "Transportado em 17/10/2026 09:10" and v[16] == "https://x/1"
    assert 14 not in v and 2 not in v


def test_ids_da_agenda():
    assert tr.ids_da_agenda(["1;2", None, "", 3, "4"]) == {"1", "2", "3", "4"}


def test_desfazer_por_id_apos_reordenacao():
    razao = tr.Razao(quando="2026-10-17T09:10", paralela="p.xlsx",
                     linhas=[{"aba": "Publicações", "linha_agenda": 3230, "ids": ["111"], "valores": {}, "linha_paralela": 8}])
    coluna = [None] * 4 + ["111"] + [None] * 10     # "111" agora está na linha 8 + 4 = 12
    assert tr.linhas_para_desfazer(razao, coluna) == [12]


def test_razao_salvar_carregar(tmp_path):
    r = tr.Razao(quando="2026-10-17T09:10", paralela="p.xlsx", linhas=[{"aba": "Publicações", "linha_agenda": 1, "ids": ["1"],
                                                                        "valores": {"1": "2026-10-29"}, "linha_paralela": 8}])
    caminho = str(tmp_path / "r.json")
    r.salvar(caminho)
    assert tr.Razao.carregar(caminho) == r


def test_ler_paralela_gerada():
    caminho = os.path.join(RAIZ, "tests", "_saida", "CACT - Agenda Paralela 16-10-2026.xlsx")
    if not os.path.exists(caminho):
        pytest.skip("gere a Paralela local primeiro (Task 6, step 6)")
    linhas = tr.ler_paralela(caminho)
    assert len(linhas) >= 30
    l = linhas[0]
    assert l["aba"] == "Lançamentos" and l["linha"] == 8 and len(l["a_k"]) == 11 and l["ids"] and l["decisao"] == "" and l["lote"] == ""
    assert all(x["aba"] in ("Lançamentos", "Pendentes") for x in linhas)


def test_ids_normaliza_numero_do_excel():
    # o Excel devolve a célula numérica como float: 759139484.0 tem de casar com "759139484"
    assert tr.ids_da_agenda([759139484.0, "1;2", 3]) == {"759139484", "1", "2", "3"}
    assert tr.linhas_para_desfazer(tr.Razao(quando="x", paralela="p", linhas=[{"ids": ["759139484"]}]), [None, 759139484.0]) == [9]

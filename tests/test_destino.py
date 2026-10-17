# -*- coding: utf-8 -*-
import pytest

from paralela.destino import (RefProc, bloco, decidir_destino, indexar_agenda_completa, lado_nosso,
                              ESCRITORIO, CONTRARIA, NOVO, DISPENSADA)
from paralela.polo import agrupar_atos

PROC1 = "00000011120265180001"


@pytest.fixture
def atos(sintetico):
    return {a.ids[0]: a for a in agrupar_atos(sintetico)}


def ref(js=("AUTOR(A)",), conferido=True):
    return RefProc(ref={"Reclamante": "AUTOR TESTE", "Reclamado": "REU TESTE LTDA", "Advogado": "EDER", "Prioridade": "NORMAL",
                        "Cliente": "/".join(js), "Despacho": ""}, js=set(js), conferido=conferido, inPub=True, inPF=False, inArq=False)


def test_lado_nosso():
    assert lado_nosso(ref(("AUTOR(A)",))) == "A"
    assert lado_nosso(ref(("RÉU",))) == "P"
    assert lado_nosso(ref(("AUTOR(A)", "RÉU"))) is None
    assert lado_nosso(ref(())) is None
    assert lado_nosso(None) is None


def test_escritorio_quando_polo_igual_j(atos, cfg):
    d = decidir_destino(atos["9001"], ref(), cfg, "Informar dados bancários (alvará)")
    assert d.rotulo == ESCRITORIO and d.motivos == [] and d.fonte_polo == "DJEN + coluna J"


def test_contraria_quando_j_conferida(atos, cfg):
    d = decidir_destino(atos["9002"], ref(), cfg, "Impugnação aos cálculos - aguardar")
    assert d.rotulo == CONTRARIA and d.motivos == []


def test_nao_conferida_vira_escritorio_com_motivo(atos, cfg):
    d = decidir_destino(atos["9002"], ref(conferido=False), cfg, "Impugnação aos cálculos")
    assert d.rotulo == ESCRITORIO and any("J não conferida" in m for m in d.motivos)


def test_duas_partes_escritorio(atos, cfg):
    d = decidir_destino(atos["9003"], ref(), cfg, "Embargos de Declaração")
    assert d.rotulo == ESCRITORIO and d.motivos == []


def test_prazo_comum_so_contraria_listada(atos, cfg):
    d = decidir_destino(atos["9006"], ref(), cfg, "Impugnação aos cálculos")
    assert d.rotulo == ESCRITORIO and any("prazo comum" in m for m in d.motivos)


def test_sem_item_estruturado(cfg):
    pub = {"processo": "0000001-11.2026.5.18.0001", "texto": "INTIMAÇÃO Fica V. Sa. intimado. Prazo: 5 dias.", "item": {}}
    ato = agrupar_atos([pub])[0]
    assert ato.polos == set()
    d = decidir_destino(ato, ref(), cfg, "Manifestação")
    assert d.rotulo == ESCRITORIO and any("não identificado" in m for m in d.motivos) and d.fonte_polo == ""


def test_j_contraditoria(atos, cfg):
    d = decidir_destino(atos["9001"], ref(("AUTOR(A)", "RÉU")), cfg, "Manifestação")
    assert d.rotulo == ESCRITORIO and any("contraditória" in m for m in d.motivos)


def test_processo_novo_oab_escritorio(atos, cfg):
    assert bloco(atos["9001"], None, cfg) == "LANCAMENTOS"
    d = decidir_destino(atos["9001"], None, cfg, "Manifestação")
    assert d.rotulo == NOVO and any("fora da Agenda" in m for m in d.motivos)


def test_fora_da_agenda_parceiro(atos, cfg):
    assert bloco(atos["9005"], None, cfg) == "FORA"


def test_bloco_processo_monitorado(atos, cfg):
    cfg["processos_monitorados"] = ["0000002-11.2026.8.09.0001"]
    assert bloco(atos["9005"], None, cfg) == "LANCAMENTOS"


def test_bloco_na_agenda(atos, cfg):
    assert bloco(atos["9005"], ref(), cfg) == "LANCAMENTOS"


def test_dispensada_por_tipo(atos, cfg):
    cfg["dispensas"] = ["Intimação"]
    d = decidir_destino(atos["9001"], ref(), cfg, "Manifestação")
    assert d.rotulo == DISPENSADA and d.motivos


def test_ciencia_flag(atos, cfg):
    assert decidir_destino(atos["9001"], ref(), cfg, "Ciência (ofício)").ciencia is True
    assert decidir_destino(atos["9001"], ref(), cfg, "Manifestação").ciencia is False


def test_indexar_le_tres_abas_e_coluna_N(cfg, agenda_amostra):
    cfg["arquivo_agenda"] = agenda_amostra
    cfg["aba_prazos_feitos_arquivo"] = "Prazos Feitos (Arquivo)"
    ref_ = indexar_agenda_completa(cfg)
    # processo corrigido e marcado "Polo conferido" em 16/10/2026 (rodada 1 do saneamento J)
    r = ref_["00008275720255180018"]      # Leonardo Louro Xavier x Fujioka
    assert r.js == {"RÉU"} and r.conferido is True
    # processo só da aba Arquivo entra no índice
    assert any(v.inArq and not v.inPub and not v.inPF for v in ref_.values())
    # J normalizada: nunca "REU" sem acento
    assert all(j in {"AUTOR(A)", "RÉU"} for v in ref_.values() for j in v.js)


def test_cruzar_nao_inclui_processo_so_do_arquivo():
    """O Cruzamento (produção) continua cruzando só Publicações e Prazos Feitos; a aba Arquivo é só da Paralela."""
    import cact_rotina as cr
    ref = {"1": {"ref": {}, "inPub": False, "inPF": False, "inArq": True}, "2": {"ref": {}, "inPub": True, "inPF": False, "inArq": False}}
    pubs = [{"processo": "1", "texto": "a", "item": {}}, {"processo": "2", "texto": "b", "item": {}}]
    import re
    cr_norm = cr.norm_proc
    cr.norm_proc = lambda s: str(s)
    try:
        assert [m["processo"] for m in cr.cruzar(pubs, ref)] == ["2"]
    finally:
        cr.norm_proc = cr_norm

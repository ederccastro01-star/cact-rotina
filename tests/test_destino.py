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


# --------------------------------------------------------------------------- 2º grau (precatório/RPV, final .0000)
from paralela.destino import (MOTIVO_2G_SEM, segundo_grau, vincular_segundo_grau)
from paralela.polo import Ato


def _ato_2g(texto, destinatarios=(), advogados=()):
    return Ato(processo="0012147-32.2023.5.18.0000", norm="00121473220235180000", texto=texto, ids=["7001"], tribunal="TRT18",
               destinatarios=list(destinatarios), advogados=list(advogados), polos={"A"})


TEXTO_PRECAT = ("PODER JUDICIÁRIO TRT 18ª REGIÃO OJC RPV E PRECATÓRIOS Precat 0012147-32.2023.5.18.0000 REQUERENTE: CARLOS EDUARDO "
                "FERREIRA GANDARA REQUERIDO: COMPANHIA DE URBANIZACAO DE GOIANIA - COMURG INTIMAÇÃO Fica V. Sa. intimado ... "
                "autos de origem 0010472-12.2020.5.18.0009 ...")


def test_segundo_grau_e_sempre_lancamentos(cfg):
    ato = _ato_2g(TEXTO_PRECAT)
    assert segundo_grau(ato) and not segundo_grau(Ato(processo="x", norm="00000011120265180001", texto=""))
    assert bloco(ato, None, cfg) == "LANCAMENTOS"                     # sem OAB nossa e fora da Agenda: mesmo assim no radar


def test_vincula_pelo_processo_de_origem_citado():
    indice = {"00104721220205180009": ref(("AUTOR(A)",))}
    rp, origem, motivo = vincular_segundo_grau(_ato_2g(TEXTO_PRECAT), indice)
    assert rp is indice["00104721220205180009"] and origem == "00104721220205180009"
    assert "0010472-12.2020.5.18.0009" in motivo and motivo.startswith("2º grau")


def test_vincula_pelo_nome_da_parte_quando_nao_cita_origem():
    rp0 = RefProc(ref={"Reclamante": "Carlos Eduardo Ferreira Gândara", "Reclamado": "COMURG", "Cliente": "AUTOR(A)"}, js={"AUTOR(A)"}, conferido=True, inPub=True)
    indice = {"00104721220205180009": rp0, "00000011120265180001": ref()}
    texto = TEXTO_PRECAT.split("autos de origem")[0]
    rp, origem, motivo = vincular_segundo_grau(_ato_2g(texto, destinatarios=[{"nome": "CARLOS EDUARDO FERREIRA GANDARA", "polo": "A"}]), indice)
    assert rp is rp0 and origem == "00104721220205180009" and "GANDARA" in motivo.upper() and "0010472-12.2020.5.18.0009" in motivo


def test_sem_vinculo_devolve_motivo_de_conferencia():
    rp, origem, motivo = vincular_segundo_grau(_ato_2g("Precat 0012147-32.2023.5.18.0000 REQUERENTE: NINGUEM REQUERIDO: X"), {"00000011120265180001": ref()})
    assert rp is None and origem is None and motivo == MOTIVO_2G_SEM

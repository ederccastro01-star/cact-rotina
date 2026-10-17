# -*- coding: utf-8 -*-
from paralela.polo import (agrupar_atos, detectar_prazo_comum, extrair_intimados,
                           oabs_do_ato, polo_dos_intimados)

PROC1 = "00000011120265180001"


def _atos_proc1(sintetico):
    return [a for a in agrupar_atos(sintetico) if a.norm == PROC1]


def _por_id(atos, ident):
    return next(a for a in atos if str(ident) in a.ids)


def test_agrupa_duplicata_em_um_ato(sintetico):
    atos = _atos_proc1(sintetico)
    assert len(atos) == 4            # itens 1+4 (mesmo ato), 2, 3, 6
    a1 = _por_id(atos, 9001)
    assert sorted(a1.ids) == ["9001", "9004"]


def test_dois_atos_mesmo_processo_geram_duas_linhas(sintetico):
    atos = _atos_proc1(sintetico)
    assert _por_id(atos, 9001) is not _por_id(atos, 9002)


def test_polo_autor(sintetico):
    atos = _atos_proc1(sintetico)
    assert _por_id(atos, 9001).polos == {"A"}
    assert _por_id(atos, 9002).polos == {"P"}
    assert _por_id(atos, 9003).polos == {"A", "P"}


def test_tjgo_sem_bloco_usa_destinatarios(sintetico):
    ato = next(a for a in agrupar_atos(sintetico) if a.norm == "00000021120268090001")
    assert ato.intimados == []
    assert ato.polos == {"A"}
    assert ato.tribunal == "TJGO"


def test_prazo_comum(sintetico):
    atos = _atos_proc1(sintetico)
    assert _por_id(atos, 9006).prazo_comum is True
    assert _por_id(atos, 9001).prazo_comum is False


def test_detectar_prazo_comum_art_879():
    assert detectar_prazo_comum("intimado(a) nos termos do art. 879, §2º, da CLT. Prazo e fins legais.")
    assert not detectar_prazo_comum("Prazo de 05 dias.")


def test_nome_com_sufixo_casa():
    polos = polo_dos_intimados(["BALTAZAR DE OLIVEIRA CAMPOS E OUTROS"],
                               [{"nome": "BALTAZAR DE OLIVEIRA CAMPOS", "polo": "A"},
                                {"nome": "LBR - LACTEOS BRASIL S/A EM RECUPERACAO JUDICIAL", "polo": "P"}])
    assert polos == {"A"}


def test_extrair_intimados_dois_nomes():
    txt = "INTIMAÇÃO ... Intimado(s) / Citado(s)\n - MANOEL ARCHANJO DAMA FILHO\n - MEGS SERVICOS DE COBRANCA EIRELI\n"
    assert extrair_intimados(txt) == ["MANOEL ARCHANJO DAMA FILHO", "MEGS SERVICOS DE COBRANCA EIRELI"]


def test_oabs_do_ato(sintetico):
    atos = _atos_proc1(sintetico)
    assert oabs_do_ato(_por_id(atos, 9001)) == {"23147/GO", "99999/GO"}


def test_ato_campos_basicos(sintetico):
    a = _por_id(_atos_proc1(sintetico), 9003)
    assert a.processo == "0000001-11.2026.5.18.0001"
    assert a.tribunal == "TRT18"
    assert a.tipo_documento == "Acórdão"
    assert a.link.endswith("/9003")
    assert a.data_disp.isoformat() == "2026-10-16"


def test_reais_maioria_tem_polo(reais):
    atos = [a for a in agrupar_atos(reais) if a.tribunal == "TRT18" and a.intimados]
    com_polo = [a for a in atos if a.polos]
    assert len(atos) > 20
    assert len(com_polo) / len(atos) >= 0.9

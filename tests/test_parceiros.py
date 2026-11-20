# -*- coding: utf-8 -*-
"""Aba "Parceiros - Liquidação" (lista acumulada) e o desenho de colunas de 20/11/2026 (L nova; link em O)."""
import os
from datetime import date
from types import SimpleNamespace

import openpyxl

from paralela import gerar, parceiros as pa
from paralela.planilha import LARGURAS_EXCEL, Linha, gerar_paralela
from pc import agenda_paralela as ap
from pc import transporte as tr

ORDEM = ("PODER JUDICIÁRIO. Vistos. Transitada em julgado a sentença, intime-se o reclamante para apresentar os cálculos de "
         "liquidação, no prazo de 8 dias, sob pena de preclusão. Intimado(s) / Citado(s)\n - FULANO DE TAL\n\n")
CEJUSC = ("Audiência de conciliação designada. OBS.: Tendo em vista o volume de processos na Secretaria de Cálculos deste Regional, "
          "sugere-se às partes que procedam a liquidação da sentença/acórdão, para auxiliar na busca de uma composição amigável.")
SENTENCA = "Julgo procedentes em parte os pedidos. Liquidação por cálculos. Custas pela reclamada. Intimem-se as partes."


def _ato(texto, ids="1", processo="0000123-45.2026.5.18.0001", oab=("14725", "GO"), nome="WELLINGTON ALVES RIBEIRO"):
    return SimpleNamespace(texto=texto, processo=processo, norm=processo.replace("-", "").replace(".", ""), ids=ids.split(";"),
                           polos={"A"}, link="https://comunica.pje.jus.br/x/" + ids,
                           destinatarios=[{"nome": "FULANO DE TAL", "polo": "A"}, {"nome": "EMPRESA LTDA", "polo": "P"}],
                           advogados=[{"advogado": {"nome": nome, "numero_oab": oab[0], "uf_oab": oab[1]}},
                                      {"advogado": {"nome": "EDER CASTRO", "numero_oab": "23147", "uf_oab": "GO"}}])


# ----------------------------------------------------------------------------- filtro e regra sem IA
def test_filtro_por_palavras_e_regra_estreita():
    assert pa.candidato(ORDEM) and pa.ordem_direta(ORDEM)
    assert not pa.candidato(CEJUSC)                                   # aviso padrão do CEJUSC: 31 falsos em 19/11/2026
    assert pa.candidato(SENTENCA) and not pa.ordem_direta(SENTENCA)   # candidato para a IA, mas não entra sem ela
    assert not pa.candidato("Intime-se a parte para se manifestar sobre o laudo pericial em 5 dias.")


def test_detectar_sem_ia_so_ordem_direta(cfg):
    regs = pa.detectar([_ato(ORDEM, "1"), _ato(SENTENCA, "2"), _ato(CEJUSC, "3")], cfg, date(2026, 11, 23), date(2026, 11, 19), usar_ia=False)
    assert [r.ids for r in regs] == ["1"]
    r = regs[0]
    assert r.origem == pa.ORIGEM_REGRA and r.situacao == pa.AGUARDANDO and r.polo == "AUTOR"
    assert r.prazo == date(2026, 12, 3)                               # 23/11 + 8 dias úteis
    assert "Wellington Alves Ribeiro (OAB 14725/GO)" in r.parceiro and "Eder" not in r.parceiro      # advogado do escritório fica fora
    assert r.reclamante == "FULANO DE TAL" and r.reclamado == "EMPRESA LTDA"


def test_detectar_com_ia_decide(cfg, monkeypatch):
    respostas = {"1": {"apresentar_calculos": True, "quem": "autor", "prazo_dias": 10, "resumo": "Autor intimado a liquidar em 10 dias"},
                 "2": {"apresentar_calculos": False, "quem": "outro", "prazo_dias": None, "resumo": ""}}
    monkeypatch.setattr(pa, "classificar_ia", lambda texto, polo, cfg: respostas["1" if texto == ORDEM else "2"])
    regs = pa.detectar([_ato(ORDEM, "1"), _ato(SENTENCA, "2")], cfg, date(2026, 11, 23), date(2026, 11, 19), usar_ia=True)
    assert len(regs) == 1 and regs[0].origem == pa.ORIGEM_IA and regs[0].resumo.startswith("Autor intimado")
    assert regs[0].prazo == date(2026, 12, 7)                         # prazo lido pela IA: 10 dias úteis


def test_parceiro_pelo_cadastro(cfg):
    ato = _ato(ORDEM)
    ato.advogados.append({"advogado": {"nome": "OUTRO ADVOGADO", "numero_oab": "999", "uf_oab": "GO"}})
    assert pa.advogados_parceiros(ato, cfg, {"nomes": [], "oabs": [{"uf": "GO", "oab": "14725"}]}) == "Wellington Alves Ribeiro (OAB 14725/GO)"
    assert "Outro Advogado" in pa.advogados_parceiros(ato, cfg, None)          # sem Cadastro: todos, menos os do escritório


# ----------------------------------------------------------------------------- acumulado
def _reg(ids, processo, prazo, situacao=pa.AGUARDANDO, obs=""):
    return pa.Registro(situacao=situacao, detectado=date(2026, 11, 10), prazo=prazo, processo=processo, parceiro="Parceiro X",
                       resumo="r", link="https://x/" + ids, observacoes=obs, ids=ids, origem=pa.ORIGEM_IA)


def test_atualizar_situacoes():
    hoje = date(2026, 11, 20)
    antigos = [_reg("1", "0000001-11.2026.5.18.0001", date(2026, 11, 18)),                       # venceu e não chegou
               _reg("2", "0000002-11.2026.5.18.0002", date(2026, 11, 18)),                       # chegou à Agenda
               _reg("3", "0000003-11.2026.5.18.0003", date(2026, 11, 25)),                       # dentro do prazo
               _reg("4", "0000004-11.2026.5.18.0004", date(2026, 11, 1), obs="Baixar - já cobrado"),
               _reg("5", "0000005-11.2026.5.18.0005", date(2026, 10, 1), situacao="Chegou à Agenda (visto em 10/10/2026)"),   # > 30 dias
               _reg("6", "0000006-11.2026.5.18.0006", date(2026, 11, 5), situacao="Chegou à Agenda (visto em 12/11/2026)")]
    novos = [_reg("3", "0000003-11.2026.5.18.0003", date(2026, 11, 25)), _reg("7", "0000007-11.2026.5.18.0007", date(2026, 11, 30))]
    out = pa.atualizar(antigos, novos, {"00000021120265180002"}, hoje)
    assert [(r.ids, r.situacao) for r in out] == [
        ("1", pa.VENCIDO), ("3", pa.AGUARDANDO), ("7", pa.AGUARDANDO),
        ("6", "Chegou à Agenda (visto em 12/11/2026)"), ("2", "Chegou à Agenda (visto em 20/11/2026)")]


def test_aba_vai_e_volta_com_link_e_observacoes(tmp_path):
    c = str(tmp_path / "CACT - Agenda Paralela 19-11-2026.xlsx")
    gerar_paralela([], [], c, date(2026, 11, 19), {"ok": True}, [_reg("1;2", "0000001-11.2026.5.18.0001", date(2026, 11, 25))], "nota")
    wb = openpyxl.load_workbook(c); ws = wb[pa.ABA]
    assert [ws.cell(7, i).value for i in range(1, 14)] == pa.CABECALHOS and ws["J8"].value == "abrir teor" and ws["A5"].value == "nota"
    ws["K8"] = "liguei para o parceiro em 19/11"; wb.save(c)
    regs = pa.ler_aba(c)
    assert len(regs) == 1 and regs[0].link == "https://x/1;2" and regs[0].ids == "1;2" and regs[0].prazo == date(2026, 11, 25)
    assert regs[0].observacoes == "liguei para o parceiro em 19/11"


def test_paralela_sem_lista_nao_tem_aba_e_a_busca_pula_o_arquivo(cfg, tmp_path):
    """Paralela criada pelo PC (fim de semana) sai sem a aba: a rotina de segunda continua da lista de sexta."""
    pasta = str(tmp_path)
    gerar_paralela([], [], os.path.join(pasta, "CACT - Agenda Paralela 20-11-2026.xlsx"), date(2026, 11, 20), {"ok": True},
                   [_reg("1", "0000001-11.2026.5.18.0001", date(2026, 11, 25))])
    gerar_paralela([], [], os.path.join(pasta, "CACT - Agenda Paralela 21-11-2026.xlsx"), date(2026, 11, 21), {"ok": True})   # sábado, PC
    assert pa.ABA not in openpyxl.load_workbook(os.path.join(pasta, "CACT - Agenda Paralela 21-11-2026.xlsx")).sheetnames
    cfg["pasta_saida"] = pasta
    caminho, regs = pa.achar_anterior(cfg, date(2026, 11, 23))
    assert os.path.basename(caminho) == "CACT - Agenda Paralela 20-11-2026.xlsx" and [r.ids for r in regs] == ["1"]
    assert pa.achar_anterior(cfg, date(2026, 11, 20)) == (None, [])           # nunca o próprio dia nem um posterior


def test_gerar_acumula_de_um_dia_para_o_outro(sintetico, cfg, tmp_path, monkeypatch):
    """O ato Fora da Agenda do dia 16 entra na lista; no dia 19 a lista continua e a situação é recalculada."""
    monkeypatch.setattr(gerar, "indexar_agenda_completa", lambda _cfg: {})
    cfg["pasta_saida"] = str(tmp_path); cfg["anthropic_api_key"] = ""; cfg["arquivo_cadastro"] = ""
    fora = [p for p in sintetico if str(p["item"]["id"]) == "9005"]
    fora[0]["texto"] = ORDEM; fora[0]["item"]["texto"] = ORDEM
    c1 = gerar.gerar_agenda_paralela(fora, cfg, date(2026, 10, 16), usar_ia=False)
    ws = openpyxl.load_workbook(c1)[pa.ABA]
    assert ws["A8"].value == pa.AGUARDANDO and ws["L8"].value == "9005" and "Primeira lista" in ws["A5"].value
    proc = ws["D8"].value
    # dia seguinte: nenhuma publicação nova do processo, mas ele agora está na Agenda
    from paralela.destino import RefProc
    monkeypatch.setattr(gerar, "indexar_agenda_completa", lambda _cfg: {proc.replace("-", "").replace(".", ""): RefProc(js={"AUTOR(A)"})})
    outro = [p for p in sintetico if str(p["item"]["id"]) == "9003"]
    c2 = gerar.gerar_agenda_paralela(outro, cfg, date(2026, 10, 19), usar_ia=False)
    ws2 = openpyxl.load_workbook(c2)[pa.ABA]
    assert ws2["D8"].value == proc and ws2["A8"].value == "Chegou à Agenda (visto em 19/10/2026)"
    assert "Continuação de: CACT - Agenda Paralela 16-10-2026.xlsx" in ws2["A5"].value


# ----------------------------------------------------------------------------- desenho de colunas
def _linha(ids, **kw):
    base = dict(prazo=date(2026, 11, 26), peca="Contrarrazões", publicado=date(2026, 11, 16), processo="0000001-11.2026.5.18.0001",
                reclamante="A", reclamado="R", advogado="E", prioridade="NORMAL", cliente="AUTOR(A)", despacho="resumo",
                destino="PRAZO DO ESCRITÓRIO", ids=ids, link="https://x/" + ids)
    base.update(kw)
    return Linha(**base)


def test_larguras_e_despacho_anterior_na_coluna_L(tmp_path):
    c = str(tmp_path / "p.xlsx")
    gerar_paralela([_linha("1", despacho_anterior="LS entregue em 03/11")], [], c, date(2026, 11, 19), {"ok": True})
    ws = openpyxl.load_workbook(c)["Lançamentos"]
    assert ws["L7"].value == "Despacho anterior" and ws["L8"].value == "LS entregue em 03/11"
    assert ws["O7"].value == "Link (teor)" and ws["O8"].value == "abrir teor" and ws["O8"].hyperlink.target == "https://x/1"
    pedidas = [7.71, 8.00, 13.29, 9.43, 20, 20, 20, 20, 11.00, 9.29, 33, 33, 11.43, 8.29, 8.00, 12.14, 8.29]      # A a Q, como o Excel mostra
    assert LARGURAS_EXCEL[:17] == pedidas
    for letra, w in zip("ABCDEFGHIJKLMNOPQ", pedidas):
        assert abs(ws.column_dimensions[letra].width - (w + 5 / 7)) < 0.006


def test_despacho_anterior_vem_da_linha_mais_recente_de_prazos_feitos(cfg, tmp_path):
    import cact_rotina as cr
    ag = str(tmp_path / "agenda.xlsx")
    wb = openpyxl.Workbook(); pub = wb.active; pub.title = "Publicações"; pf = wb.create_sheet("Prazos Feitos")
    proc = "0000001-11.2026.5.18.0001"
    linha = lambda prazo, k: [prazo, None, "LS", prazo, proc, "A", "R", "E", "NORMAL", "AUTOR(A)", k]  # noqa: E731
    pub.append(linha(date(2026, 12, 1), "despacho de Publicações (não é este)"))
    pf.append(linha(date(2026, 11, 3), "LS entregue em 03/11"))
    pf.append(linha(date(2026, 9, 1), "IC antiga"))
    pf.append(linha(date(2026, 11, 10), None))                         # K vazia não apaga a anterior
    wb.save(ag)
    cfg["arquivo_agenda"] = ag
    ref = cr.indexar_agenda(cfg)
    assert ref["00000011120265180001"]["despacho_pf"] == "LS entregue em 03/11"


def _paralela_antiga(caminho, ids, decisao=None, lote=None):
    """Arquivo no desenho anterior a 20/11/2026: L Destino, M Decisão, N Motivo, U Link, V ID, W Lote, X Obs, Y, Z."""
    antigos = ["Destino", "Decisão", "Motivo CONFERIR", "Polo intimado", "Intimado(s)", "Fonte do polo", "Prazo no texto", "Prazo legal",
               "Base legal", "Link (teor)", "ID DJEN", "Lote", "Observações", "Prazo calculado", "Corrigir J para"]
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Lançamentos"
    for nome in ("Lançamentos", "Pendentes"):
        w = ws if nome == "Lançamentos" else wb.create_sheet(nome)
        for i, h in enumerate(["Prazo", "Dia/Sem", "Peça", "Publicado", "Processo ", "Reclamante", "Reclamado", "Advogado ", "Prioridade ",
                               "Cliente", "Despacho"] + antigos, start=1):
            w.cell(7, i, h)
    vals = [date(2026, 11, 26), None, "Contrarrazões", date(2026, 11, 16), "0000001-11.2026.5.18.0001", "A", "R", "E", "NORMAL", "AUTOR(A)",
            "resumo", "PRAZO DO ESCRITÓRIO", decisao, "motivo antigo", "AUTOR", "A", "DJEN", None, 8, "CLT", "abrir teor", ids, lote, None,
            date(2026, 11, 26), None]
    for i, v in enumerate(vals, start=1):
        ws.cell(8, i, v)
    ws.cell(8, 21).hyperlink = "https://x/" + ids
    wb.save(caminho)


def test_paralela_do_desenho_antigo_continua_legivel(cfg, tmp_path, monkeypatch):
    cfg["pc"] = dict(cfg.get("pc") or {}); cfg["pc"]["pasta_controladoria"] = str(tmp_path)
    os.makedirs(tmp_path / "CACT - Agenda Paralela", exist_ok=True)
    monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None)
    antiga = ap.caminho_paralela(cfg, date(2026, 11, 18))
    _paralela_antiga(antiga, "77")
    l = tr.ler_paralela(antiga)[0]
    assert l["cols"]["decisao"] == 13 and l["cols"]["ids"] == 22 and "despacho_anterior" not in l["cols"]
    assert l["ids"] == ["77"] and l["motivo"] == "motivo antigo" and l["destino"] == "PRAZO DO ESCRITÓRIO"
    assert l["link"] == "https://x/77"                                 # o endereço, não o texto "abrir teor"
    # a linha sem decisão vai para Pendentes do arquivo de hoje, que já é do desenho novo
    hoje = ap.caminho_paralela(cfg, date(2026, 11, 19))
    gerar_paralela([_linha("88")], [], hoje, date(2026, 11, 19), {"ok": True})
    assert [p["ids"] for p in ap.montar_pendentes(cfg, date(2026, 11, 19))] == [["77"]]
    ws = openpyxl.load_workbook(hoje)["Pendentes"]
    assert ws["W8"].value == "77" and ws["P8"].value == "motivo antigo" and ws["O8"].value == "https://x/77"
    assert ws["Y8"].value == "Paralela 18-11-2026 (sem decisão)" and ws["X8"].value in (None, "")
    # o carimbo de Lote numa Paralela antiga vai para a coluna W dela (a X é Observações naquele arquivo)
    ap._marcar_lote_na_paralela(antiga, [{"aba": "Lançamentos", "linha": 8}], "Transportado em 19/11/2026 08:00")
    wa = openpyxl.load_workbook(antiga)["Lançamentos"]
    assert wa["W8"].value == "Transportado em 19/11/2026 08:00" and wa["X8"].value is None
    assert tr.ler_paralela(antiga)[0]["lote"] == "Transportado em 19/11/2026 08:00"


def test_linha_transportada_leva_o_endereco_do_teor(tmp_path):
    c = str(tmp_path / "p.xlsx")
    gerar_paralela([_linha("5")], [], c, date(2026, 11, 19), {"ok": True})
    l = tr.ler_paralela(c)[0]
    assert tr.valores_agenda(l, "lote")[tr.COL["link"]] == "https://x/5"


def test_linhas_de_publicacao_centradas_na_vertical(tmp_path):
    c = str(tmp_path / "p.xlsx")
    gerar_paralela([_linha("1", despacho_anterior="x")], [_linha("2", destino="FORA DA AGENDA")], c, date(2026, 11, 19), {"ok": True},
                   [_reg("3", "0000003-11.2026.5.18.0003", date(2026, 11, 25))])
    wb = openpyxl.load_workbook(c)
    for aba, colunas in (("Lançamentos", 27), ("Fora da Agenda", 27), (pa.ABA, 13)):
        assert all(wb[aba].cell(8, i).alignment.vertical == "center" for i in range(1, colunas + 1)), aba
    assert wb["Lançamentos"]["K8"].alignment.horizontal == "left" and wb["Lançamentos"]["A8"].alignment.horizontal == "center"


def test_pendentes_tambem_ficam_centradas(cfg, tmp_path, monkeypatch):
    cfg["pc"] = dict(cfg.get("pc") or {}, pasta_controladoria=str(tmp_path))
    os.makedirs(tmp_path / "CACT - Agenda Paralela", exist_ok=True)
    monkeypatch.setattr(ap, "_livro_aberto_no_excel", lambda c: None)
    gerar_paralela([_linha("1")], [], ap.caminho_paralela(cfg, date(2026, 11, 18)), date(2026, 11, 18), {"ok": True})
    hoje = ap.caminho_paralela(cfg, date(2026, 11, 19))
    gerar_paralela([_linha("2")], [], hoje, date(2026, 11, 19), {"ok": True})
    assert len(ap.montar_pendentes(cfg, date(2026, 11, 19))) == 1
    ws = openpyxl.load_workbook(hoje)["Pendentes"]
    assert ws["W8"].value == "1" and all(ws.cell(8, i).alignment.vertical == "center" for i in range(1, 28))

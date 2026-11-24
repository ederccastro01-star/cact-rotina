# -*- coding: utf-8 -*-
"""Escrita do arquivo da Agenda Paralela no formato visual da Agenda oficial (A–K) + colunas de controle."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from cact_rotina import limpar_texto

COLUNAS_AGENDA = ["Prazo", "Dia/Sem", "Peça", "Publicado", "Processo ", "Reclamante", "Reclamado", "Advogado ",
                  "Prioridade ", "Cliente", "Despacho"]
# Desenho de 20/11/2026 (pedido do advogado): L = último despacho do processo em "Prazos Feitos" (coluna K da Agenda);
# o link do teor veio de U para O. As Paralelas anteriores têm o desenho antigo: quem lê acha a coluna pelo cabeçalho
# (pc.transporte.colunas_do_cabecalho), nunca pela posição.
COLUNAS_CONTROLE = ["Despacho anterior", "Destino", "Decisão", "Link (teor)", "Motivo CONFERIR", "Polo intimado",
                    "Intimado(s)", "Fonte do polo", "Prazo no texto", "Prazo legal", "Base legal", "ID DJEN", "Lote",
                    "Observações", "Prazo calculado", "Corrigir J para"]
COL = {nome: i for i, nome in enumerate(COLUNAS_AGENDA + COLUNAS_CONTROLE, start=1)}
C_DESP_ANT, C_DECISAO, C_LINK = COL["Despacho anterior"], COL["Decisão"], COL["Link (teor)"]
C_PRAZO_CALC, C_CORRIGIR_J = COL["Prazo calculado"], COL["Corrigir J para"]
# larguras como o Excel as MOSTRA (A a Q ditadas pelo advogado); no arquivo vai a largura mostrada + 5/7 (margem da célula)
LARGURAS_EXCEL = [7.71, 8.00, 13.29, 9.43, 20, 20, 20, 20, 11.00, 9.29, 33, 33, 11.43, 8.29, 8.00, 12.14, 8.29,
                  34, 16, 10, 10, 13, 22, 22, 30, 12, 14]
LARGURAS = [round(w + 5 / 7, 2) for w in LARGURAS_EXCEL]
DECISOES = '"Agendar,Descartar,Corrigir e agendar"'
OPCOES_J = '"AUTOR(A),RÉU"'
ORDEM_DESTINO = {"PROCESSO NOVO": 0, "PRAZO DO ESCRITÓRIO": 2, "DISPENSADA": 4, "PARTE CONTRÁRIA": 3, "FORA DA AGENDA": 5}

ARIAL = Font(name="Arial", size=10)
COURIER = Font(name="Courier New", size=10)
HDR_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="0070C0")
CTRL_FILL = PatternFill("solid", fgColor="595959")
FILL = {"PRAZO DO ESCRITÓRIO": "E2EFDA", "PARTE CONTRÁRIA": "F2F2F2", "PROCESSO NOVO": "FFF2CC", "CONFERIR": "FCE4D6",
        "DISPENSADA": "F2F2F2", "FORA DA AGENDA": "FFFFFF"}
AMARELO = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BORDA = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
# Linhas de publicação (pedido do advogado, 20/11/2026): texto corrido JUSTIFICADO (resumo, despacho anterior, motivo,
# observações); tudo o mais centralizado; sempre centro na vertical. A altura da linha NÃO é gravada: sem altura no
# arquivo o Excel ajusta cada linha ao texto quando abre (conferido numa instância do Excel em 20/11/2026).
JUST = Alignment(horizontal="justify", vertical="center", wrap_text=True)
COLS_JUST = {COL["Despacho"], C_DESP_ANT, COL["Motivo CONFERIR"], COL["Observações"]}

LEIA_ME = [
    ("O que é este arquivo", True),
    ("Agenda Paralela do dia, gerada pela rotina do CACT no computador do escritório, ao ligar, a partir das publicações do DJEN. Toda publicação de processo "
     "do escritório virou uma linha; nada foi gravado na Agenda oficial. As colunas A a K têm o formato da Agenda; de L em "
     "diante são controle. A coluna L repete o último despacho do processo na aba Prazos Feitos da Agenda.", False),
    ("", False),
    ("Como usar, de manhã", True),
    ("1. Aba Lançamentos: amarelo = PROCESSO NOVO; laranja = CONFERIR (há dúvida registrada na coluna P); verde = PRAZO DO "
     "ESCRITÓRIO; cinza = PARTE CONTRÁRIA, só para conhecimento.", False),
    ("2. Preencha a coluna N (Decisão): Agendar (A), Descartar (D) ou Corrigir e agendar (C). Em 'Corrigir e agendar', "
     "ajuste antes a data (A), a peça (C) ou a nota (K) na própria linha.", False),
    ("3. Salve e clique no atalho 'Transportar Paralela agora' (Área de Trabalho). O script do PC grava as linhas 'Agendar' e 'Corrigir e agendar' na aba Publicações da Agenda oficial, "
     "copiando A a K como estão, e carimba a coluna X (Lote) com 'Transportado em ...'.", False),
    ("4. Linha sem decisão não se perde: amanhã aparece na aba Pendentes e no aviso da manhã. Em Pendentes, processo que "
     "ainda NÃO CONSTA DA AGENDA vem no topo, em vermelho, com o aviso na linha 6 e a guia vermelha: agende primeiro.", False),
    ("5. Os contadores no alto da aba Lançamentos (L1 a M5) mostram o total de linhas e quantas já têm Agendar, Descartar e "
     "Corrigir e agendar; 'Falta Decisão' fica vermelho enquanto houver linha sem decisão.", False),
    ("", False),
    ("O que a máquina decide e o que você decide", True),
    ("Máquina: polo do intimado (DJEN), destino (polo x coluna J), data (dias úteis, feriados, tabela legal; na dúvida a "
     "menor), agrupamento de comunicações repetidas num só ato, conciliação (toda publicação tem linha). IA: só o rótulo "
     "da peça (C) e o resumo (K). Você: a decisão e qualquer ajuste.", False),
    ("", False),
    ("Colunas de controle", True),
    ("L Despacho anterior (Prazos Feitos, coluna K) · M Destino · N Decisão (você) · O Link para o teor · P Motivo CONFERIR · "
     "Q Polo intimado · R Intimado(s) · S Fonte do polo · T Prazo no texto · U Prazo legal · V Base legal · W ID DJEN · "
     "X Lote/Transportado · Y Observações · Z Prazo calculado · AA Corrigir J para (você).", False),
    ("", False),
    ("Aba Parceiros - Liquidação", True),
    ("Intimações para APRESENTAR CÁLCULOS DE LIQUIDAÇÃO em processos da carteira dos parceiros que ainda não estão na Agenda. "
     "A lista é acumulada: cada intimação fica até você tirar. Situação: 'Aguardando' (dentro do prazo), 'Chegou à Agenda em "
     "dd/mm' (o processo passou a constar na Agenda) ou 'PRAZO VENCIDO SEM CHEGAR'. A coluna Observações é sua e é levada "
     "para a Paralela do dia seguinte; escreva 'Baixar' nela para a linha sair da lista.", False),
]


@dataclass
class Linha:
    prazo: date | None
    peca: str
    publicado: date
    processo: str
    reclamante: str
    reclamado: str
    advogado: str
    prioridade: str
    cliente: str
    despacho: str
    destino: str
    motivo: str = ""
    polo_intimado: str = "?"
    intimados: str = ""
    fonte_polo: str = ""
    prazo_texto: int | None = None
    prazo_legal: int | None = None
    base_legal: str | None = None
    link: str = ""
    ids: str = ""
    lote: str = ""
    observacoes: str = ""
    despacho_anterior: str = ""


def _ordem(l: Linha):
    if l.destino == "PROCESSO NOVO":
        grupo = 0
    elif l.motivo and l.destino == "PRAZO DO ESCRITÓRIO":
        grupo = 1
    else:
        grupo = ORDEM_DESTINO.get(l.destino, 9)
    return (grupo, l.prazo or date.max, l.processo)


def _cabecalho(ws, titulo: str, subtitulo: str):
    ws["A1"] = titulo; ws["A1"].font = Font(name="Arial", size=12, bold=True)
    ws["A2"] = subtitulo; ws["A2"].font = Font(name="Arial", size=9, italic=True, color="595959")
    ws["A4"] = ("Colunas A a K: iguais à Agenda oficial. Colunas a partir de L: controle da Paralela "
                "(não vão para a Agenda, exceto ID DJEN, Link e Lote).")
    ws["A4"].font = Font(name="Arial", size=9, color="595959")
    ws["A5"] = ("Legenda: verde = PRAZO DO ESCRITÓRIO; cinza = PARTE CONTRÁRIA; amarelo = PROCESSO NOVO; laranja = CONFERIR. "
                "Você preenche só a coluna N (Decisão) e, se quiser, A/C/K e Y (Observações).")
    ws["A5"].font = Font(name="Arial", size=9, color="595959")
    for i, (h, w) in enumerate(zip(COLUNAS_AGENDA + COLUNAS_CONTROLE, LARGURAS), start=1):
        c = ws.cell(row=7, column=i, value=h)
        c.font = HDR_FONT; c.fill = HDR_FILL if i <= 11 else CTRL_FILL; c.alignment = CENTRO; c.border = BORDA
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[7].height = 30
    ws.freeze_panes = "B8"      # fixa só a coluna A (Prazo) e o cabeçalho (pedido do advogado, 20/11/2026; antes F8)


def _escrever(ws, row: int, l: Linha):
    cor = FILL["CONFERIR"] if (l.motivo and l.destino == "PRAZO DO ESCRITÓRIO") else FILL.get(l.destino, "FFFFFF")
    fill = PatternFill("solid", fgColor=cor)
    vals = [l.prazo, None, l.peca, l.publicado, l.processo, l.reclamante, l.reclamado, l.advogado, l.prioridade, l.cliente,
            l.despacho, l.despacho_anterior, l.destino, None, l.link, l.motivo, l.polo_intimado, l.intimados, l.fonte_polo,
            l.prazo_texto, l.prazo_legal, l.base_legal, l.ids, l.lote, l.observacoes,
            l.prazo,       # Z = prazo calculado (fica mesmo se A for corrigida)
            None]          # AA = Corrigir J para (o advogado preenche)
    for i, v in enumerate(vals, start=1):
        if isinstance(v, str):
            v = limpar_texto(v)
        c = ws.cell(row=row, column=i, value=v)
        c.font = COURIER if i == 5 else ARIAL; c.border = BORDA; c.fill = fill
        c.alignment = JUST if i in COLS_JUST else CENTRO
        if i in (1, 4, C_PRAZO_CALC):
            c.number_format = "dd/mm/yy;@"
        if i == 2:
            c.value = f'=IF(A{row}="","",A{row})'; c.number_format = "ddd"
        if i == 5:
            c.number_format = "@"
        if i == C_LINK and v:
            c.hyperlink = v; c.value = "abrir teor"
            c.font = Font(name="Arial", size=9, color="0563C1", underline="single")
    ws.cell(row=row, column=C_DECISAO).fill = AMARELO
    ws.cell(row=row, column=C_CORRIGIR_J).fill = AMARELO


def _aba(wb, titulo, subtitulo, nome, linhas: list[Linha], ordenar=True):
    ws = wb.create_sheet(nome)
    _cabecalho(ws, titulo, subtitulo)
    row = 8
    for l in (sorted(linhas, key=_ordem) if ordenar else linhas):
        _escrever(ws, row, l); row += 1
    if linhas:
        dv = DataValidation(type="list", formula1=DECISOES, allow_blank=True)
        dv.prompt = ("Agendar: o script grava na Agenda oficial. Descartar: só registro. "
                     "Corrigir e agendar: corrija A, C ou K nesta linha antes.")
        ld = get_column_letter(C_DECISAO); ws.add_data_validation(dv); dv.add(f"{ld}8:{ld}{row - 1}")
        dvj = DataValidation(type="list", formula1=OPCOES_J, allow_blank=True)
        dvj.prompt = "Lado que o escritório representa neste processo. O transporte corrige a coluna J em todas as linhas do processo na Agenda e marca 'Polo conferido'."
        lj = get_column_letter(C_CORRIGIR_J); ws.add_data_validation(dvj); dvj.add(f"{lj}8:{lj}{row - 1}")
    return ws, row


# Contadores da aba Lançamentos (pedido do advogado, 24/11/2026): linhas 1 a 5, rótulo em L e número em M. São fórmulas:
# mudam sozinhos quando a coluna N (Decisão) é preenchida. Intervalo largo para contar também as linhas que a conferência
# PUSH acrescenta depois. A linha da conciliação, embaixo da tabela, fica na coluna A e não entra na conta.
ULTIMA_LINHA_CONTADOR = 5000
CONTADOR_FILL = PatternFill("solid", fgColor="DDEBF7")


def _contadores(ws) -> None:
    e = f"$E$8:$E${ULTIMA_LINHA_CONTADOR}"
    n = f"${get_column_letter(C_DECISAO)}$8:${get_column_letter(C_DECISAO)}${ULTIMA_LINHA_CONTADOR}"
    itens = [("Total das Publicações", f"=COUNTA({e})"),
             ("Agendar", f'=COUNTIF({n},"Agendar")'),
             ("Descartar", f'=COUNTIF({n},"Descartar")'),
             ("Corrigir e agendar", f'=COUNTIF({n},"Corrigir e agendar")'),
             ("Falta Decisão", f'=COUNTIFS({e},"<>",{n},"")')]
    c_rot, c_num = C_DESP_ANT, C_DESP_ANT + 1          # L e M
    for i, (rotulo, formula) in enumerate(itens, start=1):
        r = ws.cell(row=i, column=c_rot, value=rotulo)
        r.font = Font(name="Arial", size=10, bold=True); r.fill = CONTADOR_FILL; r.border = BORDA
        r.alignment = Alignment(horizontal="right", vertical="center")
        v = ws.cell(row=i, column=c_num, value=formula)
        v.font = Font(name="Arial", size=11, bold=True); v.fill = CONTADOR_FILL; v.border = BORDA; v.alignment = CENTRO
    from openpyxl.formatting.rule import CellIsRule
    alvo = f"{get_column_letter(c_num)}5"
    ws.conditional_formatting.add(alvo, CellIsRule(operator="greaterThan", formula=["0"],
                                                   font=Font(name="Arial", size=11, bold=True, color="C00000"),
                                                   fill=PatternFill(fill_type="solid", start_color="FFC7CE", end_color="FFC7CE")))   # na formatação condicional o Excel lê o fundo em bgColor (end_color)


def gerar_paralela(lancamentos: list[Linha], fora: list[Linha], caminho: str, data_disp: date, conciliacao: dict,
                   parceiros: list | None = None, nota_parceiros: str = "") -> str:
    """`parceiros`: registros da aba "Parceiros - Liquidação" (lista acumulada). None = arquivo sem a aba (Paralela criada
    pelo PC no fim de semana, ou falha na montagem): a rotina do dia seguinte procura a lista na Paralela anterior."""
    wb = openpyxl.Workbook()
    wl = wb.active; wl.title = "Leia-me"
    for i, (t, b) in enumerate(LEIA_ME, start=1):
        c = wl.cell(row=i, column=1, value=t)
        c.font = Font(name="Arial", size=11 if b else 10, bold=b); c.alignment = Alignment(wrap_text=True, vertical="top")
    wl.column_dimensions["A"].width = 140

    ws, row = _aba(wb, f"CACT - Agenda Paralela {data_disp:%d-%m-%Y}",
                   "Gerada pela rotina do CACT ao ligar o computador. Ordem: PROCESSO NOVO e CONFERIR no topo; depois PRAZO DO ESCRITÓRIO "
                   "por data; PARTE CONTRÁRIA por último. Polo e data por código; peça e resumo pela IA.",
                   "Lançamentos", lancamentos)
    _contadores(ws)
    n_esc = sum(1 for l in lancamentos if l.destino == "PRAZO DO ESCRITÓRIO")
    n_con = sum(1 for l in lancamentos if l.destino == "PARTE CONTRÁRIA")
    n_novo = sum(1 for l in lancamentos if l.destino == "PROCESSO NOVO")
    n_conf = sum(1 for l in lancamentos if l.motivo and l.destino == "PRAZO DO ESCRITÓRIO")
    ws.cell(row=row + 1, column=1, value=(
        f"Conciliação: {conciliacao.get('atos', len(lancamentos) + len(fora))} atos de {conciliacao.get('processos', '?')} processos, "
        f"{conciliacao.get('ids', '?')} comunicações; {n_esc} prazo do escritório, {n_con} parte contrária, {n_novo} processo novo, "
        f"{n_conf} CONFERIR; {len(fora)} fora da Agenda.")).font = Font(name="Arial", size=9, italic=True)
    if not conciliacao.get("ok", True):
        faltam = ", ".join(str(x) for x in conciliacao.get("faltam", []))[:300]
        c = ws["A6"]
        c.value = f"ATENÇÃO: conciliação falhou — comunicações sem linha: {faltam}"
        c.font = Font(name="Arial", size=10, bold=True, color="C00000")

    _aba(wb, "Fora da Agenda — carteira dos parceiros em processos em que o escritório ainda não atuou",
         "Só para consulta; sem decisão obrigatória. Se um deles virar trabalho, lance na Agenda como hoje e a partir daí "
         "ele entra em Lançamentos.", "Fora da Agenda", fora)
    _aba(wb, "Pendentes — linhas de dias anteriores sem decisão",
         "Preenchida pelo script do PC a partir das Paralelas anteriores. Some quando a decisão é dada.", "Pendentes", [])
    if parceiros is not None:
        from paralela.parceiros import escrever_aba
        escrever_aba(wb, parceiros, data_disp, nota_parceiros)
    wb.save(caminho)
    return caminho

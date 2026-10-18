# -*- coding: utf-8 -*-
"""Parsing puro dos e-mails de movimentação processual (spec §9c): PJe Push, TRT-Push legado, Jusbrasil (layout
antigo em linha única e layout novo em blocos) e outros avisos de tribunal (PROJUDI etc.).

Não toca em Outlook nem em arquivo: recebe um `Email` e devolve `Registro`s, um por evento/publicação, no
vocabulário das 18 colunas do `Relatorio_PUSH_*.xlsx` que `pc/push.py` já consome.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime

CNJ = re.compile(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}")
ENDERECO = re.compile(r"[\w.+-]+@[\w.-]+\.\w+")
DATA_HORA = re.compile(r"^(\d{2}/\d{2}/\d{4} \d{2}:\d{2}(?::\d{2})?)\s+(.*)$")
URL = re.compile(r"<\s*https?://[^>]*>|https?://\S+")

ORIGEM_PJE = "PJe Push"
ORIGEM_LEGADO = "TRT-Push (legado)"
ORIGEM_JUSBRASIL = "Jusbrasil"
ORIGEM_OUTROS = "Outros"

REMETENTES_JUSBRASIL = ("publicacoes-diarios@jusbrasil.com.br", "andamentos@jusbrasil.com.br")
MARKETING = ("nao-responda@jusbrasil.com.br", "@descubra.jusbrasil.com.br", "ola@jusbrasil.com.br")
ASSUNTOS_ELEGIVEIS = ("[PUSH]", "PUSH - ANDAMENTO PROCESSUAL", "NOVAS PUBLICA", "NOVA PUBLICA", "[PROJUDI]")

RODAPE = ("Para acessar este processo", "ATENÇÃO: este e-mail", "Caso não tenha mais interesse", "Analisar publicação",
          "Atenciosamente", "08006206622", "AVENIDA T1", "Ler publicação na íntegra", "Jus IA ●", "Próximos passos",
          "Obs.: Informa", "Aten��o: Este", "Atenção: Este é um e-mail", "Ocorreu um erro ou voc",
          "Sistema PROJUDI", "Esta é um mensagem automática", "Por favor não responder")

# ordem importa: os dois primeiros achados formam o tipo ("Intimação / Expedição de documento")
TIPOS = [
    ("Sentença", r"\bSENTEN[ÇC]A\b"), ("Acórdão", r"\bAC[ÓO]RD[ÃA]O\b"), ("Despacho", r"\bDESPACHO\b"),
    ("Decisão", r"\bDECIS[ÃA]O\b"), ("Pauta/Audiência", r"\bAUDI[ÊE]NCIA\b|\bPAUTA\b|\bSESS[ÃA]O DE JULGAMENTO\b"),
    ("Intimação", r"\bINTIMA[ÇC][ÃA]O\b|\bINTIME-SE\b|\bINTIMAD"), ("Citação", r"\bCITA[ÇC][ÃA]O\b|\bCITE-SE\b"),
    ("Perícia", r"\bPER[ÍI]CIA\b|\bPERITO\b|\bLAUDO\b"),
    ("Recurso", r"\bRECURSO ORDIN[ÁA]RIO\b|\bRECURSO DE REVISTA\b|\bAGRAVO\b|\bEMBARGOS\b|\bRECURSO\b"),
    ("Execução", r"\bEXECU[ÇC][ÃA]O\b|\bPENHORA\b|\bALVAR[ÁA]\b|\bBACENJUD\b|\bSISBAJUD\b"),
    ("Publicação/Disponibilização", r"\bPUBLICAD|\bDISPONIBILIZAD"), ("Decurso de prazo", r"\bDECORRIDO O PRAZO\b|\bDECURSO DE PRAZO\b"),
    ("Juntada de petição", r"\bJUNTAD[AO]\b"), ("Conclusão", r"\bCONCLUS[ÃA]O\b|\bCONCLUSOS\b"),
    ("Expedição de documento", r"\bEXPEDI[ÇC][ÃA]O\b|\bEXPEDID"),
    ("Remessa/Distribuição", r"\bREMESSA\b|\bREMETID|\b(?:RE)?DISTRIBU(?:[IÍ]D[OA]S?\b|I[ÇC][ÃA]O\b)"),   # não "DISTRIBUIDORA" (nome de parte)
    ("Arquivamento", r"\bARQUIV"), ("Documento sigiloso", r"\bDOCUMENTO SIGILOSO\b|\bSIGILOSO\b"),
]
_TIPOS = [(nome, re.compile(rx)) for nome, rx in TIPOS]

# acentuação corrompida na origem (TRT-Push legado): só o que é inequívoco; o resto fica como veio
_U = "�"
MOJIBAKE = {
    "informa" + _U + _U + "es": "informações", "intima" + _U + _U + "o": "intimação", "cita" + _U + _U + "o": "citação",
    "dispon" + _U + "veis": "disponíveis", "altera" + _U + _U + "es": "alterações", "aten" + _U + _U + "o": "atenção",
    "autom" + _U + "tico": "automático", "permiss" + _U + "o": "permissão", "regi" + _U + "o": "região",
    "decis" + _U + "o": "decisão", "certid" + _U + "o": "certidão", "peti" + _U + _U + "o": "petição",
    "senten" + _U + "a": "sentença", "audi" + _U + "ncia": "audiência", "ac" + _U + "rd" + _U + "o": "acórdão",
    "execu" + _U + _U + "o": "execução", "conclus" + _U + "o": "conclusão", "publica" + _U + _U + "o": "publicação",
    "manifesta" + _U + _U + "o": "manifestação", "per" + _U + "cia": "perícia", "distribui" + _U + _U + "o": "distribuição",
    "expedi" + _U + _U + "o": "expedição", "liquida" + _U + _U + "o": "liquidação", "homologa" + _U + _U + "o": "homologação",
    "n" + _U + "o": "não", "voc" + _U: "você", "at" + _U: "até", "j" + _U: "já", "v.s" + _U: "v.sª",
    "18" + _U: "18ª", "ordin" + _U + "rio": "ordinário", "sum" + _U + "rio": "sumário", "jur" + _U + "dico": "jurídico",
    "p" + _U + "blico": "público", "p" + _U + "blica": "pública", "eletr" + _U + "nico": "eletrônico",
}
_MOJIBAKE = [(re.compile(r"\b" + re.escape(k) + r"(?!\w)", re.I), v) for k, v in MOJIBAKE.items()]


@dataclass
class Email:
    remetente: str
    assunto: str
    corpo: str
    data: datetime
    message_id: str
    conta: str = ""
    pasta: str = ""


@dataclass
class Registro:
    data: datetime | None
    autuacao: datetime | None
    processo: str
    tribunal: str
    ramo: str
    grau: str
    classe: str
    orgao: str
    tipo: str
    teor: str
    autor: str
    reu: str
    origem: str
    data_email: datetime
    assunto: str
    remetente: str
    message_id: str


# ----------------------------------------------------------------------------- utilidades
def _linhas(corpo: str) -> list[str]:
    return [re.sub(r"[ \t ]+", " ", l).strip() for l in corpo.replace("\r\n", "\n").replace("\r", "\n").split("\n")]


def _flat(texto: str) -> str:
    return re.sub(r"\s+", " ", texto or "").strip()


def _dt(texto: str) -> datetime | None:
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(texto.strip(), fmt)
        except ValueError:
            continue
    return None


def remetente_real(remetente: str, corpo: str) -> str:
    """E-mail encaminhado (Hotmail → Gmail, parceiro → Hotmail): o remetente verdadeiro está no 'De:' do corpo."""
    for linha in _linhas(corpo)[:80]:
        m = re.match(r"^(?:De|From):\s*(.*)$", linha, re.I)
        if m:
            e = ENDERECO.search(m.group(1))
            if e:
                return e.group().lower()
    return (remetente or "").lower()


def normalizar_acentos(texto: str) -> str:
    if _U not in (texto or ""):
        return texto
    for rx, certo in _MOJIBAKE:
        def _sub(m, certo=certo):
            return certo.capitalize() if m.group()[0].isupper() else certo
        texto = rx.sub(_sub, texto)
    return texto


def limpar_rodape(texto: str) -> str:
    corte = len(texto)
    for marca in RODAPE:
        i = texto.find(marca)
        if 0 <= i < corte:
            corte = i
    return texto[:corte].rstrip()


def tipo_andamento(texto: str) -> str:
    up = (texto or "").upper()
    achados = [nome for nome, rx in _TIPOS if rx.search(up)]
    return " / ".join(dict.fromkeys(achados[:2])) if achados else "Outros"


def ramo(sigla: str) -> str:
    s = (sigla or "").upper()
    if s.startswith(("TST", "STJ", "STF", "TSE", "STM")):
        return "Tribunal Superior"
    if s.startswith("TRT"):
        return "Justiça do Trabalho"
    if s.startswith("TRF"):
        return "Justiça Federal"
    if s.startswith("TJ"):
        return "Justiça Estadual"
    return "Outros"


_DOMINIO_SIGLA = {"trtsp": "TRT02", "tst": "TST", "tjgo": "TJGO"}


def sigla_tribunal(remetente: str, assunto: str, corpo: str) -> str:
    m = re.search(r"\[(TRT\d{1,2}|TST|STJ|STF|TRF\d|TJ[A-Z]{2})\]", assunto or "", re.I)
    if m:
        return m.group(1).upper()
    if re.search(r"Tribunal Superior do Trabalho", corpo or "", re.I):
        return "TST"
    m = re.search(r"Trabalho\s*(?:da|-)\s*(\d{1,2})", corpo or "")
    if m:
        return f"TRT{int(m.group(1)):02d}" if len(m.group(1)) == 2 else f"TRT{m.group(1)}"
    m = re.search(r"@(?:mail\.)?([a-z]+\d*)\.jus\.br", remetente or "", re.I)
    if m:
        dom = m.group(1).lower()
        return _DOMINIO_SIGLA.get(dom, dom.upper())
    return ""


def _grau(corpo: str) -> str:
    m = re.search(r"(\d)[ºo°]\s*Grau", corpo or "")
    if m:
        return f"{m.group(1)}º Grau"
    m = re.search(r"p_grau(?:_pje)?=(\d)", corpo or "")
    return f"{m.group(1)}º Grau" if m else ""


def origem(remetente_real_: str, assunto: str, corpo: str) -> str:
    rem = (remetente_real_ or "").lower()
    if "trtpush@" in rem or "sistema TRT-Push" in (corpo or ""):
        return ORIGEM_LEGADO
    if re.search(r"(nao-?responda|naoresponder)[\w.-]*@[\w.-]*jus\.br", rem) or \
            "Acompanhamento de Atualizações de Informações Processuais" in (corpo or "") or "[PUSH]" in (assunto or "").upper():
        return ORIGEM_PJE
    if "jusbrasil" in rem:
        return ORIGEM_JUSBRASIL
    return ORIGEM_OUTROS


def elegivel(email: Email) -> bool:
    rem = remetente_real(email.remetente, email.corpo)
    if any(m in rem for m in MARKETING):
        return False
    o = origem(rem, email.assunto, email.corpo)
    if o == ORIGEM_JUSBRASIL:
        return rem in REMETENTES_JUSBRASIL
    if o in (ORIGEM_PJE, ORIGEM_LEGADO):
        return True
    up = (email.assunto or "").upper()
    return rem.endswith("jus.br") or any(a in up for a in ASSUNTOS_ELEGIVEIS)


def anunciadas(assunto: str) -> int:
    m = re.search(r"(\d+)\s+novas?\s+publica", assunto or "", re.I)
    if m:
        return int(m.group(1))
    return 1 if re.search(r"nova publica", assunto or "", re.I) else 0


def faltantes(email: Email) -> int:
    m = re.search(r"Mais (\d+) publica", email.corpo or "", re.I)
    return int(m.group(1)) if m else 0


# ----------------------------------------------------------------------------- extratores
def _base(email: Email, rem: str, orig: str, **campos) -> Registro:
    valores = dict(data=None, autuacao=None, processo="", tribunal="", ramo="", grau="", classe="", orgao="", tipo="", teor="",
                   autor="", reu="", origem=orig, data_email=email.data, assunto=email.assunto or "", remetente=rem,
                   message_id=email.message_id)
    valores.update(campos)
    if not valores["ramo"]:
        valores["ramo"] = ramo(valores["tribunal"])
    if not valores["tipo"]:
        valores["tipo"] = tipo_andamento(valores["teor"])
    return Registro(**valores)


def _campo(texto: str, rotulo: str) -> str:
    m = re.search(r"^" + rotulo + r":\s*(.+)$", texto, re.M)
    return _flat(m.group(1)) if m else ""


def _pje(email: Email, rem: str) -> list[Registro]:
    linhas = _linhas(email.corpo)
    texto = "\n".join(linhas)
    m = re.search(r"Número do Processo:\s*(" + CNJ.pattern + ")", texto) or CNJ.search(email.assunto or "")
    processo = m.group(1) if m and m.lastindex else (m.group() if m else "")
    comum = dict(processo=processo, tribunal=sigla_tribunal(rem, email.assunto, texto), grau=_grau(texto),
                 classe=_campo(texto, "Classe Judicial"), orgao=_campo(texto, "Órgão Julgador"),
                 autuacao=_dt(_campo(texto, "Data de Autuação")), autor=_campo(texto, "Autor"), reu=_campo(texto, "Réu"))
    eventos: list[tuple[datetime | None, str]] = []
    dentro = False
    for l in linhas:
        if not dentro:
            dentro = l.startswith("Eventos")
            continue
        if not l or l in ("Data Evento", "Data", "Evento"):
            continue
        if any(l.startswith(r) for r in RODAPE):
            break
        m = DATA_HORA.match(l)
        if m:
            eventos.append((_dt(m.group(1)), m.group(2).strip()))
        elif eventos:
            eventos[-1] = (eventos[-1][0], (eventos[-1][1] + " " + l).strip())
    return [_base(email, rem, ORIGEM_PJE, data=d, teor=limpar_rodape(t), **comum) for d, t in eventos]


def _legado(email: Email, rem: str) -> list[Registro]:
    linhas = _linhas(email.corpo)
    tribunal = sigla_tribunal(rem, email.assunto, email.corpo)
    regs: list[Registro] = []
    bloco: dict | None = None
    for bruto in linhas:
        if bruto.startswith("Processo:"):
            m = CNJ.search(bruto)
            bloco = dict(processo=m.group() if m else re.sub(r"^Processo:\s*", "", URL.sub("", bruto)).strip(),
                         grau=_grau(bruto), autor="", reu="")
            continue
        if bloco is None:
            continue
        l = normalizar_acentos(URL.sub("", bruto)).strip()
        if l.startswith("AUTOR:"):
            bloco["autor"] = _flat(l[6:])
        elif re.match(r"^R[ÉE]U:", l):
            bloco["reu"] = _flat(l[4:])
        elif DATA_HORA.match(l):
            m = DATA_HORA.match(l)
            regs.append(_base(email, rem, ORIGEM_LEGADO, data=_dt(m.group(1)), teor=limpar_rodape(m.group(2)), tribunal=tribunal, **bloco))
        elif any(l.startswith(r) for r in ("Obs.:", "Aten", "Tribunal Regional")):
            bloco = None
    return regs


_JUS_ANTIGO = re.compile(
    r"Processo (" + CNJ.pattern + r") Termos encontrados (.*?) Data de disponibilização (\d{2}/\d{2}/\d{2}) "
    r"Data de publicação (\d{2}/\d{2}/\d{2}) Diário (\S+) · (.*?) Publicação (.*?)"
    r"(?= Analisar publicação| Processo \d{7}-\d{2}\.| Atenciosamente| Este e-mail foi enviado|$)")
_JUS_NOVO_CAB = re.compile(r"^([A-Z]{2,5}\d{0,2}) · (.+?) (\d{2}/\d{2}/\d{2})$")


def _orgao_jusbrasil(teor: str) -> str:
    i = teor.find("NÚMERO ÚNICO")
    if i < 0:
        m = CNJ.search(teor)
        i = m.start() if m else -1
    return teor[:i].strip(" :-") [:200] if i > 0 else ""


def _partes(texto: str) -> tuple[str, str]:
    m = re.search(r"POLO ATIVO (.*?) POLO PASSIVO (.*?)(?: ADVOGADO|$)", URL.sub("", texto))
    return (_flat(m.group(1)), _flat(m.group(2))) if m else ("", "")


def _jusbrasil_antigo(email: Email, rem: str) -> list[Registro]:
    flat = _flat(email.corpo)
    regs = []
    for m in _JUS_ANTIGO.finditer(flat):
        teor = limpar_rodape(m.group(7).strip())
        autor, reu = _partes(teor)
        regs.append(_base(email, rem, ORIGEM_JUSBRASIL, processo=m.group(1), data=_dt(m.group(4)), tribunal=m.group(5).upper(),
                          orgao=_orgao_jusbrasil(teor), teor=teor, autor=autor, reu=reu))
    return regs


def _jusbrasil_novo(email: Email, rem: str) -> list[Registro]:
    linhas = [l for l in _linhas(email.corpo) if l]
    regs = []
    i = 0
    while i < len(linhas):
        if re.match(r"^Mais \d+ publica", linhas[i]):
            break
        cab = _JUS_NOVO_CAB.match(linhas[i])
        if not cab:
            i += 1
            continue
        sigla, orgao, data = cab.group(1), cab.group(2).strip(), _dt(cab.group(3))
        j = i + 1
        processo = autor = reu = ""
        teor_partes: list[str] = []
        while j < len(linhas) and not _JUS_NOVO_CAB.match(linhas[j]) and not re.match(r"^Mais \d+ publica", linhas[j]):
            l = linhas[j]
            if l.startswith("Ler publicação"):
                j += 1
                break
            mp = re.match(r"^(?:(.+?) x (.+?) · )?(" + CNJ.pattern + r")$", l)
            if mp and not processo:
                autor, reu, processo = (mp.group(1) or "").strip(), (mp.group(2) or "").strip(), mp.group(3)
            elif processo:
                teor_partes.append(l)
            j += 1
        if processo:
            teor = limpar_rodape(" ".join(teor_partes))
            regs.append(_base(email, rem, ORIGEM_JUSBRASIL, processo=processo, data=data, tribunal=sigla.upper(),
                              orgao=orgao[:200], teor=teor, autor=autor, reu=reu))
        i = j
    return regs


def _outros(email: Email, rem: str) -> list[Registro]:
    texto = _flat(email.corpo)
    m = CNJ.search(texto) or CNJ.search(email.assunto or "")
    processo = m.group() if m else ""
    if not processo:
        m = re.search(r"processo n[ºo°]?\.?\s*([\d][\d.\-/]*\d)", texto, re.I)
        processo = m.group(1) if m else ""
    teor = limpar_rodape(_flat(URL.sub("", email.corpo)))
    return [_base(email, rem, ORIGEM_OUTROS, processo=processo, tribunal=sigla_tribunal(rem, email.assunto, email.corpo),
                  data=email.data, teor=teor[:2000])]


def extrair(email: Email) -> list[Registro]:
    rem = remetente_real(email.remetente, email.corpo)
    o = origem(rem, email.assunto, email.corpo)
    if o == ORIGEM_PJE:
        return _pje(email, rem)
    if o == ORIGEM_LEGADO:
        return _legado(email, rem)
    if o == ORIGEM_JUSBRASIL:
        return _jusbrasil_antigo(email, rem) if "Termos encontrados" in email.corpo else _jusbrasil_novo(email, rem)
    return _outros(email, rem)

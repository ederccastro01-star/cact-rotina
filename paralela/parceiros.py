# -*- coding: utf-8 -*-
"""Aba "Parceiros - Liquidação": intimações para APRESENTAR CÁLCULOS DE LIQUIDAÇÃO em processos da carteira dos
parceiros que ainda não estão na Agenda (pedido do advogado em 20/11/2026: saber se o parceiro está passando o
serviço a terceiros ou deixando de encaminhá-lo).

A lista é ACUMULADA: a rotina lê a aba da Paralela anterior, junta as intimações do dia e recalcula a situação de
cada linha contra a Agenda ("Aguardando", "Chegou à Agenda (visto em dd/mm)", "PRAZO VENCIDO SEM CHEGAR").
Quem decide se o ato manda apresentar cálculos é a IA, só sobre os candidatos de um filtro por palavras (em
19/11/2026 o filtro sozinho marcou 31 de 117 processos, quase todos falsos: aviso padrão do CEJUSC, sentença que só
cita a liquidação). Sem IA a linha entra pela regra estreita, marcada "sem IA: conferir".
"""
from __future__ import annotations

import json
import os
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from cact_rotina import limpar_texto

ABA = "Parceiros - Liquidação"
PRIMEIRA_LINHA = 8
AGUARDANDO = "Aguardando"
VENCIDO = "PRAZO VENCIDO SEM CHEGAR"
CHEGOU = "Chegou à Agenda"
BAIXAR = "baixar"
PECA = "Apresentar cálculos de liquidação"
PRAZO_PADRAO = 8                       # CLT 879 §1º-B, quando o texto não fixa prazo
ORIGEM_IA = "IA"
ORIGEM_REGRA = "sem IA: conferir"

CABECALHOS = ["Situação", "Detectado em", "Prazo", "Processo", "Reclamante", "Reclamado", "Parceiro (advogado intimado)",
              "Polo intimado", "Resumo da intimação", "Link (teor)", "Observações (suas)", "ID DJEN", "Origem"]
LARGURAS = [24, 11, 11, 26, 26, 26, 30, 10, 60, 10, 34, 20, 14]
CAMPOS = ["situacao", "detectado", "prazo", "processo", "reclamante", "reclamado", "parceiro", "polo", "resumo", "link",
          "observacoes", "ids", "origem"]

# ----------------------------------------------------------------------------- filtro por palavras
_VERBO = r"(?:apresent|elabor|junt|traz|trag|forne|proced|promov|liquid)\w*"
RE_CANDIDATO = re.compile(
    _VERBO + r"[^.;]{0,160}\bcalculos?\b|\bcalculos?\b[^.;]{0,80}\bliquidac|\bliquidac\w+\s+(?:por|de|da|do)\s+(?:calculos?|sentenca|julgado)"
    r"|\bart(?:igo)?\.?\s*879\b|\bmemoria\s+(?:de|do)\s+calculos?")
# regra estreita (só quando a IA não respondeu): ordem direta ao intimado para apresentar os cálculos
RE_ORDEM = re.compile(
    r"(?:intim\w+|notifi\w+|fic\w+\s+intimad\w+|devera\w*|apresent\w+)[^.;]{0,200}\b(?:apresent|elabor|junt|traz|trag)\w*[^.;]{0,80}"
    r"\bcalculos?\s+(?:de\s+)?liquidac|(?:apresent|elabor|junt)\w*[^.;]{0,60}\bcalculos?\s+de\s+liquidac[^.;]{0,120}\bprazo\b")
RE_RUIDO = re.compile(r"sugere-se\s+as\s+partes\s+que\s+procedam\s+a\s+liquidac[^.]*\.?")

PROMPT = """Você é assistente do CACT (Castro Advocacia & Cálculos Trabalhistas), escritório que elabora cálculos de liquidação para advogados parceiros. Recebe trechos de uma intimação publicada no DJEN e devolve SÓ um JSON.

PERGUNTA: este ato INTIMA a parte (ou as partes) a APRESENTAR, ELABORAR ou JUNTAR CÁLCULOS DE LIQUIDAÇÃO agora (liquidação de sentença por cálculos, CLT art. 879 e §1º-B; cumprimento de sentença em que o juiz manda a parte trazer a memória de cálculo)?

CONTA (true): "intime-se o reclamante para apresentar os cálculos de liquidação em 8 dias"; "apresentem as partes os cálculos"; "à parte autora para liquidar o julgado"; ordem para apresentar cálculos retificados ou adequados ao acórdão.
NÃO CONTA (false): sentença ou acórdão que só diz que a liquidação será feita por cálculos no futuro; intimação para se manifestar sobre cálculos já apresentados ou impugná-los (art. 879 §2º); cálculos a cargo da contadoria, da secretaria ou de perito; sugestão genérica do CEJUSC para que as partes liquidem antes da audiência; homologação de cálculos; mera atualização do débito pela secretaria; qualquer outro ato.

Responda SOMENTE com o JSON:
{"apresentar_calculos": true|false, "quem": "autor|réu|ambas|outro", "prazo_dias": null|int, "resumo": "até 250 caracteres: quem foi intimado, o que deve apresentar e em que prazo"}"""


def _asc(s) -> str:
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()


def candidato(texto: str) -> bool:
    return bool(RE_CANDIDATO.search(RE_RUIDO.sub(" ", _asc(texto))))


def ordem_direta(texto: str) -> bool:
    return bool(RE_ORDEM.search(RE_RUIDO.sub(" ", _asc(texto))))


def trechos(texto: str, limite: int = 9000) -> str:
    """Início do ato + janelas em volta de cada menção a cálculos (sentença longa não cabe inteira no pedido)."""
    plano = re.sub(r"[ \t]+", " ", str(texto or ""))
    if len(plano) <= limite:
        return plano
    asc = _asc(plano)                      # NFKD + ascii pode encurtar o texto: as posições são aproximadas, a janela é larga
    faixas = [(0, 1500)]
    for m in RE_CANDIDATO.finditer(asc):
        faixas.append((max(0, m.start() - 1200), min(len(plano), m.end() + 1200)))
    faixas.sort()
    unidas = [list(faixas[0])]
    for a, b in faixas[1:]:
        if a <= unidas[-1][1]:
            unidas[-1][1] = max(unidas[-1][1], b)
        else:
            unidas.append([a, b])
    return "\n[...]\n".join(plano[a:b] for a, b in unidas)[:limite]


# ----------------------------------------------------------------------------- IA
def classificar_ia(texto: str, polo_intimado: str, cfg: dict) -> dict | None:
    """{"apresentar_calculos", "quem", "prazo_dias", "resumo"} ou None (sem chave, erro, recusa, resposta ilegível)."""
    from cact_rotina import com_tentativas, log
    from paralela import ia
    pergunta = f"Polo do intimado: {polo_intimado or 'desconhecido'}\n\nTexto da intimação:\n{trechos(texto)}"
    if ia.via_claude_code(cfg):
        conteudo = ia.chamar_claude_code(PROMPT, pergunta, cfg, "Claude Code (Parceiros)")
        return _ler_resposta(conteudo) if conteudo else None
    chave = (cfg.get("anthropic_api_key") or "").strip()
    if not chave:
        return None
    corpo = {
        "model": cfg.get("modelo_claude", "claude-opus-5-5"),
        "max_tokens": ia.MAX_TOKENS,
        "output_config": {"effort": "low"},
        "system": PROMPT,
        "messages": [{"role": "user", "content": pergunta}],
    }
    headers = {"x-api-key": chave, "anthropic-version": "2023-06-01", "content-type": "application/json"}
    try:
        r = com_tentativas(lambda: ia._post("https://api.anthropic.com/v1/messages", headers=headers, json=corpo,
                                            timeout=cfg.get("timeout_segundos", 60)), "Claude (Parceiros)")
        if r.status_code != 200:
            log(f"  ! Claude (Parceiros) HTTP {r.status_code}: {str(getattr(r, 'text', ''))[:160]}")
            return None
        dados = r.json()
        if dados.get("stop_reason") in ("refusal", "max_tokens"):
            log(f"  ! Claude (Parceiros) sem resposta útil ({dados.get('stop_reason')}).")
            return None
        return _ler_resposta("\n".join(b.get("text", "") for b in dados.get("content", []) if b.get("type") == "text"))
    except Exception as e:  # noqa: BLE001
        log(f"  ! Claude (Parceiros) erro: {e}")
        return None


def _ler_resposta(conteudo: str) -> dict | None:
    m = re.search(r"\{.*\}", conteudo or "", re.S)
    try:
        d = json.loads(m.group(0)) if m else None
    except json.JSONDecodeError:
        return None
    if not isinstance(d, dict) or "apresentar_calculos" not in d:
        return None
    try:
        prazo = int(d.get("prazo_dias")) if d.get("prazo_dias") not in (None, "") else None
    except (TypeError, ValueError):
        prazo = None
    return {"apresentar_calculos": d["apresentar_calculos"] is True, "quem": str(d.get("quem") or ""),
            "prazo_dias": prazo if prazo and 1 <= prazo <= 60 else None, "resumo": str(d.get("resumo") or "").strip()[:300]}


# ----------------------------------------------------------------------------- registro
@dataclass
class Registro:
    situacao: str = AGUARDANDO
    detectado: date | None = None
    prazo: date | None = None
    processo: str = ""
    reclamante: str = ""
    reclamado: str = ""
    parceiro: str = ""
    polo: str = ""
    resumo: str = ""
    link: str = ""
    observacoes: str = ""
    ids: str = ""
    origem: str = ""

    @property
    def norm(self) -> str:
        return re.sub(r"\D", "", self.processo or "")


def _data(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})", str(v or "").strip())
    return date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


def advogados_parceiros(ato, cfg: dict, cadastro: dict | None) -> str:
    """Advogados da comunicação que estão no Cadastro (nome ou OAB), fora os do escritório; sem Cadastro, todos."""
    do_escritorio = set(cfg.get("oabs_escritorio") or [])
    nomes_cad = {_asc(n).strip() for n in (cadastro or {}).get("nomes") or []}
    oabs_cad = {f"{str(o.get('oab')).lstrip('0')}/{str(o.get('uf')).upper()}" for o in (cadastro or {}).get("oabs") or []}
    todos, parceiros = [], []
    for a in getattr(ato, "advogados", None) or []:
        adv = a.get("advogado") or a
        nome = re.sub(r"\s+", " ", str(adv.get("nome") or "")).strip()
        oab = f"{str(adv.get('numero_oab') or '').lstrip('0')}/{str(adv.get('uf_oab') or '').upper()}"
        if not nome or oab in do_escritorio:
            continue
        rotulo = f"{nome.title()} (OAB {oab})" if oab != "/" else nome.title()
        if rotulo not in todos:
            todos.append(rotulo)
        if (oab in oabs_cad or _asc(nome) in nomes_cad) and rotulo not in parceiros:
            parceiros.append(rotulo)
    return "; ".join(parceiros or todos)[:200]


def detectar(atos_fora: list, cfg: dict, publicado: date, data_disp: date, usar_ia: bool, cadastro: dict | None = None) -> list[Registro]:
    """Atos do bloco 'Fora da Agenda' que mandam o intimado apresentar cálculos de liquidação."""
    from paralela.prazos import calcular_prazo, prazos_no_texto, tribunal_do_processo
    out: list[Registro] = []
    _polo = lambda a: " e ".join({"A": "AUTOR", "P": "RÉU"}.get(p, p) for p in sorted(a.polos)) or "?"  # noqa: E731
    candidatos = [a for a in atos_fora if candidato(a.texto)]
    respostas: list = [None] * len(candidatos)
    if usar_ia and candidatos:
        from concurrent.futures import ThreadPoolExecutor
        from paralela import ia
        with ThreadPoolExecutor(max_workers=ia.paralelismo(cfg)) as pool:      # módulo.função: os testes trocam classificar_ia
            respostas = list(pool.map(lambda a: classificar_ia(a.texto, _polo(a), cfg), candidatos))
    for ato, r in zip(candidatos, respostas):
        polo = _polo(ato)
        if r is None:
            if not ordem_direta(ato.texto):
                continue
            origem, n_ia = ORIGEM_REGRA, None
            m = RE_ORDEM.search(RE_RUIDO.sub(" ", _asc(ato.texto)))
            resumo = "Intimação para apresentar cálculos de liquidação (regra por palavras): " + m.group(0)[:200]
        elif not r["apresentar_calculos"]:
            continue
        else:
            origem, n_ia, resumo = ORIGEM_IA, r["prazo_dias"], r["resumo"] or "Intimação para apresentar cálculos de liquidação."
        n_textos = prazos_no_texto(ato.texto)
        # o prazo que vale é o da ordem de apresentar os cálculos: a IA o leu no trecho certo; sem IA, o menor do texto
        n = n_ia or (min(n_textos) if n_textos else None) or PRAZO_PADRAO
        pr = calcular_prazo(publicado, PECA, n, None, tribunal_do_processo(ato.norm), cfg)
        nome = lambda p: " / ".join(d.get("nome", "") for d in ato.destinatarios if d.get("polo") == p)[:80]  # noqa: E731
        out.append(Registro(situacao=AGUARDANDO, detectado=data_disp, prazo=pr.data, processo=ato.processo, reclamante=nome("A"),
                            reclamado=nome("P"), parceiro=advogados_parceiros(ato, cfg, cadastro), polo=polo, resumo=resumo,
                            link=ato.link, ids=";".join(ato.ids), origem=origem))
    return out


# ----------------------------------------------------------------------------- acumulado
def ler_aba(caminho: str) -> list[Registro] | None:
    """Registros da aba numa Paralela anterior; None se o arquivo não tem a aba (Paralela antiga ou criada pelo PC)."""
    copia = None
    try:                                   # no PC o arquivo pode estar aberto no Excel com AutoSave: só CopyFileW o lê
        from pc import arquivos
        copia = arquivos.copia_temporaria(caminho, prefixo="parceiros_")
    except Exception:  # noqa: BLE001 — na VM (Linux, sem o pacote pc) lê direto
        copia = None
    wb = openpyxl.load_workbook(copia or caminho, data_only=True)      # sem read_only: o link está no hyperlink da célula
    try:
        if ABA not in wb.sheetnames:
            return None
        ws = wb[ABA]
        cab = {str(ws.cell(PRIMEIRA_LINHA - 1, c).value or "").strip(): c for c in range(1, ws.max_column + 1)}
        col = {campo: cab.get(titulo) for campo, titulo in zip(CAMPOS, CABECALHOS)}
        if not col["processo"]:
            return None
        out = []
        for r in range(PRIMEIRA_LINHA, ws.max_row + 1):
            g = lambda campo: ws.cell(r, col[campo]).value if col[campo] else None  # noqa: E731
            if not g("processo"):
                continue
            cel_link = ws.cell(r, col["link"]) if col["link"] else None
            link = (cel_link.hyperlink.target if cel_link is not None and cel_link.hyperlink else None) or (g("link") if str(g("link") or "").startswith("http") else "")
            out.append(Registro(situacao=str(g("situacao") or AGUARDANDO), detectado=_data(g("detectado")), prazo=_data(g("prazo")),
                                processo=str(g("processo")).strip(), reclamante=str(g("reclamante") or ""), reclamado=str(g("reclamado") or ""),
                                parceiro=str(g("parceiro") or ""), polo=str(g("polo") or ""), resumo=str(g("resumo") or ""),
                                link=str(link or ""), observacoes=str(g("observacoes") or ""), ids=str(g("ids") or ""),
                                origem=str(g("origem") or "")))
        return out
    finally:
        wb.close()
        if copia:
            import shutil
            shutil.rmtree(os.path.dirname(copia), ignore_errors=True)


def _ordem(r: Registro):
    grupo = 0 if r.situacao == VENCIDO else (1 if r.situacao == AGUARDANDO else 2)
    return (grupo, r.prazo or date.max, r.processo)


def atualizar(anteriores: list[Registro], novos: list[Registro], na_agenda: set[str], hoje: date, dias_chegou: int = 30) -> list[Registro]:
    """Junta o acumulado com as intimações do dia e recalcula a situação. `na_agenda`: processos (20 dígitos) da Agenda.
    Sai da lista: linha com 'Baixar' nas Observações e linha que chegou à Agenda há mais de `dias_chegou` dias."""
    vistos: set[str] = set()
    out: list[Registro] = []
    for r in list(anteriores) + list(novos):
        ids = {i for i in re.split(r"[;,\s]+", r.ids or "") if i}
        if ids & vistos or _asc(r.observacoes).strip().startswith(BAIXAR):
            continue
        vistos |= ids
        if r.situacao.startswith(CHEGOU):
            visto = _data((re.search(r"\d{2}/\d{2}/\d{4}", r.situacao) or [""])[0])
            if visto and hoje - visto > timedelta(days=dias_chegou):
                continue
        elif r.norm and r.norm in na_agenda:
            r.situacao = f"{CHEGOU} (visto em {hoje:%d/%m/%Y})"
        elif r.prazo and r.prazo < hoje:
            r.situacao = VENCIDO
        else:
            r.situacao = AGUARDANDO
        out.append(r)
    return sorted(out, key=_ordem)


def achar_anterior(cfg: dict, data_disp: date) -> tuple[str | None, list[Registro]]:
    """A Paralela mais recente, anterior a `data_disp`, que tem a aba. Procura em `cfg["paralelas_anteriores"]` (a nuvem
    baixa as últimas para lá) e na pasta de saída (execução local)."""
    pastas = [cfg.get("paralelas_anteriores"), cfg.get("pasta_saida"),
              os.path.join(cfg.get("pasta_saida") or "", cfg.get("subpasta_paralela", "CACT - Agenda Paralela"))]
    achados = []
    for pasta in pastas:
        if not pasta or not os.path.isdir(pasta):
            continue
        for nome in os.listdir(pasta):
            m = re.fullmatch(r"CACT - Agenda Paralela (\d{2})-(\d{2})-(\d{4})\.xlsx", nome)
            if m:
                d = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                if d < data_disp:
                    achados.append((d, os.path.join(pasta, nome)))
    for _, caminho in sorted(achados, reverse=True)[:15]:
        try:
            regs = ler_aba(caminho)
        except Exception:  # noqa: BLE001 — arquivo ilegível: tenta o anterior
            continue
        if regs is not None:
            return caminho, regs
    return None, []


# ----------------------------------------------------------------------------- escrita
_THIN = Side(style="thin", color="BFBFBF")
_BORDA = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_COR = {VENCIDO: "F8CBAD", AGUARDANDO: "FFF2CC"}


def escrever_aba(wb, registros: list[Registro], data_disp: date, origem: str = "") -> None:
    ws = wb.create_sheet(ABA)
    ws["A1"] = "Parceiros - intimações para apresentar cálculos de liquidação em processos que não estão na Agenda"
    ws["A1"].font = Font(name="Arial", size=12, bold=True)
    ws["A2"] = (f"Lista acumulada até {data_disp:%d/%m/%Y}. Situação recalculada todo dia contra a Agenda: vermelho = prazo vencido sem o "
                "serviço chegar; amarelo = aguardando; verde = o processo passou a constar na Agenda.")
    ws["A2"].font = Font(name="Arial", size=9, italic=True, color="595959")
    ws["A4"] = ("A coluna K (Observações) é sua e segue para a Paralela do dia seguinte. Escreva 'Baixar' nela para tirar a linha da lista. "
                "Linha que chegou à Agenda sai sozinha depois de 30 dias.")
    ws["A4"].font = Font(name="Arial", size=9, color="595959")
    if origem:
        ws["A5"] = origem; ws["A5"].font = Font(name="Arial", size=9, color="595959")
    for i, (h, w) in enumerate(zip(CABECALHOS, LARGURAS), start=1):
        c = ws.cell(row=PRIMEIRA_LINHA - 1, column=i, value=h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="0070C0")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = _BORDA
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[PRIMEIRA_LINHA - 1].height = 30
    ws.freeze_panes = f"E{PRIMEIRA_LINHA}"
    esquerda = {5, 6, 7, 9, 11}
    for row, r in enumerate(registros, start=PRIMEIRA_LINHA):
        fill = PatternFill("solid", fgColor=_COR.get(r.situacao, "E2EFDA"))
        for i, campo in enumerate(CAMPOS, start=1):
            v = getattr(r, campo)
            if isinstance(v, str):
                v = limpar_texto(v)
            c = ws.cell(row=row, column=i, value=v if v != "" else None)
            c.font = Font(name="Courier New" if campo == "processo" else "Arial", size=10, bold=(campo == "situacao" and r.situacao == VENCIDO))
            c.border = _BORDA; c.fill = fill
            c.alignment = Alignment(horizontal="left" if i in esquerda else "center", vertical="center", wrap_text=True)
            if campo in ("detectado", "prazo"):
                c.number_format = "dd/mm/yy;@"
            if campo in ("processo", "ids"):
                c.number_format = "@"
            if campo == "link" and v:
                c.hyperlink = v; c.value = "abrir teor"
                c.font = Font(name="Arial", size=9, color="0563C1", underline="single")
        ws.cell(row=row, column=CAMPOS.index("observacoes") + 1).fill = PatternFill("solid", fgColor="FFFF00")
        ws.row_dimensions[row].height = 54
    n_v = sum(1 for r in registros if r.situacao == VENCIDO); n_a = sum(1 for r in registros if r.situacao == AGUARDANDO)
    ws.cell(row=PRIMEIRA_LINHA + len(registros) + 1, column=1,
            value=f"{len(registros)} intimação(ões) na lista: {n_v} com prazo vencido sem chegar, {n_a} aguardando, "
                  f"{len(registros) - n_v - n_a} que chegaram à Agenda.").font = Font(name="Arial", size=9, italic=True)


def montar(atos_fora: list, cfg: dict, na_agenda: set[str], publicado: date, data_disp: date, usar_ia: bool) -> tuple[list[Registro], str]:
    """Lista do dia: acumulado da Paralela anterior + detecções de hoje. Devolve (registros, nota de origem para a aba)."""
    cadastro = None
    try:
        if cfg.get("arquivo_cadastro") and os.path.exists(cfg["arquivo_cadastro"]):
            from cact_rotina import carregar_cadastro
            cadastro = carregar_cadastro(cfg)
    except Exception:  # noqa: BLE001 — sem Cadastro a coluna Parceiro lista todos os advogados da comunicação
        cadastro = None
    novos = detectar(atos_fora, cfg, publicado, data_disp, usar_ia, cadastro)
    anterior, antigos = achar_anterior(cfg, data_disp)
    nota = (f"Continuação de: {os.path.basename(anterior)}." if anterior
            else "Primeira lista (não achei Paralela anterior com esta aba): só as intimações de hoje.")
    return atualizar(antigos, novos, na_agenda, data_disp, int(cfg.get("parceiros_dias_chegou", 30))), nota

# -*- coding: utf-8 -*-
"""Lógica pura do transporte Paralela → Agenda: o que transportar, idempotência por ID DJEN e desfazer.

Sem COM aqui (testável sem Excel). Quem grava é ``pc.agenda_paralela``.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
from dataclasses import asdict, dataclass, field

import openpyxl

# colunas da Agenda oficial (spec, Global Constraints): N já existe (Polo conferido); O–S são novas
COL = {"polo_conferido": 14, "id_djen": 15, "link": 16, "origem": 17, "lote": 18, "conferir": 19}
# colunas da Paralela (paralela.planilha: A–K + controle). Desenho de 20/11/2026: L = Despacho anterior, link em O.
# As Paralelas anteriores têm o desenho antigo (COL_PARALELA_ANTIGA): NUNCA usar estas posições direto num arquivo;
# quem lê ou grava acha a coluna pelo cabeçalho da linha 7 (colunas_do_cabecalho / colunas_paralela / linha["cols"]).
CABECALHOS_PARALELA = {"Despacho anterior": "despacho_anterior", "Destino": "destino", "Decisão": "decisao", "Link (teor)": "link",
                       "Motivo CONFERIR": "motivo", "Polo intimado": "polo", "Intimado(s)": "intimados", "Fonte do polo": "fonte_polo",
                       "Prazo no texto": "prazo_texto", "Prazo legal": "prazo_legal", "Base legal": "base_legal", "ID DJEN": "ids",
                       "Lote": "lote", "Observações": "obs", "Prazo calculado": "prazo_calculado", "Corrigir J para": "corrigir_j"}
COL_PARALELA = {chave: 12 + i for i, chave in enumerate(CABECALHOS_PARALELA.values())}
COL_PARALELA_ANTIGA = {"destino": 12, "decisao": 13, "motivo": 14, "polo": 15, "intimados": 16, "fonte_polo": 17, "prazo_texto": 18,
                       "prazo_legal": 19, "base_legal": 20, "link": 21, "ids": 22, "lote": 23, "obs": 24, "prazo_calculado": 25,
                       "corrigir_j": 26}
LINHA_CABECALHO = 7
MARCA_APLICADA = "✔"
DECISOES_QUE_AGENDAM = {"agendar", "corrigir e agendar"}
ABAS_LIDAS = ("Lançamentos", "Pendentes")
PRIMEIRA_LINHA = 8


@dataclass
class Razao:
    quando: str
    paralela: str
    linhas: list[dict] = field(default_factory=list)   # {aba, linha_agenda, ids, valores, linha_paralela, aba_paralela}
    tipo: str = "transporte"

    def salvar(self, caminho: str) -> None:
        os.makedirs(os.path.dirname(os.path.abspath(caminho)), exist_ok=True)
        with open(caminho, "w", encoding="utf-8") as fh:
            json.dump(asdict(self), fh, ensure_ascii=False, indent=1, default=str)

    @classmethod
    def carregar(cls, caminho: str) -> "Razao":
        with open(caminho, encoding="utf-8") as fh:
            d = json.load(fh)
        return cls(**d)


def normalizar_j(valor) -> str | None:
    """'réu', 'Ré', 'RÉU' → 'RÉU'; 'autor', 'AUTOR(A)', 'autora' → 'AUTOR(A)'; outro → None."""
    import unicodedata
    v = unicodedata.normalize("NFKD", str(valor or "")).encode("ascii", "ignore").decode().strip().upper()
    if v.startswith("AUTOR"):
        return "AUTOR(A)"
    if v in ("RE", "REU", "RES", "REUS") or v.startswith("REU") or v.startswith("RE "):
        return "RÉU"
    return None


def _ids(valor) -> list[str]:
    """IDs de uma célula. O Excel devolve célula numérica como float (759139484.0): normaliza para "759139484"."""
    if valor in (None, ""):
        return []
    if isinstance(valor, float) and valor.is_integer():
        valor = str(int(valor))
    out = []
    for p in str(valor).replace(",", ";").split(";"):
        p = p.strip()
        if p.endswith(".0") and p[:-2].isdigit():
            p = p[:-2]
        if p:
            out.append(p)
    return out


def colunas_do_cabecalho(cabecalho) -> dict[str, int]:
    """chave → coluna (1 = A), pelos títulos da linha 7 da aba. Serve aos dois desenhos da Paralela (o de antes e o de
    depois de 20/11/2026). Linha sem título reconhecível (aba que não veio de `paralela.planilha`): desenho atual."""
    cols = {}
    for i, titulo in enumerate(cabecalho or (), start=1):
        chave = CABECALHOS_PARALELA.get(str(titulo or "").strip())
        if chave and chave not in cols:
            cols[chave] = i
    return cols if "ids" in cols and "decisao" in cols else dict(COL_PARALELA)


def colunas_paralela(caminho: str) -> dict[str, dict[str, int]]:
    """aba → colunas daquele arquivo (lido por cópia). Use antes de gravar numa Paralela: o arquivo pode ser do desenho antigo."""
    from pc import arquivos
    tmp = arquivos.copia_temporaria(caminho, prefixo="paralela_")
    wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
    try:
        return {ws.title: colunas_do_cabecalho(next(ws.iter_rows(min_row=LINHA_CABECALHO, max_row=LINHA_CABECALHO, values_only=True), ()))
                for ws in wb.worksheets}
    finally:
        wb.close()
        shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)


def ler_paralela(caminho: str) -> list[dict]:
    """Lê Lançamentos e Pendentes de uma CÓPIA do arquivo (o original pode estar aberto no Excel). As colunas de
    controle são achadas pelo cabeçalho; cada linha leva em "cols" o mapa de colunas da sua aba, para quem for gravar."""
    from pc import arquivos
    tmp = arquivos.copia_temporaria(caminho, prefixo="paralela_")   # CopyFileW: arquivo aberto no Excel com AutoSave
    wb = openpyxl.load_workbook(tmp, data_only=True)                # sem read_only: o endereço do link está no hyperlink da célula
    out = []
    try:
        for aba in ABAS_LIDAS:
            if aba not in wb.sheetnames:
                continue
            ws = wb[aba]
            cols = colunas_do_cabecalho([c.value for c in ws[LINHA_CABECALHO]])
            for i, celulas in enumerate(ws.iter_rows(min_row=PRIMEIRA_LINHA), start=PRIMEIRA_LINHA):
                row = [c.value for c in celulas]
                if not row or len(row) < 5 or not row[4]:
                    continue

                def g(chave, row=row, cols=cols):
                    c = cols.get(chave)
                    return row[c - 1] if c and len(row) >= c else None

                c_link = cols.get("link")
                cel_link = celulas[c_link - 1] if c_link and len(celulas) >= c_link else None
                link = (cel_link.hyperlink.target if cel_link is not None and cel_link.hyperlink else None) or g("link")
                cj = str(g("corrigir_j") or "")
                out.append({
                    "linha": i, "aba": aba, "cols": cols, "a_k": row[:11] + [None] * (11 - len(row[:11])),
                    "decisao": str(g("decisao") or "").strip(),
                    "motivo": str(g("motivo") or ""),
                    "link": str(link or ""),
                    "ids": _ids(g("ids")),
                    "lote": str(g("lote") or "").strip(),
                    "destino": str(g("destino") or ""),
                    "origem": str(g("obs") or "").strip(),
                    "prazo_calculado": g("prazo_calculado"),
                    "despacho_anterior": str(g("despacho_anterior") or ""),
                    "corrigir_j": None if cj.startswith(MARCA_APLICADA) else normalizar_j(cj),
                    "j_corrigida": cj.startswith(MARCA_APLICADA),
                })
    finally:
        wb.close()
        shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)
    return out


def ids_da_agenda(coluna: list) -> set[str]:
    out: set[str] = set()
    for v in coluna:
        out.update(_ids(v))
    return out


CNJ_RE = re.compile(r"\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}")


def selecionar(linhas: list[dict], ids_na_agenda: set[str]) -> tuple[list[dict], list[str]]:
    """Linhas com decisão de agendar, ainda não transportadas e sem id já presente na Agenda."""
    sel, avisos = [], []
    for l in linhas:
        if l.get("decisao", "").strip().lower() not in DECISOES_QUE_AGENDAM:
            continue
        if l.get("lote"):
            avisos.append(f"{l['aba']} linha {l['linha']}: já transportada ({l['lote']})")
            continue
        repetidos = [i for i in l.get("ids", []) if i in ids_na_agenda]
        if repetidos:
            avisos.append(f"{l['aba']} linha {l['linha']}: id {', '.join(repetidos)} já existe na Agenda")
            continue
        if not CNJ_RE.search(str((l.get("a_k") or [None] * 5)[4] or "")):
            avisos.append(f"{l['aba']} linha {l['linha']}: preencha o número do processo (coluna E) antes de agendar")
            continue
        sel.append(l)
    return sel, avisos


def valores_agenda(linha: dict, lote: str) -> dict[int, object]:
    """Mapa coluna→valor da linha nova da Agenda. B (fórmula) e N (Polo conferido) ficam fora: o gravador cuida de B."""
    v: dict[int, object] = {}
    for c, val in enumerate(linha["a_k"], start=1):
        if c == 2:
            continue
        v[c] = val
    v[COL["id_djen"]] = ";".join(linha.get("ids", []))
    v[COL["link"]] = linha.get("link", "")
    v[COL["origem"]] = "AUTO"
    v[COL["lote"]] = lote
    v[COL["conferir"]] = linha.get("motivo", "")
    return v


def linhas_para_desfazer(razao: Razao, ids_coluna: list) -> list[int]:
    """Linhas ATUAIS da Agenda cujo ID DJEN está no razão (busca por id; o usuário pode ter reordenado)."""
    alvo = {i for l in razao.linhas for i in l.get("ids", [])}
    out = []
    for idx, v in enumerate(ids_coluna):
        if set(_ids(v)) & alvo:
            out.append(PRIMEIRA_LINHA + idx)
    return out

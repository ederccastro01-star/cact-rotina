# -*- coding: utf-8 -*-
"""Acesso ao Outlook clássico via COM (pywin32) para a coleta do Push (spec §9c).

Sem senha nem token: o Outlook já guarda as credenciais das contas (Hotmail hoje; Gmail quando o advogado adicionar).
Cada fonte é {"conta": "<e-mail da conta>", "pastas": ["Caixa de Entrada", "Tribunais/PUSH", ...]}.
`listar` devolve `Email`s já com remetente, assunto, corpo (texto), data local e Message-ID da Internet (estável entre
pastas e contas, chave do livro-razão). `mover` arquiva e-mails processados em uma pasta da mesma conta.
"""
from __future__ import annotations

import re
from datetime import datetime, timedelta

from pc.push_fontes import Email

PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
ASSUNTOS = ("%PUSH%", "%Andamento Processual%", "%novas publica%", "%nova publica%", "%PROJUDI%")
REMETENTES = ("%jus.br%", "%jusbrasil%")
INBOX = 6
_RASTRO: dict[str, tuple[str, str]] = {}      # message_id -> (StoreID, EntryID), para `mover`


def _app():
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    return win32com.client.Dispatch("Outlook.Application")


def _ns():
    return _app().GetNamespace("MAPI")


def contas_disponiveis(ns=None) -> list[str]:
    ns = ns or _ns()
    return [str(st.DisplayName) for st in ns.Stores]


def _store(ns, conta: str):
    alvo = (conta or "").lower()
    for st in ns.Stores:
        if alvo and alvo in str(st.DisplayName).lower():
            return st
    return None


def _pasta(store, nome: str):
    if nome.lower() in ("caixa de entrada", "inbox", ""):
        return store.GetDefaultFolder(INBOX)
    atual = store.GetRootFolder()
    for parte in re.split(r"[\\/]", nome):
        prox = None
        for f in atual.Folders:
            if str(f.Name).lower() == parte.lower():
                prox = f
                break
        if prox is None:
            return None
        atual = prox
    return atual


def _criar_pasta(store, nome: str):
    atual = store.GetRootFolder()
    for parte in re.split(r"[\\/]", nome):
        prox = None
        for f in atual.Folders:
            if str(f.Name).lower() == parte.lower():
                prox = f
                break
        atual = prox if prox is not None else atual.Folders.Add(parte)
    return atual


def _dasl(desde: datetime, ate: datetime | None = None) -> str:
    q = '"urn:schemas:httpmail:'
    data = f"({q}datereceived\" >= '{desde:%m/%d/%Y %H:%M}')"
    if ate:
        data += f" AND ({q}datereceived\" < '{ate:%m/%d/%Y %H:%M}')"
    padroes = [f"{q}subject\" LIKE '{p}'" for p in ASSUNTOS] + [f"{q}fromemail\" LIKE '{p}'" for p in REMETENTES]
    return "@SQL=" + data + " AND (" + " OR ".join(padroes) + ")"


def _local(rt) -> datetime:
    try:
        return datetime.fromtimestamp(rt.timestamp())
    except (AttributeError, OSError, OverflowError, ValueError):
        return datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)


def _remetente(item) -> str:
    try:
        if str(item.SenderEmailType or "").upper() == "EX":
            return str(item.PropertyAccessor.GetProperty(PR_SENDER_SMTP) or "").lower()
    except Exception:  # noqa: BLE001
        pass
    return str(item.SenderEmailAddress or "").lower()


def _message_id(item) -> str:
    try:
        mid = str(item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID) or "").strip()
    except Exception:  # noqa: BLE001
        mid = ""
    return mid or f"<entryid:{item.EntryID}>"


def _itens(pasta, desde: datetime, ate: datetime | None = None):
    itens = pasta.Items.Restrict(_dasl(desde, ate))
    itens.Sort("[ReceivedTime]", True)
    return itens


def listar(fontes: list[dict], desde: datetime, ns=None) -> list[Email]:
    """Conta ausente no Outlook (ex.: Gmail ainda não adicionado) é avisada no log e pulada; erro só se nenhuma existir."""
    from pc.push_coletar import _log
    ns = ns or _ns()
    saida: list[Email] = []
    ausentes: list[str] = []
    for fonte in fontes:
        store = _store(ns, fonte.get("conta", ""))
        if store is None:
            ausentes.append(str(fonte.get("conta")))
            _log(f"AVISO: conta {fonte.get('conta')!r} não encontrada no Outlook (contas: {', '.join(contas_disponiveis(ns))}); pulada")
            continue
        for nome in fonte.get("pastas") or ["Caixa de Entrada"]:
            pasta = _pasta(store, nome)
            if pasta is None:
                raise RuntimeError(f"pasta {nome!r} não existe na conta {fonte.get('conta')!r}")
            for item in _itens(pasta, desde):
                if getattr(item, "Class", 43) != 43:          # olMail
                    continue
                mid = _message_id(item)
                _RASTRO[mid] = (str(store.StoreID), str(item.EntryID))
                saida.append(Email(remetente=_remetente(item), assunto=str(item.Subject or ""), corpo=str(item.Body or ""),
                                   data=_local(item.ReceivedTime), message_id=mid, conta=fonte.get("conta", ""), pasta=nome))
    if fontes and len(ausentes) == len(fontes):
        raise RuntimeError(f"nenhuma das contas {ausentes} está no Outlook (contas: {', '.join(contas_disponiveis(ns))})")
    return saida


def contar_backlog(fontes: list[dict], antes: datetime, ns=None, dias: int = 365) -> int:
    ns = ns or _ns()
    total = 0
    for fonte in fontes:
        store = _store(ns, fonte.get("conta", ""))
        if store is None:
            continue
        for nome in fonte.get("pastas") or ["Caixa de Entrada"]:
            pasta = _pasta(store, nome)
            if pasta is not None:
                total += int(pasta.Items.Restrict(_dasl(antes - timedelta(days=dias), antes)).Count)
    return total


def mover(emails: list[Email], destino: str, ns=None) -> int:
    """Move cada e-mail processado para `destino` (caminho de pastas na mesma conta; cria se faltar). Nunca exclui."""
    ns = ns or _ns()
    movidos = 0
    for e in emails:
        rastro = _RASTRO.get(e.message_id)
        if not rastro:
            continue
        store_id, entry_id = rastro
        item = ns.GetItemFromID(entry_id, store_id)
        store = item.Parent.Store
        alvo = _pasta(store, destino) or _criar_pasta(store, destino)
        if str(item.Parent.EntryID) != str(alvo.EntryID):
            item.Move(alvo)
            movidos += 1
    return movidos

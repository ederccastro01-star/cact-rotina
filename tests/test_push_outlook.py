# -*- coding: utf-8 -*-
"""Integração real com o Outlook clássico (só leitura). `pytest -m outlook tests/test_push_outlook.py`."""
from datetime import datetime, timedelta

import pytest

from pc import push_outlook as po
from pc.push_fontes import elegivel, extrair

pytestmark = pytest.mark.outlook


def test_lista_hotmail_e_extrai_registros():
    fontes = [{"conta": "ederccastro@hotmail.com", "pastas": ["Caixa de Entrada"]}]
    emails = po.listar(fontes, datetime.now() - timedelta(days=10))
    assert emails, "nenhum e-mail elegível nos últimos 10 dias"
    assert all(e.message_id.startswith("<") and e.data.tzinfo is None for e in emails)
    assert all(e.remetente and "@" in e.remetente for e in emails)
    eleg = [e for e in emails if elegivel(e)]
    regs = [r for e in eleg for r in extrair(e)]
    assert regs and all(r.processo for r in regs)


def test_conta_inexistente_da_erro_claro():
    with pytest.raises(RuntimeError, match="está no Outlook"):
        po.listar([{"conta": "ninguem@nada.com", "pastas": ["Caixa de Entrada"]}], datetime.now() - timedelta(days=1))


def test_contar_backlog_e_inteiro():
    n = po.contar_backlog([{"conta": "ederccastro@hotmail.com", "pastas": ["Caixa de Entrada"]}], datetime.now() - timedelta(days=7))
    assert isinstance(n, int) and n >= 0


def test_conta_ausente_e_pulada_quando_outra_existe():
    fontes = [{"conta": "ninguem@nada.com", "pastas": ["Caixa de Entrada"]}, {"conta": "ederccastro@hotmail.com", "pastas": ["Caixa de Entrada"]}]
    emails = po.listar(fontes, datetime.now() - timedelta(days=10))
    assert emails and {e.conta for e in emails} == {"ederccastro@hotmail.com"}

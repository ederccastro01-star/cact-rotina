# -*- coding: utf-8 -*-
import json

import openpyxl

from paralela import ia
from paralela.polo import agrupar_atos

COMPLETO = {
    "peca": "Contrarrazões", "prazo_dias": 8, "prazo_origem": "texto", "dias_uteis": True, "prazo_comum": False,
    "ente_publico_intimado": False, "tipo_ato": "decisão", "teor_presente": True, "favoravel": "neutro",
    "data_retorno": None, "resumo": "Intimado: X (Autor) · Ato: Decisão · Determinação: contraminuta · Prazo: 8 dias (texto)",
    "honorarios": False,
}


def test_parse_completo():
    assert ia.parse_resposta(json.dumps(COMPLETO)) == COMPLETO


def test_parse_incompleto():
    d = ia.parse_resposta('{"peca": "Contrarrazões"}')
    assert d["peca"] == "Contrarrazões"
    assert d["prazo_dias"] is None and d["prazo_origem"] == "nenhum"
    assert d["dias_uteis"] is True and d["prazo_comum"] is False
    assert d["tipo_ato"] == "notificação" and d["teor_presente"] is True
    assert d["favoravel"] == "neutro" and d["data_retorno"] is None and d["resumo"] == ""


def test_parse_lixo_em_volta():
    d = ia.parse_resposta('Segue:\n{"peca": "Manifestação", "prazo_dias": "5"}\nObrigado')
    assert d["peca"] == "Manifestação" and d["prazo_dias"] == 5


def test_parse_sem_json():
    assert ia.parse_resposta("não consigo") is None


def test_fallback_decisao(sintetico):
    ato = next(a for a in agrupar_atos(sintetico) if "9003" in a.ids)
    d = ia.fallback(ato.texto, ato)
    assert d["peca"] == "Embargos de Declaração"
    assert d["tipo_ato"] == "acórdão"
    assert d["prazo_origem"] == "nenhum"
    assert d["resumo"].startswith("Intimado: AUTOR TESTE")


def test_fallback_honorarios(sintetico):
    ato = next(a for a in agrupar_atos(sintetico) if "9001" in a.ids)
    d = ia.fallback(ato.texto, ato)
    assert d["peca"].startswith("Atenção! Honorários - ") and d["honorarios"] is True


def test_formatar_resumo_limite():
    d = dict(COMPLETO, resumo="x" * 1000)
    r = ia.formatar_resumo(d, ["FULANO DE TAL"], "A")
    assert len(r) <= 350 and r.startswith("Intimado: FULANO DE TAL (Autor)")


def test_formatar_resumo_quando_ia_ja_formatou():
    r = ia.formatar_resumo(COMPLETO, ["X"], "A")
    assert r == COMPLETO["resumo"]


def test_vocabulario_padrao_sem_aba(cfg, tmp_path):
    wb = openpyxl.Workbook(); wb.active.title = "Relação de advogados"; wb.save(tmp_path / "cad.xlsx")
    cfg["arquivo_cadastro"] = str(tmp_path / "cad.xlsx")
    voc = ia.carregar_vocabulario(cfg)
    assert len(voc) >= 10
    assert any(v["rotulo"] == "Atenção! Honorários" for v in voc)


def test_vocabulario_da_aba(cfg, tmp_path):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Vocabulário"
    ws.append(["Rótulo", "Significado", "Prazo (dias)", "Lado"])
    ws.append(["IC", "Impugnação aos cálculos", 8, "ambos"])
    ws.append(["CMAP", "Contraminuta ao agravo de petição", 8, "nosso"])
    wb.save(tmp_path / "cad.xlsx")
    cfg["arquivo_cadastro"] = str(tmp_path / "cad.xlsx")
    voc = ia.carregar_vocabulario(cfg)
    assert voc[0] == {"rotulo": "IC", "significado": "Impugnação aos cálculos", "prazo_dias": 8, "lado": "ambos"}
    assert len(voc) == 2


def test_classificar_sem_chave_retorna_none(cfg):
    cfg["anthropic_api_key"] = ""
    assert ia.classificar("texto", "AUTOR", "AUTOR", [], cfg) is None


def test_classificar_monta_corpo_opus(cfg, monkeypatch):
    capturado = {}

    class Resp:
        status_code = 200

        def json(self):
            return {"stop_reason": "end_turn", "content": [{"type": "text", "text": json.dumps(COMPLETO)}]}

    def post(url, headers=None, json=None, timeout=None):
        capturado["url"] = url; capturado["corpo"] = json; capturado["headers"] = headers
        return Resp()

    monkeypatch.setattr(ia, "_post", post)
    cfg["anthropic_api_key"] = "sk-teste"; cfg["modelo_claude"] = "claude-opus-5-5"
    d = ia.classificar("texto da intimação", "RÉU", "AUTOR", [{"rotulo": "IC", "significado": "x", "prazo_dias": 8, "lado": "ambos"}], cfg)
    corpo = capturado["corpo"]
    assert corpo["model"] == "claude-opus-5-5"
    assert "thinking" not in corpo
    assert corpo["output_config"] == {"effort": "medium"}
    assert corpo["max_tokens"] >= 4000          # revisão I5: o raciocínio do Opus 5.5 conta no limite
    assert corpo["system"] == ia.PROMPT_PARALELA
    assert "IC" in corpo["messages"][0]["content"] and "RÉU" in corpo["messages"][0]["content"]
    assert capturado["headers"]["x-api-key"] == "sk-teste"
    assert d["peca"] == "Contrarrazões"


def test_classificar_refusal_retorna_none(cfg, monkeypatch):
    class Resp:
        status_code = 200

        def json(self):
            return {"stop_reason": "refusal", "content": []}

    monkeypatch.setattr(ia, "_post", lambda *a, **k: Resp())
    cfg["anthropic_api_key"] = "sk-teste"
    assert ia.classificar("t", "AUTOR", "AUTOR", [], cfg) is None


def test_parse_sem_json_retorna_none():
    # falha de leitura NÃO pode virar "Manifestação" silenciosa: devolve None para o fallback por regras entrar
    assert ia.parse_resposta("não consigo") is None
    assert ia.parse_resposta("{isto não é json}") is None


def test_classificar_max_tokens_retorna_none(cfg, monkeypatch):
    class Resp:
        status_code = 200

        def json(self):
            return {"stop_reason": "max_tokens", "content": [{"type": "thinking", "thinking": ""}, {"type": "text", "text": '{"peca": "x"'}]}

    monkeypatch.setattr(ia, "_post", lambda *a, **k: Resp())
    cfg["anthropic_api_key"] = "sk-teste"
    assert ia.classificar("t", "AUTOR", "AUTOR", [], cfg) is None


def test_classificar_ignora_bloco_thinking_e_max_tokens_folgado(cfg, monkeypatch):
    capturado = {}

    class Resp:
        status_code = 200

        def json(self):
            return {"stop_reason": "end_turn", "content": [{"type": "thinking", "thinking": ""}, {"type": "text", "text": json.dumps(COMPLETO)}]}

    def post(url, headers=None, json=None, timeout=None):
        capturado["corpo"] = json; return Resp()

    monkeypatch.setattr(ia, "_post", post)
    cfg["anthropic_api_key"] = "sk-teste"
    d = ia.classificar("t", "AUTOR", "AUTOR", [], cfg)
    assert d["peca"] == "Contrarrazões" and capturado["corpo"]["max_tokens"] >= 4000

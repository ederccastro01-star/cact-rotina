# -*- coding: utf-8 -*-
"""Rotina no PC sem API paga (pc.rodar_local) e leitura pelo Claude Code (cfg["ia_via"] == "claude_code")."""
import json
import os
from datetime import date, datetime

import openpyxl

import cact_rotina as cr
from paralela import ia
from paralela import parceiros as pa
from pc import agenda_paralela as ap
from pc import rodar_local as rl


def _pc(cfg, tmp_path):
    base = tmp_path / "Controladoria"
    os.makedirs(base / "CACT - Agenda Paralela", exist_ok=True)
    cfg["pc"] = dict(cfg.get("pc") or {}, pasta_controladoria=str(base))
    for nome in (cfg.get("nome_agenda", "CACT - Agenda de Prazos.xlsx"), cfg.get("nome_cadastro", "CACT - Cadastro Publicações.xlsx")):
        openpyxl.Workbook().save(str(base / nome))
    return str(base)


def _falsa_rotina(chamadas):
    """Substitui executar_rotina: grava os três arquivos do dia na pasta de saída, como a rotina de verdade."""
    def rotina(c, data_disp=None, fonte="djen", arquivo=None, sem_ia=False):
        chamadas.append((data_disp, c["ia_via"], c["anthropic_api_key"], os.path.basename(c["arquivo_agenda"])))
        out = []
        for nome in (f"relatorio-detalhado-castro ({data_disp:%d.%m.%Y}).xlsx", f"CACT - Relatório Cruzamento Publicações {data_disp:%d-%m-%Y}.xlsx",
                     f"CACT - Agenda Paralela {data_disp:%d-%m-%Y}.xlsx"):
            caminho = os.path.join(c["pasta_saida"], nome)
            with open(caminho, "w", encoding="utf-8") as fh:
                fh.write("novo")
            out.append(caminho)
        return out
    return rotina


def test_dias_perdidos_do_mais_antigo_para_o_mais_novo(cfg, tmp_path):
    base = _pc(cfg, tmp_path)
    par = os.path.join(base, "CACT - Agenda Paralela")
    open(os.path.join(par, "CACT - Agenda Paralela 17-11-2026.xlsx"), "w").close()
    open(os.path.join(par, "CACT - Agenda Paralela 18-11-2026 (reprocessada 09h10).xlsx"), "w").close()
    # 20/11 é feriado (config), 21 e 22 fim de semana: faltam 16/11, 19/11 e 23/11
    assert rl.dias_a_gerar(cfg, date(2026, 11, 23), par) == [date(2026, 11, 16), date(2026, 11, 19), date(2026, 11, 23)]
    assert rl.dias_a_gerar(cfg, date(2026, 11, 21), par, dias=0) == []                 # sábado: nada


def test_rodar_grava_sem_sobrescrever_e_sem_api(cfg, tmp_path, monkeypatch):
    base = _pc(cfg, tmp_path)
    monkeypatch.setattr(rl, "LOG", str(tmp_path / "rodar.log")); monkeypatch.setattr(rl, "TRAVA", str(tmp_path / "rodar.lock"))
    cfg["anthropic_api_key"] = "sk-nao-usar"; cfg["pc"]["dias_recuperar"] = 1
    cruz = os.path.join(base, "CACT - Relatório Cruzamento Publicações 19-11-2026.xlsx")
    with open(cruz, "w", encoding="utf-8") as fh:
        fh.write("da VM")                                                              # já existia: fica como está
    chamadas = []
    monkeypatch.setattr(cr, "executar_rotina", _falsa_rotina(chamadas))
    feitos = rl.rodar(cfg, date(2026, 11, 19))
    assert [d for d, _ in feitos] == [date(2026, 11, 18), date(2026, 11, 19)]
    assert chamadas == [(date(2026, 11, 18), "claude_code", "", "agenda.xlsx"), (date(2026, 11, 19), "claude_code", "", "agenda.xlsx")]
    assert open(cruz, encoding="utf-8").read() == "da VM"
    assert os.path.exists(os.path.join(base, "CACT - Agenda Paralela", "CACT - Agenda Paralela 19-11-2026.xlsx"))
    assert os.path.exists(os.path.join(base, cfg["subpasta_historico"], "relatorio-detalhado-castro (19.11.2026).xlsx"))
    assert rl.rodar(cfg, date(2026, 11, 19)) == [] and len(chamadas) == 2              # segunda passada: tudo existe, nem consulta o DJEN
    assert cr.log.__name__ == "log"                                                     # o registro da rotina foi devolvido


def test_ensaio_em_pasta_separada_nao_toca_nas_pastas_reais(cfg, tmp_path, monkeypatch):
    base = _pc(cfg, tmp_path)
    monkeypatch.setattr(rl, "LOG", str(tmp_path / "rodar.log")); monkeypatch.setattr(rl, "TRAVA", str(tmp_path / "rodar.lock"))
    monkeypatch.setattr(cr, "executar_rotina", _falsa_rotina([]))
    saida = str(tmp_path / "ensaio")
    feitos = rl.rodar(cfg, date(2026, 11, 19), so_o_dia=True, saida=saida)
    assert feitos == [(date(2026, 11, 19), os.path.join(saida, "paralela", "CACT - Agenda Paralela 19-11-2026.xlsx"))]
    assert os.listdir(os.path.join(base, "CACT - Agenda Paralela")) == []


def test_vigia_gera_no_pc_so_com_a_chave_ligada(cfg, tmp_path, monkeypatch):
    _pc(cfg, tmp_path)
    cfg["pc"]["hora_alerta_sem_arquivo"] = "23:59"; cfg["pc"]["gerar_local"] = False      # o config real está ligado desde a virada
    chamadas = []
    monkeypatch.setattr(rl, "rodar", lambda c, dia: chamadas.append(dia) or [])
    monkeypatch.setattr(ap, "verificar_persistencia", lambda *a, **k: [])
    monkeypatch.setattr(ap, "caminho_push", lambda *a, **k: "ja-coletado")
    monkeypatch.setattr(ap, "conferir_push", lambda *a, **k: [])
    dia = date(2026, 11, 19)
    estado = ap.Estado()
    ap.passo_vigia(cfg, dia, datetime(2026, 11, 19, 8, 0), estado)
    assert chamadas == []                                                               # padrão: espera a VM
    cfg["pc"]["gerar_local"] = True
    ap.passo_vigia(cfg, dia, datetime(2026, 11, 19, 8, 1), estado)
    ap.passo_vigia(cfg, dia, datetime(2026, 11, 19, 8, 2), estado)                      # 1 min depois: ainda não tenta de novo
    ap.passo_vigia(cfg, dia, datetime(2026, 11, 19, 8, 17), estado)                     # 15 min depois: nova tentativa
    assert chamadas == [dia, dia]
    ap.passo_vigia(cfg, dia, datetime(2026, 11, 19, 9, 0), estado)                      # depois da 2ª: só de 2 em 2 horas
    assert len(chamadas) == 2
    for h in (10, 12, 14, 16, 18, 20, 22):
        ap.passo_vigia(cfg, dia, datetime(2026, 11, 19, h, 30), estado)
    assert len(chamadas) == 6                                                           # no máximo 6 por dia


# ----------------------------------------------------------------------------- leitura pelo Claude Code
class _Proc:
    def __init__(self, stdout, returncode=0):
        self.stdout, self.stderr, self.returncode = stdout, "", returncode


def test_chamar_claude_code_usa_arquivo_de_instrucoes_e_ambiente_sem_chave(monkeypatch):
    visto = {}

    def run(cmd, input=None, env=None, cwd=None, **kw):
        visto.update(cmd=cmd, input=input, env=env, sistema=open(cmd[cmd.index("--system-prompt-file") + 1], encoding="utf-8").read())
        return _Proc(json.dumps({"is_error": False, "result": '{"peca": "Manifestação"}'}))

    import subprocess
    monkeypatch.setattr(subprocess, "run", run)
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-x"); monkeypatch.setenv("CLAUDE_CODE_SESSION_ID", "abc")
    r = ia.chamar_claude_code("instruções\ncom duas linhas", "texto da intimação", {"claude_code_exe": "claude.cmd", "ia_via": "claude_code"})
    assert r == '{"peca": "Manifestação"}' and visto["sistema"] == "instruções\ncom duas linhas" and visto["input"] == "texto da intimação"
    assert "ANTHROPIC_API_KEY" not in visto["env"] and not any(k.upper().startswith("CLAUDE") for k in visto["env"])
    assert "--system-prompt" not in visto["cmd"] and visto["cmd"][visto["cmd"].index("--tools") + 1] == ""
    assert not os.path.exists(visto["cmd"][visto["cmd"].index("--system-prompt-file") + 1])     # pasta temporária apagada
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: _Proc(json.dumps({"is_error": True, "result": "limite do plano atingido"})))
    assert ia.chamar_claude_code("s", "e", {"claude_code_exe": "claude.cmd"}) is None    # erro → None → regra sem IA
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: _Proc("não é json", 1))
    assert ia.chamar_claude_code("s", "e", {"claude_code_exe": "claude.cmd"}) is None


def test_tres_leituras_vao_pelo_claude_code(monkeypatch):
    cfg = {"ia_via": "claude_code", "anthropic_api_key": ""}
    rotulos = []

    def falso(sistema, entrada, cfg_, rotulo="x"):
        rotulos.append(rotulo)
        if "Parceiros" in rotulo:
            return '{"apresentar_calculos": true, "quem": "autor", "prazo_dias": 8, "resumo": "autor intimado a liquidar"}'
        if "Cruzamento" in rotulo:
            return '```json\n{"destinatario": "FULANO (Autor)", "peca": "Manifestação", "prazo_dias": "5", "resumo": "r", "honorarios": false}\n```'
        return '{"peca": "Contrarrazões", "prazo_dias": 8, "resumo": "Intimado: X"}'

    monkeypatch.setattr(ia, "chamar_claude_code", falso)
    assert ia.tem_ia(cfg) and ia.paralelismo(cfg) == 4 and ia.paralelismo({"anthropic_api_key": "k"}) == 1
    assert ia.classificar("texto", "AUTOR", "AUTOR", [], cfg)["peca"] == "Contrarrazões"
    assert pa.classificar_ia("apresentar cálculos de liquidação", "AUTOR", cfg) == {
        "apresentar_calculos": True, "quem": "autor", "prazo_dias": 8, "resumo": "autor intimado a liquidar"}
    c = cr.classificar_claude("texto", cfg, "AUTOR(A)")
    assert c["peca"] == "Manifestação" and c["prazo_dias"] == 5 and c["destinatario"] == "FULANO (Autor)"
    assert rotulos == ["Claude Code (Paralela)", "Claude Code (Parceiros)", "Claude Code (Cruzamento)"]
    monkeypatch.setattr(ia, "chamar_claude_code", lambda *a, **k: None)
    assert ia.classificar("t", "AUTOR", "AUTOR", [], cfg) is None and cr.classificar_claude("t", cfg) is None


# ----------------------------------------------------------------------------- cadernos do DJEN (busca em zero)
def test_busca_em_zero_le_os_cadernos(monkeypatch):
    from pc import cadernos
    lidos = []
    monkeypatch.setattr(cadernos, "varrer_cadernos", lambda d, cad, extra: lidos.append(d) or [{"processo": "p", "texto": "t", "item": {}}])
    com_resultado = cadernos.busca_com_cadernos(lambda cfg, d, cad, extra: [{"processo": "x"}])
    assert com_resultado({}, date(2026, 11, 19), {}, []) == [{"processo": "x"}] and lidos == []      # busca normal: cadernos nem são lidos
    em_zero = cadernos.busca_com_cadernos(lambda cfg, d, cad, extra: [])
    assert len(em_zero({}, date(2026, 11, 19), {}, [])) == 1 and lidos == [date(2026, 11, 19)]


def test_caderno_filtra_pelo_cadastro(tmp_path):
    import zipfile
    from pc import cadernos
    itens = [{"id": 1, "numero_processo": "00000011120265180001", "texto": "do processo cadastrado", "destinatarioadvogados": []},
             {"id": 2, "numero_processo": "00000022220265180002", "texto": "da OAB cadastrada",
              "destinatarioadvogados": [{"advogado": {"nome": "Outro", "numero_oab": "0014725", "uf_oab": "go"}}]},
             {"id": 3, "numero_processo": "00000033320265180003", "texto": "do nome cadastrado",
              "destinatarioadvogados": [{"advogado": {"nome": "José  da Sílva", "numero_oab": "1", "uf_oab": "SP"}}]},
             {"id": 4, "numero_processo": "00000044420265180004", "texto": "de ninguém", "destinatarioadvogados": []},
             {"id": 2, "numero_processo": "00000022220265180002", "texto": "repetida", "destinatarioadvogados": [
                 {"advogado": {"nome": "Outro", "numero_oab": "14725", "uf_oab": "GO"}}]}]
    zp = str(tmp_path / "c.zip")
    with zipfile.ZipFile(zp, "w") as z:
        z.writestr("caderno.json", json.dumps({"items": itens}))
    oabs, nomes, procs = cadernos.filtros({"oabs": [{"uf": "GO", "oab": "14725"}], "nomes": ["JOSE DA SILVA"],
                                           "processos": ["0000001-11.2026.5.18.0001"]}, [])
    lidas, pubs = cadernos.ler_zip(zp, oabs, nomes, procs, set())
    assert lidas == 5 and [p["item"]["id"] for p in pubs] == [1, 2, 3]
    assert pubs[0]["item"]["numeroprocessocommascara"] == "0000001-11.2026.5.18.0001" and pubs[1]["texto"] == "da OAB cadastrada"


def test_cadernos_incompletos_nao_geram_nada_e_devolvem_a_busca(cfg, tmp_path, monkeypatch):
    from pc import cadernos
    _pc(cfg, tmp_path)
    monkeypatch.setattr(rl, "LOG", str(tmp_path / "rodar.log")); monkeypatch.setattr(rl, "TRAVA", str(tmp_path / "rodar.lock"))
    original = cr.buscar_djen

    def rotina(c, data_disp=None, fonte="djen", arquivo=None, sem_ia=False):
        assert cr.buscar_djen is not original                                           # durante a rotina a busca está embrulhada
        raise cadernos.CadernosIncompletos("TJSP/D: Em processamento")

    monkeypatch.setattr(cr, "executar_rotina", rotina)
    assert rl.rodar(cfg, date(2026, 11, 19), so_o_dia=True) == []
    assert cr.buscar_djen is original
    assert "cadernos ainda incompletos (TJSP/D: Em processamento)" in open(rl.LOG, encoding="utf-8").read()

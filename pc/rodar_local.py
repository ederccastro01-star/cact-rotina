# -*- coding: utf-8 -*-
"""Rotina de publicações rodando no PC, sem a API paga (20/11/2026).

Faz no PC o que `runner_nuvem.py` faz na VM: consulta o DJEN (o PC tem IP brasileiro), gera o relatório detalhado, o
Relatório de Cruzamento e a Agenda Paralela. A leitura de cada publicação (peça, resumo, aba de parceiros) é feita pelo
Claude Code sem janela, com o login da assinatura (`pc.ia_via` = "claude_code"; ver `paralela.ia.chamar_claude_code`).

Regras:
- Gera o dia pedido e, antes dele, os dias úteis dos últimos `pc.dias_recuperar` (7) dias que ficaram sem Paralela
  (computador desligado). Do mais antigo para o mais novo: a aba "Parceiros - Liquidação" é acumulada.
- NUNCA sobrescreve: dia que já tem Paralela é pulado; Cruzamento e detalhado que já existem ficam como estão.
- Agenda e Cadastro são lidos por CÓPIA (CopyFileW): ficam abertos no Excel.
- Um por vez (trava `pc/rodar_local.lock`). Registro em `pc/rodar_local.log`.

Uso:  python -m pc.rodar_local                      gera hoje e os dias perdidos, nas pastas reais
      python -m pc.rodar_local --data 19.11.2026    só esse dia
      python -m pc.rodar_local --saida C:\\teste    ensaio: grava tudo nessa pasta, não toca nas pastas reais
      python -m pc.rodar_local --sem-ia             regras, sem leitura pela IA
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile
from datetime import date, datetime, timedelta

AQUI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG = os.path.join(AQUI, "pc", "rodar_local.log")
TRAVA = os.path.join(AQUI, "pc", "rodar_local.lock")


def _log(msg: str) -> None:
    linha = msg if msg.startswith("[") else f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {msg}"
    try:
        print(linha, flush=True)
    except Exception:  # noqa: BLE001 — pythonw não tem saída
        pass
    try:
        with open(LOG, "a", encoding="utf-8") as fh:
            fh.write(linha + "\n")
    except OSError:
        pass


def _tem_paralela(pasta: str, dia: date) -> bool:
    """Já há Paralela desse dia (a normal ou uma reprocessada)?"""
    prefixo = f"CACT - Agenda Paralela {dia:%d-%m-%Y}"
    return os.path.isdir(pasta) and any(n.startswith(prefixo) and n.endswith(".xlsx") for n in os.listdir(pasta))


def dias_a_gerar(cfg: dict, ate: date, pasta_paralelas: str, dias: int | None = None) -> list[date]:
    """Dias úteis sem Paralela, de `ate - dias` até `ate`, do mais antigo para o mais novo."""
    import cact_rotina as cr
    feriados = cr.carregar_feriados(cfg)
    n = int((cfg.get("pc") or {}).get("dias_recuperar", 7)) if dias is None else dias
    out = []
    for k in range(n, -1, -1):
        d = ate - timedelta(days=k)
        if cr.eh_util(d, feriados) and not _tem_paralela(pasta_paralelas, d):
            out.append(d)
    return out


def _cfg_do_dia(cfg: dict, tmp: str, pasta_paralelas: str, sem_ia: bool) -> dict:
    from pc import arquivos
    pc = cfg.get("pc") or {}
    base = pc.get("pasta_controladoria", "")
    c = dict(cfg)
    c["arquivo_agenda"] = arquivos.copiar(os.path.join(base, cfg.get("nome_agenda", "CACT - Agenda de Prazos.xlsx")), os.path.join(tmp, "agenda.xlsx"))
    c["arquivo_cadastro"] = arquivos.copiar(os.path.join(base, cfg.get("nome_cadastro", "CACT - Cadastro Publicações.xlsx")), os.path.join(tmp, "cadastro.xlsx"))
    c["pasta_saida"] = c["pasta_historico"] = os.path.join(tmp, "out")
    os.makedirs(c["pasta_saida"], exist_ok=True)
    c["paralelas_anteriores"] = pasta_paralelas          # a lista de parceiros continua da Paralela anterior (lida por cópia)
    c["anthropic_api_key"] = ""                          # no PC a leitura nunca vai pela API paga
    c["ia_via"] = "" if sem_ia else str(pc.get("ia_via") or "claude_code")
    for chave in ("modelo_claude_code", "ia_paralelo", "claude_code_exe", "claude_code_timeout_seg"):
        if pc.get(chave) is not None:
            c[chave] = pc[chave]
    return c


def gerar_dia(cfg: dict, dia: date, destinos: dict, sem_ia: bool = False) -> str | None:
    """Gera os arquivos de um dia numa pasta temporária e os move para `destinos` sem sobrescrever nada.
    Devolve o caminho da Paralela gravada, ou None (dia sem publicação, falha no DJEN, Paralela já existente)."""
    import cact_rotina as cr
    tmp = tempfile.mkdtemp(prefix="cact_local_")
    try:
        c = _cfg_do_dia(cfg, tmp, destinos["paralela"], sem_ia)
        _log(f"gerando {dia:%d/%m/%Y} (leitura: {c['ia_via'] or 'regras, sem IA'})")
        try:
            gerados = cr.executar_rotina(c, data_disp=dia, fonte="djen", sem_ia=sem_ia)
        except SystemExit as e:              # executar_rotina encerra o processo quando falta a biblioteca requests
            raise RuntimeError(f"a rotina encerrou (código {e.code}); falta a biblioteca 'requests'?") from e
        paralela = None
        for caminho in gerados:
            nome = os.path.basename(caminho)
            if nome.lower().startswith("relatorio-detalhado"):
                pasta = destinos["historico"]
            elif nome.startswith("CACT - Agenda Paralela"):
                pasta = destinos["paralela"]
            else:
                pasta = destinos["base"]
            os.makedirs(pasta, exist_ok=True)
            destino = os.path.join(pasta, nome)
            if os.path.exists(destino) or (pasta == destinos["paralela"] and _tem_paralela(pasta, dia)):
                _log(f"  já existe, não sobrescrevo: {nome}")
                continue
            shutil.copyfile(caminho, destino)
            _log(f"  gravado: {destino}")
            if pasta == destinos["paralela"]:
                paralela = destino
        if not gerados:
            _log(f"  nada gerado para {dia:%d/%m/%Y} (sem publicações ou DJEN fora do ar)")
        return paralela
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def destinos_reais(cfg: dict) -> dict:
    base = (cfg.get("pc") or {}).get("pasta_controladoria", "")
    return {"base": base, "historico": os.path.join(base, cfg.get("subpasta_historico", "CACT - Histórioco de Publicações")),
            "paralela": os.path.join(base, cfg.get("subpasta_paralela", "CACT - Agenda Paralela"))}


def rodar(cfg: dict, ate: date, so_o_dia: bool = False, saida: str | None = None, sem_ia: bool = False) -> list[tuple[date, str]]:
    """Gera `ate` e os dias úteis perdidos antes dele. `saida`: ensaio numa pasta à parte. Devolve [(dia, Paralela)]."""
    import cact_rotina as cr
    from pc.agenda_paralela import obter_trava, soltar_trava
    trava = obter_trava(TRAVA)
    if trava is None:
        _log("outra geração em andamento; saindo")
        return []
    log_original = cr.log
    cr.log = lambda msg: _log(f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {msg}")      # o registro da rotina vai para o arquivo (pythonw)
    try:
        destinos = destinos_reais(cfg)
        if saida:
            destinos = {"base": saida, "historico": os.path.join(saida, "historico"), "paralela": os.path.join(saida, "paralela")}
        dias = [ate] if so_o_dia else dias_a_gerar(cfg, ate, destinos["paralela"])
        if so_o_dia and _tem_paralela(destinos["paralela"], ate):
            _log(f"{ate:%d/%m/%Y} já tem Paralela em {destinos['paralela']}; nada a fazer")
            return []
        feitos = []
        for d in dias:
            caminho = gerar_dia(cfg, d, destinos, sem_ia)
            if caminho:
                feitos.append((d, caminho))
        return feitos
    finally:
        cr.log = log_original
        soltar_trava(trava)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="CACT - rotina de publicações no PC (sem API paga)")
    ap.add_argument("--data", help="DD.MM.AAAA: gera só esse dia")
    ap.add_argument("--saida", help="ensaio: grava nessa pasta em vez das pastas reais")
    ap.add_argument("--sem-ia", action="store_true", help="classificação por regras, sem leitura pela IA")
    args = ap.parse_args(argv)
    from pc.agenda_paralela import carregar_cfg
    cfg = carregar_cfg()
    dia = datetime.strptime(args.data, "%d.%m.%Y").date() if args.data else date.today()
    feitos = rodar(cfg, dia, so_o_dia=bool(args.data), saida=args.saida, sem_ia=args.sem_ia)
    for d, c in feitos:
        print(f"{d:%d/%m/%Y}: {c}")
    if not feitos:
        print("nenhuma Paralela gerada (veja pc\\rodar_local.log)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

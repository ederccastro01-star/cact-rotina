# -*- coding: utf-8 -*-
"""Virada da rotina de publicações: servidor (VM, API paga) ⇄ PC (Claude Code da assinatura). 20/11/2026.

Os dois NUNCA ficam ligados juntos: o PC geraria a Paralela antes de o OneDrive trazer a da VM e haveria dois arquivos
com o mesmo nome. Por isso cada comando faz as duas pontas na ordem segura e confere antes de seguir.

  python -m pc.virada estado  --vm ubuntu@IP --chave C:\\Users\\...\\.ssh\\chave
  python -m pc.virada para-pc --vm ubuntu@IP --chave ...     desliga a VM (5h30 e contingência) e SÓ ENTÃO liga pc.gerar_local
  python -m pc.virada para-vm --vm ubuntu@IP --chave ...     desliga pc.gerar_local e SÓ ENTÃO religa a VM

Na VM só se movem os dois arquivos de agendamento (/etc/cron.d/cact e cact-contingencia) para /opt/cact/desligado/;
o keep-alive continua, para a Oracle não recolher a máquina, que fica de reserva. Nada é apagado.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys

AQUI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG = os.path.join(AQUI, "config_nuvem.json")
CRONS = ("cact", "cact-contingencia")
GUARDA = "/opt/cact/desligado"


def _ssh(vm: str, chave: str, comando: str) -> str:
    r = subprocess.run(["ssh", "-o", "BatchMode=yes", "-o", "ConnectTimeout=15", "-i", os.path.expanduser(chave), vm, comando],
                       capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=90)
    if r.returncode != 0:
        raise RuntimeError(f"ssh falhou (código {r.returncode}): {(r.stderr or r.stdout).strip()[:300]}")
    return r.stdout


def crons_ativos(vm: str, chave: str) -> list[str]:
    saida = _ssh(vm, chave, "ls /etc/cron.d/")
    return [c for c in CRONS if c in saida.split()]


def desligar_vm(vm: str, chave: str) -> None:
    _ssh(vm, chave, f"sudo mkdir -p {GUARDA} && for f in {' '.join(CRONS)}; do [ -f /etc/cron.d/$f ] && sudo mv /etc/cron.d/$f {GUARDA}/$f; done; true")
    ativos = crons_ativos(vm, chave)
    if ativos:
        raise RuntimeError(f"a VM continua com agendamento ativo: {ativos}")


def ligar_vm(vm: str, chave: str) -> None:
    _ssh(vm, chave, f"for f in {' '.join(CRONS)}; do [ -f {GUARDA}/$f ] && sudo mv {GUARDA}/$f /etc/cron.d/$f && sudo chown root:root /etc/cron.d/$f "
                    f"&& sudo chmod 644 /etc/cron.d/$f; done; true")
    if "cact" not in crons_ativos(vm, chave):
        raise RuntimeError("não religuei o agendamento das 5h30 na VM (/etc/cron.d/cact ausente)")


def gerar_local_ligado() -> bool:
    with open(CONFIG, encoding="utf-8") as fh:
        return bool((json.load(fh).get("pc") or {}).get("gerar_local", False))


def definir_gerar_local(ligado: bool) -> None:
    with open(CONFIG, encoding="utf-8") as fh:
        cfg = json.load(fh)
    pc = cfg.setdefault("pc", {})
    pc["gerar_local"] = bool(ligado)
    pc.setdefault("ia_via", "claude_code")
    pc.setdefault("dias_recuperar", 7)
    pc["_comentario_gerar_local"] = ("gerar_local true (pc.virada): a Paralela é gerada pelo PC ao ligar (pc.rodar_local), leitura pelo Claude Code "
                                     "da assinatura; os agendamentos da VM ficam em /opt/cact/desligado. NUNCA os dois ligados juntos.")
    with open(CONFIG, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(cfg, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


def estado(vm: str, chave: str) -> str:
    ativos = crons_ativos(vm, chave)
    local = gerar_local_ligado()
    linhas = [f"VM: {'LIGADA (' + ', '.join(ativos) + ')' if ativos else 'desligada (agendamentos guardados em ' + GUARDA + ')'}",
              f"PC: geração local {'LIGADA' if local else 'desligada'}"]
    if ativos and local:
        linhas.append("ATENÇÃO: os dois ligados ao mesmo tempo. Rode 'para-pc' ou 'para-vm'.")
    if not ativos and not local:
        linhas.append("ATENÇÃO: ninguém está gerando a Paralela. Rode 'para-pc' ou 'para-vm'.")
    return "\n".join(linhas)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="CACT - virada VM ⇄ PC")
    ap.add_argument("comando", choices=["estado", "para-pc", "para-vm"])
    ap.add_argument("--vm", required=True, help="usuário@IP da VM (ex.: ubuntu@1.2.3.4)")
    ap.add_argument("--chave", required=True, help="chave privada SSH")
    a = ap.parse_args(argv)
    if a.comando == "para-pc":
        desligar_vm(a.vm, a.chave)          # primeiro a VM; se falhar, o PC continua desligado e nada muda
        definir_gerar_local(True)
    elif a.comando == "para-vm":
        definir_gerar_local(False)          # primeiro o PC
        ligar_vm(a.vm, a.chave)
    print(estado(a.vm, a.chave))
    return 0


if __name__ == "__main__":
    sys.exit(main())

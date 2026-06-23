# -*- coding: utf-8 -*-
"""
Runner da nuvem (GitHub Actions).

Fluxo:
  1. Autentica no OneDrive (Microsoft Graph) com o refresh token.
  2. Baixa os insumos: 'CACT - Agenda de Prazos.xlsx' e 'CACT - Cadastro Publicações.xlsx'.
  3. Roda a mesma lógica do cact_rotina.py (busca DJEN + cruzamento + classificação IA).
  4. Envia de volta ao OneDrive o relatório de cruzamento e o relatório detalhado.

Segredos (variáveis de ambiente — configurados no GitHub):
  ANTHROPIC_API_KEY, GRAPH_CLIENT_ID, GRAPH_REFRESH_TOKEN
  (opcional) ONEDRIVE_BASE para sobrepor a pasta base do config_nuvem.json
"""

import os
import sys
import json
import tempfile
from datetime import datetime, timezone, timedelta

import cact_rotina as cr
import cact_onedrive as od

AQUI = os.path.dirname(os.path.abspath(__file__))


def carregar_config_nuvem():
    with open(os.path.join(AQUI, "config_nuvem.json"), "r", encoding="utf-8") as fh:
        return json.load(fh)


def main():
    base_cfg = carregar_config_nuvem()
    base = os.environ.get("ONEDRIVE_BASE", base_cfg["onedrive_base"]).strip("/")

    # data de hoje no fuso de Brasília (UTC-3)
    hoje = datetime.now(timezone(timedelta(hours=-3))).date()
    cr.log(f"Iniciando rotina na nuvem — {hoje:%d/%m/%Y} (Brasília).")

    tok = od.get_access_token()
    cr.log("Autenticado no OneDrive (Microsoft Graph).")

    tmp = tempfile.mkdtemp(prefix="cact_")
    out = os.path.join(tmp, "out")
    os.makedirs(out, exist_ok=True)

    # 1) baixa os insumos
    ag = f"{base}/{base_cfg['nome_agenda']}"
    cad = f"{base}/{base_cfg['nome_cadastro']}"
    with open(os.path.join(tmp, "agenda.xlsx"), "wb") as fh:
        fh.write(od.download(tok, ag))
    with open(os.path.join(tmp, "cadastro.xlsx"), "wb") as fh:
        fh.write(od.download(tok, cad))
    cr.log("Insumos baixados do OneDrive.")

    # 2) monta o cfg local apontando para os arquivos temporários
    cfg = dict(base_cfg)
    cfg["anthropic_api_key"] = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    cfg["arquivo_agenda"] = os.path.join(tmp, "agenda.xlsx")
    cfg["arquivo_cadastro"] = os.path.join(tmp, "cadastro.xlsx")
    cfg["pasta_saida"] = out
    cfg["pasta_historico"] = out

    # 3) roda a rotina
    gerados = cr.executar_rotina(cfg, data_disp=hoje, fonte="djen")

    if not gerados:
        cr.log("Nada gerado hoje (sem dia útil, sem publicações ou sem cruzamento).")
        return

    # 4) envia os arquivos gerados de volta ao OneDrive
    for caminho in gerados:
        nome = os.path.basename(caminho)
        if nome.lower().startswith("relatorio-detalhado"):
            destino = f"{base}/{base_cfg['subpasta_historico']}/{nome}"
        else:
            destino = f"{base}/{nome}"
        with open(caminho, "rb") as fh:
            od.upload(tok, destino, fh.read())
        cr.log(f"Enviado ao OneDrive: {destino}")

    cr.log("Rotina na nuvem concluída com sucesso.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        cr.log("ERRO NA NUVEM:\n" + traceback.format_exc())
        sys.exit(1)

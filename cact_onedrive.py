# -*- coding: utf-8 -*-
"""
Integração com o OneDrive via Microsoft Graph.
Usado pela rotina na nuvem (GitHub Actions) para baixar os insumos
(Agenda e Cadastro) e enviar os relatórios gerados de volta ao OneDrive.

Autenticação: fluxo de "refresh token" de um app público (sem segredo).
Variáveis de ambiente necessárias:
  GRAPH_CLIENT_ID       -> ID do aplicativo registrado no Azure (Entra ID)
  GRAPH_REFRESH_TOKEN   -> refresh token obtido uma única vez (obter_refresh_token.py)
"""

import os
import requests

GRAPH = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
SCOPE = "offline_access Files.ReadWrite User.Read"


# Onde guardar o refresh token mais recente (para ele NÃO expirar com o tempo).
# A cada renovação a Microsoft devolve um novo refresh token; salvamos e reusamos
# sempre o mais recente, mantendo o acesso vivo indefinidamente.
TOKEN_STORE = os.environ.get("REFRESH_TOKEN_FILE", "/opt/cact/refresh_token_atual.txt")


def _carregar_refresh_token():
    try:
        if os.path.exists(TOKEN_STORE):
            with open(TOKEN_STORE, "r", encoding="utf-8") as fh:
                t = fh.read().strip()
                if t:
                    return t
    except Exception:
        pass
    return os.environ.get("GRAPH_REFRESH_TOKEN", "").strip()


def _salvar_refresh_token(rt):
    try:
        with open(TOKEN_STORE, "w", encoding="utf-8") as fh:
            fh.write(rt)
    except Exception:
        pass


def get_access_token():
    cid = os.environ["GRAPH_CLIENT_ID"]
    rt = _carregar_refresh_token()
    if not rt:
        raise RuntimeError("Refresh token do OneDrive não encontrado (nem no arquivo nem no ambiente).")
    data = {
        "client_id": cid,
        "grant_type": "refresh_token",
        "refresh_token": rt,
        "scope": SCOPE,
    }
    r = requests.post(TOKEN_URL, data=data, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Falha ao renovar token ({r.status_code}): {r.text[:300]}")
    j = r.json()
    novo_rt = j.get("refresh_token")
    if novo_rt:
        _salvar_refresh_token(novo_rt)   # mantém o token sempre atual (evita a expiração)
    return j["access_token"]


def _h(tok):
    return {"Authorization": f"Bearer {tok}"}


def _enc(rel_path):
    # Graph aceita o caminho relativo à raiz do drive; espaços e acentos são
    # tratados pelo requests, mas garantimos a ausência de barras iniciais.
    return rel_path.lstrip("/")


def download(tok, rel_path):
    """Baixa um arquivo do OneDrive (caminho relativo à raiz). Retorna bytes."""
    url = f"{GRAPH}/me/drive/root:/{_enc(rel_path)}:/content"
    r = requests.get(url, headers=_h(tok), timeout=180)
    if r.status_code != 200:
        raise RuntimeError(f"Falha ao baixar '{rel_path}' ({r.status_code}): {r.text[:200]}")
    return r.content


def upload(tok, rel_path, data):
    """Envia (cria/sobrescreve) um arquivo no OneDrive. Usa upload simples até
    4 MB e sessão de upload para arquivos maiores."""
    rel_path = _enc(rel_path)
    if len(data) <= 4 * 1024 * 1024:
        url = f"{GRAPH}/me/drive/root:/{rel_path}:/content"
        r = requests.put(
            url,
            headers={**_h(tok), "Content-Type": "application/octet-stream"},
            data=data, timeout=300,
        )
        if r.status_code not in (200, 201):
            raise RuntimeError(f"Falha ao enviar '{rel_path}' ({r.status_code}): {r.text[:200]}")
        return r.json()
    # arquivos grandes -> sessão de upload (fragmentos de ~3,2 MB)
    url = f"{GRAPH}/me/drive/root:/{rel_path}:/createUploadSession"
    r = requests.post(url, headers=_h(tok), json={"item": {"@microsoft.graph.conflictBehavior": "replace"}}, timeout=60)
    r.raise_for_status()
    upload_url = r.json()["uploadUrl"]
    chunk = 3276800
    total = len(data)
    for ini in range(0, total, chunk):
        fim = min(ini + chunk, total)
        headers = {
            "Content-Length": str(fim - ini),
            "Content-Range": f"bytes {ini}-{fim-1}/{total}",
        }
        rr = requests.put(upload_url, headers=headers, data=data[ini:fim], timeout=300)
        if rr.status_code not in (200, 201, 202):
            raise RuntimeError(f"Falha no envio fragmentado ({rr.status_code}): {rr.text[:200]}")
    return {"ok": True}

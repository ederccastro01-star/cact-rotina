# -*- coding: utf-8 -*-
"""Leitura dos CADERNOS do DJEN, para quando a busca (/api/v1/comunicacao) devolve zero.

Trazido em 20/11/2026 de `/opt/cact-fallback/contingencia.py` (a contingência da VM, que ficava fora do repositório),
quando a rotina passou a rodar no PC. Mesma regra: lê os cadernos (meios D e E) de todos os tribunais e filtra pelo
Cadastro (OAB, nome, processo). Não devolve nada pela metade: se algum caderno ainda não estiver fechado, ou se a
leitura não bater com o total informado pelo DJEN, levanta `CadernosIncompletos` e quem chamou tenta mais tarde.
Leitura item a item (ijson): o caderno de um tribunal grande não cabe folgado na memória.

Uso (só varre e conta; não grava nada):  python -m pc.cadernos AAAA-MM-DD
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import tempfile
import time
import unicodedata
import urllib.error
import urllib.request
import zipfile
from datetime import date, datetime

import cact_rotina as cr

API = "https://comunicaapi.pje.jus.br/api/v1/caderno"
HDR = {"User-Agent": "Mozilla/5.0", "Accept": "application/json", "Referer": "https://comunica.pje.jus.br/"}
UFS = "AC AL AP AM BA CE DF ES GO MA MT MS MG PA PB PR PE PI RJ RN RS RO RR SC SP SE TO".split()
SIGLAS = (["STF", "STJ", "TST", "TSE", "STM", "CNJ", "CJF", "CSJT"]
          + [f"TRF{i}" for i in range(1, 7)] + [f"TRT{i}" for i in range(1, 25)]
          + ["TJ" + u for u in UFS if u != "DF"] + ["TJDFT"]
          + [f"TRE-{u}" for u in UFS] + ["TJMSP", "TJMMG", "TJMRS"])
FECHADOS = ("Processado", "Sem comunicações")
PAUSA = 3            # segundos entre consultas de situação (o DJEN devolve 429 com menos)


class CadernosIncompletos(Exception):
    pass


def norm(s) -> str:
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.upper()).strip()


def mascara(n) -> str:
    d = re.sub(r"\D", "", n or "")
    return f"{d[:7]}-{d[7:9]}.{d[9:13]}.{d[13]}.{d[14:16]}.{d[16:]}" if len(d) == 20 else (n or "")


def get_json(url: str, tentativas: int = 8) -> dict:
    ultimo = ""
    for tent in range(tentativas):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=60) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            corpo = e.read()
            ultimo = f"HTTP {e.code}"
            if e.code == 429 or e.code >= 500:
                time.sleep(20 * (tent + 1))
                continue
            try:
                return json.loads(corpo)
            except Exception:  # noqa: BLE001
                return {"erro": ultimo}
        except Exception as e:  # noqa: BLE001
            ultimo = f"{type(e).__name__}: {e}"
            time.sleep(10 * (tent + 1))
    return {"erro": f"falha após {tentativas} tentativas ({ultimo})"}


def baixar(url: str, destino: str) -> bool:
    for _ in range(4):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=HDR), timeout=900) as r, open(destino, "wb") as f:
                shutil.copyfileobj(r, f, 1 << 20)
            return True
        except Exception:  # noqa: BLE001
            time.sleep(15)
    return False


def do_escritorio(it: dict, oabs: set, nomes: set, procs: set) -> bool:
    """A comunicação é de processo do Cadastro, ou tem advogado do Cadastro (OAB ou nome)?"""
    if re.sub(r"\D", "", it.get("numero_processo") or "") in procs:
        return True
    for a in it.get("destinatarioadvogados") or []:
        ad = a.get("advogado") or {}
        num = re.sub(r"\D", "", str(ad.get("numero_oab") or "")) or "0"
        if ((ad.get("uf_oab") or "").upper(), str(int(num))) in oabs or norm(ad.get("nome")) in nomes:
            return True
    return False


def filtros(cadastro: dict, processos_extra) -> tuple[set, set, set]:
    oabs = {(o["uf"].upper(), str(int(o["oab"]))) for o in cadastro.get("oabs", [])}
    nomes = {norm(n) for n in cadastro.get("nomes", [])}
    procs = {re.sub(r"\D", "", p) for p in list(cadastro.get("processos", [])) + list(processos_extra or [])}
    return oabs, nomes, procs


def ler_zip(zp: str, oabs: set, nomes: set, procs: set, vistos: set) -> tuple[int, list[dict]]:
    """(comunicações lidas, publicações do escritório) de um caderno baixado."""
    import ijson
    lidas, pubs = 0, []
    with zipfile.ZipFile(zp) as z:
        for inf in z.infolist():
            with z.open(inf) as fh:
                for it in ijson.items(fh, "items.item", use_float=True):
                    lidas += 1
                    if not do_escritorio(it, oabs, nomes, procs) or it.get("id") in vistos:
                        continue
                    vistos.add(it.get("id"))
                    it.setdefault("numeroprocessocommascara", mascara(it.get("numero_processo")))
                    texto = cr._texto_item(it)
                    if texto:
                        pubs.append({"processo": cr._processo_item(it), "texto": texto, "item": it})
    return lidas, pubs


def varrer_cadernos(data_disp: date, cadastro: dict, processos_extra, siglas=None) -> list[dict]:
    """Publicações do escritório no dia, no mesmo formato de `cr.buscar_djen`. Levanta CadernosIncompletos."""
    data = data_disp.strftime("%Y-%m-%d")
    siglas = list(siglas or SIGLAS)
    oabs, nomes, procs = filtros(cadastro, processos_extra)

    # 0) o DJEN responde? (PC sem internet ao ligar: sair logo, em vez de esperar 8 tentativas por caderno)
    primeiro = get_json(f"{API}/TRT18/{data}/D", tentativas=2)
    if "erro" in primeiro:
        raise CadernosIncompletos(f"DJEN inacessível ({primeiro['erro']})")

    # 1) todos os cadernos precisam estar fechados antes de qualquer download
    metas = {}
    for sig in siglas:
        for meio in ("D", "E"):
            m = get_json(f"{API}/{sig}/{data}/{meio}")
            time.sleep(PAUSA)
            if "erro" in m or m.get("status") not in FECHADOS:
                raise CadernosIncompletos(f"{sig}/{meio}: {m.get('erro') or m.get('status')}")
            metas[(sig, meio)] = m
    if ("TRT18", "D") in metas and metas[("TRT18", "D")].get("status") != "Processado":
        raise CadernosIncompletos("caderno D do TRT18 ainda sem comunicações")
    cr.log(f"CADERNOS: {len(metas)} cadernos fechados, {sum(int(m.get('total_comunicacoes') or 0) for m in metas.values())} comunicações a ler.")

    # 2) leitura item a item
    pubs, vistos, por_trib, lidas_total = [], set(), {}, 0
    tmp = tempfile.mkdtemp(prefix="cact_cadernos_")
    try:
        for (sig, meio), m in metas.items():
            total = int(m.get("total_comunicacoes") or 0)
            if m["status"] != "Processado" or total == 0:
                continue
            url = get_json(f"{API}/{sig}/{data}/{meio}").get("url", "")   # a url expira em 5 min
            zp = os.path.join(tmp, f"{sig}_{meio}.zip")
            if not url or not baixar(url, zp):
                raise CadernosIncompletos(f"{sig}/{meio}: download falhou")
            lidas, achadas = ler_zip(zp, oabs, nomes, procs, vistos)
            os.remove(zp)
            if lidas != total:
                raise CadernosIncompletos(f"{sig}/{meio}: lidas {lidas} de {total}")
            lidas_total += lidas
            pubs += achadas
            if achadas:
                por_trib[sig] = por_trib.get(sig, 0) + len(achadas)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    cr.log(f"CADERNOS: {len(metas)} cadernos conferidos, {lidas_total} comunicações lidas. "
           f"Do escritório, por tribunal: {dict(sorted(por_trib.items()))}")
    return pubs


def busca_com_cadernos(busca_original):
    """Embrulha `cr.buscar_djen`: se a busca devolver zero, lê os cadernos. CadernosIncompletos sobe para quem chamou."""
    def buscar(cfg, data_disp, cadastro, processos_extra):
        pubs = busca_original(cfg, data_disp, cadastro, processos_extra)
        if pubs:
            return pubs
        cr.log(f"CADERNOS: a busca do DJEN devolveu zero; lendo os cadernos ({len(SIGLAS)} tribunais, meios D e E).")
        pubs = varrer_cadernos(data_disp, cadastro, processos_extra)
        cr.log(f"Publicações capturadas (após dedupe): {len(pubs)}")
        return pubs
    return buscar


def main(argv=None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) < 1:
        print(__doc__)
        return 2
    from pc import arquivos
    from pc.agenda_paralela import carregar_cfg
    cfg = carregar_cfg()
    d = datetime.strptime(argv[0], "%Y-%m-%d").date()
    base = (cfg.get("pc") or {}).get("pasta_controladoria", "")
    tmp = arquivos.copia_temporaria(os.path.join(base, cfg.get("nome_cadastro", "CACT - Cadastro Publicações.xlsx")), prefixo="cad_")
    try:
        cad = cr.carregar_cadastro(dict(cfg, arquivo_cadastro=tmp))
    finally:
        arquivos.apagar_copia(tmp)
    t = time.time()
    pubs = varrer_cadernos(d, cad, cfg.get("processos_monitorados") or [], siglas=argv[1:] or None)
    print(f"TOTAL {len(pubs)} publicações em {time.time() - t:.0f}s")
    saida = os.environ.get("CACT_CADERNOS_IDS")          # conferência: grava os ids achados para comparar com a busca
    if saida:
        with open(saida, "w", encoding="utf-8") as fh:
            json.dump(sorted(str(p["item"].get("id")) for p in pubs), fh)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except CadernosIncompletos as e:
        print(f"cadernos incompletos: {e}")
        sys.exit(1)

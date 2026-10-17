# -*- coding: utf-8 -*-
"""Cópia de arquivos que o Excel mantém trancados.

Fato medido em 17/10/2026: com o AutoSave do OneDrive ligado, o Excel abre `CACT - Agenda de Prazos.xlsx` de modo
que `open()`/`shutil.copyfile` do Python (e `File.Open` do .NET) falham com "Permission denied" em qualquer modo de
compartilhamento, enquanto `Copy-Item` e a API Win32 `CopyFileW` copiam normalmente. Toda leitura da Agenda (e da
Paralela, que também fica aberta no Excel do advogado) passa por aqui: copia-se com `CopyFileW` e lê-se a cópia.
"""
from __future__ import annotations

import ctypes
import os
import shutil
import sys
import tempfile
import time


def copiar(origem: str, destino: str, tentativas: int = 6, espera: float = 2.0) -> str:
    """Copia `origem` para `destino` (substitui). No Windows usa CopyFileW; noutros sistemas, shutil."""
    ultimo: Exception | None = None
    for _ in range(max(1, tentativas)):
        try:
            if sys.platform == "win32":
                ok = ctypes.windll.kernel32.CopyFileW(ctypes.c_wchar_p(origem), ctypes.c_wchar_p(destino), False)
                if not ok:
                    raise OSError(ctypes.GetLastError(), f"CopyFileW falhou (erro {ctypes.GetLastError()}) ao copiar {origem}")
            else:
                shutil.copyfile(origem, destino)
            return destino
        except OSError as e:
            ultimo = e
            time.sleep(espera)
    raise ultimo  # type: ignore[misc]


def copia_temporaria(caminho: str, prefixo: str = "cact_") -> str:
    """Cópia em pasta temporária própria, com o mesmo nome de arquivo; apagar com `apagar_copia`."""
    tmp = os.path.join(tempfile.mkdtemp(prefix=prefixo), os.path.basename(caminho))
    return copiar(caminho, tmp)


def apagar_copia(tmp: str) -> None:
    shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)

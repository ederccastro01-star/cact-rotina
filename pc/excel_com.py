# -*- coding: utf-8 -*-
"""Camada fina sobre o Excel via COM (pywin32), resistente ao Excel ocupado.

O Excel recusa automação (RPC_E_CALL_REJECTED) enquanto uma célula está em edição ou um diálogo
está aberto. Toda chamada passa por :func:`com`, que espera e repete. A Agenda é gravada SÓ por
aqui, nunca por openpyxl (26 abas, formatação e histórico que o openpyxl perderia).

Nunca devolvemos um livro aberto só-leitura (trava do OneDrive/Excel que ainda não soltou o
arquivo): ``Save``/``Close(True)`` nele abrem um diálogo e a automação fica presa para sempre.
"""
from __future__ import annotations

import os
import time

REJEITADAS = {-2147418111, -2147417848, -2147023170}
TENTATIVAS_SO_LEITURA = 6
_avisou = {"ocupado": False}


def _hresult(e: Exception):
    hr = getattr(e, "hresult", None)
    if hr is None and getattr(e, "args", None) and isinstance(e.args[0], int):
        hr = e.args[0]
    return hr


def com(fn, *a, espera_max: int = 600, **k):
    """Executa ``fn(*a, **k)`` repetindo a cada 2 s enquanto o Excel estiver ocupado (até ``espera_max`` s)."""
    inicio = time.monotonic()
    while True:
        try:
            return fn(*a, **k)
        except Exception as e:  # noqa: BLE001
            if _hresult(e) not in REJEITADAS or time.monotonic() - inicio > espera_max:
                raise
            if not _avisou["ocupado"]:
                nome = getattr(fn, "__name__", None) or getattr(getattr(fn, "_oleobj_", None), "__class__", type(fn)).__name__
                print(f"  Excel ocupado em {nome} (célula em edição ou diálogo aberto); aguardando...", flush=True)
                _avisou["ocupado"] = True
            time.sleep(2)


def set_cell(ws, r: int, c: int, v) -> None:
    ws.Cells(r, c).Value = v


def set_text(ws, r: int, c: int, v) -> None:
    """Grava como TEXTO (formato @): IDs e números de processo não podem virar número no Excel."""
    cel = ws.Cells(r, c)
    cel.NumberFormat = "@"
    cel.Value = str(v)


def _bloqueado(caminho: str) -> bool:
    """True se outro processo (Excel que acabou de fechar, OneDrive) ainda segura o arquivo."""
    if not os.path.exists(caminho):
        return False                      # open(..., "a") criaria o arquivo
    try:
        with open(caminho, "a"):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


def get_range(ws, r1: int, c1: int, r2: int, c2: int) -> list[list]:
    v = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Value2
    if not isinstance(v, (list, tuple)):
        return [[v]]
    return [list(linha) if isinstance(linha, (list, tuple)) else [linha] for linha in v]


class _GUID(__import__("ctypes").Structure):
    import ctypes as _c
    _fields_ = [("Data1", _c.c_ulong), ("Data2", _c.c_ushort), ("Data3", _c.c_ushort), ("Data4", _c.c_ubyte * 8)]


def _get_active_object():
    """A instância VISÍVEL do Excel (a do usuário), pela janela XLMAIN → EXCEL7 e a API de acessibilidade.

    Não usa GetActiveObject/ROT de propósito: a ROT pode apontar para uma instância oculta que está
    morrendo, e a chamada RPC seguinte fica presa. Lança RuntimeError se não há Excel visível."""
    import ctypes
    import pythoncom
    import win32com.client
    import win32gui
    pythoncom.CoInitialize()
    iid = _GUID(0x00020400, 0, 0, (ctypes.c_ubyte * 8)(0xC0, 0, 0, 0, 0, 0, 0, 0x46))  # IID_IDispatch
    hwnd = win32gui.FindWindow("XLMAIN", None)
    while hwnd:
        if win32gui.IsWindowVisible(hwnd):
            desk = win32gui.FindWindowEx(hwnd, 0, "XLDESK", None)
            child = win32gui.FindWindowEx(desk, 0, "EXCEL7", None) if desk else 0
            if child:
                ptr = ctypes.c_void_p()
                ctypes.oledll.oleacc.AccessibleObjectFromWindow(child, 0xFFFFFFF0, ctypes.byref(iid), ctypes.byref(ptr))
                disp = pythoncom.ObjectFromAddress(ptr.value, pythoncom.IID_IDispatch)
                return win32com.client.Dispatch(disp).Application
        hwnd = win32gui.FindWindowEx(0, hwnd, "XLMAIN", None)
    raise RuntimeError("nenhuma instância visível do Excel")


def _livro_na_instancia_ativa(caminho: str):
    """(xl, wb) se o livro já está aberto na instância VISÍVEL do usuário; senão None.
    Uma instância sem janela é automação (nossa ou alheia, talvez morrendo): nunca é a do usuário."""
    nome = os.path.basename(caminho).lower()
    try:
        xl = _get_active_object()
        if not com(lambda: xl.Visible):
            return None
        for w in com(lambda: list(xl.Workbooks)):
            if w.Name.lower() == nome:
                return xl, w
    except Exception:  # noqa: BLE001 — sem instância ativa
        return None
    return None


def _pid_de(xl) -> int | None:
    try:
        import ctypes
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(int(xl.Hwnd), ctypes.byref(pid))
        return pid.value or None
    except Exception:  # noqa: BLE001
        return None


def _k32():
    import ctypes
    k32 = ctypes.windll.kernel32
    k32.OpenProcess.restype = ctypes.c_void_p
    k32.OpenProcess.argtypes = [ctypes.c_uint32, ctypes.c_int, ctypes.c_uint32]
    k32.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    k32.CloseHandle.argtypes = [ctypes.c_void_p]
    k32.TerminateProcess.argtypes = [ctypes.c_void_p, ctypes.c_uint32]
    return k32


def _processo_vivo(pid: int) -> bool:
    k32 = _k32()
    h = k32.OpenProcess(0x100000, False, pid)  # SYNCHRONIZE
    if not h:
        return False
    try:
        return k32.WaitForSingleObject(h, 0) != 0
    finally:
        k32.CloseHandle(h)


def _encerrar_processo(pid: int) -> None:
    k32 = _k32()
    h = k32.OpenProcess(0x0001, False, pid)  # PROCESS_TERMINATE
    if h:
        try:
            k32.TerminateProcess(h, 0)
        finally:
            k32.CloseHandle(h)


def _nova_instancia(visivel: bool):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = visivel
    xl.DisplayAlerts = False
    return xl


def abrir_livro(caminho: str, visivel: bool = False):
    """Conecta ao livro já aberto na instância ativa do Excel; senão abre numa instância própria, escondida.
    Devolve (xl, wb, aberto_por_mim). Quem abriu fecha (``fechar_livro``). Nunca devolve livro só-leitura."""
    ativo = _livro_na_instancia_ativa(caminho)
    if ativo:
        return ativo[0], ativo[1], False
    for _ in range(30):                       # espera o arquivo ser solto (Quit do Excel é assíncrono; OneDrive)
        if not _bloqueado(caminho):
            break
        time.sleep(1)
    xl = _nova_instancia(visivel)
    for tentativa in range(TENTATIVAS_SO_LEITURA):
        wb = com(xl.Workbooks.Open, os.path.abspath(caminho), UpdateLinks=0, ReadOnly=False)
        if not com(lambda: wb.ReadOnly):
            return xl, wb, True
        com(wb.Close, False)
        time.sleep(2)
    com(xl.Quit)
    raise RuntimeError(f"Excel só conseguiu abrir {os.path.basename(caminho)} em modo só-leitura "
                       f"(trava do OneDrive ou de outra instância); nada foi gravado.")


def fechar_livro(xl, wb, aberto_por_mim: bool, espera_saida: int = 3) -> None:
    """Fecha SEM salvar (quem gravou já chamou wb.Save) e encerra a instância própria; nunca Close(True).

    O processo do Excel só sai quando todas as referências COM morrem; enquanto vive, o próximo
    GetActiveObject pode cair nele e travar. Espera a saída e, se não sair, encerra o processo
    (é a NOSSA instância escondida; a do usuário nunca passa por aqui)."""
    if not aberto_por_mim:
        return
    pid = _pid_de(xl)
    try:
        com(wb.Close, False)
    finally:
        com(xl.Quit)
        import gc
        gc.collect()          # solta proxies COM pendentes: com eles vivos o Excel não sai sozinho
        if pid:
            for _ in range(max(1, espera_saida)):
                if not _processo_vivo(pid):
                    break
                time.sleep(1)
            else:
                _encerrar_processo(pid)
                for _ in range(10):                 # confirma a morte: um GetActiveObject sobre processo morrendo trava
                    if not _processo_vivo(pid):
                        break
                    time.sleep(0.5)
        time.sleep(2)  # ROT e trava do arquivo levam alguns instantes para sumir depois da saída


def proxima_linha_tabela(ws, nome: str = "Tabela1") -> int:
    """Acrescenta uma linha ao fim da tabela e devolve o número da linha nova na planilha."""
    tab = com(lambda: ws.ListObjects(nome))
    nova = com(tab.ListRows.Add)
    return int(com(lambda: nova.Range.Row))


def coluna_valores(ws, col: int, primeira: int = 8) -> list:
    """Valores da coluna ``col`` de ``primeira`` até a última linha usada da coluna E (processo)."""
    ultima = int(com(lambda: ws.Cells(ws.Rows.Count, 5).End(-4162).Row))  # xlUp
    if ultima < primeira:
        return []
    return [linha[0] for linha in com(get_range, ws, primeira, col, ultima, col)]

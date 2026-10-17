# -*- coding: utf-8 -*-
"""Cópia de arquivos trancados pelo Excel (AutoSave do OneDrive): CopyFileW do Windows em vez de shutil."""
import os

from pc import arquivos, comparar_manual, push, transporte


def test_copiar_substitui_destino_e_preserva_bytes(tmp_path):
    src = tmp_path / "a.bin"; src.write_bytes(b"\x00\x01agenda" * 1000)
    dst = tmp_path / "b.bin"; dst.write_bytes(b"velho")
    arquivos.copiar(str(src), str(dst))
    assert dst.read_bytes() == src.read_bytes()


def test_copia_temporaria_cria_pasta_propria_e_mesmo_nome(tmp_path):
    src = tmp_path / "CACT - Agenda de Prazos.xlsx"; src.write_bytes(b"x")
    tmp = arquivos.copia_temporaria(str(src))
    assert os.path.basename(tmp) == "CACT - Agenda de Prazos.xlsx" and os.path.dirname(tmp) != str(tmp_path) and open(tmp, "rb").read() == b"x"
    arquivos.apagar_copia(tmp)
    assert not os.path.exists(os.path.dirname(tmp))


def test_modulos_que_leem_a_agenda_usam_a_copia_resistente():
    import inspect
    for mod in (push, comparar_manual, transporte):
        fonte = inspect.getsource(mod)
        assert "shutil.copyfile" not in fonte, mod.__name__
        assert "arquivos.cop" in fonte or "copia_temporaria" in fonte, mod.__name__

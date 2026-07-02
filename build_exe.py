# -*- coding: utf-8 -*-
"""
Crea l'eseguibile unico  dist/Convertitore PDF.exe

Gestisce il caso (tipico dei venv su Windows) in cui PyInstaller non include
automaticamente tkinter / _tkinter e i dati Tcl/Tk: qui vengono individuati
nel Python di base e aggiunti esplicitamente.
"""
import os
import sys

import PyInstaller.__main__

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = sys.base_prefix   # installazione Python di base (non il venv)
SEP = ";" if os.name == "nt" else ":"


def find_tcl_tk():
    """Individua le cartelle dati di Tcl e Tk nel Python di base."""
    tcl_data = tk_data = None
    tcl_root = os.path.join(BASE, "tcl")
    if os.path.isdir(tcl_root):
        for name in os.listdir(tcl_root):
            full = os.path.join(tcl_root, name)
            if not os.path.isdir(full):
                continue
            if name.startswith("tcl8") and os.path.isfile(os.path.join(full, "init.tcl")):
                tcl_data = full
            elif name.startswith("tk8"):
                tk_data = full
    return tcl_data, tk_data


def main():
    tcl_data, tk_data = find_tcl_tk()

    args = [
        os.path.join(HERE, "convertitore.py"),
        "--noconfirm", "--clean", "--onefile", "--noconsole",
        "--name", "Convertitore PDF",
        "--icon", os.path.join(HERE, "assets", "icon.ico"),
        "--add-data", f"{os.path.join(HERE, 'assets')}{SEP}assets",
        # tkinter / _tkinter dal Python di base
        "--paths", os.path.join(BASE, "Lib"),
        "--paths", os.path.join(BASE, "DLLs"),
        "--hidden-import", "_tkinter",
        "--collect-all", "tkinter",
        # librerie con file binari
        "--collect-all", "tkinterdnd2",
        "--collect-all", "pypdfium2",
        "--collect-all", "pypdfium2_raw",
    ]
    if tcl_data:
        args += ["--add-data", f"{tcl_data}{SEP}_tcl_data"]
    if tk_data:
        args += ["--add-data", f"{tk_data}{SEP}_tk_data"]

    print("Python di base :", BASE)
    print("Dati Tcl       :", tcl_data)
    print("Dati Tk        :", tk_data)
    print("Avvio di PyInstaller...\n")
    PyInstaller.__main__.run(args)


if __name__ == "__main__":
    main()

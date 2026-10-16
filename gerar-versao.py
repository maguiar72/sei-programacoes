"""Grava _versao.py com a data/hora da compilação (chamado por gerar-exe.bat)."""
from datetime import datetime

with open("_versao.py", "w", encoding="utf-8") as f:
    f.write('VERSAO = "1.1"\nCOMPILACAO = "%s"\n' % datetime.now().strftime("%d/%m/%Y %H:%M"))

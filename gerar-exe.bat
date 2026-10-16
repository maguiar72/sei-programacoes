@echo off
rem Gera dist\SEI-Programacoes.exe (rodar no Windows, com Python 3.10+ no PATH)
python -m pip install --upgrade pyinstaller playwright openpyxl xlwt || goto :erro
taskkill /f /t /im SEI-Programacoes.exe >nul 2>&1
del /q dist\SEI-Programacoes.exe >nul 2>&1
python gerar-versao.py || goto :erro
python -m PyInstaller --onefile --windowed --name SEI-Programacoes --collect-all playwright sei_programacoes.py || goto :erro
echo.
echo Pronto: dist\SEI-Programacoes.exe (janela grafica; embute o Playwright e usa o Edge ja instalado)
pause
exit /b 0
:erro
echo FALHOU
pause
exit /b 1

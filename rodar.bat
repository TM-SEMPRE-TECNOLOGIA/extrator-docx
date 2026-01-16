@echo off
REM ==============================
REM Rodar Extrator Word -> Excel
REM ==============================

REM Garante que rode na pasta do arquivo
cd /d "%~dp0"

REM Executa com Python do sistema
python extrair_itens_docx.py

REM Mantém a janela aberta só se der erro
if errorlevel 1 pause

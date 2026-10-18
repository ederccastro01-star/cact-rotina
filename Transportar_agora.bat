@echo off
rem Transporta JA para a Agenda oficial todas as decisoes (Agendar / Corrigir e agendar) ainda sem Lote,
rem da Paralela de hoje e das dos ultimos 7 dias. O vigia continua como rede de seguranca.
chcp 65001 >nul
cd /d "%~dp0"
python -m pc.agenda_paralela agora
if errorlevel 1 (echo. & echo Houve erro: veja pc\agenda_paralela.log & pause) else (timeout /t 6 >nul)

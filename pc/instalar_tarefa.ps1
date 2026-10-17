# Instala as tarefas agendadas da Agenda Paralela no PC (rodar uma vez, como o usuário).
# 1) "CACT - Agenda Paralela": ao fazer logon, roda o vigia (espera o arquivo do dia, abre, transporta).
# 2) "CACT - Agenda Paralela (repetir)": das 08:00 às 23:00, a cada 15 min, garante que o vigia está vivo
#    (se já houver um em execução, o segundo sai na hora pelo arquivo de trava).
$clone  = Split-Path -Parent $PSScriptRoot
$python = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $python) { $python = (Get-Command python.exe).Source }
$cmd = "`"$python`" -m pc.agenda_paralela vigiar"

schtasks /Delete /TN "CACT - Agenda Paralela" /F 2>$null | Out-Null
schtasks /Delete /TN "CACT - Agenda Paralela (repetir)" /F 2>$null | Out-Null
schtasks /Create /TN "CACT - Agenda Paralela" /SC ONLOGON /TR "cmd /c cd /d `"$clone`" && $cmd" /RL LIMITED /F | Out-Null
schtasks /Create /TN "CACT - Agenda Paralela (repetir)" /SC MINUTE /MO 15 /ST 08:00 /ET 23:00 /TR "cmd /c cd /d `"$clone`" && $cmd" /RL LIMITED /F | Out-Null
schtasks /Query /TN "CACT - Agenda Paralela"
schtasks /Query /TN "CACT - Agenda Paralela (repetir)"
Write-Host "Instalado. Clone: $clone | Python: $python"

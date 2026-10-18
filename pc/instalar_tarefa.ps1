# Instala as tarefas agendadas da Agenda Paralela no PC (rodar uma vez, como o usuário).
# 1) "CACT - Agenda Paralela (repetir)": TODO DIA, das 08:00 às 23:00, a cada 15 min, garante que o vigia está vivo
#    (se já houver um em execução, o segundo sai na hora pelo arquivo de trava pc\vigia.lock).
# 2) "CACT - Agenda Paralela": ao fazer logon (opcional; exige administrador).
#
# Lição de 18/10/2026: "/SC MINUTE /MO 15 /ST 08:00 /ET 23:00" cria um gatilho de UM dia só (EndBoundary na data da
# instalação) e a tarefa nunca mais dispara. O certo é DAILY com repetição: /SC DAILY /ST 08:00 /RI 15 /DU 15:00.
$clone  = Split-Path -Parent $PSScriptRoot
$python = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $python) { $python = (Get-Command python.exe).Source }
$cmd = "`"$python`" -m pc.agenda_paralela vigiar"
$nome = "CACT - Agenda Paralela (repetir)"

schtasks /Delete /TN "CACT - Agenda Paralela" /F 2>$null | Out-Null
schtasks /Delete /TN $nome /F 2>$null | Out-Null
schtasks /Create /TN $nome /SC DAILY /ST 08:00 /RI 15 /DU 15:00 /TR "cmd /c cd /d `"$clone`" && $cmd" /RL LIMITED /F | Out-Null
if ($LASTEXITCODE -ne 0) { Write-Host "ERRO: nao criei a tarefa '$nome'." -ForegroundColor Red; exit 1 }

# Notebook na bateria e PC que estava desligado na hora: a tarefa deve rodar mesmo assim.
$ajustes = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
    -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 16)
Set-ScheduledTask -TaskName $nome -Settings $ajustes | Out-Null

# Gatilho "ao fazer logon" e opcional e exige privilegio de administrador; se falhar, segue sem ele.
schtasks /Create /TN "CACT - Agenda Paralela" /SC ONLOGON /TR "cmd /c cd /d `"$clone`" && $cmd" /RL LIMITED /F 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) { Write-Host "Aviso: gatilho ao logon nao criado (precisa de administrador); a tarefa a cada 15 min cobre o caso." }

$t = Get-ScheduledTask -TaskName $nome
$g = $t.Triggers[0]
Write-Host ("Gatilho: inicio {0} | fim '{1}' | repete a cada {2} por {3}" -f $g.StartBoundary, $g.EndBoundary, $g.Repetition.Interval, $g.Repetition.Duration)
if ($g.EndBoundary) { Write-Host "ERRO: o gatilho tem data de fim; a tarefa pararia de disparar." -ForegroundColor Red; exit 1 }
Write-Host ("Proximo disparo: {0}" -f (Get-ScheduledTaskInfo $nome).NextRunTime)
Write-Host "Instalado. Clone: $clone | Python: $python"

<#
.SYNOPSIS
  Despliega el monitor de energia WUG como tarea programada de Windows.

.DESCRIPTION
  1. Verifica Python >= 3.10
  2. Crea .venv e instala requirements.txt
  3. Crea .env desde .env.example si no existe (y se detiene para que lo llenes)
  4. Verifica que exista el catalogo Excel
  5. Registra la tarea programada (arranca con el equipo, reinicia si falla)

  Ejecutar desde PowerShell *como Administrador* en la carpeta del proyecto:
    .\deploy.ps1
    .\deploy.ps1 -IntervalSec 300 -TaskName "MonitorEnergiaWUG"
    .\deploy.ps1 -SinAlertas   # registra la tarea sin canal de alertas (solo genera el Excel)
    .\deploy.ps1 -Uninstall
#>
[CmdletBinding()]
param(
    [string]$TaskName = "MonitorEnergiaWUG",
    [int]$IntervalSec = 300,
    [switch]$SinAlertas,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$ProjectDir = $PSScriptRoot
$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$Script     = Join-Path $ProjectDir "paso1_filtrar_y_ordenar.py"
$EnvFile    = Join-Path $ProjectDir ".env"
$Catalogo   = Join-Path $ProjectDir "Logs de energía.xlsx"

function Step($msg) { Write-Host "`n==> $msg" -ForegroundColor Cyan }

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { throw "Ejecuta PowerShell como Administrador." }

if ($Uninstall) {
    Step "Eliminando tarea '$TaskName'"
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Listo. (.venv, .env y estado_alertas.json no se tocaron)"
    return
}

Step "Verificando Python >= 3.10"
$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) { throw "Python no esta en PATH. Instalalo desde python.org (marca 'Add to PATH')." }
$ver = [version](& python -c "import sys;print('%d.%d'%sys.version_info[:2])")
if ($ver -lt [version]"3.10") { throw "Se requiere Python 3.10+, encontrado $ver" }
Write-Host "Python $ver OK"

Step "Creando entorno virtual e instalando dependencias"
if (-not (Test-Path $VenvPython)) { & python -m venv (Join-Path $ProjectDir ".venv") }
& $VenvPython -m pip install --quiet --upgrade pip
& $VenvPython -m pip install --quiet -r (Join-Path $ProjectDir "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "Fallo la instalacion de dependencias." }

Step "Verificando configuracion"
if (-not (Test-Path $EnvFile)) {
    Copy-Item (Join-Path $ProjectDir ".env.example") $EnvFile
    # Restringe el .env (contiene credenciales). Se usan SID porque los nombres de grupo
    # dependen del idioma de Windows ("Administrators" vs "Administradores").
    # SYSTEM (S-1-5-18) debe poder leerlo: es la cuenta con la que corre la tarea.
    icacls $EnvFile /inheritance:r /grant:r "$($env:USERNAME):(R,W)" "*S-1-5-18:(F)" "*S-1-5-32-544:(F)" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "No se pudieron aplicar permisos a .env" }
    Write-Warning "Se creo .env. Llena WUG_USER/WUG_PASS (o WUG_COOKIE) y los canales de alerta, luego vuelve a ejecutar deploy.ps1."
    return
}
# Lee el .env linea por linea (tolera BOM, espacios y finales de linea CR/LF/CRLF).
$cfg = @{}
foreach ($line in (Get-Content $EnvFile)) {
    $l = $line.Trim([char]0xFEFF).Trim()
    if ($l -match '^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$') { $cfg[$Matches[1]] = $Matches[2].Trim() }
}
function Has($k) { return ($cfg.ContainsKey($k) -and $cfg[$k] -ne '') }
$hasAuth = (Has 'WUG_COOKIE') -or ((Has 'WUG_USER') -and (Has 'WUG_PASS'))
if (-not $hasAuth) { throw ".env sin credenciales WUG (WUG_USER+WUG_PASS o WUG_COOKIE). Claves leidas: $($cfg.Keys -join ', ')" }
$hasTelegram = (Has 'TELEGRAM_BOT_TOKEN') -and (Has 'TELEGRAM_CHAT_ID')
$hasTeams    = Has 'TEAMS_WEBHOOK_URL'
if ($SinAlertas) {
    Write-Warning "Modo -SinAlertas: el monitor generara el Excel pero NO avisara a nadie."
} elseif (-not ($hasTelegram -or $hasTeams)) {
    throw ".env sin canal de alertas (TELEGRAM_BOT_TOKEN+TELEGRAM_CHAT_ID o TEAMS_WEBHOOK_URL). Sin canal el monitor correria pero nunca avisaria. Si es intencional, usa -SinAlertas."
}
if (-not (Test-Path $Catalogo)) { throw "Falta '$Catalogo' (hoja 'Nodos NI', columna 'Nombre del Nodo')." }

if (-not $SinAlertas) {
    Step "Probando notificaciones"
    & $VenvPython $Script --test-notify "Deploy monitor energia: prueba de notificacion"
    if ($LASTEXITCODE -ne 0) { throw "La prueba de notificacion fallo. Revisa .env (Telegram/Teams)." }
}

Step "Registrando tarea programada '$TaskName'"
$action  = New-ScheduledTaskAction -Execute $VenvPython `
           -Argument "`"$Script`" --loop --interval-sec $IntervalSec" -WorkingDirectory $ProjectDir
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew `
            -ExecutionTimeLimit ([TimeSpan]::Zero) `
            -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force | Out-Null

Start-ScheduledTask -TaskName $TaskName
Start-Sleep -Seconds 5
Write-Host "Estado: $((Get-ScheduledTask -TaskName $TaskName).State)" -ForegroundColor Green
Write-Host "Logs:   $(Join-Path $ProjectDir 'energia_loop.log')"
Write-Host "Quitar: .\deploy.ps1 -Uninstall"

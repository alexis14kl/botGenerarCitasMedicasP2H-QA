@echo off
setlocal

title Login medico automatico (timing reforzado)

set "URL=https://mp-stg.telesalud.gob.sv/"
set "WAIT_OPEN=14"
set "WAIT_BETWEEN=300"

echo [1/5] Abriendo portal...
start "" "%URL%"

echo [2/5] Esperando carga inicial (%WAIT_OPEN%s)...
timeout /t %WAIT_OPEN% /nobreak >nul

echo [3/5] Ejecutando autologin con reintentos por timing...
powershell -NoProfile -ExecutionPolicy Bypass -Command "
$ws = New-Object -ComObject WScript.Shell
$step = [int]$env:WAIT_BETWEEN

function SleepMs([int]$ms){ Start-Sleep -Milliseconds $ms }
function FocusLogin(){
  return $ws.AppActivate('Account login')
}
function K([string]$k, [int]$ms){
  $ws.SendKeys($k)
  SleepMs $ms
}
function DoLoginPass(){
  K 'MEDICO09' $step
  K '{TAB}' $step
  K 'ISSS202' $step
  K '{TAB}' ($step + 150)

  # Empresa
  K '{DOWN}' $step
  K '{ENTER}' ($step + 120)

  # Departamento
  K '{TAB}' ($step + 120)
  K '{DOWN}' $step
  K '{ENTER}' ($step + 120)

  # Workaround: re-seleccionar empresa
  K '+{TAB}' ($step + 100)
  K '{DOWN}' $step
  K '{ENTER}' ($step + 120)

  # Ir a boton iniciar sesion
  K '{TAB}' ($step + 120)
  K '{TAB}' ($step + 120)
  K '{ENTER}' ($step + 350)
}

$ok = $false
for($try=1; $try -le 3; $try++){
  if(-not (FocusLogin())){
    Write-Host ('No se detecta Account login en intento ' + $try + '. Esperando...')
    SleepMs 1500
    continue
  }

  SleepMs 600
  DoLoginPass
  SleepMs 2200

  # Si ya no puede activar Account login, asumimos que entro
  if(-not (FocusLogin())){
    $ok = $true
    break
  }

  # Si sigue en login, un Enter extra y esperar
  K '{ENTER}' 600
  SleepMs 1500
  if(-not (FocusLogin())){
    $ok = $true
    break
  }

  Write-Host ('Sigue en login tras intento ' + $try + ', reintentando...')
  SleepMs 1200
}

if($ok){ exit 0 } else { exit 3 }
"
set "RC=%ERRORLEVEL%"

echo [4/5] Codigo de salida: %RC%
if "%RC%"=="0" (
  echo Login probablemente exitoso.
) else (
  echo No se completo el login automatico. Sugerencia: subir WAIT_OPEN o WAIT_BETWEEN.
)

echo [5/5] Fin.
pause
endlocal

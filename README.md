# Monitor de Energia WUG

Script Python para procesar eventos `UP/DOWN` desde WhatsUp Gold/WUG, normalizar nodos contra un catalogo Excel, generar un archivo procesado y enviar alertas incrementales por Telegram o Power Automate/Teams.

## Que hace

- Consulta el mapa de energia en WUG.
- Filtra lineas `UP`, `DOWN`, `SVCUP` y `SVCDOWN`.
- Normaliza nombres de nodos y aplica aliases conocidos.
- Cruza los eventos contra `Logs de energia.xlsx`, hoja `Nodos NI`.
- Genera `logs_energia_procesados.xlsx` con las hojas `eventos_up_down` y `Detalle`.
- Mantiene estado incremental en `estado_alertas.json` para no repetir alertas.
- Envia alarma cuando un nodo lleva mas de `ALERT_AFTER_MIN` minutos abajo.
- Envia restablecimiento cuando vuelve `UP` una caida que ya habia sido notificada.

## Archivos esperados

El script busca estos archivos en la misma carpeta:

- `paso1_filtrar_y_ordenar.py`
- `Logs de energia.xlsx`
- `TICKETS DE ENERGIA DESDE OCT 2025.xlsx` *(actualmente esta definido en el script, pero no se usa)*

El proceso genera:

- `logs_energia_procesados.xlsx`
- `estado_alertas.json`
- `energia_loop.log`

Por seguridad, los Excel, Power BI, logs y estado local estan ignorados por Git en `.gitignore`.

## Paso a paso para implementar

### 1. Instalar Python

Instala Python 3.10 o superior en el servidor/equipo donde va a correr el monitor.

Verifica:

```powershell
python --version
```

### 2. Crear entorno virtual

Desde la carpeta del proyecto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Instalar dependencias

```powershell
pip install -r requirements.txt
```

### 4. Copiar archivos operativos

Coloca en la misma carpeta del script:

- `Logs de energia.xlsx`
- `TICKETS DE ENERGIA DESDE OCT 2025.xlsx`, si se decide usarlo mas adelante.

La hoja del catalogo debe llamarse `Nodos NI` y debe contener la columna `Nombre del Nodo`.

### 5. Configurar credenciales

Puedes usar variables de entorno o crear un archivo `.env` local tomando como base `.env.example`.

Minimo requerido:

```powershell
$env:WUG_USER="usuario"
$env:WUG_PASS="password"
```

O usando cookie:

```powershell
$env:WUG_COOKIE="cookie_de_sesion"
```

Opcional para notificaciones:

```powershell
$env:TELEGRAM_BOT_TOKEN="token"
$env:TELEGRAM_CHAT_ID="chat_id"
$env:TEAMS_WEBHOOK_URL="url_power_automate"
```

### 6. Probar una ejecucion

```powershell
python .\paso1_filtrar_y_ordenar.py
```

Si todo esta correcto, debe generar o actualizar `logs_energia_procesados.xlsx`.

### 7. Probar notificaciones

```powershell
python .\paso1_filtrar_y_ordenar.py --test-notify "Prueba monitor energia"
```

### 8. Ejecutar en loop

```powershell
python .\paso1_filtrar_y_ordenar.py --loop --interval-sec 300
```

El intervalo puede bajar automaticamente si hay caidas abiertas o eventos cerca de generar alerta.

### 9. Ejecutarlo como tarea programada de Windows

Crea una tarea en el Programador de tareas con:

- Programa: ruta a `python.exe` del entorno virtual, por ejemplo `C:\ruta\proyecto\.venv\Scripts\python.exe`
- Argumentos: `paso1_filtrar_y_ordenar.py --loop --interval-sec 300`
- Iniciar en: carpeta del proyecto

Configura la tarea para iniciar con Windows o con el usuario operativo.

## Variables utiles

| Variable | Default | Uso |
| --- | --- | --- |
| `WUG_USER` | vacio | Usuario WUG |
| `WUG_PASS` | vacio | Password WUG |
| `WUG_COOKIE` | vacio | Cookie de sesion WUG |
| `ALERT_AFTER_MIN` | `10` | Minutos antes de alertar una caida |
| `LOOP_INTERVAL_SEC` | `300` | Intervalo base del loop |
| `EXPORT_ONLY_IF_NEW` | `1` | Evita reexportar Excel si no hay eventos nuevos |
| `LOG_LEVEL` | `INFO` | Nivel de logs |
| `TEAMS_WEBHOOK_URL` | vacio | URL de Power Automate |
| `TELEGRAM_BOT_TOKEN` | vacio | Token bot Telegram |
| `TELEGRAM_CHAT_ID` | vacio | Chat destino Telegram |

## Notas de operacion

- Si `estado_alertas.json` queda con una caida abierta que ya fue resuelta fuera del rango de logs, puede ser necesario revisarlo manualmente.
- El archivo `TICKETS DE ENERGIA DESDE OCT 2025.xlsx` esta referenciado como constante, pero no participa en el procesamiento actual.
- No subas `.env`, logs, estados ni archivos Excel con datos internos a repositorios publicos.

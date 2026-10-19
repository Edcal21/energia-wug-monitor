# Monitor de Energia WUG

Script Python para procesar eventos `UP/DOWN` desde WhatsUp Gold/WUG, normalizar nodos contra un catalogo Excel, generar un archivo procesado y enviar alertas incrementales por Telegram o Power Automate/Teams.

## Que hace

- Consulta el mapa de energia en WUG.
- Filtra lineas `UP`, `DOWN`, `SVCUP` y `SVCDOWN`.
- Normaliza nombres de nodos y aplica aliases conocidos.
- Cruza los eventos contra `Logs de energía.xlsx`, hoja `Nodos NI`.
- Genera `logs_energia_procesados.xlsx` con las hojas `eventos_up_down` y `Detalle`.
- Mantiene estado incremental en `estado_alertas.json` para no repetir alertas.
- Envia alarma cuando un nodo lleva mas de `ALERT_AFTER_MIN` minutos abajo.
- Envia restablecimiento cuando vuelve `UP` una caida que ya habia sido notificada.

## Archivos esperados

El script busca estos archivos en la misma carpeta:

- `paso1_filtrar_y_ordenar.py`
- `Logs de energía.xlsx`
- `TICKETS DE ENERGIA DESDE OCT 2025.xlsx` *(actualmente esta definido en el script, pero no se usa)*

El proceso genera:

- `logs_energia_procesados.xlsx`
- `estado_alertas.json`
- `energia_loop.log`

Por seguridad, los Excel, Power BI, logs y estado local estan ignorados por Git en `.gitignore`.

## Paso a paso para implementar

## Despliegue rapido (Windows)

Desde PowerShell **como Administrador** en la carpeta del proyecto:

```powershell
.\deploy.ps1
```

La primera vez crea `.venv`, instala dependencias y genera `.env`; llena las credenciales y vuelve a ejecutarlo. En la segunda corrida valida configuracion y catalogo, envia una notificacion de prueba y registra la tarea programada `MonitorEnergiaWUG` (inicia con el equipo, reinicia si falla). Para quitarla: `.\deploy.ps1 -Uninstall`.

Los pasos manuales siguen abajo por si prefieres hacerlo a mano.

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

- `Logs de energía.xlsx`
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
| `ESCALATION_AFTER_MIN` | `60` | Minutos de caida activa antes de enviar el aviso de escalamiento (`0` lo desactiva) |
| `ESCALATION_CONTACT` | vacio | Persona que debe dar seguimiento; aparece en el mensaje de escalamiento |
| `WUG_URLS` | ver `.env.example` | Endpoints de WUG separados por coma (se lee del `.env`) |
| `PLAN_ENTRADA_AFTER_MIN` | `0` | Minutos de caida antes del aviso de entrada del plan de autonomia (subir si hay muchos parpadeos) |
| `PLAN_AUTONOMIA_FILE` | `datos/plan_autonomia.csv` | CSV de autonomia y movilizacion por nodo |
| `HEARTBEAT_AFTER_MIN` | `30` | Minutos sin un ciclo exitoso antes de avisar que el monitor mismo esta ciego (`0` lo desactiva) |
| `HEARTBEAT_REPEAT_MIN` | `120` | Cada cuantos minutos recuerda la alerta mientras siga sin funcionar |
| `BASELINE_MAX_AGE_HOURS` | `24` | Al arrancar sin `estado_alertas.json`, las caidas en curso (ultimo evento DOWN) de las ultimas N horas se registran y alarman; las mas viejas se ignoran |
| `LOOP_INTERVAL_SEC` | `300` | Intervalo base del loop |
| `EXPORT_ONLY_IF_NEW` | `1` | Evita reexportar Excel si no hay eventos nuevos |
| `LOG_LEVEL` | `INFO` | Nivel de logs |
| `TEAMS_WEBHOOK_URL` | vacio | URL de Power Automate |
| `TELEGRAM_BOT_TOKEN` | vacio | Token bot Telegram |
| `TELEGRAM_CHAT_ID` | vacio | Chat destino Telegram |

## Plan de autonomia (hora limite por nodo)

`datos/plan_autonomia.csv` guarda, por nodo, la **autonomia** de baterias/generador y el **tiempo de movilizacion** de la cuadrilla desde su base (ambos en horas; viene del Excel `Energia_Plan_Mundial.xlsx`, sin coordenadas). Con eso, para cada caida:

```
hora limite = hora de caida + autonomia - movilizacion
```

- **Aviso de entrada** (al detectar la caida): cuanto tiempo queda para normalizar o mover la cuadrilla, la hora limite y la hora estimada de apagado. Si el nodo tiene acceso restringido (p. ej. solo diurno) lo indica.
- **Aviso de hora limite**: si llega esa hora y el nodo no ha levantado, envia "MUEVE LA CUADRILLA YA o se apaga el nodo".
- Si autonomia <= movilizacion, o el monitor se entera tarde, sale un solo aviso urgente en vez de los dos.
- Cada aviso sale una vez por caida (se guarda en `estado_alertas.json`). Se suman a la alarma de 10 min y al escalamiento de 60 min, no los reemplazan.

Los nodos se cruzan por **nombre normalizado** (sin tildes ni mayusculas). Columnas del CSV: `nodo, sector, autonomia_hrs, movilizacion_hrs, acceso, alias_wug`. Si el nombre en WUG/catalogo no coincide, pon el nombre alterno en `alias_wug` (varios separados por `|`). El CSV se recarga solo al editarlo.

Para ver que nodos no hacen match (en ambos sentidos), sin red ni notificaciones:

```powershell
python .\paso1_filtrar_y_ordenar.py --verificar-plan
```

Un nodo que cae y no esta en el plan queda registrado como WARNING en el log y no genera aviso de hora limite.

## Heartbeat (el monitor vigila al monitor)

Si WUG o la VPN dejan de responder, el monitor no puede ver caidas de nodos y el silencio parece "todo bien". Por eso, si pasan `HEARTBEAT_AFTER_MIN` minutos sin un ciclo exitoso, envia por el mismo canal una **ALERTA** (con el ultimo error), un recordatorio cada `HEARTBEAT_REPEAT_MIN` minutos y un aviso **RESUELTO** al recuperarse. Si el envio falla, reintenta en el siguiente ciclo.

Limite: si el proceso muere o el equipo se apaga/duerme, no hay nadie que avise. Para eso hace falta un servidor siempre encendido o un monitoreo externo.

## Diagnostico: "no sono la alarma de un nodo"

Solo lectura (no toca estado, Excel ni envia nada). Muestra las lineas del log de WUG que mencionan el nodo y si el filtro las descarta o las acepta:

```powershell
python .\paso1_filtrar_y_ordenar.py --diagnosticar tipitapa
```

## Notas de operacion

- Si `estado_alertas.json` queda con una caida abierta que ya fue resuelta fuera del rango de logs, puede ser necesario revisarlo manualmente.
- El archivo `TICKETS DE ENERGIA DESDE OCT 2025.xlsx` esta referenciado como constante, pero no participa en el procesamiento actual.
- No subas `.env`, logs, estados ni archivos Excel con datos internos a repositorios publicos.

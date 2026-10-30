# file: paso1_filtrar_y_ordenar.py
"""
Procesa logs UP/DOWN desde WUG, normaliza nodos con catálogo, genera Excel y envía alertas
incrementales por Telegram y/o Power Automate (Teams/Flow).

Bucle infinito:
  python paso1_filtrar_y_ordenar.py --loop --interval-sec 300

Logging:
  set LOG_FILE=energia_loop.log
  set LOG_LEVEL=INFO
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import time as time_module
from dataclasses import dataclass
from datetime import time
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Dict
from urllib.parse import urljoin

import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
from requests.exceptions import RequestException

# ================= CONFIG =================
DEFAULT_WUG_URLS = [
    "http://monitoreo.alliance.com.ni:8022/map.asp",
    "http://whatsup.alliance.com.ni:8022/map.asp",
]
WUG_PARAMS = {"MapViewMode": "0", "map": "4- RED ENERGIA COMERCIAL.wup"}
BASE_DIR = Path(__file__).resolve().parent

def _load_dotenv(path: Path) -> None:
    """Carga KEY=VALUE de un .env local sin pisar variables ya definidas en el entorno."""
    if not path.is_file():
        return
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip().strip("\"'"))


_load_dotenv(BASE_DIR / ".env")

WUG_URLS = [
    u.strip()
    for u in os.getenv("WUG_URLS", ",".join(DEFAULT_WUG_URLS)).split(",")
    if u.strip()
]

USERNAME = os.getenv("WUG_USER")
PASSWORD = os.getenv("WUG_PASS")
WUG_COOKIE = os.getenv("WUG_COOKIE", "").strip()
if not WUG_COOKIE and (not USERNAME or not PASSWORD):
    raise ValueError("Faltan variables de entorno WUG_USER / WUG_PASS o WUG_COOKIE.")

CATALOGO_SHEET = "Nodos NI"
CATALOGO_FILE = str(BASE_DIR / "Logs de energía.xlsx")
TICKETS_FILE = str(BASE_DIR / "TICKETS DE ENERGIA DESDE OCT 2025.xlsx")
OUT_FILE = str(BASE_DIR / "logs_energia_procesados.xlsx")

SHEET_EVENTOS = "eventos_up_down"
SHEET_DETALLE = "Detalle"

# ======= Telegram / Power Automate (ENV) =======
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "").strip()

TEAMS_WEBHOOK_URL = os.getenv("TEAMS_WEBHOOK_URL", "").strip()

PA_PAYLOAD_MODE = os.getenv("PA_PAYLOAD_MODE", "json").strip().lower()  # json | text
PA_TEXT_FIELD = os.getenv("PA_TEXT_FIELD", "text").strip()
PA_EXTRA_FIELDS_RAW = os.getenv("PA_EXTRA_FIELDS", "").strip()  # JSON string: {"k":"v", ...}
PA_RAISE_ON_ERROR = os.getenv("PA_RAISE_ON_ERROR", "1").strip() in {"1", "true", "True", "yes", "YES"}

ALERT_AFTER_MIN = int(os.getenv("ALERT_AFTER_MIN", "10"))

# ======= Escalamiento: segundo aviso si la caida sigue activa (0 = desactivado) =======
ESCALATION_AFTER_MIN = int(os.getenv("ESCALATION_AFTER_MIN", "60"))
ESCALATION_CONTACT = os.getenv("ESCALATION_CONTACT", "").strip()
# Al arrancar sin estado, las caidas en curso con DOWN mas reciente que esto se siguen alertando.
BASELINE_MAX_AGE_HOURS = float(os.getenv("BASELINE_MAX_AGE_HOURS", "24"))

# ======= Plan de autonomia: hora limite = caida + autonomia - movilizacion de la cuadrilla =======
PLAN_FILE = Path(os.getenv("PLAN_AUTONOMIA_FILE", str(BASE_DIR / "datos" / "plan_autonomia.csv")))
# Minutos de caida antes del aviso de entrada (0 = al detectar la caida; subir si hay muchos parpadeos).
PLAN_ENTRADA_AFTER_MIN = float(os.getenv("PLAN_ENTRADA_AFTER_MIN", "0"))

# ======= Heartbeat: avisa si el monitor mismo deja de funcionar (0 = desactivado) =======
HEARTBEAT_AFTER_MIN = float(os.getenv("HEARTBEAT_AFTER_MIN", "30"))
HEARTBEAT_REPEAT_MIN = float(os.getenv("HEARTBEAT_REPEAT_MIN", "120"))

STATE_FILE = Path(os.getenv("ENERGIA_ALERT_STATE", str(BASE_DIR / "estado_alertas.json")))
EXPORT_ONLY_IF_NEW = os.getenv("EXPORT_ONLY_IF_NEW", "1").strip() in {"1", "true", "True", "yes", "YES"}

# ======= Dynamic interval (ENV) =======
DYN_INTERVAL_ENABLED = os.getenv("DYN_INTERVAL_ENABLED", "1").strip() in {"1", "true", "True", "yes", "YES"}
DYN_INTERVAL_MIN_SEC = int(os.getenv("DYN_INTERVAL_MIN_SEC", "60"))  # mínimo absoluto del loop
DYN_INTERVAL_OPEN_SEC = int(os.getenv("DYN_INTERVAL_OPEN_SEC", "120"))  # hay open_down cualquiera
DYN_INTERVAL_PREALERT_SEC = int(os.getenv("DYN_INTERVAL_PREALERT_SEC", "60"))  # cerca de cruzar ALERT_AFTER_MIN
DYN_INTERVAL_PREALERT_WINDOW_MIN = float(os.getenv("DYN_INTERVAL_PREALERT_WINDOW_MIN", "3"))  # ventana "cerca"

# ======= WUG HTTP robustness (ENV) =======
WUG_TIMEOUT_SEC = int(os.getenv("WUG_TIMEOUT_SEC", "60"))
WUG_RETRIES = int(os.getenv("WUG_RETRIES", "3"))
WUG_RETRY_SLEEP_SEC = int(os.getenv("WUG_RETRY_SLEEP_SEC", "5"))

# ======= Logging (ENV) =======
LOG_FILE = os.getenv("LOG_FILE", str(BASE_DIR / "energia_loop.log")).strip()
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").strip().upper()
LOG_MAX_BYTES = int(os.getenv("LOG_MAX_BYTES", str(5 * 1024 * 1024)))  # 5MB
LOG_BACKUP_COUNT = int(os.getenv("LOG_BACKUP_COUNT", "5"))


# ================= LOGGER =================
def setup_logging() -> logging.Logger:
    logger = logging.getLogger("energia")
    if logger.handlers:
        return logger

    logger.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))

    fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File (rotating)
    try:
        fh = RotatingFileHandler(
            LOG_FILE,
            maxBytes=LOG_MAX_BYTES,
            backupCount=LOG_BACKUP_COUNT,
            encoding="utf-8",
        )
        fh.setFormatter(fmt)
        fh.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))
        logger.addHandler(fh)
    except Exception:
        # If file handler fails, continue with console logging only.
        pass

    # Console
    ch = logging.StreamHandler(stream=sys.stdout)
    ch.setFormatter(fmt)
    ch.setLevel(getattr(logging, LOG_LEVEL, logging.INFO))
    logger.addHandler(ch)

    logger.propagate = False
    return logger


logger = setup_logging()

# ================= ALIAS =================
ALIAS_MAP_RAW = {
    "plaza inter": "Plaza Inter JDC",
    "nueva guinea": "Nueva Guinea",
    "san miguelito": "San Miguelito",
    "crucero": "El Crucero",
    "el crucero": "El Crucero",
    "oficina ufinet": "Planes de Altamira",
    "ufinet oficina": "Planes de Altamira",
    "planes de altamira": "Planes de Altamira",
    "pmaderas": "Playa Maderas",
    "p maderas": "Playa Maderas",
    "playa maderas": "Playa Maderas",
    "sj del sur": "San Juan del Sur",
    "san juan del sur": "San Juan del Sur",
    "stomas": "Santo Tomás",
    "stomás": "Santo Tomás",
    "santo tomas": "Santo Tomás",
    "santo tomás": "Santo Tomás",
}

# ================= UTILIDADES =================
def norm_text(x: Any) -> str:
    return str(x).lower().strip()


def norm_series(s: pd.Series) -> pd.Series:
    return (
        s.astype(str)
        .str.lower()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )


def clean_key_series(s: pd.Series) -> pd.Series:
    x = norm_series(s)
    x = x.str.replace(r"^(mini\s+)?n\s+", "", regex=True)
    x = x.str.replace(r"^nodo\s+", "", regex=True)
    x = x.str.replace(r"\bnva\b", "nueva", regex=True)
    x = x.str.replace(r"\bsto\b", "santo", regex=True)
    x = x.str.replace(r"\bs\s+miguelito\b", "san miguelito", regex=True)
    x = x.str.replace(r"\s+", " ", regex=True).str.strip()
    return x


def is_yes(x: Any) -> bool:
    v = norm_text(x)
    return v in ["si", "sí", "s", "yes", "y", "1", "true"]


def clasificar_turno(dt: pd.Timestamp) -> str:
    t = dt.time()
    if time(6, 0) <= t <= time(13, 59, 59):
        return "DIURNO 1"
    if time(14, 0) <= t <= time(21, 59, 59):
        return "DIURNO 2"
    return "NOCTURNO"


BASURA_KEYWORDS = [
    "interface(",
    "gigabitethernet",
    "hundredgig",
    "tengige",
    "te0/",
    "hu0/",
    "ncs540",
    "ncs 540",
    "mpl",
    "nni-",
    "trk-",
    "port",
    "pv:",
    "efp",
    "vlan",
    "rectificador",
    "router",
    "wimax",
    "sw2948",
    "swr2948",
    "sw_2924",
    "sw_r2924",
    "sw_r2948",
    "sw r2924",
    "sw r2948",
    "pe yota",
    "yota nodos unica",
]


def is_basura_candidato(txt: str) -> bool:
    s = norm_text(txt)
    if not s:
        return True
    return any(k in s for k in BASURA_KEYWORDS)


# ================= WUG =================
def _extract_filter_link(html: str) -> str | None:
    m = re.search(r'href="([^"]*logview\.asp\?[^"]*ViewMapFilter=true[^"]*)"', html, flags=re.I)
    if m:
        return m.group(1)
    m = re.search(
        r'href="([^"]*logview\.asp\?[^"]*)"\s*>\s*Filter\s+Activity\s+Log\s+data',
        html,
        flags=re.I,
    )
    if m:
        return m.group(1)
    return None


def _to_logview_url(wug_url: str) -> str:
    return re.sub(r"/map\.asp(\?.*)?$", "/logview.asp", wug_url, flags=re.I)


def _get_with_retries(
    session: requests.Session,
    url: str,
    *,
    params: dict[str, Any] | None = None,
    context: str,
) -> requests.Response:
    attempts = max(1, WUG_RETRIES + 1)

    for attempt in range(1, attempts + 1):
        try:
            response = session.get(url, params=params, timeout=WUG_TIMEOUT_SEC)
            if response.status_code == 401:
                raise PermissionError(f"401 Unauthorized en {url}")
            response.raise_for_status()
            return response
        except PermissionError:
            raise
        except RequestException as exc:
            if attempt >= attempts:
                raise RuntimeError(
                    f"No se pudo consultar WUG en '{context}' tras {attempts} intentos. "
                    "El servidor cerró la conexión o no respondió correctamente."
                ) from exc

            sleep_sec = max(1, WUG_RETRY_SLEEP_SEC * attempt)
            logger.warning(
                "Fallo consultando WUG en %s (intento %s/%s): %s. Reintentando en %s s...",
                context,
                attempt,
                attempts,
                exc,
                sleep_sec,
            )
            time_module.sleep(sleep_sec)


def _check_log_access(text: str, url: str) -> str:
    """WUG responde HTTP 200 con una pagina 'Error: Unauthorized Access' cuando el usuario no tiene permiso sobre el log."""
    if "Unauthorized Access" in text:
        raise PermissionError(f"WUG respondio 'Unauthorized Access' en la vista de log: el usuario no tiene permiso ({url})")
    return text


def fetch_filtered_log_text(session: requests.Session) -> str:
    errors: list[str] = []

    for wug_url in WUG_URLS:
        logger.info("Probando endpoint WUG: %s", wug_url)
        try:
            r0 = _get_with_retries(session, wug_url, params=WUG_PARAMS, context=f"actividad inicial [{wug_url}]")

            href = _extract_filter_link(r0.text)
            if href:
                filtered_url = urljoin(wug_url, href)
                r1 = _get_with_retries(session, filtered_url, context=f"vista filtrada [{wug_url}]")
                logger.info("Endpoint WUG autenticado correctamente: %s", wug_url)
                return _check_log_access(r1.text, filtered_url)

            fallback_url = _to_logview_url(wug_url)
            params = dict(WUG_PARAMS)
            params["ViewMapFilter"] = "true"
            r1 = _get_with_retries(
                session,
                fallback_url,
                params=params,
                context=f"fallback ViewMapFilter=true [{fallback_url}]",
            )
            logger.info("Endpoint WUG autenticado correctamente: %s", wug_url)
            return _check_log_access(r1.text, fallback_url)
        except PermissionError as exc:
            logger.warning("Credenciales rechazadas en %s: %s", wug_url, exc)
            errors.append(f"{wug_url}: {exc}")
        except Exception as exc:
            logger.warning("Fallo consultando %s: %s", wug_url, exc)
            errors.append(f"{wug_url}: {exc}")

    raise RuntimeError(
        "No fue posible autenticarse ni consultar WUG en ninguno de los endpoints configurados. "
        + " | ".join(errors)
    )


def is_up_down_line(line: str) -> bool:
    line = line.strip()
    if not re.match(r"^\d{8}\s+\d{6}\b", line):
        return False
    parts = line.split()
    if len(parts) < 3:
        return False
    return parts[2].upper() in {"UP", "DOWN", "SVCUP", "SVCDOWN"}


def parse_up_down(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    pat = re.compile(
        r"^(?P<date>\d{8})\s+(?P<time>\d{6})\s+(?P<estado>SVCUP|SVCDOWN|UP|DOWN)\s+(?P<msg>.+)$",
        re.IGNORECASE,
    )
    df = df.join(df["raw"].astype(str).str.extract(pat))
    df = df.dropna(subset=["estado"]).copy()

    df["estado"] = df["estado"].astype(str).str.upper().replace({"SVCDOWN": "DOWN", "SVCUP": "UP"})
    df["datetime"] = pd.to_datetime(df["date"] + df["time"], format="%Y%m%d%H%M%S", errors="coerce")

    msg = df["msg"].astype(str)
    cand = msg.str.extract(
        r"\bN:(.+?)(?:\s+F:|\s+TIP\b|\s+\d{1,3}(?:\.\d{1,3}){3}\b|$)",
        expand=False,
    )
    cand2 = msg.str.extract(
        r"\bN\s+(.+?)(?:\s+F:|\s+TIP\b|\s+\d{1,3}(?:\.\d{1,3}){3}\b|$)",
        expand=False,
    )
    df["candidato"] = cand.fillna(cand2).fillna("").astype(str).str.strip()
    df["ip"] = df["raw"].astype(str).str.extract(r"(\d{1,3}(?:\.\d{1,3}){3})", expand=False)
    return df


# ================= CATÁLOGO =================
def load_catalogo_excel(path: str, sheet_name: str) -> pd.DataFrame:
    if not Path(path).exists():
        raise FileNotFoundError(f"No existe el archivo de catálogo: {path}")

    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    header_row = raw.iloc[1].tolist()
    data = raw.iloc[2:].copy()
    data.columns = header_row

    if "Nombre del Nodo" not in data.columns:
        raise ValueError("No encuentro la columna 'Nombre del Nodo' en el catálogo.")

    cat = pd.DataFrame(
        {
            "nodo_catalogo": data["Nombre del Nodo"].astype(str).str.strip(),
            "latitud": pd.to_numeric(data.get("Latitud"), errors="coerce"),
            "longitud": pd.to_numeric(data.get("Longitud"), errors="coerce"),
            "sector": data.get("Sector"),
            "tiene_dwdm": data.get("Tiene DWDM ? (S/N)"),
        }
    ).dropna(subset=["nodo_catalogo"])

    cat["key_match"] = clean_key_series(cat["nodo_catalogo"])
    return cat


# ================= NOTIFY =================
def _parse_extra_fields(raw: str) -> Dict[str, Any]:
    if not raw:
        return {}
    try:
        obj = json.loads(raw)
        return obj if isinstance(obj, dict) else {}
    except Exception:
        return {}


def _post_power_automate(url: str, text: str) -> None:
    extra = _parse_extra_fields(PA_EXTRA_FIELDS_RAW)

    if PA_PAYLOAD_MODE == "text":
        r = requests.post(
            url,
            data=text.encode("utf-8"),
            headers={"Content-Type": "text/plain; charset=utf-8"},
            timeout=20,
        )
    else:
        payload = {PA_TEXT_FIELD: text, **extra}
        r = requests.post(
            url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=20,
        )

    if not (200 <= r.status_code < 300):
        raise RuntimeError(f"Power Automate POST failed: {r.status_code} - {r.text}")


def notify(text: str) -> None:
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        r = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "text": text}, timeout=15)
        r.raise_for_status()

    if TEAMS_WEBHOOK_URL:
        try:
            _post_power_automate(TEAMS_WEBHOOK_URL, text)
        except Exception:
            if PA_RAISE_ON_ERROR:
                raise


# ================= PLAN DE AUTONOMIA =================
def _strip_parens(key: str) -> str:
    return re.sub(r"\s*\(.*?\)\s*", " ", key).strip()


def _key(name: Any) -> str:
    return clean_key_series(pd.Series([str(name)])).iloc[0]


@dataclass
class PlanNodo:
    nodo: str
    autonomia_hrs: float
    movilizacion_hrs: float
    acceso: str = ""


class PlanAutonomia:
    """Autonomia y tiempo de movilizacion por nodo, cruzados por nombre normalizado.

    Orden de busqueda: nombre exacto o alias_wug -> nombre sin parentesis (solo si es unico).
    """

    def __init__(self, df: pd.DataFrame):
        self.items: list[PlanNodo] = []
        self.exact: dict[str, PlanNodo] = {}
        loose: dict[str, list[PlanNodo]] = {}
        for _, r in df.iterrows():
            item = PlanNodo(
                nodo=str(r["nodo"]).strip(),
                autonomia_hrs=float(r["autonomia_hrs"]),
                movilizacion_hrs=float(r["movilizacion_hrs"]),
                acceso=str(r.get("acceso") or "").strip(),
            )
            self.items.append(item)
            names = [item.nodo] + [a for a in str(r.get("alias_wug") or "").split("|") if a.strip()]
            for n in names:
                self.exact[_key(n)] = item
            loose.setdefault(_strip_parens(_key(item.nodo)), []).append(item)
        self.loose = {k: v[0] for k, v in loose.items() if len(v) == 1}
        self.ambiguos = {k: [i.nodo for i in v] for k, v in loose.items() if len(v) > 1}

    def buscar(self, nodo: str) -> PlanNodo | None:
        k = _key(nodo)
        return self.exact.get(k) or self.loose.get(_strip_parens(k))


def load_plan(path: Path = PLAN_FILE) -> PlanAutonomia | None:
    """Carga el CSV del plan. Si falta o es invalido, el aviso de hora limite queda desactivado (con ERROR en el log)."""
    if not path.is_file():
        logger.warning("Plan de autonomia no encontrado (%s): avisos de hora limite desactivados.", path)
        return None
    try:
        df = pd.read_csv(path, encoding="utf-8-sig", dtype={"acceso": str, "alias_wug": str})
        faltan = {"nodo", "autonomia_hrs", "movilizacion_hrs"} - set(df.columns)
        if faltan:
            raise ValueError(f"faltan columnas {sorted(faltan)}")
        df = df.dropna(subset=["nodo", "autonomia_hrs"])
        df["movilizacion_hrs"] = df["movilizacion_hrs"].fillna(0)  # vacio solo tiene sentido con autonomia 0
        for c in ("autonomia_hrs", "movilizacion_hrs"):
            df[c] = pd.to_numeric(df[c], errors="raise")
        df = df.fillna({"acceso": "", "alias_wug": ""})
        return PlanAutonomia(df)
    except Exception:
        logger.exception("Plan de autonomia ilegible (%s): avisos de hora limite desactivados.", path)
        return None


_NUNCA = object()
_PLAN_CACHE: dict[str, Any] = {"mtime": _NUNCA, "plan": None}


def get_plan() -> PlanAutonomia | None:
    """Plan en cache; se recarga solo si cambia el CSV (editarlo no requiere reiniciar y evita repetir el warning)."""
    try:
        mtime = PLAN_FILE.stat().st_mtime
    except OSError:
        mtime = None
    if mtime != _PLAN_CACHE["mtime"]:
        _PLAN_CACHE.update(mtime=mtime, plan=load_plan(PLAN_FILE))
    return _PLAN_CACHE["plan"]


def _fmt_dt(ts: Any) -> str:
    return pd.to_datetime(ts).strftime("%d/%m %H:%M")


def _fmt_dur(minutos: float) -> str:
    m = int(round(abs(minutos)))
    h, mm = divmod(m, 60)
    return f"{h} h {mm:02d} min" if h else f"{mm} min"


def _plan_info(info: dict[str, Any], nodo: str, down_dt: pd.Timestamp, plan: PlanAutonomia | None) -> bool:
    """Guarda en la caida su hora limite (una sola vez). False si el nodo no esta en el plan."""
    if info.get("limite_dt"):
        return True
    p = plan.buscar(nodo) if plan else None
    if p is None:
        return False
    info["autonomia_hrs"] = p.autonomia_hrs
    info["movilizacion_hrs"] = p.movilizacion_hrs
    info["acceso"] = p.acceso
    info["apagado_dt"] = str(down_dt + pd.Timedelta(hours=p.autonomia_hrs))
    info["limite_dt"] = str(down_dt + pd.Timedelta(hours=p.autonomia_hrs - p.movilizacion_hrs))
    return True


def _acceso_linea(info: dict[str, Any]) -> str:
    acc = str(info.get("acceso") or "")
    return f"Acceso: {acc}\n" if acc and acc.strip().lower() != "24/7" else ""


def avisos_plan(nodo: str, info: dict[str, Any], down_dt: pd.Timestamp, now: pd.Timestamp, plan: PlanAutonomia | None) -> None:
    """Aviso de entrada (cuanto tiempo hay) y aviso de hora limite vencida. Cada uno sale una sola vez por caida."""
    if plan is None or not _plan_info(info, nodo, down_dt, plan):
        if plan is not None and not info.get("plan_sin_dato"):
            logger.warning("Nodo sin dato en el plan de autonomia: %s (no hay aviso de hora limite).", nodo)
            info["plan_sin_dato"] = True
        return

    limite = pd.to_datetime(info["limite_dt"])
    apagado = pd.to_datetime(info["apagado_dt"])
    restante_min = (limite - now).total_seconds() / 60.0
    base = (
        f"Nodo: {nodo}\n"
        f"Caida: {_fmt_dt(down_dt)}\n"
        f"Autonomia: {_fmt_dur(info['autonomia_hrs'] * 60)} | Movilizacion: {_fmt_dur(info['movilizacion_hrs'] * 60)}\n"
        f"{_acceso_linea(info)}"
    )

    if not info.get("plan_entrada"):
        if (now - down_dt).total_seconds() / 60.0 < PLAN_ENTRADA_AFTER_MIN:
            return
        if info["autonomia_hrs"] <= 0:
            # Nodo sin autonomia: no hay tiempo que calcular, la cuadrilla sale desde la caida.
            notify(f"[PLAN AUTONOMIA] MUEVE LA CUADRILLA YA: nodo sin autonomia\nNodo: {nodo}\nCaida: {_fmt_dt(down_dt)}\n{_acceso_linea(info)}Se apaga con la caida de energia; la cuadrilla debe salir desde la hora de caida.")
            info["plan_entrada"] = True
            info["plan_limite"] = True
            return
        if restante_min > 0:
            msg = (
                f"[PLAN AUTONOMIA] Caida de nodo\n{base}"
                f"Tiempo para normalizar o mover la cuadrilla: {_fmt_dur(restante_min)}\n"
                f"Hora limite: {_fmt_dt(limite)}\n"
                f"Sin energia el nodo se apaga a las {_fmt_dt(apagado)}."
            )
            notify(msg)
            info["plan_entrada"] = True
            return
        # Ya no hay margen (autonomia <= movilizacion, o el monitor se enteró tarde): un solo aviso urgente.
        msg = (
            f"[PLAN AUTONOMIA] MUEVE LA CUADRILLA YA o se apaga el nodo\n{base}"
            f"La hora limite era {_fmt_dt(limite)} (hace {_fmt_dur(restante_min)}).\n"
            f"Apagado estimado: {_fmt_dt(apagado)}."
        )
        notify(msg)
        info["plan_entrada"] = True
        info["plan_limite"] = True
        return

    if not info.get("plan_limite") and restante_min <= 0:
        msg = (
            f"[PLAN AUTONOMIA] MUEVE LA CUADRILLA YA o se apaga el nodo\n{base}"
            f"Se cumplio la hora limite ({_fmt_dt(limite)}) y el nodo sigue caido.\n"
            f"Apagado estimado: {_fmt_dt(apagado)} (en {_fmt_dur((apagado - now).total_seconds() / 60.0)})."
        )
        notify(msg)
        info["plan_limite"] = True


def verificar_plan(plan: PlanAutonomia | None) -> int:
    """Solo lectura: cruza el plan contra el catalogo y reporta nodos sin match en ambos sentidos."""
    if plan is None:
        print("No se pudo cargar el plan (ver log).")
        return 1
    cat = load_catalogo_excel(CATALOGO_FILE, CATALOGO_SHEET)
    nombres = sorted(set(cat["nodo_catalogo"]))
    en_plan, sin_plan = [], []
    for n in nombres:
        (en_plan if plan.buscar(n) else sin_plan).append(n)
    usados = {id(plan.buscar(n)) for n in nombres if plan.buscar(n)}
    plan_sin_nodo = [i.nodo for i in plan.items if id(i) not in usados]
    print(f"Plan: {len(plan.items)} nodos | Catalogo: {len(nombres)} nodos | con match: {len(en_plan)}")
    print(f"\nNodos del catalogo SIN dato en el plan ({len(sin_plan)}) -> no tendran aviso de hora limite:")
    for n in sin_plan:
        print(f"  - {n}")
    print(f"\nNodos del plan que NO aparecen en el catalogo ({len(plan_sin_nodo)}) -> revisar nombre o llenar alias_wug:")
    for n in plan_sin_nodo:
        print(f"  - {n}")
    if plan.ambiguos:
        print("\nNombres ambiguos en el plan (varios sitios con el mismo nombre base; usa el nombre completo o alias_wug):")
        for k, v in plan.ambiguos.items():
            print(f"  - {k}: {v}")
    return 0


# ================= ESTADO =================
def _empty_state() -> dict[str, Any]:
    return {"last_seen_datetime": None, "open_down": {}}


def _load_state() -> dict[str, Any]:
    if not STATE_FILE.exists():
        return _empty_state()
    try:
        s = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        if not isinstance(s, dict):
            raise ValueError("el estado no es un objeto JSON")
        s.setdefault("last_seen_datetime", None)
        s.setdefault("open_down", {})
        return s
    except Exception as e:
        # Estado ilegible: guardar copia para poder revisarlo y avisar en el log,
        # en vez de volver a baseline en silencio y perder las caidas abiertas.
        backup = STATE_FILE.with_name(f"{STATE_FILE.name}.corrupto-{time_module.strftime('%Y%m%d-%H%M%S')}")
        try:
            STATE_FILE.replace(backup)
            logger.error("Estado ilegible (%s). Copia guardada en %s; se reinicia el baseline.", e, backup)
        except OSError:
            logger.error("Estado ilegible (%s) y no se pudo respaldar; se reinicia el baseline.", e)
        return _empty_state()


def _save_state(state: dict[str, Any]) -> None:
    # Escritura atomica: se escribe a un temporal y se reemplaza, asi un corte de
    # energia a mitad de escritura nunca deja el JSON truncado.
    tmp = STATE_FILE.with_name(STATE_FILE.name + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    # En Windows un antivirus o indexador puede tener el archivo abierto un instante: reintentar.
    for intento in range(5):
        try:
            os.replace(tmp, STATE_FILE)
            return
        except PermissionError:
            if intento == 4:
                raise
            time_module.sleep(0.5)


def _parse_state_ts(x: Any) -> pd.Timestamp | None:
    if not x:
        return None
    ts = pd.to_datetime(x, errors="coerce")
    if pd.isna(ts):
        return None
    return ts


def _seed_open_down(df: pd.DataFrame, now: pd.Timestamp) -> dict[str, Any]:
    """Caidas aun abiertas segun el log: nodos cuyo ultimo evento es DOWN y fue hace <= BASELINE_MAX_AGE_HOURS."""
    seeded: dict[str, Any] = {}
    max_age = pd.Timedelta(hours=BASELINE_MAX_AGE_HOURS)
    for nodo, g in df.sort_values("datetime").groupby("nodo"):
        if str(g.iloc[-1]["estado"]).upper() != "DOWN":
            continue
        ups = g[g["estado"].astype(str).str.upper() == "UP"]
        tramo = g if ups.empty else g[g["datetime"] > ups["datetime"].max()]
        first_down = tramo.iloc[0]  # primer DOWN de la racha final (DOWN repetidos no reinician la caida)
        down_dt = pd.to_datetime(first_down["datetime"])
        if now - down_dt > max_age:
            continue
        seeded[str(nodo)] = {
            "down_dt": str(down_dt),
            "notified": False,
            "ip": str(first_down.get("ip") or ""),
            "raw_down": str(first_down.get("raw") or ""),
        }
    return seeded


def update_alerts_incremental(
    df_events_full: pd.DataFrame, state: dict[str, Any], plan: PlanAutonomia | None = None
) -> tuple[dict[str, Any], pd.DataFrame]:
    """
    - ALERTA DOWN: solo si dura >= ALERT_AFTER_MIN (1 sola vez por caída)
    - RESTABLECIDO: solo si esa caída ya había sido notificada (notified=True)
    - baseline (sin estado): setea last_seen y registra como abiertas las caidas en curso
      (ultimo evento DOWN de las ultimas BASELINE_MAX_AGE_HOURS h); esas si alarman normal
    """
    df = df_events_full.dropna(subset=["nodo", "estado", "datetime"]).copy()
    last_seen = _parse_state_ts(state.get("last_seen_datetime"))
    if df.empty:
        # Sin eventos en el log: aun asi hay que evaluar las caidas ya abiertas (alarma, plan, escalamiento).
        df_new = df.iloc[0:0].copy()
        open_down = state.get("open_down", {})
        if not open_down:
            return state, df_new
    elif last_seen is None:
        state["last_seen_datetime"] = str(df["datetime"].max())
        open_down: dict[str, Any] = _seed_open_down(df, pd.Timestamp.now())
        df_new = df.iloc[0:0].copy()
        logger.info("Baseline sin estado previo: %s caida(s) en curso registradas.", len(open_down))
    else:
        df_new = df[df["datetime"] > last_seen].copy()
        df_new = df_new.sort_values(["datetime", "nodo"]).reset_index(drop=True)
        open_down = state.get("open_down", {})

    for _, r in df_new.iterrows():
        nodo = str(r["nodo"])
        estado = str(r["estado"]).upper()
        dt = pd.to_datetime(r["datetime"], errors="coerce")
        if pd.isna(dt):
            continue

        if estado == "DOWN":
            if nodo in open_down:
                continue  # DOWN repetido: no reiniciar la caida ni volver a alarmar
            open_down[nodo] = {
                "down_dt": str(dt),
                "notified": False,
                "ip": str(r.get("ip") or ""),
                "raw_down": str(r.get("raw") or ""),
            }

        elif estado == "UP":
            if nodo in open_down:
                info = open_down[nodo]
                down_dt = pd.to_datetime(info.get("down_dt"), errors="coerce")

                if info.get("notified") and pd.notna(down_dt) and dt >= down_dt:
                    mins = round((dt - down_dt).total_seconds() / 60, 1)
                    msg = (
                        f"✅ RESTABLECIDO Energía\n"
                        f"Nodo: {nodo}\n"
                        f"Caída: {down_dt}\n"
                        f"UP: {dt}\n"
                        f"Duración: {mins} min"
                    )
                    notify(msg)

                del open_down[nodo]

    now = pd.Timestamp.now()
    for nodo, info in list(open_down.items()):
        down_dt = pd.to_datetime(info.get("down_dt"), errors="coerce")
        if pd.isna(down_dt):
            continue
        mins = (now - down_dt).total_seconds() / 60.0

        avisos_plan(nodo, info, down_dt, now, plan)
        open_down[nodo] = info

        if mins >= ALERT_AFTER_MIN and not info.get("notified"):
            mins_r = round(mins, 1)
            msg = (
                f"⚠️ ALARMA Energía (DOWN > {ALERT_AFTER_MIN} min)\n"
                f"Nodo: {nodo}\n"
                f"IP: {info.get('ip','')}\n"
                f"Desde: {down_dt}\n"
                f"Duración: {mins_r} min\n"
                f"RAW_DOWN: {info.get('raw_down','')}"
            )
            notify(msg)
            info["notified"] = True
            open_down[nodo] = info

        if ESCALATION_AFTER_MIN > 0 and mins >= ESCALATION_AFTER_MIN and not info.get("escalated_60"):
            seguimiento = f" de {ESCALATION_CONTACT}" if ESCALATION_CONTACT else ""
            msg = (
                f"[ESCALAMIENTO {ESCALATION_AFTER_MIN} MIN] Energia\n"
                f"Nodo: {nodo}\n"
                f"IP: {info.get('ip', '')}\n"
                f"Desde: {down_dt}\n"
                f"Duracion: {round(mins, 1)} min\n"
                f"La caida sigue activa. Requiere seguimiento{seguimiento}."
            )
            notify(msg)
            info["escalated_60"] = True
            open_down[nodo] = info

    if not df_new.empty:
        state["last_seen_datetime"] = str(df_new["datetime"].max())
    state["open_down"] = open_down
    return state, df_new


# ================= LOOP HELPERS =================
def compute_dynamic_interval(base_interval_sec: int, state: dict[str, Any]) -> int:
    """
    Heurística:
      - Si hay caídas no notificadas cerca de cruzar ALERT_AFTER_MIN => intervalo muy corto
      - Si hay open_down (cualquier) => intervalo medio
      - Si no hay nada => intervalo base
    """
    if not DYN_INTERVAL_ENABLED:
        return max(1, base_interval_sec)

    open_down: dict[str, Any] = state.get("open_down", {}) or {}
    if not open_down:
        return max(DYN_INTERVAL_MIN_SEC, base_interval_sec)

    now = pd.Timestamp.now()
    prealert_threshold = max(0.0, ALERT_AFTER_MIN - DYN_INTERVAL_PREALERT_WINDOW_MIN)

    has_prealert = False
    has_any_open = False

    for _, info in open_down.items():
        down_dt = pd.to_datetime(info.get("down_dt"), errors="coerce")
        if pd.isna(down_dt):
            continue
        has_any_open = True

        mins = (now - down_dt).total_seconds() / 60.0
        if (not info.get("notified")) and mins >= prealert_threshold:
            has_prealert = True
            break

        limite = pd.to_datetime(info.get("limite_dt"), errors="coerce")
        if pd.notna(limite) and not info.get("plan_limite"):
            # Hora limite a punto de cumplirse: revisar mas seguido para que el aviso no llegue tarde.
            if (limite - now).total_seconds() / 60.0 <= DYN_INTERVAL_PREALERT_WINDOW_MIN:
                has_prealert = True
                break

    if has_prealert:
        return max(DYN_INTERVAL_MIN_SEC, min(base_interval_sec, DYN_INTERVAL_PREALERT_SEC))

    if has_any_open:
        return max(DYN_INTERVAL_MIN_SEC, min(base_interval_sec, DYN_INTERVAL_OPEN_SEC))

    return max(DYN_INTERVAL_MIN_SEC, base_interval_sec)


# ================= CLI =================
def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(add_help=True)
    p.add_argument(
        "--test-notify",
        metavar="TEXT",
        type=str,
        default="",
        help="Envía una notificación de prueba (Telegram/Power Automate) y sale.",
    )
    p.add_argument(
        "--diagnosticar",
        metavar="NODO",
        type=str,
        default="",
        help="Solo lectura: muestra las lineas del log WUG que mencionan NODO y como las trata el filtro.",
    )
    p.add_argument(
        "--cambiar-credenciales",
        action="store_true",
        help="Pide por consola el usuario y la contrasena de WUG y los guarda en el .env (apto para SSH).",
    )
    p.add_argument(
        "--verificar-plan",
        action="store_true",
        help="Solo lectura: cruza el plan de autonomia contra el catalogo y lista los nodos sin match.",
    )
    p.add_argument("--loop", action="store_true", help="Ejecuta el proceso en bucle infinito.")
    p.add_argument(
        "--interval-sec",
        type=int,
        default=int(os.getenv("LOOP_INTERVAL_SEC", "300")),
        help="Segundos de espera base entre ejecuciones cuando se usa --loop.",
    )
    return p.parse_args()


@dataclass(frozen=True)
class RunSummary:
    events_total: int
    new_events: int
    open_down_total: int
    exported: bool


# ================= RUN ONCE =================
def _ciclo_sin_eventos() -> "RunSummary":
    """WUG respondio pero sin lineas UP/DOWN utiles: no es un error, pero se avisa en el log y se siguen evaluando las caidas abiertas."""
    logger.warning("WUG no devolvio lineas UP/DOWN validas; se evaluan solo las caidas ya abiertas.")
    state = _load_state()
    vacio = pd.DataFrame(columns=["nodo", "estado", "datetime"])
    state, _ = update_alerts_incremental(vacio, state, get_plan())
    _save_state(state)
    return RunSummary(
        events_total=0,
        new_events=0,
        open_down_total=len((state.get("open_down") or {}).keys()),
        exported=False,
    )


def run_once() -> RunSummary:
    cat = load_catalogo_excel(CATALOGO_FILE, CATALOGO_SHEET)

    logger.info("Descargando logs FILTRADOS desde WUG...")
    with requests.Session() as s:
        s.trust_env = False
        s.headers.update(
            {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) energia-wug-monitor/1.0",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Connection": "close",
            }
        )

        if WUG_COOKIE:
            s.headers["Cookie"] = WUG_COOKIE
            logger.info("Usando autenticacion por cookie de sesion WUG.")
        else:
            s.auth = HTTPBasicAuth(USERNAME, PASSWORD)
            logger.info("Usando autenticacion por credenciales WUG.")

        texto = fetch_filtered_log_text(s)

    lineas = [l.strip() for l in texto.splitlines() if is_up_down_line(l)]
    raw = pd.DataFrame({"raw": lineas})
    logger.info("Lineas UP/DOWN candidatas: %s", len(raw))

    if raw.empty:
        # Guardar lo que respondio WUG para poder ver si es una vista vacia, un login o un cambio de formato.
        volcado = BASE_DIR / "wug_respuesta_vacia.txt"
        try:
            volcado.write_text(texto[:200_000], encoding="utf-8")
            logger.warning("Respuesta de WUG sin lineas UP/DOWN (%s caracteres) guardada en %s", len(texto), volcado)
        except OSError:
            logger.exception("No pude guardar la respuesta vacia de WUG en %s", volcado)
        return _ciclo_sin_eventos()

    df = parse_up_down(raw)
    df = df[df["candidato"].str.len().ge(3)].copy()
    df = df[~df["candidato"].apply(is_basura_candidato)].copy()
    logger.info("Logs UP/DOWN con nodo válidos: %s", len(df))

    if df.empty:
        return _ciclo_sin_eventos()

    alias_keys = pd.Series(list(ALIAS_MAP_RAW.keys()))
    alias_norm = clean_key_series(alias_keys)
    alias_map = dict(zip(alias_norm, ALIAS_MAP_RAW.values()))

    df["key_match"] = clean_key_series(df["candidato"])
    df["nodo_oficial"] = df["key_match"].map(alias_map).fillna(df["candidato"].astype(str).str.strip())
    df["key_match"] = clean_key_series(df["nodo_oficial"])

    df = df.merge(
        cat[["key_match", "nodo_catalogo", "latitud", "longitud", "sector", "tiene_dwdm"]],
        on="key_match",
        how="left",
    )

    df["nodo"] = df["nodo_catalogo"].fillna(df["nodo_oficial"])
    df["tipo_nodo"] = df["tiene_dwdm"].apply(lambda x: "DWDM" if is_yes(x) else "RCE")
    df.loc[df["nodo_catalogo"].isna(), "tipo_nodo"] = "OTRO"

    state = _load_state()
    state, df_new = update_alerts_incremental(df, state, get_plan())
    _save_state(state)

    open_down_total = len((state.get("open_down") or {}).keys())
    logger.info("Eventos nuevos desde last_seen: %s | open_down: %s", len(df_new), open_down_total)

    if EXPORT_ONLY_IF_NEW and len(df_new) == 0 and os.path.exists(OUT_FILE):
        logger.info("Sin eventos nuevos: omito export Excel (%s).", OUT_FILE)
        return RunSummary(events_total=len(df), new_events=0, open_down_total=open_down_total, exported=False)

    df = df.sort_values(["nodo", "datetime"]).reset_index(drop=True)
    df["estado_next"] = df.groupby("nodo")["estado"].shift(-1)
    df["dt_next"] = df.groupby("nodo")["datetime"].shift(-1)
    df["raw_next"] = df.groupby("nodo")["raw"].shift(-1)

    mask = (
        (df["estado"] == "DOWN")
        & (df["estado_next"] == "UP")
        & df["datetime"].notna()
        & df["dt_next"].notna()
        & (df["dt_next"] >= df["datetime"])
    )
    caidas = df.loc[mask].copy()
    dur_sec = (caidas["dt_next"] - caidas["datetime"]).dt.total_seconds()
    caidas["dur_min"] = dur_sec / 60
    caidas["dur_horas"] = dur_sec / 3600

    df_detalle = pd.DataFrame(
        {
            "Nodo": caidas["nodo"],
            "Tipo Nodo": caidas["tipo_nodo"],
            "Latitud": caidas["latitud"],
            "Longitud": caidas["longitud"],
            "Sector": caidas["sector"],
            "Tiene DWDM": caidas.get("tiene_dwdm"),
            "Fecha Caida": caidas["datetime"].dt.date,
            "Hora Caida": caidas["datetime"].dt.time,
            "Fecha Restablecimiento": caidas["dt_next"].dt.date,
            "Hora Restablecimiento": caidas["dt_next"].dt.time,
            "Tiempo de Caída (minutos)": caidas["dur_min"].round(0).astype("Int64"),
            "Tiempo de Caída (horas)": caidas["dur_horas"].round(2),
            "Horario de Afectación": caidas["datetime"].apply(clasificar_turno),
            "IP": caidas["ip"],
            "RAW_DOWN": caidas["raw"],
            "RAW_UP": caidas["raw_next"],
        }
    )

    Path(OUT_FILE).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_FILE, engine="openpyxl", mode="w") as w:
        df.to_excel(w, index=False, sheet_name=SHEET_EVENTOS)
        df_detalle.to_excel(w, index=False, sheet_name=SHEET_DETALLE)

    logger.info("Export Excel OK: %s | Eventos: %s | Caídas: %s", OUT_FILE, len(df), len(df_detalle))
    return RunSummary(
        events_total=len(df),
        new_events=len(df_new),
        open_down_total=open_down_total,
        exported=True,
    )


# ================= HEARTBEAT =================
class Heartbeat:
    """Avisa si el propio monitor lleva demasiado tiempo sin completar un ciclo con exito.

    Sin esto, una caida de WUG/VPN deja al monitor ciego y el silencio parece "todo bien".
    Un aviso al superar `after_min`, recordatorio cada `repeat_min` y aviso al recuperarse.
    after_min <= 0 lo desactiva.
    """

    def __init__(self, after_min: float, repeat_min: float, now: float) -> None:
        self.after_sec = after_min * 60
        self.repeat_sec = repeat_min * 60
        self.last_ok = now
        self.last_alert: float | None = None

    @staticmethod
    def _fmt(sec: float) -> str:
        return f"{int(sec // 60)} min"

    def ok(self, now: float) -> str | None:
        msg = None
        if self.last_alert is not None:
            msg = (
                "RESUELTO: el monitor de energia volvio a completar ciclos "
                f"(estuvo {self._fmt(now - self.last_ok)} sin datos)."
            )
        self.last_ok = now
        self.last_alert = None
        return msg

    def failed(self, now: float, error: str) -> str | None:
        if self.after_sec <= 0:
            return None
        down = now - self.last_ok
        if down < self.after_sec:
            return None
        if self.last_alert is not None and now - self.last_alert < self.repeat_sec:
            return None
        self.last_alert = now
        return (
            f"ALERTA: el monitor de energia lleva {self._fmt(down)} sin completar un ciclo. "
            "NO se estan vigilando las caidas de nodos. "
            f"Ultimo error: {error[:300]}"
        )


def _heartbeat_notify(msg: str | None) -> bool:
    """Envia el aviso. True si se envio (o no habia nada que enviar), False si fallo."""
    if not msg:
        return True
    logger.warning("Heartbeat: %s", msg)
    try:
        notify(msg)
        return True
    except Exception:
        logger.exception("No se pudo enviar el aviso de heartbeat")
        return False


# ================= DIAGNOSTICO =================
def diagnosticar_nodo(nombre: str, max_lineas: int = 40) -> None:
    """Solo lectura: no toca estado, Excel ni notificaciones."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")

    with requests.Session() as s:
        s.trust_env = False
        if WUG_COOKIE:
            s.headers["Cookie"] = WUG_COOKIE
        else:
            s.auth = HTTPBasicAuth(USERNAME, PASSWORD)
        texto = fetch_filtered_log_text(s)

    clave = norm_series(pd.Series([nombre])).iloc[0]
    todas = texto.splitlines()
    coinciden = [l.strip() for l in todas if clave in norm_series(pd.Series([l])).iloc[0]]
    print(f"Lineas totales del log: {len(todas)} | lineas que mencionan '{nombre}': {len(coinciden)}")
    if not coinciden:
        print("-> WUG no devolvio ninguna linea de ese nodo: el script nunca lo ve (mapa/filtro de WUG).")
        return

    # El log de WUG puede venir ordenado de nuevo a viejo (o al reves): ordenar por fecha y hora.
    def _clave(linea: str) -> str:
        return " ".join(linea.split()[:2])

    coinciden.sort(key=_clave, reverse=True)
    eventos = [l for l in coinciden if is_up_down_line(l)]
    otras = len(coinciden) - len(eventos)
    print(f"Eventos UP/DOWN: {len(eventos)} | otras lineas (alertas fallidas, etc.): {otras}")
    print(f"Rango de fechas de las lineas del nodo: {_clave(coinciden[-1])} -> {_clave(coinciden[0])}")
    print(f"Se muestran los {min(max_lineas, len(eventos))} eventos UP/DOWN mas recientes:")

    for linea in eventos[:max_lineas]:
        print("-" * 70)
        print(linea)
        d = parse_up_down(pd.DataFrame({"raw": [linea]}))
        if d.empty:
            print("   [DESCARTADA] no coincide con el formato esperado")
            continue
        cand = d.iloc[0]["candidato"]
        if len(cand) < 3:
            print(f"   [DESCARTADA] no se pudo extraer el nombre del nodo (candidato='{cand}')")
        elif is_basura_candidato(cand):
            palabra = next(k for k in BASURA_KEYWORDS if k in norm_text(cand))
            print(f"   [DESCARTADA] filtro de basura: candidato='{cand}' contiene '{palabra}'")
        else:
            print(f"   [OK] {d.iloc[0]['estado']} {d.iloc[0]['datetime']} nodo='{cand}'")


# ================= CREDENCIALES =================
def cambiar_credenciales(path: Path | None = None) -> int:
    """Pide usuario y contrasena de WUG por consola (la contrasena no se muestra) y los escribe en el .env.

    Sirve para hacerlo por SSH, sin editor grafico. Se reescribe el archivo en el mismo lugar (no se
    reemplaza por uno nuevo) para conservar los permisos restringidos del .env, y se deja WUG_COOKIE vacia
    porque la cookie tiene prioridad sobre usuario/contrasena.
    """
    import getpass

    path = path or (BASE_DIR / ".env")
    if not path.is_file():
        print(f"No existe {path}. Primero ejecuta deploy.ps1 para crearlo.")
        return 1

    user = input("Usuario nuevo de WUG: ").strip()
    pw = getpass.getpass("Contrasena nueva de WUG (no se muestra al escribir): ")
    pw2 = getpass.getpass("Repite la contrasena: ")
    if not user or not pw:
        print("El usuario y la contrasena no pueden estar vacios. No se cambio nada.")
        return 1
    if pw != pw2:
        print("Las contrasenas no coinciden. No se cambio nada.")
        return 1
    if "\n" in user + pw or "\r" in user + pw:
        print("El usuario y la contrasena no pueden tener saltos de linea. No se cambio nada.")
        return 1
    if pw != pw.strip() or pw[0] in "\"'" or pw[-1] in "\"'":
        # El loader del .env recorta espacios y comillas en los extremos del valor.
        print("La contrasena empieza o termina con espacio o comilla, y el monitor la leeria mal. No se cambio nada.")
        return 1

    nuevos = {"WUG_USER": user, "WUG_PASS": pw, "WUG_COOKIE": ""}
    pendientes = dict(nuevos)
    salida: list[str] = []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        clave = raw.strip().partition("=")[0].strip()
        if "=" in raw and clave in pendientes and not raw.lstrip().startswith("#"):
            salida.append(f"{clave}={pendientes.pop(clave)}")
        else:
            salida.append(raw)
    for clave, valor in pendientes.items():
        salida.append(f"{clave}={valor}")

    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write("\r\n".join(salida) + "\r\n")
    print(f"Listo: usuario '{user}' guardado en {path} (WUG_COOKIE quedo vacia).")
    print("Reinicia la tarea para que lo use: schtasks /End /TN MonitorEnergiaWUG y luego schtasks /Run /TN MonitorEnergiaWUG")
    return 0


# ================= MAIN =================
def main() -> int:
    args = _parse_args()

    if args.diagnosticar:
        diagnosticar_nodo(args.diagnosticar)
        return 0

    if args.cambiar_credenciales:
        return cambiar_credenciales()

    if args.verificar_plan:
        return verificar_plan(get_plan())

    if args.test_notify:
        notify(args.test_notify)
        logger.info("test-notify enviado.")
        return 0

    if args.interval_sec < 1:
        raise ValueError("--interval-sec debe ser mayor o igual a 1.")

    if args.loop:
        logger.info("Modo bucle activo. Intervalo base: %s seg. Ctrl+C para detener.", args.interval_sec)
        heartbeat = Heartbeat(HEARTBEAT_AFTER_MIN, HEARTBEAT_REPEAT_MIN, time_module.time())
        try:
            while True:
                started_at = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                logger.info("=== Inicio de ciclo: %s ===", started_at)

                try:
                    summary = run_once()
                    logger.info(
                        "Resumen ciclo | events_total=%s new_events=%s open_down=%s exported=%s",
                        summary.events_total,
                        summary.new_events,
                        summary.open_down_total,
                        summary.exported,
                    )
                    if summary.events_total == 0:
                        # WUG responde pero sin una sola linea UP/DOWN: casi siempre es el servidor WUG caido o
                        # una vista rota, no "sin novedades". Cuenta como ciclo fallido para el heartbeat.
                        alert = heartbeat.failed(
                            time_module.time(),
                            "WUG respondio sin lineas UP/DOWN (servidor WUG caido o vista vacia)",
                        )
                        if not _heartbeat_notify(alert):
                            heartbeat.last_alert = None
                    else:
                        _heartbeat_notify(heartbeat.ok(time_module.time()))
                except Exception as exc:
                    logger.exception("Error en ciclo")
                    alert = heartbeat.failed(time_module.time(), f"{type(exc).__name__}: {exc}")
                    if not _heartbeat_notify(alert):
                        heartbeat.last_alert = None  # no se envio: reintentar en el siguiente ciclo

                state = _load_state()
                sleep_sec = compute_dynamic_interval(args.interval_sec, state)
                logger.info("Esperando %s segundos para el siguiente ciclo...", sleep_sec)
                time_module.sleep(sleep_sec)
        except KeyboardInterrupt:
            logger.info("Proceso detenido por el usuario.")
        return 0

    # Single-run
    try:
        run_once()
        return 0
    except Exception:
        logger.exception("La ejecución terminó con error")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

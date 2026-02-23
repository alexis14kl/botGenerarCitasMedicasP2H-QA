from __future__ import annotations

import os
import re
import shutil
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import Page, sync_playwright

URL = os.getenv("START_URL", "https://mp-stg.telesalud.gob.sv/")
USER = os.getenv("BOT_USER", "MEDICO09")
PASSWORD = os.getenv("BOT_PASSWORD", "ISSS202")

ONLY_LOGIN = os.getenv("ONLY_LOGIN", "0") == "1"
TIMEOUT_SCALE = min(1.20, max(0.25, float(os.getenv("TIMEOUT_SCALE", "0.72"))))
MIN_WAIT_MS = min(250, max(0, int(os.getenv("MIN_WAIT_MS", "30"))))
SLOW_MO_MS = min(500, max(0, int(os.getenv("SLOW_MO_MS", "100"))))
HOLD_AFTER_LOGIN_MS = int(os.getenv("HOLD_AFTER_LOGIN_MS", "60000"))
NODE_FLOW_RETRIES = max(1, int(os.getenv("NODE_FLOW_RETRIES", "3")))

LIVE_LOG_PATH = Path(os.getenv("LIVE_LOG_PATH", str(Path.cwd() / "login_medico_python_live.log")))
PATIENT_KEYS_FILE = Path(os.getenv("PATIENT_KEYS_FILE", str(Path(__file__).with_name("patient_keys.txt"))))
DEFAULT_PATIENT_KEYS = [
    "00955873-3",
    "06169373-5",
    "05608981-6",
    "06416857-7",
    "05400186-2",
    "B04676303",
    "B02661296",
    "B01700785",
    "B00838396",
    "B00491489",
]


def normalize_patient_key(raw: str) -> str:
    return re.sub(r"[^A-Z0-9-]", "", (raw or "").upper().strip().replace("'", "").replace('"', ""))


def is_likely_patient_key(key: str) -> bool:
    return bool(re.fullmatch(r"[A-Z]?[0-9][0-9-]{2,}", key))


def parse_patient_keys(raw_text: str) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for chunk in re.split(r"[\r\n,\s;]+", raw_text or ""):
        key = normalize_patient_key(chunk)
        if not key or key in seen or not is_likely_patient_key(key):
            continue
        seen.add(key)
        out.append(key)
    return out


def load_patient_keys() -> tuple[list[str], str]:
    try:
        if PATIENT_KEYS_FILE.exists():
            file_keys = parse_patient_keys(PATIENT_KEYS_FILE.read_text(encoding="utf-8", errors="ignore"))
            if file_keys:
                return file_keys, f"file:{PATIENT_KEYS_FILE}"
    except Exception:
        pass

    env_keys = parse_patient_keys(os.getenv("PATIENT_KEYS", ""))
    if env_keys:
        return env_keys, "env:PATIENT_KEYS"

    return DEFAULT_PATIENT_KEYS.copy(), "fallback:default_keys"


PATIENT_KEYS, PATIENT_KEYS_SOURCE = load_patient_keys()


def log(msg: str) -> None:
    line = f"[{datetime.now().isoformat(timespec='seconds')}] {msg}"
    print(msg)
    try:
        LIVE_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with LIVE_LOG_PATH.open("a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except Exception:
        pass


def scaled_ms(ms: int | float) -> int:
    n = int(ms)
    if n >= 30000:
        return n
    return max(MIN_WAIT_MS, int(round(n * TIMEOUT_SCALE)))


def wait(page: Page, ms: int | float, critical: bool = False) -> None:
    page.wait_for_timeout(int(ms) if critical else scaled_ms(ms))


def activate_ddl(page: Page, ddl_id: str, settle_ms: int = 120) -> None:
    root = page.locator(f"#{ddl_id}").first
    root.wait_for(state="visible", timeout=12000)
    root.click(force=True)
    wait(page, settle_ms, critical=True)


def read_ddl_state(page: Page, ddl_id: str) -> dict[str, Any]:
    return page.evaluate(
        """
        (ddlId) => {
          try {
            const ddl = window.$find && window.$find(ddlId);
            if (!ddl) return { ok: false, selectedValue: '', selectedText: '' };
            const selected = ddl.get_selectedItem ? ddl.get_selectedItem() : null;
            return {
              ok: true,
              selectedValue: selected && selected.get_value ? selected.get_value() : '',
              selectedText: selected && selected.get_text ? selected.get_text() : ''
            };
          } catch {
            return { ok: false, selectedValue: '', selectedText: '' };
          }
        }
        """,
        ddl_id,
    )


def select_ddl_via_telerik(page: Page, ddl_id: str, value: str) -> dict[str, Any]:
    return page.evaluate(
        """
        ({ ddlId, expectedValue }) => {
          try { if (typeof window.Page_ClientValidate === 'function') window.Page_ClientValidate(); } catch {}
          const btn = document.getElementById('ctl00_usercontrol2_T9500_Login_input');
          if (btn) {
            btn.disabled = false;
            btn.removeAttribute('disabled');
          }
          try {
            const ddl = window.$find && window.$find(ddlId);
            if (!ddl) return { ok: false, reason: 'ddl_not_found', value: '', text: '' };
            const item = ddl.findItemByValue ? ddl.findItemByValue(expectedValue) : null;
            if (!item) return { ok: false, reason: 'item_not_found', value: '', text: '' };
            if (ddl.trackChanges) ddl.trackChanges();
            item.select();
            if (ddl.commitChanges) ddl.commitChanges();
            if (ddl.raisePropertyChanged) ddl.raisePropertyChanged('selectedItem');
            if (ddl.postback) ddl.postback();
            const selected = ddl.get_selectedItem ? ddl.get_selectedItem() : null;
            return {
              ok: true,
              value: selected && selected.get_value ? selected.get_value() : '',
              text: selected && selected.get_text ? selected.get_text() : ''
            };
          } catch (e) {
            return { ok: false, reason: String(e), value: '', text: '' };
          }
        }
        """,
        {"ddlId": ddl_id, "expectedValue": value},
    )


def wait_department_ready_after_company(page: Page, timeout_ms: int = 350) -> bool:
    started = time.time()
    timeout_s = max(0.05, timeout_ms / 1000.0)
    while (time.time() - started) < timeout_s:
        ready = page.evaluate(
            """
            () => {
              try {
                const ddl = window.$find && window.$find('ctl00_usercontrol2_ddlDepartamento');
                const root = document.getElementById('ctl00_usercontrol2_ddlDepartamento');
                if (!ddl || !root) return false;
                const enabled = ddl.get_enabled ? !!ddl.get_enabled() : true;
                const disabledByClass = root.classList.contains('rddlDisabled');
                const disabledByAttr = root.getAttribute('disabled') !== null;
                return enabled && !disabledByClass && !disabledByAttr;
              } catch {
                return false;
              }
            }
            """
        )
        if bool(ready):
            return True
        wait(page, 70, critical=True)
    return False


def ensure_single_ddl_selected(
    page: Page,
    ddl_id: str,
    expected_value: str,
    expected_text: str,
    label: str,
    max_attempts: int = 4,
) -> dict[str, Any]:
    expected_text_norm = expected_text.lower()
    state = read_ddl_state(page, ddl_id)
    text_now = str(state.get("selectedText", "")).lower()
    value_now = str(state.get("selectedValue", ""))
    if value_now == expected_value or expected_text_norm in text_now:
        log(f'select-{label}: ya seleccionado value={value_now or "(vacio)"} "{state.get("selectedText", "")}"')
        return state

    for i in range(max_attempts):
        _ = select_ddl_via_telerik(page, ddl_id, expected_value)
        state = read_ddl_state(page, ddl_id)
        text_now = str(state.get("selectedText", "")).lower()
        value_now = str(state.get("selectedValue", ""))
        ok = value_now == expected_value or expected_text_norm in text_now
        log(f'Intento select-{label} {i + 1}: value={value_now or "(vacio)"} "{state.get("selectedText", "")}"')
        if ok:
            return state
        activate_ddl(page, ddl_id, settle_ms=65)
        wait(page, 30, critical=True)

    return read_ddl_state(page, ddl_id)


def force_login_prereqs(page: Page) -> None:
    page.evaluate(
        """
        () => {
          try { if (typeof window.Page_ClientValidate === 'function') window.Page_ClientValidate(); } catch {}
          try {
            const fire = (id) => {
              const ddl = window.$find && window.$find(id);
              if (!ddl) return;
              if (ddl.raisePropertyChanged) ddl.raisePropertyChanged('selectedItem');
              if (ddl.postback) ddl.postback();
            };
            fire('ctl00_usercontrol2_ddlCompany');
            fire('ctl00_usercontrol2_ddlDepartamento');
          } catch {}
          const btn = document.getElementById('ctl00_usercontrol2_T9500_Login_input');
          if (btn) {
            btn.disabled = false;
            btn.removeAttribute('disabled');
          }
        }
        """
    )


def click_login_once(page: Page) -> None:
    page.mouse.click(40, 40)
    wait(page, 120)
    force_login_prereqs(page)
    wait(page, 120)
    try:
        page.locator("#ctl00_usercontrol2_T9500_Login_input").first.click(timeout=7000, force=True)
    except Exception:
        page.evaluate(
            """
            () => {
              const btn = document.getElementById('ctl00_usercontrol2_T9500_Login_input');
              if (btn) {
                btn.disabled = false;
                btn.removeAttribute('disabled');
                btn.click();
              }
              if (typeof window.__doPostBack === 'function') {
                try { window.__doPostBack('ctl00$usercontrol2$T9500$Login', ''); } catch {}
              }
            }
            """
        )


def do_login_flow(page: Page) -> None:
    log("Paso 2: escribir usuario y contrasena")
    page.locator("#ctl00_usercontrol2_txt9001").fill(USER)
    page.locator("#ctl00_usercontrol2_txt9004").fill(PASSWORD)
    page.mouse.click(140, 140)
    wait(page, 420)

    log("Paso 3-4: seleccionar Empresa y Departamento (con verificacion fuerte)")
    activate_ddl(page, "ctl00_usercontrol2_ddlCompany", settle_ms=90)
    wait(page, 120, critical=True)
    company = ensure_single_ddl_selected(
        page,
        ddl_id="ctl00_usercontrol2_ddlCompany",
        expected_value="PSV",
        expected_text="MEDICAL PRACTICE",
        label="empresa",
        max_attempts=4,
    )
    company_ok = company.get("selectedValue") == "PSV" or "medical practice" in str(company.get("selectedText", "")).lower()
    if not company_ok:
        raise RuntimeError("No se pudo seleccionar Empresa (MEDICAL PRACTICE).")

    dept_ready = wait_department_ready_after_company(page, timeout_ms=120)
    log(f"Departamento ready after empresa: {1 if dept_ready else 0}")
    department = ensure_single_ddl_selected(
        page,
        ddl_id="ctl00_usercontrol2_ddlDepartamento",
        expected_value="CEX",
        expected_text="PRACTICA MEDICA",
        label="departamento",
        max_attempts=3,
    )
    dept_text = str(department.get("selectedText", "")).lower()
    dept_ok = (
        department.get("selectedValue") == "CEX"
        or "practica medica" in dept_text
        or "práctica médica" in dept_text
    )
    if not dept_ok:
        raise RuntimeError("No se pudo seleccionar Departamento (PRACTICA MEDICA).")

    log(f'Empresa final => value:{company.get("selectedValue", "(vacio)")} text:{company.get("selectedText", "(vacio)")}')
    log(f'Departamento final => value:{department.get("selectedValue", "(vacio)")} text:{department.get("selectedText", "(vacio)")}')

    btn_disabled = page.evaluate(
        """
        () => {
          const btn = document.getElementById('ctl00_usercontrol2_T9500_Login_input');
          return btn ? !!btn.disabled : true;
        }
        """
    )
    log(f"Boton login disabled: {btn_disabled}")

    log("Paso 5: iniciar sesion")
    if btn_disabled:
        force_login_prereqs(page)
        wait(page, 150, critical=True)
    for i in range(3):
        click_login_once(page)
        try:
            page.wait_for_url(re.compile(r"/Default"), timeout=20000)
            return
        except PlaywrightTimeoutError:
            log(f"Reintento click login {i + 1}/3 (sigue en Login)")
            wait(page, 350, critical=True)
    raise RuntimeError("No se pudo completar login tras 3 intentos.")


def run_node_full_flow() -> int:
    node_script = Path(__file__).with_name("login_medico_auto.js")
    if not node_script.exists():
        raise RuntimeError(f"No existe {node_script}")
    if shutil.which("node") is None:
        raise RuntimeError("No se encontro 'node' en PATH.")

    env = os.environ.copy()
    env.setdefault("ONLY_SELECT_CALENDAR_FIELD", "1")
    env.setdefault("BOT_MAIN_MODE", os.getenv("BOT_MAIN_MODE", "1"))
    env.setdefault("CANCEL_MAX_APPOINTMENTS", os.getenv("CANCEL_MAX_APPOINTMENTS", "1"))
    env.setdefault("CANCEL_ACTION_WAIT_TIMEOUT_MS", os.getenv("CANCEL_ACTION_WAIT_TIMEOUT_MS", "10000"))
    env.setdefault("CANCEL_ACTION_WAIT_INTERVAL_MS", os.getenv("CANCEL_ACTION_WAIT_INTERVAL_MS", "380"))
    env.setdefault("AUTO_CREATE_APPOINTMENT", "1")
    env.setdefault("AUTO_SAVE_APPOINTMENT", "1")
    env.setdefault("AUTO_OPEN_MODULE_AFTER_SAVE", os.getenv("AUTO_OPEN_MODULE_AFTER_SAVE", "1"))
    env.setdefault("POST_SAVE_REQUIRE_ASSIGNED_MODAL", os.getenv("POST_SAVE_REQUIRE_ASSIGNED_MODAL", "1"))
    env.setdefault("POST_SAVE_ALLOW_GENERIC_MODULO_FALLBACK", os.getenv("POST_SAVE_ALLOW_GENERIC_MODULO_FALLBACK", "0"))
    env.setdefault("POST_SAVE_MAX_RETRIES", os.getenv("POST_SAVE_MAX_RETRIES", "4"))
    env.setdefault("POST_SAVE_RETRY_INTERVAL_MS", os.getenv("POST_SAVE_RETRY_INTERVAL_MS", "90"))
    env.setdefault("POST_SAVE_MODAL_CLICK_LOOP_MAX", os.getenv("POST_SAVE_MODAL_CLICK_LOOP_MAX", "3"))
    env.setdefault("POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS", os.getenv("POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS", "70"))
    env.setdefault("CLICK_SEARCH_AFTER_KEY", "0")
    env.setdefault("ENABLE_ENTER_FALLBACK", os.getenv("ENABLE_ENTER_FALLBACK", "0"))
    env.setdefault("STRICT_NUEVA_CITA_MODAL", os.getenv("STRICT_NUEVA_CITA_MODAL", "1"))
    env.setdefault("CATALOG_LOOP_MAX", os.getenv("CATALOG_LOOP_MAX", "2"))
    env.setdefault("MAX_KEY_ATTEMPTS", os.getenv("MAX_KEY_ATTEMPTS", "0"))
    env.setdefault("PRIORITIZE_RECENT_KEYS", os.getenv("PRIORITIZE_RECENT_KEYS", "0"))
    env.setdefault("KEY_SELECTION_MODE", os.getenv("KEY_SELECTION_MODE", "random"))
    env.setdefault("KEY_RANDOM_SEED", os.getenv("KEY_RANDOM_SEED", ""))
    env.setdefault("KEY_SETTLE_MS", os.getenv("KEY_SETTLE_MS", "900"))
    env.setdefault("KEY_RESOLUTION_TIMEOUT_MS", os.getenv("KEY_RESOLUTION_TIMEOUT_MS", "4200"))
    env.setdefault("COMMENT_CLICK_RETRIES", os.getenv("COMMENT_CLICK_RETRIES", "3"))
    env.setdefault("REQUIRE_SAVE_ALERT", os.getenv("REQUIRE_SAVE_ALERT", "1"))
    env.setdefault("MODULE_LOAD_POLL_TIMEOUT_MS", os.getenv("MODULE_LOAD_POLL_TIMEOUT_MS", "90000"))
    env.setdefault("MODULE_LOAD_POLL_INTERVAL_MS", os.getenv("MODULE_LOAD_POLL_INTERVAL_MS", "350"))
    env.setdefault("AUTO_OPEN_NOTA_MEDICA_AFTER_MODULE", os.getenv("AUTO_OPEN_NOTA_MEDICA_AFTER_MODULE", "1"))
    env.setdefault("NOTA_MEDICA_DELAY_MS", os.getenv("NOTA_MEDICA_DELAY_MS", "2000"))
    env.setdefault("NOTA_MEDICA_CLICK_TIMEOUT_MS", os.getenv("NOTA_MEDICA_CLICK_TIMEOUT_MS", "20000"))
    env.setdefault("AUTO_GENERAR_PLAN_TRATAMIENTO", os.getenv("AUTO_GENERAR_PLAN_TRATAMIENTO", "1"))
    env.setdefault("PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS", os.getenv("PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS", "12000"))
    env.setdefault("AUTO_FILL_NOTA_MEDICA_FIELDS", os.getenv("AUTO_FILL_NOTA_MEDICA_FIELDS", "1"))
    env.setdefault("AUTO_CLICK_GENERAR_IA_NOTA_MEDICA", os.getenv("AUTO_CLICK_GENERAR_IA_NOTA_MEDICA", "1"))
    env.setdefault("NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS", os.getenv("NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS", "18000"))
    env.setdefault("NOTA_MEDICA_FIELDS_FILL_RETRY_MS", os.getenv("NOTA_MEDICA_FIELDS_FILL_RETRY_MS", "320"))
    env.setdefault("REVIEW_HOLD_MS", os.getenv("REVIEW_HOLD_MS", "1800000"))
    env.setdefault("ERROR_REVIEW_HOLD_MS", os.getenv("ERROR_REVIEW_HOLD_MS", "1800000"))
    env.setdefault("KEY_EXHAUST_REVIEW_HOLD_MS", os.getenv("KEY_EXHAUST_REVIEW_HOLD_MS", "0"))
    env.setdefault("STRICT_PREFERRED_SLOT", "0")
    env.setdefault("PATIENT_KEYS_FILE", str(PATIENT_KEYS_FILE))
    env.setdefault("APPOINTMENT_MEMORY_ENABLED", os.getenv("APPOINTMENT_MEMORY_ENABLED", "1"))
    env.setdefault("APPOINTMENT_MEMORY_FILE", os.getenv("APPOINTMENT_MEMORY_FILE", str(Path(__file__).with_name("appointment_memory_tmp.json"))))
    env.setdefault("APPOINTMENT_MEMORY_TTL_HOURS", os.getenv("APPOINTMENT_MEMORY_TTL_HOURS", "72"))
    env.setdefault("RESTART_FROM_LOGIN_ON_BUG", os.getenv("NODE_RESTART_FROM_LOGIN_ON_BUG", "1"))
    env.setdefault("FULL_FLOW_RETRIES", os.getenv("NODE_FULL_FLOW_RETRIES", "1"))
    # Perfil un poco más conservador para evitar condition-races que abren catálogo.
    env.setdefault("TIMEOUT_SCALE", os.getenv("NODE_TIMEOUT_SCALE", "0.92"))
    env.setdefault("MIN_WAIT_MS", str(MIN_WAIT_MS))
    env.setdefault("SLOW_MO_MS", os.getenv("NODE_SLOW_MO_MS", "140"))
    env.setdefault("LIVE_LOG_PATH", str(Path.cwd() / "login_medico_live.log"))
    cmd = ["node", str(node_script)]

    for attempt in range(1, NODE_FLOW_RETRIES + 1):
        log(
            f"Delegando flujo completo al bot Node estable (login + agenda + cita). "
            f"Intento {attempt}/{NODE_FLOW_RETRIES}"
        )
        proc = subprocess.run(
            cmd,
            cwd=str(node_script.parent),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        out = (proc.stdout or b"").decode("utf-8", errors="replace")
        err = (proc.stderr or b"").decode("utf-8", errors="replace")
        if out:
            print(out, end="")
        if err:
            print(err, end="")

        rc = int(proc.returncode)
        if rc == 0:
            return 0

        joined = f"{out}\n{err}".lower()
        transient = (
            "no se pudo recuperar \"nueva cita\" en la misma casilla" in joined
            or "catalogo de pacientes" in joined
            or "nueva_cita_modal_not_found" in joined
        )
        if attempt < NODE_FLOW_RETRIES and transient:
            log("Fallo transitorio detectado (modal/catálogo). Reintentando flujo completo...")
            time.sleep(1.5)
            continue
        return rc
    return 1


def run_python_login_only() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=SLOW_MO_MS)
        context = browser.new_context()
        page = context.new_page()

        try:
            log("Paso 1: abrir pagina inicial")
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
            page.bring_to_front()
            page.mouse.click(200, 200)
            wait(page, 600)

            on_login_page = page.locator("#ctl00_usercontrol2_txt9001").count() > 0
            if on_login_page:
                do_login_flow(page)
            else:
                log("Sesion ya iniciada o pagina /Default cargada. Saltando login.")

            page.wait_for_url(re.compile(r"/Default"), timeout=30000)
            wait(page, 1000)
            log("LOGIN_OK")
            log(f"Navegador abierto {int(HOLD_AFTER_LOGIN_MS / 1000)}s para uso/validacion...")
            page.wait_for_timeout(HOLD_AFTER_LOGIN_MS)
            return 0
        except PlaywrightTimeoutError:
            page.screenshot(path="login_medico_error.png", full_page=True)
            log("LOGIN_TIMEOUT - revisa login_medico_error.png")
            return 2
        except Exception as exc:
            page.screenshot(path="login_medico_error.png", full_page=True)
            log(f"LOGIN_ERROR: {exc}")
            return 3


def main() -> int:
    try:
        LIVE_LOG_PATH.write_text("", encoding="utf-8")
    except Exception:
        pass
    log(f"LIVE_LOG_PATH={LIVE_LOG_PATH}")
    log(f"PERF_CONFIG TIMEOUT_SCALE={TIMEOUT_SCALE} MIN_WAIT_MS={MIN_WAIT_MS} SLOW_MO_MS={SLOW_MO_MS}")
    log(
        "FLOW_GUARDS "
        f'BOT_MAIN_MODE={os.getenv("BOT_MAIN_MODE", "1")} '
        f'CANCEL_MAX_APPOINTMENTS={os.getenv("CANCEL_MAX_APPOINTMENTS", "1")} '
        f'AUTO_OPEN_MODULE_AFTER_SAVE={os.getenv("AUTO_OPEN_MODULE_AFTER_SAVE", "1")} '
        f'STRICT_NUEVA_CITA_MODAL={os.getenv("STRICT_NUEVA_CITA_MODAL", "1")} '
        f'CLICK_SEARCH_AFTER_KEY={os.getenv("CLICK_SEARCH_AFTER_KEY", "0")} '
        f'ENABLE_ENTER_FALLBACK={os.getenv("ENABLE_ENTER_FALLBACK", "0")} '
        f'CATALOG_LOOP_MAX={os.getenv("CATALOG_LOOP_MAX", "2")} '
        f'MAX_KEY_ATTEMPTS={os.getenv("MAX_KEY_ATTEMPTS", "0")}'
        f' PRIORITIZE_RECENT_KEYS={os.getenv("PRIORITIZE_RECENT_KEYS", "0")}'
        f' KEY_SELECTION_MODE={os.getenv("KEY_SELECTION_MODE", "random")}'
        f' KEY_RANDOM_SEED={os.getenv("KEY_RANDOM_SEED", "-") or "-"}'
        f' KEY_SETTLE_MS={os.getenv("KEY_SETTLE_MS", "900")}'
        f' KEY_RESOLUTION_TIMEOUT_MS={os.getenv("KEY_RESOLUTION_TIMEOUT_MS", "4200")}'
        f' COMMENT_CLICK_RETRIES={os.getenv("COMMENT_CLICK_RETRIES", "3")}'
        f' REQUIRE_SAVE_ALERT={os.getenv("REQUIRE_SAVE_ALERT", "1")}'
        f' AUTO_FILL_NOTA_MEDICA_FIELDS={os.getenv("AUTO_FILL_NOTA_MEDICA_FIELDS", "1")}'
        f' AUTO_CLICK_GENERAR_IA_NOTA_MEDICA={os.getenv("AUTO_CLICK_GENERAR_IA_NOTA_MEDICA", "1")}'
        f' NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS={os.getenv("NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS", "18000")}'
        f' NOTA_MEDICA_FIELDS_FILL_RETRY_MS={os.getenv("NOTA_MEDICA_FIELDS_FILL_RETRY_MS", "320")}'
        f' POST_SAVE_MODAL_CLICK_LOOP_MAX={os.getenv("POST_SAVE_MODAL_CLICK_LOOP_MAX", "3")}'
        f' POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS={os.getenv("POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS", "70")}'
        f' MODULE_LOAD_POLL_TIMEOUT_MS={os.getenv("MODULE_LOAD_POLL_TIMEOUT_MS", "90000")}'
        f' MODULE_LOAD_POLL_INTERVAL_MS={os.getenv("MODULE_LOAD_POLL_INTERVAL_MS", "350")}'
        f' AUTO_GENERAR_PLAN_TRATAMIENTO={os.getenv("AUTO_GENERAR_PLAN_TRATAMIENTO", "1")}'
        f' PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS={os.getenv("PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS", "12000")}'
        f' REVIEW_HOLD_MS={os.getenv("REVIEW_HOLD_MS", "1800000")}'
        f' ERROR_REVIEW_HOLD_MS={os.getenv("ERROR_REVIEW_HOLD_MS", "1800000")}'
        f' KEY_EXHAUST_REVIEW_HOLD_MS={os.getenv("KEY_EXHAUST_REVIEW_HOLD_MS", "0")}'
    )
    log(f"PATIENT_KEYS_SOURCE={PATIENT_KEYS_SOURCE} count={len(PATIENT_KEYS)}")
    log(
        "APPOINTMENT_MEMORY "
        f'enabled={os.getenv("APPOINTMENT_MEMORY_ENABLED", "1")} '
        f'file={os.getenv("APPOINTMENT_MEMORY_FILE", str(Path(__file__).with_name("appointment_memory_tmp.json")))} '
        f'ttl_h={os.getenv("APPOINTMENT_MEMORY_TTL_HOURS", "72")}'
    )

    if ONLY_LOGIN:
        return run_python_login_only()

    # Flujo completo: se apoya en el bot Node ya validado para cita.
    return run_node_full_flow()


if __name__ == "__main__":
    raise SystemExit(main())


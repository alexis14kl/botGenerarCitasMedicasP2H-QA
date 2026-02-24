const fs = require('node:fs');
const path = require('node:path');
const readline = require('node:readline/promises');
const { stdin: processStdin, stdout: processStdout } = require('node:process');
const { chromium } = require('playwright');

const URL = 'https://mp-stg.telesalud.gob.sv/';
const START_URL = process.env.START_URL || URL;
const USER = process.env.BOT_USER || 'MEDICO09';
const PASSWORD = process.env.BOT_PASSWORD || 'ISSS202';
const BOT_MAIN_MODE_ENV = (process.env.BOT_MAIN_MODE || '').toString().trim();
let BOT_MAIN_MODE = BOT_MAIN_MODE_ENV === '2' ? '2' : '1'; // 1=generar ordenes, 2=cancelar cita
const ONLY_LOGIN = process.env.ONLY_LOGIN === '1';
const ONLY_SELECT_CALENDAR_FIELD = process.env.ONLY_SELECT_CALENDAR_FIELD !== '0';
const COOP_MODE = process.env.COOP_MODE === '1';
const AUTO_CREATE_APPOINTMENT = process.env.AUTO_CREATE_APPOINTMENT !== '0';
const AUTO_SAVE_APPOINTMENT = process.env.AUTO_SAVE_APPOINTMENT !== '0';
const AUTO_OPEN_MODULE_AFTER_SAVE = process.env.AUTO_OPEN_MODULE_AFTER_SAVE !== '0';
const CLICK_SEARCH_AFTER_KEY = process.env.CLICK_SEARCH_AFTER_KEY === '1';
const ENABLE_ENTER_FALLBACK = process.env.ENABLE_ENTER_FALLBACK === '1';
const STRICT_PREFERRED_SLOT = process.env.STRICT_PREFERRED_SLOT === '1';
const STRICT_NUEVA_CITA_MODAL = process.env.STRICT_NUEVA_CITA_MODAL !== '0';
const ALLOW_SAVE_ON_UNCONFIRMED_KEY = process.env.ALLOW_SAVE_ON_UNCONFIRMED_KEY === '1';
const RESTART_FROM_LOGIN_ON_BUG = process.env.RESTART_FROM_LOGIN_ON_BUG !== '0';
const FULL_FLOW_RETRIES = (() => {
  const n = Number(process.env.FULL_FLOW_RETRIES || '1');
  if (!Number.isFinite(n)) return 1;
  return Math.min(5, Math.max(1, Math.round(n)));
})();
const TIMEOUT_SCALE = (() => {
  const n = Number(process.env.TIMEOUT_SCALE || '0.85');
  if (!Number.isFinite(n)) return 0.85;
  return Math.min(1.2, Math.max(0.25, n));
})();
const MIN_WAIT_MS = (() => {
  const n = Number(process.env.MIN_WAIT_MS || '60');
  if (!Number.isFinite(n)) return 60;
  return Math.min(250, Math.max(0, Math.round(n)));
})();
const SLOW_MO_MS = (() => {
  const n = Number(process.env.SLOW_MO_MS || '130');
  if (!Number.isFinite(n)) return 130;
  return Math.min(500, Math.max(0, Math.round(n)));
})();
const LIVE_LOG_PATH = process.env.LIVE_LOG_PATH || path.join(process.cwd(), 'login_medico_live.log');
const PATIENT_KEYS_FILE = process.env.PATIENT_KEYS_FILE || path.join(__dirname, 'patient_keys.txt');
const DEFAULT_PATIENT_KEYS = [
  '00955873-3',
  '06169373-5',
  '05608981-6',
  '06416857-7',
  '05400186-2',
  'B04676303',
  'B02661296',
  'B01700785',
  'B00838396',
  'B00491489'
];

function normalizePatientKey(raw) {
  return (raw || '')
    .toString()
    .toUpperCase()
    .trim()
    .replace(/['"]/g, '')
    .replace(/[^A-Z0-9-]/g, '');
}

function isLikelyPatientKey(key) {
  // Acepta formatos tipo 01234567-8, B01234567, 1001856, etc.
  return /^[A-Z]?[0-9][0-9-]{2,}$/.test(key);
}

function parsePatientKeys(rawText) {
  const out = [];
  const seen = new Set();
  const chunks = (rawText || '').split(/[\r\n,\s;]+/g);
  for (const chunk of chunks) {
    const normalized = normalizePatientKey(chunk);
    if (!normalized || !isLikelyPatientKey(normalized) || seen.has(normalized)) continue;
    seen.add(normalized);
    out.push(normalized);
  }
  return out;
}

function loadPatientKeys() {
  try {
    if (fs.existsSync(PATIENT_KEYS_FILE)) {
      const fileText = fs.readFileSync(PATIENT_KEYS_FILE, 'utf8');
      const fileKeys = parsePatientKeys(fileText);
      if (fileKeys.length > 0) {
        return { keys: fileKeys, source: `file:${PATIENT_KEYS_FILE}` };
      }
    }
  } catch {}

  const envKeys = parsePatientKeys(process.env.PATIENT_KEYS || '');
  if (envKeys.length > 0) {
    return { keys: envKeys, source: 'env:PATIENT_KEYS' };
  }

  return { keys: DEFAULT_PATIENT_KEYS, source: 'fallback:default_keys' };
}

const PATIENT_KEYS_LOAD = loadPatientKeys();
const PATIENT_KEYS = PATIENT_KEYS_LOAD.keys;
const PATIENT_KEYS_SOURCE = PATIENT_KEYS_LOAD.source;
const PRIORITIZE_RECENT_KEYS = process.env.PRIORITIZE_RECENT_KEYS === '1';
const KEY_SELECTION_MODE = (() => {
  const raw = (process.env.KEY_SELECTION_MODE || 'random').toString().trim().toLowerCase();
  const allowed = new Set(['sequential', 'random', 'recent_then_random']);
  return allowed.has(raw) ? raw : 'sequential';
})();
const KEY_RANDOM_SEED = (process.env.KEY_RANDOM_SEED || '').toString().trim();
const APPOINTMENT_MEMORY_ENABLED = process.env.APPOINTMENT_MEMORY_ENABLED !== '0';
const APPOINTMENT_MEMORY_FILE = process.env.APPOINTMENT_MEMORY_FILE || path.join(__dirname, 'appointment_memory_tmp.json');
const APPOINTMENT_MEMORY_TTL_HOURS = (() => {
  const n = Number(process.env.APPOINTMENT_MEMORY_TTL_HOURS || '72');
  if (!Number.isFinite(n)) return 72;
  return Math.min(24 * 30, Math.max(1, Math.round(n)));
})();
const APPOINTMENT_MEMORY_MAX_ITEMS = (() => {
  const n = Number(process.env.APPOINTMENT_MEMORY_MAX_ITEMS || '800');
  if (!Number.isFinite(n)) return 800;
  return Math.min(5000, Math.max(50, Math.round(n)));
})();
const KEY_HEALTH_ENABLED = process.env.KEY_HEALTH_ENABLED !== '0';
const KEY_HEALTH_FILE = process.env.KEY_HEALTH_FILE || path.join(__dirname, 'patient_key_health_tmp.json');
const KEY_HEALTH_TTL_HOURS = (() => {
  const n = Number(process.env.KEY_HEALTH_TTL_HOURS || '168');
  if (!Number.isFinite(n)) return 168;
  return Math.min(24 * 60, Math.max(6, Math.round(n)));
})();
const KEY_HEALTH_MAX_ITEMS = (() => {
  const n = Number(process.env.KEY_HEALTH_MAX_ITEMS || '6000');
  if (!Number.isFinite(n)) return 6000;
  return Math.min(20000, Math.max(200, Math.round(n)));
})();
const KEY_HARD_BLOCK_THRESHOLD = (() => {
  const n = Number(process.env.KEY_HARD_BLOCK_THRESHOLD || '2');
  if (!Number.isFinite(n)) return 2;
  return Math.min(8, Math.max(1, Math.round(n)));
})();
const CATALOG_LOOP_MAX = (() => {
  const n = Number(process.env.CATALOG_LOOP_MAX || '2');
  if (!Number.isFinite(n)) return 2;
  return Math.min(10, Math.max(1, Math.round(n)));
})();
const MAX_KEY_ATTEMPTS = (() => {
  const raw = process.env.MAX_KEY_ATTEMPTS;
  const autoCount = Math.min(500, Math.max(1, PATIENT_KEYS.length || 1));
  if (raw === undefined || raw === null || String(raw).trim() === '') return autoCount;
  const n = Number(raw);
  // 0 o negativo => modo automático: usa el total cargado del txt.
  if (!Number.isFinite(n) || n <= 0) return autoCount;
  return Math.min(500, Math.max(1, Math.round(n)));
})();
const KEY_SETTLE_MS = (() => {
  const n = Number(process.env.KEY_SETTLE_MS || '900');
  if (!Number.isFinite(n)) return 900;
  return Math.min(3000, Math.max(150, Math.round(n)));
})();
const KEY_RESOLUTION_TIMEOUT_MS = (() => {
  const n = Number(process.env.KEY_RESOLUTION_TIMEOUT_MS || '4200');
  if (!Number.isFinite(n)) return 4200;
  return Math.min(12000, Math.max(800, Math.round(n)));
})();
const COMMENT_CLICK_RETRIES = (() => {
  const n = Number(process.env.COMMENT_CLICK_RETRIES || '3');
  if (!Number.isFinite(n)) return 3;
  return Math.min(8, Math.max(1, Math.round(n)));
})();
const COMMENT_TEXT = (process.env.COMMENT_TEXT || 'TEST').toString().trim() || 'TEST';
const CANCEL_MAX_APPOINTMENTS = (() => {
  const n = Number(process.env.CANCEL_MAX_APPOINTMENTS || '1');
  if (!Number.isFinite(n)) return 1;
  return Math.min(25, Math.max(1, Math.round(n)));
})();
const CANCEL_SEARCH_MAX_WEEKS = (() => {
  const n = Number(process.env.CANCEL_SEARCH_MAX_WEEKS || '2');
  if (!Number.isFinite(n)) return 2;
  return Math.min(8, Math.max(1, Math.round(n)));
})();
const MODE2_SLOT_MIN_DAY_OFFSET = (() => {
  const n = Number(process.env.MODE2_SLOT_MIN_DAY_OFFSET || '0');
  if (!Number.isFinite(n)) return 1;
  return Math.min(14, Math.max(0, Math.round(n)));
})();
const MODE2_SKIP_SUNDAYS = process.env.MODE2_SKIP_SUNDAYS !== '0';
const MODE2_MAX_SEARCH_WEEKS = (() => {
  const n = Number(process.env.MODE2_MAX_SEARCH_WEEKS || '1');
  if (!Number.isFinite(n)) return 1;
  return Math.min(4, Math.max(1, Math.round(n)));
})();
const MODE2_AUTO_FILTER = process.env.MODE2_AUTO_FILTER === '1';
const MODE2_MAX_SLOT_CANDIDATES = (() => {
  const n = Number(process.env.MODE2_MAX_SLOT_CANDIDATES || '16');
  if (!Number.isFinite(n)) return 16;
  return Math.min(120, Math.max(6, Math.round(n)));
})();
const MODE2_MAX_SLOT_CANDIDATES_PER_DAY = (() => {
  const n = Number(process.env.MODE2_MAX_SLOT_CANDIDATES_PER_DAY || '4');
  if (!Number.isFinite(n)) return 4;
  return Math.min(40, Math.max(2, Math.round(n)));
})();
const MODE2_SCAN_START_MINUTES = (() => {
  const n = Number(process.env.MODE2_SCAN_START_MINUTES || '0');
  if (!Number.isFinite(n)) return 0;
  return Math.min(1439, Math.max(0, Math.round(n)));
})();
const MODE2_SCAN_END_MINUTES = (() => {
  const n = Number(process.env.MODE2_SCAN_END_MINUTES || '1080');
  if (!Number.isFinite(n)) return 1080;
  return Math.min(1439, Math.max(0, Math.round(n)));
})();
const CANCEL_SLOT_MODULO_MAX_RETRIES = (() => {
  const n = Number(process.env.CANCEL_SLOT_MODULO_MAX_RETRIES || '4');
  if (!Number.isFinite(n)) return 4;
  return Math.min(8, Math.max(1, Math.round(n)));
})();
const CANCEL_SLOT_REFOCUS_LOOP_MAX = (() => {
  const n = Number(process.env.CANCEL_SLOT_REFOCUS_LOOP_MAX || '4');
  if (!Number.isFinite(n)) return 4;
  return Math.min(8, Math.max(1, Math.round(n)));
})();
const CANCEL_SLOT_REFOCUS_RETRY_MS = (() => {
  const n = Number(process.env.CANCEL_SLOT_REFOCUS_RETRY_MS || '80');
  if (!Number.isFinite(n)) return 80;
  return Math.min(600, Math.max(20, Math.round(n)));
})();
const CANCEL_ACTION_WAIT_TIMEOUT_MS = (() => {
  const n = Number(process.env.CANCEL_ACTION_WAIT_TIMEOUT_MS || '10000');
  if (!Number.isFinite(n)) return 10000;
  return Math.min(45000, Math.max(1500, Math.round(n)));
})();
const CANCEL_ACTION_WAIT_INTERVAL_MS = (() => {
  const n = Number(process.env.CANCEL_ACTION_WAIT_INTERVAL_MS || '380');
  if (!Number.isFinite(n)) return 380;
  return Math.min(2000, Math.max(120, Math.round(n)));
})();
const DEBUG_NUEVA_CITA_CONTROLS = process.env.DEBUG_NUEVA_CITA_CONTROLS === '1';
const REQUIRE_SAVE_ALERT = process.env.REQUIRE_SAVE_ALERT !== '0';
const REVIEW_HOLD_MS = (() => {
  const n = Number(process.env.REVIEW_HOLD_MS || '1800000');
  if (!Number.isFinite(n)) return 1800000;
  if (n <= 0) return 0;
  return Math.min(24 * 60 * 60 * 1000, Math.max(10000, Math.round(n)));
})();
const ERROR_REVIEW_HOLD_MS = (() => {
  const n = Number(process.env.ERROR_REVIEW_HOLD_MS || '1800000');
  if (!Number.isFinite(n)) return 1800000;
  if (n <= 0) return 0;
  return Math.min(24 * 60 * 60 * 1000, Math.max(10000, Math.round(n)));
})();
const KEY_EXHAUST_REVIEW_HOLD_MS = (() => {
  const n = Number(process.env.KEY_EXHAUST_REVIEW_HOLD_MS || '0');
  if (!Number.isFinite(n)) return 0;
  if (n <= 0) return 0;
  return Math.min(24 * 60 * 60 * 1000, Math.max(10000, Math.round(n)));
})();
const MODULE_LOAD_POLL_TIMEOUT_MS = (() => {
  const n = Number(process.env.MODULE_LOAD_POLL_TIMEOUT_MS || '90000');
  if (!Number.isFinite(n)) return 90000;
  return Math.min(3 * 60 * 1000, Math.max(5000, Math.round(n)));
})();
const MODULE_LOAD_POLL_INTERVAL_MS = (() => {
  const n = Number(process.env.MODULE_LOAD_POLL_INTERVAL_MS || '350');
  if (!Number.isFinite(n)) return 350;
  return Math.min(2000, Math.max(120, Math.round(n)));
})();
const AUTO_OPEN_NOTA_MEDICA_AFTER_MODULE = process.env.AUTO_OPEN_NOTA_MEDICA_AFTER_MODULE !== '0';
const RELOAD_BEFORE_NOTA_MEDICA = process.env.RELOAD_BEFORE_NOTA_MEDICA !== '0';
const RELOAD_BEFORE_NOTA_MEDICA_TIMEOUT_MS = (() => {
  const n = Number(process.env.RELOAD_BEFORE_NOTA_MEDICA_TIMEOUT_MS || '20000');
  if (!Number.isFinite(n)) return 20000;
  return Math.min(60000, Math.max(5000, Math.round(n)));
})();
const RELOAD_BEFORE_NOTA_MEDICA_POLL_MS = (() => {
  const n = Number(process.env.RELOAD_BEFORE_NOTA_MEDICA_POLL_MS || '320');
  if (!Number.isFinite(n)) return 320;
  return Math.min(2000, Math.max(120, Math.round(n)));
})();
const NOTA_MEDICA_DELAY_MS = (() => {
  const n = Number(process.env.NOTA_MEDICA_DELAY_MS || '2000');
  if (!Number.isFinite(n)) return 2000;
  return Math.min(15000, Math.max(0, Math.round(n)));
})();
const NOTA_MEDICA_CLICK_TIMEOUT_MS = (() => {
  const n = Number(process.env.NOTA_MEDICA_CLICK_TIMEOUT_MS || '20000');
  if (!Number.isFinite(n)) return 20000;
  return Math.min(60000, Math.max(2000, Math.round(n)));
})();
const AUTO_FILL_NOTA_MEDICA_FIELDS = process.env.AUTO_FILL_NOTA_MEDICA_FIELDS !== '0';
const AUTO_CLICK_GENERAR_IA_NOTA_MEDICA = process.env.AUTO_CLICK_GENERAR_IA_NOTA_MEDICA !== '0';
const AUTO_GENERAR_RECETA_AFTER_IA = process.env.AUTO_GENERAR_RECETA_AFTER_IA !== '0';
const RECETA_AFTER_IA_WAIT_MS = (() => {
  const n = Number(process.env.RECETA_AFTER_IA_WAIT_MS || '3500');
  if (!Number.isFinite(n)) return 3500;
  return Math.min(30000, Math.max(500, Math.round(n)));
})();
const RECETA_CLICK_TIMEOUT_MS = (() => {
  const n = Number(process.env.RECETA_CLICK_TIMEOUT_MS || '12000');
  if (!Number.isFinite(n)) return 12000;
  return Math.min(60000, Math.max(2000, Math.round(n)));
})();
const NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS = (() => {
  const n = Number(process.env.NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS || '18000');
  if (!Number.isFinite(n)) return 18000;
  return Math.min(90000, Math.max(4000, Math.round(n)));
})();
const NOTA_MEDICA_FIELDS_FILL_RETRY_MS = (() => {
  const n = Number(process.env.NOTA_MEDICA_FIELDS_FILL_RETRY_MS || '320');
  if (!Number.isFinite(n)) return 320;
  return Math.min(3000, Math.max(120, Math.round(n)));
})();
const DEFAULT_NOTA_MEDICA_TEXT =
  'El paciente refiere inicio de los síntomas hace 3 días, caracterizados por fiebre de predominio generalizado, sin localización específica, la cual aumenta en horas de la tarde y noche y disminuye parcialmente con la administración de antipiréticos; niega factores claros que la agraven, no presenta irradiación ni extensión a otras partes del cuerpo, y asocia el cuadro a malestar general, cefalea y astenia. En una escala del 1 al 10, el paciente califica la intensidad de sus síntomas como 7/10, refiriendo que interfieren con sus actividades habituales pero permiten el descanso parcial.';
const NOTA_MEDICA_FIELDS_TEXT = (process.env.NOTA_MEDICA_FIELDS_TEXT || DEFAULT_NOTA_MEDICA_TEXT).toString().trim() || DEFAULT_NOTA_MEDICA_TEXT;
const AUTO_GENERAR_PLAN_TRATAMIENTO = process.env.AUTO_GENERAR_PLAN_TRATAMIENTO !== '0';
const DEFAULT_PLAN_TRATAMIENTO_TEXT =
  'Plan breve: manejo ambulatorio, hidratacion, reposo relativo y control en 48 horas.';
const PLAN_TRATAMIENTO_TEXT =
  (process.env.PLAN_TRATAMIENTO_TEXT || DEFAULT_PLAN_TRATAMIENTO_TEXT).toString().trim() || DEFAULT_PLAN_TRATAMIENTO_TEXT;
const PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS = (() => {
  const n = Number(process.env.PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS || '12000');
  if (!Number.isFinite(n)) return 12000;
  return Math.min(60000, Math.max(2000, Math.round(n)));
})();
const POST_SAVE_REQUIRE_ASSIGNED_MODAL = process.env.POST_SAVE_REQUIRE_ASSIGNED_MODAL !== '0';
const POST_SAVE_ALLOW_GENERIC_MODULO_FALLBACK = process.env.POST_SAVE_ALLOW_GENERIC_MODULO_FALLBACK === '1';
const POST_SAVE_MAX_RETRIES = (() => {
  const n = Number(process.env.POST_SAVE_MAX_RETRIES || '4');
  if (!Number.isFinite(n)) return 4;
  return Math.min(4, Math.max(1, Math.round(n)));
})();
const POST_SAVE_RETRY_INTERVAL_MS = (() => {
  const n = Number(process.env.POST_SAVE_RETRY_INTERVAL_MS || '90');
  if (!Number.isFinite(n)) return 90;
  return Math.min(1000, Math.max(20, Math.round(n)));
})();
const POST_SAVE_MODAL_CLICK_LOOP_MAX = (() => {
  const n = Number(process.env.POST_SAVE_MODAL_CLICK_LOOP_MAX || '3');
  if (!Number.isFinite(n)) return 3;
  return Math.min(6, Math.max(1, Math.round(n)));
})();
const POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS = (() => {
  const n = Number(process.env.POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS || '70');
  if (!Number.isFinite(n)) return 70;
  return Math.min(400, Math.max(20, Math.round(n)));
})();
const POST_SAVE_STRATEGY = {
  maxAttempts: POST_SAVE_MAX_RETRIES,
  popupSettleMs: 45
};

function toLogText(value) {
  if (typeof value === 'string') return value;
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}

function installLiveLogger() {
  try {
    fs.writeFileSync(LIVE_LOG_PATH, '', 'utf8');
  } catch {}

  const baseLog = console.log.bind(console);
  const baseError = console.error.bind(console);

  const append = (level, args) => {
    const line = `[${new Date().toISOString()}] [${level}] ${args.map(toLogText).join(' ')}`;
    try {
      fs.appendFileSync(LIVE_LOG_PATH, `${line}\n`, 'utf8');
    } catch {}
  };

  console.log = (...args) => {
    baseLog(...args);
    append('INFO', args);
  };

  console.error = (...args) => {
    baseError(...args);
    append('ERROR', args);
  };

  baseLog(`LIVE_LOG_PATH=${LIVE_LOG_PATH}`);
  append('INFO', [`LIVE_LOG_PATH=${LIVE_LOG_PATH}`]);
}

installLiveLogger();

const patchedWaitTargets = new WeakSet();
const rawWaitForTimeoutMap = new WeakMap();
function scaleWaitMs(ms) {
  const n = Number(ms);
  if (!Number.isFinite(n) || n <= 0) return 0;
  // No escalar pausas largas de inspección/manual.
  if (n >= 30000) return Math.round(n);
  return Math.max(MIN_WAIT_MS, Math.round(n * TIMEOUT_SCALE));
}

function installScaledWaitForTimeout(target) {
  try {
    if (!target || typeof target.waitForTimeout !== 'function') return;
    if (patchedWaitTargets.has(target)) return;
    const original = target.waitForTimeout.bind(target);
    rawWaitForTimeoutMap.set(target, original);
    target.waitForTimeout = (ms) => original(scaleWaitMs(ms));
    patchedWaitTargets.add(target);
  } catch {}
}

function waitForTimeoutRaw(target, ms) {
  try {
    const raw = rawWaitForTimeoutMap.get(target);
    if (raw) return raw(ms);
  } catch {}
  return target.waitForTimeout(ms);
}

function sleepRaw(ms) {
  const t = Math.max(0, Number(ms) || 0);
  return new Promise((resolve) => setTimeout(resolve, t));
}

function isPageClosedSafe(page) {
  try {
    return !page || page.isClosed();
  } catch {
    return true;
  }
}

/**
 * Inyecta/actualiza un overlay flotante en la página para mostrar el estado del bot.
 * @param {import('playwright').Page} page
 * @param {'working'|'success'|'error'|'waiting'|'info'} status
 * @param {string} message - Texto a mostrar
 */
async function updateBotStatusOverlay(page, status, message) {
  if (isPageClosedSafe(page)) return;
  try {
    // Esperar a que la página esté lista antes de inyectar
    await page.waitForLoadState('domcontentloaded', { timeout: 3000 }).catch(() => {});
    const injected = await page.evaluate(({ status, message }) => {
      const root = document.body || document.documentElement;
      if (!root) return false;
      const OVERLAY_ID = '__noyecodito_bot_overlay__';
      let overlay = document.getElementById(OVERLAY_ID);
      if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = OVERLAY_ID;
        Object.assign(overlay.style, {
          position: 'fixed',
          top: '4px',
          left: '50%',
          transform: 'translateX(-50%)',
          zIndex: '2147483647',
          padding: '6px 14px',
          borderRadius: '0 0 10px 10px',
          fontFamily: "'Segoe UI', Tahoma, sans-serif",
          fontSize: '12px',
          fontWeight: '600',
          color: '#fff',
          boxShadow: '0 3px 12px rgba(0,0,0,0.35)',
          pointerEvents: 'none',
          transition: 'all 0.3s ease',
          display: 'flex',
          alignItems: 'center',
          gap: '6px',
          maxWidth: '420px',
          whiteSpace: 'nowrap',
          backdropFilter: 'blur(6px)',
          border: '1px solid rgba(255,255,255,0.25)',
          borderTop: 'none'
        });
        root.appendChild(overlay);
        // Inyectar animación CSS
        const animId = '__noyecodito_pulse_anim__';
        if (!document.getElementById(animId)) {
          const styleEl = document.createElement('style');
          styleEl.id = animId;
          styleEl.textContent = '@keyframes noyecoditoPulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.85;transform:scale(1.02)}}';
          (document.head || root).appendChild(styleEl);
        }
      }
      const configs = {
        working:  { bg: 'linear-gradient(135deg, #0088ff, #00c6ff)', emoji: '\uD83E\uDD16', pulse: true },
        success:  { bg: 'linear-gradient(135deg, #00b894, #00cec9)', emoji: '\u2705',       pulse: false },
        error:    { bg: 'linear-gradient(135deg, #e74c3c, #fd79a8)', emoji: '\u274C',       pulse: false },
        waiting:  { bg: 'linear-gradient(135deg, #fdcb6e, #e17055)', emoji: '\u23F3',       pulse: true },
        info:     { bg: 'linear-gradient(135deg, #6c5ce7, #a29bfe)', emoji: '\uD83D\uDCCB', pulse: false }
      };
      const cfg = configs[status] || configs.info;
      overlay.style.background = cfg.bg;
      overlay.style.animation = cfg.pulse ? 'noyecoditoPulse 2s ease-in-out infinite' : 'none';
      overlay.innerHTML = '<span style="font-size:18px">' + cfg.emoji + '</span><span>Noyecodito ' + message + '</span>';
      return true;
    }, { status, message });
    if (injected) {
      console.log(`BOT_OVERLAY status=${status} msg="${message}"`);
    }
  } catch (e) {
    // No bloquear el flujo por error de overlay
    console.log(`BOT_OVERLAY_ERR ${e?.message || 'unknown'}`);
  }
}

function normalizeText(value) {
  return (value || '')
    .toString()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function shouldRestartFromLogin(errorMessage) {
  if (!RESTART_FROM_LOGIN_ON_BUG) return false;
  const msg = normalizeText(errorMessage);
  const patterns = [
    'timeout',
    'no se pudo asegurar calendario',
    'no se pudo abrir modal "nueva cita"',
    'no se encontro casilla disponible',
    'nueva_cita_modal_not_found',
    'catalogo de pacientes',
    'execution context was destroyed',
    'target page, context or browser has been closed',
    'socket hang up',
    'net::err',
    'frame was detached',
    'stale'
  ];
  return patterns.some((p) => msg.includes(p));
}

function sanitizeSlotForMemory(slot) {
  if (!slot || typeof slot !== 'object') return null;
  const out = {};
  if (typeof slot.dayIso === 'string') out.dayIso = slot.dayIso.trim();
  if (Number.isFinite(slot.minutes)) out.minutes = Number(slot.minutes);
  if (Number.isInteger(slot.colIdx)) out.colIdx = Number(slot.colIdx);
  if (Number.isFinite(slot.x) && Number.isFinite(slot.y)) out.point = { x: Math.round(slot.x), y: Math.round(slot.y) };
  return Object.keys(out).length > 0 ? out : null;
}

function pruneAppointmentMemoryRecords(records, nowMs = Date.now()) {
  const ttlMs = APPOINTMENT_MEMORY_TTL_HOURS * 60 * 60 * 1000;
  const filtered = Array.isArray(records)
    ? records.filter((r) => r && Number.isFinite(r.ts) && nowMs - r.ts <= ttlMs)
    : [];
  filtered.sort((a, b) => a.ts - b.ts);
  if (filtered.length > APPOINTMENT_MEMORY_MAX_ITEMS) {
    filtered.splice(0, filtered.length - APPOINTMENT_MEMORY_MAX_ITEMS);
  }
  return filtered;
}

function loadAppointmentMemoryState() {
  if (!APPOINTMENT_MEMORY_ENABLED) {
    return { enabled: false, records: [], dirty: false };
  }
  try {
    if (!fs.existsSync(APPOINTMENT_MEMORY_FILE)) {
      return { enabled: true, records: [], dirty: false };
    }
    const raw = fs.readFileSync(APPOINTMENT_MEMORY_FILE, 'utf8');
    const parsed = JSON.parse(raw);
    const original = Array.isArray(parsed?.records) ? parsed.records : [];
    const pruned = pruneAppointmentMemoryRecords(original);
    return { enabled: true, records: pruned, dirty: pruned.length !== original.length };
  } catch {
    return { enabled: true, records: [], dirty: false };
  }
}

const APPOINTMENT_MEMORY_STATE = loadAppointmentMemoryState();

function persistAppointmentMemory() {
  if (!APPOINTMENT_MEMORY_STATE.enabled) return false;
  try {
    APPOINTMENT_MEMORY_STATE.records = pruneAppointmentMemoryRecords(APPOINTMENT_MEMORY_STATE.records);
    const payload = {
      version: 1,
      updatedAt: new Date().toISOString(),
      ttlHours: APPOINTMENT_MEMORY_TTL_HOURS,
      maxItems: APPOINTMENT_MEMORY_MAX_ITEMS,
      records: APPOINTMENT_MEMORY_STATE.records
    };
    fs.writeFileSync(APPOINTMENT_MEMORY_FILE, JSON.stringify(payload, null, 2), 'utf8');
    APPOINTMENT_MEMORY_STATE.dirty = false;
    return true;
  } catch {
    return false;
  }
}

function rememberCreatedAppointment({ key, number, slot, status = 'success' }) {
  if (!APPOINTMENT_MEMORY_STATE.enabled) {
    return { ok: false, reason: 'memory_disabled' };
  }
  const normalizedKey = normalizePatientKey(key || '');
  if (!normalizedKey) {
    return { ok: false, reason: 'empty_key' };
  }
  const appointmentNumber = String(number || '').trim();
  const slotData = sanitizeSlotForMemory(slot);
  const nowMs = Date.now();

  // Evitar duplicados obvios.
  if (appointmentNumber) {
    const existingByNumber = APPOINTMENT_MEMORY_STATE.records.find((r) => String(r.appointmentNumber || '') === appointmentNumber);
    if (existingByNumber) {
      return { ok: true, dedup: true, reason: 'already_exists_by_number', total: APPOINTMENT_MEMORY_STATE.records.length };
    }
  }
  const existingRecent = APPOINTMENT_MEMORY_STATE.records.find((r) => {
    if (!r) return false;
    const sameKey = String(r.patientKey || '') === normalizedKey;
    const sameDay = String(r.slot?.dayIso || '') === String(slotData?.dayIso || '');
    const closeTime = Number.isFinite(r.ts) && Math.abs(nowMs - r.ts) <= 2 * 60 * 1000;
    return sameKey && sameDay && closeTime;
  });
  if (existingRecent) {
    if (!existingRecent.appointmentNumber && appointmentNumber) {
      existingRecent.appointmentNumber = appointmentNumber;
      existingRecent.status = status;
      existingRecent.updatedAt = new Date(nowMs).toISOString();
      APPOINTMENT_MEMORY_STATE.dirty = true;
      const persisted = persistAppointmentMemory();
      return { ok: true, updated: true, persisted, total: APPOINTMENT_MEMORY_STATE.records.length };
    }
    return { ok: true, dedup: true, reason: 'already_exists_recent', total: APPOINTMENT_MEMORY_STATE.records.length };
  }

  const record = {
    ts: nowMs,
    createdAt: new Date(nowMs).toISOString(),
    patientKey: normalizedKey,
    appointmentNumber,
    status,
    slot: slotData
  };
  APPOINTMENT_MEMORY_STATE.records.push(record);
  APPOINTMENT_MEMORY_STATE.records = pruneAppointmentMemoryRecords(APPOINTMENT_MEMORY_STATE.records, nowMs);
  APPOINTMENT_MEMORY_STATE.dirty = true;
  const persisted = persistAppointmentMemory();
  return {
    ok: true,
    persisted,
    total: APPOINTMENT_MEMORY_STATE.records.length,
    appointmentNumber: appointmentNumber || ''
  };
}

function pruneKeyHealthRecords(records, nowMs = Date.now()) {
  const ttlMs = KEY_HEALTH_TTL_HOURS * 60 * 60 * 1000;
  const filtered = Array.isArray(records)
    ? records.filter((r) => {
      if (!r) return false;
      const key = normalizePatientKey(r.key || '');
      if (!key) return false;
      if (!Number.isFinite(r.ts)) return false;
      return nowMs - r.ts <= ttlMs;
    })
    : [];
  filtered.sort((a, b) => Number(a.ts || 0) - Number(b.ts || 0));
  if (filtered.length > KEY_HEALTH_MAX_ITEMS) {
    filtered.splice(0, filtered.length - KEY_HEALTH_MAX_ITEMS);
  }
  return filtered;
}

function loadKeyHealthState() {
  if (!KEY_HEALTH_ENABLED) {
    return { enabled: false, records: [], dirty: false };
  }
  try {
    if (!fs.existsSync(KEY_HEALTH_FILE)) {
      return { enabled: true, records: [], dirty: false };
    }
    const raw = fs.readFileSync(KEY_HEALTH_FILE, 'utf8');
    const parsed = JSON.parse(raw);
    const original = Array.isArray(parsed?.records) ? parsed.records : [];
    const pruned = pruneKeyHealthRecords(original);
    return { enabled: true, records: pruned, dirty: pruned.length !== original.length };
  } catch {
    return { enabled: true, records: [], dirty: false };
  }
}

const KEY_HEALTH_STATE = loadKeyHealthState();

function persistKeyHealth() {
  if (!KEY_HEALTH_STATE.enabled) return false;
  try {
    KEY_HEALTH_STATE.records = pruneKeyHealthRecords(KEY_HEALTH_STATE.records);
    const payload = {
      version: 1,
      updatedAt: new Date().toISOString(),
      ttlHours: KEY_HEALTH_TTL_HOURS,
      hardBlockThreshold: KEY_HARD_BLOCK_THRESHOLD,
      maxItems: KEY_HEALTH_MAX_ITEMS,
      records: KEY_HEALTH_STATE.records
    };
    fs.writeFileSync(KEY_HEALTH_FILE, JSON.stringify(payload, null, 2), 'utf8');
    KEY_HEALTH_STATE.dirty = false;
    return true;
  } catch {
    return false;
  }
}

function classifyKeyOutcome(reason, severityHint = '') {
  const hint = normalizeText(severityHint);
  if (hint === 'success') return { severity: 'success', tag: 'success' };
  const msg = normalizeText(reason || '');
  if (!msg) return { severity: 'soft', tag: 'unknown' };

  if (
    msg.includes('patient_not_found') ||
    msg.includes('no encontrado') ||
    msg.includes('404') ||
    msg.includes('key_not_set') ||
    msg.includes('key_not_confirmed') ||
    msg.includes('guardar_not_enabled') ||
    msg.includes('key_rejected:key_not_set')
  ) {
    return { severity: 'hard', tag: 'invalid_or_disabled' };
  }

  if (msg.includes('already_scheduled')) return { severity: 'soft', tag: 'already_scheduled' };
  if (msg.includes('catalog_opened')) return { severity: 'soft', tag: 'catalog_noise' };
  if (msg.includes('modal_still_open') || msg.includes('closed_without_success_alert')) return { severity: 'soft', tag: 'no_confirmation' };

  return { severity: 'soft', tag: 'other' };
}

function rememberKeyOutcome(key, reason, severityHint = '') {
  if (!KEY_HEALTH_STATE.enabled) return { ok: false, reason: 'disabled' };
  const normalizedKey = normalizePatientKey(key || '');
  if (!normalizedKey) return { ok: false, reason: 'empty_key' };

  const outcome = String(reason || '').trim() || 'unknown';
  const cls = classifyKeyOutcome(outcome, severityHint);
  const nowMs = Date.now();

  const duplicate = KEY_HEALTH_STATE.records.find((r) => {
    if (!r) return false;
    if (String(r.key || '') !== normalizedKey) return false;
    if (String(r.outcome || '') !== outcome) return false;
    return Number.isFinite(r.ts) && Math.abs(nowMs - r.ts) <= 45 * 1000;
  });
  if (duplicate) return { ok: true, dedup: true, severity: cls.severity, tag: cls.tag };

  KEY_HEALTH_STATE.records.push({
    ts: nowMs,
    createdAt: new Date(nowMs).toISOString(),
    key: normalizedKey,
    outcome,
    severity: cls.severity,
    tag: cls.tag
  });
  KEY_HEALTH_STATE.records = pruneKeyHealthRecords(KEY_HEALTH_STATE.records, nowMs);
  KEY_HEALTH_STATE.dirty = true;
  const persisted = persistKeyHealth();
  return { ok: true, persisted, severity: cls.severity, tag: cls.tag };
}

function getKeyHealthMetrics(key) {
  const normalizedKey = normalizePatientKey(key || '');
  if (!normalizedKey || !KEY_HEALTH_STATE.enabled) {
    return { hardFails: 0, softFails: 0, score: 0, blocked: false, hasSuccess: false };
  }
  const records = (KEY_HEALTH_STATE.records || [])
    .filter((r) => String(r.key || '') === normalizedKey)
    .sort((a, b) => Number(a.ts || 0) - Number(b.ts || 0));

  let lastSuccessTs = 0;
  for (const r of records) {
    if (String(r.severity || '') === 'success' && Number.isFinite(r.ts)) {
      lastSuccessTs = Math.max(lastSuccessTs, Number(r.ts));
    }
  }

  let hardFails = 0;
  let softFails = 0;
  for (const r of records) {
    const ts = Number(r.ts || 0);
    if (!Number.isFinite(ts) || ts <= lastSuccessTs) continue;
    const sev = String(r.severity || '');
    if (sev === 'hard') hardFails += 1;
    if (sev === 'soft') softFails += 1;
  }

  const score = hardFails * 100 + softFails * 14;
  const blocked = hardFails >= KEY_HARD_BLOCK_THRESHOLD;
  return { hardFails, softFails, score, blocked, hasSuccess: lastSuccessTs > 0 };
}

function applyKeyHealthToPlan(keys) {
  const base = dedupePatientKeys(keys);
  if (!KEY_HEALTH_STATE.enabled || base.length <= 1) {
    return { ordered: base, blocked: 0, total: base.length };
  }

  const scored = base.map((key, idx) => ({ key, idx, ...getKeyHealthMetrics(key) }));
  const allowed = scored
    .filter((x) => !x.blocked)
    .sort((a, b) => {
      if (a.score !== b.score) return a.score - b.score;
      return a.idx - b.idx;
    })
    .map((x) => x.key);
  const blocked = scored
    .filter((x) => x.blocked)
    .sort((a, b) => {
      if (a.score !== b.score) return a.score - b.score;
      return a.idx - b.idx;
    })
    .map((x) => x.key);

  const ordered = [...allowed, ...blocked];
  return { ordered, blocked: blocked.length, total: ordered.length };
}

function dedupePatientKeys(keys) {
  const base = Array.isArray(keys) ? keys.map((k) => normalizePatientKey(k)).filter(Boolean) : [];
  const seen = new Set();
  const out = [];

  const push = (k) => {
    const key = normalizePatientKey(k);
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(key);
  };

  for (const key of base) push(key);
  return out;
}

function getRecentSuccessPatientKeys() {
  if (!APPOINTMENT_MEMORY_STATE.enabled) return [];
  const seen = new Set();
  const out = [];
  const push = (key) => {
    const normalized = normalizePatientKey(key);
    if (!normalized || seen.has(normalized)) return;
    seen.add(normalized);
    out.push(normalized);
  };

  const recentSuccess = (APPOINTMENT_MEMORY_STATE.records || [])
    .filter((r) => r && String(r.status || '').startsWith('success'))
    .sort((a, b) => Number(b.ts || 0) - Number(a.ts || 0))
    .map((r) => normalizePatientKey(r.patientKey || ''))
    .filter(Boolean);

  for (const key of recentSuccess) push(key);
  return out;
}

function prioritizePatientKeys(keys) {
  const base = dedupePatientKeys(keys);
  if (!PRIORITIZE_RECENT_KEYS || !APPOINTMENT_MEMORY_STATE.enabled) return base;

  const seen = new Set();
  const out = [];
  const push = (k) => {
    const key = normalizePatientKey(k);
    if (!key || seen.has(key)) return;
    seen.add(key);
    out.push(key);
  };

  // Prioriza claves con éxito reciente solo cuando se habilita explícitamente.
  const recentSuccess = (APPOINTMENT_MEMORY_STATE.records || [])
    .filter((r) => r && String(r.status || '').startsWith('success'))
    .sort((a, b) => Number(b.ts || 0) - Number(a.ts || 0))
    .map((r) => normalizePatientKey(r.patientKey || ''))
    .filter(Boolean);

  for (const key of recentSuccess) push(key);
  for (const key of base) push(key);
  return out;
}

function hashStringToUInt32(input) {
  const text = String(input || '');
  let h = 2166136261;
  for (let i = 0; i < text.length; i += 1) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0) || 1;
}

function createSeededRandom(seedText) {
  let state = hashStringToUInt32(seedText);
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

function shuffleKeys(keys, rng) {
  const arr = Array.isArray(keys) ? [...keys] : [];
  const rand = typeof rng === 'function' ? rng : Math.random;
  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(rand() * (i + 1));
    const tmp = arr[i];
    arr[i] = arr[j];
    arr[j] = tmp;
  }
  return arr;
}

function buildPatientKeyAttemptPlan(keys) {
  const sequential = prioritizePatientKeys(keys);
  const rng = KEY_RANDOM_SEED ? createSeededRandom(KEY_RANDOM_SEED) : Math.random;
  const withHealth = (plan) => applyKeyHealthToPlan(plan).ordered;

  if (KEY_SELECTION_MODE === 'sequential') {
    return withHealth(sequential);
  }

  if (KEY_SELECTION_MODE === 'random') {
    // Si hay priorización reciente activa, mantiene ese encabezado y randomiza el resto.
    if (PRIORITIZE_RECENT_KEYS && APPOINTMENT_MEMORY_STATE.enabled) {
      const recent = getRecentSuccessPatientKeys();
      const seen = new Set();
      const head = [];
      for (const key of recent) {
        if (!sequential.includes(key) || seen.has(key)) continue;
        seen.add(key);
        head.push(key);
      }
      const tail = sequential.filter((k) => !seen.has(k));
      return withHealth([...head, ...shuffleKeys(tail, rng)]);
    }
    return withHealth(shuffleKeys(sequential, rng));
  }

  // recent_then_random:
  // fuerza recientes primero (desde memoria), luego randomiza el resto.
  const base = dedupePatientKeys(keys);
  const recent = getRecentSuccessPatientKeys();
  const seen = new Set();
  const head = [];
  for (const key of recent) {
    if (!base.includes(key) || seen.has(key)) continue;
    seen.add(key);
    head.push(key);
  }
  const tail = base.filter((k) => !seen.has(k));
  return withHealth([...head, ...shuffleKeys(tail, rng)]);
}

function normalizeMainMode(raw) {
  const value = String(raw || '').trim();
  return value === '2' ? '2' : '1';
}

function isInteractiveTerminal() {
  return Boolean(processStdin?.isTTY && processStdout?.isTTY);
}

async function resolveMainModeSelection() {
  const envMode = String(BOT_MAIN_MODE_ENV || '').trim();
  if (envMode === '1' || envMode === '2') {
    BOT_MAIN_MODE = normalizeMainMode(envMode);
    return BOT_MAIN_MODE;
  }

  if (!isInteractiveTerminal()) {
    BOT_MAIN_MODE = '1';
    return BOT_MAIN_MODE;
  }

  let rl;
  try {
    rl = readline.createInterface({ input: processStdin, output: processStdout });
    processStdout.write('\n');
    processStdout.write('=============================\n');
    processStdout.write(' Menu Principal del Bot\n');
    processStdout.write('=============================\n');
    processStdout.write(' 1) Generar ordenes\n');
    processStdout.write(' 2) Nota médica + Finalizar cita existente\n');
    const answer = await rl.question('Selecciona opcion [1/2] (default 1): ');
    BOT_MAIN_MODE = normalizeMainMode(answer);
  } catch {
    BOT_MAIN_MODE = '1';
  } finally {
    try {
      rl?.close();
    } catch {}
  }
  return BOT_MAIN_MODE;
}

async function clickIfVisible(page, text, timeout = 8000) {
  const el = page.locator(`text=${text}`).first();
  await el.waitFor({ state: 'visible', timeout });
  await el.click();
}

async function activateRadDropdown(page, ddlId, optionText, timing = {}) {
  const preClickWaitMs = Number.isFinite(Number(timing.preClickWaitMs)) ? Number(timing.preClickWaitMs) : 160;
  const firstPopupWaitMs = Number.isFinite(Number(timing.firstPopupWaitMs)) ? Number(timing.firstPopupWaitMs) : 900;
  const reopenWaitMs = Number.isFinite(Number(timing.reopenWaitMs)) ? Number(timing.reopenWaitMs) : 180;
  const popupWaitMs = Number.isFinite(Number(timing.popupWaitMs)) ? Number(timing.popupWaitMs) : 1500;
  const optionWaitMs = Number.isFinite(Number(timing.optionWaitMs)) ? Number(timing.optionWaitMs) : 2000;
  const optionSettleMs = Number.isFinite(Number(timing.optionSettleMs)) ? Number(timing.optionSettleMs) : 120;
  const popupHideWaitMs = Number.isFinite(Number(timing.popupHideWaitMs)) ? Number(timing.popupHideWaitMs) : 1000;
  const finalSettleMs = Number.isFinite(Number(timing.finalSettleMs)) ? Number(timing.finalSettleMs) : 120;

  const root = page.locator(`#${ddlId}`);
  await root.waitFor({ state: 'visible', timeout: 10000 });

  // 1) Activar control (1 click; segundo solo si no abre).
  await root.click({ force: true });
  await page.waitForTimeout(preClickWaitMs);

  // 2) Si hay popup visible, seleccionar por texto. Si no, reintentar abrir una vez.
  const popup = page.locator(`#${ddlId}_DropDown`);
  try {
    await popup.waitFor({ state: 'visible', timeout: firstPopupWaitMs });
  } catch {
    await root.click({ force: true });
    await page.waitForTimeout(reopenWaitMs);
  }

  try {
    await popup.waitFor({ state: 'visible', timeout: popupWaitMs });

    const option = popup.getByText(optionText, { exact: false }).first();
    await option.waitFor({ state: 'visible', timeout: optionWaitMs });
    await option.click({ force: true });
    await page.waitForTimeout(optionSettleMs);
  } catch {
    // Fallback teclado si popup no se pudo operar por texto.
    await page.keyboard.press('ArrowDown');
    await page.waitForTimeout(90);
    await page.keyboard.press('Enter');
  }

  // 3) Confirmar cierre (sin click extra para evitar bug visual).
  try {
    await popup.waitFor({ state: 'hidden', timeout: popupHideWaitMs });
  } catch {}

  await page.waitForTimeout(finalSettleMs);
}

async function forceSelectDdlByValue(page, ddlId, expectedValue, fallbackText) {
  return await page.evaluate(({ ddlId, expectedValue, fallbackText }) => {
    const ddl = window.$find && window.$find(ddlId);
    if (!ddl) return { ok: false, reason: 'ddl_not_found' };
    let item = null;
    try {
      if (ddl.findItemByValue) item = ddl.findItemByValue(expectedValue);
      if (!item && ddl.findItemByText) item = ddl.findItemByText(fallbackText);
      if (!item) return { ok: false, reason: 'item_not_found' };
      if (ddl.trackChanges) ddl.trackChanges();
      item.select();
      if (ddl.commitChanges) ddl.commitChanges();
      if (ddl.raisePropertyChanged) ddl.raisePropertyChanged('selectedItem');
      return { ok: true };
    } catch (e) {
      return { ok: false, reason: String(e) };
    }
  }, { ddlId, expectedValue, fallbackText });
}

async function ensureDdlSelected(page, ddlId, optionText, expectedValue) {
  for (let i = 0; i < 3; i += 1) {
    await activateRadDropdown(page, ddlId, optionText);
    const state = await getDdlState(page, ddlId);
    if (state.selectedValue === expectedValue) return state;

    await forceSelectDdlByValue(page, ddlId, expectedValue, optionText);
    const forced = await getDdlState(page, ddlId);
    if (forced.selectedValue === expectedValue) return forced;

    await page.waitForTimeout(250);
  }
  return await getDdlState(page, ddlId);
}

async function getDdlState(page, ddlId) {
  return await page.evaluate((ddlId) => {
    const out = {};
    const root = document.getElementById(ddlId);
    const cs = document.getElementById(`${ddlId}_ClientState`);
    out.rootValue = root ? (root.getAttribute('value') || root.textContent || '').trim() : '';

    if (cs?.value) {
      out.clientStateRaw = cs.value;
      try {
        const parsed = JSON.parse(cs.value);
        out.selectedText = decodeURIComponent(parsed.selectedText || '').trim();
        out.selectedValue = parsed.selectedValue || '';
      } catch {
        out.selectedText = '';
        out.selectedValue = '';
      }
    } else {
      out.clientStateRaw = '';
      out.selectedText = '';
      out.selectedValue = '';
    }

    return out;
  }, ddlId);
}

async function ensureDropdownReadyForLogin(page) {
  // Trigger validaciones/estado con API Telerik si existe.
  await page.evaluate(() => {
    const safePost = (id) => {
      try {
        const ddl = window.$find && window.$find(id);
        if (ddl && ddl.postback) ddl.postback();
      } catch {}
    };
    safePost('ctl00_usercontrol2_ddlCompany');
    safePost('ctl00_usercontrol2_ddlDepartamento');
  });
  await page.waitForTimeout(500);
}

async function ensureCalendarContext(page) {
  const calendarSelector = '.rsContentTable td, .k-scheduler-table td, .k-scheduler-content td, td[role="gridcell"]';
  const started = Date.now();
  const hasCalendar = async () => {
    try {
      const first = page.locator(calendarSelector).first();
      if ((await first.count()) === 0) return false;
      return await first.isVisible();
    } catch {
      return false;
    }
  };

  const waitCalendar = async (ms = 2600) => {
    const step = 250;
    const ticks = Math.ceil(ms / step);
    for (let i = 0; i < ticks; i += 1) {
      if (await hasCalendar()) return true;
      await page.waitForTimeout(step);
    }
    return false;
  };

  const clickPracticaCard = async () => {
    try {
      return await page.evaluate(() => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .trim();
        const visible = (el) => {
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 80;
        };

        const cards = Array.from(document.querySelectorAll('div,section,article'))
          .filter((n) => visible(n))
          .map((n) => ({ n, t: normalize(n.textContent || ''), r: n.getBoundingClientRect() }))
          .filter((x) => x.t.includes('practica medica'))
          .filter((x) => x.r.width >= 220 && x.r.width <= 680 && x.r.height >= 120 && x.r.height <= 360);
        if (!cards.length) return false;
        cards.sort((a, b) => (a.r.width * a.r.height) - (b.r.width * b.r.height));
        const card = cards[0].n;
        const assist = Array.from(card.querySelectorAll('button,a,span,div'))
          .find((n) => normalize(n.textContent || '').includes('assist'));
        card.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        card.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        card.click();
        if (assist instanceof HTMLElement) {
          assist.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          assist.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          assist.click();
          assist.click();
        } else {
          card.click();
        }
        return true;
      });
    } catch {
      return false;
    }
  };

  const clickAgendaOption = async () => {
    try {
      return await page.evaluate(() => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .trim();
        const visible = (el) => {
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
        };

        const nodes = Array.from(
          document.querySelectorAll(
            '.option-title, .settings-option, .CoreWelcomeFormItem, a, button, div[role="option"], li'
          )
        ).filter(visible);
        const target = nodes.find((n) => normalize(n.textContent || '') === 'agenda medica')
          || nodes.find((n) => normalize(n.textContent || '').includes('agenda medica'));
        if (!(target instanceof HTMLElement)) return false;
        target.scrollIntoView({ block: 'center', inline: 'center' });
        target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        target.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        target.click();
        return true;
      });
    } catch {
      return false;
    }
  };

  if (await hasCalendar()) {
    console.log(`Paso 6 calendar ya visible (${Date.now() - started}ms)`);
    return;
  }

  // Flujo fijo: seleccionar tarjeta "Práctica médica" y luego opción "Agenda médica".
  for (let i = 0; i < 3; i += 1) {
    const card = await clickPracticaCard();
    console.log(`Paso 6 action practica_card ok=${card}`);
    await page.waitForTimeout(220);

    const agenda = await clickAgendaOption();
    console.log(`Paso 6 action agenda_option ok=${agenda}`);
    if (agenda && await waitCalendar(4200)) {
      console.log(`Paso 6 calendar listo (${Date.now() - started}ms)`);
      return;
    }

    // Fallback mínimo por texto para evitar bloqueo del panel lateral.
    try {
      const byText = page.getByText(/Agenda m[eé]dica/i).first();
      await byText.waitFor({ state: 'visible', timeout: 900 });
      await byText.click({ force: true, timeout: 1000 });
    } catch {}
    if (await waitCalendar(2600)) {
      console.log(`Paso 6 calendar listo (${Date.now() - started}ms)`);
      return;
    }
  }

  throw new Error('No se pudo asegurar calendario en Paso 6.');
}

async function getWeekStatus(page) {
  try {
    return await page.evaluate(() => {
      const normalize = (s) => (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
      };

      // Contar eventos "NO DISPONIBLE" usando solo .k-event (Kendo scheduler)
      const kEvents = Array.from(document.querySelectorAll('.k-event')).filter(visible);
      const noDispEvents = kEvents.filter((e) => normalize(e.textContent).includes('no disponible'));

      // Determinar columnas de día por posiciones X únicas de los k-event
      let dayColumns = 0;
      if (kEvents.length > 0) {
        const uniqueLefts = new Set(kEvents.map((e) => Math.round(e.getBoundingClientRect().left / 10)));
        dayColumns = uniqueLefts.size;
      }
      if (dayColumns <= 0) dayColumns = 7; // fallback: vista semanal

      // Fingerprint para detectar calendario estancado
      const fingerprint = kEvents
        .map((e) => {
          const r = e.getBoundingClientRect();
          return `${Math.round(r.left)}:${(e.textContent || '').trim().slice(0, 15)}`;
        })
        .sort()
        .join('|');

      const blocked = noDispEvents.length >= Math.min(dayColumns, 5);

      return { blocked, noDispEvents: noDispEvents.length, dayColumns, totalKEvents: kEvents.length, fingerprint };
    });
  } catch {
    return { blocked: false, noDispEvents: 0, dayColumns: 7, totalKEvents: 0, fingerprint: '' };
  }
}

async function getVisibleWeekInfo(page) {
  return await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
    };
    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const parseMd = (txt) => {
      const m = (txt || '').match(/(\d{1,2})\s*\/\s*(\d{1,2})/);
      if (!m) return null;
      const mm = Number(m[1]);
      const dd = Number(m[2]);
      if (!mm || !dd) return null;
      let year = todayStart.getFullYear();
      let d = new Date(year, mm - 1, dd);
      if (d < new Date(todayStart.getTime() - 1000 * 60 * 60 * 24 * 300)) d = new Date(year + 1, mm - 1, dd);
      return d;
    };

    const rangeNode = Array.from(document.querySelectorAll('.k-lg-date-format, .k-nav-current, .rsDateHeader')).find(visible);
    const rangeLabel = rangeNode ? (rangeNode.textContent || '').trim() : '';

    const headers = Array.from(document.querySelectorAll('thead th, .k-scheduler-header th, .rsHeader')).filter(visible);
    const dates = [];
    for (const h of headers) {
      const d = parseMd(normalize(h.textContent || ''));
      if (d) dates.push(d);
    }
    dates.sort((a, b) => a.getTime() - b.getTime());

    const start = dates.length ? dates[0] : null;
    const end = dates.length ? dates[dates.length - 1] : null;
    return {
      label: rangeLabel,
      startIso: start ? start.toISOString().slice(0, 10) : '',
      endIso: end ? end.toISOString().slice(0, 10) : '',
      startTs: start ? start.getTime() : 0,
      endTs: end ? end.getTime() : 0
    };
  });
}

async function applyAgendaFilter(page) {
  try {
    const filterBtn = page.locator('button:has-text("Filtrar"), input[value="Filtrar"]').first();
    await filterBtn.waitFor({ state: 'visible', timeout: 2500 });
    await filterBtn.click({ force: true });
    await page.waitForTimeout(600);
    console.log('AGENDA_FILTER_OK');
    return true;
  } catch {
    return false;
  }
}

async function isAppointmentModalVisible(page) {
  return await page.evaluate(() => {
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 50 && r.height > 30;
    };
    const dialogs = Array.from(
      document.querySelectorAll('[role="dialog"], .k-window, .k-dialog, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, .modal')
    ).filter(visible);
    if (!dialogs.length) return false;
    const text = dialogs.map((d) => (d.textContent || '').toLowerCase()).join(' ');
    return /cita|paciente|documento|guardar|comentario/.test(text) || dialogs.length > 0;
  });
}

async function waitForManualWindow(page, ms = 20000) {
  const step = 600;
  const ticks = Math.ceil(ms / step);
  for (let i = 0; i < ticks; i += 1) {
    if (await isAppointmentModalVisible(page)) return true;
    await page.waitForTimeout(step);
  }
  return false;
}

async function goToNextCalendarRange(page) {
  const getRangeMeta = async () => {
    return await page.evaluate(() => {
      const text = (document.querySelector('.k-lg-date-format, .k-nav-current, .rsDateHeader')?.textContent || '').trim();
      const m = text.match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/);
      let ts = 0;
      if (m) {
        const mm = Number(m[1]);
        const dd = Number(m[2]);
        const yy = Number(m[3]);
        const d = new Date(yy, mm - 1, dd);
        if (!Number.isNaN(d.getTime())) ts = d.getTime();
      }
      return { text, ts };
    });
  };

  const before = await getRangeMeta();
  const clicked = await page.evaluate(() => {
    const candidates = [
      '.k-nav-next',
      '.rsNext',
      'button[title="Next"]',
      'button[aria-label="Next"]',
      '[aria-label="Next"]',
      '.k-link.k-nav-next'
    ];
    for (const sel of candidates) {
      const el = document.querySelector(sel);
      if (el instanceof HTMLElement) {
        el.click();
        return true;
      }
    }
    // fallback por texto
    const byText = Array.from(document.querySelectorAll('button, a, span')).find((n) =>
      /next|siguiente|›|»/i.test((n.textContent || '').trim())
    );
    if (byText instanceof HTMLElement) {
      byText.click();
      return true;
    }
    return false;
  });
  if (!clicked) return false;

  for (let i = 0; i < 20; i += 1) {
    await page.waitForTimeout(180);
    const after = await getRangeMeta();
    if (after.text && after.text !== before.text) {
      if (before.ts > 0 && after.ts > 0) {
        const days = Math.round((after.ts - before.ts) / (24 * 60 * 60 * 1000));
        if (days > 8) {
          console.log(`WEEK_JUMP_DETECTED before="${before.text}" after="${after.text}" days=${days}`);
          // Si el salto fue doble, intenta retroceder una vez y reporta fallo para reintento controlado.
          await page.evaluate(() => {
            const prev = document.querySelector('.k-nav-prev, .rsPrev, button[title="Previous"], button[aria-label="Previous"]');
            if (prev instanceof HTMLElement) prev.click();
          });
          await page.waitForTimeout(500);
          return false;
        }
      }
      return true;
    }
  }

  return true;
}

async function goToPreviousCalendarRange(page) {
  const getRangeMeta = async () => {
    return await page.evaluate(() => {
      const text = (document.querySelector('.k-lg-date-format, .k-nav-current, .rsDateHeader')?.textContent || '').trim();
      const m = text.match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/);
      let ts = 0;
      if (m) {
        const mm = Number(m[1]);
        const dd = Number(m[2]);
        const yy = Number(m[3]);
        const d = new Date(yy, mm - 1, dd);
        if (!Number.isNaN(d.getTime())) ts = d.getTime();
      }
      return { text, ts };
    });
  };

  const before = await getRangeMeta();
  const clicked = await page.evaluate(() => {
    const candidates = [
      '.k-nav-prev',
      '.rsPrev',
      'button[title="Previous"]',
      'button[aria-label="Previous"]',
      '[aria-label="Previous"]',
      '.k-link.k-nav-prev'
    ];
    for (const sel of candidates) {
      const el = document.querySelector(sel);
      if (el instanceof HTMLElement) {
        el.click();
        return true;
      }
    }
    const byText = Array.from(document.querySelectorAll('button, a, span')).find((n) =>
      /prev|previous|anterior|‹|«/i.test((n.textContent || '').trim())
    );
    if (byText instanceof HTMLElement) {
      byText.click();
      return true;
    }
    return false;
  });
  if (!clicked) return false;

  for (let i = 0; i < 20; i += 1) {
    await page.waitForTimeout(180);
    const after = await getRangeMeta();
    if (after.text && after.text !== before.text) {
      if (before.ts > 0 && after.ts > 0) {
        const days = Math.round((before.ts - after.ts) / (24 * 60 * 60 * 1000));
        if (days > 8) {
          console.log(`WEEK_BACK_JUMP_DETECTED before="${before.text}" after="${after.text}" days=${days}`);
          await page.evaluate(() => {
            const next = document.querySelector('.k-nav-next, .rsNext, button[title="Next"], button[aria-label="Next"]');
            if (next instanceof HTMLElement) next.click();
          });
          await page.waitForTimeout(500);
          return false;
        }
      }
      return true;
    }
  }

  return true;
}

async function clickCalendarToday(page) {
  const clicked = await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
    };
    const nodes = Array.from(
      document.querySelectorAll(
        '.k-nav-today, .rsToday, button, a, span, div[role="button"], [title], [aria-label]'
      )
    ).filter(visible);
    const direct = nodes.find((n) => {
      const txt = normalize(
        `${n.textContent || ''} ${n.getAttribute('title') || ''} ${n.getAttribute('aria-label') || ''} ${n.id || ''}`
      );
      return txt === 'hoy' || txt.includes(' hoy') || txt.startsWith('hoy ') || txt.includes('today');
    });
    if (!(direct instanceof HTMLElement)) return false;
    direct.click();
    return true;
  });
  if (!clicked) return false;
  await page.waitForTimeout(420);
  return true;
}

async function ensureCalendarOnCurrentWeek(page, options = {}) {
  const maxAdjustments = Math.min(8, Math.max(2, Number(options?.maxAdjustments || 6)));
  const shouldApplyFilter = options?.applyFilter === true;
  const todayDate = new Date();
  const todayStart = new Date(todayDate.getFullYear(), todayDate.getMonth(), todayDate.getDate()).getTime();

  for (let i = 0; i < maxAdjustments; i += 1) {
    const info = await getVisibleWeekInfo(page);
    const inWeek = info.startTs > 0 && info.endTs > 0 && todayStart >= info.startTs && todayStart <= info.endTs;
    if (inWeek) {
      console.log(`CALENDAR_CURRENT_WEEK_OK start=${info.startIso || '-'} end=${info.endIso || '-'} step=${i}`);
      return true;
    }

    let action = '';
    if (i === 0) {
      const todayClicked = await clickCalendarToday(page);
      if (todayClicked) action = 'today';
    }
    if (!action) {
      if (info.startTs > 0 && todayStart < info.startTs) {
        const moved = await goToPreviousCalendarRange(page);
        if (moved) action = 'prev';
      } else if (info.endTs > 0 && todayStart > info.endTs) {
        const moved = await goToNextCalendarRange(page);
        if (moved) action = 'next';
      } else {
        const todayClicked = await clickCalendarToday(page);
        if (todayClicked) action = 'today';
      }
    }

    console.log(
      `CALENDAR_CURRENT_WEEK_ADJUST step=${i + 1}/${maxAdjustments} action=${action || 'none'} start=${info.startIso || '-'} end=${info.endIso || '-'} label="${info.label || ''}"`
    );
    if (!action) break;
    await page.waitForTimeout(760);
    if (shouldApplyFilter) await applyAgendaFilter(page);
    await ensureWorkingHoursVisible(page);
    await page.waitForTimeout(320);
  }

  const finalInfo = await getVisibleWeekInfo(page);
  const finalOk = finalInfo.startTs > 0 && finalInfo.endTs > 0 && todayStart >= finalInfo.startTs && todayStart <= finalInfo.endTs;
  console.log(
    `CALENDAR_CURRENT_WEEK_${finalOk ? 'OK' : 'WARN'} start=${finalInfo.startIso || '-'} end=${finalInfo.endIso || '-'} label="${finalInfo.label || ''}"`
  );
  return finalOk;
}

async function dismissNetworkBanners(page) {
  try {
    const count = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
      let dismissed = 0;
      // Buscar banners/toasts de intermitencia, error de red, etc.
      const candidates = Array.from(
        document.querySelectorAll('.k-notification, .toast, .alert, [role="alert"], .notification, .banner, .swal2-container, .ajs-message')
      );
      // También buscar divs con fondo naranja/rojo posicionados arriba
      const allDivs = Array.from(document.querySelectorAll('div, aside, section'));
      for (const el of allDivs) {
        const r = el.getBoundingClientRect();
        if (r.width < 150 || r.height < 30 || r.height > 200) continue;
        const st = getComputedStyle(el);
        const bg = st.backgroundColor;
        // Detectar fondos naranja/rojo/warning
        const isWarningBg = bg.includes('rgb(255, 1') || bg.includes('rgb(255, 8') || bg.includes('rgb(230,') || bg.includes('rgb(243,') || bg.includes('orange');
        if (isWarningBg && !candidates.includes(el)) candidates.push(el);
      }
      for (const el of candidates) {
        const txt = normalize(el.textContent || '');
        if (txt.includes('intermitencia') || txt.includes('intermitente') || txt.includes('conexion') || txt.includes('red')) {
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          if (st.display === 'none' || st.visibility === 'hidden' || r.width < 50) continue;
          // Buscar botón de cerrar
          const closeBtn = el.querySelector('.close, .btn-close, [aria-label="close"], [aria-label="Close"], button');
          if (closeBtn instanceof HTMLElement) {
            closeBtn.click();
            dismissed += 1;
          } else {
            el.style.display = 'none';
            dismissed += 1;
          }
        }
      }
      return dismissed;
    });
    if (count > 0) {
      console.log(`DISMISS_NETWORK_BANNERS_OK count=${count}`);
      await page.waitForTimeout(200);
    }
    return count;
  } catch {
    return 0;
  }
}

async function dismissStaleP2HPopup(page, options = {}) {
  const clickAbrirModulo = options?.clickAbrirModulo === true;
  try {
    // Paso 1: detectar si el popup existe (solo detección, sin click)
    const detected = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();

      // Estrategia 1: clase P2H directa
      const popup = document.querySelector('.div_hos930AvisoPaciente');
      if (popup) {
        const st = getComputedStyle(popup);
        const r = popup.getBoundingClientRect();
        if (st.display !== 'none' && st.visibility !== 'hidden' && r.width >= 50 && r.height >= 50) {
          const txt = normalize(popup.textContent || '');
          if (txt.includes('nueva cita') || txt.includes('cita asignada') || txt.includes('paciente')) {
            return { found: true, via: 'p2h_class' };
          }
        }
      }

      // Estrategia 2: alertify dialog con header "Nueva cita asignada"
      const dialogs = document.querySelectorAll('.ajs-dialog');
      for (const d of dialogs) {
        const dst = getComputedStyle(d);
        if (dst.display === 'none' || dst.visibility === 'hidden') continue;
        const header = d.querySelector('.ajs-header');
        if (header && normalize(header.textContent).includes('nueva cita asignada')) {
          return { found: true, via: 'alertify' };
        }
      }

      return { found: false };
    });

    if (!detected?.found) return false;

    // Paso 2: click "Abrir módulo" usando Playwright locator (dispara postback Telerik)
    if (clickAbrirModulo) {
      const btnSelectors = [
        '[id$="MP_HOS930_btnModulo"]',
        'button:has-text("Abrir módulo")',
        'a:has-text("Abrir módulo")'
      ];
      for (const sel of btnSelectors) {
        try {
          const loc = page.locator(sel).first();
          if ((await loc.count()) === 0) continue;
          if (!(await loc.isVisible())) continue;
          await loc.click({ force: true, timeout: 2000 });
          console.log(`DISMISS_STALE_P2H_POPUP action=abrir_modulo via=${detected.via} sel="${sel}"`);
          await page.waitForTimeout(800);
          return { found: true, action: 'abrir_modulo', via: detected.via };
        } catch {}
      }
      console.log(`DISMISS_STALE_P2H_POPUP action=abrir_modulo_FAIL via=${detected.via}`);
    }

    // Modo 1 o fallback: ocultar/cerrar
    try {
      await page.evaluate(() => {
        const popup = document.querySelector('.div_hos930AvisoPaciente');
        if (popup) {
          popup.style.display = 'none';
          popup.style.visibility = 'hidden';
          popup.style.pointerEvents = 'none';
        }
        const closeBtn = document.querySelector('.ajs-dialog .ajs-close') || document.querySelector('[id$="MP_HOS930_btnCerrar"]');
        if (closeBtn instanceof HTMLElement) closeBtn.click();
      });
    } catch {}
    console.log(`DISMISS_STALE_P2H_POPUP action=hidden via=${detected.via}`);
    await page.waitForTimeout(200);
    return { found: true, action: 'hidden', via: detected.via };
  } catch {
    return false;
  }
}

async function ensureWorkingHoursVisible(page) {
  try {
    const btn = page.locator('button:has-text("Mostrar horas laborales"), button:has-text("Mostrar horas laborables")').first();
    await btn.waitFor({ state: 'visible', timeout: 1800 });
    await btn.click({ force: true });
    await page.waitForTimeout(500);
    console.log('WORKING_HOURS_TOGGLE_OK');
    return true;
  } catch {
    return false;
  }
}

async function findAvailableCalendarCell(page) {
  return await page.evaluate(() => {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    const visible = (el) => {
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 18 && r.height > 12;
    };
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();

    const parseHeaderDate = (headerText, attrDate) => {
      if (attrDate) {
        const n = Number(attrDate);
        if (!Number.isNaN(n) && String(attrDate).length >= 10) {
          const d = new Date(String(attrDate).length > 10 ? n : n * 1000);
          if (!Number.isNaN(d.getTime())) return d;
        }
        const p = Date.parse(attrDate);
        if (!Number.isNaN(p)) return new Date(p);
      }
      const text = headerText || '';
      const m = text.match(/(\d{1,2})\s*\/\s*(\d{1,2})/); // mm/dd
      if (!m) return null;
      const mm = Number(m[1]);
      const dd = Number(m[2]);
      if (!mm || !dd) return null;
      let year = today.getFullYear();
      let d = new Date(year, mm - 1, dd);
      // Manejo de cruce de anio (diciembre/enero)
      if (d < new Date(today.getTime() - 1000 * 60 * 60 * 24 * 300)) d = new Date(year + 1, mm - 1, dd);
      return d;
    };

    const selectors = [
      '.rsContentTable td',
      '.k-scheduler-table td',
      '.k-scheduler-content td',
      'td[role="gridcell"]'
    ];

    let cells = [];
    for (const sel of selectors) {
      cells = Array.from(document.querySelectorAll(sel)).filter(visible);
      if (cells.length) break;
    }
    if (!cells.length) return { ok: false, reason: 'no_cells' };

    const rows = new Map();
    for (const td of cells) {
      const tr = td.parentElement;
      if (!tr) continue;
      if (!rows.has(tr)) rows.set(tr, []);
      rows.get(tr).push(td);
    }

    const headers = Array.from(document.querySelectorAll('thead th, .rsHeader')).filter(visible);

    const hasEvent = (td) => Boolean(td.querySelector('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]'));
    const hardBadClass = (cls) => /disabled|othermonth|holiday|blocked|busy|outside/i.test(cls || '');
    const softBadClass = (cls) => /nonwork|off/i.test(cls || '');
    const eventRects = Array.from(
      document.querySelectorAll('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]')
    )
      .filter(visible)
      .map((el) => el.getBoundingClientRect())
      .map((r) => ({ left: r.left, right: r.right, top: r.top, bottom: r.bottom }));
    const coveredByEvent = (td) => {
      const r = td.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      return eventRects.some((er) => cx >= er.left && cx <= er.right && cy >= er.top && cy <= er.bottom);
    };

    const strictCandidates = [];
    const fallbackCandidates = [];
    const lenientCandidates = [];
    const samples = [];

    for (const td of cells) {
      const cls = td.className || '';
      const text = (td.textContent || '').trim();
      const col = td.cellIndex ?? 0;
      const row = td.parentElement;
      const rowIdx = row && row.parentElement ? Array.from(row.parentElement.children).indexOf(row) : 0;

      // Header por columna
      const header = headers.find((h) => (h.cellIndex ?? -1) === col) || headers[col] || null;
      const headerText = header ? (header.textContent || '') : '';
      const headerNorm = normalize(headerText);
      const attrDate = header ? (header.getAttribute('data-date') || header.getAttribute('date') || header.getAttribute('aria-label') || '') : '';
      const parsed = parseHeaderDate(headerText, attrDate);

      const dayStart = parsed ? new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()) : null;
      const isSunday = parsed ? parsed.getDay() === 0 : headerNorm.startsWith('dom');
      const isFutureOrToday = dayStart ? dayStart >= today : false;
      const occupied = hasEvent(td);
      const covered = coveredByEvent(td);

      let score = 0;
      if (isFutureOrToday) score += 25;
      if (isSunday) score -= 30;
      if (!occupied) score += 12;
      if (occupied) score -= 20;
      if (!covered) score += 8;
      if (covered) score -= 25;
      if (!text) score += 4;
      if (softBadClass(cls)) score -= 10;
      if (hardBadClass(cls)) score -= 20;
      // preferir filas medias/altas para reducir scroll y tiempos
      score += Math.max(0, 8 - Math.abs(8 - rowIdx));

      const candidate = (() => {
        const r = td.getBoundingClientRect();
        return {
          x: Math.floor(r.left + r.width / 2),
          y: Math.floor(r.top + r.height / 2),
          className: cls,
          text: text.slice(0, 50),
          header: headerText.trim().slice(0, 60),
          dayIso: dayStart ? dayStart.toISOString().slice(0, 10) : '',
          score
        };
      })();

      if (samples.length < 8) {
        samples.push({
          className: cls.slice(0, 120),
          header: headerText.trim().slice(0, 60),
          dayIso: dayStart ? dayStart.toISOString().slice(0, 10) : '',
          occupied,
          covered,
          isSunday,
          isFutureOrToday,
          text: text.slice(0, 30)
        });
      }

      // Nivel 1: estricto (futuro/hoy, no domingo, libre, sin clases duras/blandas)
      if (isFutureOrToday && !isSunday && !occupied && !covered && !text && !hardBadClass(cls) && !softBadClass(cls)) {
        strictCandidates.push(candidate);
        continue;
      }

      // Nivel 2: futuro/hoy + libre + no duro (acepta nonwork/off)
      if (isFutureOrToday && !occupied && !covered && !text && !hardBadClass(cls)) {
        fallbackCandidates.push(candidate);
        continue;
      }

      // Nivel 3: libre + no duro (si no se pudo parsear fecha/header)
      if (!occupied && !covered && !text && !hardBadClass(cls)) {
        lenientCandidates.push(candidate);
      }
    }

    const pickBest = (arr) => {
      if (!arr.length) return null;
      return arr.sort((a, b) => b.score - a.score)[0];
    };

    const bestStrict = pickBest(strictCandidates);
    if (bestStrict) return { ok: true, strategy: 'strict', ...bestStrict };

    const bestFallback = pickBest(fallbackCandidates);
    if (bestFallback) return { ok: true, strategy: 'fallback_future', ...bestFallback };

    const bestLenient = pickBest(lenientCandidates);
    if (bestLenient) return { ok: true, strategy: 'lenient_any_free', ...bestLenient };

    const emergency = cells.find((td) => {
      const cls = td.className || '';
      const txt = (td.textContent || '').trim();
      return !hasEvent(td) && !coveredByEvent(td) && !txt && !/disabled|othermonth/i.test(cls);
    });
    if (emergency) {
      const r = emergency.getBoundingClientRect();
      return {
        ok: true,
        strategy: 'emergency_any_clickable',
        x: Math.floor(r.left + r.width / 2),
        y: Math.floor(r.top + r.height / 2),
        className: emergency.className || '',
        text: ((emergency.textContent || '').trim()).slice(0, 50),
        header: '',
        dayIso: ''
      };
    }

    return {
      ok: false,
      reason: 'no_available_future_cell',
      debug: {
        cells: cells.length,
        headers: headers.length,
        strictCandidates: strictCandidates.length,
        fallbackCandidates: fallbackCandidates.length,
        lenientCandidates: lenientCandidates.length,
        samples
      }
    };
  });
}

async function selectValidCalendarField(page) {
  const clickByDomFallback = async () => {
    return await page.evaluate(() => {
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 18 && r.height > 12;
      };
      const hasEvent = (td) => Boolean(td.querySelector('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]'));
      const hardBadClass = (cls) => /disabled|othermonth|holiday|blocked|busy|outside/i.test(cls || '');
      const timeLike = (s) => /\d{1,2}:\d{2}\s*(am|pm)/i.test((s || '').trim());

      const pick = (arr) => {
        if (!arr.length) return null;
        const byMiddle = arr.sort((a, b) => Math.abs(a.row - 8) - Math.abs(b.row - 8));
        return byMiddle[0] || null;
      };

      const cells = Array.from(
        document.querySelectorAll('.rsContentTable tbody td, .k-scheduler-table tbody td, .k-scheduler-content tbody td, td[role="gridcell"]')
      ).filter(visible);
      const candidates = [];

      for (const td of cells) {
        const cls = td.className || '';
        if (hardBadClass(cls) || hasEvent(td)) continue;
        if ((td.cellIndex ?? 0) <= 0) continue;
        const text = (td.textContent || '').trim();
        if (timeLike(text)) continue;
        const row = td.parentElement && td.parentElement.parentElement
          ? Array.from(td.parentElement.parentElement.children).indexOf(td.parentElement)
          : 0;
        candidates.push({ td, cls, text: text.slice(0, 40), row });
      }

      const chosen = pick(candidates);
      if (!chosen) return { ok: false };

      const el = chosen.td;
      const fire = (type) => el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
      fire('mousedown');
      fire('mouseup');
      fire('click');
      fire('mousedown');
      fire('mouseup');
      fire('click');
      return { ok: true, className: chosen.cls, text: chosen.text };
    });
  };

  // Busca en rango actual, luego siguiente rango si hace falta.
  for (let pass = 0; pass < 2; pass += 1) {
    const spot = await findAvailableCalendarCell(page);
    if (spot.ok) {
      await page.mouse.click(spot.x, spot.y);
      await page.waitForTimeout(350);
      // Click 2 controlado en la misma casilla para abrir modal.
      if (!(await isAppointmentModalVisible(page))) {
        await page.mouse.click(spot.x, spot.y);
        await page.waitForTimeout(220);
      }

      if (!(await isAppointmentModalVisible(page))) {
        const js = await clickByDomFallback();
        if (js.ok) {
          console.log(`CALENDAR_JS_CLICK_OK class="${js.className}" text="${js.text}"`);
          await page.waitForTimeout(500);
        }
      }
      return spot;
    }

    if (pass === 0) {
      const moved = await goToNextCalendarRange(page);
      if (!moved) {
        if (spot.debug) {
          console.log(`CALENDAR_DEBUG ${JSON.stringify(spot.debug)}`);
        }
        throw new Error(`No se encontro casilla disponible y no se pudo avanzar calendario: ${spot.reason}`);
      }
      continue;
    }

    if (spot.debug) {
      console.log(`CALENDAR_DEBUG ${JSON.stringify(spot.debug)}`);
    }
    throw new Error(`No se encontro casilla disponible: ${spot.reason}`);
  }
}

async function openAppointmentModalByHumanLikeClicks(page) {
  const selectors = [
    '.rsContentTable tbody tr td:not(:first-child)',
    '.k-scheduler-content table tbody tr td:not(:first-child)',
    '.k-scheduler-table tbody tr td:not(:first-child)',
    'td[role="gridcell"]'
  ];

  const badClass = /disabled|othermonth|holiday|blocked|busy|outside|readonly/i;
  const timeLike = /\d{1,2}:\d{2}\s*(am|pm)/i;

  const started = Date.now();
  const maxMs = 6000;

  for (const selector of selectors) {
    if (Date.now() - started > maxMs) break;
    const cells = page.locator(selector);
    const total = await cells.count();
    const limit = Math.min(total, 14);
    if (!limit) continue;

    for (let i = 0; i < limit; i += 1) {
      if (Date.now() - started > maxMs) break;
      const cell = cells.nth(i);
      try {
        if (!(await cell.isVisible())) continue;
        const cls = (await cell.getAttribute('class')) || '';
        if (badClass.test(cls)) continue;

        const hasEvent =
          (await cell.locator('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]').count()) > 0;
        if (hasEvent) continue;

        const txt = ((await cell.innerText()) || '').trim();
        if (timeLike.test(txt)) continue;

        await cell.click({ force: true, timeout: 1200 });
        await page.waitForTimeout(80);

        await cell.click({ force: true, timeout: 900 });
        await page.waitForTimeout(180);

        let modalOpen = (await isNuevaCitaModalVisible(page)) || (await isNuevaCitaAsignadaModalVisible(page));
        if (!modalOpen) {
          // Fallback adicional: disparar eventos y on* handlers del td.
          await cell.evaluate((el) => {
            const fire = (type) =>
              el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
            fire('mousedown');
            fire('mouseup');
            fire('click');
            fire('mousedown');
            fire('mouseup');
            fire('click');
            try {
              if (typeof el.ondblclick === 'function') el.ondblclick(new MouseEvent('dblclick'));
            } catch {}
            try {
              const raw = el.getAttribute('ondblclick');
              if (raw) new Function(raw).call(el);
            } catch {}
            try {
              const raw = el.getAttribute('onclick');
              if (raw) new Function(raw).call(el);
            } catch {}
          });
          await page.waitForTimeout(220);
          modalOpen = (await isNuevaCitaModalVisible(page)) || (await isNuevaCitaAsignadaModalVisible(page));
        }

        if (modalOpen) {
          console.log(`CALENDAR_HUMAN_CLICK_OK selector="${selector}" idx=${i} class="${cls}" text="${txt.slice(0, 35)}"`);
          return { ok: true, selector, idx: i, className: cls, text: txt.slice(0, 50) };
        }
      } catch {}
    }
  }

  return { ok: false };
}

async function getTopCalendarCandidates(page, limit = 8) {
  return await page.evaluate((limit) => {
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 18 && r.height > 12;
    };
    const hardBadClass = (cls) => /disabled|othermonth|holiday|blocked|busy|outside|readonly/i.test(cls || '');
    const timeLike = (s) => /\d{1,2}:\d{2}\s*(am|pm)/i.test((s || '').trim());
    const eventRects = Array.from(
      document.querySelectorAll('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]')
    )
      .filter(visible)
      .map((el) => el.getBoundingClientRect())
      .map((r) => ({ left: r.left, right: r.right, top: r.top, bottom: r.bottom }));
    const coveredByEvent = (td) => {
      const r = td.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      return eventRects.some((er) => cx >= er.left && cx <= er.right && cy >= er.top && cy <= er.bottom);
    };

    const cells = Array.from(
      document.querySelectorAll('.rsContentTable tbody tr td:not(:first-child), .k-scheduler-content tbody tr td:not(:first-child), .k-scheduler-table tbody tr td:not(:first-child), td[role="gridcell"]')
    ).filter(visible);

    const out = [];
    for (const td of cells) {
      const cls = td.className || '';
      if (hardBadClass(cls)) continue;
      if (coveredByEvent(td)) continue;
      const txt = (td.textContent || '').trim();
      if (timeLike(txt) || txt) continue;
      const hasEvent = Boolean(td.querySelector('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]'));
      if (hasEvent) continue;

      const r = td.getBoundingClientRect();
      const tr = td.parentElement;
      const rowIdx = tr && tr.parentElement ? Array.from(tr.parentElement.children).indexOf(tr) : 0;
      const colIdx = td.cellIndex ?? 1;
      // Priorizar horas de media mañana/tarde temprana y columnas de días cercanos.
      const score = 100 - Math.abs(9 - rowIdx) * 4 - colIdx;
      out.push({
        x: Math.floor(r.left + r.width / 2),
        y: Math.floor(r.top + r.height / 2),
        className: cls,
        text: txt.slice(0, 40),
        rowIdx,
        colIdx,
        score
      });
    }

    out.sort((a, b) => b.score - a.score);
    return out.slice(0, limit);
  }, limit);
}

async function clickTopCalendarCandidateByRank(page, rank = 0) {
  return await page.evaluate((rank) => {
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 18 && r.height > 12;
    };
    const hardBadClass = (cls) => /disabled|othermonth|holiday|blocked|busy|outside|readonly/i.test(cls || '');
    const timeLike = (s) => /\d{1,2}:\d{2}\s*(am|pm)/i.test((s || '').trim());
    const eventRects = Array.from(
      document.querySelectorAll('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]')
    )
      .filter(visible)
      .map((el) => el.getBoundingClientRect())
      .map((r) => ({ left: r.left, right: r.right, top: r.top, bottom: r.bottom }));
    const coveredByEvent = (td) => {
      const r = td.getBoundingClientRect();
      const cx = r.left + r.width / 2;
      const cy = r.top + r.height / 2;
      return eventRects.some((er) => cx >= er.left && cx <= er.right && cy >= er.top && cy <= er.bottom);
    };

    const pool = Array.from(
      document.querySelectorAll('.rsContentTable tbody tr td:not(:first-child), .k-scheduler-content tbody tr td:not(:first-child), .k-scheduler-table tbody tr td:not(:first-child), td[role="gridcell"]')
    ).filter(visible);
    const headers = Array.from(document.querySelectorAll('thead th, .rsHeader')).filter(visible);
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const headerTextByCol = new Map();
    const headerDateByCol = new Map();
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const parseHeaderDate = (text) => {
      const m = (text || '').match(/(\d{1,2})\s*\/\s*(\d{1,2})/);
      if (!m) return null;
      const mm = Number(m[1]);
      const dd = Number(m[2]);
      if (!mm || !dd) return null;
      let y = today.getFullYear();
      let d = new Date(y, mm - 1, dd);
      if (d < new Date(today.getTime() - 1000 * 60 * 60 * 24 * 300)) d = new Date(y + 1, mm - 1, dd);
      return d;
    };
    for (const h of headers) {
      const col = h.cellIndex ?? -1;
      if (col >= 0) {
        const htxt = normalize(h.textContent || '');
        headerTextByCol.set(col, htxt);
        headerDateByCol.set(col, parseHeaderDate(htxt));
      }
    }
    const parseMinutes = (timeText) => {
      const raw = normalize(timeText || '');
      const m = raw.match(/(\d{1,2})\s*:\s*(\d{2})\s*(am|pm)/);
      if (!m) return null;
      let hh = Number(m[1]) % 12;
      const mm = Number(m[2]);
      const ampm = m[3];
      if (ampm === 'pm') hh += 12;
      return hh * 60 + mm;
    };

    const candidates = [];
    for (const td of pool) {
      const cls = td.className || '';
      if (hardBadClass(cls)) continue;
      if (coveredByEvent(td)) continue;
      const txt = (td.textContent || '').trim();
      if (timeLike(txt) || txt) continue;
      const hasEvent = Boolean(td.querySelector('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]'));
      if (hasEvent) continue;
      const tr = td.parentElement;
      const rowIdx = tr && tr.parentElement ? Array.from(tr.parentElement.children).indexOf(tr) : 0;
      const colIdx = td.cellIndex ?? 1;
      // Evitar primera columna (normalmente domingo o columna no util).
      if (colIdx <= 0) continue;
      // Rango laboral aproximado (evita madrugada/no disponibles).
      if (rowIdx < 12 || rowIdx > 38) continue;
      const htxt = headerTextByCol.get(colIdx) || '';
      const hdate = headerDateByCol.get(colIdx) || null;
      const isSunday = htxt.startsWith('dom');
      if (isSunday) continue;
      const timeCell = tr ? tr.querySelector('td:first-child, th:first-child') : null;
      const mins = parseMinutes(timeCell ? timeCell.textContent || '' : '');

      const isPastDay = hdate ? hdate < today : false;
      const isPastTimeToday =
        hdate &&
        mins !== null &&
        hdate.getFullYear() === today.getFullYear() &&
        hdate.getMonth() === today.getMonth() &&
        hdate.getDate() === today.getDate() &&
        mins < (now.getHours() * 60 + now.getMinutes());
      if (isPastDay || isPastTimeToday) continue;

      candidates.push({
        td,
        cls,
        txt,
        rowIdx,
        colIdx,
        htxt,
        dayIso: hdate ? hdate.toISOString().slice(0, 10) : '',
        minutes: mins !== null ? mins : rowIdx
      });
    }

    // Orden natural del DOM: primera casilla visible, luego la siguiente, etc.
    // Esto sigue exactamente el flujo "si hay data, saltar a la siguiente".
    const chosen = candidates[rank];
    if (!chosen) return { ok: false };

    const el = chosen.td;
    const fire = (type) => el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
    // Primer click: posicionarse en la casilla.
    fire('mousedown');
    fire('mouseup');
    fire('click');

    const clsAfter = (el.className || '').toLowerCase();
    const selectedAfterFirst =
      clsAfter.includes('selected') || clsAfter.includes('active') || el.getAttribute('aria-selected') === 'true';

    return {
      ok: true,
      rowIdx: chosen.rowIdx,
      colIdx: chosen.colIdx,
      header: chosen.htxt,
      dayIso: chosen.dayIso,
      minutes: chosen.minutes,
      className: chosen.cls,
      text: (chosen.txt || '').slice(0, 40),
      totalCandidates: candidates.length,
      selectedAfterFirst
    };
  }, rank);
}

async function openNuevaCitaViaSchedulerApi(page) {
  const result = await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 12 && r.height > 10;
    };

    const jq = window.jQuery || window.$;
    if (!jq) return { ok: false, reason: 'jq_missing' };

    const schedulerHost = Array.from(document.querySelectorAll('.k-scheduler'))
      .find((el) => {
        try {
          return !!jq(el).data('kendoScheduler');
        } catch {
          return false;
        }
      });
    if (!(schedulerHost instanceof HTMLElement)) return { ok: false, reason: 'scheduler_not_found' };
    const scheduler = jq(schedulerHost).data('kendoScheduler');
    if (!scheduler) return { ok: false, reason: 'scheduler_data_missing' };

    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    const nowMinutes = now.getHours() * 60 + now.getMinutes();

    const cells = Array.from(
      schedulerHost.querySelectorAll('.k-scheduler-content td[role="gridcell"], .k-scheduler-content td, .rsContentTable td[role="gridcell"], .rsContentTable td')
    ).filter(visible);
    if (!cells.length) return { ok: false, reason: 'no_cells' };

    const checkAvailabilityFn = typeof window.checkAvailability === 'function' ? window.checkAvailability : null;
    const candidates = [];
    for (const td of cells) {
      try {
        const slot = scheduler.slotByElement(td);
        if (!slot || !slot.startDate || !slot.endDate) continue;
        const start = new Date(slot.startDate);
        const end = new Date(slot.endDate);
        const dayStart = new Date(start.getFullYear(), start.getMonth(), start.getDate()).getTime();
        if (dayStart < todayStart) continue;
        if (start.getDay() === 0) continue; // domingo
        const minutes = start.getHours() * 60 + start.getMinutes();
        if (minutes > 18 * 60) continue;
        if (dayStart === todayStart && minutes < nowMinutes) continue;

        let available = true;
        if (checkAvailabilityFn) {
          try {
            available = Boolean(checkAvailabilityFn(start, end, null));
          } catch {
            available = false;
          }
        }
        if (!available) continue;

        candidates.push({
          startTs: start.getTime(),
          startIso: start.toISOString(),
          endIso: end.toISOString(),
          minutes
        });
      } catch {}
    }

    if (!candidates.length) {
      return { ok: false, reason: 'no_available_slot_by_checkAvailability', checked: cells.length };
    }
    candidates.sort((a, b) => a.startTs - b.startTs);
    const chosen = candidates[0];

    try {
      scheduler.addEvent({
        start: new Date(chosen.startTs),
        end: new Date(new Date(chosen.startTs).getTime() + 30 * 60 * 1000),
        title: ''
      });
    } catch (e) {
      return { ok: false, reason: `scheduler_addEvent_error:${String(e)}` };
    }

    return {
      ok: true,
      chosenStart: chosen.startIso,
      chosenEnd: chosen.endIso,
      totalCandidates: candidates.length
    };
  });

  if (!result.ok) {
    console.log(`SCHEDULER_API_OPEN_FAIL reason=${result.reason || 'unknown'}`);
    return { ok: false };
  }

  await page.waitForTimeout(450);
  const opened = await isNuevaCitaModalVisible(page);
  if (opened) {
    console.log(
      `SCHEDULER_API_OPEN_OK start=${result.chosenStart || '-'} end=${result.chosenEnd || '-'} candidates=${result.totalCandidates || 0}`
    );
    return { ok: true };
  }

  console.log('SCHEDULER_API_OPEN_NO_MODAL');
  return { ok: false };
}

async function fastOpenNuevaCitaFromCandidates(page) {
  const selectors = [
    '.k-scheduler-content table tbody td:not(:first-child)',
    '.k-scheduler-table tbody td:not(:first-child)',
    '.rsContentTable tbody td:not(:first-child)',
    'td[role="gridcell"]'
  ];

  let activeSelector = '';
  for (const sel of selectors) {
    const l = page.locator(sel);
    const n = await l.count();
    if (n > 0) {
      activeSelector = sel;
      break;
    }
  }
  if (!activeSelector) return { ok: false };

  const result = await page.evaluate(({ sel }) => {
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 18 && r.height > 12;
    };
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();

    const parseHeaderDate = (text, todayDate) => {
      const m = (text || '').match(/(\d{1,2})\s*\/\s*(\d{1,2})/);
      if (!m) return null;
      const mm = Number(m[1]);
      const dd = Number(m[2]);
      if (!mm || !dd) return null;
      let year = todayDate.getFullYear();
      let d = new Date(year, mm - 1, dd);
      if (d < new Date(todayDate.getTime() - 1000 * 60 * 60 * 24 * 300)) d = new Date(year + 1, mm - 1, dd);
      return d;
    };

    const parseMinutes = (timeText) => {
      const t = normalize(timeText || '');
      const m = t.match(/(\d{1,2})\s*:\s*(\d{2})\s*(am|pm)/);
      if (!m) return null;
      let hh = Number(m[1]) % 12;
      const mm = Number(m[2]);
      if (m[3] === 'pm') hh += 12;
      return hh * 60 + mm;
    };

    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const todayIso = today.toISOString().slice(0, 10);
    const nowMinutes = now.getHours() * 60 + now.getMinutes();

    const headers = Array.from(document.querySelectorAll('thead th, .k-scheduler-header th, .rsHeader')).filter(visible);
    const dayByCol = new Map();
    for (const h of headers) {
      const col = h.cellIndex ?? -1;
      if (col < 0) continue;
      const txt = normalize(h.textContent || '');
      const date = parseHeaderDate(txt, today);
      dayByCol.set(col, { text: txt, date });
    }

    const eventRects = Array.from(
      document.querySelectorAll('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]')
    )
      .filter(visible)
      .map((e) => e.getBoundingClientRect())
      .map((r) => ({ left: r.left, right: r.right, top: r.top, bottom: r.bottom }));

    const overlapsEvent = (rect) => {
      const cx = rect.left + rect.width / 2;
      const cy = rect.top + rect.height / 2;
      return eventRects.some((er) => cx >= er.left && cx <= er.right && cy >= er.top && cy <= er.bottom);
    };

    const cells = Array.from(document.querySelectorAll(sel)).filter(visible);
    const candidates = [];

    for (let i = 0; i < cells.length; i += 1) {
      const td = cells[i];
      const cls = (td.className || '').toLowerCase();
      if (/disabled|othermonth|holiday|blocked|busy|outside|readonly/.test(cls)) continue;

      const row = td.parentElement;
      const rowIdx = row && row.parentElement ? Array.from(row.parentElement.children).indexOf(row) : 0;
      const colIdx = td.cellIndex ?? 0;
      if (colIdx <= 0) continue;

      // Casilla vacia visible.
      const txt = (td.textContent || '').trim();
      if (txt) continue;

      // Ocupacion por evento dentro del td.
      const hasEvent = Boolean(td.querySelector('.k-event, .rsApt, [class*="appointment"], [class*="Appointment"], [class*="event"]'));
      if (hasEvent) continue;

      // Ocupacion por overlap visual (evento flotante sobre la casilla).
      const rect = td.getBoundingClientRect();
      if (overlapsEvent(rect)) continue;

      // Rango horario 00:00 - 18:00 (y para hoy, evitar horas ya pasadas).
      const firstCell = row ? row.querySelector('td:first-child, th:first-child') : null;
      const parsedMinutes = parseMinutes(firstCell ? firstCell.textContent || '' : '');
      const minutes = parsedMinutes !== null ? parsedMinutes : (rowIdx * 30);
      if (minutes < 0 || minutes > 18 * 60) continue;

      const dayInfo = dayByCol.get(colIdx) || { text: '', date: null };
      const dayIso = dayInfo.date
        ? new Date(dayInfo.date.getFullYear(), dayInfo.date.getMonth(), dayInfo.date.getDate()).toISOString().slice(0, 10)
        : '';

      if (dayIso && dayIso < todayIso) continue;
      if (dayIso === todayIso && minutes < nowMinutes) continue;

      candidates.push({
        domIdx: i,
        rowIdx,
        colIdx,
        minutes,
        dayIso,
        dayText: dayInfo.text,
        x: Math.floor(rect.left + rect.width / 2),
        y: Math.floor(rect.top + rect.height / 2)
      });
    }

    if (!candidates.length) {
      return { startCol: -1, candidates: [] };
    }

    // Iniciar en columna del dia actual; si no existe, primera columna >= hoy.
    const todayCols = Array.from(dayByCol.entries())
      .filter(([, v]) => v.date && new Date(v.date.getFullYear(), v.date.getMonth(), v.date.getDate()).toISOString().slice(0, 10) === todayIso)
      .map(([col]) => col);
    let startCol = todayCols.length ? Math.min(...todayCols) : -1;
    if (startCol < 0) {
      const futureCols = Array.from(dayByCol.entries())
        .filter(([, v]) => v.date && new Date(v.date.getFullYear(), v.date.getMonth(), v.date.getDate()).toISOString().slice(0, 10) >= todayIso)
        .map(([col]) => col);
      if (futureCols.length) startCol = Math.min(...futureCols);
    }
    if (startCol < 0) startCol = Math.min(...candidates.map((c) => c.colIdx));

    // Secuencia: desde startCol por dia, y dentro del dia por hora ascendente.
    candidates.sort((a, b) => {
      const ao = a.colIdx >= startCol ? a.colIdx : a.colIdx + 1000;
      const bo = b.colIdx >= startCol ? b.colIdx : b.colIdx + 1000;
      if (ao !== bo) return ao - bo;
      if (a.minutes !== b.minutes) return a.minutes - b.minutes;
      return a.rowIdx - b.rowIdx;
    });

    return { startCol, candidates };
  }, { sel: activeSelector });

  const candidates = result.candidates || [];
  console.log(`FAST_CANDIDATES total=${candidates.length} startCol=${result.startCol} selector="${activeSelector}"`);
  if (!candidates.length) return { ok: false };

  const maxSteps = Math.min(candidates.length, 40);
  const startedAt = Date.now();
  for (let i = 0; i < maxSteps; i += 1) {
    if (Date.now() - startedAt > 45000) {
      console.log(`FAST_LOOP_TIMEOUT step=${i}/${maxSteps}`);
      break;
    }

    // Si el modal ya se abrio (auto o manual), cortar loop y continuar con clave.
    if (await isNuevaCitaModalVisible(page)) {
      console.log(`FAST_MODAL_DETECTADO step=${i}`);
      return { ok: true, opened: true, assigned: false, slot: null };
    }

    if (await isNuevaCitaAsignadaModalVisible(page)) {
      console.log(`FAST_MODAL_ASIGNADA_DETECTADO step=${i}`);
      return { ok: true, opened: false, assigned: true, slot: null };
    }

    const c = candidates[i];
    try {
      // Click 1: posicionar.
      await page.mouse.click(c.x, c.y, { delay: 30 });
      await page.waitForTimeout(240);

      // Click 2: abrir modal en la misma casilla.
      await page.mouse.click(c.x, c.y, { delay: 30 });
      await page.waitForTimeout(320);

      let opened = await isNuevaCitaModalVisible(page);
      let assigned = await isNuevaCitaAsignadaModalVisible(page);

      if (!opened && !assigned) {
        // Doble click real sobre la casilla para UI que no abre con doble click sintético rápido.
        await page.mouse.click(c.x, c.y, { clickCount: 2, delay: 70 });
        await page.waitForTimeout(300);
        opened = await isNuevaCitaModalVisible(page);
        assigned = await isNuevaCitaAsignadaModalVisible(page);
      }

      if (!opened && !assigned) {
        // Fallback: click del locator por índice DOM detectado.
        const cell = page.locator(activeSelector).nth(c.domIdx);
        await cell.click({ force: true, timeout: 900 });
        await page.waitForTimeout(220);
        await cell.click({ force: true, timeout: 900, clickCount: 2 });
        await page.waitForTimeout(260);
        opened = await isNuevaCitaModalVisible(page);
        assigned = await isNuevaCitaAsignadaModalVisible(page);
      }

      if (!opened && !assigned && ENABLE_ENTER_FALLBACK) {
        await page.keyboard.press('Enter');
        await page.waitForTimeout(160);
        opened = await isNuevaCitaModalVisible(page);
        assigned = await isNuevaCitaAsignadaModalVisible(page);
      }

      if (opened) {
        console.log(
          `FAST_SLOT_OK idx=${c.domIdx} day=${c.dayIso || '-'} col=${c.colIdx} min=${c.minutes} step=${i + 1}/${candidates.length}`
        );
        return {
          ok: true,
          opened: true,
          assigned: false,
          slot: {
            x: c.x,
            y: c.y,
            domIdx: c.domIdx,
            selector: activeSelector,
            colIdx: c.colIdx,
            minutes: c.minutes,
            dayIso: c.dayIso || ''
          }
        };
      }

      if (assigned) {
        console.log(`FAST_SLOT_ASSIGNED idx=${c.domIdx} day=${c.dayIso || '-'} col=${c.colIdx} min=${c.minutes}`);
        await closeNuevaCitaAsignadaModal(page);
        await page.waitForTimeout(140);
      } else {
        console.log(`FAST_SLOT_SKIP idx=${c.domIdx} day=${c.dayIso || '-'} col=${c.colIdx} min=${c.minutes}`);
      }
    } catch {}
  }

  return { ok: false };
}

async function waitForNuevaCitaModal(page, ms = 6000) {
  const step = 250;
  const ticks = Math.ceil(ms / step);
  for (let i = 0; i < ticks; i += 1) {
    const opened = await isNuevaCitaModalVisible(page);
    const assigned = await isNuevaCitaAsignadaModalVisible(page);
    if (opened || assigned) return { opened, assigned };
    await page.waitForTimeout(step);
  }
  return { opened: false, assigned: false };
}

async function clickModuloButton(page, options = {}) {
  const waitBeforeMsRaw = Number(options?.waitBeforeMs);
  const waitBeforeMs = Number.isFinite(waitBeforeMsRaw) ? Math.max(0, Math.round(waitBeforeMsRaw)) : 700;
  // Espera configurable; en popup de celda debe ser mínima para que no se cierre.
  if (waitBeforeMs > 0) {
    await page.waitForTimeout(waitBeforeMs);
  }

  const tryClick = async (locator, label) => {
    try {
      await locator.waitFor({ state: 'visible', timeout: 2500 });
      await locator.click({ force: true });
      console.log(`MODULO_CLICK_OK via ${label}`);
      return true;
    } catch {
      return false;
    }
  };

  const candidates = [
    { label: 'role-button', locator: page.getByRole('button', { name: /m[oó]dulo/i }).first() },
    { label: 'role-link', locator: page.getByRole('link', { name: /m[oó]dulo/i }).first() },
    { label: 'text-regex', locator: page.locator('text=/M[oó]dulo/i').first() },
    {
      label: 'css-text',
      locator: page.locator(
        'button:has-text("Módulo"), button:has-text("Modulo"), a:has-text("Módulo"), a:has-text("Modulo")'
      ).first()
    }
  ];

  for (const c of candidates) {
    if (await tryClick(c.locator, c.label)) return true;
  }

  // Fallback por JS en elemento visible que contenga "Modulo"/"Módulo".
  const jsClicked = await page.evaluate(() => {
    const visible = (el) => {
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 18 && r.height > 12;
    };
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const nodes = Array.from(document.querySelectorAll('button,a,span,div')).filter(visible);
    const target = nodes.find((n) => normalize(n.textContent).includes('modulo'));
    if (!target) return false;
    target.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
    target.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
    target.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
    return true;
  });

  if (jsClicked) {
    console.log('MODULO_CLICK_OK via js-fallback');
    return true;
  }

  console.log('MODULO_CLICK_FAIL');
  return false;
}

async function readModuloLoadState(page) {
  return await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
    };

    const schedulerVisible = Array.from(
      document.querySelectorAll('.k-scheduler, .k-scheduler-content, .rsContentTable')
    ).some(visible);

    const nodes = document.querySelectorAll('h1,h2,h3,h4,h5,button,a,label,span,div,li');
    const seen = new Set();
    const texts = [];
    for (const n of nodes) {
      if (texts.length >= 360) break;
      if (!(n instanceof HTMLElement)) continue;
      if (!visible(n)) continue;
      const t = normalize(n.textContent || n.getAttribute('aria-label') || n.getAttribute('title') || '');
      if (!t || t.length < 3) continue;
      const k = t.slice(0, 110);
      if (seen.has(k)) continue;
      seen.add(k);
      texts.push(k);
    }
    const blob = texts.join(' | ');
    const tokens = [
      'nota medica',
      'apreciacion diagnostica',
      'estudios complementarios',
      'finalizar cita',
      'receta medica',
      'laboratorio',
      'imagenologia',
      'ordenes'
    ];
    const tokenHit = tokens.find((t) => blob.includes(t)) || '';
    const hash = String(location.hash || '').toLowerCase();
    const hashSignal = /#t(3|4|5|6|7|8|9|10|11|12)\b/.test(hash);

    // Detectar tab "Tablero Médico" en RadTabStrip (puede existir pero no estar activo)
    let tableroTabExists = false;
    const tabLinks = document.querySelectorAll('.rtsLink, [role="tab"]');
    for (const tab of tabLinks) {
      const txt = normalize(tab.textContent || '');
      if (txt.includes('tablero medico') || txt.includes('tablero médico')) {
        tableroTabExists = true;
        // Si el tab existe pero no está activo, clickearlo para activarlo
        const isActive = tab.classList.contains('rtsSelected') ||
          tab.parentElement?.classList.contains('rtsSelected');
        if (!isActive) {
          try { tab.click(); } catch {}
        }
        // Limpiar selección de texto que puede quedar por el click
        try { window.getSelection()?.removeAllRanges(); } catch {}
        break;
      }
    }

    return {
      loaded: Boolean(tokenHit) || hashSignal || tableroTabExists,
      signal: tokenHit ? `token:${tokenHit}` : (hashSignal ? `hash:${hash}` : (tableroTabExists ? 'tab:tablero_medico' : 'none')),
      schedulerVisible,
      tableroTabExists,
      url: location.href
    };
  });
}

async function openNotaMedicaWhenInPatientModule(page, origin = '') {
  if (isPageClosedSafe(page)) return false;
  let moduleState = null;
  try {
    moduleState = await readModuloLoadState(page);
  } catch {}

  const inModule = Boolean(moduleState?.loaded);
  console.log(
    `NOTA_MEDICA_STEP_START origin=${origin || '-'} in_module=${inModule ? 1 : 0} signal=${moduleState?.signal || 'none'}`
  );
  if (!inModule) return false;

  if (RELOAD_BEFORE_NOTA_MEDICA) {
    console.log(`NOTA_MEDICA_PRE_RELOAD origin=${origin || '-'} timeout=${RELOAD_BEFORE_NOTA_MEDICA_TIMEOUT_MS}ms`);
    let reloadOk = false;
    try {
      await page.reload({ waitUntil: 'domcontentloaded', timeout: RELOAD_BEFORE_NOTA_MEDICA_TIMEOUT_MS });
      await waitForTimeoutRaw(page, 380);
      reloadOk = true;
    } catch (e) {
      console.log(`NOTA_MEDICA_PRE_RELOAD_FAIL origin=${origin || '-'} err=${e?.message || e}`);
    }

    if (reloadOk) {
      const reloadStarted = Date.now();
      let readyAfterReload = false;
      let lastSignal = 'none';
      while ((Date.now() - reloadStarted) < RELOAD_BEFORE_NOTA_MEDICA_TIMEOUT_MS) {
        if (isPageClosedSafe(page)) break;
        let st = null;
        try {
          st = await readModuloLoadState(page);
        } catch {}
        if (st?.signal && st.signal !== 'none') lastSignal = st.signal;
        if (st?.loaded) {
          readyAfterReload = true;
          console.log(
            `NOTA_MEDICA_POST_RELOAD_READY origin=${origin || '-'} signal=${st.signal} elapsed=${Date.now() - reloadStarted}ms`
          );
          break;
        }
        await sleepRaw(RELOAD_BEFORE_NOTA_MEDICA_POLL_MS);
      }
      if (!readyAfterReload) {
        console.log(
          `NOTA_MEDICA_POST_RELOAD_TIMEOUT origin=${origin || '-'} elapsed=${Date.now() - reloadStarted}ms last_signal=${lastSignal}`
        );
      }
    } else {
      console.log(
        `NOTA_MEDICA_POST_RELOAD_SOFT_FAIL origin=${origin || '-'} continue_click=1`
      );
    }
  }

  const notaOpened = await openNotaMedicaFromSidebar(page, origin || moduleState?.signal || '-');
  if (!notaOpened) return false;
  const filled = await fillNotaMedicaAntecedentesAndGenerateIA(page, origin || moduleState?.signal || '-');
  if (!filled) return false;
  const recetaOk = await generarReceta(page, origin || moduleState?.signal || '-');
  if (!recetaOk) return false;
  return true;
}

async function waitForModuloLoaded(page, origin = '', options = {}) {
  const autoPostModule = options?.autoPostModule !== false;
  const started = Date.now();
  let polls = 0;
  let lastSignal = 'none';
  const pollLogEvery = Math.max(1, Math.round(3000 / MODULE_LOAD_POLL_INTERVAL_MS));

  while ((Date.now() - started) < MODULE_LOAD_POLL_TIMEOUT_MS) {
    let pages = [];
    try {
      const ctx = page?.context?.();
      pages = ctx?.pages?.().filter((p) => p && !p.isClosed()) || [];
    } catch {}
    if ((!pages || !pages.length) && page && !page.isClosed()) pages = [page];

    let firstState = null;
    for (let i = 0; i < pages.length; i += 1) {
      const targetPage = pages[i];
      let state;
      try {
        state = await readModuloLoadState(targetPage);
      } catch {
        continue;
      }
      if (!firstState) firstState = state;
      if (state.signal && state.signal !== 'none') lastSignal = state.signal;
      if (state.loaded) {
        const elapsed = Date.now() - started;
        console.log(
          `MODULO_CARGADO_OK origin=${origin || '-'} signal=${state.signal} elapsed=${elapsed}ms page_idx=${i + 1}/${pages.length}`
        );
        console.log('ACCESO_MODULO_PACIENTE_OK');
        if (!autoPostModule) {
          console.log(`MODULO_POST_ACTION_SKIP origin=${origin || '-'} auto_post=0`);
          return true;
        }
        const notaOk = await openNotaMedicaWhenInPatientModule(targetPage, origin || state.signal);
        if (!notaOk) {
          console.log(`NOTA_MEDICA_STEP_FAIL origin=${origin || '-'} elapsed=${Date.now() - started}ms`);
          return false;
        }
        return true;
      }
    }

    if (polls % pollLogEvery === 0) {
      const elapsed = Date.now() - started;
      const state = firstState || { signal: 'none', schedulerVisible: false };
      console.log(
        `MODULO_CARGA_POLL origin=${origin || '-'} elapsed=${elapsed}ms signal=${state.signal} scheduler=${state.schedulerVisible ? 1 : 0} pages=${pages.length}`
      );
    }
    polls += 1;
    await sleepRaw(MODULE_LOAD_POLL_INTERVAL_MS);
  }

  console.log(
    `MODULO_CARGADO_TIMEOUT origin=${origin || '-'} elapsed=${Date.now() - started}ms last_signal=${lastSignal}`
  );
  return false;
}

async function getLoadedModuloPage(page, timeoutMs = 9000) {
  const started = Date.now();
  while ((Date.now() - started) < timeoutMs) {
    if (isPageClosedSafe(page)) break;
    let pages = [];
    try {
      const ctx = page?.context?.();
      pages = ctx?.pages?.().filter((p) => p && !p.isClosed()) || [];
    } catch {}
    if ((!pages || !pages.length) && page && !page.isClosed()) pages = [page];

    for (let i = 0; i < pages.length; i += 1) {
      const p = pages[i];
      try {
        const st = await readModuloLoadState(p);
        if (st?.loaded) {
          console.log(
            `MODULO_PAGE_TARGET_OK page_idx=${i + 1}/${pages.length} signal=${st.signal || 'none'} elapsed=${Date.now() - started}ms`
          );
          return { page: p, state: st, index: i, total: pages.length };
        }
      } catch {}
    }
    await sleepRaw(Math.max(120, MODULE_LOAD_POLL_INTERVAL_MS));
  }
  console.log(`MODULO_PAGE_TARGET_FAIL timeout=${timeoutMs}ms`);
  return null;
}

async function readAntecedentesMenuState(page) {
  if (isPageClosedSafe(page)) return { notaSelected: false, alergiasSelected: false, activeLabel: '' };
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };

      const nodes = Array.from(
        document.querySelectorAll(
          '.active, .selected, .k-state-selected, .k-selected, [aria-selected="true"], [aria-current="page"]'
        )
      ).filter(visible);

      const leftNodes = nodes
        .map((n) => {
          const r = n.getBoundingClientRect();
          return {
            txt: normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || ''),
            left: r.left,
            top: r.top,
            width: r.width,
            area: r.width * r.height
          };
        })
        .filter((x) => x.left >= 20 && x.left <= 360 && x.top >= 120 && x.width <= 340)
        .sort((a, b) => a.area - b.area);

      const activeLabel = leftNodes[0]?.txt || '';
      const notaSelected = leftNodes.some((x) => x.txt === 'nota medica' || x.txt.startsWith('nota medica'));
      const alergiasSelected = leftNodes.some((x) => x.txt === 'alergias' || x.txt.startsWith('alergias'));

      return { notaSelected, alergiasSelected, activeLabel };
    });
  } catch {
    return { notaSelected: false, alergiasSelected: false, activeLabel: '' };
  }
}

async function isNotaMedicaViewActive(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    const state = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };

      const activeNodes = Array.from(
        document.querySelectorAll(
          '.active, .selected, .k-state-selected, .k-selected, [aria-selected="true"], [aria-current="page"]'
        )
      ).filter(visible);
      const activeTxt = activeNodes
        .map((n) => {
          const r = n.getBoundingClientRect();
          return {
            t: normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || ''),
            left: r.left,
            top: r.top,
            width: r.width
          };
        })
        .filter((x) => x.left >= 20 && x.left <= 360 && x.top >= 120 && x.width <= 340);

      const hasNotaSelected = activeTxt.some((x) => x.t === 'nota medica' || x.t.startsWith('nota medica'));
      const hasAlergiasSelected = activeTxt.some((x) => x.t === 'alergias' || x.t.startsWith('alergias'));

      // Señal de formulario en panel de contenido (no barra lateral).
      const labels = Array.from(document.querySelectorAll('label,span,div,p,th,td,h1,h2,h3')).filter((n) => {
        if (!visible(n)) return false;
        const r = n.getBoundingClientRect();
        return r.left >= 220;
      });
      const blob = labels
        .map((n) => normalize(n.textContent || ''))
        .filter(Boolean)
        .slice(0, 900)
        .join(' | ');
      const hasNotaForm =
        blob.includes('consulta por') &&
        blob.includes('triage') &&
        (blob.includes('presente enfermedad') || blob.includes('enfermedad presente')) &&
        blob.includes('apreciacion diagnostica') &&
        blob.includes('diagnostico principal');

      return { hasNotaSelected, hasAlergiasSelected, hasNotaForm };
    });

    if (state?.hasNotaSelected) return true;
    if (state?.hasAlergiasSelected) return false;
    return Boolean(state?.hasNotaForm);
  } catch {
    return false;
  }
}

async function clickNotaMedicaInLeftMenu(page, origin = '', elapsedBase = 0) {
  if (isPageClosedSafe(page)) return false;
  const target = await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
    };
    const elementText = (el) => normalize(el?.textContent || el?.getAttribute?.('title') || el?.getAttribute?.('aria-label') || '');
    const hasAntecedentesContext = (el) => {
      let p = el;
      for (let i = 0; i < 8 && p; i += 1) {
        const t = elementText(p);
        if (
          (t.includes('antecedentes') && t.includes('alergias')) ||
          (t.includes('alergias') && t.includes('estudios complementarios'))
        ) {
          return true;
        }
        p = p.parentElement;
      }
      return false;
    };
    const hasSiblingContext = (el) => {
      const parent = el?.parentElement;
      if (!parent) return false;
      const items = Array.from(parent.children).map((n) => elementText(n));
      const hasAlergias = items.some((t) => t.includes('alergias'));
      const hasEstudios = items.some((t) => t.includes('estudios complementarios'));
      return hasAlergias && hasEstudios;
    };

    const nodes = Array.from(
      document.querySelectorAll('a,button,li,div,span,[role="menuitem"],[role="tab"],[title],[aria-label]')
    ).filter(visible);

    const alergiasY = nodes
      .filter((n) => {
        const txt = elementText(n);
        const r = n.getBoundingClientRect();
        return (txt === 'alergias' || txt.startsWith('alergias')) && r.left >= 20 && r.left <= 360;
      })
      .map((n) => n.getBoundingClientRect().top);
    const estudiosY = nodes
      .filter((n) => {
        const txt = elementText(n);
        const r = n.getBoundingClientRect();
        return txt.includes('estudios complementarios') && r.left >= 20 && r.left <= 360;
      })
      .map((n) => n.getBoundingClientRect().top);

    const yMin = alergiasY.length ? (Math.min(...alergiasY) - 30) : 220;
    const yMax = estudiosY.length ? (Math.max(...estudiosY) + 120) : 980;

    const candidates = [];
    for (const n of nodes) {
      const txt = elementText(n);
      const exact = txt === 'nota medica';
      const close = txt.startsWith('nota medica') && txt.length <= 36;
      if (!(exact || close)) continue;

      const r = n.getBoundingClientRect();
      if (r.left < 20 || r.left > 360) continue;
      if (r.top < yMin || r.top > yMax) continue;
      if (r.width < 70 || r.width > 280) continue;
      if (r.height < 22 || r.height > 95) continue;

      const siblingContext = hasSiblingContext(n);
      const antecedentesContext = hasAntecedentesContext(n);
      if (!siblingContext && !antecedentesContext) continue;

      let score = 0;
      if (exact) score += 260;
      if (close) score += 130;
      if (r.left <= 280) score += 120;
      if (r.width >= 90 && r.width <= 230) score += 80;
      if (r.height >= 28 && r.height <= 72) score += 80;
      if (siblingContext) score += 220;
      if (antecedentesContext) score += 160;

      candidates.push({
        x: Math.round(r.left + (r.width / 2)),
        y: Math.round(r.top + (r.height / 2)),
        score,
        txt
      });
    }
    if (!candidates.length) return { ok: false };
    candidates.sort((a, b) => b.score - a.score);
    return { ok: true, x: candidates[0].x, y: candidates[0].y, score: candidates[0].score, txt: candidates[0].txt };
  });

  if (!target?.ok || !Number.isFinite(target.x) || !Number.isFinite(target.y)) return false;
  try {
    // Click robusto: hover + click + reintento rápido sobre mismo punto.
    for (let i = 0; i < 2; i += 1) {
      await page.mouse.move(target.x, target.y);
      await waitForTimeoutRaw(page, 32);
      await page.mouse.click(target.x, target.y, { delay: 22 });
      await waitForTimeoutRaw(page, 170);
      if (await isNotaMedicaViewActive(page)) {
        console.log(
          `NOTA_MEDICA_CLICK_OK via=left_menu_xy origin=${origin || '-'} elapsed=${elapsedBase}ms score=${target.score} try=${i + 1}`
        );
        return true;
      }
    }

    await page.evaluate(({ x, y }) => {
      const el = document.elementFromPoint(x, y);
      if (!(el instanceof HTMLElement)) return;
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
      el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
      el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      el.click();
    }, { x: target.x, y: target.y });
    await waitForTimeoutRaw(page, 180);
    if (await isNotaMedicaViewActive(page)) {
      console.log(
        `NOTA_MEDICA_CLICK_OK via=left_menu_xy_js origin=${origin || '-'} elapsed=${elapsedBase}ms score=${target.score}`
      );
      return true;
    }
  } catch {}
  return false;
}

async function openNotaMedicaFromSidebar(page, origin = '') {
  if (!AUTO_OPEN_NOTA_MEDICA_AFTER_MODULE) {
    console.log('NOTA_MEDICA_STEP_SKIP auto=0');
    return true;
  }

  if (await isNotaMedicaViewActive(page)) {
    console.log(`NOTA_MEDICA_ALREADY_ACTIVE origin=${origin || '-'}`);
    return true;
  }

  if (NOTA_MEDICA_DELAY_MS > 0) {
    await waitForTimeoutRaw(page, NOTA_MEDICA_DELAY_MS);
  }

  const started = Date.now();
  const endBy = started + NOTA_MEDICA_CLICK_TIMEOUT_MS;
  let attempts = 0;

  while (Date.now() < endBy) {
    attempts += 1;

    // Helper: verificar si Nota médica está activa (vista o form visible)
    const checkNotaActive = async () => {
      if (await isNotaMedicaViewActive(page)) return true;
      // Fallback: verificar si los campos del form de Nota médica están visibles
      try {
        return await page.evaluate(() => {
          const normalize = (s) => (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
          const labels = Array.from(document.querySelectorAll('label,span,div,p,strong,h2,h3,fieldset')).map(n => normalize(n.textContent || '')).filter(Boolean);
          const blob = labels.join(' ');
          return blob.includes('consulta por') && blob.includes('presente enfermedad');
        });
      } catch { return false; }
    };

    // Estrategia 1: click por ID exacto del botón "Nota médica"
    try {
      const btnLoc = page.locator('#btnIdMenuLate5MP_TableroMedico, [id$="btnIdMenuLate5MP_TableroMedico"]').first();
      if ((await btnLoc.count()) > 0 && (await btnLoc.isVisible())) {
        await btnLoc.click({ force: true, timeout: 2000 });
        await waitForTimeoutRaw(page, 400);
        if (await checkNotaActive()) {
          console.log(`NOTA_MEDICA_CLICK_OK via=btn_id origin=${origin || '-'} elapsed=${Date.now() - started}ms`);
          return true;
        }
      }
    } catch {}

    // Estrategia 2: click por estructura del sidebar (TabsHeader li[3])
    try {
      const tabsSelectors = [
        '#TabsHeader ul li:nth-child(3) span',
        '#TabsHeader ul li:nth-child(3) div',
        '#TabsHeader ul li:nth-child(3)'
      ];
      for (const sel of tabsSelectors) {
        const loc = page.locator(sel).first();
        if ((await loc.count()) === 0) continue;
        if (!(await loc.isVisible())) continue;
        const txt = await loc.textContent().catch(() => '');
        if (txt && /nota\s*m[eé]dica/i.test(txt)) {
          await loc.click({ force: true, timeout: 2000 });
          await waitForTimeoutRaw(page, 400);
          if (await checkNotaActive()) {
            console.log(`NOTA_MEDICA_CLICK_OK via=tabs_header_li3 origin=${origin || '-'} elapsed=${Date.now() - started}ms`);
            return true;
          }
        }
      }
    } catch {}

    // Estrategia 3: scoring por texto y posición
    if (await clickNotaMedicaInLeftMenu(page, origin, Date.now() - started)) return true;

    // Estrategia 4: click por texto exacto regex
    try {
      const target = page.locator('text=/^\\s*Nota\\s*M[eé]dica\\s*$/i').first();
      if ((await target.count()) > 0) {
        await target.click({ force: true, timeout: 900 });
        await waitForTimeoutRaw(page, 300);
        if (await checkNotaActive()) {
          console.log(`NOTA_MEDICA_CLICK_OK via=text-exact origin=${origin || '-'} elapsed=${Date.now() - started}ms`);
          return true;
        }
      }
    } catch {}

    if (attempts % 3 === 0) {
      const menu = await readAntecedentesMenuState(page);
      console.log(
        `NOTA_MEDICA_CLICK_RETRY attempts=${attempts} origin=${origin || '-'} elapsed=${Date.now() - started}ms active="${menu.activeLabel || '-'}"`
      );
    }
    await waitForTimeoutRaw(page, 200);
  }

  const menu = await readAntecedentesMenuState(page);
  console.log(
    `NOTA_MEDICA_CLICK_TIMEOUT origin=${origin || '-'} timeout=${NOTA_MEDICA_CLICK_TIMEOUT_MS}ms active="${menu.activeLabel || '-'}"`
  );
  return false;
}

async function resolveNotaFieldPoints(page, requestedKeys = []) {
  if (isPageClosedSafe(page)) return { fields: {}, generar: null };
  try {
    return await page.evaluate((requested) => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        const inViewport = r.bottom > 0 && r.top < window.innerHeight && r.right > 0 && r.left < window.innerWidth;
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 6 && r.height > 6 && inViewport;
      };
      const isEditable = (el) => {
        if (!(el instanceof HTMLElement) || !visible(el)) return false;
        if (el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true') return false;
        if (el.getAttribute('readonly') !== null || el.getAttribute('aria-readonly') === 'true') return false;
        if (el instanceof HTMLTextAreaElement) return true;
        if (el instanceof HTMLInputElement) {
          const t = (el.type || 'text').toLowerCase();
          return t === 'text' || t === 'search' || t === 'email' || t === 'tel' || t === 'url' || t === 'number';
        }
        return el.getAttribute('contenteditable') === 'true';
      };
      const rectDistance = (a, b) => {
        if (!a || !b) return 99999;
        const ax = a.left + a.width / 2;
        const ay = a.top + a.height / 2;
        const bx = b.left + b.width / 2;
        const by = b.top + b.height / 2;
        const dx = ax - bx;
        const dy = ay - by;
        return Math.sqrt(dx * dx + dy * dy);
      };
      const getClosestContainer = (el) => {
        if (!(el instanceof HTMLElement)) return null;
        return (
          el.closest('.form-group, .form-row, .row, tr, td, .k-form-field, .k-edit-field, .k-window-content, [class*="field"], [class*="group"], section, article') ||
          el.parentElement ||
          null
        );
      };
      const meta = (el) =>
        normalize(
          `${el?.id || ''} ${el?.getAttribute?.('name') || ''} ${el?.getAttribute?.('aria-label') || ''} ${el?.getAttribute?.('placeholder') || ''} ${el?.getAttribute?.('title') || ''}`
        );
      const defs = {
        consulta_por: { patterns: ['consulta por'], keywords: ['consulta'] },
        triage: { patterns: ['triage'], keywords: ['triage'] },
        presente_enfermedad: { patterns: ['presente enfermedad', 'enfermedad presente'], keywords: ['presente', 'enfermedad'] },
        apreciacion_diagnostica: { patterns: ['apreciacion diagnostica', 'apreciación diagnóstica'], keywords: ['apreciacion', 'diagnostic'] },
        diagnostico_principal: { patterns: ['diagnostico principal', 'diagnóstico principal'], keywords: ['diagnostico principal', 'diagnostic principal'] }
      };
      const wanted = Array.isArray(requested) && requested.length ? requested : Object.keys(defs);

      const nodes = Array.from(document.querySelectorAll('label, span, div, p, strong, td, th, li, h1, h2, h3'))
        .filter(visible)
        .map((n) => ({ n, t: normalize(n.textContent || '') }))
        .filter((x) => x.t && x.t.length >= 3 && x.t.length <= 130);
      const findFieldNode = (patterns) => {
        let best = null;
        for (const item of nodes) {
          const hits = patterns.filter((p) => item.t.includes(normalize(p))).length;
          if (!hits) continue;
          const score = hits * 100 - item.t.length;
          if (!best || score > best.score) best = { node: item.n, score };
        }
        return best ? best.node : null;
      };

      const allIframes = Array.from(document.querySelectorAll('iframe'));
      const editables = Array.from(
        document.querySelectorAll(
          'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], .k-input-inner, .k-input'
        )
      ).filter(isEditable);
      const editors = allIframes
        .map((el, frameIdx) => ({ el, frameIdx }))
        .filter((x) => visible(x.el));

      const fields = {};
      for (const key of wanted) {
        const def = defs[key];
        if (!def) continue;
        const node = findFieldNode(def.patterns);
        if (!(node instanceof HTMLElement)) {
          fields[key] = [];
          continue;
        }
        const nodeRect = node.getBoundingClientRect();
        const nodeContainer = getClosestContainer(node);
        const scoredEditable = editables
          .map((el) => {
            const r = el.getBoundingClientRect();
            let score = 0;
            for (const k of def.keywords) {
              if (meta(el).includes(normalize(k))) score += 120;
            }
            score += Math.max(0, 260 - rectDistance(nodeRect, r));
            if (r.left >= 220) score += 20;
            if (nodeContainer instanceof HTMLElement && nodeContainer.contains(el)) score += 120;
            if (key === 'presente_enfermedad') {
              if (el instanceof HTMLTextAreaElement) score += 220;
              if (el.getAttribute('contenteditable') === 'true') score += 160;
              if (r.height >= 70) score += 80;
            }
            if (key === 'apreciacion_diagnostica') {
              if (el instanceof HTMLTextAreaElement) score += 180;
              if (el.getAttribute('contenteditable') === 'true') score += 180;
              if (r.height >= 60) score += 60;
            }
            return {
              x: Math.round(r.left + r.width / 2),
              y: Math.round(r.top + Math.min(r.height / 2, 16)),
              score,
              kind: 'editable'
            };
          })
          .filter((x) => x.score > 20);

        const scoredEditors = editors
          .map(({ el, frameIdx }) => {
            const r = el.getBoundingClientRect();
            const m = meta(el);
            let score = 0;
            for (const k of def.keywords) {
              if (m.includes(normalize(k))) score += 180;
            }
            score += Math.max(0, 240 - rectDistance(nodeRect, r));
            if (nodeContainer instanceof HTMLElement && nodeContainer.contains(el)) score += 140;
            if (key === 'consulta_por') {
              if (m.includes('consulta')) score += 260;
            }
            if (key === 'apreciacion_diagnostica') {
              if (m.includes('apreciacion') || m.includes('diagnostic')) score += 260;
            }
            return {
              x: Math.round(r.left + r.width / 2),
              y: Math.round(r.top + r.height / 2),
              score,
              kind: 'iframe',
              frameIdx
            };
          })
          .filter((x) => x.score > 30);

        fields[key] = [...scoredEditable, ...scoredEditors]
          .sort((a, b) => b.score - a.score)
          .slice(0, 10);
      }

      // Botón Generar IA del lado derecho/debajo de "Presente enfermedad".
      const presenteNode = findFieldNode(defs.presente_enfermedad.patterns);
      const presenteRect = presenteNode instanceof HTMLElement ? presenteNode.getBoundingClientRect() : null;
      const generarButtons = Array.from(
        document.querySelectorAll('button, a, span, input[type="button"], input[type="submit"], [role="button"]')
      )
        .filter(visible)
        .map((b) => {
          const txt = normalize(`${b.textContent || ''} ${b.getAttribute('title') || ''} ${b.getAttribute('aria-label') || ''}`);
          if (!txt.includes('generar')) return null;
          const r = b.getBoundingClientRect();
          let score = 100;
          if (txt.includes('ia')) score += 90;
          if (presenteRect) {
            if (r.left >= (presenteRect.left + presenteRect.width * 0.35)) score += 220;
            if (r.top >= (presenteRect.top - 10)) score += 140;
            if (r.top <= (presenteRect.top + 360)) score += 80;
            score += Math.max(0, 260 - rectDistance(presenteRect, r));
          }
          return {
            x: Math.round(r.left + r.width / 2),
            y: Math.round(r.top + r.height / 2),
            score
          };
        })
        .filter(Boolean)
        .sort((a, b) => b.score - a.score);

      return {
        fields,
        generar: generarButtons[0] || null,
        generarCandidates: generarButtons.slice(0, 4)
      };
    }, requestedKeys);
  } catch {
    return { fields: {}, generar: null };
  }
}

async function scrollNotaMedicaToTop(page) {
  if (isPageClosedSafe(page)) return;
  try {
    await page.evaluate(() => {
      const nodes = Array.from(document.querySelectorAll('div,section,main,article'))
        .filter((n) => {
          const r = n.getBoundingClientRect();
          return r.width > 200 && r.height > 160;
        })
        .map((n) => {
          const st = getComputedStyle(n);
          const overflowY = st.overflowY || '';
          const canScroll = (overflowY.includes('auto') || overflowY.includes('scroll')) && n.scrollHeight > (n.clientHeight + 60);
          const txt = (n.textContent || '').toLowerCase();
          const score =
            (canScroll ? 300 : 0) +
            (txt.includes('consulta por') ? 150 : 0) +
            (txt.includes('apreciacion') ? 130 : 0) +
            (txt.includes('diagnostico principal') ? 110 : 0) +
            (txt.includes('triage') ? 90 : 0);
          return { n, score };
        })
        .filter((x) => x.score > 0)
        .sort((a, b) => b.score - a.score);

      if (nodes.length) {
        try { nodes[0].n.scrollTop = 0; } catch {}
      }
      try { window.scrollTo(0, 0); } catch {}
    });
  } catch {}
  await waitForTimeoutRaw(page, 120);
}

async function touchFieldHumanAtPoint(page, point) {
  if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y)) return false;
  try {
    await page.mouse.move(point.x, point.y);
    await waitForTimeoutRaw(page, 30);
    await page.mouse.click(point.x, point.y, { delay: 24 });
    await waitForTimeoutRaw(page, 35);
    await page.keyboard.press('End');
    await page.keyboard.press(' ');
    await page.keyboard.press('Backspace');
    await waitForTimeoutRaw(page, 20);
    await page.keyboard.press('Tab');
    await waitForTimeoutRaw(page, 80);
    return true;
  } catch {
    return false;
  }
}

async function ensureTrustedTouchOnNotaFields(page, keys = []) {
  await scrollNotaMedicaToTop(page);
  const plan = await resolveNotaFieldPoints(page, keys);
  let touched = 0;
  const total = Array.isArray(keys) ? keys.length : 0;
  for (const key of keys) {
    const candidates = Array.isArray(plan?.fields?.[key]) ? plan.fields[key] : [];
    let ok = false;
    for (let i = 0; i < candidates.length && i < 3 && !ok; i += 1) {
      ok = await touchFieldHumanAtPoint(page, candidates[i]);
    }
    if (ok) touched += 1;
  }
  return { touched, total };
}

async function keyboardCaptureAtPoint(page, point, value) {
  const captureInIframe = async (frameIdx) => {
    try {
      if (!Number.isInteger(frameIdx) || frameIdx < 0) return false;
      const frameLoc = page.locator('iframe').nth(frameIdx);
      if ((await frameLoc.count()) === 0) return false;
      await frameLoc.scrollIntoViewIfNeeded();
      await frameLoc.click({ force: true, timeout: 1000 });
      await waitForTimeoutRaw(page, 45);

      const handle = await frameLoc.elementHandle();
      const frame = await handle?.contentFrame();
      if (!frame) return false;

      let body = frame.locator('body').first();
      if ((await body.count()) === 0) {
        body = frame.locator('[contenteditable="true"]').first();
      }
      if ((await body.count()) === 0) return false;

      await body.click({ force: true, timeout: 1000 });
      await waitForTimeoutRaw(page, 40);
      try { await page.keyboard.press('Control+A'); } catch {}
      await waitForTimeoutRaw(page, 20);
      try { await page.keyboard.press('Backspace'); } catch {}
      await waitForTimeoutRaw(page, 20);
      const text = String(value || '');
      const head = text.slice(0, 42);
      const tail = text.slice(42);
      if (head) await page.keyboard.type(head, { delay: 7 });
      if (tail) await page.keyboard.insertText(tail);
      await waitForTimeoutRaw(page, 35);
      try { await page.keyboard.press('Tab'); } catch {}
      await waitForTimeoutRaw(page, 65);

      const ok = await frame.evaluate((txt) => {
        const norm = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const probe = norm(String(txt || '')).slice(0, 28);
        const bodyText = norm(document.body?.innerText || document.body?.textContent || '');
        return probe ? bodyText.includes(probe) : bodyText.length > 0;
      }, String(value || ''));
      return Boolean(ok);
    } catch {
      return false;
    }
  };

  if (point?.kind === 'iframe' && Number.isInteger(point?.frameIdx)) {
    return captureInIframe(point.frameIdx);
  }

  if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y)) return false;
  try {
    await page.mouse.move(point.x, point.y);
    await waitForTimeoutRaw(page, 28);
    await page.mouse.click(point.x, point.y, { delay: 20 });
    await waitForTimeoutRaw(page, 35);
    try { await page.keyboard.press('Control+A'); } catch {}
    await waitForTimeoutRaw(page, 16);
    try { await page.keyboard.press('Backspace'); } catch {}
    await waitForTimeoutRaw(page, 16);
    const text = String(value || '');
    const head = text.slice(0, 42);
    const tail = text.slice(42);
    if (head) await page.keyboard.type(head, { delay: 7 });
    if (tail) await page.keyboard.insertText(tail);
    await waitForTimeoutRaw(page, 18);
    try { await page.keyboard.press('Enter'); } catch {}
    await waitForTimeoutRaw(page, 30);
    try { await page.keyboard.press('Tab'); } catch {}
    await waitForTimeoutRaw(page, 65);
    return true;
  } catch {
    return false;
  }
}

async function captureNotaFieldsByKeyboard(page, keys = [], textValue = '') {
  await scrollNotaMedicaToTop(page);
  const plan = await resolveNotaFieldPoints(page, keys);
  let captured = 0;
  const total = Array.isArray(keys) ? keys.length : 0;
  for (const key of keys) {
    const candidates = Array.isArray(plan?.fields?.[key]) ? plan.fields[key] : [];
    // Skip capture if field already has data
    const hasExisting = await page.evaluate((fieldKey) => {
      const normalize = (s) => (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 6 && r.height > 6;
      };
      const defs = {
        consulta_por: ['consulta por'],
        triage: ['triage'],
        presente_enfermedad: ['presente enfermedad', 'enfermedad presente'],
        apreciacion_diagnostica: ['apreciacion diagnostica'],
        diagnostico_principal: ['diagnostico principal']
      };
      const patterns = defs[fieldKey] || [];
      if (!patterns.length) return false;
      const labels = Array.from(document.querySelectorAll('label,span,div,p,strong,td,th,li'))
        .filter(visible).map((n) => ({ n, t: normalize(n.textContent || '') }));
      let best = null;
      for (const item of labels) {
        const hits = patterns.filter((p) => item.t.includes(normalize(p))).length;
        if (!hits) continue;
        const score = hits * 100 - item.t.length;
        if (!best || score > best.score) best = { node: item.n, score };
      }
      if (!best) return false;
      const nearSel = 'textarea, input[type="text"], input:not([type]), [contenteditable="true"]';
      const container = best.node.closest('.form-group, .form-row, .row, tr, td, section') || best.node.parentElement;
      const controls = container ? Array.from(container.querySelectorAll(nearSel)).filter(visible) : [];
      for (const ctrl of controls) {
        let val = '';
        if (ctrl instanceof HTMLTextAreaElement || ctrl instanceof HTMLInputElement) val = normalize(ctrl.value || '');
        else if (ctrl.getAttribute('contenteditable') === 'true') val = normalize(ctrl.textContent || '');
        if (val && val.length >= 3) return true;
      }
      return false;
    }, key).catch(() => false);
    if (hasExisting) {
      captured += 1;
      continue;
    }
    let ok = false;
    for (let i = 0; i < candidates.length && i < 3 && !ok; i += 1) {
      ok = await keyboardCaptureAtPoint(page, candidates[i], textValue);
    }
    if (ok) captured += 1;
  }
  return { captured, total };
}

/**
 * Detecta y cierra el modal "Catálogo de diagnósticos" si está abierto.
 * Este modal se abre accidentalmente cuando el bot clickea el campo de búsqueda
 * de diagnóstico en vez del botón "Generar". Usa 2-phase: detectar con evaluate,
 * cerrar con Playwright locator para asegurar el postback Telerik.
 */
async function dismissCatalogoDiagnosticosModal(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    const modalInfo = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 20;
      };

      // Buscar ventanas Kendo/Telerik visibles
      const windows = Array.from(document.querySelectorAll('.k-window, [role="dialog"], .RadWindow'))
        .filter(visible);

      for (const win of windows) {
        const titleBar = win.querySelector('.k-window-titlebar, .rwTitlebar, .k-dialog-titlebar');
        if (!titleBar) continue;
        const titleText = normalize(titleBar.textContent || '');
        if (!titleText.includes('catalogo de diagnostico') && !titleText.includes('catálogo de diagnóstico')) continue;

        // Encontrado - buscar botón X de cierre
        const closeSelectors = [
          '.k-window-titlebar .k-window-titlebar-actions a',
          '.k-window-titlebar .k-window-titlebar-actions button',
          '.k-window-titlebar .k-window-action',
          '.k-window-titlebar [aria-label="Close"]',
          '.k-window-titlebar .k-i-close',
          '.k-window-titlebar .k-svg-icon.k-svg-i-x',
          '.rwTitlebar .rwCloseButton',
          'button[aria-label="Close"]',
          'a[title*="close" i]',
          'a[title*="cerrar" i]'
        ];

        let closeBtnId = null;
        for (const sel of closeSelectors) {
          const btn = win.querySelector(sel);
          if (btn instanceof HTMLElement && visible(btn)) {
            if (!btn.id) btn.setAttribute('data-catalogo-close-tmp', '1');
            closeBtnId = btn.id || null;
            return { found: true, closeBtnId, hasTmpAttr: !btn.id, title: titleText.slice(0, 40) };
          }
        }

        // Fallback: buscar cualquier botón/enlace con X o close
        const allBtns = Array.from(win.querySelectorAll('button, a, span')).filter(visible);
        const xBtn = allBtns.find((n) => {
          const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || '');
          return t === 'x' || t === '×' || t === '✕' || t.includes('close') || t.includes('cerrar');
        });
        if (xBtn) {
          if (!xBtn.id) xBtn.setAttribute('data-catalogo-close-tmp', '1');
          return { found: true, closeBtnId: xBtn.id || null, hasTmpAttr: !xBtn.id, title: titleText.slice(0, 40) };
        }

        return { found: true, closeBtnId: null, hasTmpAttr: false, title: titleText.slice(0, 40) };
      }
      return { found: false };
    });

    if (!modalInfo?.found) return false;

    console.log(`CATALOGO_DIAGNOSTICOS_MODAL_DETECTED title="${modalInfo.title || ''}"`);

    // Paso 2: cerrar con Playwright locator
    if (modalInfo.closeBtnId) {
      try {
        const loc = page.locator(`#${CSS.escape(modalInfo.closeBtnId)}`).first();
        if ((await loc.count()) > 0) {
          await loc.click({ force: true, timeout: 2000 });
          await waitForTimeoutRaw(page, 300);
          console.log('CATALOGO_DIAGNOSTICOS_MODAL_CLOSED via=locator_id');
          return true;
        }
      } catch {}
    }

    if (modalInfo.hasTmpAttr) {
      try {
        const loc = page.locator('[data-catalogo-close-tmp="1"]').first();
        if ((await loc.count()) > 0) {
          await loc.click({ force: true, timeout: 2000 });
          await waitForTimeoutRaw(page, 300);
          try { await page.evaluate(() => { document.querySelectorAll('[data-catalogo-close-tmp]').forEach(e => e.removeAttribute('data-catalogo-close-tmp')); }); } catch {}
          console.log('CATALOGO_DIAGNOSTICOS_MODAL_CLOSED via=locator_tmp_attr');
          return true;
        }
      } catch {}
    }

    // Fallback: selectores genéricos de cierre en ventana Kendo
    const fallbackSelectors = [
      '.k-window .k-window-titlebar .k-window-action',
      '.k-window .k-window-titlebar .k-i-close',
      '.k-window .k-window-titlebar [aria-label="Close"]'
    ];
    for (const sel of fallbackSelectors) {
      try {
        const loc = page.locator(sel).first();
        if ((await loc.count()) > 0 && (await loc.isVisible())) {
          await loc.click({ force: true, timeout: 2000 });
          await waitForTimeoutRaw(page, 300);
          console.log(`CATALOGO_DIAGNOSTICOS_MODAL_CLOSED via=locator_fallback sel="${sel}"`);
          return true;
        }
      } catch {}
    }

    // Último fallback: Escape
    try {
      await page.keyboard.press('Escape');
      await waitForTimeoutRaw(page, 300);
      console.log('CATALOGO_DIAGNOSTICOS_MODAL_CLOSED via=escape');
      return true;
    } catch {}

    console.log('CATALOGO_DIAGNOSTICOS_MODAL_CLOSE_FAIL');
    return false;
  } catch {
    return false;
  }
}

async function clickGenerarIaByHumanAction(page) {
  if (isPageClosedSafe(page)) return false;

  // Pre-check: cerrar modal "Catálogo de diagnósticos" si está abierto
  await dismissCatalogoDiagnosticosModal(page);

  await scrollNotaMedicaToTop(page);

  // Paso 1: detectar botones "Generar" y obtener el mejor candidato (solo detección, sin click)
  let bestBtnId = null;
  try {
    const detected = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 6 && r.height > 6;
      };

      const labels = Array.from(document.querySelectorAll('label,span,div,p,strong,td,th,li,h1,h2,h3'))
        .filter(visible)
        .map((n) => ({ n, t: normalize(n.textContent || '') }));
      const apreciacionLabel = labels.find((x) => x.t.includes('apreciacion diagnostica') || x.t.includes('apreciación diagnóstica'))?.n || null;
      const anchor = apreciacionLabel instanceof HTMLElement ? apreciacionLabel.getBoundingClientRect() : null;

      const nodes = Array.from(
        document.querySelectorAll('button,a,span,input[type="button"],input[type="submit"],[role="button"]')
      ).filter(visible);

      const candidates = [];
      for (const n of nodes) {
        const txt = normalize(
          `${n.textContent || ''} ${n.getAttribute?.('title') || ''} ${n.getAttribute?.('aria-label') || ''} ${n.id || ''} ${n.className || ''}`
        );
        if (!txt.includes('generar')) continue;
        const r = n.getBoundingClientRect();
        let score = 100;
        if (txt.includes('generar')) score += 200;
        if (anchor) {
          const dy = Math.abs(r.top - anchor.top);
          if (dy < 80) score += 300;
          else if (dy < 200) score += 150;
        }
        // Dar un id temporal para poder encontrarlo después con locator
        if (!n.id) n.setAttribute('data-generar-ia-tmp', '1');
        candidates.push({
          score,
          id: n.id || null,
          hasTmpAttr: !n.id,
          tag: n.tagName.toLowerCase(),
          txt: txt.slice(0, 40)
        });
      }
      if (!candidates.length) return { found: false };
      candidates.sort((a, b) => b.score - a.score);
      return { found: true, best: candidates[0] };
    });

    if (detected?.found && detected.best) {
      bestBtnId = detected.best.id;
    }
  } catch {}

  // Paso 2: click con Playwright locator (dispara postback Telerik correctamente)
  // Estrategia 2a: por ID si lo tenemos
  if (bestBtnId) {
    try {
      const loc = page.locator(`#${CSS.escape(bestBtnId)}`).first();
      if ((await loc.count()) > 0 && (await loc.isVisible())) {
        await loc.click({ force: true, timeout: 2000 });
        await waitForTimeoutRaw(page, 300);
        console.log(`GENERAR_IA_CLICK_OK via=locator_id id="${bestBtnId}"`);
        return true;
      }
    } catch {}
  }

  // Estrategia 2b: por atributo temporal
  try {
    const tmpLoc = page.locator('[data-generar-ia-tmp="1"]').first();
    if ((await tmpLoc.count()) > 0 && (await tmpLoc.isVisible())) {
      await tmpLoc.click({ force: true, timeout: 2000 });
      await waitForTimeoutRaw(page, 300);
      console.log('GENERAR_IA_CLICK_OK via=locator_tmp_attr');
      // Limpiar atributo temporal
      try { await page.evaluate(() => { document.querySelectorAll('[data-generar-ia-tmp]').forEach(e => e.removeAttribute('data-generar-ia-tmp')); }); } catch {}
      return true;
    }
  } catch {}

  // Estrategia 2c: primer botón "Generar" visible con Playwright locator
  try {
    const loc = page.locator('button:has-text("Generar"), a:has-text("Generar"), [role="button"]:has-text("Generar")').first();
    if ((await loc.count()) > 0 && (await loc.isVisible())) {
      await loc.click({ force: true, timeout: 2000 });
      await waitForTimeoutRaw(page, 300);
      console.log('GENERAR_IA_CLICK_OK via=locator_text');
      return true;
    }
  } catch {}

  // Fallback por coordenadas calculadas
  const plan = await resolveNotaFieldPoints(page, ['presente_enfermedad', 'apreciacion_diagnostica']);
  const candidates = [];
  if (plan?.generar) candidates.push(plan.generar);
  if (Array.isArray(plan?.generarCandidates)) {
    for (const c of plan.generarCandidates) candidates.push(c);
  }
  for (let i = 0; i < candidates.length && i < 5; i += 1) {
    const c = candidates[i];
    if (!c || !Number.isFinite(c.x) || !Number.isFinite(c.y)) continue;
    try {
      await page.mouse.move(c.x, c.y);
      await waitForTimeoutRaw(page, 30);
      await page.mouse.click(c.x, c.y, { delay: 26 });
      await waitForTimeoutRaw(page, 300);
      console.log(`GENERAR_IA_CLICK_OK via=coordinates x=${c.x} y=${c.y}`);
      return true;
    } catch {}
  }

  console.log('GENERAR_IA_CLICK_FAIL');
  return false;
}

async function hasNotaRequiredFieldAlerts(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 4 && r.height > 4;
      };
      const nodes = Array.from(
        document.querySelectorAll('.k-tooltip-validation,.field-validation-error,.validation-summary-errors,.error,.alert,.k-notification,[role="alert"],span,div,p,li')
      ).filter(visible);
      const texts = nodes
        .map((n) => normalize(n.textContent || ''))
        .filter(Boolean)
        .slice(0, 500);
      return texts.some((t) => {
        const captureMsg = t.includes('no se ha capturado') || t.includes('no se capturo');
        const generic = t.includes('obligatorio') || t.includes('requerido') || t.includes('debe ingresar') || t.includes('complete') || t.includes('campo');
        const nota =
          t.includes('consulta') ||
          t.includes('triage') ||
          t.includes('presente enfermedad') ||
          t.includes('apreciacion diagnostica') ||
          t.includes('diagnostico principal');
        if (captureMsg && nota) return true;
        return generic && nota;
      });
    });
  } catch {
    return false;
  }
}

async function isRecetaModalVisible(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 20;
      };
      const dialogs = Array.from(
        document.querySelectorAll(
          '.k-window, .k-window-content, .modal, .modal-content, [role="dialog"], .swal2-popup, .ui-dialog'
        )
      ).filter(visible);
      const texts = dialogs
        .map((n) => normalize(n.textContent || ''))
        .filter(Boolean)
        .slice(0, 120);
      return texts.some((t) => t.includes('receta')) || texts.some((t) => t.includes('prescrip'));
    });
  } catch {
    return false;
  }
}

async function readTopDialogRecetaInfo(page) {
  if (isPageClosedSafe(page)) return { hasDialog: false, title: '', text: '', isReceta: false, isCatalogo: false };
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 80;
      };
      const asNum = (raw) => {
        const n = Number.parseInt(String(raw || ''), 10);
        return Number.isFinite(n) ? n : -999999;
      };

      const dialogs = Array.from(
        document.querySelectorAll(
          '.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]'
        )
      )
        .filter(visible)
        .map((d, idx) => {
          const title = normalize(
            d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || ''
          );
          const text = normalize(d.textContent || '');
          return { idx, title, text, z: asNum(getComputedStyle(d).zIndex) };
        });

      if (!dialogs.length) return { hasDialog: false, title: '', text: '', isReceta: false, isCatalogo: false };

      dialogs.sort((a, b) => {
        if (a.z !== b.z) return a.z - b.z;
        return a.idx - b.idx;
      });
      const top = dialogs[dialogs.length - 1];
      const blob = `${top.title} ${top.text}`;
      const isReceta = blob.includes('receta') || blob.includes('prescrip');
      const isCatalogo = blob.includes('catalogo de pacientes');
      return { hasDialog: true, title: top.title || '', text: top.text || '', isReceta, isCatalogo };
    });
  } catch {
    return { hasDialog: false, title: '', text: '', isReceta: false, isCatalogo: false };
  }
}

async function closeTopNonRecetaDialog(page, reason = '') {
  if (isPageClosedSafe(page)) return false;
  const info = await readTopDialogRecetaInfo(page);
  if (!info?.hasDialog || info?.isReceta) return false;

  if (info?.isCatalogo) {
    const closedCatalog = await closeCatalogPacientesModal(page);
    if (closedCatalog) {
      console.log(`RECETA_WRONG_MODAL_CLOSED type=catalogo reason=${reason || '-'}`);
      return true;
    }
  }

  let closed = false;
  try {
    closed = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };
      const visibleDialog = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 80;
      };
      const asNum = (raw) => {
        const n = Number.parseInt(String(raw || ''), 10);
        return Number.isFinite(n) ? n : -999999;
      };
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        try {
          el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          el.click();
          return true;
        } catch {
          return false;
        }
      };

      const dialogs = Array.from(
        document.querySelectorAll(
          '.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]'
        )
      )
        .filter(visibleDialog)
        .map((d, idx) => ({ d, idx, z: asNum(getComputedStyle(d).zIndex) }));
      if (!dialogs.length) return false;

      dialogs.sort((a, b) => {
        if (a.z !== b.z) return a.z - b.z;
        return a.idx - b.idx;
      });
      const topDialog = dialogs[dialogs.length - 1].d;

      const closeSelectors = [
        '.k-window-titlebar .k-window-titlebar-actions a',
        '.k-window-titlebar .k-window-titlebar-actions button',
        '.k-window-titlebar .k-window-action',
        '.k-window-titlebar [aria-label="Close"]',
        '.k-window-titlebar .k-i-close',
        '.k-window-titlebar .k-svg-icon.k-svg-i-x, .k-window-titlebar .k-svg-i-x',
        '.k-dialog-titlebar .k-window-action',
        '.rwTitlebar .rwCloseButton',
        'button[aria-label="Close"]',
        'button[title*="cerrar" i]',
        'button[title*="close" i]',
        'a[title*="cerrar" i]',
        'a[title*="close" i]'
      ];

      let closeBtn = null;
      for (const sel of closeSelectors) {
        const node = topDialog.querySelector(sel);
        if (node instanceof HTMLElement && visible(node)) {
          closeBtn = node;
          break;
        }
      }

      if (!closeBtn) {
        const allButtons = Array.from(topDialog.querySelectorAll('button, a, span, div')).filter(visible);
        closeBtn =
          allButtons.find((n) => {
            const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || '');
            return t === 'x' || t.includes('cerrar') || t.includes('close') || t.includes('cancelar');
          }) || null;
      }
      return safeClick(closeBtn);
    });
  } catch {}

  if (!closed) {
    try {
      await page.keyboard.press('Escape');
      closed = true;
    } catch {}
  }
  if (closed) {
    await waitForTimeoutRaw(page, 160);
    console.log(`RECETA_WRONG_MODAL_CLOSED type=generic reason=${reason || '-'} title="${info?.title || ''}"`);
  }
  return closed;
}

async function clickRecetaButton(page) {
  const directSelectors = [
    { label: 'role_button_receta', selector: 'button:has-text("Receta"), [role="button"]:has-text("Receta")' },
    { label: 'role_link_receta', selector: 'a:has-text("Receta"), a[title*="receta" i], a[aria-label*="receta" i]' },
    { label: 'text_receta', selector: 'button:has-text("Receta médica"), button:has-text("Receta medica"), span:has-text("Receta"), div:has-text("Receta"), li:has-text("Receta")' },
    { label: 'attr_receta', selector: '[title*="receta" i],[aria-label*="receta" i],[id*="receta" i],[name*="receta" i],[data-title*="receta" i]' },
    { label: 'attr_prescrip', selector: '[title*="prescrip" i],[aria-label*="prescrip" i],[id*="prescrip" i],[name*="prescrip" i],[data-title*="prescrip" i]' },
    { label: 'attr_medicamento', selector: '[title*="medicamento" i],[aria-label*="medicamento" i],[id*="medicamento" i],[name*="medicamento" i],[data-title*="medicamento" i]' }
  ];

  const tryClickBySelector = async ({ label, selector }) => {
    try {
      const loc = page.locator(selector);
      const count = Math.min(await loc.count(), 12);
      for (let i = 0; i < count; i += 1) {
        const item = loc.nth(i);
        let visible = false;
        try {
          visible = await item.isVisible();
        } catch {}
        if (!visible) continue;
        try {
          await item.scrollIntoViewIfNeeded();
          await item.click({ force: true, timeout: 1200 });
          await waitForTimeoutRaw(page, 220);
          if (await isRecetaModalVisible(page)) {
            console.log(`RECETA_CLICK_OK via=selector:${label}:idx${i + 1}`);
            return true;
          }
          await closeTopNonRecetaDialog(page, `selector:${label}:idx${i + 1}`);
        } catch {}
      }
    } catch {}
    return false;
  };

  for (const selectorSpec of directSelectors) {
    if (await tryClickBySelector(selectorSpec)) return true;
  }

  const getToolbarRecetaCandidates = async () => {
    if (isPageClosedSafe(page)) return [];
    try {
      return await page.evaluate(() => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const visible = (el) => {
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
        };
        const isLikelyClickable = (el) => {
          if (!(el instanceof HTMLElement)) return false;
          const role = normalize(el.getAttribute('role') || '');
          const hasClick = typeof el.onclick === 'function' || el.hasAttribute('onclick');
          const tabIdx = Number.parseInt(String(el.getAttribute('tabindex') || ''), 10);
          const cls = normalize(el.className || '');
          const cursor = normalize(getComputedStyle(el).cursor || '');
          return (
            role === 'button' ||
            hasClick ||
            Number.isFinite(tabIdx) ||
            cursor === 'pointer' ||
            cls.includes('btn') ||
            cls.includes('button') ||
            cls.includes('icon') ||
            cls.includes('k-icon') ||
            cls.includes('fa')
          );
        };

        const nodes = Array.from(
          document.querySelectorAll('button,a,span,div,li,i,[role="button"],[title],[aria-label],[onclick],[tabindex]')
        ).filter(visible);
        const anchorCandidates = nodes
          .map((n) => ({ n, t: normalize(`${n.textContent || ''} ${n.getAttribute?.('title') || ''} ${n.getAttribute?.('aria-label') || ''}`) }))
          .filter((x) => x.t.includes('abrir videollamada') || x.t.includes('videollamada'));
        const anchor = anchorCandidates[0]?.n || null;
        const anchorRect = anchor instanceof HTMLElement ? anchor.getBoundingClientRect() : null;

        const out = [];
        for (const n of nodes) {
          if (!isLikelyClickable(n)) continue;
          const r = n.getBoundingClientRect();
          const txt = normalize(`${n.textContent || ''} ${n.getAttribute?.('title') || ''} ${n.getAttribute?.('aria-label') || ''} ${n.id || ''} ${n.getAttribute?.('name') || ''}`);
          const cls = normalize(n.className || '');
          let score = 0;

          if (txt.includes('receta')) score += 320;
          if (txt.includes('prescrip')) score += 280;
          if (txt.includes('medicamento')) score += 200;
          if (cls.includes('receta')) score += 170;
          if (cls.includes('prescrip')) score += 130;
          if (txt.includes('laboratorio') || txt.includes('imagenolog') || txt.includes('ordenes')) score -= 120;
          if (txt.includes('abrir videollamada') || txt.includes('videollamada')) score -= 380;
          if (txt.includes('nota medica') || txt.includes('estudios complementarios')) score -= 90;
          if (r.top <= 0 || r.top > 650) score -= 90;

          if (anchorRect) {
            const sameRow = Math.abs((r.top + r.height / 2) - (anchorRect.top + anchorRect.height / 2)) <= 70;
            const rightSide = r.left >= (anchorRect.right - 40);
            const near = Math.abs(r.left - anchorRect.right) <= 680;
            if (sameRow) score += 140;
            if (rightSide) score += 120;
            if (near) score += 40;
            if (!sameRow) score -= 80;
            if (!rightSide) score -= 65;
          }

          if (score <= 20) continue;
          out.push({
            x: Math.round(r.left + r.width / 2),
            y: Math.round(r.top + r.height / 2),
            score,
            txt
          });
        }
        return out.sort((a, b) => b.score - a.score).slice(0, 16);
      });
    } catch {
      return [];
    }
  };

  const readTooltipHint = async () => {
    if (isPageClosedSafe(page)) return '';
    try {
      return await page.evaluate(() => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const visible = (el) => {
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
        };
        const tips = Array.from(document.querySelectorAll('[role="tooltip"],.k-tooltip,.tooltip,.popover,.k-animation-container'))
          .filter(visible)
          .map((n) => normalize(n.textContent || ''))
          .filter(Boolean);
        return tips.join(' | ');
      });
    } catch {
      return '';
    }
  };

  const toolbarCandidates = await getToolbarRecetaCandidates();
  for (const c of toolbarCandidates) {
    if (!Number.isFinite(c.x) || !Number.isFinite(c.y)) continue;
    try {
      await page.mouse.move(c.x, c.y, { steps: 4 });
      await waitForTimeoutRaw(page, 320);
      const tip = await readTooltipHint();
      const looksReceta = /receta|prescrip|medicamento/.test(tip || '');
      if (looksReceta) {
        await page.mouse.click(c.x, c.y, { delay: 24 });
        await waitForTimeoutRaw(page, 220);
        if (await isRecetaModalVisible(page)) {
          console.log(`RECETA_CLICK_OK via=tooltip score=${c.score || 0} txt="${String(c.txt || '').slice(0, 80)}"`);
          return true;
        }
        await closeTopNonRecetaDialog(page, `tooltip_match score=${c.score || 0}`);
      }
    } catch {}
  }

  for (const c of toolbarCandidates.slice(0, 12)) {
    if (!Number.isFinite(c.x) || !Number.isFinite(c.y)) continue;
    try {
      await page.mouse.move(c.x, c.y, { steps: 2 });
      await waitForTimeoutRaw(page, 90);
      await page.mouse.click(c.x, c.y, { delay: 22 });
      await waitForTimeoutRaw(page, 220);
      if (await isRecetaModalVisible(page)) {
        console.log(`RECETA_CLICK_OK via=toolbar_xy score=${c.score || 0} txt="${String(c.txt || '').slice(0, 80)}"`);
        return true;
      }
      await closeTopNonRecetaDialog(page, `toolbar_xy score=${c.score || 0}`);
    } catch {}
  }

  try {
    const clickPoints = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };
      const nodes = Array.from(
        document.querySelectorAll('button,a,span,div,li,i,[role="button"],[title],[aria-label],[onclick],[tabindex]')
      ).filter(visible);
      const anchor = nodes.find((n) => {
        const t = normalize(`${n.textContent || ''} ${n.getAttribute?.('title') || ''} ${n.getAttribute?.('aria-label') || ''}`);
        return t.includes('abrir videollamada') || t.includes('videollamada');
      }) || null;
      const ar = anchor instanceof HTMLElement ? anchor.getBoundingClientRect() : null;

      const scored = [];
      for (const n of nodes) {
        const txt = normalize(
          `${n.textContent || ''} ${n.getAttribute?.('title') || ''} ${n.getAttribute?.('aria-label') || ''} ${n.id || ''} ${n.getAttribute?.('name') || ''} ${n.className || ''}`
        );
        if (!txt) continue;
        let score = 0;
        if (txt.includes('receta')) score += 260;
        if (txt.includes('prescrip')) score += 220;
        if (txt.includes('medicamento')) score += 180;
        if (txt.includes('laboratorio') || txt.includes('imagenologia') || txt.includes('ordenes')) score -= 130;
        if (txt.includes('videollamada')) score -= 260;

        const r = n.getBoundingClientRect();
        if (ar) {
          const sameRow = Math.abs((r.top + r.height / 2) - (ar.top + ar.height / 2)) <= 70;
          const rightSide = r.left >= (ar.right - 30);
          if (sameRow) score += 120;
          if (rightSide) score += 95;
          if (!sameRow) score -= 80;
        }
        if (r.top <= 0 || r.top > 650) score -= 90;
        if (score <= 20) continue;

        scored.push({
          x: Math.round(r.left + r.width / 2),
          y: Math.round(r.top + r.height / 2),
          score,
          txt
        });
      }
      return scored.sort((a, b) => b.score - a.score).slice(0, 12);
    });

    for (const p of clickPoints) {
      if (!Number.isFinite(p?.x) || !Number.isFinite(p?.y)) continue;
      try {
        await page.mouse.move(p.x, p.y, { steps: 2 });
        await waitForTimeoutRaw(page, 80);
        await page.mouse.click(p.x, p.y, { delay: 18 });
        await waitForTimeoutRaw(page, 220);
        if (await isRecetaModalVisible(page)) {
          console.log(`RECETA_CLICK_OK via=fallback_xy score=${p.score || 0} txt="${String(p.txt || '').slice(0, 80)}"`);
          return true;
        }
        await closeTopNonRecetaDialog(page, `fallback_xy score=${p.score || 0}`);
      } catch {}
    }
  } catch {}

  try {
    const clicked = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        try {
          el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          el.click();
          return true;
        } catch {
          return false;
        }
      };
      const nodes = Array.from(
        document.querySelectorAll('button,a,span,div,li,i,[role="button"],[title],[aria-label]')
      ).filter(visible);
      const scored = [];
      for (const n of nodes) {
        const txt = normalize(
          `${n.textContent || ''} ${n.getAttribute?.('title') || ''} ${n.getAttribute?.('aria-label') || ''} ${n.id || ''} ${n.getAttribute?.('name') || ''} ${n.className || ''}`
        );
        if (!txt) continue;
        let score = 0;
        if (txt.includes('receta')) score += 220;
        if (txt.includes('prescrip')) score += 170;
        if (txt.includes('medicamento')) score += 130;
        if (txt.includes('laboratorio') || txt.includes('imagenologia') || txt.includes('videollamada')) score -= 150;
        if (score <= 20) continue;
        scored.push({ el: n, score });
      }
      if (!scored.length) return false;
      scored.sort((a, b) => b.score - a.score);
      return safeClick(scored[0].el);
    });
    if (clicked) {
      await waitForTimeoutRaw(page, 220);
      if (await isRecetaModalVisible(page)) {
        console.log('RECETA_CLICK_OK via=fallback_dom');
        return true;
      }
      await closeTopNonRecetaDialog(page, 'fallback_dom');
    }
  } catch {}

  return false;
}

async function generarReceta(page, origin = '') {
  if (!AUTO_GENERAR_RECETA_AFTER_IA) {
    console.log('RECETA_STEP_SKIP auto=0');
    return true;
  }
  if (isPageClosedSafe(page)) return false;

  console.log(`RECETA_STEP_START origin=${origin || '-'} wait_ms=${RECETA_AFTER_IA_WAIT_MS}`);
  await waitForTimeoutRaw(page, RECETA_AFTER_IA_WAIT_MS);

  const started = Date.now();
  const endBy = started + RECETA_CLICK_TIMEOUT_MS;
  let tries = 0;
  while (Date.now() < endBy) {
    tries += 1;
    if (await isRecetaModalVisible(page)) {
      console.log(`RECETA_MODAL_OK origin=${origin || '-'} tries=${tries} elapsed=${Date.now() - started}ms`);
      return true;
    }
    const clicked = await clickRecetaButton(page);
    if (clicked) {
      await waitForTimeoutRaw(page, 300);
      if (await isRecetaModalVisible(page)) {
        console.log(`RECETA_MODAL_OK origin=${origin || '-'} tries=${tries} elapsed=${Date.now() - started}ms`);
        return true;
      }
    }
    await waitForTimeoutRaw(page, 220);
  }
  console.log(`RECETA_MODAL_TIMEOUT origin=${origin || '-'} timeout=${RECETA_CLICK_TIMEOUT_MS}ms`);
  return false;
}

async function fillNotaMedicaAntecedentesAndGenerateIA(page, origin = '') {
  if (!AUTO_FILL_NOTA_MEDICA_FIELDS) {
    console.log('NOTA_MEDICA_FIELDS_FILL_SKIP auto=0');
    return true;
  }
  if (isPageClosedSafe(page)) return false;

  const started = Date.now();
  const endBy = started + NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS;
  let attempts = 0;
  let lastResult = null;

  console.log(
    `NOTA_MEDICA_FIELDS_FILL_START origin=${origin || '-'} click_generar=${AUTO_CLICK_GENERAR_IA_NOTA_MEDICA ? 1 : 0}`
  );
  await updateBotStatusOverlay(page, 'working', 'llenando campos Nota médica...');

  while (Date.now() < endBy) {
    attempts += 1;
    await scrollNotaMedicaToTop(page);

    const result = await page.evaluate(({ textValue, clickGenerar }) => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 6 && r.height > 6;
      };
      const isEditable = (el) => {
        if (!(el instanceof HTMLElement) || !visible(el)) return false;
        if (el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true') return false;
        if (el.getAttribute('readonly') !== null || el.getAttribute('aria-readonly') === 'true') return false;
        if (el instanceof HTMLTextAreaElement) return true;
        if (el instanceof HTMLInputElement) {
          const t = (el.type || 'text').toLowerCase();
          return t === 'text' || t === 'search' || t === 'email' || t === 'tel' || t === 'url' || t === 'number';
        }
        return el.getAttribute('contenteditable') === 'true';
      };
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        try {
          el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          el.click();
          return true;
        } catch {
          return false;
        }
      };
      const readValue = (el) => {
        if (!(el instanceof HTMLElement)) return '';
        try {
          if (el instanceof HTMLTextAreaElement || el instanceof HTMLInputElement) return normalize(el.value || '');
          if (el.getAttribute('contenteditable') === 'true') return normalize(el.textContent || '');
        } catch {}
        return '';
      };
      const setControlValue = (el, value) => {
        const readValue = (target) => {
          if (!(target instanceof HTMLElement)) return '';
          try {
            if (target instanceof HTMLTextAreaElement || target instanceof HTMLInputElement) {
              return normalize(target.value || '');
            }
            if (target.getAttribute('contenteditable') === 'true') {
              return normalize(target.textContent || '');
            }
          } catch {}
          return '';
        };
        const fireKey = (target, type, key) => {
          try {
            target.dispatchEvent(
              new KeyboardEvent(type, {
                bubbles: true,
                cancelable: true,
                key,
                code: key === ' ' ? 'Space' : (key === 'Backspace' ? 'Backspace' : 'KeyA')
              })
            );
          } catch {}
        };
        const fireInput = (target, data, inputType) => {
          try {
            target.dispatchEvent(
              new InputEvent('input', {
                bubbles: true,
                cancelable: true,
                data,
                inputType
              })
            );
          } catch {
            try {
              target.dispatchEvent(new Event('input', { bubbles: true }));
            } catch {}
          }
        };
        const commitAsTypedInput = (target, finalValue) => {
          if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) return;
          const proto = target instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
          const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
          const setVal = (v) => {
            if (setter) setter.call(target, v);
            else target.value = v;
          };
          const base = finalValue || '';
          try {
            target.focus?.();
            target.selectionStart = base.length;
            target.selectionEnd = base.length;
          } catch {}
          // Simula una edición real: espacio + backspace.
          fireKey(target, 'keydown', ' ');
          fireKey(target, 'keypress', ' ');
          setVal(`${base} `);
          fireInput(target, ' ', 'insertText');
          fireKey(target, 'keyup', ' ');
          fireKey(target, 'keydown', 'Backspace');
          setVal(base);
          fireInput(target, null, 'deleteContentBackward');
          fireKey(target, 'keyup', 'Backspace');
        };
        const commitAsTypedContentEditable = (target, finalValue) => {
          if (!(target instanceof HTMLElement)) return;
          if (target.getAttribute('contenteditable') !== 'true') return;
          const base = finalValue || '';
          try {
            target.focus?.();
          } catch {}
          fireKey(target, 'keydown', ' ');
          fireKey(target, 'keypress', ' ');
          target.textContent = `${base} `;
          fireInput(target, ' ', 'insertText');
          fireKey(target, 'keyup', ' ');
          fireKey(target, 'keydown', 'Backspace');
          target.textContent = base;
          fireInput(target, null, 'deleteContentBackward');
          fireKey(target, 'keyup', 'Backspace');
        };
        if (!(el instanceof HTMLElement)) return false;
        try {
          const expected = normalize(value);
          el.focus?.();
          if (el instanceof HTMLTextAreaElement || el instanceof HTMLInputElement) {
            const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
            const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
            if (setter) setter.call(el, value);
            else el.value = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            commitAsTypedInput(el, value);
            el.dispatchEvent(new Event('blur', { bubbles: true }));
            const got = readValue(el);
            if (!got) return false;
            const probe = expected.slice(0, Math.min(42, expected.length));
            return probe ? got.includes(probe) : got.length > 0;
          }
          if (el.getAttribute('contenteditable') === 'true') {
            el.textContent = value;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            commitAsTypedContentEditable(el, value);
            el.dispatchEvent(new Event('blur', { bubbles: true }));
            const got = readValue(el);
            if (!got) return false;
            const probe = expected.slice(0, Math.min(42, expected.length));
            return probe ? got.includes(probe) : got.length > 0;
          }
        } catch {}
        return false;
      };
      const getClosestContainer = (el) => {
        if (!(el instanceof HTMLElement)) return null;
        return (
          el.closest('.form-group, .form-row, .row, tr, td, .k-form-field, .k-edit-field, .k-window-content, [class*="field"], [class*="group"], section, article') ||
          el.parentElement ||
          null
        );
      };
      const dedupe = (arr) => {
        const out = [];
        const seen = new Set();
        for (const el of arr) {
          if (!(el instanceof HTMLElement)) continue;
          if (seen.has(el)) continue;
          seen.add(el);
          out.push(el);
        }
        return out;
      };
      const findFieldNode = (patterns) => {
        const nodes = Array.from(document.querySelectorAll('label, span, div, p, strong, td, th, li'))
          .filter(visible)
          .map((n) => ({ n, t: normalize(n.textContent || '') }))
          .filter((x) => x.t && x.t.length >= 3 && x.t.length <= 120);
        let best = null;
        for (const item of nodes) {
          const hits = patterns.filter((p) => item.t.includes(p)).length;
          if (!hits) continue;
          const score = hits * 100 - item.t.length;
          if (!best || score > best.score) best = { node: item.n, score, text: item.t };
        }
        return best ? best.node : null;
      };
      const candidateControlsFromFieldNode = (node, keywords = []) => {
        const candidates = [];
        if (!(node instanceof HTMLElement)) return candidates;
        const htmlFor = node.getAttribute('for');
        if (htmlFor) {
          const direct = document.getElementById(htmlFor);
          if (direct instanceof HTMLElement) candidates.push(direct);
        }

        const nearSelectors =
          'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], .k-input-inner, .k-input';
        const near1 = node.nextElementSibling;
        if (near1 instanceof HTMLElement) {
          candidates.push(...Array.from(near1.querySelectorAll(nearSelectors)));
          candidates.push(near1);
        }

        const container = getClosestContainer(node);
        if (container) {
          candidates.push(...Array.from(container.querySelectorAll(nearSelectors)));
        }

        let p = node.parentElement;
        for (let i = 0; i < 3 && p; i += 1) {
          candidates.push(...Array.from(p.querySelectorAll(nearSelectors)));
          p = p.parentElement;
        }

        if (keywords.length) {
          const globalInputs = Array.from(
            document.querySelectorAll(
              'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], [id], [name], [aria-label], [placeholder], [title]'
            )
          );
          for (const g of globalInputs) {
            if (!(g instanceof HTMLElement)) continue;
            const meta = normalize(
              `${g.id || ''} ${g.getAttribute('name') || ''} ${g.getAttribute('aria-label') || ''} ${g.getAttribute('placeholder') || ''} ${g.getAttribute('title') || ''}`
            );
            if (keywords.some((k) => meta.includes(k))) candidates.push(g);
          }
        }

        return dedupe(candidates);
      };
      const rectDistance = (a, b) => {
        if (!a || !b) return 99999;
        const ax = a.left + a.width / 2;
        const ay = a.top + a.height / 2;
        const bx = b.left + b.width / 2;
        const by = b.top + b.height / 2;
        const dx = ax - bx;
        const dy = ay - by;
        return Math.sqrt(dx * dx + dy * dy);
      };
      const nearestGenerarDistance = (ctrl) => {
        if (!(ctrl instanceof HTMLElement)) return 99999;
        const ctrlRect = ctrl.getBoundingClientRect();
        const btns = Array.from(document.querySelectorAll('button,a,span,[role="button"],input[type="button"],input[type="submit"]'))
          .filter(visible)
          .map((b) => {
            const txt = normalize(`${b.textContent || ''} ${b.getAttribute?.('title') || ''} ${b.getAttribute?.('aria-label') || ''}`);
            if (!(txt.includes('generar') || txt.includes(' ia ') || txt.endsWith(' ia') || txt.includes('ia'))) return null;
            return b.getBoundingClientRect();
          })
          .filter(Boolean);
        if (!btns.length) return 99999;
        let best = 99999;
        for (const br of btns) {
          const d = rectDistance(ctrlRect, br);
          if (d < best) best = d;
        }
        return best;
      };
      const controlMeta = (el) =>
        normalize(
          `${el?.id || ''} ${el?.getAttribute?.('name') || ''} ${el?.getAttribute?.('aria-label') || ''} ${el?.getAttribute?.('placeholder') || ''} ${el?.getAttribute?.('title') || ''}`
        );
      const scoreControlForTarget = (ctrl, node, targetDef) => {
        if (!(ctrl instanceof HTMLElement)) return -99999;
        const r = ctrl.getBoundingClientRect();
        if (!visible(ctrl)) return -99999;
        let score = 0;
        const meta = controlMeta(ctrl);
        for (const k of targetDef.keywords.map(normalize)) {
          if (meta.includes(k)) score += 140;
        }
        const nodeRect = node instanceof HTMLElement ? node.getBoundingClientRect() : null;
        const d = rectDistance(nodeRect, r);
        score += Math.max(0, 220 - d);
        if (r.left >= 220) score += 20;
        if (r.width >= 140) score += 18;
        if (ctrl instanceof HTMLTextAreaElement) score += 60;
        if (ctrl.getAttribute('contenteditable') === 'true') score += 50;
        if (targetDef.key === 'presente_enfermedad') {
          if (ctrl instanceof HTMLTextAreaElement) score += 220;
          if (ctrl.getAttribute('contenteditable') === 'true') score += 160;
          if (r.height >= 70) score += 80;
          if (meta.includes('presente') || meta.includes('enfermedad')) score += 220;
          if (r.height < 28) score -= 200;
        }
        if (targetDef.key === 'apreciacion_diagnostica') {
          if (ctrl instanceof HTMLTextAreaElement) score += 180;
          if (ctrl.getAttribute('contenteditable') === 'true') score += 180;
          if (r.height >= 60) score += 55;
          if (meta.includes('apreciacion') || meta.includes('diagnostic')) score += 250;
          const dg = nearestGenerarDistance(ctrl);
          score += Math.max(0, 240 - dg);
          if (r.height < 26) score -= 120;
        }
        return score;
      };
      const pickBestControl = (controls, node, targetDef, allowUsed = false) => {
        const pool = (controls || [])
          .filter((c) => c instanceof HTMLElement)
          .filter(isEditable)
          .filter((c) => allowUsed || !usedControls.has(c))
          .map((c) => ({ c, score: scoreControlForTarget(c, node, targetDef) }))
          .sort((a, b) => b.score - a.score);
        return pool[0]?.c || null;
      };
      const targets = [
        { key: 'consulta_por', patterns: ['consulta por'], keywords: ['consulta'] },
        { key: 'triage', patterns: ['triage'], keywords: ['triage'] },
        { key: 'presente_enfermedad', patterns: ['presente enfermedad', 'enfermedad presente'], keywords: ['presente', 'enfermedad'] },
        {
          key: 'apreciacion_diagnostica',
          patterns: ['apreciacion diagnostica', 'apreciación diagnóstica'],
          keywords: ['apreciacion', 'diagnostic']
        },
        {
          key: 'diagnostico_principal',
          patterns: ['diagnostico principal', 'diagnóstico principal'],
          keywords: ['diagnostico principal', 'diagnostic principal']
        }
      ];

      const filledMap = new Map();
      const controlByKey = new Map();
      const missing = [];
      const usedControls = new Set();
      const skippedExisting = [];
      const strictKeys = ['triage', 'presente_enfermedad', 'apreciacion_diagnostica'];

      for (const t of targets) {
        const node = findFieldNode(t.patterns.map(normalize));
        const controls = candidateControlsFromFieldNode(node, t.keywords.map(normalize));
        const ctrl = pickBestControl(controls, node, t, false) || pickBestControl(controls, node, t, true) || null;
        if (!(ctrl instanceof HTMLElement)) {
          missing.push(t.key);
          filledMap.set(t.key, false);
          continue;
        }
        controlByKey.set(t.key, ctrl);
        const existing = readValue(ctrl);
        if (existing && existing.length >= 3) {
          filledMap.set(t.key, true);
          usedControls.add(ctrl);
          skippedExisting.push(t.key);
          continue;
        }
        const ok = setControlValue(ctrl, textValue);
        filledMap.set(t.key, Boolean(ok));
        if (ok) usedControls.add(ctrl);
      }

      // Si dos campos sensibles cayeron sobre el mismo control, forzar reasignación única.
      const strictByCtrl = new Map();
      for (const key of strictKeys) {
        const ctrl = controlByKey.get(key);
        if (!(ctrl instanceof HTMLElement)) continue;
        const arr = strictByCtrl.get(ctrl) || [];
        arr.push(key);
        strictByCtrl.set(ctrl, arr);
      }
      const strictReassign = [];
      for (const keys of strictByCtrl.values()) {
        if (keys.length <= 1) continue;
        for (let i = 1; i < keys.length; i += 1) strictReassign.push(keys[i]);
      }
      for (const key of strictReassign) {
        filledMap.set(key, false);
        controlByKey.delete(key);
      }

      // Pase de refuerzo para los campos más sensibles del flujo.
      for (const key of strictKeys) {
        if (filledMap.get(key)) continue;
        const targetDef = targets.find((t) => t.key === key);
        if (!targetDef) continue;
        const reservedStrict = new Set(
          strictKeys
            .map((k) => controlByKey.get(k))
            .filter((x) => x instanceof HTMLElement)
        );
        const fallback = Array.from(
          document.querySelectorAll(
            'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], [id], [name], [aria-label], [placeholder], [title]'
          )
        )
          .filter((el) => el instanceof HTMLElement)
          .filter((el) => !usedControls.has(el))
          .filter((el) => !reservedStrict.has(el))
          .map((el) => {
            const meta = normalize(
              `${el.id || ''} ${el.getAttribute?.('name') || ''} ${el.getAttribute?.('aria-label') || ''} ${el.getAttribute?.('placeholder') || ''} ${el.getAttribute?.('title') || ''}`
            );
            const r = el.getBoundingClientRect?.() || { left: 9999, top: 9999, width: 0, height: 0 };
            let score = 0;
            for (const k of targetDef.keywords.map(normalize)) {
              if (meta.includes(k)) score += 120;
            }
            if (r.left >= 220) score += 25; // panel de contenido (no menú lateral)
            if (r.width >= 120) score += 10;
            if (key === 'presente_enfermedad') {
              if (el instanceof HTMLTextAreaElement) score += 220;
              if (el.getAttribute?.('contenteditable') === 'true') score += 160;
              if (r.height >= 70) score += 80;
            }
            return { el, score };
          })
          .filter((x) => x.score > 0)
          .sort((a, b) => b.score - a.score);
        for (const picked of fallback) {
          if (!(picked?.el instanceof HTMLElement)) continue;
          if (!isEditable(picked.el)) continue;
          const existingFb = readValue(picked.el);
          if (existingFb && existingFb.length >= 3) {
            filledMap.set(key, true);
            controlByKey.set(key, picked.el);
            usedControls.add(picked.el);
            skippedExisting.push(key);
            break;
          }
          const ok = setControlValue(picked.el, textValue);
          if (!ok) continue;
          filledMap.set(key, true);
          controlByKey.set(key, picked.el);
          usedControls.add(picked.el);
          break;
        }
      }

      // Pase final fuerte: garantiza intento dirigido para los 2 campos que suelen fallar.
      const finalHardKeys = ['presente_enfermedad', 'apreciacion_diagnostica'];
      for (const key of finalHardKeys) {
        if (filledMap.get(key)) continue;
        const targetDef = targets.find((t) => t.key === key);
        if (!targetDef) continue;
        const node = findFieldNode(targetDef.patterns.map(normalize));
        const controlsNear = candidateControlsFromFieldNode(node, targetDef.keywords.map(normalize));
        const controlsGlobal = Array.from(
          document.querySelectorAll(
            'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], [id], [name], [aria-label], [placeholder], [title]'
          )
        ).filter((el) => el instanceof HTMLElement && isEditable(el));
        const pool = dedupe([...controlsNear, ...controlsGlobal])
          .map((el) => ({ el, score: scoreControlForTarget(el, node, targetDef) }))
          .filter((x) => x.score > -5000)
          .sort((a, b) => b.score - a.score)
          .slice(0, 16);

        for (const item of pool) {
          const ctrl = item.el;
          if (!(ctrl instanceof HTMLElement)) continue;
          // Evita pisar control ya asignado a otro campo sensible.
          const takenByOtherStrict = strictKeys.some((k) => k !== key && controlByKey.get(k) === ctrl);
          if (takenByOtherStrict) continue;
          const existingHard = readValue(ctrl);
          if (existingHard && existingHard.length >= 3) {
            filledMap.set(key, true);
            controlByKey.set(key, ctrl);
            usedControls.add(ctrl);
            skippedExisting.push(key);
            break;
          }
          const ok = setControlValue(ctrl, textValue);
          if (!ok) continue;
          filledMap.set(key, true);
          controlByKey.set(key, ctrl);
          usedControls.add(ctrl);
          break;
        }
      }

      // Verificación anclada: los dos campos largos deben estar realmente cerca de su label.
      const anchoredKeys = ['presente_enfermedad', 'apreciacion_diagnostica'];
      for (const key of anchoredKeys) {
        const targetDef = targets.find((t) => t.key === key);
        if (!targetDef) continue;
        const node = findFieldNode(targetDef.patterns.map(normalize));
        const ctrl = controlByKey.get(key);
        if (!(node instanceof HTMLElement) || !(ctrl instanceof HTMLElement)) {
          filledMap.set(key, false);
          continue;
        }
        const nodeContainer = getClosestContainer(node);
        const ctrlContainer = getClosestContainer(ctrl);
        const sameContainer =
          (nodeContainer instanceof HTMLElement && nodeContainer.contains(ctrl)) ||
          (ctrlContainer instanceof HTMLElement && ctrlContainer.contains(node)) ||
          (nodeContainer instanceof HTMLElement && ctrlContainer instanceof HTMLElement && nodeContainer === ctrlContainer);
        const d = rectDistance(node.getBoundingClientRect(), ctrl.getBoundingClientRect());
        if (!sameContainer && d > 760) {
          filledMap.set(key, false);
          controlByKey.delete(key);
        }
      }

      const strictDistinctCount = new Set(
        strictKeys
          .map((k) => controlByKey.get(k))
          .filter((x) => x instanceof HTMLElement)
      ).size;
      const strictDistinctOk = strictDistinctCount >= strictKeys.length;

      const readyForGenerar = targets.every((t) => Boolean(filledMap.get(t.key))) && strictDistinctOk;
      let generarClicked = false;
      if (clickGenerar && readyForGenerar) {
        // Deja que el DOM aplique bindings del último set antes de generar IA.
        try {
          const t0 = performance.now();
          while (performance.now() - t0 < 220) {}
        } catch {}
        const presenteCtrl = controlByKey.get('presente_enfermedad');
        const appCtrl = controlByKey.get('apreciacion_diagnostica');
        const scoreGenerarButton = (btn, anchorRect = null) => {
          const txt = normalize(
            `${btn.textContent || ''} ${btn.getAttribute('title') || ''} ${btn.getAttribute('aria-label') || ''} ${btn.id || ''} ${btn.getAttribute('name') || ''}`
          );
          if (!txt.includes('generar')) return -9999;
          const r = btn.getBoundingClientRect();
          let score = 0;
          score += 120;
          if (txt.includes('ia')) score += 80;
          if (txt.includes('diagnostic') || txt.includes('apreciacion')) score += 35;
          if (anchorRect) {
            // Priorización pedida: botón debajo de "Presente enfermedad" y hacia la derecha.
            if (r.top >= anchorRect.top - 10) score += 40;
            if (r.top >= anchorRect.top + 28) score += 55;
            if (r.left >= (anchorRect.left + anchorRect.width * 0.35)) score += 55;
            if (r.left <= (anchorRect.left + anchorRect.width + 420)) score += 25;
            const d = rectDistance(r, anchorRect);
            score += Math.max(0, 220 - d);
          }
          return score;
        };
        const clickBestGenerar = (anchorRect = null) => {
          const all = Array.from(
            document.querySelectorAll('button, a, span, input[type="button"], input[type="submit"], [role="button"]')
          ).filter(visible);
          if (anchorRect) {
            const zone = [];
            for (const b of all) {
              const txt = normalize(
                `${b.textContent || ''} ${b.getAttribute('title') || ''} ${b.getAttribute('aria-label') || ''} ${b.id || ''} ${b.getAttribute('name') || ''}`
              );
              if (!txt.includes('generar')) continue;
              const r = b.getBoundingClientRect();
              const inRight = r.left >= (anchorRect.left + anchorRect.width * 0.35);
              const inBelow = r.top >= (anchorRect.top - 10);
              const inBand = r.top <= (anchorRect.top + Math.max(340, anchorRect.height + 260));
              if (!(inRight && inBelow && inBand)) continue;
              zone.push({ b, score: scoreGenerarButton(b, anchorRect) + 260 });
            }
            if (zone.length) {
              zone.sort((a, b) => b.score - a.score);
              return safeClick(zone[0].b);
            }
          }
          const scored = [];
          for (const b of all) {
            const sc = scoreGenerarButton(b, anchorRect);
            if (sc <= 0) continue;
            scored.push({ b, score: sc });
          }
          if (!scored.length) return false;
          scored.sort((a, b) => b.score - a.score);
          return safeClick(scored[0].b);
        };
        const tryButtonsNear = (rootEl) => {
          if (!(rootEl instanceof HTMLElement)) return false;
          const buttons = Array.from(
            rootEl.querySelectorAll('button, a, span, input[type="button"], input[type="submit"], [role="button"], [id], [name], [title], [aria-label]')
          ).filter(visible);
          const scored = [];
          for (const b of buttons) {
            const txt = normalize(`${b.textContent || ''} ${b.getAttribute('title') || ''} ${b.getAttribute('aria-label') || ''} ${b.id || ''} ${b.getAttribute('name') || ''}`);
            if (!txt) continue;
            let score = 0;
            if (txt.includes('generar')) score += 120;
            if (txt.includes('ia')) score += 70;
            if (txt.includes('apreciacion')) score += 35;
            if (txt.includes('diagnostic')) score += 25;
            if (score <= 0) continue;
            scored.push({ b, score });
          }
          if (!scored.length) return false;
          scored.sort((a, b) => b.score - a.score);
          return safeClick(scored[0].b);
        };

        if (presenteCtrl instanceof HTMLElement && !generarClicked) {
          generarClicked = clickBestGenerar(presenteCtrl.getBoundingClientRect());
        }
        if (appCtrl instanceof HTMLElement && !generarClicked) {
          generarClicked = clickBestGenerar(appCtrl.getBoundingClientRect());
        }

        if (appCtrl instanceof HTMLElement) {
          let p = appCtrl.parentElement;
          for (let i = 0; i < 4 && p && !generarClicked; i += 1) {
            generarClicked = tryButtonsNear(p);
            p = p.parentElement;
          }
        }

        if (!generarClicked) {
          const all = Array.from(document.querySelectorAll('button, a, span, input[type="button"], input[type="submit"], [role="button"]')).filter(visible);
          const scored = [];
          for (const b of all) {
            const txt = normalize(`${b.textContent || ''} ${b.getAttribute('title') || ''} ${b.getAttribute('aria-label') || ''}`);
            let score = 0;
            if (txt.includes('generar')) score += 110;
            if (txt.includes('ia')) score += 80;
            if (txt.includes('apreciacion')) score += 25;
            if (score <= 0) continue;
            const r = b.getBoundingClientRect();
            score += Math.max(0, 500 - (r.left + r.top) / 4);
            scored.push({ b, score });
          }
          if (scored.length) {
            scored.sort((a, b) => b.score - a.score);
            generarClicked = safeClick(scored[0].b);
          }
        }
      }

      const filledCount = targets.filter((t) => filledMap.get(t.key)).length;
      const allFilled = filledCount === targets.length;
      const missingFinal = targets.filter((t) => !filledMap.get(t.key)).map((t) => t.key);
      return {
        ok: allFilled && strictDistinctOk && (!clickGenerar || generarClicked),
        allFilled,
        readyForGenerar,
        strictDistinctOk,
        strictDistinctCount,
        filledCount,
        total: targets.length,
        clickedGenerar: generarClicked,
        missing: missingFinal,
        skippedExisting
      };
    }, { textValue: NOTA_MEDICA_FIELDS_TEXT, clickGenerar: false });

    lastResult = result;
    if (result?.skippedExisting?.length) {
      console.log(`NOTA_MEDICA_FIELD_SKIP_EXISTING origin=${origin || '-'} skipped=${result.skippedExisting.join(',')}`);
    }
    const okFill = result?.allFilled && result?.strictDistinctOk;
    let generarHumanOk = !AUTO_CLICK_GENERAR_IA_NOTA_MEDICA;
    let touchStats = { touched: 0, total: 0 };
    let captureStats = { captured: 0, total: 0 };
    if (AUTO_CLICK_GENERAR_IA_NOTA_MEDICA && result?.readyForGenerar && okFill) {
      await updateBotStatusOverlay(page, 'working', 'click en Generar IA...');
      // Captura fuerte por teclado real para que el sistema marque los campos como "capturados".
      captureStats = await captureNotaFieldsByKeyboard(
        page,
        ['consulta_por', 'triage', 'presente_enfermedad', 'apreciacion_diagnostica', 'diagnostico_principal'],
        NOTA_MEDICA_FIELDS_TEXT
      );
      for (let g = 0; g < 4 && !generarHumanOk; g += 1) {
        touchStats = await ensureTrustedTouchOnNotaFields(page, ['triage', 'presente_enfermedad', 'apreciacion_diagnostica']);
        generarHumanOk = await clickGenerarIaByHumanAction(page);
        const hasCaptureAlerts = await hasNotaRequiredFieldAlerts(page);
        if (hasCaptureAlerts) {
          console.log(`NOTA_MEDICA_GENERAR_RETRY_CAPTURE_ALERT origin=${origin || '-'} click_try=${g + 1}`);
          generarHumanOk = false;
          const recapture = await captureNotaFieldsByKeyboard(
            page,
            ['consulta_por', 'triage', 'presente_enfermedad', 'apreciacion_diagnostica'],
            NOTA_MEDICA_FIELDS_TEXT
          );
          captureStats = {
            captured: Math.max(captureStats.captured, recapture.captured || 0),
            total: Math.max(captureStats.total, recapture.total || 0)
          };
          await waitForTimeoutRaw(page, 220);
        }
      }
    }
    if (okFill && generarHumanOk) {
      console.log(
        `NOTA_MEDICA_FIELDS_FILL_OK origin=${origin || '-'} attempts=${attempts} filled=${result.filledCount}/${result.total} strict=${result.strictDistinctCount || 0}/3 ready=${result.readyForGenerar ? 1 : 0} captured=${captureStats.captured}/${captureStats.total} touched=${touchStats.touched}/${touchStats.total} generar=${generarHumanOk ? 1 : 0}`
      );
      return true;
    }

    if (attempts % 3 === 0) {
      const missingText = Array.isArray(result?.missing) ? result.missing.join(',') : '-';
      console.log(
        `NOTA_MEDICA_FIELDS_FILL_RETRY origin=${origin || '-'} attempts=${attempts} filled=${result?.filledCount || 0}/${result?.total || 5} strict=${result?.strictDistinctCount || 0}/3 ready=${result?.readyForGenerar ? 1 : 0} captured=${captureStats.captured || 0}/${captureStats.total || 0} touched=${touchStats.touched || 0}/${touchStats.total || 0} generar=${generarHumanOk ? 1 : 0} missing=${missingText}`
      );
    }
    await sleepRaw(NOTA_MEDICA_FIELDS_FILL_RETRY_MS);
  }

  const missingText = Array.isArray(lastResult?.missing) ? lastResult.missing.join(',') : '-';
  const timeoutButReady = Boolean(lastResult?.allFilled && lastResult?.readyForGenerar);
  console.log(
    `NOTA_MEDICA_FIELDS_FILL_TIMEOUT origin=${origin || '-'} attempts=${attempts} filled=${lastResult?.filledCount || 0}/${lastResult?.total || 5} strict=${lastResult?.strictDistinctCount || 0}/3 ready=${lastResult?.readyForGenerar ? 1 : 0} missing=${missingText} accept_as_ready=${timeoutButReady ? 1 : 0}`
  );
  // Si los campos están todos llenos y ready, aceptar como éxito aunque Generar IA haya fallado
  return timeoutButReady;
}

async function readNotaMedicaRequiredState(page) {
  if (isPageClosedSafe(page)) {
    return { allFilled: false, filledCount: 0, total: 5, missing: ['consulta_por', 'triage', 'presente_enfermedad', 'apreciacion_diagnostica', 'diagnostico_principal'] };
  }
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 6 && r.height > 6;
      };
      const isEditable = (el) => {
        if (!(el instanceof HTMLElement) || !visible(el)) return false;
        if (el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true') return false;
        if (el.getAttribute('readonly') !== null || el.getAttribute('aria-readonly') === 'true') return false;
        if (el instanceof HTMLTextAreaElement) return true;
        if (el instanceof HTMLInputElement) {
          const t = (el.type || 'text').toLowerCase();
          return t === 'text' || t === 'search' || t === 'email' || t === 'tel' || t === 'url' || t === 'number';
        }
        return el.getAttribute('contenteditable') === 'true';
      };
      const readValue = (el) => {
        if (!(el instanceof HTMLElement)) return '';
        try {
          if (el instanceof HTMLTextAreaElement || el instanceof HTMLInputElement) return normalize(el.value || '');
          if (el.getAttribute('contenteditable') === 'true') return normalize(el.textContent || '');
        } catch {}
        return '';
      };
      const getClosestContainer = (el) => {
        if (!(el instanceof HTMLElement)) return null;
        return (
          el.closest('.form-group, .form-row, .row, tr, td, .k-form-field, .k-edit-field, .k-window-content, [class*="field"], [class*="group"], section, article') ||
          el.parentElement ||
          null
        );
      };
      const dedupe = (arr) => {
        const out = [];
        const seen = new Set();
        for (const el of arr) {
          if (!(el instanceof HTMLElement)) continue;
          if (seen.has(el)) continue;
          seen.add(el);
          out.push(el);
        }
        return out;
      };
      const rectDistance = (a, b) => {
        if (!a || !b) return 99999;
        const ax = a.left + a.width / 2;
        const ay = a.top + a.height / 2;
        const bx = b.left + b.width / 2;
        const by = b.top + b.height / 2;
        const dx = ax - bx;
        const dy = ay - by;
        return Math.sqrt(dx * dx + dy * dy);
      };
      const controlMeta = (el) =>
        normalize(
          `${el?.id || ''} ${el?.getAttribute?.('name') || ''} ${el?.getAttribute?.('aria-label') || ''} ${el?.getAttribute?.('placeholder') || ''} ${el?.getAttribute?.('title') || ''}`
        );
      const findFieldNode = (patterns) => {
        const nodes = Array.from(document.querySelectorAll('label, span, div, p, strong, td, th, li, h1, h2, h3'))
          .filter(visible)
          .map((n) => ({ n, t: normalize(n.textContent || '') }))
          .filter((x) => x.t && x.t.length >= 3 && x.t.length <= 130);
        let best = null;
        for (const item of nodes) {
          const hits = patterns.filter((p) => item.t.includes(normalize(p))).length;
          if (!hits) continue;
          const score = hits * 100 - item.t.length;
          if (!best || score > best.score) best = { node: item.n, score };
        }
        return best ? best.node : null;
      };
      const candidateControlsFromFieldNode = (node, keywords = []) => {
        const candidates = [];
        if (!(node instanceof HTMLElement)) return candidates;
        const htmlFor = node.getAttribute('for');
        if (htmlFor) {
          const direct = document.getElementById(htmlFor);
          if (direct instanceof HTMLElement) candidates.push(direct);
        }

        const nearSelectors =
          'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], .k-input-inner, .k-input';
        const near1 = node.nextElementSibling;
        if (near1 instanceof HTMLElement) {
          candidates.push(...Array.from(near1.querySelectorAll(nearSelectors)));
          candidates.push(near1);
        }

        const container = getClosestContainer(node);
        if (container) candidates.push(...Array.from(container.querySelectorAll(nearSelectors)));

        let p = node.parentElement;
        for (let i = 0; i < 3 && p; i += 1) {
          candidates.push(...Array.from(p.querySelectorAll(nearSelectors)));
          p = p.parentElement;
        }

        if (keywords.length) {
          const globalInputs = Array.from(
            document.querySelectorAll(
              'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], [id], [name], [aria-label], [placeholder], [title]'
            )
          );
          for (const g of globalInputs) {
            if (!(g instanceof HTMLElement)) continue;
            const meta = controlMeta(g);
            if (keywords.some((k) => meta.includes(normalize(k)))) candidates.push(g);
          }
        }
        return dedupe(candidates);
      };
      const scoreControlForTarget = (ctrl, node, targetDef) => {
        if (!(ctrl instanceof HTMLElement)) return -99999;
        const r = ctrl.getBoundingClientRect();
        if (!visible(ctrl)) return -99999;
        let score = 0;
        const meta = controlMeta(ctrl);
        for (const k of targetDef.keywords.map(normalize)) {
          if (meta.includes(k)) score += 140;
        }
        const nodeRect = node instanceof HTMLElement ? node.getBoundingClientRect() : null;
        score += Math.max(0, 220 - rectDistance(nodeRect, r));
        if (r.left >= 220) score += 20;
        if (ctrl instanceof HTMLTextAreaElement) score += 70;
        if (ctrl.getAttribute('contenteditable') === 'true') score += 60;
        if (targetDef.key === 'presente_enfermedad') {
          if (ctrl instanceof HTMLTextAreaElement) score += 220;
          if (ctrl.getAttribute('contenteditable') === 'true') score += 150;
          if (r.height >= 65) score += 70;
        }
        if (targetDef.key === 'apreciacion_diagnostica') {
          if (ctrl instanceof HTMLTextAreaElement) score += 190;
          if (ctrl.getAttribute('contenteditable') === 'true') score += 170;
          if (r.height >= 55) score += 55;
        }
        return score;
      };

      const targets = [
        { key: 'consulta_por', patterns: ['consulta por'], keywords: ['consulta'] },
        { key: 'triage', patterns: ['triage'], keywords: ['triage'] },
        { key: 'presente_enfermedad', patterns: ['presente enfermedad', 'enfermedad presente'], keywords: ['presente', 'enfermedad'] },
        { key: 'apreciacion_diagnostica', patterns: ['apreciacion diagnostica', 'apreciación diagnóstica'], keywords: ['apreciacion', 'diagnostic'] },
        { key: 'diagnostico_principal', patterns: ['diagnostico principal', 'diagnóstico principal'], keywords: ['diagnostico principal', 'diagnostic principal'] }
      ];

      const missing = [];
      const values = {};
      for (const t of targets) {
        const node = findFieldNode(t.patterns);
        const controls = candidateControlsFromFieldNode(node, t.keywords);
        const best = controls
          .filter((c) => c instanceof HTMLElement && isEditable(c))
          .map((c) => ({ c, score: scoreControlForTarget(c, node, t) }))
          .sort((a, b) => b.score - a.score)[0]?.c;

        const v = readValue(best);
        values[t.key] = v.slice(0, 80);
        if (!v) missing.push(t.key);
      }

      const filledCount = targets.length - missing.length;
      return {
        allFilled: missing.length === 0,
        filledCount,
        total: targets.length,
        missing,
        values
      };
    });
  } catch {
    return { allFilled: false, filledCount: 0, total: 5, missing: ['consulta_por', 'triage', 'presente_enfermedad', 'apreciacion_diagnostica', 'diagnostico_principal'] };
  }
}

async function isTableroMedicoTabActive(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    return await page.evaluate(() => {
      const normalize = (s) => (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
      // Buscar tab "Tablero Médico" seleccionado en las tabs de Telerik RadTabStrip
      const tabs = document.querySelectorAll('.rtsLink, [role="tab"]');
      for (const t of tabs) {
        const txt = normalize(t.textContent || '');
        if (txt.includes('tablero medico') || txt.includes('tablero médico')) {
          // Verificar que esté seleccionado
          if (t.classList.contains('rtsSelected') || t.parentElement?.classList.contains('rtsSelected')) return true;
        }
      }
      return false;
    });
  } catch { return false; }
}

async function closeTableroMedicoTab(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    // Buscar la X de cierre del tab "Tablero Médico"
    const closed = await page.evaluate(() => {
      const normalize = (s) => (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
      // Buscar el tab de Tablero Médico y su botón X
      const tabs = document.querySelectorAll('.rtsLI, [role="tab"]');
      for (const tab of tabs) {
        const txt = normalize(tab.textContent || '');
        if (!txt.includes('tablero medico') && !txt.includes('tablero médico')) continue;
        // Buscar el botón X dentro del tab
        const closeBtn = tab.querySelector('.rtsClose, .rtsCloseButton, [class*="close"], a[title="Close"], button[title="Close"]');
        if (closeBtn instanceof HTMLElement) {
          closeBtn.click();
          return { clicked: true, via: 'close_btn_in_tab' };
        }
      }
      // Fallback: buscar cualquier X cercana al texto "Tablero Médico"
      const links = document.querySelectorAll('a, button, span');
      for (const el of links) {
        const txt = normalize(el.textContent || '');
        const title = normalize(el.getAttribute('title') || '');
        if (txt === '×' || txt === 'x' || txt === '✕' || title === 'close') {
          const r = el.getBoundingClientRect();
          if (r.top < 100 && r.width > 5 && r.height > 5) {
            el.click();
            return { clicked: true, via: 'close_x_header' };
          }
        }
      }
      return { clicked: false };
    });

    if (!closed?.clicked) {
      // Fallback: Playwright locator para la X del tab
      const xLoc = page.locator('.rtsClose, [class*="rtsClose"]').first();
      if ((await xLoc.count()) > 0) {
        await xLoc.click({ force: true, timeout: 2000 });
        console.log('CLOSE_TABLERO_TAB_OK via=locator_rtsClose');
        await waitForTimeoutRaw(page, 500);
        return true;
      }
    } else {
      console.log(`CLOSE_TABLERO_TAB_OK via=${closed.via}`);
      await waitForTimeoutRaw(page, 500);
      return true;
    }
  } catch {}

  // Último fallback: click en tab "Agenda médica" directamente
  try {
    const agendaLoc = page.locator('text=/Agenda\\s*m[eé]dica/i').first();
    if ((await agendaLoc.count()) > 0) {
      await agendaLoc.click({ force: true, timeout: 2000 });
      console.log('CLOSE_TABLERO_TAB_OK via=click_agenda_tab');
      await waitForTimeoutRaw(page, 500);
      return true;
    }
  } catch {}

  console.log('CLOSE_TABLERO_TAB_FAIL');
  return false;
}

async function waitForTableroMedicoSidebar(page, timeoutMs = 30000) {
  const started = Date.now();
  while ((Date.now() - started) < timeoutMs) {
    if (isPageClosedSafe(page)) return false;
    try {
      const found = await page.evaluate(() => {
        const normalize = (s) => (s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
        // Estrategia 1: buscar por ID directo del botón "Nota médica"
        const btnById = document.querySelector('#btnIdMenuLate5MP_TableroMedico, [id$="btnIdMenuLate5MP_TableroMedico"]');
        if (btnById) {
          const st = getComputedStyle(btnById);
          const r = btnById.getBoundingClientRect();
          if (st.display !== 'none' && st.visibility !== 'hidden' && r.width > 5 && r.height > 5) return { ready: true, via: 'id' };
        }
        // Estrategia 2: buscar cualquier elemento del sidebar del Tablero con texto "Nota médica"
        const items = document.querySelectorAll('[id*="MenuLate"], [id*="TableroMedico"], [role="option"], [role="listitem"], li, a, span, div');
        for (const el of items) {
          const txt = normalize(el.textContent || '');
          if (txt === 'nota medica' || txt === 'nota médica') {
            const st = getComputedStyle(el);
            const r = el.getBoundingClientRect();
            if (st.display !== 'none' && st.visibility !== 'hidden' && r.width > 10 && r.height > 10) return { ready: true, via: 'text' };
          }
        }
        return { ready: false };
      });
      if (found?.ready) {
        console.log(`TABLERO_MEDICO_SIDEBAR_READY elapsed=${Date.now() - started}ms via=${found.via}`);
        return true;
      }
    } catch {}
    await waitForTimeoutRaw(page, 800);
  }
  console.log(`TABLERO_MEDICO_SIDEBAR_TIMEOUT timeout=${timeoutMs}ms`);
  return false;
}

async function ensureNotaMedicaReadyForFinalize(page, origin = '') {
  if (isPageClosedSafe(page)) return false;

  // Paso 1: esperar a que el Tablero Médico cargue (sidebar visible con "Nota médica")
  const sidebarReady = await waitForTableroMedicoSidebar(page, 60000);
  if (!sidebarReady) {
    await updateBotStatusOverlay(page, 'error', 'Tablero Médico no cargó');
    console.log(`NOTA_MEDICA_FAIL origin=${origin || '-'} reason=tablero_medico_not_loaded`);
    return false;
  }
  await updateBotStatusOverlay(page, 'info', 'Tablero Médico listo');

  // Paso 2: click en "Nota médica"
  const notaOpened = await openNotaMedicaFromSidebar(page, origin);
  if (!notaOpened) {
    await updateBotStatusOverlay(page, 'error', 'no se pudo abrir Nota médica');
    return false;
  }
  await updateBotStatusOverlay(page, 'success', 'Nota médica abierta');

  // Paso 2.5: cerrar modal "Catálogo de diagnósticos" si se abrió accidentalmente
  await dismissCatalogoDiagnosticosModal(page);

  await updateBotStatusOverlay(page, 'working', 'leyendo campos...');
  const initialState = await readNotaMedicaRequiredState(page);
  console.log(
    `NOTA_MEDICA_STATE_PRE origin=${origin || '-'} filled=${initialState.filledCount || 0}/${initialState.total || 5} missing=${(initialState.missing || []).join(',') || '-'}`
  );

  // Validar campo "Consulta por" y mostrar estado en overlay
  const consultaPorValue = (initialState.values?.consulta_por || '').trim();
  if (consultaPorValue) {
    const snippet = consultaPorValue.slice(0, 35);
    await updateBotStatusOverlay(page, 'success', `"Consulta por" lleno: "${snippet}..."`);
    console.log(`CONSULTA_POR_CHECK origin=${origin || '-'} status=filled value="${snippet}"`);
  } else {
    await updateBotStatusOverlay(page, 'warning', '"Consulta por" está vacío');
    console.log(`CONSULTA_POR_CHECK origin=${origin || '-'} status=empty`);
  }
  await waitForTimeoutRaw(page, 1200);

  // Validar campo "Triage" y mostrar estado en overlay
  const triageValue = (initialState.values?.triage || '').trim();
  if (triageValue) {
    const snippet = triageValue.slice(0, 35);
    await updateBotStatusOverlay(page, 'success', `"Triage" lleno: "${snippet}..."`);
    console.log(`TRIAGE_CHECK origin=${origin || '-'} status=filled value="${snippet}"`);
  } else {
    await updateBotStatusOverlay(page, 'warning', '"Triage" está vacío');
    console.log(`TRIAGE_CHECK origin=${origin || '-'} status=empty`);
  }
  await waitForTimeoutRaw(page, 1200);

  // Validar campo "Presente enfermedad" y mostrar estado en overlay
  const presenteEnfValue = (initialState.values?.presente_enfermedad || '').trim();
  if (presenteEnfValue) {
    const snippet = presenteEnfValue.slice(0, 35);
    await updateBotStatusOverlay(page, 'success', `"Presente enfermedad" lleno: "${snippet}..."`);
    console.log(`PRESENTE_ENFERMEDAD_CHECK origin=${origin || '-'} status=filled value="${snippet}"`);
  } else {
    await updateBotStatusOverlay(page, 'warning', '"Presente enfermedad" está vacío');
    console.log(`PRESENTE_ENFERMEDAD_CHECK origin=${origin || '-'} status=empty`);
  }
  await waitForTimeoutRaw(page, 1200);

  // Validar campo "Apreciación diagnóstica" y mostrar estado en overlay
  const apreciacionValue = (initialState.values?.apreciacion_diagnostica || '').trim();
  if (apreciacionValue) {
    const snippet = apreciacionValue.slice(0, 35);
    await updateBotStatusOverlay(page, 'success', `"Apreciación diagnóstica" lleno: "${snippet}..."`);
    console.log(`APRECIACION_DIAGNOSTICA_CHECK origin=${origin || '-'} status=filled value="${snippet}"`);
  } else {
    await updateBotStatusOverlay(page, 'warning', '"Apreciación diagnóstica" está vacío');
    console.log(`APRECIACION_DIAGNOSTICA_CHECK origin=${origin || '-'} status=empty`);
  }
  await waitForTimeoutRaw(page, 1200);

  // Validar si el input "Diagnóstico principal" (mp_nm_DiagP_wrapper) ya tiene data
  const diagInputData = await page.evaluate(() => {
    // Buscar el wrapper del diagnóstico principal
    const wrapper = document.querySelector('[id$="mp_nm_DiagP_wrapper"]');
    if (!wrapper) return { found: false, hasData: false, text: '' };
    // El texto puede estar en el wrapper, en un input hijo, o en un span hijo
    const txt = (wrapper.textContent || wrapper.innerText || '').trim();
    const input = wrapper.querySelector('input');
    const inputVal = input ? (input.value || '').trim() : '';
    const data = inputVal || txt;
    return { found: true, hasData: data.length > 0, text: data.slice(0, 60) };
  });
  console.log(`DIAG_INPUT_CHECK origin=${origin || '-'} found=${diagInputData.found ? 1 : 0} hasData=${diagInputData.hasData ? 1 : 0} text="${diagInputData.text}"`);

  if (diagInputData.found && diagInputData.hasData) {
    // Input tiene data = diagnóstico ya existe → no generar
    await updateBotStatusOverlay(page, 'success', `diagnóstico existente: "${diagInputData.text.slice(0, 35)}..."`);
    console.log(`DIAG_ALREADY_EXISTS origin=${origin || '-'} text="${diagInputData.text}"`);
    await waitForTimeoutRaw(page, 1200);
  } else {
    // Input vacío = necesita generar diagnóstico → click y esperar
    await updateBotStatusOverlay(page, 'working', 'generando diagnóstico con IA...');
    console.log(`DIAG_EMPTY_GENERATING origin=${origin || '-'}`);
    const generarClicked = await clickGenerarIaByHumanAction(page);
    if (generarClicked) {
      // Esperar hasta que el input tenga data (= IA terminó de generar)
      const generarStart = Date.now();
      const GENERAR_IA_WAIT_MS = 30000;
      let diagGenerated = false;
      let lastLog = 0;
      while ((Date.now() - generarStart) < GENERAR_IA_WAIT_MS) {
        if (isPageClosedSafe(page)) break;
        // Verificar si el input ya tiene data
        const check = await page.evaluate(() => {
          const w = document.querySelector('[id$="mp_nm_DiagP_wrapper"]');
          if (!w) return { hasData: false };
          const txt = (w.textContent || w.innerText || '').trim();
          const input = w.querySelector('input');
          const val = input ? (input.value || '').trim() : '';
          return { hasData: (val || txt).length > 0 };
        });

        if (check.hasData) {
          diagGenerated = true;
          break;
        }

        const elapsed = Date.now() - generarStart;
        if (elapsed - lastLog >= 3000) {
          const secs = Math.round(elapsed / 1000);
          await updateBotStatusOverlay(page, 'working', `esperando diagnóstico IA... (${secs}s)`);
          console.log(`GENERAR_IA_WAITING origin=${origin || '-'} elapsed=${elapsed}ms`);
          lastLog = elapsed;
        }
        await sleepRaw(600);
      }
      if (diagGenerated) {
        await updateBotStatusOverlay(page, 'success', 'diagnóstico generado por IA!');
        console.log(`GENERAR_IA_RESULT origin=${origin || '-'} status=generated elapsed=${Date.now() - generarStart}ms`);
      } else {
        await updateBotStatusOverlay(page, 'warning', 'diagnóstico no se generó a tiempo (30s)');
        console.log(`GENERAR_IA_RESULT origin=${origin || '-'} status=timeout elapsed=${Date.now() - generarStart}ms`);
      }
    } else {
      await updateBotStatusOverlay(page, 'warning', 'no se pudo clickear Generar IA');
      console.log(`GENERAR_IA_BTN_CLICK_FAIL origin=${origin || '-'}`);
    }
    await waitForTimeoutRaw(page, 1200);
  }

  // Validar campo "Diagnóstico principal" (resultado de Generar IA)
  const finalState = await readNotaMedicaRequiredState(page);
  const diagValue = (finalState.values?.diagnostico_principal || '').trim();
  if (diagValue) {
    const snippet = diagValue.slice(0, 35);
    await updateBotStatusOverlay(page, 'success', `"Diagnóstico principal" lleno: "${snippet}..."`);
    console.log(`DIAGNOSTICO_PRINCIPAL_CHECK origin=${origin || '-'} status=filled value="${snippet}"`);
  } else {
    await updateBotStatusOverlay(page, 'warning', '"Diagnóstico principal" está vacío');
    console.log(`DIAGNOSTICO_PRINCIPAL_CHECK origin=${origin || '-'} status=empty`);
  }
  await waitForTimeoutRaw(page, 1200);

  // Verificar estado final de todos los campos
  console.log(
    `NOTA_MEDICA_STATE_FINAL origin=${origin || '-'} filled=${finalState.filledCount || 0}/${finalState.total || 5} missing=${(finalState.missing || []).join(',') || '-'}`
  );
  const allFieldsReady = finalState.allFilled && diagValue;
  if (allFieldsReady) {
    await updateBotStatusOverlay(page, 'success', `todos los campos listos (${finalState.filledCount}/${finalState.total})`);
    console.log(`NOTA_MEDICA_READY_OK origin=${origin || '-'} via=refactored_flow`);
  } else {
    // Fallback: intentar llenado automático si faltan campos
    await updateBotStatusOverlay(page, 'working', `faltan ${finalState.missing?.length || '?'} campos, llenando...`);
    const filled = await fillNotaMedicaAntecedentesAndGenerateIA(page, origin);
    if (!filled) {
      await updateBotStatusOverlay(page, 'error', 'falló llenando campos');
      return false;
    }
    await updateBotStatusOverlay(page, 'success', 'campos llenados OK');
  }
  return true;
}

async function resolvePlanTratamientoPoints(page) {
  if (isPageClosedSafe(page)) return { fieldCandidates: [], generarCandidates: [] };
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 6 && r.height > 6;
      };
      const isEditable = (el) => {
        if (!(el instanceof HTMLElement) || !visible(el)) return false;
        if (el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true') return false;
        if (el.getAttribute('readonly') !== null || el.getAttribute('aria-readonly') === 'true') return false;
        if (el instanceof HTMLTextAreaElement) return true;
        if (el instanceof HTMLInputElement) {
          const t = (el.type || 'text').toLowerCase();
          return t === 'text' || t === 'search' || t === 'email' || t === 'tel' || t === 'url' || t === 'number';
        }
        return el.getAttribute('contenteditable') === 'true';
      };
      const getClosestContainer = (el) => {
        if (!(el instanceof HTMLElement)) return null;
        return (
          el.closest('.form-group, .form-row, .row, tr, td, .k-form-field, .k-edit-field, .k-window-content, [class*="field"], [class*="group"], section, article') ||
          el.parentElement ||
          null
        );
      };
      const rectDistance = (a, b) => {
        if (!a || !b) return 99999;
        const ax = a.left + a.width / 2;
        const ay = a.top + a.height / 2;
        const bx = b.left + b.width / 2;
        const by = b.top + b.height / 2;
        const dx = ax - bx;
        const dy = ay - by;
        return Math.sqrt(dx * dx + dy * dy);
      };
      const controlMeta = (el) =>
        normalize(
          `${el?.id || ''} ${el?.getAttribute?.('name') || ''} ${el?.getAttribute?.('aria-label') || ''} ${el?.getAttribute?.('placeholder') || ''} ${el?.getAttribute?.('title') || ''}`
        );
      const dedupe = (arr) => {
        const out = [];
        const seen = new Set();
        for (const el of arr) {
          if (!(el instanceof HTMLElement)) continue;
          if (seen.has(el)) continue;
          seen.add(el);
          out.push(el);
        }
        return out;
      };

      const labels = Array.from(document.querySelectorAll('label, span, div, p, strong, td, th, li, h1, h2, h3'))
        .filter(visible)
        .map((n) => ({ n, t: normalize(n.textContent || '') }))
        .filter((x) => x.t && x.t.length >= 4 && x.t.length <= 160);

      let anchorNode = null;
      let anchorScore = -1;
      for (const item of labels) {
        const t = item.t;
        let score = 0;
        if (t.includes('plan de tratamiento')) score += 500;
        if (t.includes('plan tratamiento')) score += 420;
        if (t.includes('plan') && t.includes('tratamiento')) score += 280;
        if (score > anchorScore) {
          anchorNode = item.n;
          anchorScore = score;
        }
      }
      if (!(anchorNode instanceof HTMLElement) || anchorScore <= 0) {
        return { fieldCandidates: [], generarCandidates: [] };
      }

      const anchorRect = anchorNode.getBoundingClientRect();
      const anchorContainer = getClosestContainer(anchorNode);
      const nearSelectors =
        'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], .k-input-inner, .k-input';

      const pool = [];
      const htmlFor = anchorNode.getAttribute('for');
      if (htmlFor) {
        const direct = document.getElementById(htmlFor);
        if (direct instanceof HTMLElement) pool.push(direct);
      }
      const sibling = anchorNode.nextElementSibling;
      if (sibling instanceof HTMLElement) {
        pool.push(sibling);
        pool.push(...Array.from(sibling.querySelectorAll(nearSelectors)));
      }
      if (anchorContainer instanceof HTMLElement) {
        pool.push(...Array.from(anchorContainer.querySelectorAll(nearSelectors)));
      }
      let p = anchorNode.parentElement;
      for (let i = 0; i < 4 && p; i += 1) {
        pool.push(...Array.from(p.querySelectorAll(nearSelectors)));
        p = p.parentElement;
      }
      pool.push(
        ...Array.from(
          document.querySelectorAll(
            'textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], [id], [name], [aria-label], [placeholder], [title]'
          )
        )
      );

      const editables = dedupe(pool).filter(isEditable);
      const fieldCandidates = [];
      for (const el of editables) {
        const r = el.getBoundingClientRect();
        const meta = controlMeta(el);
        let score = 100;
        if (meta.includes('plan')) score += 190;
        if (meta.includes('tratamiento')) score += 220;
        score += Math.max(0, 250 - rectDistance(anchorRect, r));
        if (anchorContainer instanceof HTMLElement && anchorContainer.contains(el)) score += 120;
        if (el instanceof HTMLTextAreaElement) score += 120;
        if (el.getAttribute('contenteditable') === 'true') score += 90;
        if (r.height >= 55) score += 70;
        if (r.left >= 220) score += 20;
        fieldCandidates.push({
          x: Math.round(r.left + r.width / 2),
          y: Math.round(r.top + Math.min(r.height / 2, 16)),
          score,
          kind: 'editable'
        });
      }

      const iframeCandidates = Array.from(document.querySelectorAll('iframe'))
        .map((el, frameIdx) => ({ el, frameIdx }))
        .filter((x) => visible(x.el))
        .map(({ el, frameIdx }) => {
          const r = el.getBoundingClientRect();
          const meta = controlMeta(el);
          let score = 60;
          if (meta.includes('plan')) score += 180;
          if (meta.includes('tratamiento')) score += 200;
          score += Math.max(0, 220 - rectDistance(anchorRect, r));
          return {
            x: Math.round(r.left + r.width / 2),
            y: Math.round(r.top + r.height / 2),
            score,
            kind: 'iframe',
            frameIdx
          };
        });

      const buttons = Array.from(
        document.querySelectorAll('button, a, span, input[type="button"], input[type="submit"], [role="button"], [id], [name], [title], [aria-label]')
      ).filter(visible);
      const generarCandidates = [];
      for (const b of buttons) {
        const txt = normalize(
          `${b.textContent || ''} ${b.getAttribute('title') || ''} ${b.getAttribute('aria-label') || ''} ${b.id || ''} ${b.getAttribute('name') || ''}`
        );
        if (!txt) continue;
        let score = 0;
        if (txt.includes('generar plan')) score += 520;
        else if (txt.includes('generar') && txt.includes('plan')) score += 390;
        else if (txt.includes('plan') && txt.includes('tratamiento')) score += 260;
        else if (txt.includes('generar')) score += 120;
        if (score <= 0) continue;
        const r = b.getBoundingClientRect();
        score += Math.max(0, 260 - rectDistance(anchorRect, r));
        if (r.left >= anchorRect.left + (anchorRect.width * 0.20)) score += 80;
        if (r.top >= anchorRect.top - 20 && r.top <= anchorRect.top + 420) score += 70;
        generarCandidates.push({
          x: Math.round(r.left + r.width / 2),
          y: Math.round(r.top + r.height / 2),
          score
        });
      }

      const sortedFields = [...fieldCandidates, ...iframeCandidates]
        .sort((a, b) => b.score - a.score)
        .slice(0, 12);
      const sortedGenerar = generarCandidates.sort((a, b) => b.score - a.score).slice(0, 8);

      return {
        fieldCandidates: sortedFields,
        generarCandidates: sortedGenerar
      };
    });
  } catch {
    return { fieldCandidates: [], generarCandidates: [] };
  }
}

async function readPlanTratamientoState(page) {
  if (isPageClosedSafe(page)) return { fieldFound: false, filled: false, valueLength: 0 };
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 6 && r.height > 6;
      };
      const isEditable = (el) => {
        if (!(el instanceof HTMLElement) || !visible(el)) return false;
        if (el.hasAttribute('disabled') || el.getAttribute('aria-disabled') === 'true') return false;
        if (el.getAttribute('readonly') !== null || el.getAttribute('aria-readonly') === 'true') return false;
        if (el instanceof HTMLTextAreaElement) return true;
        if (el instanceof HTMLInputElement) {
          const t = (el.type || 'text').toLowerCase();
          return t === 'text' || t === 'search' || t === 'email' || t === 'tel' || t === 'url' || t === 'number';
        }
        return el.getAttribute('contenteditable') === 'true';
      };
      const readValue = (el) => {
        if (!(el instanceof HTMLElement)) return '';
        try {
          if (el instanceof HTMLTextAreaElement || el instanceof HTMLInputElement) return normalize(el.value || '');
          if (el.getAttribute('contenteditable') === 'true') return normalize(el.textContent || '');
        } catch {}
        return '';
      };
      const getClosestContainer = (el) => {
        if (!(el instanceof HTMLElement)) return null;
        return (
          el.closest('.form-group, .form-row, .row, tr, td, .k-form-field, .k-edit-field, .k-window-content, [class*="field"], [class*="group"], section, article') ||
          el.parentElement ||
          null
        );
      };
      const rectDistance = (a, b) => {
        if (!a || !b) return 99999;
        const ax = a.left + a.width / 2;
        const ay = a.top + a.height / 2;
        const bx = b.left + b.width / 2;
        const by = b.top + b.height / 2;
        const dx = ax - bx;
        const dy = ay - by;
        return Math.sqrt(dx * dx + dy * dy);
      };
      const controlMeta = (el) =>
        normalize(
          `${el?.id || ''} ${el?.getAttribute?.('name') || ''} ${el?.getAttribute?.('aria-label') || ''} ${el?.getAttribute?.('placeholder') || ''} ${el?.getAttribute?.('title') || ''}`
        );

      const labels = Array.from(document.querySelectorAll('label, span, div, p, strong, td, th, li, h1, h2, h3'))
        .filter(visible)
        .map((n) => ({ n, t: normalize(n.textContent || '') }))
        .filter((x) => x.t && x.t.length >= 4 && x.t.length <= 160);
      const anchor = labels
        .map((x) => {
          let score = 0;
          if (x.t.includes('plan de tratamiento')) score += 500;
          if (x.t.includes('plan tratamiento')) score += 420;
          if (x.t.includes('plan') && x.t.includes('tratamiento')) score += 280;
          return { n: x.n, score };
        })
        .sort((a, b) => b.score - a.score)[0];
      if (!(anchor?.n instanceof HTMLElement) || (anchor.score || 0) <= 0) {
        return { fieldFound: false, filled: false, valueLength: 0 };
      }

      const anchorRect = anchor.n.getBoundingClientRect();
      const anchorContainer = getClosestContainer(anchor.n);
      const controls = Array.from(
        document.querySelectorAll('textarea, input[type="text"], input:not([type]), input[type="search"], [contenteditable="true"], [id], [name], [aria-label], [placeholder], [title]')
      )
        .filter((el) => el instanceof HTMLElement && isEditable(el))
        .map((el) => {
          const r = el.getBoundingClientRect();
          const meta = controlMeta(el);
          let score = 90;
          if (meta.includes('plan')) score += 190;
          if (meta.includes('tratamiento')) score += 220;
          score += Math.max(0, 250 - rectDistance(anchorRect, r));
          if (anchorContainer instanceof HTMLElement && anchorContainer.contains(el)) score += 120;
          if (el instanceof HTMLTextAreaElement) score += 120;
          if (el.getAttribute('contenteditable') === 'true') score += 90;
          return { el, score };
        })
        .sort((a, b) => b.score - a.score);

      const best = controls[0]?.el;
      const value = readValue(best);
      const len = value.length;
      return { fieldFound: true, filled: len >= 6, valueLength: len };
    });
  } catch {
    return { fieldFound: false, filled: false, valueLength: 0 };
  }
}

async function clickGenerarPlanButton(page, points = []) {
  if (isPageClosedSafe(page)) return false;

  const selectors = [
    // Selector directo por ID parcial (el más confiable)
    '[id$="btnMostrarPlan"]',
    '[id$="btnMostrarPlan"] > div',
    // Selectores por texto "Mostrar plan"
    'button:has-text("Mostrar plan"), a:has-text("Mostrar plan"), [role="button"]:has-text("Mostrar plan")',
    'button:has-text("Mostrar Plan"), a:has-text("Mostrar Plan"), [role="button"]:has-text("Mostrar Plan")',
    // Selectores por texto "Generar plan" (fallback)
    'button:has-text("Generar plan"), a:has-text("Generar plan"), [role="button"]:has-text("Generar plan")',
    'button:has-text("Generar Plan"), a:has-text("Generar Plan"), [role="button"]:has-text("Generar Plan")',
    '[title*="generar plan" i], [aria-label*="generar plan" i], [id*="generarplan" i], [name*="generarplan" i]',
    '[title*="mostrar plan" i], [aria-label*="mostrar plan" i], [id*="mostrarplan" i]'
  ];
  for (const sel of selectors) {
    try {
      const loc = page.locator(sel).first();
      if ((await loc.count()) === 0) continue;
      if (!(await loc.isVisible())) continue;
      await loc.click({ force: true, timeout: 1200 });
      await waitForTimeoutRaw(page, 220);
      console.log(`PLAN_TRATAMIENTO_GENERAR_CLICK_OK via=selector "${sel}"`);
      return true;
    } catch {}
  }

  for (let i = 0; i < points.length && i < 6; i += 1) {
    const p = points[i];
    if (!Number.isFinite(p?.x) || !Number.isFinite(p?.y)) continue;
    try {
      await page.mouse.move(p.x, p.y);
      await waitForTimeoutRaw(page, 26);
      await page.mouse.click(p.x, p.y, { delay: 20 });
      await waitForTimeoutRaw(page, 220);
      console.log(`PLAN_TRATAMIENTO_GENERAR_CLICK_OK via=xy score=${p.score || 0}`);
      return true;
    } catch {}
  }

  try {
    const clicked = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        try {
          el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          el.click();
          return true;
        } catch {
          return false;
        }
      };
      const nodes = Array.from(
        document.querySelectorAll('button, a, span, div, input[type="button"], input[type="submit"], [role="button"], [title], [aria-label], [id], [name]')
      ).filter(visible);
      const scored = [];
      for (const n of nodes) {
        const txt = normalize(
          `${n.textContent || ''} ${n.getAttribute('title') || ''} ${n.getAttribute('aria-label') || ''} ${n.id || ''} ${n.getAttribute('name') || ''}`
        );
        if (!txt) continue;
        let score = 0;
        if (txt.includes('mostrar plan')) score += 600;
        else if (txt.includes('mostrar') && txt.includes('plan')) score += 460;
        else if (txt.includes('generar plan')) score += 500;
        else if (txt.includes('generar') && txt.includes('plan')) score += 360;
        else if (txt.includes('generar') && txt.includes('tratamiento')) score += 300;
        // Bonus por ID que contiene btnMostrarPlan
        if ((n.id || '').toLowerCase().includes('mostrarplan')) score += 700;
        if (score <= 0) continue;
        scored.push({ n, score });
      }
      if (!scored.length) return false;
      scored.sort((a, b) => b.score - a.score);
      return safeClick(scored[0].n);
    });
    if (clicked) {
      await waitForTimeoutRaw(page, 220);
      console.log('PLAN_TRATAMIENTO_GENERAR_CLICK_OK via=dom_scored');
      return true;
    }
  } catch {}

  console.log('PLAN_TRATAMIENTO_GENERAR_CLICK_FAIL');
  return false;
}

async function hasPlanTratamientoRequiredAlert(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 4 && r.height > 4;
      };
      const nodes = Array.from(
        document.querySelectorAll('.k-tooltip-validation,.field-validation-error,.validation-summary-errors,.error,.alert,.k-notification,[role="alert"],span,div,p,li')
      ).filter(visible);
      const text = normalize(nodes.map((n) => n.textContent || '').join(' | '));
      const mentionsPlan = text.includes('plan de tratamiento') || (text.includes('plan') && text.includes('tratamiento'));
      const required =
        text.includes('no se ha capturado') ||
        text.includes('no se capturo') ||
        text.includes('obligatorio') ||
        text.includes('requerido') ||
        text.includes('debe ingresar') ||
        text.includes('complete');
      return Boolean(mentionsPlan && required);
    });
  } catch {
    return false;
  }
}

async function ensurePlanTratamientoAndGenerate(page, origin = '') {
  if (!AUTO_GENERAR_PLAN_TRATAMIENTO) {
    console.log('PLAN_TRATAMIENTO_STEP_SKIP auto=0');
    return true;
  }
  if (isPageClosedSafe(page)) return false;

  // Paso 1: Buscar el campo "Plan de tratamiento"
  await updateBotStatusOverlay(page, 'working', 'buscando campo Plan de tratamiento...');
  let state = await readPlanTratamientoState(page);
  if (!state.fieldFound) {
    const fieldStart = Date.now();
    while (!state.fieldFound && (Date.now() - fieldStart) < 6000) {
      await waitForTimeoutRaw(page, 500);
      state = await readPlanTratamientoState(page);
    }
  }
  if (!state.fieldFound) {
    await updateBotStatusOverlay(page, 'warning', 'campo Plan de tratamiento no encontrado');
    console.log(`PLAN_TRATAMIENTO_FIELD_NOT_FOUND origin=${origin || '-'}`);
    await waitForTimeoutRaw(page, 1200);
    return false;
  }

  // Paso 2: Validar si ya tiene data
  if (state.filled) {
    await updateBotStatusOverlay(page, 'success', `"Plan de tratamiento" lleno (${state.valueLength} chars)`);
    console.log(`PLAN_TRATAMIENTO_CHECK origin=${origin || '-'} status=filled len=${state.valueLength}`);
    await waitForTimeoutRaw(page, 600);
    return true;
  }

  // Paso 3: Plan vacío → click en btnSolIma (Imagenología) para crear registro
  await updateBotStatusOverlay(page, 'working', 'Plan vacío, abriendo Imagenología (btnSolIma)...');
  console.log(`PLAN_TRATAMIENTO_CHECK origin=${origin || '-'} status=empty action=click_btnSolIma`);
  await waitForTimeoutRaw(page, 600);

  let clickedIma = false;
  // Intento 1: por ID exacto
  try {
    const imaBtn = page.locator('[id$="btnSolIma"]');
    const cnt = await imaBtn.count();
    if (cnt > 0) {
      await imaBtn.first().click({ timeout: 4000 });
      clickedIma = true;
      console.log('PLAN_TRATAMIENTO_BTNSOLIMA_OK via=id_locator');
    }
  } catch (e) {
    console.log(`PLAN_TRATAMIENTO_BTNSOLIMA_ID_ERR ${(e.message || '').slice(0, 80)}`);
  }

  // Intento 2: por ID con wrapper
  if (!clickedIma) {
    try {
      const imaWrap = page.locator('[id$="btnSolIma_wrapper"] a, [id$="btnSolIma_wrapper"] button');
      const cnt2 = await imaWrap.count();
      if (cnt2 > 0) {
        await imaWrap.first().click({ timeout: 4000 });
        clickedIma = true;
        console.log('PLAN_TRATAMIENTO_BTNSOLIMA_OK via=wrapper_locator');
      }
    } catch (e) {
      console.log(`PLAN_TRATAMIENTO_BTNSOLIMA_WRAP_ERR ${(e.message || '').slice(0, 80)}`);
    }
  }

  // Intento 3: por evaluate + Playwright click
  if (!clickedIma) {
    try {
      const info = await page.evaluate(() => {
        const candidates = [
          document.querySelector('[id$="btnSolIma"]'),
          document.querySelector('[id$="btnSolIma_wrapper"] a'),
          document.querySelector('[id$="btnSolIma_wrapper"] button'),
        ].filter(Boolean);
        for (const el of candidates) {
          const r = el.getBoundingClientRect();
          if (r.width > 5 && r.height > 5) {
            return { id: el.id, x: r.x + r.width / 2, y: r.y + r.height / 2 };
          }
        }
        return null;
      });
      if (info) {
        await page.mouse.click(info.x, info.y);
        clickedIma = true;
        console.log(`PLAN_TRATAMIENTO_BTNSOLIMA_OK via=mouse_click id=${info.id}`);
      }
    } catch (e) {
      console.log(`PLAN_TRATAMIENTO_BTNSOLIMA_MOUSE_ERR ${(e.message || '').slice(0, 80)}`);
    }
  }

  if (!clickedIma) {
    await updateBotStatusOverlay(page, 'warning', 'no se pudo clickear btnSolIma');
    console.log(`PLAN_TRATAMIENTO_BTNSOLIMA_FAIL origin=${origin || '-'}`);
    await waitForTimeoutRaw(page, 1200);
    console.log(`PLAN_TRATAMIENTO_READY_OK origin=${origin || '-'} clicked_ima=0`);
    return true;
  }

  // Paso 4: Verificar que el modal "Solicitud de estudios de imagenología" abrió
  await updateBotStatusOverlay(page, 'working', 'verificando modal de Imagenología...');
  await waitForTimeoutRaw(page, 1500);

  let modalOpen = false;
  const modalStart = Date.now();
  while (!modalOpen && (Date.now() - modalStart) < 8000) {
    modalOpen = await page.evaluate(() => {
      // Buscar ventana Telerik/Kendo visible con título que contenga "imagenolog"
      const windows = document.querySelectorAll('.k-window, .k-dialog, [role="dialog"]');
      for (const w of windows) {
        const st = getComputedStyle(w);
        if (st.display === 'none' || st.visibility === 'hidden') continue;
        const r = w.getBoundingClientRect();
        if (r.width < 100 || r.height < 100) continue;
        const title = (w.querySelector('.k-window-title, .k-dialog-title, [class*="title"]')?.textContent || '').toLowerCase();
        if (title.includes('imagenolog') || title.includes('solicitud')) return true;
      }
      // Fallback: buscar el campo Estudio con catalogButton visible
      const catBtn = document.querySelector('[id$="mp_lab_txt_4_catalogButton"]');
      if (catBtn) {
        const r = catBtn.getBoundingClientRect();
        if (r.width > 5 && r.height > 5) return true;
      }
      return false;
    });
    if (!modalOpen) await waitForTimeoutRaw(page, 500);
  }

  if (!modalOpen) {
    await updateBotStatusOverlay(page, 'warning', 'modal de Imagenología no detectado');
    console.log(`PLAN_IMA_MODAL_NOT_FOUND origin=${origin || '-'}`);
    await waitForTimeoutRaw(page, 1000);
    console.log(`PLAN_TRATAMIENTO_READY_OK origin=${origin || '-'} clicked_ima=1 modal=0`);
    return true;
  }

  console.log(`PLAN_IMA_MODAL_OPEN origin=${origin || '-'}`);

  // Paso 5: Click en botón catálogo de "Estudio" (mp_lab_txt_4_catalogButton)
  await updateBotStatusOverlay(page, 'working', 'click en catálogo Estudio...');
  await waitForTimeoutRaw(page, 600);

  let clickedEstudio = false;
  // Intento 1: locator por ID
  try {
    const estBtn = page.locator('[id$="mp_lab_txt_4_catalogButton"]');
    const cnt = await estBtn.count();
    if (cnt > 0) {
      await estBtn.first().click({ timeout: 4000 });
      clickedEstudio = true;
      console.log('PLAN_IMA_ESTUDIO_CATALOG_OK via=id_locator');
    }
  } catch (e) {
    console.log(`PLAN_IMA_ESTUDIO_CATALOG_ID_ERR ${(e.message || '').slice(0, 80)}`);
  }

  // Intento 2: evaluate + mouse click
  if (!clickedEstudio) {
    try {
      const info = await page.evaluate(() => {
        const btn = document.querySelector('[id$="mp_lab_txt_4_catalogButton"]');
        if (!btn) return null;
        const r = btn.getBoundingClientRect();
        if (r.width < 5 || r.height < 5) return null;
        return { x: r.x + r.width / 2, y: r.y + r.height / 2, id: btn.id || '' };
      });
      if (info) {
        await page.mouse.click(info.x, info.y);
        clickedEstudio = true;
        console.log(`PLAN_IMA_ESTUDIO_CATALOG_OK via=mouse_click id=${info.id}`);
      }
    } catch (e) {
      console.log(`PLAN_IMA_ESTUDIO_CATALOG_MOUSE_ERR ${(e.message || '').slice(0, 80)}`);
    }
  }

  if (clickedEstudio) {
    await updateBotStatusOverlay(page, 'success', 'catálogo Estudio abierto');
    await waitForTimeoutRaw(page, 1200);
  } else {
    await updateBotStatusOverlay(page, 'warning', 'no se pudo abrir catálogo Estudio');
    console.log(`PLAN_IMA_ESTUDIO_CATALOG_FAIL origin=${origin || '-'}`);
    await waitForTimeoutRaw(page, 1000);
  }

  console.log(`PLAN_TRATAMIENTO_READY_OK origin=${origin || '-'} clicked_ima=1 modal=1 estudio=${clickedEstudio ? 1 : 0}`);
  return true;
}

async function clickFinalizarCitaInModule(page) {
  if (isPageClosedSafe(page)) return false;

  // Paso 1: detectar btnResolver y quitar disabled si es necesario (solo detección)
  try {
    const detected = await page.evaluate(() => {
      const btn = document.querySelector('[id$="btnResolver"]');
      if (!btn || !(btn instanceof HTMLElement)) return { found: false };
      const st = getComputedStyle(btn);
      const r = btn.getBoundingClientRect();
      const isVisible = st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      if (!isVisible) return { found: true, visible: false, reason: 'not_visible' };
      const isDisabled = btn.disabled || btn.classList.contains('k-state-disabled');
      if (isDisabled) {
        btn.disabled = false;
        btn.classList.remove('k-state-disabled');
        btn.removeAttribute('aria-disabled');
      }
      return { found: true, visible: true, wasDisabled: isDisabled, id: btn.id };
    });

    if (detected?.found && detected?.visible) {
      // Paso 2: click con Playwright locator (dispara postback Telerik correctamente)
      const loc = page.locator('[id$="btnResolver"]').first();
      try {
        await loc.click({ force: true, timeout: 2000 });
        await waitForTimeoutRaw(page, 300);
        console.log(`FINALIZAR_CLICK_OK via=btnResolver_locator wasDisabled=${detected.wasDisabled ? 1 : 0}`);
        return true;
      } catch (e) {
        console.log(`FINALIZAR_BTNRESOLVER_LOCATOR_FAIL err=${(e?.message || '').slice(0, 80)}`);
      }
    }
    if (detected?.found && !detected?.visible) {
      console.log(`FINALIZAR_BTNRESOLVER_FOUND_BUT reason=${detected.reason || 'not_visible'}`);
    }
  } catch {}

  // Estrategia 2: selectores directos por texto (fallback)
  const directSelectors = [
    'button:has-text("Finalizar cita"), a:has-text("Finalizar cita"), [role="button"]:has-text("Finalizar cita")',
    '[title*="finalizar" i], [aria-label*="finalizar" i]'
  ];
  for (const sel of directSelectors) {
    try {
      const loc = page.locator(sel).first();
      if ((await loc.count()) === 0) continue;
      if (!(await loc.isVisible())) continue;
      await loc.click({ force: true, timeout: 1200 });
      await waitForTimeoutRaw(page, 220);
      console.log(`FINALIZAR_CLICK_OK via=selector "${sel}"`);
      return true;
    } catch {}
  }

  console.log('FINALIZAR_CLICK_FAIL');
  return false;
}

async function clickFinalizarCitaInModuleWithRetry(page, options = {}) {
  if (isPageClosedSafe(page)) return false;
  const timeoutMs = (() => {
    const n = Number(options?.timeoutMs ?? CANCEL_ACTION_WAIT_TIMEOUT_MS);
    if (!Number.isFinite(n)) return CANCEL_ACTION_WAIT_TIMEOUT_MS;
    return Math.min(45000, Math.max(1000, Math.round(n)));
  })();
  const intervalMs = (() => {
    const n = Number(options?.intervalMs ?? CANCEL_ACTION_WAIT_INTERVAL_MS);
    if (!Number.isFinite(n)) return CANCEL_ACTION_WAIT_INTERVAL_MS;
    return Math.min(2000, Math.max(120, Math.round(n)));
  })();

  const started = Date.now();
  let tries = 0;
  while ((Date.now() - started) < timeoutMs) {
    if (isPageClosedSafe(page)) return false;
    tries += 1;

    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 120);
    }
    if (await isNuevaCitaModalVisible(page)) {
      await closeNuevaCitaModalIfOpen(page);
      await waitForTimeoutRaw(page, 110);
    }

    const clicked = await clickFinalizarCitaInModule(page);
    if (clicked) {
      console.log(`FINALIZAR_CLICK_RETRY_OK tries=${tries} elapsed=${Date.now() - started}ms`);
      return true;
    }

    if (tries % 3 === 0) {
      console.log(`FINALIZAR_CLICK_RETRY_WAIT tries=${tries} elapsed=${Date.now() - started}ms`);
    }
    await sleepRaw(intervalMs);
  }

  console.log(`FINALIZAR_CLICK_RETRY_TIMEOUT elapsed=${Date.now() - started}ms`);
  return false;
}

async function processNotaMedicaAndFinalizar(page, origin = '') {
  if (isPageClosedSafe(page)) return false;
  console.log(`MODE2_NOTA_FINALIZAR_START origin=${origin || '-'}`);
  // El overlay se actualiza DENTRO de ensureNotaMedicaReadyForFinalize con mensajes específicos

  const ready = await ensureNotaMedicaReadyForFinalize(page, origin);
  if (!ready) {
    console.log(`MODE2_NOTA_FINALIZAR_FAIL origin=${origin || '-'} step=nota_ready`);
    await updateBotStatusOverlay(page, 'error', 'falló en Nota médica');
    return false;
  }

  await updateBotStatusOverlay(page, 'working', 'generando Plan de tratamiento...');
  const planReady = await ensurePlanTratamientoAndGenerate(page, origin);
  if (!planReady) {
    console.log(`MODE2_NOTA_FINALIZAR_FAIL origin=${origin || '-'} step=plan_tratamiento`);
    await updateBotStatusOverlay(page, 'error', 'falló en Plan tratamiento');
    return false;
  }
  await updateBotStatusOverlay(page, 'success', 'Plan generado!');

  // Intentar Finalizar con reintentos (no volver a validar campos/plan si ya están OK)
  const FINALIZAR_MAX_RETRIES = 3;
  for (let fAttempt = 1; fAttempt <= FINALIZAR_MAX_RETRIES; fAttempt++) {
    if (isPageClosedSafe(page)) return false;

    await updateBotStatusOverlay(page, 'working', fAttempt === 1 ? 'click en Finalizar...' : `reintentando Finalizar... (${fAttempt}/${FINALIZAR_MAX_RETRIES})`);
    console.log(`FINALIZAR_ATTEMPT origin=${origin || '-'} attempt=${fAttempt}/${FINALIZAR_MAX_RETRIES}`);

    const clickedFinalizar = await clickFinalizarCitaInModuleWithRetry(page);
    if (!clickedFinalizar) {
      console.log(`FINALIZAR_CLICK_FAIL origin=${origin || '-'} attempt=${fAttempt}`);
      if (fAttempt < FINALIZAR_MAX_RETRIES) {
        await updateBotStatusOverlay(page, 'warning', 'no se pudo clickear Finalizar, reintentando...');
        await waitForTimeoutRaw(page, 1500);
        continue;
      }
      await updateBotStatusOverlay(page, 'error', 'falló al click Finalizar');
      return false;
    }

    await updateBotStatusOverlay(page, 'working', 'esperando diálogo de confirmación...');
    await waitForTimeoutRaw(page, 1200);

    // Reintentar confirmar diálogo hasta 3 veces (puede tardar en aparecer)
    let confirmed = false;
    for (let cAttempt = 1; cAttempt <= 3; cAttempt++) {
      confirmed = await confirmCancellationDialog(page);
      if (confirmed) break;
      console.log(`CONFIRM_DIALOG_RETRY attempt=${cAttempt}/3`);
      await waitForTimeoutRaw(page, 800);
    }

    if (!confirmed) {
      console.log(`FINALIZAR_CONFIRM_FAIL origin=${origin || '-'} attempt=${fAttempt} - diálogo no confirmado`);
      await updateBotStatusOverlay(page, 'warning', 'no se pudo confirmar Sí, reintentando...');
      if (fAttempt < FINALIZAR_MAX_RETRIES) {
        await waitForTimeoutRaw(page, 1500);
        continue;
      }
      await updateBotStatusOverlay(page, 'error', 'falló al confirmar Finalizar');
      return false;
    }

    await updateBotStatusOverlay(page, 'working', 'esperando confirmación del sistema...');
    const feedback = await waitForCancellationFeedback(page, 5500);
    console.log(
      `MODE2_NOTA_FINALIZAR_OK origin=${origin || '-'} clicked=1 confirmed=${confirmed ? 1 : 0} feedback=${feedback ? 1 : 0} attempt=${fAttempt}`
    );

    // ── Click "Continuar" en modal "Seleccionar iniciar descanso" ──
    await updateBotStatusOverlay(page, 'working', 'esperando modal Continuar...');
    await waitForTimeoutRaw(page, 1500);

    let continuarClicked = false;
    for (let cntAttempt = 1; cntAttempt <= 5; cntAttempt++) {
      if (isPageClosedSafe(page)) break;

      // Intento directo por ID exacto
      try {
        const btnContinuar = page.locator('#ctl00_nc003_MP_TableroMedico_v10NotaMedicaMP_mp_segop_btnContinuar');
        const cnt = await btnContinuar.count();
        if (cnt > 0) {
          await btnContinuar.click({ timeout: 3000 });
          continuarClicked = true;
          console.log(`CONTINUAR_CLICK_OK via=id attempt=${cntAttempt}`);
          break;
        }
      } catch {}

      // Intento por selector parcial ID
      try {
        const btnPartial = page.locator('[id$="btnContinuar"] button, [id$="btnContinuar"]');
        const cnt2 = await btnPartial.count();
        if (cnt2 > 0) {
          await btnPartial.first().click({ timeout: 3000 });
          continuarClicked = true;
          console.log(`CONTINUAR_CLICK_OK via=partial_id attempt=${cntAttempt}`);
          break;
        }
      } catch {}

      // Intento por texto "Continuar"
      try {
        const btnText = page.locator('button:has-text("Continuar")');
        const cnt3 = await btnText.count();
        if (cnt3 > 0) {
          await btnText.first().click({ timeout: 3000 });
          continuarClicked = true;
          console.log(`CONTINUAR_CLICK_OK via=text attempt=${cntAttempt}`);
          break;
        }
      } catch {}

      console.log(`CONTINUAR_WAIT attempt=${cntAttempt}/5`);
      await waitForTimeoutRaw(page, 800);
    }

    if (continuarClicked) {
      await waitForTimeoutRaw(page, 500);
      await updateBotStatusOverlay(page, 'success', 'cita finalizada exitosamente!');
    } else {
      console.log('CONTINUAR_NOT_FOUND - modal may not have appeared');
      await updateBotStatusOverlay(page, 'success', 'cita finalizada!');
    }
    return true;
  }

  console.log(`MODE2_NOTA_FINALIZAR_FAIL origin=${origin || '-'} step=finalizar_all_retries`);
  await updateBotStatusOverlay(page, 'error', 'falló al finalizar después de reintentos');
  return false;
}

async function holdBrowserForReview(page, reason = 'validacion visual', holdMs = REVIEW_HOLD_MS) {
  const msRaw = Number(holdMs);
  const ms = Number.isFinite(msRaw) ? msRaw : REVIEW_HOLD_MS;
  if (ms <= 0) {
    console.log(`Navegador abierto sin límite para ${reason}. Cierra la ventana para terminar.`);
    while (true) {
      if (isPageClosedSafe(page)) break;
      await sleepRaw(1000);
    }
    return;
  }
  const sec = Math.round(ms / 1000);
  console.log(`Navegador abierto ${sec}s para ${reason}.`);
  try {
    await waitForTimeoutRaw(page, ms);
  } catch {}
}

// Post-Save Strategy:
// 1) Posicionar en casilla guardada.
// 2) Mantener popup activo.
// 3) Click forzado y rápido en botón "Módulo".
async function focusSavedAppointmentSlotForVideoModal(page, slot) {
  if (!slot) return { clicked: false, opened: false, via: 'none' };

  // 1) Coordenadas guardadas del slot: posicionar + hover (evita abrir "Editar cita").
  if (Number.isFinite(slot.x) && Number.isFinite(slot.y)) {
    try {
      await page.mouse.move(slot.x, slot.y);
      await waitForTimeoutRaw(page, 90);
      const opened = await isNuevaCitaAsignadaModalVisible(page);
      return { clicked: true, opened, via: 'coords_hover' };
    } catch {}
  }

  // 2) Fallback por selector/índice: hover sin click.
  try {
    if (slot.selector && Number.isInteger(slot.domIdx)) {
      const cell = page.locator(slot.selector).nth(slot.domIdx);
      await cell.hover({ force: true, timeout: 1000 });
      await waitForTimeoutRaw(page, 90);
      const opened = await isNuevaCitaAsignadaModalVisible(page);
      return { clicked: true, opened, via: 'locator_hover' };
    }
  } catch {}

  return { clicked: false, opened: false, via: 'none' };
}

async function forceClickModuloFromSlotPopup(page, slot) {
  const x = Number.isFinite(slot?.x) ? slot.x : null;
  const y = Number.isFinite(slot?.y) ? slot.y : null;

  const result = await page.evaluate(({ x, y }) => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 10 && r.height > 8;
    };
    const zNum = (el) => {
      try {
        const n = Number.parseInt(getComputedStyle(el).zIndex || '0', 10);
        return Number.isFinite(n) ? n : 0;
      } catch {
        return 0;
      }
    };
    const distToPoint = (rect, px, py) => {
      if (!Number.isFinite(px) || !Number.isFinite(py)) return 99999;
      const cx = rect.left + rect.width / 2;
      const cy = rect.top + rect.height / 2;
      const dx = cx - px;
      const dy = cy - py;
      return Math.sqrt(dx * dx + dy * dy);
    };
    const safeClick = (el) => {
      if (!(el instanceof HTMLElement)) return false;
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
      el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
      el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      el.click();
      return true;
    };

    // Refresca foco en la celda para mantener popup activo.
    if (Number.isFinite(x) && Number.isFinite(y)) {
      const origin = document.elementFromPoint(x, y);
      if (origin instanceof HTMLElement) {
        origin.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        origin.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        origin.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      }
    }

    const nodes = Array.from(
      document.querySelectorAll('button, a, span, input[type="button"], input[type="submit"], div[role="button"]')
    ).filter(visible);

    const candidates = [];
    for (const n of nodes) {
      const txt = normalize(n.textContent || n.value || n.getAttribute?.('title') || n.getAttribute?.('aria-label') || '');
      const isModulo = txt.includes('abrir modulo') || txt.includes('abrir módulo') || txt === 'modulo' || txt.includes(' modulo ');
      if (!isModulo) continue;
      const rect = n.getBoundingClientRect();
      candidates.push({
        el: n,
        txt,
        z: zNum(n),
        dist: distToPoint(rect, x, y),
        priority: txt.includes('abrir modulo') || txt.includes('abrir módulo') ? 2 : 1
      });
    }

    if (!candidates.length) return { ok: false, reason: 'modulo_candidate_not_found' };

    candidates.sort((a, b) => {
      if (a.priority !== b.priority) return b.priority - a.priority;
      if (a.dist !== b.dist) return a.dist - b.dist;
      return b.z - a.z;
    });

    const target = candidates[0].el;
    const clicked = safeClick(target);
    if (!clicked) return { ok: false, reason: 'modulo_click_failed' };
    return { ok: true, via: 'popup_force_click', text: candidates[0].txt };
  }, { x, y });

  if (result?.ok) {
    console.log(`MODULO_CLICK_OK via ${result.via}`);
    return true;
  }
  return false;
}

async function forceClickModuloFromSavedSlotTooltip(page, slot, options = {}) {
  if (isPageClosedSafe(page)) return { ok: false, via: 'page_closed' };
  const appointmentNumber = String(options.appointmentNumber || '').trim();
  const blockFinalizada = options?.blockFinalizada === true;
  const maxEventCandidates = (() => {
    const n = Number(options.maxEventCandidates || 12);
    if (!Number.isFinite(n)) return 12;
    return Math.min(20, Math.max(3, Math.round(n)));
  })();

  try {
    const result = await page.evaluate(async ({ slot, appointmentNumber, maxEventCandidates, blockFinalizada }) => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const digitsOnly = (s) => String(s || '').replace(/\D+/g, '');
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };
      const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        try {
          el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          el.click();
          return true;
        } catch {
          return false;
        }
      };
      const fireMouse = (el, type, px, py) => {
        try {
          el.dispatchEvent(
            new MouseEvent(type, {
              bubbles: true,
              cancelable: true,
              view: window,
              clientX: px,
              clientY: py,
              button: 0,
              buttons: 1
            })
          );
        } catch {}
      };
      const firePointer = (el, type, px, py) => {
        try {
          el.dispatchEvent(
            new PointerEvent(type, {
              bubbles: true,
              cancelable: true,
              view: window,
              pointerId: 1,
              pointerType: 'mouse',
              isPrimary: true,
              clientX: px,
              clientY: py,
              button: 0,
              buttons: 1
            })
          );
        } catch {}
      };
      const parseMinutes = (raw) => {
        const t = normalize(raw || '');
        const m = t.match(/(\d{1,2})\s*:\s*(\d{2})\s*(am|pm)/);
        if (!m) return null;
        let hh = Number(m[1]) % 12;
        const mm = Number(m[2]);
        if (m[3] === 'pm') hh += 12;
        return hh * 60 + mm;
      };
      const toNumber = (v) => {
        const n = Number(v);
        return Number.isFinite(n) ? n : NaN;
      };
      const pointCandidates = [];
      const pushPoint = (x, y, label) => {
        const px = toNumber(x);
        const py = toNumber(y);
        if (!Number.isFinite(px) || !Number.isFinite(py)) return;
        pointCandidates.push({ x: Math.round(px), y: Math.round(py), label });
      };

      pushPoint(slot?.x, slot?.y, 'slot_center');
      if (Number.isFinite(toNumber(slot?.x)) && Number.isFinite(toNumber(slot?.y))) {
        pushPoint(slot.x - 4, slot.y, 'slot_left');
        pushPoint(slot.x + 4, slot.y, 'slot_right');
        pushPoint(slot.x, slot.y - 4, 'slot_top');
        pushPoint(slot.x, slot.y + 4, 'slot_bottom');
      }
      try {
        if (slot?.selector && Number.isInteger(slot?.domIdx)) {
          const cell = document.querySelectorAll(slot.selector)[slot.domIdx];
          if (cell instanceof HTMLElement && visible(cell)) {
            const r = cell.getBoundingClientRect();
            pushPoint(r.left + r.width / 2, r.top + r.height / 2, 'slot_locator_center');
          }
        }
      } catch {}

      const primaryPoint = pointCandidates[0] || null;
      const appointmentDigits = digitsOnly(appointmentNumber);

      const colBounds = (() => {
        if (!Number.isInteger(slot?.colIdx)) return null;
        const headers = Array.from(
          document.querySelectorAll('thead th, .k-scheduler-header th, .rsHeader, .rsHeaderTable th')
        ).filter(visible);
        for (const h of headers) {
          if ((h.cellIndex ?? -1) !== slot.colIdx) continue;
          const r = h.getBoundingClientRect();
          if (r.width < 12) continue;
          return { left: r.left, right: r.right };
        }
        return null;
      })();

      const expectedY = (() => {
        if (!Number.isFinite(Number(slot?.minutes))) return null;
        const rows = Array.from(
          document.querySelectorAll('.k-scheduler-content table tbody tr, .k-scheduler-table tbody tr, .rsContentTable tbody tr')
        );
        let best = null;
        const targetMinutes = Number(slot.minutes);
        for (const row of rows) {
          if (!(row instanceof HTMLElement)) continue;
          const timeCell = row.querySelector('td:first-child, th:first-child');
          const mins = parseMinutes(timeCell ? timeCell.textContent || '' : '');
          if (mins === null) continue;
          const rr = row.getBoundingClientRect();
          const cy = rr.top + rr.height / 2;
          const diff = Math.abs(mins - targetMinutes);
          if (!best || diff < best.diff) best = { diff, cy };
        }
        return best ? best.cy : null;
      })();

      const eventSelector = '.k-event, .rsApt, [class*="k-event"], [class*="Apt"], [class*="appointment"], [class*="Appointment"], [class*="event"]';
      const candidateMap = new Map();
      const upsert = (eventEl, via, baseScore = 0) => {
        if (!(eventEl instanceof HTMLElement) || !visible(eventEl)) return;
        const r = eventEl.getBoundingClientRect();
        const cx = r.left + r.width / 2;
        const cy = r.top + r.height / 2;
        const txt = normalize(
          `${eventEl.textContent || ''} ${eventEl.getAttribute('title') || ''} ${eventEl.getAttribute('aria-label') || ''} ${eventEl.id || ''}`
        );
        const txtDigits = digitsOnly(txt);
        let score = Number(baseScore) || 0;
        let dist = 99999;
        if (primaryPoint) {
          const dx = cx - primaryPoint.x;
          const dy = cy - primaryPoint.y;
          dist = Math.sqrt(dx * dx + dy * dy);
          score += Math.max(-120, 260 - dist);
          if (dist <= 24) score += 80;
        }
        if (colBounds) {
          if (cx >= (colBounds.left - 3) && cx <= (colBounds.right + 3)) score += 170;
          else score -= 90;
        }
        if (Number.isFinite(expectedY)) {
          score += Math.max(-80, 120 - Math.abs(cy - expectedY));
        }
        if (appointmentDigits && txtDigits.includes(appointmentDigits)) score += 900;
        if (txt.includes('no disponible')) score -= 900;
        if (txt.includes('programada') || txt.includes('videollamada')) score += 25;

        const existing = candidateMap.get(eventEl);
        if (existing) {
          if (score > existing.score) existing.score = score;
          if (dist < existing.dist) existing.dist = dist;
          if (!existing.vias.includes(via)) existing.vias.push(via);
          return;
        }
        candidateMap.set(eventEl, {
          el: eventEl,
          score,
          dist,
          cx: Math.round(cx),
          cy: Math.round(cy),
          text: txt,
          vias: [via]
        });
      };

      for (const pt of pointCandidates) {
        const stack = Array.from(document.elementsFromPoint(pt.x, pt.y) || []);
        for (let i = 0; i < stack.length; i += 1) {
          const el = stack[i];
          if (!(el instanceof HTMLElement)) continue;
          const eventLike = el.closest(eventSelector);
          if (!(eventLike instanceof HTMLElement)) continue;
          upsert(eventLike, `point:${pt.label}`, 680 - (i * 25));
        }
      }

      const allEvents = Array.from(document.querySelectorAll(eventSelector));
      for (const eventEl of allEvents) {
        upsert(eventEl, 'global_scan', 120);
      }

      const rankedEvents = Array.from(candidateMap.values())
        .sort((a, b) => {
          if (a.score !== b.score) return b.score - a.score;
          return a.dist - b.dist;
        })
        .slice(0, maxEventCandidates);
      if (!rankedEvents.length) return { ok: false, reason: 'no_event_candidates' };

      const pickTooltipWithModulo = () => {
        const roots = Array.from(
          document.querySelectorAll('.k-widget.k-tooltip, .k-tooltip, .k-popup, [role="tooltip"], .k-animation-container, .div_hos930AvisoPaciente')
        ).filter(visible);
        const scored = [];
        for (const root of roots) {
          const txt = normalize(root.textContent || '');
          if (!txt) continue;
          if (txt.includes('catalogo de pacientes') || txt.includes('catálogo de pacientes')) continue;
          if (txt.includes('nueva cita') && !txt.includes('programada') && !txt.includes('finalizada')) continue;
          if (blockFinalizada && txt.includes('finalizada')) continue;
          const controls = Array.from(
            root.querySelectorAll('button, a, span, div, input[type="button"], input[type="submit"], [role="button"]')
          ).filter(visible);
          const modulo = controls.find((n) => {
            const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || n.value || '');
            return t === 'modulo' || t.includes('abrir modulo') || t.includes('abrir módulo');
          });
          if (!(modulo instanceof HTMLElement)) continue;

          const r = root.getBoundingClientRect();
          const area = r.width * r.height;
          let score = 0;
          if (txt.includes('programada')) score += 35;
          if (txt.includes('finalizada')) score += blockFinalizada ? -250 : 35;
          if (txt.includes('videollamada')) score += 25;
          if (txt.includes('registro')) score += 20;
          if (txt.includes('expediente')) score += 20;
          if (appointmentDigits && digitsOnly(txt).includes(appointmentDigits)) score += 110;
          if (area <= 180000) score += Math.max(0, 180000 - area) / 7000;
          scored.push({ root, modulo, score, area });
        }
        if (!scored.length) return null;
        scored.sort((a, b) => {
          if (a.score !== b.score) return b.score - a.score;
          return a.area - b.area;
        });
        return scored[0];
      };

      for (let i = 0; i < rankedEvents.length; i += 1) {
        const candidate = rankedEvents[i];
        const target = (candidate.el.querySelector('div') || candidate.el);
        const hoverTargets = target === candidate.el ? [candidate.el] : [candidate.el, target];

        for (const hoverTarget of hoverTargets) {
          if (!(hoverTarget instanceof HTMLElement) || !visible(hoverTarget)) continue;
          try {
            hoverTarget.scrollIntoView({ block: 'center', inline: 'center' });
          } catch {}
          firePointer(hoverTarget, 'pointerover', candidate.cx, candidate.cy);
          firePointer(hoverTarget, 'pointerenter', candidate.cx, candidate.cy);
          firePointer(hoverTarget, 'pointermove', candidate.cx, candidate.cy);
          fireMouse(hoverTarget, 'mouseover', candidate.cx, candidate.cy);
          fireMouse(hoverTarget, 'mouseenter', candidate.cx, candidate.cy);
          fireMouse(hoverTarget, 'mousemove', candidate.cx, candidate.cy);
          try {
            const jq = window.jQuery || window.$;
            if (jq) {
              jq(hoverTarget).trigger('mouseover');
              jq(hoverTarget).trigger('mouseenter');
              jq(hoverTarget).trigger('mousemove');
            }
          } catch {}
        }

        for (let poll = 0; poll < 7; poll += 1) {
          await sleep(poll === 0 ? 45 : 68);
          const tooltip = pickTooltipWithModulo();
          if (!tooltip) continue;
          const mr = tooltip.modulo.getBoundingClientRect();
          const domClicked = safeClick(tooltip.modulo);
          return {
            ok: true,
            via: `force_tooltip:${candidate.vias[0] || 'event'}:poll${poll + 1}`,
            domClicked,
            moduloPoint: { x: Math.round(mr.left + mr.width / 2), y: Math.round(mr.top + mr.height / 2) }
          };
        }
      }

      const lastTooltip = pickTooltipWithModulo();
      if (lastTooltip) {
        const mr = lastTooltip.modulo.getBoundingClientRect();
        const domClicked = safeClick(lastTooltip.modulo);
        return {
          ok: true,
          via: 'force_tooltip:last_visible',
          domClicked,
          moduloPoint: { x: Math.round(mr.left + mr.width / 2), y: Math.round(mr.top + mr.height / 2) }
        };
      }

      return { ok: false, reason: blockFinalizada ? 'modulo_btn_not_found_or_finalizada_tooltip' : 'modulo_btn_not_found_on_tooltip' };
    }, { slot: slot || {}, appointmentNumber, maxEventCandidates, blockFinalizada });

    if (result?.ok) {
      const viaBase = result.via || 'force_tooltip';
      const mx = Number(result?.moduloPoint?.x);
      const my = Number(result?.moduloPoint?.y);
      // Ensure ASP.NET postback fires and hide P2H popup
      const postClickP2H = async () => {
        try {
          await page.evaluate(() => {
            const btn = document.getElementById('ctl00_nc002_MP_HOS930_btnModulo');
            if (btn instanceof HTMLElement) {
              try { btn.click(); } catch {}
              try { __doPostBack('ctl00$nc002$MP_HOS930$btnModulo', ''); } catch {}
            }
            const popup = document.querySelector('.div_hos930AvisoPaciente');
            if (popup) popup.style.display = 'none';
          });
        } catch {}
      };
      if (Number.isFinite(mx) && Number.isFinite(my)) {
        try {
          await page.mouse.move(mx, my);
          await waitForTimeoutRaw(page, 26);
          await page.mouse.click(mx, my, { delay: 18 });
          await postClickP2H();
          console.log(`MODULO_CLICK_OK via=${viaBase}:mouse_xy`);
          return { ok: true, via: `${viaBase}:mouse_xy` };
        } catch {}
      }
      if (result.domClicked) {
        await postClickP2H();
        console.log(`MODULO_CLICK_OK via=${viaBase}:dom_click`);
        return { ok: true, via: `${viaBase}:dom_click` };
      }
      return { ok: false, via: `${viaBase}:no_effective_click` };
    }
    return { ok: false, via: result?.reason || 'force_tooltip_failed' };
  } catch {
    return { ok: false, via: 'force_tooltip_exception' };
  }
}

async function clickModuloFromSavedSlotQuickAction(page, slot, options = {}) {
  if (isPageClosedSafe(page)) return { ok: false, via: 'page_closed' };
  if (!slot || !Number.isFinite(slot.x) || !Number.isFinite(slot.y)) return { ok: false, via: 'no_slot' };
  const appointmentNumber = String(options.appointmentNumber || '').trim();
  const disallowFinalizada = options?.disallowFinalizada === true;
  const skipForceTooltip = options?.skipForceTooltip === true;
  const assumeQuickVisible = options?.assumeQuickVisible === true;

  const clickFromVisibleQuick = async () => {
    try {
      return await page.evaluate(({ appointmentNumber, disallowFinalizada }) => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const digitsOnly = (s) => String(s || '').replace(/\D+/g, '');
        const visible = (el) => {
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 16;
        };
        const safeClick = (el) => {
          if (!(el instanceof HTMLElement)) return false;
          try {
            el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
            el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
            el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
            el.click();
            return true;
          } catch {
            return false;
          }
        };

        const popups = Array.from(
          document.querySelectorAll('.k-widget.k-tooltip, .k-tooltip, .k-popup, [role="tooltip"], .k-animation-container, .div_hos930AvisoPaciente')
        ).filter(visible);
        if (!popups.length) {
          // Fallback directo P2H: buscar botón por ID específico
          const directBtn = document.getElementById('ctl00_nc002_MP_HOS930_btnModulo');
          if (directBtn instanceof HTMLElement) {
            const directVisible = (() => {
              const st = getComputedStyle(directBtn);
              const r = directBtn.getBoundingClientRect();
              return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 16;
            })();
            if (directVisible) {
              safeClick(directBtn);
              try { __doPostBack('ctl00$nc002$MP_HOS930$btnModulo', ''); } catch {}
              return { ok: true, via: 'direct_btn_modulo_p2h' };
            }
          }
          return { ok: false, via: 'quick_modal_not_visible' };
        }

        const wantedDigits = digitsOnly(appointmentNumber);
        const candidates = [];
        for (const popup of popups) {
          const txt = normalize(popup.textContent || '');
          if (txt.includes('catalogo de pacientes') || txt.includes('catálogo de pacientes')) continue;
          if (txt.includes('nueva cita') && !txt.includes('modulo') && !txt.includes('módulo')) continue;

          const controls = Array.from(
            popup.querySelectorAll('button, a, span, div, input[type="button"], input[type="submit"], [role="button"]')
          ).filter(visible);
          const moduloBtn = controls.find((n) => {
            const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || n.value || '');
            return t === 'modulo' || t.includes('abrir modulo') || t.includes('abrir módulo');
          });
          if (!(moduloBtn instanceof HTMLElement)) continue;

          const isProgramada = txt.includes('programada');
          const isVideollamada = txt.includes('videollamada');
          const isFinalizada = txt.includes('finalizada') || txt.includes('inasistencia');
          if (disallowFinalizada && isFinalizada) continue;

          const txtDigits = digitsOnly(txt);
          let score = 0;
          if (isProgramada) score += 18;
          if (isVideollamada) score += 16;
          if (isFinalizada) score -= disallowFinalizada ? 120 : 10;
          if (txt.includes('modulo') || txt.includes('módulo')) score += 40;
          if (txt.includes('editar cita')) score += 22;
          if (txt.includes('registro')) score += 12;
          if (txt.includes('expediente')) score += 12;
          if (wantedDigits && txtDigits.includes(wantedDigits)) score += 120;

          const r = popup.getBoundingClientRect();
          candidates.push({ popup, moduloBtn, txt, score, area: r.width * r.height });
        }

        if (!candidates.length) return { ok: false, via: 'quick_modal_without_modulo' };

        candidates.sort((a, b) => {
          if (a.score !== b.score) return b.score - a.score;
          return a.area - b.area;
        });
        const chosen = candidates[0];
        const clicked = safeClick(chosen.moduloBtn);
        if (!clicked) return { ok: false, via: 'quick_modulo_click_failed' };
        // Trigger ASP.NET postback for P2H btnModulo
        const btnModulo = document.getElementById('ctl00_nc002_MP_HOS930_btnModulo');
        if (btnModulo instanceof HTMLElement) {
          try { safeClick(btnModulo); } catch {}
          try { __doPostBack('ctl00$nc002$MP_HOS930$btnModulo', ''); } catch {}
        }
        // Hide P2H popup so it doesn't block module load
        try {
          const p2hPopup = chosen.popup.closest('.div_hos930AvisoPaciente') || chosen.popup;
          if (p2hPopup.classList?.contains('div_hos930AvisoPaciente')) p2hPopup.style.display = 'none';
        } catch {}
        return { ok: true, via: 'quick_visible_modal_click' };
      }, { appointmentNumber, disallowFinalizada });
    } catch {
      return { ok: false, via: 'quick_visible_modal_exception' };
    }
  };

  if (assumeQuickVisible) {
    const visibleClick = await clickFromVisibleQuick();
    if (visibleClick?.ok) return visibleClick;
    return { ok: false, via: visibleClick?.via || 'quick_visible_click_failed' };
  }
  if (!skipForceTooltip) {
    const forcedTooltipClick = await forceClickModuloFromSavedSlotTooltip(page, slot, {
      appointmentNumber,
      maxEventCandidates: 12,
      blockFinalizada: disallowFinalizada
    });
    if (forcedTooltipClick.ok) return forcedTooltipClick;
  }
  const x = slot.x;
  const y = slot.y;

  try {
    // 1) Activar tooltip contextual de la cita guardada (programada/videollamada).
    const tooltipTarget = await page.evaluate(({ x, y }) => {
      const fireMouse = (el, type, px, py) =>
        el.dispatchEvent(
          new MouseEvent(type, {
            bubbles: true,
            cancelable: true,
            view: window,
            clientX: px,
            clientY: py,
            button: 0,
            buttons: 1
          })
        );
      const firePointer = (el, type, px, py) => {
        try {
          el.dispatchEvent(
            new PointerEvent(type, {
              bubbles: true,
              cancelable: true,
              view: window,
              pointerId: 1,
              pointerType: 'mouse',
              isPrimary: true,
              clientX: px,
              clientY: py,
              button: 0,
              buttons: 1
            })
          );
        } catch {}
      };
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };

      const stack = Array.from(document.elementsFromPoint(x, y) || []);
      const eventLike = [];
      for (const el of stack) {
        if (!(el instanceof HTMLElement)) continue;
        const node = el.closest('.k-event, .rsApt, [class*="k-event"], [class*="Apt"], [class*="event"], [class*="appointment"]');
        if (node instanceof HTMLElement && !eventLike.includes(node)) eventLike.push(node);
      }
      if (!eventLike.length) return { ok: false, reason: 'event_under_point_not_found' };

      const target = eventLike[0];
      const inner = (target.querySelector('div') || target);
      const rect = inner.getBoundingClientRect();
      const cx = Math.round(rect.left + rect.width / 2);
      const cy = Math.round(rect.top + rect.height / 2);
      inner.scrollIntoView({ block: 'center', inline: 'center' });
      firePointer(inner, 'pointerover', cx, cy);
      firePointer(inner, 'pointerenter', cx, cy);
      firePointer(inner, 'pointermove', cx, cy);
      fireMouse(inner, 'mouseover', cx, cy);
      fireMouse(inner, 'mouseenter', cx, cy);
      fireMouse(inner, 'mousemove', cx, cy);

      return { ok: true, cx, cy };
    }, { x, y });

    if (!tooltipTarget?.ok) {
      return { ok: false, via: tooltipTarget?.reason || 'tooltip_activation_failed' };
    }

    await sleepRaw(420);

    // 2) Click en "Modulo" dentro del tooltip pequeño de cita (no en contenedores globales).
    const tooltipClick = await page.evaluate(({ appointmentNumber, disallowFinalizada }) => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const digitsOnly = (s) => String(s || '').replace(/\D+/g, '');
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 10 && r.height > 8;
      };

      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        el.click();
        return true;
      };

      const popups = Array.from(
        document.querySelectorAll('.k-widget.k-tooltip, .k-tooltip, .k-popup, [role="tooltip"], .k-animation-container, .div_hos930AvisoPaciente')
      )
        .filter(visible)
        .map((n) => {
          const r = n.getBoundingClientRect();
          const txt = normalize(n.textContent || '');
          const txtDigits = digitsOnly(txt);
          const wantedDigits = digitsOnly(appointmentNumber);
          let score = 0;
          if (txt.includes('programada')) score += 30;
          if (txt.includes('finalizada')) score += disallowFinalizada ? -220 : 30;
          if (txt.includes('videollamada')) score += 20;
          if (txt.includes('registro')) score += 15;
          if (txt.includes('expediente')) score += 15;
          if (wantedDigits && txtDigits.includes(wantedDigits)) score += 120;
          return { n, txt, area: r.width * r.height, score };
        })
        .filter((x) => {
          if (disallowFinalizada && x.txt.includes('finalizada')) return false;
          const hasState = x.txt.includes('programada') || x.txt.includes('videollamada') || x.txt.includes('finalizada');
          const hasContext =
            x.txt.includes('expediente') ||
            x.txt.includes('registro') ||
            x.txt.includes('fecha') ||
            x.txt.includes('hora inicio') ||
            x.txt.includes('hora fin') ||
            x.txt.includes('paciente') ||
            x.txt.includes('cita programada');
          return hasState && hasContext;
        });
      if (!popups.length) return { ok: false, reason: 'tooltip_not_found' };

      popups.sort((a, b) => {
        if (a.score !== b.score) return b.score - a.score;
        return a.area - b.area;
      });
      const popup = popups[0].n;

      const controls = Array.from(
        popup.querySelectorAll('button, a, span, div, input[type="button"], input[type="submit"], [role="button"]')
      ).filter(visible);

      const modulo = controls.find((n) => {
        const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || n.value || '');
        if (!(t === 'modulo' || t.includes('abrir modulo') || t.includes('abrir módulo'))) return false;
        const r = n.getBoundingClientRect();
        return r.width >= 55 && r.width <= 260 && r.height >= 22 && r.height <= 90;
      });
      if (!(modulo instanceof HTMLElement)) return { ok: false, reason: 'modulo_not_found_in_tooltip' };

      safeClick(modulo);
      return { ok: true, via: 'tooltip_modulo_click' };
    }, { appointmentNumber, disallowFinalizada });

    if (tooltipClick?.ok) {
      console.log(`MODULO_CLICK_OK via=${tooltipClick.via}`);
      return { ok: true, via: tooltipClick.via };
    }

    // 3) Fallback final opcional (coordenadas) si tooltip no se pudo resolver.
    if (!POST_SAVE_ALLOW_GENERIC_MODULO_FALLBACK) {
      return { ok: false, via: tooltipClick?.reason || 'tooltip_click_failed_no_generic_fallback' };
    }

    // 3) Fallback final (coordenadas).
    const fallback = await page.evaluate(({ x, y, disallowFinalizada }) => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 10 && r.height > 8;
      };
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        el.click();
        return true;
      };

      const controls = Array.from(
        document.querySelectorAll('button,a,span,div,input[type="button"],input[type="submit"],[role="button"]')
      ).filter(visible);
      const modulo = controls.find((n) => {
        const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || n.value || '');
        if (t !== 'modulo' && !t.includes('abrir modulo') && !t.includes('abrir módulo')) return false;
        let p = n.parentElement;
        let foundEditar = false;
        for (let i = 0; i < 10 && p; i += 1) {
          const ptxt = normalize(p.textContent || '');
          if (disallowFinalizada && ptxt.includes('finalizada')) return false;
          if (ptxt.includes('editar cita') || ptxt.includes('programada') || ptxt.includes('finalizada')) {
            foundEditar = true;
            break;
          }
          p = p.parentElement;
        }
        return foundEditar;
      });
      if (!(modulo instanceof HTMLElement)) return { ok: false, reason: 'fallback_modulo_not_found' };
      const r = modulo.getBoundingClientRect();
      return safeClick(modulo)
        ? { ok: true, via: 'fallback_scoped_click', x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) }
        : { ok: false, reason: 'fallback_click_failed' };
    }, { x, y, disallowFinalizada });

    if (fallback?.ok) {
      console.log(`MODULO_CLICK_OK via=${fallback.via}`);
      return { ok: true, via: fallback.via };
    }

    return { ok: false, via: fallback?.reason || tooltipClick?.reason || 'modulo_popup_not_found' };
  } catch {
    return { ok: false, via: 'quick_action_exception' };
  }
}

async function isProgramadaQuickActionVisible(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 16;
      };

      const popups = Array.from(
        document.querySelectorAll('.k-widget.k-tooltip, .k-tooltip, .k-popup, [role="tooltip"], .k-animation-container, .div_hos930AvisoPaciente')
      ).filter(visible);

      for (const popup of popups) {
        const txt = normalize(popup.textContent || '');
        if (!txt) continue;
        if (txt.includes('catalogo de pacientes') || txt.includes('catálogo de pacientes')) continue;
        if (!(txt.includes('programada') || txt.includes('videollamada') || txt.includes('finalizada'))) continue;

        const controls = Array.from(
          popup.querySelectorAll('button, a, span, div, input[type="button"], input[type="submit"], [role="button"]')
        ).filter(visible);
        const hasModulo = controls.some((n) => {
          const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || n.value || '');
          return t === 'modulo' || t.includes('abrir modulo') || t.includes('abrir módulo');
        });
        if (hasModulo) return true;
      }
      return false;
    });
  } catch {
    return false;
  }
}

async function closeNuevaCitaModalIfOpen(page) {
  if (!(await isNuevaCitaModalVisible(page))) return true;

  try {
    const closeBtn = page.getByRole('button', { name: /cerrar/i }).first();
    await closeBtn.waitFor({ state: 'visible', timeout: 1200 });
    await closeBtn.click({ force: true });
    await waitForTimeoutRaw(page, 220);
  } catch {}

  if (await isNuevaCitaModalVisible(page)) {
    try {
      await page.keyboard.press('Escape');
      await waitForTimeoutRaw(page, 220);
    } catch {}
  }
  return !(await isNuevaCitaModalVisible(page));
}

async function tryOpenAssignedModalFromSavedSlot(page, slot) {
  if (isPageClosedSafe(page)) return { opened: false, via: 'page_closed' };
  if (!slot) return { opened: false, via: 'no_slot' };
  if (await isNuevaCitaAsignadaModalVisible(page)) return { opened: true, via: 'already_open' };

  const waitAssigned = async (via, timeoutMs = 950) => {
    const started = Date.now();
    while ((Date.now() - started) < timeoutMs) {
      if (await isNuevaCitaAsignadaModalVisible(page)) return { opened: true, via };
      await waitForTimeoutRaw(page, 85);
    }
    return { opened: false, via };
  };

  // Intento 1: disparar eventos sintéticos en la casilla/evento del slot (no depende del mouse humano).
  if (Number.isFinite(slot.x) && Number.isFinite(slot.y)) {
    try {
      const js = await page.evaluate(({ x, y }) => {
        const visible = (el) => {
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
        };
        const fire = (el, type) => {
          el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
        };
        const uniq = (arr) => {
          const out = [];
          const seen = new Set();
          for (const el of arr) {
            if (!(el instanceof HTMLElement)) continue;
            if (seen.has(el)) continue;
            seen.add(el);
            out.push(el);
          }
          return out;
        };

        const stack = Array.from(document.elementsFromPoint(x, y) || []);
        const nearby = [];
        for (const el of stack) {
          if (!(el instanceof HTMLElement)) continue;
          const eventLike = el.closest('.k-event, .rsApt, [class*="k-event"], [class*="Apt"], [class*="event"]');
          if (eventLike) nearby.push(eventLike);
          nearby.push(el);
        }
        const candidates = uniq([...nearby]).filter(visible);
        if (!candidates.length) return { ok: false, reason: 'no_candidates' };

        candidates.sort((a, b) => {
          const as = /k-event|rsapt|event|apt/i.test(a.className || '') ? 2 : 1;
          const bs = /k-event|rsapt|event|apt/i.test(b.className || '') ? 2 : 1;
          if (as !== bs) return bs - as;
          const ar = a.getBoundingClientRect();
          const br = b.getBoundingClientRect();
          return (ar.width * ar.height) - (br.width * br.height);
        });

        const target = candidates[0];
        const seq = ['mouseover', 'mouseenter', 'mousemove', 'mousedown', 'mouseup', 'click'];
        for (const ev of seq) fire(target, ev);
        target.click();
        return { ok: true, via: 'script_sequence' };
      }, { x: slot.x, y: slot.y });

      if (js?.ok) {
        const w = await waitAssigned(js.via);
        if (w.opened) return w;
      }
    } catch {}
  }

  // Intento 2: clicks reales sobre coordenada del slot.
  if (Number.isFinite(slot.x) && Number.isFinite(slot.y)) {
    try {
      await page.mouse.move(slot.x, slot.y);
      await waitForTimeoutRaw(page, 40);
      await page.mouse.click(slot.x, slot.y, { delay: 20 });
      let w = await waitAssigned('coords_single_click');
      if (w.opened) return w;

    } catch {}
  }

  // Intento 3: locator del slot capturado durante selección.
  if (slot.selector && Number.isInteger(slot.domIdx)) {
    const cell = page.locator(slot.selector).nth(slot.domIdx);
    try {
      await cell.click({ force: true, timeout: 900 });
      let w = await waitAssigned('locator_single_click');
      if (w.opened) return w;
    } catch {}

    try {
      const jsCell = await cell.evaluate((el) => {
        if (!(el instanceof HTMLElement)) return { ok: false, reason: 'no_element' };
        const fire = (type) => {
          el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
        };
        const seq = ['mouseover', 'mouseenter', 'mousemove', 'mousedown', 'mouseup', 'click'];
        for (const ev of seq) fire(ev);
        el.click();
        return { ok: true, via: 'locator_script_sequence' };
      });
      if (jsCell?.ok) {
        const w = await waitAssigned(jsCell.via);
        if (w.opened) return w;
      }
    } catch {}
  }

  return { opened: false, via: 'none' };
}

async function openModuloAfterAppointmentSave(page, slot, options = {}) {
  if (!AUTO_OPEN_MODULE_AFTER_SAVE) return false;
  if (isPageClosedSafe(page)) return false;
  await waitForTimeoutRaw(page, 140);
  const appointmentNumber = String(options.appointmentNumber || '').trim();

  const moduloState = {
    clicked: false,
    loaded: false,
    clickVia: 'none',
    loadedVia: 'none',
    attempts: 0
  };

  for (let attempt = 0; attempt < POST_SAVE_STRATEGY.maxAttempts; attempt += 1) {
    if (isPageClosedSafe(page)) {
      console.log(`MODULO_FLOW_ABORT page_closed attempt=${attempt + 1}`);
      break;
    }
    const attemptNo = attempt + 1;
    moduloState.attempts = attemptNo;
    let clickedThisAttempt = false;
    let clickVia = 'none';

    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 160);
    }
    if (await isNuevaCitaModalVisible(page)) {
      await closeNuevaCitaModalIfOpen(page);
      await waitForTimeoutRaw(page, 120);
    }
    // Si por error se abre modal grande de "Editar cita", cerrarlo antes de buscar "Modulo" del tooltip.
    try {
      const editClose = page.getByRole('button', { name: /cerrar/i }).first();
      if (await editClose.isVisible({ timeout: 260 })) {
        await editClose.click({ force: true, timeout: 700 });
        await waitForTimeoutRaw(page, 120);
      }
    } catch {}

    let forceLoopTry = 0;
    let clickLoopFlag = false;
    let modalClosedLoopFlag = false;
    do {
      forceLoopTry += 1;
      const slotFocus = await focusSavedAppointmentSlotForVideoModal(page, slot);
      console.log(
        `POST_SAVE_SLOT_REFOCUS attempt=${attemptNo}.${forceLoopTry} clicked=${slotFocus.clicked ? 1 : 0} opened=${slotFocus.opened ? 1 : 0} via=${slotFocus.via}`
      );

      const quickModalBefore = await isProgramadaQuickActionVisible(page);
      const quickModulo = await clickModuloFromSavedSlotQuickAction(page, slot, { appointmentNumber });
      clickLoopFlag = quickModulo.ok;
      clickedThisAttempt = clickLoopFlag;
      clickVia = quickModulo.via || 'quick_action_unknown';
      const quickModalAfter = await isProgramadaQuickActionVisible(page);
      modalClosedLoopFlag = !clickLoopFlag && quickModalBefore && !quickModalAfter;

      console.log(`POST_SAVE_QUICK_MODULO attempt=${attemptNo}.${forceLoopTry} ok=${quickModulo.ok ? 1 : 0} via=${clickVia}`);
      console.log(
        `MODULO_FORCE_LOOP attempt=${attemptNo}.${forceLoopTry} clicked=${clickLoopFlag ? 1 : 0} modal_before=${quickModalBefore ? 1 : 0} modal_after=${quickModalAfter ? 1 : 0} modal_closed=${modalClosedLoopFlag ? 1 : 0}`
      );
      console.log(`MODULO_BTN_FLAG attempt=${attemptNo}.${forceLoopTry} clicked=${clickedThisAttempt ? 1 : 0} source=quick_action`);

      if (!clickLoopFlag && modalClosedLoopFlag && forceLoopTry < POST_SAVE_MODAL_CLICK_LOOP_MAX) {
        await sleepRaw(Math.max(30, POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS));
      }
    } while (!clickLoopFlag && modalClosedLoopFlag && forceLoopTry < POST_SAVE_MODAL_CLICK_LOOP_MAX);

    if (!clickedThisAttempt && POST_SAVE_REQUIRE_ASSIGNED_MODAL) {
      const assignedModal = await tryOpenAssignedModalFromSavedSlot(page, slot);
      if (assignedModal.opened) {
        const assignedClick = await clickAbrirModuloDesdeNuevaCitaAsignada(page);
        if (assignedClick) {
          clickedThisAttempt = true;
          clickVia = `assigned_modal:${assignedModal.via || 'opened'}`;
          console.log(`POST_SAVE_ASSIGNED_MODULO attempt=${attemptNo} ok=1 via=${clickVia}`);
          console.log(`MODULO_BTN_FLAG attempt=${attemptNo} clicked=1 source=assigned_modal`);
        } else {
          console.log(`POST_SAVE_ASSIGNED_MODULO attempt=${attemptNo} ok=0 via=${assignedModal.via || 'opened_no_click'}`);
        }
      }
    }

    if (clickedThisAttempt) {
      moduloState.clicked = true;
      moduloState.clickVia = clickVia;
      if (await waitForModuloLoaded(page, `quick_action:${clickVia}`)) {
        moduloState.loaded = true;
        moduloState.loadedVia = `quick_action:${clickVia}`;
        console.log(
          `MODULO_BTN_FLAG_FINAL clicked=${moduloState.clicked ? 1 : 0} loaded=${moduloState.loaded ? 1 : 0} attempts=${attemptNo} click_via=${moduloState.clickVia} load_via=${moduloState.loadedVia}`
        );
        return true;
      }
      console.log(`MODULO_LOAD_FLAG attempt=${attemptNo} loaded=0 after_click=1 via=${clickVia}`);
    }

    if (!clickedThisAttempt) console.log(`MODULO_BTN_FLAG attempt=${attemptNo} clicked=0 source=no_click_path`);
    await sleepRaw(POST_SAVE_RETRY_INTERVAL_MS);
  }

  console.log(
    `MODULO_BTN_FLAG_FINAL clicked=${moduloState.clicked ? 1 : 0} loaded=${moduloState.loaded ? 1 : 0} attempts=${moduloState.attempts} click_via=${moduloState.clickVia} load_via=${moduloState.loadedVia}`
  );
  console.log('MODULO_AFTER_SAVE_FAIL');
  return false;
}

async function isNuevaCitaModalVisible(page) {
  if (isPageClosedSafe(page)) return false;
  const exactInputId = '#ctl00_nc002_MP_HOS930_MP_HOS930_panelExpExistente_MP_HOS930_altaclavedoc';

  // Deteccion fuerte por input exacto en documento principal.
  try {
    const input = page.locator(exactInputId).first();
    if ((await input.count()) > 0 && (await input.isVisible())) return true;
  } catch {}

  // Deteccion fuerte por input exacto dentro de iframes.
  try {
    for (const frame of page.frames()) {
      try {
        const input = frame.locator(exactInputId).first();
        if ((await input.count()) > 0 && (await input.isVisible())) return true;
      } catch {}
    }
  } catch {}

  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 80;
      };

      const dialogs = Array.from(
        document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')
      ).filter(visible);
      const exactInputId = 'ctl00_nc002_MP_HOS930_MP_HOS930_panelExpExistente_MP_HOS930_altaclavedoc';
      return dialogs.some((d) => {
        const t = normalize(d.textContent || '');
        const title = normalize(
          (d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || '')
        );
        const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
        if (!isNuevaCitaTitle) return false;
        if (d.querySelector(`#${exactInputId}`)) return true;
        if (t.includes('clave documento')) return true;
        if (t.includes('paciente') && t.includes('guardar')) return true;
        return false;
      });
    });
  } catch {
    return false;
  }
}

async function isNuevaCitaAsignadaModalVisible(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    return await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 80;
    };

    const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
    return dialogs.some((d) => normalize(d.textContent || '').includes('nueva cita asignada'));
    });
  } catch {
    return false;
  }
}

async function closeNuevaCitaAsignadaModal(page) {
  if (!(await isNuevaCitaAsignadaModalVisible(page))) return true;

  try {
    const closeBtn = page.getByRole('button', { name: /cerrar/i }).first();
    await closeBtn.waitFor({ state: 'visible', timeout: 1500 });
    await closeBtn.click({ force: true });
    await page.waitForTimeout(350);
  } catch {}

  if (await isNuevaCitaAsignadaModalVisible(page)) {
    await page.keyboard.press('Escape');
    await page.waitForTimeout(300);
  }

  return !(await isNuevaCitaAsignadaModalVisible(page));
}

async function clickAbrirModuloDesdeNuevaCitaAsignada(page) {
  if (!(await isNuevaCitaAsignadaModalVisible(page))) return false;

  const clicked = await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
    };

    const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
    const root = dialogs.find((d) => normalize(d.textContent || '').includes('nueva cita asignada'));
    if (!root) return false;

    const nodes = Array.from(root.querySelectorAll('button, a, span, input[type="button"]')).filter(visible);
    const btn = nodes.find((n) => {
      const t = normalize(n.textContent || n.value || '');
      return t.includes('abrir modulo') || t.includes('abrir módulo');
    });
    if (!(btn instanceof HTMLElement)) return false;
    btn.click();
    return true;
  });

  if (!clicked) return false;
  await page.waitForTimeout(900);
  return true;
}

async function isCatalogPacientesModalVisible(page) {
  if (isPageClosedSafe(page)) return false;
  const hasCatalogOnScope = async (scope) => {
    try {
      return await scope.evaluate(() => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const visible = (el) => {
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 80 && r.height > 40;
        };
        const looksLikeDialog = (el) => {
          if (!(el instanceof HTMLElement)) return false;
          const cls = normalize(el.className || '');
          return (
            el.matches?.('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, .RadWindow, [role="dialog"], .k-window-content, .k-animation-container, .ui-dialog') ||
            cls.includes('dialog') ||
            cls.includes('window') ||
            cls.includes('modal') ||
            cls.includes('popup') ||
            cls.includes('overlay')
          );
        };
        const hasCatalogText = (el) => {
          const text = normalize(el.textContent || '');
          const title = normalize(
            el.querySelector?.('.k-window-title, .k-dialog-title, .modal-title')?.textContent ||
            el.getAttribute?.('aria-label') ||
            el.getAttribute?.('title') ||
            ''
          );
          return text.includes('catalogo de pacientes') || title.includes('catalogo de pacientes');
        };

        const dialogCandidates = Array.from(
          document.querySelectorAll(
            '.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"], .k-window-content, .k-animation-container, .ui-dialog, [class*="dialog"], [class*="modal"], [class*="window"]'
          )
        )
          .filter(visible)
          .filter(looksLikeDialog)
          .slice(0, 180);

        if (dialogCandidates.some((d) => hasCatalogText(d))) return true;

        const textNodes = Array.from(document.querySelectorAll('h1,h2,h3,div,span,label,strong,b'))
          .filter(visible)
          .slice(0, 900);
        for (const node of textNodes) {
          const t = normalize(node.textContent || '');
          if (!t.includes('catalogo de pacientes')) continue;
          const dialogRoot =
            node.closest?.(
              '.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, .RadWindow, [role="dialog"], .k-window-content, .k-animation-container, .ui-dialog, [class*="dialog"], [class*="modal"], [class*="window"]'
            ) || node.parentElement;
          if (dialogRoot && visible(dialogRoot)) return true;
          return true;
        }
        return false;
      });
    } catch {
      return false;
    }
  };

  if (await hasCatalogOnScope(page)) return true;
  for (const frame of page.frames()) {
    try {
      if (await hasCatalogOnScope(frame)) return true;
    } catch {}
  }
  return false;
}

async function closeCatalogPacientesModal(page) {
  if (isPageClosedSafe(page)) return true;
  if (!(await isCatalogPacientesModalVisible(page))) return true;

  const tryCloseFromDom = async (scope) => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    const visibleDialog = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 90 && r.height > 50;
    };
    const visibleControl = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
    };
    const looksLikeDialog = (el) => {
      if (!(el instanceof HTMLElement)) return false;
      const cls = normalize(el.className || '');
      return (
        el.matches?.('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, .RadWindow, [role="dialog"], .k-window-content, .k-animation-container, .ui-dialog') ||
        cls.includes('dialog') ||
        cls.includes('window') ||
        cls.includes('modal') ||
        cls.includes('popup') ||
        cls.includes('overlay')
      );
    };
    const safeClick = (el) => {
      if (!(el instanceof HTMLElement)) return false;
      try {
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        el.click();
        return true;
      } catch {
        return false;
      }
    };

    const dialogs = Array.from(
      document.querySelectorAll(
        '.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"], .k-window-content, .k-animation-container, .ui-dialog, [class*="dialog"], [class*="modal"], [class*="window"]'
      )
    )
      .filter(visibleDialog)
      .filter(looksLikeDialog)
      .slice(0, 220);

    let dlg = dialogs.find((d) => {
      const txt = normalize(d.textContent || '');
      const title = normalize(
        d.querySelector?.('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute?.('aria-label') || d.getAttribute?.('title') || ''
      );
      return txt.includes('catalogo de pacientes') || title.includes('catalogo de pacientes');
    }) || null;

    if (!dlg) {
      const textNodes = Array.from(document.querySelectorAll('h1,h2,h3,div,span,label,strong,b'))
        .filter(visibleDialog)
        .slice(0, 1200);
      for (const node of textNodes) {
        const t = normalize(node.textContent || '');
        if (!t.includes('catalogo de pacientes')) continue;
        const parent =
          node.closest?.(
            '.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, .RadWindow, [role="dialog"], .k-window-content, .k-animation-container, .ui-dialog, [class*="dialog"], [class*="modal"], [class*="window"]'
          ) || node.parentElement;
        if (parent instanceof HTMLElement && visibleDialog(parent)) {
          dlg = parent;
          break;
        }
      }
    }
    if (!dlg) return false;

    const closeCandidates = [
      '.k-window-titlebar .k-window-titlebar-actions a',
      '.k-window-titlebar .k-window-titlebar-actions button',
      '.k-window-titlebar .k-window-titlebar-actions .k-window-action',
      '.k-window-titlebar .k-window-titlebar-actions .k-window-action',
      '.k-window-titlebar .k-window-actions a',
      '.k-window-titlebar .k-window-actions button',
      '.k-window-titlebar [aria-label="Close"]',
      '.k-window-titlebar .k-i-close',
      '.k-window-titlebar .k-svg-icon.k-svg-i-x, .k-window-titlebar .k-svg-i-x',
      '.rwTitlebar .rwCloseButton',
      '.k-dialog-titlebar .k-window-action',
      'button[aria-label="Close"]',
      'button[aria-label*="cerrar" i]',
      'button[aria-label*="close" i]',
      'button[title*="Cerrar"]',
      'button[title*="Close"]',
      'a[title*="Cerrar"]',
      'a[title*="Close"]',
      '[title*="cerrar" i]',
      '[title*="close" i]',
      '[aria-label*="cerrar" i]',
      '[aria-label*="close" i]',
      '.fa-times',
      '.icon-close'
    ];
    let closeBtn = null;
    for (const sel of closeCandidates) {
      const node = dlg.querySelector(sel);
      if (node instanceof HTMLElement && visibleControl(node)) {
        closeBtn = node;
        break;
      }
    }

    if (!closeBtn) {
      const allButtons = Array.from(dlg.querySelectorAll('button, a, span, div, i')).filter(visibleControl);
      closeBtn = allButtons.find((n) => {
        const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || '');
        return t === 'x' || t.includes('cerrar') || t.includes('close');
      }) || null;
    }

    if (closeBtn instanceof HTMLElement && safeClick(closeBtn)) return true;

    // Fallback: click en esquina superior derecha del modal (donde suele estar X).
    try {
      const rr = dlg.getBoundingClientRect();
      const x = Math.max(2, Math.round(rr.right - 26));
      const y = Math.max(2, Math.round(rr.top + 24));
      const target = document.elementFromPoint(x, y);
      if (target instanceof HTMLElement && safeClick(target)) return true;
    } catch {}

    return false;
  };

  for (let i = 0; i < 8; i += 1) {
    const scopes = [page, ...page.frames()];
    for (const scope of scopes) {
      try {
        await scope.evaluate(tryCloseFromDom);
      } catch {}
    }
    await page.waitForTimeout(160);

    if (!(await isCatalogPacientesModalVisible(page))) return true;

    try {
      const closeBtn = page.getByRole('button', { name: /cerrar|close|x/i }).first();
      await closeBtn.waitFor({ state: 'visible', timeout: 500 });
      await closeBtn.click({ force: true, timeout: 700 });
      await page.waitForTimeout(160);
    } catch {}

    if (!(await isCatalogPacientesModalVisible(page))) return true;

    try {
      await page.keyboard.press('Escape');
    } catch {}
    await page.waitForTimeout(220);
    if (!(await isCatalogPacientesModalVisible(page))) return true;
  }

  return !(await isCatalogPacientesModalVisible(page));
}

async function closeCatalogIfAppearsSoon(page, waitMs = 1600, reason = '') {
  const started = Date.now();
  while ((Date.now() - started) < waitMs) {
    if (await isCatalogPacientesModalVisible(page)) {
      console.log(`CATALOGO_PACIENTES_DETECTADO_LATE reason=${reason || 'unspecified'} -> cerrando`);
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 180);
      return true;
    }
    await waitForTimeoutRaw(page, 120);
  }
  return false;
}

async function getTopVisibleDialogInfo(page) {
  const readTopDialogOnScope = async (scope) => {
    return await scope.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 80;
      };
      const asNum = (raw) => {
        const n = Number.parseInt(String(raw || ''), 10);
        return Number.isFinite(n) ? n : -999999;
      };

      const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]'))
        .filter(visible)
        .map((d, idx) => {
          const title = normalize(
            d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || ''
          );
          const text = normalize(d.textContent || '');
          const zIndex = asNum(getComputedStyle(d).zIndex);
          let type = 'other';
          if (title === 'nueva cita' || title.startsWith('nueva cita')) type = 'nueva_cita';
          if (title.includes('nueva cita asignada') || text.includes('nueva cita asignada')) type = 'nueva_cita_asignada';
          if (title.includes('catalogo de pacientes') || text.includes('catalogo de pacientes')) type = 'catalogo_pacientes';
          return { idx, zIndex, title, type };
        });

      if (!dialogs.length) return { type: 'none', title: '', zIndex: -999999, idx: -1 };

      dialogs.sort((a, b) => {
        if (a.zIndex !== b.zIndex) return a.zIndex - b.zIndex;
        return a.idx - b.idx;
      });

      return dialogs[dialogs.length - 1];
    });
  };

  let top = { type: 'none', title: '', zIndex: -999999, idx: -1 };
  const scopes = [page, ...page.frames()];
  for (const scope of scopes) {
    try {
      const candidate = await readTopDialogOnScope(scope);
      if (!candidate || candidate.type === 'none') continue;
      if (candidate.zIndex > top.zIndex || (candidate.zIndex === top.zIndex && candidate.idx > top.idx)) {
        top = candidate;
      }
    } catch {}
  }
  return top;
}

async function ensureStrictNuevaCitaContext(page, stage = '') {
  if (!STRICT_NUEVA_CITA_MODAL) return await isNuevaCitaModalVisible(page);

  for (let i = 0; i < 3; i += 1) {
    const top = await getTopVisibleDialogInfo(page);
    const nuevaVisible = await isNuevaCitaModalVisible(page);
    const catalogVisible = await isCatalogPacientesModalVisible(page);

    if (top.type === 'catalogo_pacientes' || catalogVisible) {
      console.log(
        `STRICT_MODAL_GUARD stage=${stage || '-'} action=close_catalog attempt=${i + 1} title="${top.title || ''}"`
      );
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 150);
      continue;
    }

    if (top.type === 'nueva_cita') return true;
    if (nuevaVisible) return true;

    if (top.type === 'none') {
      await waitForTimeoutRaw(page, 140);
      continue;
    }

    if (top.type === 'nueva_cita_asignada') {
      console.log(
        `STRICT_MODAL_GUARD stage=${stage || '-'} action=block top=${top.type} title="${top.title || ''}"`
      );
      return false;
    }

    console.log(
      `STRICT_MODAL_GUARD stage=${stage || '-'} action=wait top=${top.type} title="${top.title || ''}" attempt=${i + 1}`
    );
    await waitForTimeoutRaw(page, 120);
  }

  return (await isNuevaCitaModalVisible(page)) && !(await isCatalogPacientesModalVisible(page));
}

async function isAlreadyScheduledAlertVisible(page) {
  return await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 80;
    };

    const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
    return dialogs.some((d) => normalize(d.textContent || '').includes('el paciente ya tiene una cita programada'));
  });
}

async function isPatientNotFoundAlertVisible(page) {
  const hasOnScope = async (scope) => {
    return await scope.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 40 && r.height > 20;
      };
      const matchText = (s) => {
        const t = normalize(s || '');
        return t.includes('paciente no encontrado') ||
          t.includes('no se encontro paciente') ||
          t.includes('clave documento no encontrada') ||
          t.includes('error 404') ||
          /^404\b/.test(t);
      };

      const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
      if (dialogs.some((d) => matchText(d.textContent || ''))) return true;

      const banners = Array.from(document.querySelectorAll('.k-notification, .toast, .alert, .swal2-container, .ajs-message, [role="alert"]')).filter(visible);
      return banners.some((b) => matchText(b.textContent || ''));
    });
  };

  if (await hasOnScope(page)) return true;
  for (const frame of page.frames()) {
    try {
      if (await hasOnScope(frame)) return true;
    } catch {}
  }
  return false;
}

async function closePatientNotFoundAlert(page) {
  const closeOnScope = async (scope) => {
    return await scope.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
      };
      const isTargetText = (s) => {
        const t = normalize(s || '');
        return t.includes('paciente no encontrado') ||
          t.includes('no se encontro paciente') ||
          t.includes('clave documento no encontrada') ||
          t.includes('error 404') ||
          /^404\b/.test(t);
      };
      const isCloseText = (s) => {
        const t = normalize(s || '');
        return t.includes('aceptar') || t.includes('cerrar') || t === 'ok' || t.includes('continuar');
      };

      const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
      for (const d of dialogs) {
        if (!isTargetText(d.textContent || '')) continue;
        const btns = Array.from(d.querySelectorAll('button, a, input[type="button"], input[type="submit"]')).filter(visible);
        const btn = btns.find((b) => isCloseText((b.textContent || b.value || b.getAttribute('aria-label') || b.getAttribute('title') || ''))) || btns[0];
        if (btn instanceof HTMLElement) {
          btn.click();
          return true;
        }
      }
      return false;
    });
  };

  let closedSomething = false;
  try {
    closedSomething = (await closeOnScope(page)) || closedSomething;
  } catch {}
  for (const frame of page.frames()) {
    try {
      closedSomething = (await closeOnScope(frame)) || closedSomething;
    } catch {}
  }

  if (!closedSomething) {
    try { await page.keyboard.press('Escape'); } catch {}
  }
  await waitForTimeoutRaw(page, 180);

  if (await isCatalogPacientesModalVisible(page)) {
    await closeCatalogPacientesModal(page);
    await waitForTimeoutRaw(page, 180);
  }

  return !(await isPatientNotFoundAlertVisible(page));
}

async function closeAlreadyScheduledAlert(page) {
  if (!(await isAlreadyScheduledAlertVisible(page))) return false;

  try {
    const btn = page.getByRole('button', { name: /ok|aceptar|cerrar|continuar/i }).first();
    await btn.waitFor({ state: 'visible', timeout: 1200 });
    await btn.click({ force: true, timeout: 1200 });
  } catch {}

  try {
    await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 80;
      };
      const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
      const alert = dialogs.find((d) => normalize(d.textContent || '').includes('el paciente ya tiene una cita programada'));
      if (!alert) return;
      const close =
        alert.querySelector('button, a, .k-button, .ajs-button') ||
        alert.querySelector('.k-window-actions a, .k-window-actions button');
      if (close instanceof HTMLElement) close.click();
    });
  } catch {}

  await waitForTimeoutRaw(page, 250);
  return !(await isAlreadyScheduledAlertVisible(page));
}

async function ensureNuevaCitaModalOpenForRetry(page) {
  if (await isNuevaCitaModalVisible(page)) return true;

  for (let attempt = 0; attempt < 4; attempt += 1) {
    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 220);
    }
    if (await isNuevaCitaAsignadaModalVisible(page)) {
      await closeNuevaCitaAsignadaModal(page);
      await waitForTimeoutRaw(page, 180);
    }

    const fast = await fastOpenNuevaCitaFromCandidates(page);
    if (!fast.ok) continue;

    const waited = await waitForNuevaCitaModal(page, 3200);
    if (waited.opened) return true;

    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 220);
      continue;
    }
    if (await isNuevaCitaAsignadaModalVisible(page)) {
      await closeNuevaCitaAsignadaModal(page);
      await waitForTimeoutRaw(page, 180);
      continue;
    }
  }

  return false;
}

async function reopenNuevaCitaOnPreferredSlot(page, slot) {
  if (!slot) return false;
  if (await isNuevaCitaModalVisible(page)) return true;

  for (let attempt = 0; attempt < 4; attempt += 1) {
    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 180);
    }
    if (await isNuevaCitaAsignadaModalVisible(page)) {
      await closeNuevaCitaAsignadaModal(page);
      await waitForTimeoutRaw(page, 160);
    }

    try {
      if (typeof slot.x === 'number' && typeof slot.y === 'number') {
        await page.mouse.click(slot.x, slot.y, { delay: 35 });
        await waitForTimeoutRaw(page, 140);
        await page.mouse.click(slot.x, slot.y, { delay: 35 });
        await waitForTimeoutRaw(page, 240);

        if (await isNuevaCitaModalVisible(page)) return true;
        if (await isCatalogPacientesModalVisible(page)) continue;

        await page.mouse.click(slot.x, slot.y, { clickCount: 2, delay: 70 });
        await waitForTimeoutRaw(page, 260);
        if (await isNuevaCitaModalVisible(page)) return true;
      }
    } catch {}

    try {
      if (slot.selector && Number.isInteger(slot.domIdx)) {
        const cell = page.locator(slot.selector).nth(slot.domIdx);
        await cell.click({ force: true, timeout: 900 });
        await waitForTimeoutRaw(page, 150);
        await cell.click({ force: true, timeout: 900 });
        await waitForTimeoutRaw(page, 200);
        if (await isNuevaCitaModalVisible(page)) return true;
        await cell.click({ force: true, timeout: 900, clickCount: 2 });
        await waitForTimeoutRaw(page, 260);
        if (await isNuevaCitaModalVisible(page)) return true;
      }
    } catch {}

    const waited = await waitForNuevaCitaModal(page, 1100);
    if (waited.opened) return true;
    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 160);
      continue;
    }
  }

  return false;
}

async function getCitaGeneradaSuccessInfo(page) {
  return await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 80 && r.height > 24;
    };

    const nodes = Array.from(document.querySelectorAll('div, span, p, li')).filter(visible);
    const hit = nodes.find((n) => {
      const t = normalize(n.textContent || '');
      return t.includes('se ha generado la cita con el numero');
    });
    if (!hit) return { ok: false, number: '' };

    const text = (hit.textContent || '').trim();
    const m = text.match(/\[\s*(\d+)\s*\]/);
    return { ok: true, number: m ? m[1] : '', text };
  });
}

async function waitForCitaGeneradaSuccessInfo(page, timeoutMs = 5200) {
  const started = Date.now();
  while ((Date.now() - started) < timeoutMs) {
    const info = await getCitaGeneradaSuccessInfo(page);
    if (info.ok) return info;
    await waitForTimeoutRaw(page, 170);
  }
  return { ok: false, number: '' };
}

async function setClaveDocumentoAndTriggerSearch(page, key) {
  const keyValue = String(key || '').trim().toUpperCase();
  if (!keyValue) return { ok: false, reason: 'empty_key' };

  const exactId = '#ctl00_nc002_MP_HOS930_MP_HOS930_panelExpExistente_MP_HOS930_altaclavedoc';
  const exactIdRaw = exactId.replace('#', '');
  const normalizeKey = (v) => String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const keyNorm = normalizeKey(keyValue);
  const shouldSearch = false; // Flujo validado: NO usar lupa/buscar; va por Comentarios + Guardar.

  let strictContextOk = false;
  for (let i = 0; i < 4; i += 1) {
    if (await ensureStrictNuevaCitaContext(page, `set_key_start:${i + 1}`)) {
      strictContextOk = true;
      break;
    }
    await waitForTimeoutRaw(page, 150);
  }
  if (!strictContextOk) {
    return { ok: false, reason: 'invalid_modal_context_before_key' };
  }

  const fillKeyOnScope = async (scope) => {
    const input = scope.locator(`${exactId}:visible`).first();
    await input.waitFor({ state: 'visible', timeout: 2200 });
    await input.click({ force: true, timeout: 1200 });
    await input.fill('');
    await input.type(keyValue, { delay: 28 });
    await page.waitForTimeout(80);
    await input.blur();
    await page.waitForTimeout(90);

    let currentValue = '';
    try {
      currentValue = await input.inputValue();
    } catch {}
    let keyAccepted = normalizeKey(currentValue).includes(keyNorm);

    // Fallback para controles Kendo MaskedTextBox que no persisten con fill/type normal.
    if (!keyAccepted) {
      const patched = await scope.evaluate(({ id, val }) => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .trim();
        const visible = (el) => {
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 10 && r.height > 10;
        };

        const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
        const root = dialogs.find((d) => {
          const title = normalize(
            d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || ''
          );
          return title === 'nueva cita' || title.startsWith('nueva cita');
        });
        if (!root) return { ok: false, value: '' };

        const direct = root.querySelector(`#${id}`);
        const semantic = Array.from(root.querySelectorAll('input[type="text"], input:not([type])')).find((n) => (
          n instanceof HTMLInputElement &&
          visible(n) &&
          /clave|documento|expediente/i.test(`${n.id} ${n.name} ${n.placeholder}`)
        ));
        const input = direct || semantic;
        if (!(input instanceof HTMLInputElement)) return { ok: false, value: '' };

        try {
          if (window.$ && input.id) {
            const jq = window.$(`#${input.id}`);
            const kMasked = jq && jq.data && jq.data('kendoMaskedTextBox');
            if (kMasked && typeof kMasked.value === 'function') {
              kMasked.value(val);
              if (typeof kMasked.trigger === 'function') kMasked.trigger('change');
            }
          }
        } catch {}

        const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
        input.focus();
        if (nativeSetter) nativeSetter.call(input, val);
        else input.value = val;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new Event('blur', { bubbles: true }));
        return { ok: true, value: input.value || '' };
      }, { id: exactIdRaw, val: keyValue });

      currentValue = patched?.value || currentValue;
      keyAccepted = normalizeKey(currentValue).includes(keyNorm);
    }

    let clickedSearch = false;
    if (shouldSearch) {
      try {
        const searchBtn = scope.locator(
          '[id*="buscaexpediente"]:visible, [name*="buscaexpediente"]:visible, a[href*="buscaexpediente"]:visible'
        ).first();
        await searchBtn.waitFor({ state: 'visible', timeout: 1200 });
        await searchBtn.click({ force: true, timeout: 1000 });
        clickedSearch = true;
      } catch {}
    }

    return {
      ok: keyAccepted,
      reason: keyAccepted ? undefined : 'key_not_set_visible_input',
      keyAccepted,
      keyInputId: exactIdRaw,
      valueSet: keyValue,
      currentValue,
      clickedSearch
    };
  };

  // Camino principal: documento actual.
  try {
    const direct = await fillKeyOnScope(page);
    if (direct.ok) {
      if (STRICT_NUEVA_CITA_MODAL && (await isCatalogPacientesModalVisible(page))) {
        await closeCatalogPacientesModal(page);
        return { ...direct, ok: false, keyAccepted: false, reason: 'catalog_opened_after_key_set' };
      }
      return direct;
    }
  } catch {}

  // Variante: el modal puede estar dentro de un iframe.
  for (const frame of page.frames()) {
    try {
      const fromFrame = await fillKeyOnScope(frame);
      if (fromFrame.ok) {
        if (STRICT_NUEVA_CITA_MODAL && (await isCatalogPacientesModalVisible(page))) {
          await closeCatalogPacientesModal(page);
          return { ...fromFrame, ok: false, keyAccepted: false, reason: 'catalog_opened_after_key_set' };
        }
        return fromFrame;
      }
    } catch {}
  }

  // Fallback: inyección DOM/JS dentro del modal.
  const fallbackResult = await page.evaluate(({ keyValue, shouldSearch }) => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
    };

    const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
    const root = dialogs.find((d) => {
      const t = normalize(d.textContent || '');
      const title = normalize(
        (d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || '')
      );
      const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
      return isNuevaCitaTitle && t.includes('clave documento');
    });
    if (!root) return { ok: false, reason: 'nueva_cita_modal_not_found' };

    const exactId = 'ctl00_nc002_MP_HOS930_MP_HOS930_panelExpExistente_MP_HOS930_altaclavedoc';
    const exactInput = root.querySelector(`#${exactId}`);
    const textInputs = Array.from(root.querySelectorAll('input[type="text"], input:not([type])'))
      .filter(visible)
      .filter((i) => !i.disabled);
    if (!textInputs.length) return { ok: false, reason: 'key_input_not_found' };

    const bySemantic = textInputs.find((i) => /clave|documento|expediente/i.test(`${i.id} ${i.name} ${i.placeholder}`));
    const keyInput = exactInput || bySemantic || textInputs[0];
    const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;

    // Si es Kendo MaskedTextBox, usar la API del widget para persistir valor internamente.
    try {
      if (window.$ && keyInput.id) {
        const jq = window.$(`#${keyInput.id}`);
        const kMasked = jq && jq.data && jq.data('kendoMaskedTextBox');
        if (kMasked && typeof kMasked.value === 'function') {
          kMasked.value(keyValue);
          if (typeof kMasked.trigger === 'function') kMasked.trigger('change');
        }
      }
    } catch {}

    keyInput.focus();
    if (nativeSetter) nativeSetter.call(keyInput, '');
    else keyInput.value = '';
    keyInput.dispatchEvent(new Event('input', { bubbles: true }));
    if (nativeSetter) nativeSetter.call(keyInput, keyValue);
    else keyInput.value = keyValue;
    keyInput.dispatchEvent(new Event('input', { bubbles: true }));
    keyInput.dispatchEvent(new Event('change', { bubbles: true }));
    keyInput.dispatchEvent(new Event('blur', { bubbles: true }));

    let searchTarget = null;
    let clickedSearch = false;
    if (shouldSearch) {
      searchTarget =
        root.querySelector('[id*="buscaexpediente"]') ||
        root.querySelector('[name*="buscaexpediente"]') ||
        root.querySelector('a[href*="buscaexpediente"]');
      if (searchTarget instanceof HTMLElement) {
        searchTarget.click();
        searchTarget.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        clickedSearch = true;
      }
    }

    const normalizeKey = (v) => String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    const keyAccepted = normalizeKey(keyInput.value || '').includes(normalizeKey(keyValue));
    return {
      ok: keyAccepted,
      reason: keyAccepted ? undefined : 'js_fallback_key_not_set',
      keyAccepted,
      keyInputId: keyInput.id || '',
      valueSet: keyValue,
      currentValue: keyInput.value || '',
      clickedSearch
    };
  }, { keyValue, shouldSearch });

  if (STRICT_NUEVA_CITA_MODAL && (await isCatalogPacientesModalVisible(page))) {
    await closeCatalogPacientesModal(page);
    return { ...fallbackResult, ok: false, keyAccepted: false, reason: 'catalog_opened_after_key_set' };
  }

  return fallbackResult;
}

async function disableSearchControlsInNuevaCita(page) {
  const disableOnScope = async (scope) => {
    return await scope.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 18 && r.height > 10;
      };
      const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
      const root = dialogs.find((d) => {
        const t = normalize(d.textContent || '');
        const title = normalize(
          (d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || '')
        );
        const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
        return isNuevaCitaTitle && t.includes('clave documento');
      });
      if (!root) return 0;

      const candidates = Array.from(
        root.querySelectorAll('[id*="buscaexpediente"], [name*="buscaexpediente"], a[href*="buscaexpediente"], .k-i-search, .k-icon.k-i-search, [aria-label*="buscar"]')
      );
      let changed = 0;
      for (const n of candidates) {
        if (!(n instanceof HTMLElement)) continue;
        if (n.dataset && n.dataset.codexSearchDisabled === '1') continue;
        if (n.dataset) n.dataset.codexSearchDisabled = '1';
        n.style.pointerEvents = 'none';
        n.style.opacity = n.style.opacity || '0.85';
        n.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          e.stopImmediatePropagation?.();
        }, true);
        n.addEventListener('mousedown', (e) => {
          e.preventDefault();
          e.stopPropagation();
          e.stopImmediatePropagation?.();
        }, true);
        if ('disabled' in n) {
          try { n.disabled = true; } catch {}
        }
        changed += 1;
      }
      return changed;
    });
  };

  let changed = 0;
  try { changed += await disableOnScope(page); } catch {}
  for (const frame of page.frames()) {
    try { changed += await disableOnScope(frame); } catch {}
  }
  if (changed > 0) {
    console.log(`SEARCH_CONTROL_DISABLED count=${changed}`);
  }
  return changed > 0;
}

async function installCatalogAutoCloseGuard(page) {
  const installOnScope = async (scope) => {
    return await scope.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 80 && r.height > 40;
      };

      try {
        if (window.__codexCatalogGuardTimer) return false;
      } catch {}

      const closeCatalogNow = () => {
        try {
          const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
          for (const d of dialogs) {
            const text = normalize(d.textContent || '');
            const title = normalize(
              d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || ''
            );
            if (!(title.includes('catalogo de pacientes') || text.includes('catalogo de pacientes'))) continue;

            const closeBtn =
              d.querySelector('.k-window-actions a, .k-window-actions button, .rwCloseButton, [aria-label*="cerrar"], [title*="cerrar"]') ||
              d.querySelector('button, a');
            if (closeBtn instanceof HTMLElement) {
              closeBtn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
              closeBtn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
              closeBtn.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
            }
            d.style.display = 'none';
            d.style.visibility = 'hidden';
          }
        } catch {}
      };

      window.__codexCatalogGuardTimer = window.setInterval(closeCatalogNow, 80);
      window.setTimeout(() => {
        try {
          if (window.__codexCatalogGuardTimer) {
            clearInterval(window.__codexCatalogGuardTimer);
            window.__codexCatalogGuardTimer = null;
          }
        } catch {}
      }, 30000);

      closeCatalogNow();
      return true;
    });
  };

  let installed = false;
  try { installed = (await installOnScope(page)) || installed; } catch {}
  for (const frame of page.frames()) {
    try { installed = (await installOnScope(frame)) || installed; } catch {}
  }
  if (installed) console.log('CATALOG_AUTOCLOSE_GUARD=ON');
  return installed;
}

async function fillComentarioTextInNuevaCita(page, value = COMMENT_TEXT) {
  if (!(await ensureStrictNuevaCitaContext(page, 'comentarios_fill_start'))) return false;
  const commentValue = String(value || '').trim() || 'TEST';

  // Camino directo (más estable): el textarea de Comentarios suele estar visible en el modal.
  try {
    const directTextArea = page.locator('textarea:visible').first();
    if ((await directTextArea.count()) > 0 && (await directTextArea.isVisible())) {
      await directTextArea.fill(commentValue);
      await waitForTimeoutRaw(page, 120);
      return true;
    }
  } catch {}
  for (const frame of page.frames()) {
    try {
      const directTextArea = frame.locator('textarea:visible').first();
      if ((await directTextArea.count()) > 0 && (await directTextArea.isVisible())) {
        await directTextArea.fill(commentValue);
        await waitForTimeoutRaw(page, 120);
        return true;
      }
    } catch {}
  }

  const fillOnScope = async (scope) => {
    return await scope.evaluate(({ commentValue }) => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
      };
      const setInputLikeValue = (input, val) => {
        const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
        const nativeTextAreaSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set;
        input.focus();
        if (input instanceof HTMLTextAreaElement && nativeTextAreaSetter) nativeTextAreaSetter.call(input, val);
        else if (input instanceof HTMLInputElement && nativeSetter) nativeSetter.call(input, val);
        else input.value = val;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new Event('blur', { bubbles: true }));
      };

      const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
      const root = dialogs.find((d) => {
        const t = normalize(d.textContent || '');
        const title = normalize(
          d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || ''
        );
        const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
        return isNuevaCitaTitle && t.includes('clave documento');
      });
      if (!root) return false;

      const fields = Array.from(
        root.querySelectorAll('textarea:not([disabled]), input[type="text"]:not([disabled]), [contenteditable="true"]')
      ).filter(visible);
      if (!fields.length) return false;

      const bySemantic = fields.find((f) => {
        const meta = normalize(
          `${f.id || ''} ${f.getAttribute?.('name') || ''} ${f.getAttribute?.('placeholder') || ''} ${f.className || ''}`
        );
        return meta.includes('coment') || meta.includes('observ') || meta.includes('nota') || meta.includes('descripcion');
      });
      const target = bySemantic || fields.find((f) => f instanceof HTMLTextAreaElement) || fields[0];
      if (!target) return false;

      if (target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement) {
        setInputLikeValue(target, commentValue);
        return true;
      }
      if (target instanceof HTMLElement && target.isContentEditable) {
        target.focus();
        target.textContent = commentValue;
        target.dispatchEvent(new Event('input', { bubbles: true }));
        target.dispatchEvent(new Event('blur', { bubbles: true }));
        return true;
      }
      return false;
    }, { commentValue });
  };

  for (let i = 0; i < 4; i += 1) {
    try {
      if (await fillOnScope(page)) return true;
    } catch {}
    for (const frame of page.frames()) {
      try {
        if (await fillOnScope(frame)) return true;
      } catch {}
    }
    await waitForTimeoutRaw(page, 180);
  }
  return false;
}

async function clickComentariosInNuevaCita(page) {
  if (!(await ensureStrictNuevaCitaContext(page, 'comentarios_start'))) return false;
  if (await isCatalogPacientesModalVisible(page)) {
    await closeCatalogPacientesModal(page);
    return false;
  }

  const clickOnScope = async (scope) => {
    return await scope.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
      };

      const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
      const root = dialogs.find((d) => {
        const t = normalize(d.textContent || '');
        const title = normalize(
          d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || ''
        );
        const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
        return isNuevaCitaTitle && t.includes('clave documento');
      });
      if (!root) return false;

      const candidates = Array.from(
        root.querySelectorAll(
          'button, a, label, [id*="coment"], [name*="coment"], [title*="Coment"], [aria-label*="Coment"]'
        )
      ).filter(visible);
      const target = candidates.find((n) => {
        const text = normalize(`${n.textContent || ''} ${n.getAttribute?.('title') || ''} ${n.getAttribute?.('aria-label') || ''}`);
        const idName = normalize(`${n.id || ''} ${n.getAttribute?.('name') || ''}`);
        return text.includes('coment') || text.includes('observ') || idName.includes('coment') || idName.includes('observ');
      });

      if (target instanceof HTMLElement) {
        target.click();
        target.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        return true;
      }

      const textarea = root.querySelector('textarea:not([disabled]), [contenteditable="true"]');
      if (textarea instanceof HTMLElement) {
        textarea.focus();
        return true;
      }
      return false;
    });
  };

  let clicked = false;
  try {
    clicked = await clickOnScope(page);
  } catch {}
  if (!clicked) {
    for (const frame of page.frames()) {
      try {
        if (await clickOnScope(frame)) {
          clicked = true;
          break;
        }
      } catch {}
    }
  }

  await waitForTimeoutRaw(page, 220);
  const filled = await fillComentarioTextInNuevaCita(page, COMMENT_TEXT);
  if (filled) console.log(`COMENTARIO_SET_OK value="${COMMENT_TEXT}"`);
  return clicked || filled;
}

async function isPatientLoadedInNuevaCita(page) {
  if (!(await ensureStrictNuevaCitaContext(page, 'check_patient_loaded'))) return false;
  return await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
    };

    const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
    const root = dialogs.find((d) => {
      const t = normalize(d.textContent || '');
      const title = normalize(
        (d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || '')
      );
      const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
      return isNuevaCitaTitle && t.includes('clave documento');
    });
    if (!root) return false;

    const disabledFilled = Array.from(root.querySelectorAll('input[type="text"], input:not([type])')).some((i) => {
      if (!(i instanceof HTMLInputElement)) return false;
      const v = (i.value || '').trim();
      return i.disabled && v.length >= 4;
    });

    return disabledFilled;
  });
}

async function isGuardarEnabledInNuevaCita(page) {
  if (!(await ensureStrictNuevaCitaContext(page, 'check_guardar_enabled'))) return false;

  // Camino directo: botón Guardar visible en modal.
  try {
    const btn = page.getByRole('button', { name: /guardar/i }).first();
    if ((await btn.count()) > 0 && (await btn.isVisible())) {
      return !(await btn.isDisabled());
    }
  } catch {}
  for (const frame of page.frames()) {
    try {
      const btn = frame.getByRole('button', { name: /guardar/i }).first();
      if ((await btn.count()) > 0 && (await btn.isVisible())) {
        return !(await btn.isDisabled());
      }
    } catch {}
  }

  return await page.evaluate(() => {
    const normalize = (s) =>
      (s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
    const visible = (el) => {
      if (!el) return false;
      const st = getComputedStyle(el);
      const r = el.getBoundingClientRect();
      return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
    };

    const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
    const root = dialogs.find((d) => {
      const t = normalize(d.textContent || '');
      const title = normalize(
        (d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || '')
      );
      const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
      return isNuevaCitaTitle && t.includes('clave documento');
    });
    if (!root) return false;

    const nodes = Array.from(root.querySelectorAll('button, a, span, input[type="button"], input[type="submit"], [id*="guardar"], [name*="guardar"]')).filter(visible);
    let btn = nodes.find((n) => {
      const t = normalize(n.textContent || n.value || '');
      if (t.includes('guardar')) return true;
      const idName = normalize(`${n.id || ''} ${n.getAttribute?.('name') || ''} ${n.getAttribute?.('title') || ''}`);
      return idName.includes('guardar');
    });
    if (!btn) {
      const footerButtons = Array.from(
        root.querySelectorAll('.k-window-actions button, .k-window-actions a, .k-window-actions input, .modal-footer button, .modal-footer a, .modal-footer input')
      ).filter(visible);
      btn =
        footerButtons.find((n) => {
          const t = normalize(n.textContent || n.value || '');
          return !t.includes('cerrar') && !t.includes('cancelar') && !t.includes('salir');
        }) || footerButtons[footerButtons.length - 1];
    }
    if (!btn) return false;
    return !(btn.disabled || btn.getAttribute('aria-disabled') === 'true');
  });
}

async function waitForNuevaCitaKeyResolution(page, timeoutMs = KEY_RESOLUTION_TIMEOUT_MS) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (await isCatalogPacientesModalVisible(page)) {
      return { ok: true, state: 'catalog_opened', elapsedMs: Date.now() - started };
    }
    if (!(await isNuevaCitaModalVisible(page))) {
      return { ok: true, state: 'modal_closed', elapsedMs: Date.now() - started };
    }
    if (await isPatientLoadedInNuevaCita(page)) {
      return { ok: true, state: 'patient_loaded', elapsedMs: Date.now() - started };
    }
    if (await isGuardarEnabledInNuevaCita(page)) {
      return { ok: true, state: 'guardar_enabled', elapsedMs: Date.now() - started };
    }
    await waitForTimeoutRaw(page, 220);
  }
  return { ok: false, state: 'timeout', elapsedMs: Date.now() - started };
}

async function debugDumpNuevaCitaControls(page, stage = '') {
  if (!DEBUG_NUEVA_CITA_CONTROLS) return;
  try {
    const dump = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };
      const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
      const root = dialogs.find((d) => {
        const t = normalize(d.textContent || '');
        const title = normalize(
          d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || ''
        );
        const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
        return isNuevaCitaTitle && t.includes('clave documento');
      });
      if (!root) return null;

      const rows = [];
      const nodes = Array.from(
        root.querySelectorAll('button, a, label, textarea, input, [contenteditable=\"true\"], [id], [name]')
      ).filter(visible);
      for (const n of nodes.slice(0, 200)) {
        const text = (n.textContent || n.value || '').replace(/\s+/g, ' ').trim().slice(0, 70);
        const meta = `${n.tagName || ''} id=${n.id || ''} name=${n.getAttribute?.('name') || ''} title=${n.getAttribute?.('title') || ''} aria=${n.getAttribute?.('aria-label') || ''}`.trim();
        if (!text && !/coment|observ|nota|guardar|clave|doc|exped/i.test(normalize(meta))) continue;
        rows.push({ text, meta });
      }
      return rows;
    });
    if (!dump) {
      console.log(`DEBUG_NUEVA_CITA_CONTROLS stage=${stage} root=none`);
      return;
    }
    console.log(`DEBUG_NUEVA_CITA_CONTROLS stage=${stage} items=${dump.length}`);
    for (const row of dump) {
      console.log(`DEBUG_CTRL text="${row.text}" meta="${row.meta}"`);
    }
  } catch (e) {
    console.log(`DEBUG_NUEVA_CITA_CONTROLS stage=${stage} error=${e?.message || e}`);
  }
}

async function loadFirstPatientFromKeys(page, keys, options = {}) {
  const preferredSlot = options.preferredSlot || null;
  const keyPlan = buildPatientKeyAttemptPlan(keys);
  const keysToTry = keyPlan.slice(0, MAX_KEY_ATTEMPTS);
  const makeAttemptResult = (key, appointmentNumber = '') => ({
    key: String(key || '').trim(),
    appointmentNumber: String(appointmentNumber || '').trim()
  });
  let saveAttemptInProgress = false;
  let catalogHits = 0;
  const reopenModalForRetry = async (reason) => {
    if (preferredSlot) {
      const reopenedSameSlot = await reopenNuevaCitaOnPreferredSlot(page, preferredSlot);
      if (reopenedSameSlot) return true;

      if (STRICT_PREFERRED_SLOT) {
        console.log(`REOPEN_STRICT_FAIL reason=${reason}`);
        return false;
      }

      const reopenedAny = await ensureNuevaCitaModalOpenForRetry(page);
      if (reopenedAny) {
        console.log(`REOPEN_FALLBACK_ANY_SLOT reason=${reason}`);
        return true;
      }
      return false;
    }
    return await ensureNuevaCitaModalOpenForRetry(page);
  };
  const handleCatalogLoop = async (reason, key = '') => {
    catalogHits += 1;
    const closedCatalog = await closeCatalogPacientesModal(page);
    if (!closedCatalog) {
      throw new Error(`No se pudo cerrar "Catálogo de pacientes" (${reason}) con clave "${key}".`);
    }
    if (key) rememberKeyOutcome(key, `catalog_opened:${reason}`);
    await waitForTimeoutRaw(page, 220);
    if (catalogHits >= CATALOG_LOOP_MAX) {
      throw new Error(
        `Catalogo de pacientes recurrente (${catalogHits}) en ${reason}. Reinicio recomendado desde login.`
      );
    }
  };
  const recoverNuevaCitaAfterReject = async (reason, key = '') => {
    for (let i = 0; i < 3; i += 1) {
      if (await isNuevaCitaModalVisible(page)) return true;
      if (await isCatalogPacientesModalVisible(page)) {
        await handleCatalogLoop(`${reason}_catalog_retry_${i + 1}`, key);
      }
      if (await isNuevaCitaAsignadaModalVisible(page)) {
        await closeNuevaCitaAsignadaModal(page);
        await waitForTimeoutRaw(page, 160);
      }
      const reopened = await reopenModalForRetry(`${reason}_retry_${i + 1}`);
      if (reopened && (await isNuevaCitaModalVisible(page))) return true;
      await waitForTimeoutRaw(page, 180);
    }
    return await isNuevaCitaModalVisible(page);
  };
  console.log(
    `KEY_ATTEMPTS plan=${keysToTry.length} limit=${MAX_KEY_ATTEMPTS} prioritize_recent=${PRIORITIZE_RECENT_KEYS ? 1 : 0} mode=${KEY_SELECTION_MODE} seed=${KEY_RANDOM_SEED || '-'}`
  );
  if (KEY_HEALTH_STATE.enabled) {
    const blockedInPlan = keysToTry.reduce((acc, k) => {
      const m = getKeyHealthMetrics(k);
      return acc + (m.blocked ? 1 : 0);
    }, 0);
    console.log(
      `KEY_HEALTH_PLAN enabled=1 file="${KEY_HEALTH_FILE}" ttl_h=${KEY_HEALTH_TTL_HOURS} hard_block=${KEY_HARD_BLOCK_THRESHOLD} blocked_in_plan=${blockedInPlan} records=${KEY_HEALTH_STATE.records.length}`
    );
  }

  for (const key of keysToTry) {
    if (saveAttemptInProgress) {
      const recovered = await reopenModalForRetry('save_flag_recovery');
      saveAttemptInProgress = false;
      if (!recovered) {
        throw new Error('No se pudo recuperar modal tras save flag activo.');
      }
    }

    if (await isCatalogPacientesModalVisible(page)) {
      console.log('CATALOGO_PACIENTES_DETECTADO antes de probar clave, cerrando...');
      await handleCatalogLoop('before_key', key);
    }
    if (!(await isNuevaCitaModalVisible(page))) {
      const reopened = await reopenModalForRetry('before_key');
      if (!reopened) {
        throw new Error('No se pudo abrir/reabrir el modal "Nueva cita" antes de intentar clave.');
      }
      await waitForTimeoutRaw(page, 250);
    }

    await installCatalogAutoCloseGuard(page);
    // Guardia dura: deshabilitar lupa/buscar para evitar apertura accidental de catálogo.
    await disableSearchControlsInNuevaCita(page);

    console.log(`INTENTANDO_CLAVE_PACIENTE "${key}"`);
    const result = await setClaveDocumentoAndTriggerSearch(page, key);
    if (!result.ok || !result.keyAccepted) {
      console.log(`CLAVE_RECHAZADA "${key}" reason=${result.reason || 'unknown'}`);
      rememberKeyOutcome(key, `key_rejected:${result.reason || 'unknown'}`);
      if (await isCatalogPacientesModalVisible(page)) {
        console.log(`CATALOGO_PACIENTES_DETECTADO tras rechazo de clave "${key}", cerrando...`);
        await handleCatalogLoop('after_key_rejected', key);
      }
      const recovered = await recoverNuevaCitaAfterReject('key_rejected', key);
      if (!recovered) {
        throw new Error('No se pudo recuperar el modal "Nueva cita" tras rechazo de clave.');
      }
      continue;
    }

    if (await closeCatalogIfAppearsSoon(page, 900, `after_key_set:${key}`)) {
      rememberKeyOutcome(key, 'catalog_opened:after_key_set_late');
      const recovered = await recoverNuevaCitaAfterReject('late_catalog_after_key_set', key);
      if (!recovered) {
        throw new Error('No se pudo recuperar "Nueva cita" tras catálogo tardío después de setear clave.');
      }
      continue;
    }

    if (await isPatientNotFoundAlertVisible(page)) {
      console.log(`ALERTA_404_PACIENTE_NO_ENCONTRADO tras set clave "${key}", cerrando y probando siguiente clave`);
      rememberKeyOutcome(key, 'patient_not_found_after_key');
      await closePatientNotFoundAlert(page);
      if (await isCatalogPacientesModalVisible(page)) {
        await handleCatalogLoop('patient_not_found_after_key', key);
      }
      const recovered = await recoverNuevaCitaAfterReject('patient_not_found_after_key', key);
      if (!recovered) {
        throw new Error('No se pudo recuperar "Nueva cita" tras alerta 404/paciente no encontrado.');
      }
      continue;
    }

    console.log(
      `CLAVE_SET_OK "${result.valueSet}" inputId="${result.keyInputId || ''}" current="${result.currentValue || ''}" search=${result.clickedSearch}`
    );
    await debugDumpNuevaCitaControls(page, `after_key:${key}`);

    await waitForTimeoutRaw(page, KEY_SETTLE_MS);
    let keyResolution = await waitForNuevaCitaKeyResolution(page, KEY_RESOLUTION_TIMEOUT_MS);
    if (keyResolution.state === 'catalog_opened') {
      console.log(`CATALOGO_PACIENTES_DETECTADO con clave "${key}", cerrando y probando siguiente clave`);
      await handleCatalogLoop('after_key_set', key);
      continue;
    }

    let patientLoaded = keyResolution.state === 'patient_loaded';
    let guardarEnabled = keyResolution.state === 'guardar_enabled';
    if (!patientLoaded && !guardarEnabled) {
      patientLoaded = await isPatientLoadedInNuevaCita(page);
      guardarEnabled = await isGuardarEnabledInNuevaCita(page);
    }
    if (!patientLoaded && !guardarEnabled) {
      if (!ALLOW_SAVE_ON_UNCONFIRMED_KEY) {
        console.log(
          `CLAVE_NO_CONFIRMADA "${key}" state=${keyResolution.state} elapsed=${keyResolution.elapsedMs}ms -> saltando Guardar y probando siguiente clave`
        );
        rememberKeyOutcome(key, `key_not_confirmed:${keyResolution.state}`);
        const recovered = await recoverNuevaCitaAfterReject('key_not_confirmed_after_resolution', key);
        if (!recovered) {
          throw new Error('No se pudo recuperar el modal "Nueva cita" tras clave no confirmada.');
        }
        continue;
      }
      console.log(
        `CLAVE_NO_CONFIRMADA "${key}" state=${keyResolution.state} elapsed=${keyResolution.elapsedMs}ms (override ALLOW_SAVE_ON_UNCONFIRMED_KEY=1)`
      );
    }

    if (!(await ensureStrictNuevaCitaContext(page, `before_comments:${key}`))) {
      console.log(`STRICT_MODAL_BLOQUEO "${key}" antes de Comentarios; reintentando con siguiente clave`);
      const recovered = await recoverNuevaCitaAfterReject('strict_before_comments', key);
      if (!recovered) {
        throw new Error('No se pudo recuperar el modal "Nueva cita" tras bloqueo estricto antes de Comentarios.');
      }
      continue;
    }

    let commentsClicked = await fillComentarioTextInNuevaCita(page, COMMENT_TEXT);
    for (let cTry = 0; cTry < COMMENT_CLICK_RETRIES && !commentsClicked; cTry += 1) {
      // Importante: evitar clicks heurísticos en controles ambiguos (ej. lupa de búsqueda).
      commentsClicked = await fillComentarioTextInNuevaCita(page, COMMENT_TEXT);
      if (commentsClicked) break;
      await waitForTimeoutRaw(page, 220);
    }
    if (!commentsClicked) {
      console.log(`COMENTARIOS_NO_DETECTADO "${key}" en textarea, continuando con Guardar sin clicks extra.`);
    }
    await debugDumpNuevaCitaControls(page, `after_comments_try:${key}`);
    await waitForTimeoutRaw(page, 320);

    keyResolution = await waitForNuevaCitaKeyResolution(page, Math.max(1200, Math.round(KEY_RESOLUTION_TIMEOUT_MS * 0.55)));
    if (await closeCatalogIfAppearsSoon(page, 800, `after_comments:${key}`)) {
      rememberKeyOutcome(key, 'catalog_opened:after_comments_late');
      const modalReady = await reopenModalForRetry('late_catalog_after_comments');
      if (!modalReady) {
        throw new Error('No se pudo reabrir "Nueva cita" tras catálogo tardío después de Comentarios.');
      }
      continue;
    }
    if (await isPatientNotFoundAlertVisible(page)) {
      console.log(`ALERTA_404_PACIENTE_NO_ENCONTRADO tras Comentarios con clave "${key}", cerrando y probando siguiente clave`);
      rememberKeyOutcome(key, 'patient_not_found_after_comments');
      await closePatientNotFoundAlert(page);
      if (await isCatalogPacientesModalVisible(page)) {
        await handleCatalogLoop('patient_not_found_after_comments', key);
      }
      const modalReady = await reopenModalForRetry('patient_not_found_after_comments');
      if (!modalReady) {
        throw new Error('No se pudo reabrir "Nueva cita" tras alerta 404/paciente no encontrado después de Comentarios.');
      }
      continue;
    }
    if (keyResolution.state === 'catalog_opened' || await isCatalogPacientesModalVisible(page)) {
      console.log(`CATALOGO_PACIENTES_DETECTADO tras click Comentarios con clave "${key}", cerrando y probando siguiente clave`);
      await handleCatalogLoop('after_comments', key);
      const modalReady = await reopenModalForRetry('catalog_after_comments');
      if (!modalReady) {
        throw new Error('No se pudo reabrir "Nueva cita" tras cerrar catálogo.');
      }
      continue;
    }

    const guardarEnabledAfterComments = await isGuardarEnabledInNuevaCita(page);
    const patientLoadedAfterComments = await isPatientLoadedInNuevaCita(page);
    if (!guardarEnabledAfterComments && !patientLoadedAfterComments) {
      if (!ALLOW_SAVE_ON_UNCONFIRMED_KEY) {
        console.log(`GUARDAR_NO_HABILITADO "${key}" tras Comentarios -> saltando Guardar (modo estricto).`);
        rememberKeyOutcome(key, 'guardar_not_enabled_after_comments');
        const recovered = await recoverNuevaCitaAfterReject('guardar_not_enabled_after_comments', key);
        if (!recovered) {
          throw new Error('No se pudo recuperar el modal "Nueva cita" tras Guardar no habilitado.');
        }
        continue;
      }
      console.log(`GUARDAR_NO_HABILITADO "${key}" tras Comentarios. Se intentará Guardar por override.`);
    }

    if (!AUTO_SAVE_APPOINTMENT) return makeAttemptResult(key, '');

    if (!(await ensureStrictNuevaCitaContext(page, `before_save:${key}`))) {
      console.log(`STRICT_MODAL_BLOQUEO "${key}" antes de Guardar; reintentando con siguiente clave`);
      const recovered = await recoverNuevaCitaAfterReject('strict_before_save', key);
      if (!recovered) {
        throw new Error('No se pudo recuperar el modal "Nueva cita" tras bloqueo estricto antes de Guardar.');
      }
      continue;
    }

    await disableSearchControlsInNuevaCita(page);

    let saved = false;
    saveAttemptInProgress = true;
    try {
      saved = await clickGuardarNuevaCita(page);
    } finally {
      saveAttemptInProgress = false;
    }
      if (saved) {
        rememberKeyOutcome(key, 'success', 'success');
        const successAfterSave = await getCitaGeneradaSuccessInfo(page);
        const successNumber = successAfterSave.ok ? String(successAfterSave.number || '').trim() : '';
        if (successAfterSave.ok) {
          console.log(`ALERTA_CITA_EXITO detectada con clave "${key}" numero="${successAfterSave.number || '?'}"`);
        }
        const memSaved = rememberCreatedAppointment({
          key,
          number: successNumber,
          slot: preferredSlot,
          status: successAfterSave.ok ? 'success_alert' : 'success_inferred'
        });
      if (memSaved.ok && !memSaved.dedup) {
        console.log(
          `APPOINTMENT_MEMORY_SAVE key="${key}" numero="${successAfterSave.number || '?'}" total=${memSaved.total} persisted=${memSaved.persisted ? 1 : 0}`
        );
      }
      return makeAttemptResult(key, successNumber);
    }

    if (await closeCatalogIfAppearsSoon(page, 1400, `after_save:${key}`)) {
      rememberKeyOutcome(key, 'catalog_opened:after_save_late');
      const modalReady = await reopenModalForRetry('late_catalog_after_save');
      if (!modalReady) {
        throw new Error('No se pudo reabrir "Nueva cita" tras catálogo tardío después de Guardar.');
      }
      continue;
    }

    if (await isPatientNotFoundAlertVisible(page)) {
      console.log(`ALERTA_404_PACIENTE_NO_ENCONTRADO tras Guardar con clave "${key}", cerrando y probando siguiente clave`);
      rememberKeyOutcome(key, 'patient_not_found_after_save');
      await closePatientNotFoundAlert(page);
      if (await isCatalogPacientesModalVisible(page)) {
        await handleCatalogLoop('patient_not_found_after_save', key);
      }
      const modalReady = await reopenModalForRetry('patient_not_found_after_save');
      if (!modalReady) {
        throw new Error('No se pudo reabrir "Nueva cita" tras alerta 404/paciente no encontrado después de Guardar.');
      }
      continue;
    }
    const successInfo = await getCitaGeneradaSuccessInfo(page);
    if (successInfo.ok) {
      rememberKeyOutcome(key, 'success', 'success');
      console.log(`ALERTA_CITA_EXITO detectada con clave "${key}" numero="${successInfo.number || '?'}"`);
      const memSaved = rememberCreatedAppointment({
        key,
        number: successInfo.number || '',
        slot: preferredSlot,
        status: 'success_alert'
      });
      if (memSaved.ok && !memSaved.dedup) {
        console.log(
          `APPOINTMENT_MEMORY_SAVE key="${key}" numero="${successInfo.number || '?'}" total=${memSaved.total} persisted=${memSaved.persisted ? 1 : 0}`
        );
      }
      return makeAttemptResult(key, String(successInfo.number || '').trim());
    }

    // Ventana de gracia: evita falso fallo cuando la alerta verde aparece tarde.
    const lateSuccessInfo = await waitForCitaGeneradaSuccessInfo(page, 5600);
    if (lateSuccessInfo.ok) {
      rememberKeyOutcome(key, 'success', 'success');
      console.log(`ALERTA_CITA_EXITO_LATE detectada con clave "${key}" numero="${lateSuccessInfo.number || '?'}"`);
      const memSaved = rememberCreatedAppointment({
        key,
        number: lateSuccessInfo.number || '',
        slot: preferredSlot,
        status: 'success_alert_late'
      });
      if (memSaved.ok && !memSaved.dedup) {
        console.log(
          `APPOINTMENT_MEMORY_SAVE key="${key}" numero="${lateSuccessInfo.number || '?'}" total=${memSaved.total} persisted=${memSaved.persisted ? 1 : 0}`
        );
      }
      return makeAttemptResult(key, String(lateSuccessInfo.number || '').trim());
    }

    if (await isAlreadyScheduledAlertVisible(page)) {
      console.log(`ALERTA_CITA_PROGRAMADA con clave "${key}", cerrando alerta y probando siguiente clave`);
      rememberKeyOutcome(key, 'already_scheduled_alert');
      await closeAlreadyScheduledAlert(page);
      const modalReady = await reopenModalForRetry('already_scheduled_alert');
      if (!modalReady) {
        throw new Error('No se pudo reabrir el modal "Nueva cita" después de alerta de cita programada.');
      }
      continue;
    }

    try {
      if (page.isClosed()) return makeAttemptResult(key, '');
    } catch {}
    let modalVisible = false;
    try {
      modalVisible = await isNuevaCitaModalVisible(page);
    } catch {
      modalVisible = false;
    }
    if (!modalVisible) {
      console.log(`MODAL_CERRADO_TRAS_FALLO_GUARDAR "${key}" -> reabriendo para intentar otra clave`);
    }
    let assigned = false;
    try {
      assigned = await isNuevaCitaAsignadaModalVisible(page);
    } catch {
      assigned = false;
    }
    if (assigned) {
      console.log(`MODAL_NUEVA_CITA_ASIGNADA detectado tras clave "${key}" -> cerrar y reintentar con otra clave`);
      await closeNuevaCitaAsignadaModal(page);
      const modalReady = await reopenModalForRetry('assigned_after_save');
      if (!modalReady) {
        throw new Error('No se pudo reabrir el modal "Nueva cita" tras detectar "Nueva cita asignada" después de Guardar.');
      }
      continue;
    }
    const reopened = await reopenModalForRetry(modalVisible ? 'save_failed_continue' : 'save_failed_modal_closed');
    rememberKeyOutcome(key, modalVisible ? 'save_failed_continue' : 'save_failed_modal_closed');
    if (!reopened) {
      throw new Error(`No se pudo reabrir el modal "Nueva cita" tras fallo de Guardar con clave "${key}".`);
    }
    console.log(`GUARDAR_FALLO_CON_CLAVE "${key}", modal recuperado, intentando siguiente clave`);
  }

  throw new Error(`No se pudo guardar la cita con ninguna clave dentro del límite (${MAX_KEY_ATTEMPTS}).`);
}

async function waitAfterGuardarSingleClick(page, timeoutMs = 3200) {
  const started = Date.now();
  while ((Date.now() - started) < timeoutMs) {
    const successInfo = await getCitaGeneradaSuccessInfo(page);
    if (successInfo.ok) return { ok: true, reason: 'success_alert' };

    if (await isPatientNotFoundAlertVisible(page)) {
      await closePatientNotFoundAlert(page);
      if (await isCatalogPacientesModalVisible(page)) {
        await closeCatalogPacientesModal(page);
      }
      return { ok: false, reason: 'patient_not_found_404' };
    }

    if (await isNuevaCitaAsignadaModalVisible(page)) {
      return { ok: false, reason: 'nueva_cita_asignada_opened' };
    }

    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      return { ok: false, reason: 'catalog_opened' };
    }

    if (!(await isNuevaCitaModalVisible(page))) {
      if (!REQUIRE_SAVE_ALERT) {
        return { ok: true, reason: 'modal_closed' };
      }
      // Modo estricto: no dar éxito por cierre de modal sin confirmar alerta verde.
      const confirmStarted = Date.now();
      while ((Date.now() - confirmStarted) < 2600) {
        const confirmed = await getCitaGeneradaSuccessInfo(page);
        if (confirmed.ok) return { ok: true, reason: 'success_alert_after_close' };
        if (await isCatalogPacientesModalVisible(page)) {
          await closeCatalogPacientesModal(page);
          return { ok: false, reason: 'catalog_opened_after_close' };
        }
        await waitForTimeoutRaw(page, 130);
      }
      return { ok: false, reason: 'closed_without_success_alert' };
    }
    await waitForTimeoutRaw(page, 160);
  }
  if (await closeCatalogIfAppearsSoon(page, 520, 'wait_after_guardar_timeout')) {
    return { ok: false, reason: 'catalog_opened_timeout' };
  }
  if (await isCatalogPacientesModalVisible(page)) {
    await closeCatalogPacientesModal(page);
    return { ok: false, reason: 'catalog_opened_timeout_direct' };
  }
  return { ok: false, reason: 'modal_still_open' };
}

async function clickGuardarNuevaCita(page) {
  if (!(await ensureStrictNuevaCitaContext(page, 'guardar_click_start'))) return false;
  if (!(await isNuevaCitaModalVisible(page))) return false;

  let clicked = false;
  let clickPath = '';

  // Click único directo por rol (prioridad 1).
  try {
    const btn = page.getByRole('button', { name: /guardar/i }).first();
    if ((await btn.count()) > 0 && (await btn.isVisible()) && !(await btn.isDisabled())) {
      await btn.click({ force: true, timeout: 1500 });
      clicked = true;
      clickPath = 'role-page';
    }
  } catch {}

  // Click único en frame solo si aún no se hizo click.
  if (!clicked) {
    for (const frame of page.frames()) {
      try {
        const btn = frame.getByRole('button', { name: /guardar/i }).first();
        if ((await btn.count()) > 0 && (await btn.isVisible()) && !(await btn.isDisabled())) {
          await btn.click({ force: true, timeout: 1500 });
          clicked = true;
          clickPath = 'role-frame';
          break;
        }
      } catch {}
    }
  }

  // Fallback DOM con click único (sin doble click de parentLink).
  if (!clicked) {
    try {
      clicked = await page.evaluate(() => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .trim();
        const visible = (el) => {
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 12;
        };

        const dialogs = Array.from(document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')).filter(visible);
        const root = dialogs.find((d) => {
          const t = normalize(d.textContent || '');
          const title = normalize(
            (d.querySelector('.k-window-title, .k-dialog-title, .modal-title')?.textContent || d.getAttribute('aria-label') || '')
          );
          const isNuevaCitaTitle = title === 'nueva cita' || title.startsWith('nueva cita');
          return isNuevaCitaTitle && t.includes('clave documento');
        });
        if (!root) return false;

        const nodes = Array.from(
          root.querySelectorAll('button, a, span, input[type="button"], input[type="submit"], [id*="guardar"], [name*="guardar"]')
        ).filter(visible);
        let btn = nodes.find((n) => {
          const t = normalize(n.textContent || n.value || '');
          if (t.includes('guardar')) return true;
          const idName = normalize(`${n.id || ''} ${n.getAttribute?.('name') || ''} ${n.getAttribute?.('title') || ''}`);
          return idName.includes('guardar');
        });
        if (!(btn instanceof HTMLElement)) {
          const footerButtons = Array.from(
            root.querySelectorAll('.k-window-actions button, .k-window-actions a, .k-window-actions input, .modal-footer button, .modal-footer a, .modal-footer input')
          ).filter(visible);
          btn =
            footerButtons.find((n) => {
              const t = normalize(n.textContent || n.value || '');
              return !t.includes('cerrar') && !t.includes('cancelar') && !t.includes('salir');
            }) || footerButtons[footerButtons.length - 1];
        }
        if (!(btn instanceof HTMLElement)) return false;
        if (btn.disabled || btn.getAttribute('aria-disabled') === 'true') return false;

        btn.click();
        return true;
      });
      if (clicked) clickPath = 'dom-fallback';
    } catch {}
  }

  if (!clicked) return false;

  const outcome = await waitAfterGuardarSingleClick(page, 3400);
  if (outcome.reason === 'nueva_cita_asignada_opened') {
    await closeNuevaCitaAsignadaModal(page);
  }
  if (!outcome.ok && outcome.reason === 'modal_still_open') {
    await closeCatalogIfAppearsSoon(page, 700, 'after_guardar_modal_still_open');
    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
    }
  }
  if (!outcome.ok) {
    console.log(`GUARDAR_SINGLE_CLICK_NO_RESULT path=${clickPath || 'unknown'} reason=${outcome.reason}`);
    return false;
  }
  console.log(`GUARDAR_SINGLE_CLICK_OK path=${clickPath || 'unknown'} reason=${outcome.reason}`);
  return true;
}

async function createAppointmentFromCalendar(page) {
  const MAX_WEEKS = 2; // semana actual + siguiente, evita saltos bruscos
  let opened = await isNuevaCitaModalVisible(page);
  let assigned = await isNuevaCitaAsignadaModalVisible(page);
  let preferredSlot = null;
  let lastFingerprint = '';
  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  const maxAllowedStart = todayStart + (7 * 24 * 60 * 60 * 1000);

  await ensureWorkingHoursVisible(page);
  await ensureCalendarOnCurrentWeek(page, { applyFilter: false });

  if (opened) {
    console.log('NUEVA_CITA_MODAL_YA_ABIERTO: pasando directo a carga de clave');
  }

  for (let week = 0; week < MAX_WEEKS && !opened; week += 1) {
    // Semana 0: revisar la semana actual primero. Desde semana 1: avanzar.
    if (week > 0) {
      await goToNextCalendarRange(page);
      await page.waitForTimeout(1200);
      await applyAgendaFilter(page);
    }
    await ensureWorkingHoursVisible(page);
    // Esperar a que los eventos del scheduler se rendericen
    await page.waitForTimeout(1500);

    const weekInfo = await getVisibleWeekInfo(page);
    console.log(`WEEK_RANGE week=${week} start=${weekInfo.startIso || '-'} end=${weekInfo.endIso || '-'} label="${weekInfo.label}"`);
    if (weekInfo.startTs && weekInfo.startTs > maxAllowedStart) {
      console.log(`WEEK_JUMP_GUARD week=${week} start=${weekInfo.startIso} > nextWeek, deteniendo búsqueda`);
      break;
    }

    // Si el usuario abrió modal manualmente o se abrió por transición, continuar.
    opened = await isNuevaCitaModalVisible(page);
    assigned = await isNuevaCitaAsignadaModalVisible(page);
    if (opened) {
      console.log(`NUEVA_CITA_MODAL_DETECTADO week=${week} antes de seleccionar casilla`);
      break;
    }

    // Obtener estado de la semana (bloqueada + fingerprint)
    const status = await getWeekStatus(page);
    console.log(`WEEK_CHECK week=${week} blocked=${status.blocked} noDisp=${status.noDispEvents} days=${status.dayColumns} kEvents=${status.totalKEvents}`);

    // Detectar calendario estancado (misma vista que la semana anterior)
    if (week > 0 && status.fingerprint === lastFingerprint) {
      console.log(`CALENDAR_STALE week=${week} - calendario no avanzó, deteniendo búsqueda`);
      break;
    }
    lastFingerprint = status.fingerprint;

    // Detección rápida de semana completamente bloqueada
    if (status.blocked) {
      console.log(`SEMANA_BLOQUEADA week=${week} - ${status.noDispEvents} eventos "NO DISPONIBLE" en ${status.dayColumns} columnas, avanzando...`);
      continue;
    }

    // Intentar abrir modal en casilla libre
    const fast = await fastOpenNuevaCitaFromCandidates(page);
    if (fast.ok) {
      if (fast.slot) preferredSlot = fast.slot;
      opened = await isNuevaCitaModalVisible(page);
      assigned = await isNuevaCitaAsignadaModalVisible(page);
    }

    // Si se abrió "Nueva cita asignada", cerrar y reintentar en misma semana
    if (!opened && assigned) {
      console.log('INFO modal "Nueva cita asignada" detectado; cerrando y reintentando casilla libre.');
      await closeNuevaCitaAsignadaModal(page);
      await page.waitForTimeout(300);
      const retry = await fastOpenNuevaCitaFromCandidates(page);
      if (retry.ok) {
        if (retry.slot) preferredSlot = retry.slot;
        opened = await isNuevaCitaModalVisible(page);
        assigned = await isNuevaCitaAsignadaModalVisible(page);
      }
    }

    if (opened) break;

    console.log(`SIN_DISPONIBILIDAD week=${week} - ninguna casilla abrió modal, avanzando...`);
  }

  if (!opened && !assigned && COOP_MODE) {
    console.log('COOP_MODE: esperando 4s por apertura manual del modal "Nueva cita"...');
    const coop = await waitForNuevaCitaModal(page, 4000);
    opened = coop.opened;
    assigned = coop.assigned;
  }

  if (!opened) {
    throw new Error(`No se pudo abrir modal "Nueva cita" tras recorrer ${MAX_WEEKS} semanas.`);
  }

  const saveResult = await loadFirstPatientFromKeys(page, PATIENT_KEYS, { preferredSlot });
  const usedKey = typeof saveResult === 'string' ? saveResult : String(saveResult?.key || '');
  const createdAppointmentNumber = typeof saveResult === 'string' ? '' : String(saveResult?.appointmentNumber || '').trim();
  if (AUTO_SAVE_APPOINTMENT) {
    console.log(`CITA_GUARDADA_CON_CLAVE "${usedKey}"`);
    if (AUTO_OPEN_MODULE_AFTER_SAVE) {
      console.log(`Paso 8: abrir módulo desde la cita recién guardada numero="${createdAppointmentNumber || '-'}"`);
      await openModuloAfterAppointmentSave(page, preferredSlot, {
        appointmentNumber: createdAppointmentNumber,
        patientKey: usedKey
      });
    }
  } else {
    console.log(`CLAVE_CARGADA_EN_MODAL "${usedKey}"`);
  }
  return usedKey;
}

async function openFirstAppointment(page) {
  const clicked = await page.evaluate(() => {
    const selectors = [
      '.k-event',
      '.rsApt',
      '[class*="appointment"]',
      '[class*="Appointment"]',
      '[id*="Appointment"]',
      '[id*="appointment"]'
    ];

    for (const s of selectors) {
      const node = document.querySelector(s);
      if (node instanceof HTMLElement) {
        node.click();
        return true;
      }
    }
    return false;
  });

  if (clicked) return;

  const patientLike = page.getByText(/-\s*\d{4,6}/).first();
  await patientLike.waitFor({ state: 'visible', timeout: 10000 });
  await patientLike.click();
}

async function isCalendarGridVisible(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    const cell = page.locator('.rsContentTable td, .k-scheduler-table td, .k-scheduler-content td, td[role="gridcell"]').first();
    if ((await cell.count()) === 0) return false;
    return await cell.isVisible();
  } catch {
    return false;
  }
}

async function readQuickActionStatus(page) {
  if (isPageClosedSafe(page)) return { visible: false, status: 'none', isFinalizada: false, isProgramada: false, isVideollamada: false, text: '' };
  try {
    return await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 16;
      };

      const popups = Array.from(
        document.querySelectorAll('.k-widget.k-tooltip, .k-tooltip, .k-popup, [role="tooltip"], .k-animation-container, .div_hos930AvisoPaciente')
      ).filter(visible);

      const candidates = [];
      for (const popup of popups) {
        const txt = normalize(popup.textContent || '');
        if (!txt) continue;
        if (txt.includes('catalogo de pacientes') || txt.includes('catálogo de pacientes')) continue;
        const controls = Array.from(
          popup.querySelectorAll('button, a, span, div, input[type="button"], input[type="submit"], [role="button"]')
        ).filter(visible);
        const hasModulo = controls.some((n) => {
          const t = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || n.value || '');
          return t === 'modulo' || t.includes('abrir modulo') || t.includes('abrir módulo');
        });
        if (!hasModulo) continue;
        const r = popup.getBoundingClientRect();
        let score = 0;
        if (txt.includes('programada')) score += 30;
        if (txt.includes('videollamada')) score += 30;
        if (txt.includes('finalizada')) score += 45;
        if (txt.includes('registro')) score += 15;
        if (txt.includes('expediente')) score += 15;
        candidates.push({ txt, area: r.width * r.height, score });
      }

      if (!candidates.length) {
        return { visible: false, status: 'none', isFinalizada: false, isProgramada: false, isVideollamada: false, text: '' };
      }
      candidates.sort((a, b) => {
        if (a.score !== b.score) return b.score - a.score;
        return a.area - b.area;
      });

      const txt = candidates[0].txt;
      const isFinalizada = txt.includes('finalizada') || txt.includes('inasistencia');
      const isProgramada = txt.includes('programada');
      const isVideollamada = txt.includes('videollamada');
      let status = 'unknown';
      if (isFinalizada) status = 'finalizada';
      else if (isProgramada && isVideollamada) status = 'programada_videollamada';
      else if (isProgramada) status = 'programada';
      else if (isVideollamada) status = 'videollamada';

      return { visible: true, status, isFinalizada, isProgramada, isVideollamada, text: txt.slice(0, 220) };
    });
  } catch {
    return { visible: false, status: 'none', isFinalizada: false, isProgramada: false, isVideollamada: false, text: '' };
  }
}

async function getExistingAppointmentSlots(page, limit = 30, options = {}) {
  if (isPageClosedSafe(page)) return [];
  const preferFinalizada = options?.preferFinalizada === true;
  const excludeFinalizada = options?.excludeFinalizada === true;
  const onlyProgramadaVideollamada = options?.onlyProgramadaVideollamada === true;
  const minDayIso = String(options?.minDayIso || '').trim();
  const excludeSundays = options?.excludeSundays === true;
  try {
    const result = await page.evaluate(({ limit, preferFinalizada, excludeFinalizada, onlyProgramadaVideollamada, minDayIso, excludeSundays }) => {
      const EVENT_SELECTOR =
        '.k-event, .rsApt, [class*="k-event"], [class*="appointment"], [class*="Appointment"], [id*="appointment"], [id*="Appointment"]';
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 14 && r.height > 10;
      };
      const parseRgb = (raw) => {
        const m = String(raw || '').match(/rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/i);
        if (!m) return null;
        return { r: Number(m[1]), g: Number(m[2]), b: Number(m[3]) };
      };
      const colorDistance = (a, b) => {
        if (!a || !b) return Number.POSITIVE_INFINITY;
        const dr = a.r - b.r;
        const dg = a.g - b.g;
        const db = a.b - b.b;
        return Math.sqrt((dr * dr) + (dg * dg) + (db * db));
      };
      const isNeutralColor = (rgb) => {
        if (!rgb) return true;
        const max = Math.max(rgb.r, rgb.g, rgb.b);
        const min = Math.min(rgb.r, rgb.g, rgb.b);
        // blancos/grises/transiciones casi neutras no sirven para clasificar estado
        return max >= 235 || (max - min) < 10;
      };
      const isPurpleLike = (raw) => {
        const rgb = parseRgb(raw);
        if (!rgb) return false;
        // Morado aproximado: componente azul alta y verde más baja.
        return rgb.b >= 120 && rgb.r >= 90 && rgb.g <= 150;
      };
      const pickLegendColor = (row) => {
        if (!(row instanceof HTMLElement)) return null;
        const sample = [row, ...Array.from(row.querySelectorAll('*'))];
        const candidates = [];
        for (const el of sample) {
          if (!(el instanceof HTMLElement)) continue;
          if (!visible(el)) continue;
          const r = el.getBoundingClientRect();
          if (r.width < 2 || r.height < 2 || r.width > 120 || r.height > 120) continue;
          const st = getComputedStyle(el);
          const bg = parseRgb(st.backgroundColor);
          const bd = parseRgb(st.borderColor);
          const bl = parseRgb(st.borderLeftColor);
          if (bg && !isNeutralColor(bg)) candidates.push({ rgb: bg, area: r.width * r.height, priority: 1 });
          if (bl && !isNeutralColor(bl)) candidates.push({ rgb: bl, area: r.width * r.height, priority: 0 });
          if (bd && !isNeutralColor(bd)) candidates.push({ rgb: bd, area: r.width * r.height, priority: 2 });
        }
        if (!candidates.length) return null;
        candidates.sort((a, b) => a.priority - b.priority || a.area - b.area);
        return candidates[0].rgb;
      };
      const findLegendRow = (label) => {
        const target = normalize(label);
        const nodes = Array.from(document.querySelectorAll('li, div, span, p, a'));
        const filtered = nodes.filter((n) => {
          if (!(n instanceof HTMLElement)) return false;
          if (!visible(n)) return false;
          const t = normalize(n.textContent || '');
          if (!t.includes(target)) return false;
          const r = n.getBoundingClientRect();
          return r.width >= 70 && r.width <= 260 && r.height >= 18 && r.height <= 120;
        });
        if (!filtered.length) return null;
        filtered.sort((a, b) => {
          const ar = a.getBoundingClientRect();
          const br = b.getBoundingClientRect();
          return (ar.width * ar.height) - (br.width * br.height);
        });
        return filtered[0];
      };
      const legendProgramadaVideoRgb = pickLegendColor(findLegendRow('programada videollamada'));
      const legendFinalizadaRgb = pickLegendColor(findLegendRow('finalizada'));
      const legendInasistenciaRgb = pickLegendColor(findLegendRow('inasistencia'));
      const legendProgramadaLlamadaRgb = pickLegendColor(findLegendRow('programada llamada directa'));

      const bannedWords = [
        'no disponible',
        'bloqueo',
        'bloqueada',
        'inasistencia',
        'no es el paciente'
      ];
      const isTimeRangeOnly = (txt) =>
        /^\d{1,2}\s*:\s*\d{2}\s*(am|pm)\s*-\s*\d{1,2}\s*:\s*\d{2}\s*(am|pm)$/.test(String(txt || '').trim());
      const stripTimeRanges = (txt) =>
        String(txt || '')
          .replace(/\b\d{1,2}\s*:\s*\d{2}\s*(am|pm)\b/g, '')
          .replace(/\s*-\s*/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
      const hasPatientNameLike = (txt) => {
        const cleaned = stripTimeRanges(txt)
          .replace(/\b(programada|videollamada|finalizada|registro|expediente|fecha|hora|inicio|fin)\b/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
        return /[a-záéíóúñ]{3,}/.test(cleaned);
      };
      const today = new Date();
      const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      const parseHeaderDate = (text) => {
        const m = String(text || '').match(/(\d{1,2})\s*\/\s*(\d{1,2})/);
        if (!m) return null;
        const mm = Number(m[1]);
        const dd = Number(m[2]);
        if (!mm || !dd) return null;
        let year = todayStart.getFullYear();
        let d = new Date(year, mm - 1, dd);
        if (d < new Date(todayStart.getTime() - 1000 * 60 * 60 * 24 * 300)) d = new Date(year + 1, mm - 1, dd);
        return d;
      };

      const headers = Array.from(document.querySelectorAll('thead th, .k-scheduler-header th, .rsHeader')).filter((h) => visible(h));
      const dayColumns = [];
      for (const h of headers) {
        const raw = (h.textContent || '').trim();
        const date = parseHeaderDate(raw);
        if (!date) continue;
        const r = h.getBoundingClientRect();
        if (!(r.width > 10 && r.height > 8)) continue;
        dayColumns.push({
          left: r.left,
          right: r.right,
          center: r.left + (r.width / 2),
          label: raw.slice(0, 40),
          dayIso: date.toISOString().slice(0, 10),
          dayTs: new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime(),
          dayDow: date.getDay()
        });
      }
      dayColumns.sort((a, b) => a.center - b.center);

      const nodes = Array.from(document.querySelectorAll(EVENT_SELECTOR))
        .map((n, domIdx) => ({ n, domIdx }))
        .filter((item) => visible(item.n));

      const dedup = new Set();
      const logicalDedup = new Set();
      const out = [];
      for (const item of nodes) {
        const n = item.n;
        const domIdx = item.domIdx;
        const txt = normalize(n.textContent || n.getAttribute('title') || n.getAttribute('aria-label') || '');
        if (!txt) continue;
        if (bannedWords.some((w) => txt.includes(w))) continue;
        if (isTimeRangeOnly(txt)) continue;
        const meta = normalize(`${n.className || ''} ${n.getAttribute('title') || ''} ${n.getAttribute('aria-label') || ''}`);
        const st = getComputedStyle(n);
        const hasProgramada = txt.includes('programada') || meta.includes('programada');
        const hasVideollamada = txt.includes('videollamada') || meta.includes('videollamada');
        const hasFinalizadaLabel = txt.includes('finalizada') || meta.includes('finalizada');
        const hasInasistenciaLabel = txt.includes('inasistencia') || meta.includes('inasistencia');

        // Recopilar TODOS los colores relevantes del elemento y sus hijos inmediatos
        const collectEventColors = (el) => {
          const colors = [];
          const addColor = (raw) => { const c = parseRgb(raw); if (c && !isNeutralColor(c)) colors.push(c); };
          const elSt = getComputedStyle(el);
          addColor(elSt.backgroundColor);
          addColor(elSt.borderLeftColor);
          addColor(elSt.borderColor);
          addColor(elSt.outlineColor);
          // Revisar hijos inmediatos (tiras de color, indicadores)
          const children = Array.from(el.children || []).slice(0, 8);
          for (const ch of children) {
            if (!(ch instanceof HTMLElement)) continue;
            const chSt = getComputedStyle(ch);
            addColor(chSt.backgroundColor);
            addColor(chSt.borderLeftColor);
            addColor(chSt.borderColor);
            // Si el hijo es un div muy delgado (tira de color), darle prioridad
            const chR = ch.getBoundingClientRect();
            if (chR.width <= 8 && chR.height > 10) {
              addColor(chSt.backgroundColor);
            }
          }
          // También revisar inline style del elemento (Telerik a veces usa inline)
          const inlineBg = el.style?.backgroundColor;
          const inlineBl = el.style?.borderLeftColor;
          const inlineBd = el.style?.borderColor;
          if (inlineBg) addColor(inlineBg);
          if (inlineBl) addColor(inlineBl);
          if (inlineBd) addColor(inlineBd);
          return colors;
        };
        const eventColors = collectEventColors(n);

        const matchesLegendColor = (legendRgb, threshold = 72) => {
          if (!legendRgb) return false;
          return eventColors.some(c => colorDistance(c, legendRgb) <= threshold);
        };
        const hasProgramadaVideollamadaByLegendColor = matchesLegendColor(legendProgramadaVideoRgb);
        const hasFinalizadaByLegendColor = matchesLegendColor(legendFinalizadaRgb);
        const hasInasistenciaByLegendColor = matchesLegendColor(legendInasistenciaRgb);
        const hasProgramadaLlamadaByLegendColor = matchesLegendColor(legendProgramadaLlamadaRgb);
        const hasFinalizadaColor =
          isPurpleLike(st.backgroundColor) || isPurpleLike(st.borderLeftColor || st.borderColor) || isPurpleLike(st.outlineColor);
        const isFinalizada = hasFinalizadaLabel || hasFinalizadaColor || hasFinalizadaByLegendColor;
        const isInasistencia = hasInasistenciaLabel || hasInasistenciaByLegendColor;
        const isProgramadaVideollamada = !isFinalizada && !isInasistencia && !hasProgramadaLlamadaByLegendColor &&
          ((hasProgramada && hasVideollamada) || hasProgramadaVideollamadaByLegendColor);

        if (excludeFinalizada && isFinalizada) continue;
        if (isInasistencia) continue; // Siempre excluir inasistencia
        if (onlyProgramadaVideollamada && !isProgramadaVideollamada) continue;

        let statusHint = '';
        if (isFinalizada) statusHint = 'finalizada';
        else if (isInasistencia) statusHint = 'inasistencia';
        else if (isProgramadaVideollamada) statusHint = 'programada_videollamada';
        else if (hasProgramadaLlamadaByLegendColor) statusHint = 'programada_llamada_directa';
        else if (hasProgramada) statusHint = 'programada';
        else if (hasVideollamada) statusHint = 'videollamada';

        const r = n.getBoundingClientRect();
        const x = Math.round(r.left + (r.width / 2));
        const y = Math.round(r.top + (r.height / 2));

        let dayMeta = null;
        if (dayColumns.length) {
          dayMeta = dayColumns.find((d) => x >= d.left && x <= d.right) || null;
          if (!dayMeta) {
            dayMeta = dayColumns
              .map((d) => ({ d, dist: Math.abs(d.center - x) }))
              .sort((a, b) => a.dist - b.dist)[0]?.d || null;
          }
        }
        const dayIso = dayMeta?.dayIso || '';
        const dayTs = Number.isFinite(dayMeta?.dayTs) ? dayMeta.dayTs : 0;
        const dayLabel = dayMeta?.label || '';
        const dayDow = Number.isFinite(dayMeta?.dayDow) ? dayMeta.dayDow : -1;

        if (minDayIso) {
          // Si activamos fecha mínima, descartar slots sin mapeo de día para no saltar a columnas incorrectas.
          if (!dayIso) continue;
          if (dayIso < minDayIso) continue;
        }
        if (excludeSundays && dayDow === 0) continue;

        const key = `${Math.round(r.left / 2)}|${Math.round(r.top / 2)}|${Math.round(r.width / 2)}|${Math.round(r.height / 2)}`;
        if (dedup.has(key)) continue;
        dedup.add(key);

        const m = txt.match(/\b(\d{4,7})\b/);
        const appointmentNumber = m ? String(m[1]) : '';
        const normalizedCore = stripTimeRanges(txt).slice(0, 80);
        const logicalKey = appointmentNumber
          ? `${dayIso || '-'}|n:${appointmentNumber}`
          : `${dayIso || '-'}|c:${normalizedCore}|row:${Math.round(r.top / 14)}`;
        if (logicalDedup.has(logicalKey)) continue;
        logicalDedup.add(logicalKey);

        const confidenceScore = (() => {
          let s = 0;
          if (appointmentNumber) s += 120;
          if (hasPatientNameLike(txt)) s += 80;
          if (isProgramadaVideollamada) s += 90;
          if (isFinalizada) s -= 120;
          if (txt.includes(' am ') || txt.includes(' pm ')) s -= 10;
          return s;
        })();
        out.push({
          x,
          y,
          top: r.top,
          left: r.left,
          width: r.width,
          height: r.height,
          text: (txt || '').slice(0, 120),
          appointmentNumber,
          statusHint,
          isFinalizada,
          isProgramadaVideollamada,
          dayIso,
          dayTs,
          dayDow,
          dayLabel,
          confidenceScore,
          selector: EVENT_SELECTOR,
          domIdx
        });
      }

      const statusRank = (status) => {
        const s = normalize(status || '');
        if (s === 'programada_videollamada') return 4;
        if (s === 'finalizada') return 3;
        if (s === 'programada') return 2;
        if (s === 'videollamada') return 1;
        return 0;
      };
      out.sort((a, b) => {
        if (preferFinalizada) {
          const ar = statusRank(a.statusHint);
          const br = statusRank(b.statusHint);
          if (ar !== br) return br - ar;
        }
        if (a.dayTs > 0 && b.dayTs > 0 && a.dayTs !== b.dayTs) return a.dayTs - b.dayTs;
        if ((a.confidenceScore || 0) !== (b.confidenceScore || 0)) return (b.confidenceScore || 0) - (a.confidenceScore || 0);
        if (Math.abs(a.top - b.top) > 4) return a.top - b.top;
        return a.left - b.left;
      });
      const _debugLegend = {
        progVideo: legendProgramadaVideoRgb ? `rgb(${legendProgramadaVideoRgb.r},${legendProgramadaVideoRgb.g},${legendProgramadaVideoRgb.b})` : 'null',
        finalizada: legendFinalizadaRgb ? `rgb(${legendFinalizadaRgb.r},${legendFinalizadaRgb.g},${legendFinalizadaRgb.b})` : 'null',
        inasistencia: legendInasistenciaRgb ? `rgb(${legendInasistenciaRgb.r},${legendInasistenciaRgb.g},${legendInasistenciaRgb.b})` : 'null',
        progLlamada: legendProgramadaLlamadaRgb ? `rgb(${legendProgramadaLlamadaRgb.r},${legendProgramadaLlamadaRgb.g},${legendProgramadaLlamadaRgb.b})` : 'null'
      };
      return { slots: out.slice(0, Math.max(1, Number(limit) || 30)), _debugLegend };
    }, { limit, preferFinalizada, excludeFinalizada, onlyProgramadaVideollamada, minDayIso, excludeSundays });
    if (result && result._debugLegend) {
      console.log(`LEGEND_COLORS ${JSON.stringify(result._debugLegend)}`);
    }
    return (result && result.slots) || [];
  } catch {
    return [];
  }
}

async function clickCancelActionInModule(page) {
  if (isPageClosedSafe(page)) return false;

  const directSelectors = [
    'button:has-text("Finalizar cita"), a:has-text("Finalizar cita"), [role="button"]:has-text("Finalizar cita")',
    'button:has-text("Cancelar cita"), a:has-text("Cancelar cita"), [role="button"]:has-text("Cancelar cita")',
    'button:has-text("Anular cita"), a:has-text("Anular cita"), [role="button"]:has-text("Anular cita")',
    '[title*="finalizar" i], [aria-label*="finalizar" i], [id*="finalizar" i], [name*="finalizar" i]',
    '[title*="cancelar" i], [aria-label*="cancelar" i], [id*="cancelar" i], [name*="cancelar" i]',
    '[title*="anular" i], [aria-label*="anular" i], [id*="anular" i], [name*="anular" i]'
  ];

  for (const sel of directSelectors) {
    try {
      const loc = page.locator(sel).first();
      if ((await loc.count()) === 0) continue;
      if (!(await loc.isVisible())) continue;
      await loc.click({ force: true, timeout: 1200 });
      await waitForTimeoutRaw(page, 220);
      console.log(`CANCEL_CLICK_OK via=selector "${sel}"`);
      return true;
    } catch {}
  }

  try {
    const clicked = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        try {
          el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          el.click();
          return true;
        } catch {
          return false;
        }
      };

      const includeWord = (txt) =>
        txt.includes('finalizar cita') ||
        txt.includes('cancelar cita') ||
        txt.includes('anular cita') ||
        txt.includes('finalizar') ||
        txt.includes('cancelar') ||
        txt.includes('anular');

      const excludeWord = (txt) =>
        txt.includes('salir') ||
        txt.includes('guardar') ||
        txt.includes('cerrar') ||
        txt.includes('receta') ||
        txt.includes('laboratorio') ||
        txt.includes('imagenologia');

      const nodes = Array.from(
        document.querySelectorAll('button, a, span, div, i, input[type="button"], input[type="submit"], [role="button"], [title], [aria-label]')
      ).filter(visible);
      const scored = [];
      for (const n of nodes) {
        const txt = normalize(
          `${n.textContent || ''} ${n.getAttribute('title') || ''} ${n.getAttribute('aria-label') || ''} ${n.id || ''} ${n.getAttribute('name') || ''}`
        );
        if (!txt) continue;
        if (!includeWord(txt)) continue;
        if (excludeWord(txt)) continue;
        const r = n.getBoundingClientRect();
        let score = 100;
        if (txt.includes('finalizar cita')) score += 220;
        if (txt.includes('cancelar cita')) score += 220;
        if (txt.includes('anular cita')) score += 210;
        if (txt.includes('finalizar')) score += 140;
        if (txt.includes('cancelar')) score += 140;
        if (txt.includes('anular')) score += 120;
        if (r.top >= 80 && r.top <= 420) score += 65; // toolbar zona alta
        if (r.left >= 450) score += 25;
        scored.push({ n, score });
      }
      if (!scored.length) return false;
      scored.sort((a, b) => b.score - a.score);
      return safeClick(scored[0].n);
    });
    if (clicked) {
      await waitForTimeoutRaw(page, 220);
      console.log('CANCEL_CLICK_OK via=dom_scored');
      return true;
    }
  } catch {}

  console.log('CANCEL_CLICK_FAIL no se encontro boton de cancelar/finalizar');
  return false;
}

async function clickCancelActionInModuleWithRetry(page, options = {}) {
  if (isPageClosedSafe(page)) return false;
  const timeoutMs = (() => {
    const n = Number(options?.timeoutMs ?? CANCEL_ACTION_WAIT_TIMEOUT_MS);
    if (!Number.isFinite(n)) return CANCEL_ACTION_WAIT_TIMEOUT_MS;
    return Math.min(45000, Math.max(1000, Math.round(n)));
  })();
  const intervalMs = (() => {
    const n = Number(options?.intervalMs ?? CANCEL_ACTION_WAIT_INTERVAL_MS);
    if (!Number.isFinite(n)) return CANCEL_ACTION_WAIT_INTERVAL_MS;
    return Math.min(2000, Math.max(120, Math.round(n)));
  })();

  const started = Date.now();
  let tries = 0;
  while ((Date.now() - started) < timeoutMs) {
    if (isPageClosedSafe(page)) return false;
    tries += 1;

    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 120);
    }
    if (await isNuevaCitaModalVisible(page)) {
      await closeNuevaCitaModalIfOpen(page);
      await waitForTimeoutRaw(page, 110);
    }

    const clicked = await clickCancelActionInModule(page);
    if (clicked) {
      console.log(`CANCEL_CLICK_RETRY_OK tries=${tries} elapsed=${Date.now() - started}ms`);
      return true;
    }

    if (tries % 3 === 0) {
      console.log(`CANCEL_CLICK_RETRY_WAIT tries=${tries} elapsed=${Date.now() - started}ms`);
    }
    await sleepRaw(intervalMs);
  }

  console.log(`CANCEL_CLICK_RETRY_TIMEOUT elapsed=${Date.now() - started}ms`);
  return false;
}

async function confirmCancellationDialog(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    let clicked = false;

    // ── Intento 1: Alertify directo (selector exacto del diálogo "¿Desea finalizar?") ──
    try {
      const ajsOk = page.locator('div.alertify button.ajs-button.ajs-ok');
      const count = await ajsOk.count();
      if (count > 0) {
        await ajsOk.first().click({ timeout: 3000 });
        clicked = true;
        console.log('CANCEL_CONFIRM_OK via=alertify_ajs_ok');
      }
    } catch (e) {
      console.log(`CANCEL_CONFIRM_ALERTIFY_FAIL err=${(e.message || '').slice(0, 80)}`);
    }

    // ── Intento 2: Buscar botón "Sí" en cualquier diálogo visible ──
    if (!clicked) {
      const btnInfo = await page.evaluate(() => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const visible = (el) => {
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 10 && r.height > 10;
        };

        const dialogSel = '.alertify, .k-window, .k-dialog, .modal, .rwDialog, .rwWindow, .RadWindow, [role="dialog"], .swal2-popup, .ui-dialog';
        const dialogs = Array.from(document.querySelectorAll(dialogSel)).filter(visible);
        if (!dialogs.length) return null;

        const top = dialogs.sort((a, b) => {
          const az = Number.parseInt(getComputedStyle(a).zIndex || '0', 10) || 0;
          const bz = Number.parseInt(getComputedStyle(b).zIndex || '0', 10) || 0;
          return bz - az;
        })[0];

        const controls = Array.from(
          top.querySelectorAll('button, a, input[type="button"], input[type="submit"], [role="button"], span')
        ).filter(visible);
        let best = null;
        let bestScore = 0;
        for (const c of controls) {
          const txt = normalize(c.textContent || c.value || c.getAttribute('title') || c.getAttribute('aria-label') || '');
          if (!txt) continue;
          let score = 0;
          if (txt === 'si' || txt === 'sí') score += 350;
          if (c.classList.contains('ajs-ok')) score += 400;
          if (txt.includes('aceptar')) score += 300;
          if (txt.includes('confirmar')) score += 280;
          if (txt === 'ok') score += 250;
          if (txt.includes('continuar')) score += 220;
          if (txt.includes('no') || txt.includes('cancelar') || txt.includes('cerrar') || txt.includes('salir')) score -= 500;
          if (score > bestScore) {
            const r = c.getBoundingClientRect();
            best = { score, txt, x: r.x + r.width / 2, y: r.y + r.height / 2 };
            bestScore = score;
          }
        }
        return best;
      });

      if (btnInfo) {
        console.log(`CANCEL_CONFIRM_FOUND btn_txt="${btnInfo.txt}" score=${btnInfo.score} x=${Math.round(btnInfo.x)} y=${Math.round(btnInfo.y)}`);

        // Click por coordenadas con Playwright
        if (btnInfo.x > 0 && btnInfo.y > 0) {
          try {
            await page.mouse.click(btnInfo.x, btnInfo.y);
            clicked = true;
            console.log(`CANCEL_CONFIRM_OK via=coordinates x=${Math.round(btnInfo.x)} y=${Math.round(btnInfo.y)}`);
          } catch {}
        }
      }
    }

    // ── Intento 3: Locators genéricos de texto "Sí" ──
    if (!clicked) {
      const siLocators = [
        page.locator('button.ajs-ok'),
        page.locator('.alertify button:has-text("Sí")'),
        page.locator('button:has-text("Sí")'),
        page.locator('button:has-text("Si")'),
        page.locator('a:has-text("Sí")'),
      ];
      for (const loc of siLocators) {
        try {
          const cnt = await loc.count();
          if (cnt > 0) {
            await loc.first().click({ timeout: 3000 });
            clicked = true;
            console.log('CANCEL_CONFIRM_OK via=text_locator');
            break;
          }
        } catch {}
      }
    }

    if (clicked) {
      await waitForTimeoutRaw(page, 400);
      return true;
    }

    console.log('CANCEL_CONFIRM_ALL_METHODS_FAIL');
  } catch (e) {
    console.log(`CANCEL_CONFIRM_ERROR err=${(e.message || '').slice(0, 100)}`);
  }
  return false;
}

async function waitForCancellationFeedback(page, timeoutMs = 4500) {
  const started = Date.now();
  while ((Date.now() - started) < timeoutMs) {
    if (isPageClosedSafe(page)) return false;
    try {
      const ok = await page.evaluate(() => {
        const normalize = (s) =>
          (s || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const visible = (el) => {
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 20 && r.height > 10;
        };
        const nodes = Array.from(
          document.querySelectorAll('div,span,p,li,[role="alert"],.k-notification,.toast,.alert,.swal2-popup,.ajs-message')
        ).filter(visible);
        const text = normalize(nodes.map((n) => n.textContent || '').join(' | '));
        const hasCita = text.includes('cita');
        const hasCancel = text.includes('cancelad') || text.includes('finalizad') || text.includes('anulad') || text.includes('inasistencia');
        return hasCita && hasCancel;
      });
      if (ok) return true;
    } catch {}
    await waitForTimeoutRaw(page, 180);
  }
  return false;
}

async function clickSalirFromPatientModule(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    const loc = page.locator('button:has-text("Salir"), a:has-text("Salir"), [role="button"]:has-text("Salir")').first();
    if ((await loc.count()) > 0 && (await loc.isVisible())) {
      await loc.click({ force: true, timeout: 1200 });
      await waitForTimeoutRaw(page, 380);
      console.log('MODULE_SALIR_OK via=selector');
      return true;
    }
  } catch {}

  try {
    const clicked = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 8 && r.height > 8;
      };
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        try {
          el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          el.click();
          return true;
        } catch {
          return false;
        }
      };
      const nodes = Array.from(document.querySelectorAll('button,a,span,div,[role="button"],[title],[aria-label]')).filter(visible);
      const target = nodes.find((n) => {
        const t = normalize(`${n.textContent || ''} ${n.getAttribute('title') || ''} ${n.getAttribute('aria-label') || ''}`);
        return t === 'salir' || t.includes(' salir ');
      });
      if (!(target instanceof HTMLElement)) return false;
      return safeClick(target);
    });
    if (clicked) {
      await waitForTimeoutRaw(page, 380);
      console.log('MODULE_SALIR_OK via=dom_scored');
      return true;
    }
  } catch {}

  console.log('MODULE_SALIR_FAIL');
  return false;
}

async function closeEditCitaLikeModalIfOpen(page) {
  if (isPageClosedSafe(page)) return false;
  try {
    const closed = await page.evaluate(() => {
      const normalize = (s) =>
        (s || '')
          .toLowerCase()
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .replace(/\s+/g, ' ')
          .trim();
      const visible = (el) => {
        if (!el) return false;
        const st = getComputedStyle(el);
        const r = el.getBoundingClientRect();
        return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 120 && r.height > 90;
      };
      const safeClick = (el) => {
        if (!(el instanceof HTMLElement)) return false;
        try {
          el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
          el.click();
          return true;
        } catch {
          return false;
        }
      };

      const dialogs = Array.from(
        document.querySelectorAll('.k-window, .k-dialog, .modal, .rwDialog, .rwWindow, [class*="rwDialog"], [class*="rwWindow"], .RadWindow, [role="dialog"]')
      ).filter(visible);
      if (!dialogs.length) return false;

      const candidates = dialogs.filter((d) => {
        const txt = normalize(d.textContent || '');
        const isQuickCard =
          (txt.includes('programada') || txt.includes('finalizada')) &&
          txt.includes('videollamada') &&
          txt.includes('modulo');
        return txt.includes('editar cita') || isQuickCard;
      });
      if (!candidates.length) return false;
      const target = candidates[candidates.length - 1];

      const controls = Array.from(
        target.querySelectorAll('button, a, span, div, input[type="button"], input[type="submit"], [role="button"], [title], [aria-label]')
      ).filter((n) => visible(n));
      const closeBtn = controls.find((n) => {
        const t = normalize(`${n.textContent || ''} ${n.getAttribute('title') || ''} ${n.getAttribute('aria-label') || ''}`);
        return t === 'cerrar' || t.includes(' cerrar ') || t === 'x' || t.includes(' close ');
      });
      if (!(closeBtn instanceof HTMLElement)) return false;
      return safeClick(closeBtn);
    });
    if (closed) {
      await waitForTimeoutRaw(page, 120);
      console.log('CANCEL_EDIT_CITA_MODAL_CLOSED');
      return true;
    }
  } catch {}
  return false;
}

async function focusExistingAppointmentSlotForQuickAction(page, slot, options = {}) {
  if (isPageClosedSafe(page)) return { focused: false, opened: false, via: 'page_closed' };
  const forceClick = options?.forceClick === true;
  const points = [];
  if (Number.isFinite(slot?.x) && Number.isFinite(slot?.y)) {
    points.push({ x: slot.x, y: slot.y, label: 'center' });
    points.push({ x: slot.x - 4, y: slot.y, label: 'left' });
    points.push({ x: slot.x + 4, y: slot.y, label: 'right' });
    points.push({ x: slot.x, y: slot.y - 4, label: 'top' });
    points.push({ x: slot.x, y: slot.y + 4, label: 'bottom' });
  }

  for (const point of points) {
    try {
      await page.mouse.move(point.x, point.y);
      await waitForTimeoutRaw(page, 36);
      await page.evaluate(({ x, y, forceClick }) => {
        const fireMouse = (el, type, px, py) => {
          try {
            el.dispatchEvent(
              new MouseEvent(type, {
                bubbles: true,
                cancelable: true,
                view: window,
                clientX: px,
                clientY: py,
                button: 0,
                buttons: 1
              })
            );
          } catch {}
        };
        const firePointer = (el, type, px, py) => {
          try {
            el.dispatchEvent(
              new PointerEvent(type, {
                bubbles: true,
                cancelable: true,
                view: window,
                pointerId: 1,
                pointerType: 'mouse',
                isPrimary: true,
                clientX: px,
                clientY: py,
                button: 0,
                buttons: 1
              })
            );
          } catch {}
        };

        const stack = Array.from(document.elementsFromPoint(x, y) || []);
        let target = null;
        for (const el of stack) {
          if (!(el instanceof HTMLElement)) continue;
          const found = el.closest('.k-event, .rsApt, [class*="k-event"], [class*="Apt"], [class*="event"], [class*="appointment"]');
          if (found instanceof HTMLElement) {
            target = found;
            break;
          }
        }
        if (!(target instanceof HTMLElement)) return;

        const rect = target.getBoundingClientRect();
        const cx = Math.round(rect.left + rect.width / 2);
        const cy = Math.round(rect.top + rect.height / 2);
        firePointer(target, 'pointerover', cx, cy);
        firePointer(target, 'pointerenter', cx, cy);
        firePointer(target, 'pointermove', cx, cy);
        fireMouse(target, 'mouseover', cx, cy);
        fireMouse(target, 'mouseenter', cx, cy);
        fireMouse(target, 'mousemove', cx, cy);

        if (forceClick) {
          firePointer(target, 'pointerdown', cx, cy);
          fireMouse(target, 'mousedown', cx, cy);
          firePointer(target, 'pointerup', cx, cy);
          fireMouse(target, 'mouseup', cx, cy);
          fireMouse(target, 'click', cx, cy);
          try { target.click(); } catch {}
        }
      }, { x: point.x, y: point.y, forceClick });

      if (forceClick) {
        await page.mouse.click(point.x, point.y, { delay: 18 });
      }
      await waitForTimeoutRaw(page, 46);
      const opened = await isProgramadaQuickActionVisible(page);
      if (opened) return { focused: true, opened: true, via: `coords_${point.label}${forceClick ? '_click' : '_hover'}` };
    } catch {}
  }

  if (slot?.selector && Number.isInteger(slot?.domIdx)) {
    try {
      const appt = page.locator(slot.selector).nth(slot.domIdx);
      if ((await appt.count()) > 0 && (await appt.isVisible())) {
        if (forceClick) await appt.click({ force: true, timeout: 900 });
        else await appt.hover({ force: true, timeout: 900 });
        await waitForTimeoutRaw(page, 56);
        const opened = await isProgramadaQuickActionVisible(page);
        if (opened) return { focused: true, opened: true, via: `locator_${forceClick ? 'click' : 'hover'}` };
      }
    } catch {}
  }

  return { focused: points.length > 0, opened: false, via: 'not_opened' };
}

async function openModuloFromExistingAppointmentSlot(page, slot, idx = 0, options = {}) {
  if (isPageClosedSafe(page)) return { ok: false, reason: 'page_closed' };
  if (!slot || !Number.isFinite(slot.x) || !Number.isFinite(slot.y)) return { ok: false, reason: 'invalid_slot' };

  const appointmentNumber = String(slot.appointmentNumber || '').trim();
  const requireVisibleQuickStatus = options?.requireVisibleQuickStatus === true;
  const requireProgramadaQuickStatus = options?.requireProgramadaQuickStatus === true;
  const skipForceTooltip = options?.skipForceTooltip === true;
  const controlledTwoStep = options?.controlledTwoStep !== false;
  const maxAttempts = (() => {
    const n = Number(options?.maxAttempts ?? CANCEL_SLOT_MODULO_MAX_RETRIES);
    if (!Number.isFinite(n)) return CANCEL_SLOT_MODULO_MAX_RETRIES;
    return Math.min(8, Math.max(1, Math.round(n)));
  })();
  const maxFocusLoops = (() => {
    const n = Number(options?.maxFocusLoops ?? CANCEL_SLOT_REFOCUS_LOOP_MAX);
    if (!Number.isFinite(n)) return CANCEL_SLOT_REFOCUS_LOOP_MAX;
    return Math.min(8, Math.max(1, Math.round(n)));
  })();
  const effectiveFocusLoops = controlledTwoStep ? Math.min(2, maxFocusLoops) : maxFocusLoops;
  let lastReason = 'unknown';

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    if (isPageClosedSafe(page)) return { ok: false, reason: 'page_closed' };

    if (await isCatalogPacientesModalVisible(page)) {
      await closeCatalogPacientesModal(page);
      await waitForTimeoutRaw(page, 150);
    }
    if (await isNuevaCitaModalVisible(page)) {
      await closeNuevaCitaModalIfOpen(page);
      await waitForTimeoutRaw(page, 110);
    }
    await closeEditCitaLikeModalIfOpen(page);

    let clickedThisAttempt = false;
    let clickVia = 'none';
    let skipByFinalizadaRuntime = false;

    for (let focusTry = 1; focusTry <= effectiveFocusLoops; focusTry += 1) {
      const forceClick = controlledTwoStep ? true : focusTry > 1;
      const focus = await focusExistingAppointmentSlotForQuickAction(page, slot, { forceClick });
      if (controlledTwoStep) await waitForTimeoutRaw(page, 16);
      const quickState = await readQuickActionStatus(page);
      console.log(
        `CANCEL_SLOT_QUICK_STATE idx=${idx} attempt=${attempt}.${focusTry} visible=${quickState.visible ? 1 : 0} status=${quickState.status || '-'} finalizada=${quickState.isFinalizada ? 1 : 0} programada=${quickState.isProgramada ? 1 : 0} videollamada=${quickState.isVideollamada ? 1 : 0}`
      );
      console.log(
        `CANCEL_SLOT_STEP1_SELECT idx=${idx} attempt=${attempt}.${focusTry} focused=${focus.focused ? 1 : 0} opened=${focus.opened ? 1 : 0} via=${focus.via} force_click=${forceClick ? 1 : 0}`
      );
      const runtimePureFinalizada = quickState.visible && quickState.isFinalizada && !quickState.isProgramada && !quickState.isVideollamada;
      const runtimeConfirmedWithSlot = quickState.visible && quickState.isFinalizada && slot?.isFinalizada === true;
      if (runtimePureFinalizada || runtimeConfirmedWithSlot) {
        lastReason = 'slot_finalizada_quick_modal';
        skipByFinalizadaRuntime = true;
        break;
      }
      if (quickState.visible && quickState.isFinalizada && !runtimePureFinalizada && slot?.isFinalizada !== true) {
        console.log(`CANCEL_SLOT_FINALIZADA_AMBIGUOUS idx=${idx} attempt=${attempt}.${focusTry} action=continue_to_modulo`);
      }
      if (requireVisibleQuickStatus && !quickState.visible) {
        lastReason = 'quick_modal_not_visible';
        await sleepRaw(Math.max(20, CANCEL_SLOT_REFOCUS_RETRY_MS));
        continue;
      }
      if (requireProgramadaQuickStatus && quickState.visible && !(quickState.isProgramada || quickState.isVideollamada)) {
        lastReason = 'quick_status_not_programada';
        await sleepRaw(Math.max(20, CANCEL_SLOT_REFOCUS_RETRY_MS));
        continue;
      }
      const quickBefore = await isProgramadaQuickActionVisible(page);
      let quick = await clickModuloFromSavedSlotQuickAction(page, slot, {
        appointmentNumber,
        disallowFinalizada: true,
        skipForceTooltip,
        assumeQuickVisible: controlledTwoStep
      });
      let quickAfter = await isProgramadaQuickActionVisible(page);
      let modalClosed = !quick.ok && quickBefore && !quickAfter;

      // Plataforma lenta: no cambiar de casilla inmediatamente.
      // Reintenta click de Modulo dentro del mismo quick-modal antes de refocus.
      if (
        controlledTwoStep &&
        !quick.ok &&
        quickAfter &&
        (quick.via === 'quick_modal_without_modulo' || quick.via === 'quick_modulo_click_failed' || quick.via === 'quick_modal_not_visible')
      ) {
        for (let settleTry = 1; settleTry <= 3; settleTry += 1) {
          await waitForTimeoutRaw(page, 220 + (settleTry * 80));
          const visibleDuringSettle = await isProgramadaQuickActionVisible(page);
          if (!visibleDuringSettle) break;
          const settleQuick = await clickModuloFromSavedSlotQuickAction(page, slot, {
            appointmentNumber,
            disallowFinalizada: true,
            skipForceTooltip: true,
            assumeQuickVisible: true
          });
          console.log(
            `CANCEL_SLOT_STEP2_SETTLE idx=${idx} attempt=${attempt}.${focusTry}.${settleTry} ok=${settleQuick.ok ? 1 : 0} via=${settleQuick.via || 'unknown'}`
          );
          if (settleQuick.ok) {
            quick = settleQuick;
            break;
          }
        }
        quickAfter = await isProgramadaQuickActionVisible(page);
        modalClosed = !quick.ok && quickBefore && !quickAfter;
      }

      // Fallback fuerte: mantener misma casilla y usar detector robusto del tooltip
      // antes de pasar a la siguiente cita.
      if (controlledTwoStep && !quick.ok && quick.via === 'quick_modal_without_modulo') {
        const fallbackQuick = await clickModuloFromSavedSlotQuickAction(page, slot, {
          appointmentNumber,
          disallowFinalizada: true,
          skipForceTooltip: false,
          assumeQuickVisible: false
        });
        console.log(
          `CANCEL_SLOT_STEP2_FALLBACK idx=${idx} attempt=${attempt}.${focusTry} ok=${fallbackQuick.ok ? 1 : 0} via=${fallbackQuick.via || 'unknown'}`
        );
        if (fallbackQuick.ok) quick = fallbackQuick;
        quickAfter = await isProgramadaQuickActionVisible(page);
        modalClosed = !quick.ok && quickBefore && !quickAfter;
      }

      console.log(
        `CANCEL_SLOT_STEP2_MODULO idx=${idx} attempt=${attempt}.${focusTry} ok=${quick.ok ? 1 : 0} via=${quick.via || 'unknown'} modal_before=${quickBefore ? 1 : 0} modal_after=${quickAfter ? 1 : 0} modal_closed=${modalClosed ? 1 : 0} controlled=${controlledTwoStep ? 1 : 0}`
      );

      if (await isCatalogPacientesModalVisible(page)) {
        await closeCatalogPacientesModal(page);
        await waitForTimeoutRaw(page, 160);
        lastReason = 'catalog_opened_unexpected';
        continue;
      }

      if (await isNuevaCitaModalVisible(page)) {
        await closeNuevaCitaModalIfOpen(page);
        await waitForTimeoutRaw(page, 120);
        lastReason = 'nueva_cita_opened_unexpected';
        continue;
      }

      if (quick.ok) {
        clickedThisAttempt = true;
        clickVia = quick.via || 'quick_action';
        break;
      }

      lastReason = quick.via || 'modulo_click_failed';
      await sleepRaw(Math.max(260, CANCEL_SLOT_REFOCUS_RETRY_MS));
    }

    if (skipByFinalizadaRuntime) {
      return { ok: false, reason: 'slot_finalizada_quick_modal' };
    }

    if (!clickedThisAttempt) {
      console.log(`CANCEL_SLOT_MODULO_RETRY idx=${idx} attempt=${attempt} reason=${lastReason}`);
      await sleepRaw(Math.max(40, CANCEL_SLOT_REFOCUS_RETRY_MS));
      continue;
    }

    await updateBotStatusOverlay(page, 'working', 'esperando apertura Tablero Médico...');
    const loaded = await waitForModuloLoaded(page, `cancel_mode_slot_${idx}_attempt_${attempt}`, { autoPostModule: false });
    console.log(
      `CANCEL_SLOT_MODULE_LOAD idx=${idx} attempt=${attempt} loaded=${loaded ? 1 : 0} click_via=${clickVia}`
    );
    if (loaded) {
      await updateBotStatusOverlay(page, 'success', 'Tablero Médico abierto!');
      // Cerrar popup P2H que puede quedar encima del módulo
      try {
        await page.evaluate(() => {
          const popup = document.querySelector('.div_hos930AvisoPaciente');
          if (popup) popup.style.display = 'none';
        });
      } catch {}
      return { ok: true, via: clickVia };
    }

    lastReason = 'module_not_loaded_after_modulo_click';
    await sleepRaw(Math.max(40, CANCEL_SLOT_REFOCUS_RETRY_MS));
  }

  return { ok: false, reason: lastReason };
}

async function cancelOneAppointmentFromSlot(page, slot, idx = 0) {
  if (!slot || !Number.isFinite(slot.x) || !Number.isFinite(slot.y)) return { ok: false, reason: 'invalid_slot' };
  if (await isCatalogPacientesModalVisible(page)) {
    await closeCatalogPacientesModal(page);
    await waitForTimeoutRaw(page, 180);
  }
  if (await isNuevaCitaModalVisible(page)) {
    await closeNuevaCitaModalIfOpen(page);
    await waitForTimeoutRaw(page, 140);
  }

  const modulo = await openModuloFromExistingAppointmentSlot(page, slot, idx);
  if (!modulo.ok) {
    console.log(
      `CANCEL_SLOT_MODULO_FAIL idx=${idx} reason=${modulo.reason || 'unknown'} number="${slot.appointmentNumber || ''}" text="${slot.text || ''}"`
    );
    return { ok: false, reason: modulo.reason || 'modulo_click_failed' };
  }

  const clickedCancel = await clickCancelActionInModuleWithRetry(page);
  if (!clickedCancel) {
    await clickSalirFromPatientModule(page);
    await ensureCalendarContext(page);
    await applyAgendaFilter(page);
    await ensureWorkingHoursVisible(page);
    await waitForTimeoutRaw(page, 360);
    return { ok: false, reason: 'cancel_button_not_found' };
  }

  const confirmed = await confirmCancellationDialog(page);
  const feedback = await waitForCancellationFeedback(page, 5000);
  await clickSalirFromPatientModule(page);
  await ensureCalendarContext(page);
  await applyAgendaFilter(page);
  await ensureWorkingHoursVisible(page);
  await waitForTimeoutRaw(page, 520);

  console.log(
    `CANCEL_SLOT_DONE idx=${idx} confirmed=${confirmed ? 1 : 0} feedback=${feedback ? 1 : 0} number="${slot.appointmentNumber || ''}" text="${slot.text || ''}"`
  );
  return { ok: true, reason: feedback ? 'feedback_ok' : 'clicked_no_feedback' };
}

function prioritizeMode2SlotCandidates(slots, options = {}) {
  const input = Array.isArray(slots) ? slots : [];
  if (!input.length) return [];
  const maxTotal = (() => {
    const n = Number(options?.maxTotal ?? MODE2_MAX_SLOT_CANDIDATES);
    if (!Number.isFinite(n)) return MODE2_MAX_SLOT_CANDIDATES;
    return Math.min(200, Math.max(1, Math.round(n)));
  })();
  const perDayCap = (() => {
    const n = Number(options?.perDayCap ?? MODE2_MAX_SLOT_CANDIDATES_PER_DAY);
    if (!Number.isFinite(n)) return MODE2_MAX_SLOT_CANDIDATES_PER_DAY;
    return Math.min(100, Math.max(1, Math.round(n)));
  })();

  const byDay = new Map();
  const dayOrder = [];
  const sampleEvenly = (arr, k) => {
    const items = Array.isArray(arr) ? arr : [];
    const n = items.length;
    if (n <= k) return items.slice();
    const out = [];
    const seen = new Set();
    for (let i = 0; i < k; i += 1) {
      const idx = Math.round((i * (n - 1)) / Math.max(1, k - 1));
      if (seen.has(idx)) continue;
      seen.add(idx);
      out.push(items[idx]);
    }
    if (!out.length) return items.slice(0, k);
    return out;
  };
  for (const slot of input) {
    const day = String(slot?.dayIso || 'undated');
    if (!byDay.has(day)) {
      byDay.set(day, []);
      dayOrder.push(day);
    }
    byDay.get(day).push(slot);
  }

  // Mantener primeros N por dia para evitar "atascarse" en un solo dia lleno de finalizadas.
  const queues = dayOrder
    .sort((a, b) => String(a).localeCompare(String(b)))
    .map((day) => ({ day, q: sampleEvenly(byDay.get(day), perDayCap) }));

  const out = [];
  let advanced = true;
  while (out.length < maxTotal && advanced) {
    advanced = false;
    for (const item of queues) {
      if (out.length >= maxTotal) break;
      if (!item.q.length) continue;
      out.push(item.q.shift());
      advanced = true;
    }
  }

  return out;
}

async function openModuleFromExistingAppointmentInCalendar(page) {
  console.log(`MODULE_OPEN_FLOW_START search_weeks=${MODE2_MAX_SEARCH_WEEKS}`);

  // Limpiar modal "Catálogo de diagnósticos" residual de intentos previos
  await dismissCatalogoDiagnosticosModal(page);

  // Check inmediato: si el Tablero Médico ya está abierto, saltar directo
  if (await isTableroMedicoTabActive(page)) {
    console.log('MODULE_OPEN_TABLERO_ALREADY_ACTIVE skipping_agenda_scan');
    return { ok: true, scanned: 0, attempted: 0, via: 'tablero_already_active' };
  }

  // Check inmediato: detectar tab Tablero Médico (puede existir pero no estar activo)
  try {
    const earlyState = await readModuloLoadState(page);
    if (earlyState?.tableroTabExists || earlyState?.loaded) {
      console.log(`MODULE_OPEN_TABLERO_TAB_DETECTED signal=${earlyState.signal || 'tab'}`);
      return { ok: true, scanned: 0, attempted: 0, via: 'tablero_tab_detected' };
    }
  } catch {}

  // Check inmediato: popup "Nueva cita asignada" → click "Abrir módulo" antes de buscar celdas
  const earlyP2H = await dismissStaleP2HPopup(page, { clickAbrirModulo: true });
  if (earlyP2H?.action === 'abrir_modulo') {
    console.log('MODULE_OPEN_VIA_RECORDATORIO_EARLY - esperando carga...');
    await updateBotStatusOverlay(page, 'working', 'esperando apertura Tablero Médico...');
    const recLoaded = await waitForModuloLoaded(page, 'recordatorio_early', { autoPostModule: false });
    if (recLoaded) {
      console.log('MODULE_OPEN_VIA_RECORDATORIO_EARLY_OK');
      await updateBotStatusOverlay(page, 'success', 'Tablero Médico abierto!');
      return { ok: true, scanned: 0, attempted: 0, via: 'recordatorio_early' };
    }
    console.log('MODULE_OPEN_VIA_RECORDATORIO_EARLY_TIMEOUT - continuando búsqueda normal');
    await updateBotStatusOverlay(page, 'waiting', 'recordatorio no cargó, buscando celda...');
  }

  let scanned = 0;
  const attempted = new Set();
  const minBaseDate = new Date();
  minBaseDate.setHours(0, 0, 0, 0);
  minBaseDate.setDate(minBaseDate.getDate() + MODE2_SLOT_MIN_DAY_OFFSET);
  const minDayIso = minBaseDate.toISOString().slice(0, 10);
  console.log(`MODULE_OPEN_DAY_FILTER min_day_iso=${minDayIso} offset=${MODE2_SLOT_MIN_DAY_OFFSET} skip_sundays=${MODE2_SKIP_SUNDAYS ? 1 : 0} auto_filter=${MODE2_AUTO_FILTER ? 1 : 0}`);
  await ensureWorkingHoursVisible(page);
  await ensureCalendarOnCurrentWeek(page, { applyFilter: false });

  for (let week = 0; week < MODE2_MAX_SEARCH_WEEKS; week += 1) {
    if (week > 0) {
      const moved = await goToNextCalendarRange(page);
      console.log(`MODULE_OPEN_NEXT_WEEK week=${week} moved=${moved ? 1 : 0}`);
      await waitForTimeoutRaw(page, 820);
      if (MODE2_AUTO_FILTER) await applyAgendaFilter(page);
    }
    await ensureWorkingHoursVisible(page);
    await waitForTimeoutRaw(page, 620);

    // Check: si el Tablero Médico ya está abierto (por popup previo u otra razón)
    if (await isTableroMedicoTabActive(page)) {
      console.log(`MODULE_OPEN_TABLERO_ALREADY_ACTIVE week=${week}`);
      return { ok: true, scanned, attempted: attempted.size, via: 'tablero_already_active' };
    }

    // Si aparece recordatorio "Nueva cita asignada", click "Abrir módulo" directo
    const weekPopup = await dismissStaleP2HPopup(page, { clickAbrirModulo: true });
    if (weekPopup?.action === 'abrir_modulo') {
      console.log(`MODULE_OPEN_VIA_RECORDATORIO week=${week}`);
      return { ok: true, scanned, attempted: attempted.size, via: 'recordatorio_cita_asignada' };
    }

    let slots = await getExistingAppointmentSlots(page, 120, {
      preferFinalizada: false,
      excludeFinalizada: true,
      onlyProgramadaVideollamada: true,
      minDayIso,
      excludeSundays: MODE2_SKIP_SUNDAYS
    });
    let strictProgramadaVideo = true;
    if (!slots.length) {
      const relaxed = await getExistingAppointmentSlots(page, 120, {
        preferFinalizada: false,
        excludeFinalizada: true,
        onlyProgramadaVideollamada: false,
        minDayIso,
        excludeSundays: MODE2_SKIP_SUNDAYS
      });
      if (relaxed.length) {
        slots = relaxed;
        strictProgramadaVideo = false;
        console.log(`MODULE_OPEN_FALLBACK non_finalizada_enabled week=${week} slots=${slots.length}`);
      }
    }
    slots = prioritizeMode2SlotCandidates(slots, {
      maxTotal: MODE2_MAX_SLOT_CANDIDATES,
      perDayCap: MODE2_MAX_SLOT_CANDIDATES_PER_DAY
    });
    console.log(
      `MODULE_OPEN_WEEK_SCAN week=${week} slots=${slots.length} strict_video=${strictProgramadaVideo ? 1 : 0} max_candidates=${MODE2_MAX_SLOT_CANDIDATES} per_day_cap=${MODE2_MAX_SLOT_CANDIDATES_PER_DAY}`
    );
    if (!slots.length) continue;
    let runtimeFinalizadaSkips = 0;

    for (let i = 0; i < slots.length; i += 1) {
      const slot = slots[i];
      const slotKey = `${Math.round(slot.x)}|${Math.round(slot.y)}|${slot.appointmentNumber || ''}|${slot.text || ''}`;
      if (attempted.has(slotKey)) continue;
      attempted.add(slotKey);
      scanned += 1;

      if (slot.isFinalizada) {
        console.log(
          `MODULE_OPEN_IGNORE_FINALIZADA slot=${i + 1}/${slots.length} scanned=${scanned} number="${slot.appointmentNumber || ''}" text="${slot.text || ''}"`
        );
        continue;
      }
      if (strictProgramadaVideo && !slot.isProgramadaVideollamada) {
        console.log(
          `MODULE_OPEN_IGNORE_NOT_PROGRAMADA_VIDEO slot=${i + 1}/${slots.length} scanned=${scanned} status="${slot.statusHint || '-'}" number="${slot.appointmentNumber || ''}" text="${slot.text || ''}"`
        );
        continue;
      }

      console.log(
        `MODULE_OPEN_TRY slot=${i + 1}/${slots.length} scanned=${scanned} day=${slot.dayIso || '-'} status="${slot.statusHint || '-'}" number="${slot.appointmentNumber || ''}" text="${slot.text || ''}"`
      );

      // Check: si el Tablero Médico ya está abierto
      if (await isTableroMedicoTabActive(page)) {
        console.log(`MODULE_OPEN_TABLERO_ALREADY_ACTIVE_SLOT slot=${i + 1}/${slots.length}`);
        return { ok: true, scanned, attempted: attempted.size, via: 'tablero_already_active' };
      }

      // Si aparece recordatorio "Nueva cita asignada", click "Abrir módulo" directo
      const slotPopup = await dismissStaleP2HPopup(page, { clickAbrirModulo: true });
      if (slotPopup?.action === 'abrir_modulo') {
        console.log(`MODULE_OPEN_VIA_RECORDATORIO_SLOT slot=${i + 1}/${slots.length} - esperando carga...`);
        await updateBotStatusOverlay(page, 'working', 'esperando apertura Tablero Médico...');
        const recLoaded = await waitForModuloLoaded(page, `recordatorio_slot_${i + 1}`, { autoPostModule: false });
        if (recLoaded) {
          console.log(`MODULE_OPEN_VIA_RECORDATORIO_LOADED slot=${i + 1}/${slots.length}`);
          await updateBotStatusOverlay(page, 'success', 'Tablero Médico abierto!');
          return { ok: true, scanned, attempted: attempted.size, via: 'recordatorio_during_slot' };
        }
        // Si no cargó, continuar buscando en celdas normales
        console.log(`MODULE_OPEN_VIA_RECORDATORIO_TIMEOUT slot=${i + 1}/${slots.length} - continuando búsqueda`);
        await updateBotStatusOverlay(page, 'waiting', 'recordatorio no cargó, buscando celda...');
      }
      if (await isCatalogPacientesModalVisible(page)) {
        await closeCatalogPacientesModal(page);
        await waitForTimeoutRaw(page, 140);
      }
      if (await isNuevaCitaModalVisible(page)) {
        await closeNuevaCitaModalIfOpen(page);
        await waitForTimeoutRaw(page, 120);
      }
      // Cerrar "Catálogo de diagnósticos" si quedó abierto de un intento previo
      await dismissCatalogoDiagnosticosModal(page);

      const opened = await openModuloFromExistingAppointmentSlot(page, slot, scanned, {
        requireVisibleQuickStatus: true,
        requireProgramadaQuickStatus: strictProgramadaVideo,
        skipForceTooltip: false,
        controlledTwoStep: true,
        maxAttempts: 2,
        maxFocusLoops: 2
      });
      if (opened.ok) {
        console.log(
          `MODULE_OPEN_OK scanned=${scanned} via=${opened.via || '-'} number="${slot.appointmentNumber || ''}" text="${slot.text || ''}"`
        );
        return { ok: true, scanned, attempted: attempted.size, slot, via: opened.via || '' };
      }

      if (opened.reason === 'slot_finalizada_quick_modal') {
        runtimeFinalizadaSkips += 1;
        console.log(
          `MODULE_OPEN_IGNORE_FINALIZADA_RUNTIME slot=${i + 1}/${slots.length} scanned=${scanned} day=${slot.dayIso || '-'} number="${slot.appointmentNumber || ''}" text="${slot.text || ''}"`
        );
        continue;
      }

      console.log(`MODULE_OPEN_SKIP scanned=${scanned} reason=${opened.reason || 'unknown'}`);
      // Cerrar modales residuales antes del siguiente intento
      await dismissCatalogoDiagnosticosModal(page);
      if (await isTableroMedicoTabActive(page)) {
        console.log('MODULE_OPEN_SKIP_CLOSE_TABLERO cleaning residual tablero tab');
        await closeTableroMedicoTab(page);
        await waitForTimeoutRaw(page, 400);
      }
      await ensureCalendarContext(page);
      if (MODE2_AUTO_FILTER) await applyAgendaFilter(page);
      await ensureWorkingHoursVisible(page);
      await waitForTimeoutRaw(page, 400);
    }

    if (runtimeFinalizadaSkips >= slots.length && slots.length > 0) {
      console.log(`MODULE_OPEN_ABORT_ALL_FINALIZADA week=${week} scanned=${scanned} slots=${slots.length}`);
      return { ok: false, scanned, attempted: attempted.size, reason: 'all_slots_finalizada_in_range' };
    }
  }

  console.log(`MODULE_OPEN_FAIL scanned=${scanned} attempted=${attempted.size}`);
  return { ok: false, scanned, attempted: attempted.size, reason: 'no_existing_slot_opened_module' };
}

async function cancelAppointmentsFromCalendar(page) {
  console.log(
    `CANCEL_FLOW_START max_cancel=${CANCEL_MAX_APPOINTMENTS} search_weeks=${CANCEL_SEARCH_MAX_WEEKS}`
  );
  let cancelled = 0;
  let scanned = 0;
  const attempted = new Set();
  if (MODE2_AUTO_FILTER) await applyAgendaFilter(page);
  await ensureWorkingHoursVisible(page);
  await ensureCalendarOnCurrentWeek(page, { applyFilter: MODE2_AUTO_FILTER });

  for (let week = 0; week < CANCEL_SEARCH_MAX_WEEKS && cancelled < CANCEL_MAX_APPOINTMENTS; week += 1) {
    if (week > 0) {
      const moved = await goToNextCalendarRange(page);
      console.log(`CANCEL_FLOW_NEXT_WEEK week=${week} moved=${moved ? 1 : 0}`);
      await waitForTimeoutRaw(page, 920);
    }
    if (MODE2_AUTO_FILTER) await applyAgendaFilter(page);
    await ensureWorkingHoursVisible(page);
    await waitForTimeoutRaw(page, 700);

    const slots = await getExistingAppointmentSlots(page, 120);
    console.log(`CANCEL_FLOW_WEEK_SCAN week=${week} slots=${slots.length}`);
    if (!slots.length) continue;

    for (let i = 0; i < slots.length && cancelled < CANCEL_MAX_APPOINTMENTS; i += 1) {
      const slot = slots[i];
      const slotKey = `${Math.round(slot.x)}|${Math.round(slot.y)}|${slot.appointmentNumber || ''}|${slot.text || ''}`;
      if (attempted.has(slotKey)) continue;
      attempted.add(slotKey);
      scanned += 1;

      console.log(
        `CANCEL_FLOW_TRY slot=${i + 1}/${slots.length} scanned=${scanned} cancelled=${cancelled} number="${slot.appointmentNumber || ''}" text="${slot.text || ''}"`
      );

      const result = await cancelOneAppointmentFromSlot(page, slot, scanned);
      if (result.ok) {
        cancelled += 1;
        console.log(`CANCEL_FLOW_OK scanned=${scanned} cancelled=${cancelled}`);
      } else {
        console.log(`CANCEL_FLOW_SKIP scanned=${scanned} reason=${result.reason || 'unknown'}`);
        await ensureCalendarContext(page);
        if (MODE2_AUTO_FILTER) await applyAgendaFilter(page);
        await ensureWorkingHoursVisible(page);
        await waitForTimeoutRaw(page, 320);
      }
    }
  }

  console.log(
    `CANCEL_FLOW_FINISH cancelled=${cancelled} scanned=${scanned} attempted=${attempted.size} max=${CANCEL_MAX_APPOINTMENTS}`
  );
  return { cancelled, scanned, attempted: attempted.size };
}

async function waitForLoginSuccess(page, timeoutMs = 30000) {
  const started = Date.now();
  while ((Date.now() - started) < timeoutMs) {
    if (isPageClosedSafe(page)) return false;
    try {
      const state = await page.evaluate(() => {
        const visibleById = (id) => {
          const el = document.getElementById(id);
          if (!el) return false;
          const st = getComputedStyle(el);
          const r = el.getBoundingClientRect();
          return st.display !== 'none' && st.visibility !== 'hidden' && r.width > 6 && r.height > 6;
        };
        const txt = (document.body?.innerText || '').toLowerCase();
        const loginVisible = visibleById('ctl00_usercontrol2_txt9001') || visibleById('ctl00_usercontrol2_txt9004');
        const hasDashboardText =
          txt.includes('mis opciones') ||
          txt.includes('práctica médica') ||
          txt.includes('practica medica') ||
          txt.includes('agenda médica') ||
          txt.includes('agenda medica') ||
          txt.includes('catalogo de reportes');
        const href = String(location.href || '').toLowerCase();
        const urlLooksLogged = href.includes('/default') || href.includes('#t2') || href.includes('#t3');
        return { loginVisible, hasDashboardText, urlLooksLogged };
      });
      if ((state.hasDashboardText || state.urlLooksLogged) && !state.loginVisible) return true;
    } catch {}
    await waitForTimeoutRaw(page, 250);
  }
  return false;
}

async function doLoginFlow(page) {
  const norm = (s) =>
    (s || '')
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim();
  const forceLoginPrereqs = async () => {
    await page.evaluate(() => {
      try {
        const fire = (id) => {
          const ddl = window.$find && window.$find(id);
          if (!ddl) return;
          try { if (ddl.raisePropertyChanged) ddl.raisePropertyChanged('selectedItem'); } catch {}
          try { if (ddl.postback) ddl.postback(); } catch {}
        };
        fire('ctl00_usercontrol2_ddlCompany');
        fire('ctl00_usercontrol2_ddlDepartamento');
      } catch {}
      try { if (typeof window.Page_ClientValidate === 'function') window.Page_ClientValidate(); } catch {}
      const btn = document.getElementById('ctl00_usercontrol2_T9500_Login_input');
      if (btn) {
        btn.disabled = false;
        btn.removeAttribute('disabled');
      }
    });
  };

  const getSelections = async () => {
    const company = await getDdlState(page, 'ctl00_usercontrol2_ddlCompany');
    const dept = await getDdlState(page, 'ctl00_usercontrol2_ddlDepartamento');
    return { company, dept };
  };

  console.log('Paso 2: escribir usuario y contrasena');
  await page.locator('#ctl00_usercontrol2_txt9001').fill(USER);
  await page.locator('#ctl00_usercontrol2_txt9004').fill(PASSWORD);
  // Click fuera del formulario para activar validadores de controles Telerik.
  await page.mouse.click(140, 140);
  await page.waitForTimeout(500);

  console.log('Paso 3-4: seleccionar Empresa y Departamento (con verificacion fuerte)');
  const isCompanyOkDdl = (state) =>
    state.selectedValue === 'PSV' || norm(state.selectedText).includes('medical practice');
  const isDeptOkDdl = (state) =>
    state.selectedValue === 'CEX' || norm(state.selectedText).includes('practica medica');

  const waitDepartmentReadyAfterCompany = async (timeoutMs = 350) => {
    const started = Date.now();
    while ((Date.now() - started) < timeoutMs) {
      const ready = await page.evaluate(() => {
        try {
          const ddl = window.$find && window.$find('ctl00_usercontrol2_ddlDepartamento');
          const root = document.getElementById('ctl00_usercontrol2_ddlDepartamento');
          if (!ddl || !root) return false;
          const enabled = ddl.get_enabled ? !!ddl.get_enabled() : true;
          const disabledByClass = root.classList.contains('rddlDisabled');
          const disabledByAttr = root.getAttribute('disabled') !== null;
          // Items puede venir vacio hasta abrir el dropdown, no bloquear por eso.
          return enabled && !disabledByClass && !disabledByAttr;
        } catch {
          return false;
        }
      });
      if (ready) return true;
      await waitForTimeoutRaw(page, 50);
    }
    return false;
  };

  const selectLoginDdl = async ({
    ddlId,
    optionText,
    expectedValue,
    label,
    isOk,
    maxAttempts = 4,
    dropdownTiming = {},
    postAttemptWaitMs = 80
  }) => {
    let state = await getDdlState(page, ddlId);
    let ok = isOk(state);
    if (ok) {
      console.log(`select-${label}: ya seleccionado value=${state.selectedValue || '(vacio)'} "${state.selectedText || ''}"`);
      return { ok, state };
    }

    for (let i = 0; i < maxAttempts && !ok; i += 1) {
      await activateRadDropdown(page, ddlId, optionText, dropdownTiming);
      state = await getDdlState(page, ddlId);
      ok = isOk(state);
      console.log(
        `Intento select-${label} ${i + 1}: value=${state.selectedValue || '(vacio)'} "${state.selectedText || ''}"`
      );
      if (ok) break;

      await forceSelectDdlByValue(page, ddlId, expectedValue, optionText);
      state = await getDdlState(page, ddlId);
      ok = isOk(state);
      if (ok) break;

      try {
        const root = page.locator(`#${ddlId}`).first();
        await root.click({ force: true, timeout: 900 });
        await waitForTimeoutRaw(page, 90);
        await page.keyboard.press('ArrowDown');
        await page.keyboard.press('Enter');
      } catch {}
      await waitForTimeoutRaw(page, postAttemptWaitMs);
      state = await getDdlState(page, ddlId);
      ok = isOk(state);
    }

    return { ok, state };
  };

  // Fase 1: Empresa (base de referencia).
  const companyResult = await selectLoginDdl({
    ddlId: 'ctl00_usercontrol2_ddlCompany',
    optionText: 'MEDICAL PRACTICE',
    expectedValue: 'PSV',
    label: 'empresa',
    isOk: isCompanyOkDdl,
    maxAttempts: 4
  });

  if (!companyResult.ok) {
    throw new Error('No se pudo seleccionar Empresa (MEDICAL PRACTICE).');
  }

  const deptReady = await waitDepartmentReadyAfterCompany(120);
  console.log(`Departamento ready after empresa: ${deptReady ? 1 : 0}`);

  // Fase 2: Departamento (misma lógica de Empresa, para evitar retraso por camino distinto).
  const deptResult = await selectLoginDdl({
    ddlId: 'ctl00_usercontrol2_ddlDepartamento',
    optionText: 'PRACTICA MEDICA',
    expectedValue: 'CEX',
    label: 'departamento',
    isOk: isDeptOkDdl,
    maxAttempts: 3,
    postAttemptWaitMs: 30,
    dropdownTiming: {
      preClickWaitMs: 55,
      firstPopupWaitMs: 180,
      reopenWaitMs: 80,
      popupWaitMs: 620,
      optionWaitMs: 760,
      optionSettleMs: 45,
      popupHideWaitMs: 260,
      finalSettleMs: 40
    }
  });

  let deptOk = deptResult.ok;
  if (!deptOk) {
    // Ultimo intento directo por valor Telerik, solo en departamento.
    await forceSelectDdlByValue(
      page,
      'ctl00_usercontrol2_ddlDepartamento',
      'CEX',
      'PRACTICA MEDICA'
    );
    const deptState = await getDdlState(page, 'ctl00_usercontrol2_ddlDepartamento');
    deptOk = isDeptOkDdl(deptState);
  }

  if (!deptOk) {
    throw new Error('No se pudo seleccionar Departamento (PRACTICA MEDICA).');
  }

  await ensureDropdownReadyForLogin(page);
  await forceLoginPrereqs();
  const finalState = await getSelections();

  console.log(
    `Empresa final => value:${finalState.company.selectedValue || '(vacio)'} text:${finalState.company.selectedText || '(vacio)'}`
  );
  console.log(
    `Departamento final => value:${finalState.dept.selectedValue || '(vacio)'} text:${finalState.dept.selectedText || '(vacio)'}`
  );

  const btnDisabled = await page.evaluate(() => {
    const btn = document.getElementById('ctl00_usercontrol2_T9500_Login_input');
    return btn ? btn.disabled : true;
  });
  console.log(`Boton login disabled: ${btnDisabled}`);

  console.log('Paso 5: iniciar sesion');
  const loginButton = page.locator('#ctl00_usercontrol2_T9500_Login_input').first();
  if (btnDisabled) {
    await forceLoginPrereqs();
    await page.waitForTimeout(150);
  }
  for (let i = 0; i < 3; i += 1) {
    await forceLoginPrereqs();
    await waitForTimeoutRaw(page, 120);
    try {
      await loginButton.click({ timeout: 7000 });
    } catch {
      await page.evaluate(() => {
        const btn = document.getElementById('ctl00_usercontrol2_T9500_Login_input');
        if (btn) {
          btn.disabled = false;
          btn.removeAttribute('disabled');
          btn.click();
        }
      });
    }

    if (await waitForLoginSuccess(page, 15000)) {
      return;
    }

    // Fallback ASP.NET postback directo cuando el click visual no navega.
    await page.evaluate(() => {
      try {
        if (typeof __doPostBack === 'function') {
          __doPostBack('ctl00$usercontrol2$T9500$Login', '');
        }
      } catch {}
    });
    if (await waitForLoginSuccess(page, 12000)) {
      return;
    }
    console.log(`Reintento click login ${i + 1}/3 (sigue en Login)`);
    await page.waitForTimeout(350);
  }
  throw new Error('No se pudo completar login tras 3 intentos.');
}

async function runSingleFlowAttempt(attempt, totalAttempts) {
  let browser;
  try {
    browser = await chromium.launch({
      headless: false,
      slowMo: SLOW_MO_MS,
      args: [
        '--disk-cache-size=1073741824',      // 1GB de cache en disco
        '--media-cache-size=524288000',      // 500MB cache de media
        '--aggressive-cache-discard=false',  // No descartar cache agresivamente
        '--disable-background-timer-throttling', // No frenar timers en background
        '--disable-renderer-backgrounding',  // No limitar rendimiento en background
        '--disable-backgrounding-occluded-windows', // No frenar ventanas ocultas
        '--max-connections-per-host=16',     // 16 conexiones paralelas (default 6)
        '--enable-features=ParallelDownloading', // Descargas paralelas
      ]
    });
    const context = await browser.newContext({
      ignoreHTTPSErrors: true,
      viewport: null, // Usar viewport nativo del navegador (se adapta al mover entre pantallas)
    });
    context.on('page', (p) => installScaledWaitForTimeout(p));
    const page = await context.newPage();
    installScaledWaitForTimeout(page);

    console.log(`PERF_CONFIG TIMEOUT_SCALE=${TIMEOUT_SCALE} MIN_WAIT_MS=${MIN_WAIT_MS} SLOW_MO_MS=${SLOW_MO_MS}`);
    console.log(
      `FLOW_GUARDS STRICT_NUEVA_CITA_MODAL=${STRICT_NUEVA_CITA_MODAL ? 1 : 0} CLICK_SEARCH_AFTER_KEY=${CLICK_SEARCH_AFTER_KEY ? 1 : 0} ENABLE_ENTER_FALLBACK=${ENABLE_ENTER_FALLBACK ? 1 : 0} ALLOW_SAVE_ON_UNCONFIRMED_KEY=${ALLOW_SAVE_ON_UNCONFIRMED_KEY ? 1 : 0} CATALOG_LOOP_MAX=${CATALOG_LOOP_MAX} MAX_KEY_ATTEMPTS=${MAX_KEY_ATTEMPTS} PRIORITIZE_RECENT_KEYS=${PRIORITIZE_RECENT_KEYS ? 1 : 0} KEY_SELECTION_MODE=${KEY_SELECTION_MODE} KEY_RANDOM_SEED=${KEY_RANDOM_SEED || '-'} REQUIRE_SAVE_ALERT=${REQUIRE_SAVE_ALERT ? 1 : 0} KEY_SETTLE_MS=${KEY_SETTLE_MS} KEY_RESOLUTION_TIMEOUT_MS=${KEY_RESOLUTION_TIMEOUT_MS} COMMENT_CLICK_RETRIES=${COMMENT_CLICK_RETRIES} REVIEW_HOLD_MS=${REVIEW_HOLD_MS} ERROR_REVIEW_HOLD_MS=${ERROR_REVIEW_HOLD_MS} KEY_EXHAUST_REVIEW_HOLD_MS=${KEY_EXHAUST_REVIEW_HOLD_MS}`
    );
    console.log(
      `FLOW_MODULE_LOAD MODULE_LOAD_POLL_TIMEOUT_MS=${MODULE_LOAD_POLL_TIMEOUT_MS} MODULE_LOAD_POLL_INTERVAL_MS=${MODULE_LOAD_POLL_INTERVAL_MS} REVIEW_HOLD_MS=${REVIEW_HOLD_MS} AUTO_OPEN_NOTA_MEDICA_AFTER_MODULE=${AUTO_OPEN_NOTA_MEDICA_AFTER_MODULE ? 1 : 0} RELOAD_BEFORE_NOTA_MEDICA=${RELOAD_BEFORE_NOTA_MEDICA ? 1 : 0} RELOAD_BEFORE_NOTA_MEDICA_TIMEOUT_MS=${RELOAD_BEFORE_NOTA_MEDICA_TIMEOUT_MS} RELOAD_BEFORE_NOTA_MEDICA_POLL_MS=${RELOAD_BEFORE_NOTA_MEDICA_POLL_MS} NOTA_MEDICA_DELAY_MS=${NOTA_MEDICA_DELAY_MS} NOTA_MEDICA_CLICK_TIMEOUT_MS=${NOTA_MEDICA_CLICK_TIMEOUT_MS} AUTO_FILL_NOTA_MEDICA_FIELDS=${AUTO_FILL_NOTA_MEDICA_FIELDS ? 1 : 0} AUTO_CLICK_GENERAR_IA_NOTA_MEDICA=${AUTO_CLICK_GENERAR_IA_NOTA_MEDICA ? 1 : 0} NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS=${NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS} NOTA_MEDICA_FIELDS_FILL_RETRY_MS=${NOTA_MEDICA_FIELDS_FILL_RETRY_MS} AUTO_GENERAR_PLAN_TRATAMIENTO=${AUTO_GENERAR_PLAN_TRATAMIENTO ? 1 : 0} PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS=${PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS} AUTO_GENERAR_RECETA_AFTER_IA=${AUTO_GENERAR_RECETA_AFTER_IA ? 1 : 0} RECETA_AFTER_IA_WAIT_MS=${RECETA_AFTER_IA_WAIT_MS} RECETA_CLICK_TIMEOUT_MS=${RECETA_CLICK_TIMEOUT_MS}`
    );
    console.log(
      `FLOW_POST_SAVE AUTO_OPEN_MODULE_AFTER_SAVE=${AUTO_OPEN_MODULE_AFTER_SAVE ? 1 : 0} POST_SAVE_REQUIRE_ASSIGNED_MODAL=${POST_SAVE_REQUIRE_ASSIGNED_MODAL ? 1 : 0} POST_SAVE_ALLOW_GENERIC_MODULO_FALLBACK=${POST_SAVE_ALLOW_GENERIC_MODULO_FALLBACK ? 1 : 0} POST_SAVE_MAX_RETRIES=${POST_SAVE_MAX_RETRIES} POST_SAVE_RETRY_INTERVAL_MS=${POST_SAVE_RETRY_INTERVAL_MS} POST_SAVE_MODAL_CLICK_LOOP_MAX=${POST_SAVE_MODAL_CLICK_LOOP_MAX} POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS=${POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS}`
    );
    console.log(
      `FLOW_MAIN_MODE mode=${BOT_MAIN_MODE} label=${BOT_MAIN_MODE === '2' ? 'nota_medica_y_finalizar_cita_existente' : 'generar_ordenes'} cancel_max=${CANCEL_MAX_APPOINTMENTS} cancel_weeks=${CANCEL_SEARCH_MAX_WEEKS} mode2_auto_filter=${MODE2_AUTO_FILTER ? 1 : 0} mode2_max_weeks=${MODE2_MAX_SEARCH_WEEKS} mode2_min_day_offset=${MODE2_SLOT_MIN_DAY_OFFSET} mode2_skip_sundays=${MODE2_SKIP_SUNDAYS ? 1 : 0} mode2_max_slot_candidates=${MODE2_MAX_SLOT_CANDIDATES} mode2_per_day_cap=${MODE2_MAX_SLOT_CANDIDATES_PER_DAY}`
    );
    console.log(`PATIENT_KEYS_SOURCE=${PATIENT_KEYS_SOURCE} count=${PATIENT_KEYS.length}`);
    if (APPOINTMENT_MEMORY_STATE.enabled && APPOINTMENT_MEMORY_STATE.dirty) {
      persistAppointmentMemory();
    }
    if (KEY_HEALTH_STATE.enabled && KEY_HEALTH_STATE.dirty) {
      persistKeyHealth();
    }
    console.log(
      `APPOINTMENT_MEMORY enabled=${APPOINTMENT_MEMORY_STATE.enabled ? 1 : 0} file="${APPOINTMENT_MEMORY_FILE}" count=${APPOINTMENT_MEMORY_STATE.records.length} ttl_h=${APPOINTMENT_MEMORY_TTL_HOURS}`
    );
    console.log(
      `KEY_HEALTH enabled=${KEY_HEALTH_STATE.enabled ? 1 : 0} file="${KEY_HEALTH_FILE}" count=${KEY_HEALTH_STATE.records.length} ttl_h=${KEY_HEALTH_TTL_HOURS} hard_block=${KEY_HARD_BLOCK_THRESHOLD}`
    );
    console.log(`FLOW_ATTEMPT ${attempt}/${totalAttempts}`);

    console.log('Paso 1: abrir pagina inicial');
    await page.goto(START_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await updateBotStatusOverlay(page, 'working', 'iniciando sesión...');
    await page.bringToFront();
    await page.mouse.click(200, 200);
    await page.waitForTimeout(600);

    const onLoginPage = (await page.locator('#ctl00_usercontrol2_txt9001').count()) > 0;
    if (onLoginPage) {
      await doLoginFlow(page);
    } else {
      console.log('Sesion ya iniciada o pagina /Default cargada. Saltando login.');
    }

    // Validación real de sesión iniciada (no depende solo de cambio de URL).
    const loginOk = await waitForLoginSuccess(page, 30000);
    if (!loginOk) throw new Error('No se detectó sesión iniciada después del login.');
    await page.waitForTimeout(1000);
    console.log('LOGIN_OK');
    await updateBotStatusOverlay(page, 'success', 'login exitoso!');

    if (ONLY_LOGIN) {
      console.log('Modo ONLY_LOGIN activo.');
      await holdBrowserForReview(page, 'validacion visual (only-login)');
      return 0;
    }

    if (ONLY_SELECT_CALENDAR_FIELD) {
      console.log('Paso 6: seleccionar Práctica médica y asegurar calendario');
      await updateBotStatusOverlay(page, 'working', 'cargando agenda médica...');
      // Espera generosa para que el portal cargue completamente tras login
      await page.waitForTimeout(3000);
      await dismissNetworkBanners(page);

      // Modo 2: detectar si el Tablero Médico ya está abierto ANTES de buscar el calendario
      let earlyModuleOpened = false;
      if (BOT_MAIN_MODE === '2') {
        // Check 1: Tablero Médico tab ya existe en RadTabStrip
        try {
          const tableroState = await readModuloLoadState(page);
          if (tableroState?.tableroTabExists || tableroState?.loaded) {
            earlyModuleOpened = true;
            await updateBotStatusOverlay(page, 'success', 'Tablero Médico ya abierto!');
            console.log(`PASO6_EARLY_TABLERO_DETECTED signal=${tableroState.signal || 'tab'}`);
          }
        } catch {}

        // Check 2: popup P2H "Nueva cita asignada" → clickear "Abrir módulo"
        if (!earlyModuleOpened) {
          const p2hResult = await dismissStaleP2HPopup(page, { clickAbrirModulo: true });
          if (p2hResult?.action === 'abrir_modulo') {
            console.log('P2H_POPUP_ABRIR_MODULO_CLICKED - esperando carga de módulo...');
            await updateBotStatusOverlay(page, 'working', 'esperando apertura Tablero Médico...');
            const loaded = await waitForModuloLoaded(page, 'p2h_popup_direct');
            if (loaded) {
              earlyModuleOpened = true;
              await updateBotStatusOverlay(page, 'success', 'Tablero Médico abierto!');
              console.log('P2H_POPUP_MODULE_LOADED_OK');
            } else {
              await updateBotStatusOverlay(page, 'waiting', 'módulo no cargó, buscando cita normal...');
              console.log('P2H_POPUP_MODULE_LOAD_FAILED - continuando con búsqueda normal');
            }
          }
        }
      }

      // Solo asegurar calendario si NO hay módulo abierto ya
      if (!earlyModuleOpened) {
        try {
          await ensureCalendarContext(page);
        } catch (e) {
          // En Modo 2: re-check si mientras tanto apareció el Tablero
          if (BOT_MAIN_MODE === '2') {
            try {
              const retryState = await readModuloLoadState(page);
              if (retryState?.tableroTabExists || retryState?.loaded) {
                earlyModuleOpened = true;
                await updateBotStatusOverlay(page, 'success', 'Tablero Médico detectado!');
                console.log(`PASO6_TABLERO_DETECTED_AFTER_CALENDAR_FAIL signal=${retryState.signal || 'tab'}`);
              }
            } catch {}

            // En Modo 2: si el calendario falla, no es fatal - podemos buscar citas sin él
            if (!earlyModuleOpened) {
              console.log(`PASO6_CALENDAR_FAIL_MODE2_CONTINUE err=${(e.message || '').slice(0, 60)}`);
              await updateBotStatusOverlay(page, 'warning', 'calendario no cargó, intentando buscar citas...');
              // No lanzar error, continuar al flujo de búsqueda
            }
          } else {
            throw e;
          }
        }
        if (BOT_MAIN_MODE !== '2' || MODE2_AUTO_FILTER) {
          await applyAgendaFilter(page);
        }
        await ensureWorkingHoursVisible(page);
        await ensureCalendarOnCurrentWeek(page, { applyFilter: BOT_MAIN_MODE !== '2' || MODE2_AUTO_FILTER });
      }

      await dismissNetworkBanners(page);
      if (BOT_MAIN_MODE !== '2') {
        // En Modo 1: solo ocultar el popup sin clickear nada
        await dismissStaleP2HPopup(page);
      }
      await dismissNetworkBanners(page);
      if (BOT_MAIN_MODE === '2') {
        // Esperar breve a que la agenda se estabilice antes de buscar citas
        if (!earlyModuleOpened) {
          await updateBotStatusOverlay(page, 'working', 'esperando agenda...');
          await waitForTimeoutRaw(page, 2200);
          // Re-check: popup P2H pudo aparecer durante la espera
          const postWaitP2H = await dismissStaleP2HPopup(page, { clickAbrirModulo: true });
          if (postWaitP2H?.action === 'abrir_modulo') {
            console.log('PASO7_PRE_WAIT_P2H_POPUP_DETECTED - esperando carga módulo...');
            await updateBotStatusOverlay(page, 'working', 'recordatorio detectado, abriendo módulo...');
            const loaded = await waitForModuloLoaded(page, 'p2h_post_calendar_wait', { autoPostModule: false });
            if (loaded) {
              earlyModuleOpened = true;
              await updateBotStatusOverlay(page, 'success', 'Tablero Médico abierto!');
              console.log('PASO7_PRE_WAIT_P2H_MODULE_LOADED');
            }
          }
          // Re-check: Tablero tab pudo activarse
          if (!earlyModuleOpened) {
            try {
              const postWaitState = await readModuloLoadState(page);
              if (postWaitState?.tableroTabExists || postWaitState?.loaded) {
                earlyModuleOpened = true;
                console.log(`PASO7_PRE_WAIT_TABLERO_DETECTED signal=${postWaitState.signal || 'tab'}`);
                await updateBotStatusOverlay(page, 'success', 'Tablero Médico detectado!');
              }
            } catch {}
          }
        }

        const MODE2_MAX_PATIENT_RETRIES = 5;
        let mode2Success = false;

        for (let patientAttempt = 0; patientAttempt < MODE2_MAX_PATIENT_RETRIES; patientAttempt++) {
          if (isPageClosedSafe(page)) break;

          console.log(`Paso 7: abrir módulo desde cita existente (intento ${patientAttempt + 1}/${MODE2_MAX_PATIENT_RETRIES})`);
          let moduleOpen;
          if (patientAttempt === 0 && earlyModuleOpened) {
            await updateBotStatusOverlay(page, 'working', 'esperando apertura Tablero Médico...');
            moduleOpen = { ok: true, scanned: 0, attempted: 0, via: 'p2h_popup_direct' };
          } else {
            await updateBotStatusOverlay(page, 'working', `buscando cita... (${patientAttempt + 1}/${MODE2_MAX_PATIENT_RETRIES})`);
            moduleOpen = await openModuleFromExistingAppointmentInCalendar(page);
          }
          console.log(
            `MODULE_OPEN_RESULT ok=${moduleOpen.ok ? 1 : 0} scanned=${moduleOpen.scanned || 0} attempted=${moduleOpen.attempted || 0} reason=${moduleOpen.reason || '-'} attempt=${patientAttempt + 1}`
          );
          if (!moduleOpen.ok) {
            await updateBotStatusOverlay(page, 'error', 'no se pudo abrir módulo');
            throw new Error('No se logró abrir el módulo de una cita existente.');
          }
          await updateBotStatusOverlay(page, 'waiting', 'esperando apertura Tablero Médico...');
          const target = await getLoadedModuloPage(page, 90000);
          if (!target?.page) {
            console.log(`MODE2_NO_MODULE_PAGE attempt=${patientAttempt + 1} - cerrando tablero y reintentando`);
            await closeTableroMedicoTab(page);
            await waitForTimeoutRaw(page, 600);
            continue;
          }
          try { await target.page.bringToFront(); } catch {}

          // Esperar a que el Tablero Médico se estabilice antes de operar
          await updateBotStatusOverlay(target.page, 'working', 'estabilizando Tablero Médico...');
          await waitForTimeoutRaw(target.page, 4000);

          // Inyectar CSS que deshabilita selección de texto (excepto inputs/textareas)
          // y limpiar cualquier selección residual
          try {
            await target.page.evaluate(() => {
              // Inyectar estilo anti-selección si no existe
              if (!document.getElementById('bot-no-select-style')) {
                const style = document.createElement('style');
                style.id = 'bot-no-select-style';
                style.textContent = `
                  *, *::before, *::after {
                    -webkit-user-select: none !important;
                    -moz-user-select: none !important;
                    -ms-user-select: none !important;
                    user-select: none !important;
                  }
                  input, textarea, [contenteditable="true"], [contenteditable=""],
                  input *, textarea *, [contenteditable="true"] *, [contenteditable=""] * {
                    -webkit-user-select: text !important;
                    -moz-user-select: text !important;
                    -ms-user-select: text !important;
                    user-select: text !important;
                  }
                `;
                document.head.appendChild(style);
              }
              // Limpiar selección existente
              try { window.getSelection()?.removeAllRanges(); } catch {}
              try { document.activeElement?.blur(); } catch {}
            });
          } catch {}

          console.log(`Paso 8: Nota médica (validar/completar) y finalizar cita (intento ${patientAttempt + 1})`);
          await updateBotStatusOverlay(target.page, 'working', 'procesando Nota médica...');
          const notaFinalizada = await processNotaMedicaAndFinalizar(target.page, 'mode2_existing_appointment');
          if (notaFinalizada) {
            await updateBotStatusOverlay(target.page, 'success', 'cita finalizada!');
            mode2Success = true;
            break;
          }

          // Falló → cerrar modales residuales, Tablero Médico y volver a agenda
          console.log(`MODE2_PATIENT_FAIL attempt=${patientAttempt + 1} - cerrando tablero y buscando otra cita`);
          await updateBotStatusOverlay(page, 'waiting', `reintentando... (${patientAttempt + 1}/${MODE2_MAX_PATIENT_RETRIES})`);
          await dismissCatalogoDiagnosticosModal(page);
          await closeTableroMedicoTab(page);
          await waitForTimeoutRaw(page, 800);
        }

        if (!mode2Success) {
          await updateBotStatusOverlay(page, 'error', 'falló después de todos los intentos');
          throw new Error(`No se logró completar Nota médica y finalizar cita después de ${MODE2_MAX_PATIENT_RETRIES} intentos.`);
        }
      } else if (AUTO_CREATE_APPOINTMENT) {
        console.log('Paso 7: crear cita (casilla libre + clave paciente + guardar)');
        const key = await createAppointmentFromCalendar(page);
        console.log(`CITA_GUARDADA_OK clave="${key}"`);
      } else {
        let modalOpened = await isAppointmentModalVisible(page);

        if (COOP_MODE) {
          console.log('Paso 7 (co-op): tienes 20s para seleccionar manualmente una casilla.');
          if (!modalOpened) {
            modalOpened = await waitForManualWindow(page, 20000);
          }
        }

        if (!modalOpened) {
          console.log('Paso 7: seleccionar campo valido para generar cita (bot fallback)');
          try {
            const selected = await selectValidCalendarField(page);
            console.log(`CAMPO_CALENDARIO_SELECCIONADO_OK class="${selected.className}" text="${selected.text}"`);
          } catch (e) {
            console.log(`WARN selección casilla: ${e.message}`);
          }
          modalOpened = await isAppointmentModalVisible(page);
        }

        console.log(`MODAL_CITA_VISIBLE=${modalOpened}`);
        console.log('Paso 8: click en botón Módulo');
        const clickedModulo = await clickModuloButton(page, { waitBeforeMs: 0 });
        if (clickedModulo) {
          await waitForModuloLoaded(page, 'manual_step8');
        }
      }
      await holdBrowserForReview(page, 'validacion visual');
      return 0;
    }

    console.log('Paso 6: entrar a Practica medica');
    await clickIfVisible(page, 'Práctica médica', 10000);
    await page.locator('text=Agenda médica').first().waitFor({ timeout: 20000 });

    console.log('Paso 7: abrir primera cita');
    await openFirstAppointment(page);
    console.log('CITA_ABIERTA_OK');

    console.log('Demo finalizada.');
    await holdBrowserForReview(page, 'revision de demo');
    return 0;
  } catch (err) {
    const message = err?.message || String(err);
    console.error(`LOGIN_ERROR: ${message}`);
    const retryable = attempt < totalAttempts && shouldRestartFromLogin(message);
    const normalizedError = normalizeText(message);
    const isKeyExhausted =
      normalizedError.includes('no se pudo guardar la cita con ninguna clave dentro del limite') ||
      normalizedError.includes('no se pudo guardar la cita con ninguna clave');
    try {
      const pages = browser?.contexts?.()[0]?.pages?.() || [];
      if (pages.length) {
        const page = pages[0];
        await updateBotStatusOverlay(page, 'error', `falló: ${message.slice(0, 60)}`);
        await page.screenshot({ path: 'login_medico_error.png', fullPage: true });
        await page.screenshot({ path: `login_medico_error_attempt${attempt}.png`, fullPage: true });
        if (retryable) {
          console.log(
            `RECOVERY_RESTART_LOGIN intento ${attempt}/${totalAttempts} por estado bugueado. Reiniciando flujo desde login...`
          );
        } else {
          console.log('Error capturado.');
          const holdMs = isKeyExhausted ? KEY_EXHAUST_REVIEW_HOLD_MS : ERROR_REVIEW_HOLD_MS;
          await holdBrowserForReview(page, isKeyExhausted ? 'sin claves válidas (espera extendida)' : 'diagnostico de error', holdMs);
        }
      }
    } catch {}
    return retryable ? 1 : 2;
  } finally {
    try {
      if (APPOINTMENT_MEMORY_STATE.enabled && APPOINTMENT_MEMORY_STATE.dirty) persistAppointmentMemory();
      if (KEY_HEALTH_STATE.enabled && KEY_HEALTH_STATE.dirty) persistKeyHealth();
    } catch {}
    try {
      await browser?.close();
    } catch {}
  }
}

(async () => {
  BOT_MAIN_MODE = await resolveMainModeSelection();
  console.log(`BOT_MAIN_MODE_SELECTED=${BOT_MAIN_MODE} (${BOT_MAIN_MODE === '2' ? 'nota_medica_y_finalizar_cita_existente' : 'generar_ordenes'})`);
  const totalAttempts = ONLY_LOGIN ? 1 : FULL_FLOW_RETRIES;
  console.log(
    `FLOW_RETRY_CONFIG RESTART_FROM_LOGIN_ON_BUG=${RESTART_FROM_LOGIN_ON_BUG ? 1 : 0} FULL_FLOW_RETRIES=${totalAttempts} BOT_MAIN_MODE=${BOT_MAIN_MODE}`
  );
  for (let attempt = 1; attempt <= totalAttempts; attempt += 1) {
    const rc = await runSingleFlowAttempt(attempt, totalAttempts);
    if (rc === 0) {
      process.exit(0);
    }
    if (rc === 1) {
      continue;
    }
    process.exit(rc);
  }
  process.exit(2);
})();


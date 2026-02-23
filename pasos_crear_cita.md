# Pasos Para Crear Cita (Agenda Medica) - Flujo Estable

## 1. Login
- Abrir `https://mp-stg.telesalud.gob.sv/`.
- Ingresar usuario y clave.
- Seleccionar:
  - Empresa: `MEDICAL PRACTICE`
  - Departamento: `PRACTICA MEDICA`
- Click en `Iniciar sesion`.

## 2. Entrar a agenda
- Click en tarjeta `Practica medica`.
- Click en opcion `Agenda medica`.
- Esperar calendario cargado.

## 3. Seleccionar casilla disponible
- Revisar semana actual; si no hay espacio, pasar a la siguiente semana.
- Buscar celda libre (sin evento ni texto ocupado).
- Click 1: posicionar celda.
- Click 2: abrir modal `Nueva cita`.

## 4. Cargar paciente y guardar
- En `Clave documento`, ingresar clave de paciente.
- No usar lupa/buscar de catalogo.
- Completar `Comentarios` (texto base: `TEST`).
- Click unico en `Guardar`.

## 5. Validacion de exito (modo estricto)
- Solo se considera exito cuando aparece alerta verde:
  - `Se ha generado la cita con el numero [XXXX]`
- Si el modal se cierra sin alerta verde, NO es exito.
- En ese caso: reabrir `Nueva cita` y probar siguiente clave.

## 6. Manejo de errores obligatorio
- Si aparece `El paciente ya tiene una cita programada`:
  - Cerrar alerta y probar siguiente clave.
- Si aparece `404` o `Paciente no encontrado`:
  - Cerrar alerta.
  - Si se abre `Catalogo de pacientes`, cerrar con la `X` superior derecha del modal.
  - Reabrir `Nueva cita` y continuar con la siguiente clave.
- Si aparece `Catalogo de pacientes` en cualquier punto:
  - Cerrar inmediatamente y continuar flujo (no usar ese modal).

## 7. Post guardado
- Guardada la cita, volver a la casilla recien creada.
- Posicionarse sobre la casilla guardada y activar popup de accion (programada/videollamada).
- Ejecutar algoritmo de fuerza para `Modulo`:
  - Identificar evento de la cita por coordenada de casilla + columna/hora + (si existe) numero de cita.
  - Activar tooltip con hover/eventos sinteticos.
  - Hacer click real por coordenada (`mouse_xy`) sobre el boton `Modulo`.
- Control por banderas en loop rapido:
  - `clicked=true|false`: si el click a `Modulo` se ejecuto.
  - `modal_closed=true|false`: si el popup se cerro antes del click.
- Si `clicked=false` y `modal_closed=true`, volver a abrir popup en la misma casilla y reintentar en loop corto.
- Esperar carga real del modulo con polling (no asumir exito solo por el click):
  - Exito: `MODULO_CARGADO_OK`
  - Si no carga en tiempo: `MODULO_CARGADO_TIMEOUT`
- Al cargar modulo:
  - esperar `2s`
  - click en `Nota medica` del menu lateral
  - logs esperados: `NOTA_MEDICA_CLICK_OK` (si falla: `NOTA_MEDICA_CLICK_TIMEOUT`)
- Log esperado al entrar: `ACCESO_MODULO_PACIENTE_OK`.
- En codigo este bloque vive como `Post-Save Strategy` (flujo rapido sobre popup de celda).
- Si el click a `Modulo` falla, reintenta rapido sobre la misma casilla guardada.
- Bandera de control en logs:
  - `MODULO_BTN_FLAG attempt=N clicked=1|0`
  - `MODULO_BTN_FLAG_FINAL clicked=1|0 loaded=1|0`
  - `MODULO_FORCE_LOOP attempt=A.B clicked=1|0 modal_before=1|0 modal_after=1|0 modal_closed=1|0`

## 8. Fuente de claves
- Archivo principal: `patient_keys.txt`
- Ruta: `C:\Users\NyGsoft\Desktop\bot ordenes m\patient_keys.txt`
- Regla actual: selección aleatoria de claves (`KEY_SELECTION_MODE=random`).

## 9. Configuracion recomendada (estabilidad)
- `REQUIRE_SAVE_ALERT=1`: exito solo con alerta verde de cita generada.
- `AUTO_OPEN_MODULE_AFTER_SAVE=1`: activar paso post-guardado (casilla -> modulo).
- `POST_SAVE_MAX_RETRIES=4`: reintentos rapidos para click en `Modulo`.
- `POST_SAVE_RETRY_INTERVAL_MS=90`: espera corta entre reintentos.
- `POST_SAVE_MODAL_CLICK_LOOP_MAX=3`: maximo de mini-reintentos internos cuando el popup se cierra antes del click.
- `POST_SAVE_MODAL_CLICK_LOOP_RETRY_MS=70`: espera corta entre mini-reintentos internos del popup.
- `POST_SAVE_ALLOW_GENERIC_MODULO_FALLBACK=0`: evita fallback generico para no tocar elementos incorrectos.
- `MODULE_LOAD_POLL_TIMEOUT_MS=45000`: tiempo maximo para detectar modulo cargado.
- `MODULE_LOAD_POLL_INTERVAL_MS=350`: frecuencia de polling de carga.
- `AUTO_OPEN_NOTA_MEDICA_AFTER_MODULE=1`: activar click automatico en `Nota medica`.
- `NOTA_MEDICA_DELAY_MS=2000`: espera antes de click en `Nota medica`.
- `NOTA_MEDICA_CLICK_TIMEOUT_MS=12000`: timeout maximo para encontrar/click `Nota medica`.
- `REVIEW_HOLD_MS=120000`: mantener navegador abierto en test para validar visualmente.
- `MAX_KEY_ATTEMPTS=0`: usa automaticamente todas las claves del txt.
- `PRIORITIZE_RECENT_KEYS=0`: mantener orden del archivo para evitar cambios de prioridad.
- `KEY_SELECTION_MODE=random`: selección aleatoria de claves (modo actual).
- `KEY_SELECTION_MODE=recent_then_random`: primero recientes exitosas (memoria), luego aleatorio.

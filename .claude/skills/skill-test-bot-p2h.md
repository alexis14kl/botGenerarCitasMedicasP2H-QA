---
name: skill-test-bot-p2h
description: Guia maestra para testing del bot de citas medicas P2H usando Google DevTools MCP. Cubre el flujo completo desde login hasta finalizacion de cita. Usar cuando se necesite testear, depurar o validar cualquier paso del bot en mp-stg.telesalud.gob.sv.
---

# Skill Test Bot P2H - Guia de Testing con DevTools MCP

## Objetivo
Servir como referencia completa para testear y depurar el bot de citas medicas P2H
(`login_medico_auto.js` - 10,889 lineas) usando herramientas de Google DevTools MCP
sobre Chrome con debugging remoto.

---

## Prerequisitos

### 1. Chrome con debugging remoto
```cmd
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%TEMP%\chrome-devtools-mcp" https://mp-stg.telesalud.gob.sv/
```

### 2. MCP configurado
Archivo `.mcp.json` en raiz del proyecto:
```json
{
  "mcpServers": {
    "google-devtools": {
      "command": "npx",
      "args": ["-y", "chrome-devtools-mcp@0.17.3", "--url=http://localhost:9222"]
    }
  }
}
```

### 3. Herramientas MCP disponibles
| Herramienta | Uso principal |
|-------------|---------------|
| `take_screenshot` | Capturar estado visual de la pagina |
| `take_snapshot` | Arbol de accesibilidad (encontrar elementos) |
| `click` | Click en elemento por selector/texto |
| `click_at` | Click por coordenadas x,y |
| `fill` | Escribir texto en inputs |
| `fill_form` | Llenar multiples campos |
| `press_key` | Teclas (Tab, Enter, etc.) |
| `hover` | Hover sobre elemento |
| `navigate_page` | Navegar a URL |
| `evaluate_script` | Ejecutar JS en la pagina |
| `wait_for` | Esperar texto visible |
| `list_network_requests` | Inspeccionar trafico HTTP |
| `get_network_request` | Detalle de request/response |
| `list_console_messages` | Ver console.log del navegador |
| `list_pages` | Listar pestanas abiertas |
| `select_page` | Cambiar de pestana |

---

## Credenciales del entorno QA/Staging
- **URL**: `https://mp-stg.telesalud.gob.sv/`
- **Usuario**: `MEDICO09`
- **Contrasena**: `ISSS202`
- **Empresa**: `MEDICAL PRACTICE` (value: `PSV`)
- **Departamento**: `PRACTICA MEDICA` (value: `CEX`)

---

## Arquitectura del Bot (login_medico_auto.js)

### Archivo principal
- `login_medico_auto.js` (~446KB, 10,889 lineas) - Monolito Node.js con Playwright
- `login_medico_auto.py` - Orquestador Python que delega al JS

### Modos de operacion
- **Modo 1** (`BOT_MAIN_MODE=1`): Generar ordenes (crear cita + nota medica + receta + plan)
- **Modo 2** (`BOT_MAIN_MODE=2`): Cancelar/finalizar citas existentes

### Funciones principales y sus lineas
| Funcion | Linea | Proposito |
|---------|-------|-----------|
| `runSingleFlowAttempt()` | 10697 | Orquestador principal del flujo |
| `doLoginFlow()` | 10466 | Login completo |
| `ensureCalendarContext()` | 1064 | Preparar vista calendario |
| `ensureCalendarOnCurrentWeek()` | 1502 | Centrar en semana actual |
| `findAvailableCalendarCell()` | 1565 | Buscar casilla libre |
| `createAppointmentFromCalendar()` | 9075 | Crear cita (modal + clave + guardar) |
| `setClaveDocumentoAndTriggerSearch()` | 7676 | Ingresar clave paciente |
| `clickGuardarNuevaCita()` | 8850 | Click en Guardar |
| `openModuloAfterAppointmentSave()` | 6771 | Abrir modulo post-guardado |
| `clickModuloButton()` | 2534 | Click boton Modulo |
| `waitForModuloLoaded()` | 2723 | Esperar carga modulo |
| `openNotaMedicaFromSidebar()` | 3065 | Abrir Nota medica |
| `fillNotaMedicaAntecedentesAndGenerateIA()` | 4216 | Llenar nota + generar IA |
| `clickGenerarIaByHumanAction()` | 3481 | Click Generar IA |
| `generarReceta()` | 4183 | Generar receta |
| `ensurePlanTratamientoAndGenerate()` | 5568 | Plan de tratamiento |
| `clickFinalizarCitaInModule()` | 5625 | Finalizar cita |
| `processNotaMedicaAndFinalizar()` | 5750 | Flujo completo nota+finalizar |
| `buildPatientKeyAttemptPlan()` | 850 | Plan de claves con health tracking |

---

## FLUJO COMPLETO DE TESTING (Modo 1 - Generar Ordenes)

### FASE 1: LOGIN
**Objetivo**: Autenticarse en el portal.

**Pasos DevTools MCP**:
1. `navigate_page` a `https://mp-stg.telesalud.gob.sv/`
2. `wait_for` texto "Nombre de usuario" o "Account login"
3. `take_snapshot` para identificar campos del formulario
4. `fill` campo usuario (`#ctl00_usercontrol2_txt9001`) con `MEDICO09`
5. `fill` campo contrasena (`#ctl00_usercontrol2_txt9004`) con `ISSS202`
6. Seleccionar Empresa via JS Telerik:
   ```js
   // evaluate_script:
   const ddl = $find('ctl00_usercontrol2_ddlCompany');
   const item = ddl.findItemByValue('PSV');
   ddl.trackChanges(); item.select(); ddl.commitChanges();
   ddl.raisePropertyChanged('selectedItem'); ddl.postback();
   ```
7. Esperar que Departamento se habilite, luego seleccionar:
   ```js
   const ddl = $find('ctl00_usercontrol2_ddlDepartamento');
   const item = ddl.findItemByValue('CEX');
   ddl.trackChanges(); item.select(); ddl.commitChanges();
   ddl.raisePropertyChanged('selectedItem'); ddl.postback();
   ```
8. Habilitar boton login:
   ```js
   const btn = document.getElementById('ctl00_usercontrol2_T9500_Login_input');
   btn.disabled = false; btn.removeAttribute('disabled');
   ```
9. `click` en `#ctl00_usercontrol2_T9500_Login_input`
10. `wait_for` texto "DoctorSV" o "Mis Opciones"
11. `take_screenshot` para confirmar login exitoso

**Validacion**: URL cambia a `/Default...`, se ve dashboard "DoctorSV".

**IDs clave del formulario login**:
- Input usuario: `#ctl00_usercontrol2_txt9001`
- Input contrasena: `#ctl00_usercontrol2_txt9004`
- Dropdown empresa: `#ctl00_usercontrol2_ddlCompany`
- Dropdown departamento: `#ctl00_usercontrol2_ddlDepartamento`
- Boton login: `#ctl00_usercontrol2_T9500_Login_input`

---

### FASE 2: NAVEGACION A AGENDA MEDICA
**Objetivo**: Llegar al calendario de citas.

**Pasos DevTools MCP**:
1. `take_snapshot` para ver opciones disponibles post-login
2. `click` en tarjeta "Practica medica" (o texto similar)
3. `click` en opcion "Agenda medica"
4. `wait_for` texto del calendario o estructura `.k-scheduler`
5. `take_screenshot` para verificar calendario visible
6. Si existe boton "Filtrar", hacer `click` en el
7. Si hay checkbox "Mostrar horas laborales", activarlo

**Validacion**: Calendario Kendo visible con dias y horas.

**Selectores del calendario (Kendo Scheduler)**:
- Contenedor: `.k-scheduler`
- Celdas del body: `.k-scheduler-content table tbody td`
- Navegacion: botones con texto "Hoy", flechas anterior/siguiente
- Rango visible: `.k-lg-date-format` o `.k-sm-date-format`

---

### FASE 3: SELECCION DE CASILLA Y CREAR CITA
**Objetivo**: Encontrar slot libre, abrir modal "Nueva cita", ingresar paciente y guardar.

**Pasos DevTools MCP**:
1. `take_snapshot` para mapear celdas del calendario
2. Verificar semana actual con `evaluate_script`:
   ```js
   document.querySelector('.k-lg-date-format')?.textContent
   ```
3. Si no es semana actual, `click` en "Hoy"
4. Buscar celda libre (sin eventos):
   ```js
   // evaluate_script - obtener celdas disponibles:
   const cells = document.querySelectorAll('.k-scheduler-content table tbody td:not(:first-child)');
   // filtrar las que no tengan eventos overlay
   ```
5. `click_at` en coordenadas de celda libre (doble click controlado)
6. `wait_for` texto "Nueva cita" (modal)
7. `take_snapshot` para ver controles del modal

**Llenado del modal**:
8. `fill` campo clave documento con clave de paciente (ej: `00955873-3`)
   - Input ID: `ctl00_nc002_MP_HOS930_MP_HOS930_panelExpExistente_MP_HOS930_altaclavedoc`
   - **IMPORTANTE**: NO hacer click en lupa/buscar. Solo fill + Tab/click fuera.
9. Esperar resolucion del paciente (~900ms KEY_SETTLE_MS)
10. `click` en campo Comentarios y `fill` con `TEST`
11. `click` en boton "Guardar" (click unico)
12. `wait_for` alerta verde: "Se ha generado la cita con el numero"
13. `take_screenshot` para confirmar

**Manejo de errores**:
- Si aparece "El paciente ya tiene una cita programada": cerrar alerta, probar otra clave
- Si aparece "Catalogo de pacientes": cerrar con X, NO usar ese modal
- Si el modal no se cierra (`modal_still_open`): reintentar con otra clave
- Si `404` o "Paciente no encontrado": siguiente clave

**Claves de paciente de referencia** (10 de 251 disponibles):
```
00955873-3, 06169373-5, 05608981-6, 06416857-7, 05400186-2
B04676303, B02661296, B01700785, B00838396, B00491489
```

**Archivo completo**: `patient_keys.txt` (251 claves)

---

### FASE 4: POST-GUARDADO - ABRIR MODULO
**Objetivo**: Desde la cita recien creada, abrir el modulo del paciente.

**Pasos DevTools MCP**:
1. Posicionar sobre la casilla de la cita guardada con `hover`
2. Esperar que aparezca tooltip/popup de la cita
3. `take_snapshot` para ver botones del popup
4. `click` en boton "Modulo" del popup
   - Si el popup se cierra antes del click, reabrir con hover y reintentar
   - Maximo 4 reintentos (`POST_SAVE_MAX_RETRIES=4`)
5. `wait_for` carga del modulo del paciente (polling hasta 45s)
6. `take_screenshot` para confirmar modulo cargado

**Validacion**: Vista del modulo del paciente con menu lateral visible.

**Banderas de control en logs**:
- `MODULO_BTN_FLAG attempt=N clicked=1|0`
- `MODULO_CARGADO_OK` / `MODULO_CARGADO_TIMEOUT`

---

### FASE 5: NOTA MEDICA
**Objetivo**: Llenar la nota medica y generar con IA.

**Pasos DevTools MCP**:
1. Esperar 2s (`NOTA_MEDICA_DELAY_MS`)
2. `click` en "Nota medica" del menu lateral izquierdo
3. `wait_for` que la vista de nota medica este activa
4. `take_snapshot` para mapear campos de nota
5. Llenar campo "Antecedentes" / "Motivo de consulta" con texto clinico:
   ```
   El paciente refiere inicio de los sintomas hace 3 dias, caracterizados por fiebre
   de predominio generalizado, sin localizacion especifica...
   ```
6. `click` en boton "Generar IA" (si `AUTO_CLICK_GENERAR_IA_NOTA_MEDICA=1`)
7. Esperar respuesta IA (hasta 18s `NOTA_MEDICA_FIELDS_FILL_TIMEOUT_MS`)
8. `take_screenshot` para verificar campos llenados por IA

**Validacion**: Campos de nota medica rellenados, sin alertas de campos requeridos.

**Log esperado**: `NOTA_MEDICA_CLICK_OK`

---

### FASE 6: RECETA
**Objetivo**: Generar receta medica.

**Pasos DevTools MCP**:
1. Esperar 3.5s despues de IA (`RECETA_AFTER_IA_WAIT_MS`)
2. `take_snapshot` para buscar boton de Receta
3. `click` en boton "Receta"
4. `wait_for` modal de receta visible
5. Completar campos si necesario
6. `take_screenshot` para confirmar

**Funciones del bot**: `generarReceta()` (L4183), `clickRecetaButton()` (L3855)

---

### FASE 7: PLAN DE TRATAMIENTO
**Objetivo**: Generar plan de tratamiento.

**Pasos DevTools MCP**:
1. `take_snapshot` para localizar seccion "Plan de tratamiento"
2. Verificar estado del formulario con `evaluate_script`
3. Si campo vacio, `fill` con texto:
   ```
   Plan breve: manejo ambulatorio, hidratacion, reposo relativo y control en 48 horas.
   ```
4. `click` en boton "Generar" del plan
5. `wait_for` confirmacion (hasta 12s `PLAN_TRATAMIENTO_GENERAR_TIMEOUT_MS`)
6. `take_screenshot`

**Funciones del bot**: `ensurePlanTratamientoAndGenerate()` (L5568)

---

### FASE 8: FINALIZAR CITA
**Objetivo**: Marcar la cita como finalizada.

**Pasos DevTools MCP**:
1. `take_snapshot` para buscar boton "Finalizar"
2. `click` en boton "Finalizar cita"
3. Si hay dialogo de confirmacion, confirmar
4. `wait_for` feedback de finalizacion exitosa
5. `take_screenshot` para evidencia final

**Funciones del bot**: `clickFinalizarCitaInModule()` (L5625)

---

## FLUJO DE TESTING (Modo 2 - Cancelar/Finalizar Citas Existentes)

### Diferencias con Modo 1
1. No crea cita nueva, busca citas existentes "Programada - Videollamada"
2. Escanea desde dia actual + offset (default +1 dia)
3. Excluye domingos (`MODE2_SKIP_SUNDAYS=1`)
4. Abre modulo desde cita existente
5. Ejecuta nota medica + receta + plan + finalizar

**Funciones clave**:
- `openModuleFromExistingAppointmentInCalendar()` (L10248)
- `cancelAppointmentsFromCalendar()` (L10377)
- `cancelOneAppointmentFromSlot()` (L10145)

---

## PATRONES DE DEPURACION

### Problema: modal_still_open al guardar
El modal "Nueva cita" no se cierra despues de click en Guardar.
**Diagnostico con DevTools**:
1. `take_screenshot` para ver estado visual
2. `evaluate_script` para verificar si hay alertas ocultas:
   ```js
   document.querySelectorAll('.rwDialogPopup, .RadAlert, .alert').length
   ```
3. `list_console_messages` para ver errores JS
4. `list_network_requests` para ver si el POST de guardar se envio

### Problema: Catalogo de pacientes aparece
La lupa de busqueda se activo involuntariamente.
**Diagnostico**:
1. `take_snapshot` para confirmar modal de catalogo
2. Cerrar con `click` en X del modal
3. Verificar que NO se hizo click en la lupa

### Problema: Dropdown no selecciona valor
Empresa o Departamento no toma el valor correcto.
**Diagnostico**:
1. `evaluate_script` para leer estado:
   ```js
   const ddl = $find('ctl00_usercontrol2_ddlCompany');
   const item = ddl.get_selectedItem();
   ({ value: item.get_value(), text: item.get_text() })
   ```
2. Verificar si Departamento esta habilitado despues de Empresa

### Problema: Calendario no carga
**Diagnostico**:
1. `evaluate_script`:
   ```js
   document.querySelector('.k-scheduler') !== null
   ```
2. `list_network_requests` para ver si hay errores de carga
3. `take_screenshot` para estado visual

---

## SISTEMA DE PERSISTENCIA

### Memoria de citas (`appointment_memory_tmp.json`)
- Registra citas exitosas con patientKey, appointmentNumber, slot, timestamp
- TTL: 72 horas
- Max: 800 registros
- Evita repetir pacientes con cita reciente

### Salud de claves (`patient_key_health_tmp.json`)
- Registra exitos/fallos por clave
- TTL: 168 horas (7 dias)
- Hard block: 2 fallos consecutivos = clave bloqueada temporalmente
- Max: 6000 registros

### Seleccion de claves (`KEY_SELECTION_MODE`)
- `random`: aleatorio (default)
- `sequential`: en orden del archivo
- `recent_then_random`: primero exitosas recientes, luego aleatorio

---

## VARIABLES DE ENTORNO CRITICAS

### Timing y performance
| Variable | Default | Descripcion |
|----------|---------|-------------|
| `TIMEOUT_SCALE` | 0.85 | Multiplicador de waits (< 1 = mas rapido) |
| `MIN_WAIT_MS` | 60 | Wait minimo en ms |
| `SLOW_MO_MS` | 130 | Slow motion de Playwright |
| `KEY_SETTLE_MS` | 900 | Espera tras ingresar clave paciente |
| `KEY_RESOLUTION_TIMEOUT_MS` | 4200 | Timeout resolucion de paciente |

### Control de flujo
| Variable | Default | Descripcion |
|----------|---------|-------------|
| `BOT_MAIN_MODE` | 1 | 1=generar ordenes, 2=cancelar cita |
| `REQUIRE_SAVE_ALERT` | 1 | Exito solo con alerta verde |
| `AUTO_OPEN_MODULE_AFTER_SAVE` | 1 | Abrir modulo post-guardado |
| `AUTO_FILL_NOTA_MEDICA_FIELDS` | 1 | Llenar nota medica |
| `AUTO_CLICK_GENERAR_IA_NOTA_MEDICA` | 1 | Click automatico en Generar IA |
| `AUTO_GENERAR_PLAN_TRATAMIENTO` | 1 | Generar plan de tratamiento |
| `AUTO_GENERAR_RECETA_AFTER_IA` | 1 | Generar receta post-IA |
| `REVIEW_HOLD_MS` | 1800000 | Pausa para revision visual (30 min) |
| `MAX_KEY_ATTEMPTS` | 0 (=auto) | 0 usa todas las claves del txt |
| `FULL_FLOW_RETRIES` | 1 | Reintentos del flujo completo |

### Post-guardado
| Variable | Default | Descripcion |
|----------|---------|-------------|
| `POST_SAVE_MAX_RETRIES` | 4 | Reintentos click en Modulo |
| `POST_SAVE_RETRY_INTERVAL_MS` | 90 | Espera entre reintentos |
| `POST_SAVE_MODAL_CLICK_LOOP_MAX` | 3 | Mini-reintentos si popup se cierra |
| `MODULE_LOAD_POLL_TIMEOUT_MS` | 45000 | Timeout carga de modulo |

---

## ARCHIVOS DEL PROYECTO

| Archivo | Proposito |
|---------|-----------|
| `login_medico_auto.js` | Core del bot (10,889 lineas) |
| `login_medico_auto.py` | Orquestador Python |
| `login_medico_auto.bat` | Lanzador basico Windows |
| `login_medico_node.bat` | Lanzador Node con config |
| `login_medico_python.bat` | Lanzador Python con config |
| `patient_keys.txt` | 251 claves de pacientes |
| `appointment_memory_tmp.json` | Memoria de citas (56 registros) |
| `patient_key_health_tmp.json` | Salud de claves (304 registros) |
| `pasos_crear_cita.md` | Documentacion del flujo de cita |
| `generar_ordenes.md` | Documentacion del flujo de ordenes |
| `package.json` | Dependencia: playwright ^1.58.2 |

---

## CHECKLIST DE TESTING RAPIDO

- [ ] Chrome abierto con `--remote-debugging-port=9222`
- [ ] MCP google-devtools conectado (verificar con `list_pages`)
- [ ] Login exitoso (URL contiene `/Default`)
- [ ] Agenda medica visible (`.k-scheduler` existe)
- [ ] Casilla libre encontrada y modal "Nueva cita" abierto
- [ ] Clave de paciente aceptada (sin catalogo, sin 404)
- [ ] Cita guardada con alerta verde
- [ ] Modulo del paciente cargado
- [ ] Nota medica llenada y IA generada
- [ ] Receta generada
- [ ] Plan de tratamiento generado
- [ ] Cita finalizada

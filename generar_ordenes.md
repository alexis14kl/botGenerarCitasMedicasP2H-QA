# Generar Ordenes - Flujo Aprendido

## Contexto
- Sistema: `mp-stg.telesalud.gob.sv`
- Usuario de práctica: `MEDICO09`
- Módulo operativo: Nota médica / barra lateral de acciones

## Flujo base
1. Entrar a cita activa desde Agenda.
2. Ir a Nota médica.
3. Abrir módulo de orden (Laboratorio / Imagenología / Procedimientos).
4. Seleccionar diagnóstico si el formulario lo requiere.
5. Agregar estudio/procedimiento y observaciones.
6. Verificar que quede en listado.
7. Guardar y confirmar modal final.

## Flujo laboratorio (practicado)
1. En `Nota médica`, abrir la acción de `Laboratorio` (u opción de órdenes de estudio).
2. En el campo `Estudio`, limpiar valor previo si existe.
3. Dar clic en `...` para abrir el catálogo de estudios.
4. Buscar y seleccionar el estudio correcto (ejemplos usados: `MONITOREO DE PRESION ARTERIAL (MAPAA)` y `COLESTEROL TOTAL`).
5. Confirmar selección y volver al formulario.
6. Dar clic en `Agregar`.
7. Dar clic en `Guardar`.
8. Ir a `Estudios complementarios` y validar que el estudio quede en estado `Pendiente` con fecha del día.
9. Si se requiere otra orden, repetir desde el paso 2.

## Atajos operativos
- Si la pantalla queda congelada o no responde: recargar la página y continuar desde el último paso confirmado.
- Si una cita/paciente no deja avanzar: usar otra clave disponible y reintentar el flujo.

## Validaciones clave
- No dejar campos obligatorios vacíos.
- Confirmar registro visible en grilla/listado.
- Si se congela o queda raro: recargar y revalidar último paso.

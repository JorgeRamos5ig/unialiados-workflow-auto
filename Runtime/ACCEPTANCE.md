# Verificación de aceptación — 2026-10-30

## Verificaciones determinísticas

- `node Runtime/workflow_static_check.cjs`: 56 nodos, 25 bloques de JavaScript compilados, cero conexiones huérfanas; workflow inactivo, envío apagado y horario deshabilitado.
- `node Runtime/acceptance_test.js __regression_only__`: PASS para preparación, español legítimo con «Esta/también», mezcla de inglés, CTA duplicado, dos viñetas, supuestos no sustentados, afirmaciones AXA internacionales y mezcla de campañas.
- `Importar_Unialiados_Optimizado.ps1`: sintaxis PowerShell válida.
- n8n 2.11.3 importó el flujo en un perfil SQLite aislado. La importación no activa ni ejecuta el workflow.
- El XLSM inspeccionado en solo lectura conserva VBA, la tabla `BaseLeads` (36 columnas) y `Enviar ahora`.

## Generación con Ollama

Modelo: `qwen3.5:4b`, API local `http://127.0.0.1:11434/api/chat`.

La última matriz completa previa a los ajustes dirigidos fue 7/9 aprobados y 2 retenidos (Unialiados THIRD por mezcla de campañas; AXA FIRST por quedar bajo el mínimo de palabras). En una re-prueba dirigida posterior, Unialiados THIRD pasó tras tres intentos, y AXA FIRST fue retenido tras tres intentos por incluir texto interno («La campaña busca…»). Una re-prueba de Unisoft SECOND aprobó formato, pero la lectura encontró una afirmación sobre archivos dispersos no sustentada; se añadió una regresión para retener esa fórmula.

No se ha repetido la matriz completa con todos los últimos ajustes. El resultado generativo final se considera pendiente, no 9/9.

## Decisión de producción

**No listo para producción.** `sendEnabled=false` y `senderEmail=PENDIENTE_CONFIGURAR`. Faltan confirmar identidad del remitente, credencial SMTP, proceso de bajas, límites diarios/mensuales, tiempos de seguimiento y matriz completa de exclusiones AXA. Se requiere selección manual de leads/campaña y revisión humana de los borradores.

No se enviaron correos, no se activó el workflow y no se realizó prueba integral con Excel/SMTP/Tavily.

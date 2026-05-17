# GUÍA: Cómo Usar Claude Code Eficientemente
# Para el Proyecto Sistema POS

---

## ¿QUÉ ES CLAUDE CODE?

Claude Code es un agente de línea de comandos que puede leer archivos de tu proyecto,
escribir código, ejecutar comandos y navegar tu repositorio de manera autónoma.
Es diferente a esta conversación de chat porque tiene acceso DIRECTO a tus archivos.

---

## ESTRUCTURA DE ARCHIVOS QUE DEBES TENER EN TU REPO

```
Sistema/
├── CLAUDE.md                    ← Memoria del agente (este proyecto)
├── skills/
│   ├── sql-server-profesional.md
│   └── vbnet-arquitectura.md
├── hooks/
│   └── guardrails.md
├── agents/
│   ├── agente-database.md
│   ├── agente-backend.md
│   └── agente-ui.md
├── Sistema.sln
├── Sistema.Entidades/
├── Sistema.Datos/
├── Sistema.Negocio/
├── Sistema.Presentacion/
└── Database/
    ├── 01_Tablas.sql
    ├── 02_Vistas.sql
    ├── 03_StoredProcedures.sql
    ├── 04_Triggers.sql
    ├── 05_DatosPrueba.sql
    └── 06_Pruebas.sql
```

> ⚠️ IMPORTANTE: El `CLAUDE.md` en la raíz del proyecto es leído automáticamente
> por Claude Code al inicio de cada sesión. Es tu "memoria persistente".

---

## 10 REGLAS DE ORO PARA USAR CLAUDE CODE

### Regla 1: Siempre da contexto al inicio
```bash
# ❌ Malo
"Crea el trigger"

# ✅ Bueno
"Estoy trabajando en el módulo de Ventas del Sistema POS (VB.NET + SQL Server).
 Necesito el trigger TRG02 que descuenta stock cuando se inserta en DetalleVenta.
 El trigger debe validar que el stock no quede negativo."
```

### Regla 2: Un módulo a la vez
No pidas múltiples módulos simultáneamente. Claude Code trabaja mejor
cuando tiene un objetivo claro y concreto.

```bash
# ❌ Malo
"Crea todas las clases de Ventas, el trigger y el formulario"

# ✅ Bueno — sesión 1
"Crea la entidad EVenta y EDetalleVenta"

# ✅ Bueno — sesión 2 (después de revisar)
"Ahora crea DVenta.vb con todos sus métodos CRUD"
```

### Regla 3: Pide revisión antes de continuar
Después de que Claude Code genere código, REVÍSALO antes de pedir más.
Es más fácil corregir 1 archivo que 5 archivos con el mismo error.

### Regla 4: Usa los archivos del Agent Dev Kit
Al inicio de sesión, di:
```
"Lee CLAUDE.md y luego el skill de [sql-server / vbnet] antes de empezar"
```

### Regla 5: Sé específico con los nombres
```bash
# ❌ Ambiguo
"Crea el stored procedure de ventas"

# ✅ Específico
"Crea sp_VentaInsertar que reciba los parámetros de EIngreso y retorne
 el @idVenta como OUTPUT, usando la plantilla de la skill sql-server-profesional.md"
```

### Regla 6: Indica el archivo de destino
```bash
"Crea el archivo Sistema.Datos/DVenta.vb con el patrón de DIngreso.vb como referencia"
```

### Regla 7: Solicita el checklist al final
```bash
"Al terminar, verifica el checklist del archivo hooks/guardrails.md"
```

### Regla 8: Actualiza CLAUDE.md después de completar cada módulo
```bash
"Actualiza CLAUDE.md marcando como ✅ los ítems de Venta que ya están completos"
```

### Regla 9: Para el proyecto integrador, pide el código con comentarios académicos
```bash
"Agrega comentarios explicativos al SP que describan el propósito del cursor,
 las estructuras de control y el manejo de excepciones — esto va en el documento"
```

### Regla 10: Sesiones de trabajo recomendadas
```
Sesión 1: Módulo Ventas — BD (SP CRUD + triggers TRG02 y TRG03)
Sesión 2: Módulo Ventas — Backend (EVenta, EDetalleVenta, DVenta, NVenta)
Sesión 3: Módulo Ventas — UI (FrmVenta — formulario maestro/detalle)
Sesión 4: Vistas para el integrador (VW01, VW02, VW03)
Sesión 5: SP con cursor (SP01: ReporteVentas, SP02: Inventario)
Sesión 6: Reportes RDLC
Sesión 7: Consultas + mejoras UI
Sesión 8: Pruebas + documentación del integrador
```

---

## PROMPTS PLANTILLA PARA SESIONES COMUNES

### Para generar un SP:
```
Lee CLAUDE.md y skills/sql-server-profesional.md.
Crea el Stored Procedure [nombre] en Database/03_StoredProcedures.sql.
Propósito: [descripción]
Parámetros de entrada: [lista]
Parámetros de salida: [lista]
Tablas involucradas: [lista]
Usar la plantilla con BEGIN TRY/CATCH y manejo de transacción.
```

### Para generar una clase DAL:
```
Lee CLAUDE.md y skills/vbnet-arquitectura.md.
Crea el archivo Sistema.Datos/D[Nombre].vb.
Debe incluir los métodos: Insertar, Actualizar, Eliminar, Listar, Buscar.
El método Insertar debe manejar maestro/detalle con transacción explícita.
Referencia la entidad E[Nombre] de Sistema.Entidades.
```

### Para generar una vista:
```
Lee CLAUDE.md y skills/sql-server-profesional.md.
Crea la vista [nombre] que cumpla los requisitos del proyecto integrador:
- JOIN entre [tabla1] y [tabla2] (mínimo)
- Función de agregación: [SUM/COUNT/AVG]
- GROUP BY por: [campos]
- HAVING: [condición]
Agrégala al archivo Database/02_Vistas.sql.
```

---

## CÓMO MEDIR TU PROGRESO

Cada semana, pide esto a Claude Code:
```
"Lee CLAUDE.md y dime: 
 1. ¿Qué porcentaje del proyecto está completo?
 2. ¿Qué módulo debo atacar hoy según las prioridades?
 3. ¿Cuántos de los 9 objetos de BD del integrador están listos?"
```

---

## ERRORES COMUNES A EVITAR

| Error | Consecuencia | Prevención |
|-------|-------------|------------|
| No leer CLAUDE.md | Código inconsistente | Siempre pedir que lo lea primero |
| Pedir todo a la vez | Código sin revisar | Sesiones enfocadas de 1-2 horas |
| No actualizar CLAUDE.md | Perder el estado | Actualizar al final de cada sesión |
| SQL inline en VB.NET | Arquitectura rota | Hook 2 del guardrails.md |
| Trigger que asume 1 fila | Bug en inserciones masivas | Hook de triggers |
| Cursor sin DEALLOCATE en CATCH | Memory leak en BD | Template del skill |

---

## PARA EL PROYECTO INTEGRADOR — CRONOGRAMA SUGERIDO

```
Semana 1: BD completa (tablas + SP + triggers + vistas)
           → Documento: sección 3 (diseño ER) + sección 4 (codificación)
           
Semana 2: Módulo Ventas + Reportes
           → Prueba desde la aplicación cada objeto de BD
           
Semana 3: Consultas + Mejoras UI + Normalización
           → Documento: sección 5 (pruebas + 3FN)
           
Semana 4: Integración final + Revisión + Preparar exposición
           → Documento: secciones 1, 2, 6, 7
```

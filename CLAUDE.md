# CLAUDE.md — Memoria del Agente
# Sistema Punto de Venta (POS) — VB.NET + SQL Server

> Este archivo es la **capa de memoria** del agente. Debe leerse PRIMERO en cada sesión
> antes de ejecutar cualquier tarea. Define el contexto completo, las reglas de trabajo,
> el estado actual y las convenciones del proyecto.

---

## 1. IDENTIDAD DEL PROYECTO

| Campo              | Valor                                              |
|--------------------|----------------------------------------------------|
| Nombre             | Sistema Punto de Venta                             |
| Tecnología UI      | VB.NET — Windows Forms (.NET Framework)            |
| Base de Datos      | SQL Server (T-SQL)                                 |
| IDE                | Visual Studio 2022                                 |
| Arquitectura       | N-Tier (4 capas)                                   |
| Estado             | ~80% completado                                    |
| Propósito dual     | Uso real en negocio + Proyecto integrador académico|

---

## 2. ARQUITECTURA DE LA SOLUCIÓN

```
Sistema.sln
├── Sistema.Entidades        ← Clases POCO / Entidades de dominio
├── Sistema.Datos            ← Acceso a datos (ADO.NET + Stored Procedures)
├── Sistema.Negocio          ← Lógica de negocio / validaciones
├── Sistema.Presentacion     ← WinForms UI (formularios, reportes RDLC)
└── Sistema.Docs             ← Scripts SQL y Documentación (NO compila)
```

### Flujo de dependencias (NUNCA romper esto):
```
Presentacion → Negocio → Datos → Entidades
     ↓              ↓        ↓
  (Forms)      (BL clases) (DAL clases)
```

### Regla crítica de arquitectura:
- La capa **Presentacion** NUNCA llama directamente a **Datos**.
- La capa **Datos** NUNCA contiene lógica de negocio.
- Toda operación de BD se hace mediante **Stored Procedures** — cero SQL inline en capas superiores.

---

## 3. MÓDULOS DEL SISTEMA

### Módulo Almacén
- **Categorías**: CRUD completo
- **Artículos**: CRUD con código de barras, stock, precio, imagen, categoría

### Módulo Compras
- **Proveedores**: CRUD (razón social, tipo doc, número doc, dirección, email, teléfono)
- **Ingresos**: Maestro/Detalle — tipo comprobante, proveedor, impuesto, detalle por artículo
  - Trigger: actualizar stock al insertar detalle
  - Funcionalidad: Anular ingreso + restaurar stock
  - Exportar PDF

### Módulo Ventas
- **Clientes**: CRUD (razón social, tipo doc, número doc, dirección, email, teléfono)
- **Ventas**: Maestro/Detalle — tipo comprobante, cliente, impuesto, detalle por artículo
  - Trigger: actualizar stock al insertar detalle
  - Funcionalidad: Anular venta + restaurar stock
  - Exportar PDF

### Módulo Acceso
- **Roles**: Administrador, Vendedor, Almacenero
- **Usuarios**: CRUD con login y password encriptado (SHA256 o similar)

### Módulo Consultas
- Consulta de ventas entre dos fechas

### Módulo Reportes
- RDLC: Artículos, Comprobante de venta, Compras, Ventas
- Exportar a PDF, Word, Excel

---

## 4. ESTADO ACTUAL — PENDIENTES

### En progreso (fase actual):
```
[COMPRAS]
✅ Entidad Ingreso + DetalleIngreso
✅ Stored Procedures CRUD básicos
⏳ Trigger actualizar stock (INSERT en DetalleIngreso)
⏳ Mostrar/Anular Ingreso
⏳ Restaurar stock al anular

[VENTAS]
✅ Entidad Venta + DetalleVenta (EVenta.vb, EDetalleVenta.vb)
✅ SP CRUD Ventas (03_StoredProcedures.sql — 6 SP)
✅ Triggers TRG02 + TRG03 (04_Triggers.sql)
✅ DAL DVenta.vb (Listar, BuscarPorFechas, ObtenerConDetalle, Anular, Actualizar, InsertarConDetalle)
✅ BL NVenta.vb (validaciones stock, cliente, detalle vacío, calcular totales)
⏳ Form: listado, búsqueda, selección clientes
⏳ Form: agregar artículos, validar stock, calcular totales
⏳ Insertar venta (UI)
⏳ Mostrar/Anular venta (UI)

[REPORTES]
⏳ Extensión RDLC en VS
⏳ Control ReportViewer
⏳ Reporte artículos
⏳ Reporte comprobante venta

[CONSULTAS]
⏳ Consulta ventas entre fechas

[MEJORAS UI]
⏳ Barra de herramientas formulario padre

[INTEGRADOR — OBJETOS BD COMPLEJOS] (Sesión 1 — 2026-05-16)
✅ SP01 — sp_ReporteVentasPorPeriodo (cursor + IF/ELSE + GROUP BY)
✅ SP02 — sp_InventarioValorizado    (cursor + IF + SUM OVER PARTITION)
✅ SP03 — sp_VentasConAnalisis       (JOIN múltiple + subconsulta correlacionada + GROUP BY + HAVING)
✅ VW01 — vw_VentasDetalladas        (CTE + JOIN 5 tablas + COUNT/SUM + HAVING)
✅ VW02 — vw_ComprasDetalladas       (CTE + JOIN 5 tablas + COUNT/SUM + HAVING)
✅ VW03 — vw_StockValorizado         (JOIN + COUNT/SUM/AVG + GROUP BY + HAVING)
✅ TRG04 — trg_Venta_AuditoriaDelete (AFTER DELETE + tabla auditoria_venta + 3FN)
Implementación integrador: 55/55 pts ✅✅✅

[IMPLEMENTACIÓN FINAL]
❌ Backup BD
❌ Setup/Instalador
```

---

## 5. OBJETOS DE BASE DE DATOS REQUERIDOS (Proyecto Integrador)

### Vistas — 3 mínimo (todas con JOIN + agregación + GROUP BY):
| ID    | Estado | Nombre                    | Archivo         | Propósito                                          |
|-------|--------|---------------------------|-----------------|----------------------------------------------------|
| VW01  | ✅     | vw_VentasDetalladas       | 02_Vistas.sql   | Ventas + cliente + artículos + totales por periodo |
| VW02  | ✅     | vw_ComprasDetalladas      | 02_Vistas.sql   | Ingresos + proveedor + artículos + costos          |
| VW03  | ✅     | vw_StockValorizado        | 02_Vistas.sql   | Stock actual + categoría + valor en almacén        |

### Triggers — 4 (INSERT x2, UPDATE, DELETE):
| ID    | Estado | Nombre                          | Tabla           | Evento        | Tipo                |
|-------|--------|---------------------------------|-----------------|---------------|---------------------|
| TRG01 | ⏳     | trg_Ingreso_ActualizarStock     | DetalleIngreso  | AFTER INSERT  | Stock ↑             |
| TRG02 | ✅     | trg_Venta_DescontarStock        | DetalleVenta    | AFTER INSERT  | Stock ↓             |
| TRG03 | ✅     | trg_Venta_RestaurarStock        | Venta           | AFTER UPDATE  | Anulación → Stock ↑ |
| TRG04 | ✅     | trg_Venta_AuditoriaDelete       | Venta           | AFTER DELETE  | Auditoría DELETE    |

### Procedimientos Almacenados — 4 totales (2 con cursor + 2 estándar):
| ID    | Estado | Nombre                          | Cursor | Propósito                                       |
|-------|--------|---------------------------------|--------|-------------------------------------------------|
| SP01  | ✅     | sp_ReporteVentasPorPeriodo      | ✅ Sí  | Itera ventas en rango, clasifica por monto       |
| SP02  | ✅     | sp_InventarioValorizado         | ✅ Sí  | Itera artículos, alerta stock, valor almacén     |
| SP03  | ✅     | sp_VentasConAnalisis            | ❌ No  | JOIN múltiple + subconsulta correlacionada       |
| SP04  | ✅     | sp_ObtenerVentaConDetalle       | ❌ No  | Retorna cabecera + detalle de una venta (app)    |

---

## 6. TABLAS DE LA BASE DE DATOS

### Tablas principales (mínimo 4 requeridas — tenemos 12):
```sql
TipoDocumento    -- Catálogo: DNI, RUC, Pasaporte, etc.
TipoComprobante  -- Catálogo: Boleta, Factura, Ticket
Rol              -- Administrador, Vendedor, Almacenero
Categoria        -- Clasificación de artículos
Usuario          -- Acceso al sistema (login + password SHA256)
Proveedor        -- Datos del proveedor
Cliente          -- Datos del cliente
Articulo         -- Inventario (FK: Categoria)
Ingreso          -- Cabecera compra (FK: Proveedor, TipoComprobante)
DetalleIngreso   -- Líneas de compra (FK: Ingreso, Articulo)
Venta            -- Cabecera venta (FK: Cliente, TipoComprobante, Usuario)
DetalleVenta     -- Líneas de venta (FK: Venta, Articulo)
auditoria_venta  -- Registro de eliminaciones de Venta (FK: Usuario) — creada en Sesión 1
```

---

## 7. CONVENCIONES DE NOMENCLATURA

### SQL Server:
```
Tablas:            PascalCase singular        → Articulo, DetalleVenta
Stored Procedures: sp_ + Entidad + Acción     → sp_ArticuloInsertar
Triggers:          trg_ + Tabla + Evento      → trg_Venta_DescontarStock
Vistas:            vw_ + Nombre               → vw_VentasDetalladas
Columnas PK:       id + Tabla                 → idArticulo
Columnas FK:       id + TablaReferenciada     → idCategoria
Columnas estado:   estado (BIT o VARCHAR)
```

### VB.NET:
```
Clases Entidad:    E + Nombre                 → EArticulo
Clases DAL:        D + Nombre                 → DArticulo
Clases BL:         N + Nombre                 → NArticulo (Negocio)
Formularios:       Frm + Nombre               → FrmArticulo
```

### Parámetros SP:
```
Todos los parámetros llevan @ y coinciden exactamente con el nombre de la columna
Ejemplo: @idArticulo, @nombre, @stock, @precio
```

---

## 8. PATRONES DE CÓDIGO OBLIGATORIOS

### DAL — Patrón estándar de conexión:
```vb
Public Function NombreMetodo(param As Tipo) As TipoRetorno
    Dim objConexion As New Conexion()
    Dim cmd As New SqlCommand()
    Try
        cmd.Connection = objConexion.AbrirConexion()
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "sp_NombreProcedimiento"
        cmd.Parameters.AddWithValue("@param", param)
        ' ... ejecutar ...
    Catch ex As Exception
        Throw ex
    Finally
        objConexion.CerrarConexion()
    End Try
End Function
```

### SP — Plantilla estándar con manejo de errores:
```sql
CREATE OR ALTER PROCEDURE sp_NombreProcedimiento
    @param1 TIPO,
    @param2 TIPO
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        -- lógica aquí
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
```

---

## 9. REGLAS DEL AGENTE

1. **Siempre leer este archivo primero** antes de cualquier tarea.
2. **Consultar el skill relevante** antes de generar código (ver `/skills/`).
3. **No romper la arquitectura N-Tier** — respetar el flujo de dependencias.
4. **Todo SQL va en SQL Server** — nunca SQL inline en VB.NET.
5. **Nombrar archivos según convención** — no inventar nombres propios.
6. **Completar un módulo antes de pasar al siguiente** — no dejar código huérfano.
7. **Documentar cada objeto de BD** — comentarios en SP, triggers y vistas.
8. **Validar siempre en dos capas**: Negocio (BL) y Base de Datos (constraints/triggers).
9. **Referenciar el estado en CLAUDE.md** — actualizar ✅ cuando un ítem se complete.
10. **Preguntar antes de refactorizar** código existente que funciona.

---

## 10. CONTEXTO ACADÉMICO (Proyecto Integrador)

Materia: Bases de Datos y Lenguajes
Evaluación del integrador:
- Introducción [2 pts]
- Especificación de requerimientos [3 pts]
- Especificación de diseños — ER + modelos [10 pts]
- Implementación: 2 SP con cursor + 1 SP + 3 triggers + 3 vistas [55 pts]
- Pruebas de software + Normalización hasta 3FN [25 pts]
- Conclusiones [3 pts] + Referencias [2 pts]
- × 70% documento + 30% exposición oral

**Restricción importante**: Los objetos de BD (SP, triggers, vistas) se ejecutan
DESDE la aplicación, NO directamente desde el gestor de BD.

---

## 11. ESTRUCTURA DE CARPETAS (Actualizado 2026-05-15)

```
Sistema/
├── Sistema.sln                          ← Solución principal
├── Sistema.Entidades/                   ← Clases de entidades
├── Sistema.Datos/                       ← Capa de datos (DAL)
├── Sistema.Negocio/                     ← Capa de negocio (BL)
├── Sistema.Presentacion/                ← Interfaz gráfica (UI/WinForms)
├── Sistema.Docs/                        ← Proyecto para documentación y scripts SQL (NO compila)
│   ├── DataBase/                        ← Scripts SQL de la base de datos
│   │   ├── 02_Vistas.sql               ← VW01, VW02, VW03 (Sesión 1)
│   │   ├── 03_StoredProcedures.sql     ← SP CRUD + SP01, SP02, SP03
│   │   ├── 04_Triggers.sql             ← TRG02, TRG03, TRG04 + tabla auditoria_venta
│   │   └── 06_PruebasSesion1.sql       ← Pruebas de objetos Sesión 1
│   └── Documentation/
│       ├── CLAUDE.md                    ← Copia dentro de la solución
│       ├── Especificacion_Requerimientos.md
│       ├── GUIA_CLAUDE_CODE.md
│       ├── README.md
│       └── Sesion1/
│           └── PROMPT_SESION1_CLAUDE_CODE.md
├── .gitignore
├── CLAUDE.md                            ← Original en la raíz (acceso rápido de Claude Code)
├── README.md
└── agent-dev-kit/                       ← Archivos del agente (ignorado en git)
```

**Nota importante**: Los archivos .md duplicados están por necesidad técnica:
- CLAUDE.md en raíz: lo encuentra Claude Code automáticamente
- CLAUDE.md en _Documentation/: está dentro de la solución como se requiere

*Última actualización: 2026-05-16 — Sesión 1: 7 objetos BD del integrador completados (3 vistas, 3 SP, 1 trigger DELETE)*

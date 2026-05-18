# CLAUDE.md — Memoria del Agente
# Sistema Punto de Venta (POS) — VB.NET + SQL Server

> Este archivo es la **capa de memoria** del agente. Debe leerse PRIMERO en cada sesión.
> Define el contexto completo, reglas de trabajo, estado actual y convenciones del proyecto.

---

## 1. IDENTIDAD DEL PROYECTO

| Campo          | Valor                                               |
|----------------|-----------------------------------------------------|
| Nombre         | Sistema Punto de Venta                              |
| Tecnología UI  | VB.NET — Windows Forms (.NET Framework 4.7.2)       |
| Base de Datos  | SQL Server (T-SQL) — dbsistema                      |
| IDE            | Visual Studio 2022                                  |
| Arquitectura   | N-Tier (4 capas)                                    |
| Propósito dual | Uso real en negocio + Proyecto integrador académico |

---

## 2. ARQUITECTURA DE LA SOLUCIÓN

```
Sistema.sln
├── Sistema.Entidades        ← Clases POCO / Entidades de dominio
├── Sistema.Datos            ← Acceso a datos (ADO.NET + Stored Procedures)
├── Sistema.Negocio          ← Lógica de negocio / validaciones
├── Sistema.Presentacion     ← WinForms UI (formularios)
└── Sistema.Docs             ← Scripts SQL y Documentación (NO compila)
    ├── DataBase/
    │   ├── 03_StoredProcedures.sql
    │   ├── 04_Triggers.sql
    │   └── [pendientes: 02_Vistas, 05_SP_Cursor, 06_TRG_Ingreso]
    └── Documentation/
        ├── CLAUDE.md
        ├── GUIA_SESION_VENTA.md
        └── Especificacion_Requerimientos.md
```

### Flujo de dependencias — NUNCA romper:
```
Presentacion → Negocio → Datos → Entidades
```

### Reglas críticas de arquitectura:
- Presentacion NUNCA llama directamente a Datos
- Datos NUNCA contiene lógica de negocio
- **CERO SQL inline en VB.NET** — solo `CommandType.StoredProcedure`
- Toda operación de BD va en un Stored Procedure

---

## 3. TABLAS DE LA BASE DE DATOS (dbsistema)

```
categoria        → idcategoria, nombre, descripcion, estado
articulo         → idarticulo, idcategoria(FK), codigo, nombre, precio_venta,
                   stock, imagen, descripcion, estado
persona          → idpersona, tipo_persona('Cliente'|'Proveedor'), nombre,
                   tipo_documento, num_documento, direccion, telefono, email
rol              → idrol, nombre, descripcion
usuario          → idusuario, idrol(FK), nombre, tipo_documento, num_documento,
                   direccion, telefono, email, clave, estado
ingreso          → idingreso, idproveedor(FK→persona.idpersona), idusuario(FK),
                   tipo_comprobante, serie_comprobante, num_comprobante, fecha,
                   impuesto, total, estado
detalle_ingreso  → iddetalle_ingreso, idingreso(FK), idarticulo(FK),
                   cantidad, precio
venta            → idventa, idcliente(FK→persona.idpersona), idusuario(FK),
                   tipo_comprobante, serie_comprobante, num_comprobante,
                   fecha, impuesto, total, estado
detalle_venta    → iddetalle_venta, idventa(FK), idarticulo(FK),
                   cantidad, precio, descuento, subtotal
```

> **CRÍTICO**: No existen tablas separadas de Cliente/Proveedor.
> La tabla `persona` usa `tipo_persona` para diferenciarlos.
> El tipo de comprobante es un VARCHAR directo (no tabla catálogo).
> La entidad VB.NET para venta es `Venta` — no existe `EVenta`.
> **IMPORTANTE**: detalle_ingreso NO tiene campo "importe" — se calcula como (precio * cantidad) en las vistas.

---

## 4. ESTADO ACTUAL DEL PROYECTO

### ✅ COMPLETADO:

```
Sistema.Entidades:
  ✅ Categoria.vb
  ✅ Articulo.vb
  ✅ Persona.vb
  ✅ Usuario.vb
  ✅ Ingreso.vb
  ✅ Venta.vb          (campos: idVenta, idCliente, idUsuario, tipo_comprobante,
  ✅ DetalleVenta.vb    serie_comprobante, num_comprobante, fecha, impuesto, total, estado)

Sistema.Datos:
  ✅ Conexion.vb
  ✅ DCategoria.vb
  ✅ DArticulo.vb
  ✅ DPersona.vb
  ✅ DRol.vb
  ✅ DUsuario.vb
  ✅ DIngreso.vb
  ✅ DVenta.vb

Sistema.Negocio:
  ✅ NCategoria.vb
  ✅ NArticulo.vb
  ✅ NPersona.vb
  ✅ NRol.vb
  ✅ NUsuario.vb
  ✅ NIngreso.vb
  ✅ NVenta.vb

Sistema.Presentacion:
  ✅ FrmLogin.vb
  ✅ FrmPrincipal.vb (MDIParent1 — MDI con menú y roles)
  ✅ FrmCategoria.vb
  ✅ FrmArticulo.vb
  ✅ FrmProveedores.vb
  ✅ FrmProveedor_Ingreso.vb
  ✅ FrmCliente.vb
  ✅ FrmRol.vb
  ✅ FrmUsuario.vb
  ✅ FrmIngreso.vb
  ✅ Variables.vb

Sistema.Docs/DataBase/:
  ✅ 03_StoredProcedures.sql (sp_VentaInsertar★+validaciones THROW 50001/50002,
                               sp_VentaActualizar, sp_VentaAnular+guard THROW 50004,
                               sp_VentaListar★, sp_DetalleVentaInsertar,
                               sp_ObtenerVentaConDetalle★, sp_VentaBuscarPorFechas★)
                               ★ = consulta compleja (JOIN + CASE/subconsulta)
  ✅ 04_Triggers.sql         (TRG02: trg_Venta_DescontarStock+guard stock insuf,
                               TRG03: trg_Venta_RestaurarStock+detección estado)
                               ⚠️ PENDIENTE: ejecutar en SSMS
```

### ❌ PENDIENTE — Lo que falta construir:

```
[CORRECCIONES A CÓDIGO EXISTENTE]
✅  DVenta.vb → BuscarPorFechas corregido → usa sp_VentaBuscarPorFechas (StoredProcedure)
✅  NVenta.vb → referencias a EVenta corregidas → usa clase Venta (entidad correcta)
✅  FrmIngreso.vb:263 → BUG-01 corregido → MsgBox error solo cuando Neg.Insertar=False
✅  sp_VentaAnular → BUG-03 corregido → THROW 50004 si venta ya anulada/inexistente
✅  sp_VentaInsertar → BUG-04/05 corregidos → THROW 50001 (sin cliente) / 50002 (total=0)
⚠️  vw_StockValorizado + sp_InventarioValorizado → verificar tipo articulo.estado (BIT vs VARCHAR)
⚠️  ingreso_insertar (SP original) → verificar compatibilidad con SqlDbType.Structured en DIngreso

[MÓDULO VENTAS — UI]
✅ FrmVenta.vb + FrmVenta.Designer.vb
   → Listado de ventas activas (sp_VentaListar con columnas: cliente, comprobante, total, clasificacion)
   → Pestaña Nueva Venta: selección de cliente, tipo comprobante, serie, número
   → Popup FrmCliente_Venta (patrón FrmProveedor_Ingreso, filtra tipo_persona='Cliente')
   → Búsqueda de artículos por código (Enter) y por panel flotante
   → DgvDetalle con columnas cantidad (editable) y descuento (editable), subtotal calculado
   → CalcularTotales: subtotal = (precio-descuento)*cantidad, IGV 16%, total
   → BtnAnular con confirmación — TRG03 restaura stock automáticamente
   → Variables.IdCliente y Variables.NombreCliente agregadas a Variables.vb
   → Handler VentasToolStripMenuItem1_Click agregado a FrmPrincipal.vb

[OBJETOS BD — PROYECTO INTEGRADOR]
✅ 02_Vistas.sql — CORREGIDO para respetar schema real
   → vw_VentasDetalladas  (5 JOINs, subqueries escalares COUNT+SUM, 14 columnas)
   → vw_ComprasDetalladas (5 JOINs, usa idproveedor, calcula importe=(precio*cantidad), 15 columnas)
   → vw_StockValorizado   (2 JOINs, CASE 4 niveles, subqueries SUM/COUNT/AVG)

✅ 05_SP_Cursor.sql
   → sp_ReporteVentasPorPeriodo (CURSOR+WHILE+IF/ELSE+tabla temporal, 2 resultsets)
   → sp_InventarioValorizado    (CURSOR+WHILE+IF+CASE 4 niveles, @stockMinimo=5)

✅ 06_TRG_Ingreso.sql
   → trg_Ingreso_ActualizarStock (AFTER INSERT detalle_ingreso, UPDATE con JOIN inserted)

✅ sp_VentaBuscarPorFechas → incluido en 03_StoredProcedures.sql

[MÓDULO CONSULTAS]
✅ FrmConsultaVentas.vb — DateTimePicker → NVenta.ConsultarDetallado (vw_VentasDetalladas) → 22 cols, suma subtotales
✅ FrmConsultaCompras.vb + Designer — DateTimePicker → NIngreso.ConsultarDetallado (vw_ComprasDetalladas) → 20 cols, suma importes
✅ FrmStockValorizado.vb + Designer — BtnActualizar → NArticulo.ConsultarStockValorizado (vw_StockValorizado) → 13 cols
✅ sp_ConsultaVentasDetalladas/ComprasDetalladas/StockValorizado → 03_StoredProcedures.sql
✅ DAL: DVenta/DIngreso.ConsultarDetallado, DArticulo.ConsultarStockValorizado
✅ BL: NVenta/NIngreso.ConsultarDetallado, NArticulo.ConsultarStockValorizado
✅ FrmPrincipal: menu Consultas ampliado con Compras y Stock Valorizado

[REPORTES]
❌ RDLC: comprobante de venta, reporte artículos, reporte ventas

[IMPLEMENTACIÓN FINAL]
❌ Backup BD (.bak)
❌ Setup/Instalador
```

---

## 5. OBJETOS BD REQUERIDOS — PROYECTO INTEGRADOR

### Vistas (3 — JOIN + agregación + GROUP BY):
| ID   | Nombre                 | Tablas                                         | Estado |
|------|------------------------|------------------------------------------------|--------|
| VW01 | vw_VentasDetalladas    | venta + persona + detalle_venta + articulo     | ✅     |
| VW02 | vw_ComprasDetalladas   | ingreso + persona + detalle_ingreso + articulo | ✅     |
| VW03 | vw_StockValorizado     | articulo + categoria                           | ✅     |

### Triggers (3 mínimo):
| ID    | Nombre                       | Tabla          | Evento       | Estado |
|-------|------------------------------|----------------|--------------|--------|
| TRG01 | trg_Ingreso_ActualizarStock  | detalle_ingreso | AFTER INSERT | ✅    |
| TRG02 | trg_Venta_DescontarStock     | detalle_venta  | AFTER INSERT | ✅    |
| TRG03 | trg_Venta_RestaurarStock     | venta          | AFTER UPDATE | ✅    |

### Stored Procedures con cursor (2 mínimo):
| ID   | Nombre                      | Cursor | Estado |
|------|-----------------------------|--------|--------|
| SP01 | sp_ReporteVentasPorPeriodo  | ✅ Sí  | ✅     |
| SP02 | sp_InventarioValorizado     | ✅ Sí  | ✅     |

---

## 6. CONVENCIONES DE NOMENCLATURA

### SQL Server (dbsistema):
```
Tablas:     snake_case singular          → articulo, detalle_venta, persona
SPs:        sp_ + Entidad + Accion       → sp_VentaInsertar, sp_VentaAnular
Triggers:   trg_ + Tabla + Evento       → trg_Venta_DescontarStock
Vistas:     vw_ + Nombre                → vw_VentasDetalladas
PKs:        id + tabla                  → idventa, idarticulo
FKs:        id + tablaReferenciada      → idcliente, idusuario
```

### VB.NET:
```
Entidades:  sin prefijo                  → Venta, Articulo, Persona
DAL:        D + Nombre                   → DVenta, DArticulo
BL:         N + Nombre                   → NVenta, NArticulo
Forms:      Frm + Nombre                 → FrmVenta, FrmArticulo
Globales:   clase Variables.vb           → Variables.IdUsuario, Variables.IdProveedor
```

---

## 7. PATRONES DE CÓDIGO OBLIGATORIOS

### DAL — Consulta (ExecuteReader):
```vb
Public Function NombreMetodo(Param As Tipo) As DataTable
    Try
        Dim Resultado As SqlDataReader
        Dim Tabla As New DataTable
        Dim Comando As New SqlCommand("sp_NombreProcedimiento", MyBase.conn)
        Comando.CommandType = CommandType.StoredProcedure
        Comando.Parameters.Add("@param", SqlDbType.Tipo).Value = Param
        MyBase.conn.Open()
        Resultado = Comando.ExecuteReader()
        Tabla.Load(Resultado)
        MyBase.conn.Close()
        Return Tabla
    Catch ex As Exception
        Throw ex
    End Try
End Function
```

### DAL — Escritura (ExecuteNonQuery):
```vb
Public Sub NombreMetodo(Obj As Entidad)
    Try
        Dim Comando As New SqlCommand("sp_NombreProcedimiento", MyBase.conn)
        Comando.CommandType = CommandType.StoredProcedure
        Comando.Parameters.Add("@campo", SqlDbType.Tipo).Value = Obj.Campo
        MyBase.conn.Open()
        Comando.ExecuteNonQuery()
        MyBase.conn.Close()
    Catch ex As Exception
        Throw ex
    End Try
End Sub
```

### DAL — Con OUTPUT (para obtener ID generado):
```vb
Dim ParamId As New SqlParameter("@idventa", SqlDbType.Int)
ParamId.Direction = ParameterDirection.Output
Comando.Parameters.Add(ParamId)
Comando.ExecuteNonQuery()
Dim NuevoId As Integer = Convert.ToInt32(ParamId.Value)
```

### DAL — Transacción explícita con detalle:
```vb
Public Sub InsertarConDetalle(Obj As Venta, Det As DataTable)
    Dim Trx As SqlTransaction = Nothing
    Try
        MyBase.conn.Open()
        Trx = MyBase.conn.BeginTransaction()
        ' 1. Insertar cabecera con SP + OUTPUT
        ' 2. Para cada fila de Det → sp_DetalleVentaInsertar
        Trx.Commit()
        MyBase.conn.Close()
    Catch ex As Exception
        If Trx IsNot Nothing Then Trx.Rollback()
        MyBase.conn.Close()
        Throw ex
    End Try
End Sub
```

### BL — Estándar:
```vb
Public Function NombreMetodo(Param As Tipo) As DataTable
    Try
        Dim Datos As New DNombreDAL
        Return Datos.NombreMetodo(Param)
    Catch ex As Exception
        MsgBox(ex.Message)
        Return Nothing
    End Try
End Function
```

### Form — Cargar listado:
```vb
Private Sub Listar()
    Try
        Dim Neg As New Negocio.NNombre
        DgvListado.DataSource = Neg.Listar()
        LblTotal.Text = "Total Registros: " & DgvListado.DataSource.Rows.Count.ToString()
        Me.Formato()
        Me.Limpiar()
    Catch ex As Exception
        MsgBox(ex.Message)
    End Try
End Sub
```

### SP — Plantilla estándar con TRY/CATCH:
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
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        DECLARE @msg NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg, 16, 1);
    END CATCH
END
GO
```

### SP con CURSOR — Plantilla:
```sql
DECLARE cur CURSOR FOR
    SELECT ... FROM ...;
OPEN cur;
FETCH NEXT FROM cur INTO @var1, @var2;
WHILE @@FETCH_STATUS = 0
BEGIN
    -- lógica por fila
    FETCH NEXT FROM cur INTO @var1, @var2;
END;
CLOSE cur;
DEALLOCATE cur;
```

---

## 8. CONTEXTO ACADÉMICO — PROYECTO INTEGRADOR

**Materia:** Bases de Datos y Lenguajes
**Restricción crítica:** Los objetos de BD (SP, triggers, vistas) se ejecutan DESDE
la aplicación VB.NET, NO directamente desde SSMS.

**Rubrica de evaluación:**
| Sección                                      | Puntos |
|----------------------------------------------|--------|
| Introducción                                 | 2      |
| Especificación de requerimientos             | 3      |
| Diseño ER + modelos                          | 10     |
| Implementación (SPs + triggers + vistas)     | 55     |
| Pruebas + Normalización 3FN                  | 25     |
| Conclusiones + Referencias                   | 5      |
| **Total**                                    | **100**|

Ponderación: 70% documento + 30% exposición oral

---

## 9. REGLAS DEL AGENTE

1. **Leer este archivo primero** antes de cualquier tarea
2. **No romper arquitectura** — flujo: Presentacion→Negocio→Datos→Entidades
3. **CERO SQL inline** — siempre `CommandType.StoredProcedure`
4. **Entidad correcta: `Venta`** — no existe `EVenta` en el proyecto
5. **`persona` para clientes Y proveedores** — no existen tablas Cliente/Proveedor
6. **Convenciones de nomenclatura** — snake_case BD, PascalCase VB
7. **Un módulo a la vez** — completar antes de pasar al siguiente
8. **Comentarios académicos** en objetos BD — propósito, estructuras de control usadas
9. **Preguntar antes de refactorizar** código existente que funciona
10. **Actualizar este CLAUDE.md** al completar cada ítem (cambiar ❌ por ✅)

---

*Última actualización: 2026-05-17 — Estado real documentado, listo para FrmVenta + Objetos BD integrador*

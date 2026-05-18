G# GUIA_SESION_VENTA.md
# Sesión de trabajo: FrmVenta + Objetos BD Integrador
# Sistema POS — VB.NET + SQL Server

> Lee CLAUDE.md primero. Esta guía complementa con instrucciones específicas
> para las tareas pendientes más importantes.

---

## ORDEN DE TRABAJO RECOMENDADO

```
Paso 1 → Corregir DVenta.vb y NVenta.vb (5 min)
Paso 2 → Agregar sp_VentaBuscarPorFechas al 03_StoredProcedures.sql
Paso 3 → Crear FrmVenta.vb (formulario maestro/detalle)
Paso 4 → Crear 02_Vistas.sql (VW01, VW02, VW03)
Paso 5 → Crear 05_SP_Cursor.sql (SP01, SP02)
Paso 6 → Crear 06_TRG_Ingreso.sql (TRG01)
Paso 7 → FrmConsultaVentas.vb
```

---

## PASO 1: CORRECCIONES URGENTES

### DVenta.vb — BuscarPorFechas (reemplazar SQL inline)

**Problema actual:**
```vb
' ❌ SQL inline — VIOLA la arquitectura
Dim Comando As New SqlCommand(
    "SELECT * FROM vw_VentasDetalladas WHERE fechaVenta BETWEEN @fechaInicio AND @fechaFin",
    MyBase.conn)
Comando.CommandType = CommandType.Text
```

**Corrección:**
```vb
' ✅ Usar SP
Public Function BuscarPorFechas(FechaInicio As Date, FechaFin As Date) As DataTable
    Try
        Dim Resultado As SqlDataReader
        Dim Tabla As New DataTable
        Dim Comando As New SqlCommand("sp_VentaBuscarPorFechas", MyBase.conn)
        Comando.CommandType = CommandType.StoredProcedure
        Comando.Parameters.Add("@fechaInicio", SqlDbType.DateTime).Value = FechaInicio
        Comando.Parameters.Add("@fechaFin", SqlDbType.DateTime).Value = FechaFin
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

### NVenta.vb — Referencia incorrecta a EVenta

**Buscar y reemplazar:** Toda referencia a `EVenta` → `Venta`
```vb
' ❌ Incorrecto
Public Function Insertar(Obj As EVenta, Det As DataTable) As Boolean

' ✅ Correcto
Public Function Insertar(Obj As Venta, Det As DataTable) As Boolean
```

---

## PASO 2: SP NUEVO — sp_VentaBuscarPorFechas

Agregar al archivo `Sistema.Docs/DataBase/03_StoredProcedures.sql`:

```sql
-- ------------------------------------------------------------
-- 7. sp_VentaBuscarPorFechas
--    Busca ventas en un rango de fechas con datos del cliente.
--    Usado por DVenta.BuscarPorFechas() para reemplazar SQL inline.
-- ------------------------------------------------------------
CREATE OR ALTER PROCEDURE sp_VentaBuscarPorFechas
    @fechaInicio DATETIME,
    @fechaFin    DATETIME
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT  V.idventa,
                P.nombre            AS cliente,
                V.tipo_comprobante,
                V.num_comprobante,
                V.fecha,
                V.impuesto,
                V.total,
                V.estado
        FROM    venta  V
        INNER JOIN persona P ON V.idcliente = P.idpersona
        WHERE   V.fecha BETWEEN @fechaInicio AND @fechaFin
          AND   V.estado = 'Activo'
        ORDER BY V.fecha DESC;
    END TRY
    BEGIN CATCH
        DECLARE @msg NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg, 16, 1);
    END CATCH
END
GO
```

---

## PASO 3: FrmVenta — Especificación completa

### Estructura visual (igual que FrmIngreso):
```
TabGeneral
├── TabPage1: "Listado"
│   ├── TxtValor + BtnBuscar
│   ├── DgvListado (ventas activas)
│   ├── ChkSeleccionar + BtnAnular
│   └── LblTotal
└── TabPage2: "Nueva Venta"
    ├── GroupBox1: "Cabecera"
    │   ├── TxtIdCliente (ReadOnly) + TxtNombreCliente (ReadOnly) + BtnBuscarCliente ("...")
    │   ├── CboTipoComprobante ("Factura"|"Boleta"|"Ticket")
    │   ├── TxtSerieComprobante + TxtNumComprobante
    │   └── TxtImpuesto (ReadOnly, valor "0.16")
    ├── GroupBox2: "Detalle"
    │   ├── TxtCodigo + BtnBuscarArticulos
    │   ├── PanelArticulos (igual que FrmIngreso, visible=False)
    │   ├── DgvDetalle (editable en cantidad y descuento)
    │   └── TxtSubTotal + TxtTotalImpuesto + TxtTotal (ReadOnly)
    └── BtnInsertar + BtnCancelar
```

### Columnas de DgvDetalle:
```
idarticulo  → oculta
codigo      → ReadOnly
articulo    → ReadOnly
cantidad    → editable
precio      → ReadOnly
descuento   → editable (default 0)
subtotal    → ReadOnly (calculado: (precio - descuento) * cantidad)
```

### Variables adicionales necesarias en Variables.vb:
```vb
Public Shared IdCliente As String
Public Shared NombreCliente As String
```

### Popup de selección de cliente:
Crear `FrmCliente_Venta.vb` con el mismo patrón que `FrmProveedor_Ingreso.vb`.
- Lista solo personas con `tipo_persona = 'Cliente'`
- Al hacer doble clic: `Variables.IdCliente` y `Variables.NombreCliente`

### Métodos clave de FrmVenta:
```vb
' Tabla en memoria para el detalle
Private DtDetalle As New DataTable

' Al cargar:
Private Sub CrearTablaDetalle()
    DtDetalle = New DataTable("Detalle")
    DtDetalle.Columns.Add("idarticulo", GetType(Integer))
    DtDetalle.Columns.Add("codigo", GetType(String))
    DtDetalle.Columns.Add("articulo", GetType(String))
    DtDetalle.Columns.Add("cantidad", GetType(Integer))
    DtDetalle.Columns.Add("precio", GetType(Decimal))
    DtDetalle.Columns.Add("descuento", GetType(Decimal))
    DtDetalle.Columns.Add("subtotal", GetType(Decimal))
    DgvDetalle.DataSource = DtDetalle
    ' ... ocultar idarticulo, ReadOnly en columnas fijas
End Sub

' Calcular totales igual que FrmIngreso:
Private Sub CalcularTotales()
    Dim Total As Decimal = 0
    For Each Fila As DataGridViewRow In DgvDetalle.Rows
        Total += CDec(Fila.Cells("subtotal").Value)
    Next
    Dim SubTotal As Decimal = Math.Round(Total / (1 + CDec(TxtImpuesto.Text)), 2)
    TxtTotal.Text = Total
    TxtSubTotal.Text = SubTotal
    TxtTotalImpuesto.Text = CStr(Total - SubTotal)
End Sub

' Al insertar — llamar a NVenta.Insertar(objVenta, DtDetalle)
' NVenta valida: cliente > 0, detalle no vacío, stock disponible
' DVenta.InsertarConDetalle: transacción explícita con sp_VentaInsertar + sp_DetalleVentaInsertar
' TRG02 (trg_Venta_DescontarStock) se activa automáticamente en cada detalle
```

---

## PASO 4: 02_Vistas.sql — Especificación

Archivo: `Sistema.Docs/DataBase/02_Vistas.sql`

### VW01 — vw_VentasDetalladas
```
Propósito: Reporte de ventas con detalle de artículos por cliente
JOINs: venta + persona(cliente) + detalle_venta + articulo + categoria
Columnas: fechaVenta, cliente, tipoComprobante, numComprobante,
          codigoArticulo, nombreArticulo, categoria, cantidad,
          precio, descuento, subtotal, totalVenta, estado
Agregación: COUNT(detalle), SUM(subtotal) — GROUP BY venta
HAVING: ventas con al menos 1 detalle
```

### VW02 — vw_ComprasDetalladas
```
Propósito: Reporte de ingresos (compras) con detalle por proveedor
JOINs: ingreso + persona(proveedor) + detalle_ingreso + articulo + categoria
Columnas: fecha, proveedor, tipoComprobante, numComprobante,
          codigoArticulo, nombreArticulo, cantidad, precio, importe, total
Agregación: COUNT(detalle), SUM(importe) — GROUP BY ingreso
HAVING: ingresos con al menos 1 detalle
```

### VW03 — vw_StockValorizado
```
Propósito: Inventario actual con valor económico por artículo y categoría
JOINs: articulo + categoria
Columnas: categoria, codigoArticulo, nombreArticulo,
          stockActual, precioVenta, valorTotal (stock * precio_venta),
          estadoStock (CASE: 'Crítico'|'Bajo'|'Normal')
Agregación: SUM(stock), SUM(stock*precio_venta) — GROUP BY categoria
HAVING: artículos activos (estado = 1)
```

---

## PASO 5: 05_SP_Cursor.sql — Especificación

Archivo: `Sistema.Docs/DataBase/05_SP_Cursor.sql`

### SP01 — sp_ReporteVentasPorPeriodo
```
Parámetros: @fechaInicio DATE, @fechaFin DATE
Propósito: Itera venta por venta en el período, clasifica por monto
Estructuras requeridas (académico):
  - CURSOR + WHILE + FETCH
  - IF/ELSE para clasificar: 'Alta' > 1000, 'Media' 500-1000, 'Baja' < 500
  - Variables acumuladoras: @totalPeriodo, @contadorVentas
  - TRY/CATCH con ROLLBACK
  - Tabla resultado: cliente, fecha, total, clasificacion, acumulado
Salida: SELECT con resultados del cursor (tabla temporal o resultado directo)
```

### SP02 — sp_InventarioValorizado
```
Parámetros: @stockMinimo INT (default 5)
Propósito: Itera artículo por artículo, detecta bajo stock, calcula valor
Estructuras requeridas (académico):
  - CURSOR + WHILE + FETCH
  - IF stock < @stockMinimo → marcar alerta
  - Variables: @valorTotal acumulado por categoría
  - CASE para estado: 'Crítico'(0-2), 'Bajo'(3-@min), 'Normal'
  - TRY/CATCH
  - Tabla resultado: categoria, articulo, stock, valorUnitario, valorTotal, alerta
```

---

## PASO 6: 06_TRG_Ingreso.sql — Especificación

Archivo: `Sistema.Docs/DataBase/06_TRG_Ingreso.sql`

### TRG01 — trg_Ingreso_ActualizarStock
```
Tabla: detalle_ingreso
Evento: AFTER INSERT
Propósito: Incrementa stock de articulo al registrar un ingreso de mercancía
Lógica:
  UPDATE articulo
  SET stock = stock + I.cantidad
  FROM articulo A
  INNER JOIN inserted I ON A.idarticulo = I.idarticulo
  -- Soporta inserciones en lote (no asumir 1 sola fila)
Con TRY/CATCH + ROLLBACK
```

---

## PROMPTS LISTOS PARA CLAUDE CODE

### Prompt 1 — Correcciones + SP nuevo:
```
Lee CLAUDE.md y GUIA_SESION_VENTA.md.

Tarea 1: En Sistema.Datos/DVenta.vb, reemplaza el método BuscarPorFechas
para que use CommandType.StoredProcedure con sp_VentaBuscarPorFechas
en lugar del SQL inline actual.

Tarea 2: En Sistema.Negocio/NVenta.vb, reemplaza toda referencia a EVenta
por la clase Venta (que es la entidad correcta del proyecto).

Tarea 3: Agrega sp_VentaBuscarPorFechas al archivo
Sistema.Docs/DataBase/03_StoredProcedures.sql siguiendo el patrón
de los SPs existentes con TRY/CATCH.
```

### Prompt 2 — FrmVenta:
```
Lee CLAUDE.md y GUIA_SESION_VENTA.md.

Crea FrmVenta.vb y FrmVenta.Designer.vb en Sistema.Presentacion/.
Sigue EXACTAMENTE el patrón de FrmIngreso.vb (que ya existe y funciona)
adaptado para ventas:
- Cabecera: cliente (con popup FrmCliente_Venta), tipo comprobante, serie, número
- Detalle: igual que FrmIngreso pero con columna descuento editable
- Cálculo: subtotal = (precio - descuento) * cantidad
- Anular: llama NVenta.Anular — TRG03 restaura stock automáticamente
- La clase de entidad es Venta (no EVenta)
- DtDetalle debe tener columnas: idarticulo, codigo, articulo, cantidad, precio, descuento, subtotal

También crea FrmCliente_Venta.vb con el mismo patrón que FrmProveedor_Ingreso.vb.

Agrega FrmVenta al menú Ventas en FrmPrincipal.vb (VentasToolStripMenuItem1_Click).
```

### Prompt 3 — Vistas:
```
Lee CLAUDE.md y GUIA_SESION_VENTA.md.

Crea el archivo Sistema.Docs/DataBase/02_Vistas.sql con:
- VW01: vw_VentasDetalladas
- VW02: vw_ComprasDetalladas
- VW03: vw_StockValorizado

Sigue la especificación exacta de GUIA_SESION_VENTA.md (sección PASO 4).
Cada vista debe tener:
- Mínimo 2 JOINs
- Al menos una función de agregación (SUM, COUNT, AVG)
- GROUP BY
- HAVING
- Comentarios académicos explicando cada elemento
```

### Prompt 4 — SPs con cursor:
```
Lee CLAUDE.md y GUIA_SESION_VENTA.md.

Crea el archivo Sistema.Docs/DataBase/05_SP_Cursor.sql con:
- SP01: sp_ReporteVentasPorPeriodo
- SP02: sp_InventarioValorizado

Sigue la especificación de GUIA_SESION_VENTA.md (sección PASO 5).
CADA SP debe incluir comentarios académicos que expliquen:
  -- CURSOR: propósito y por qué se itera registro a registro
  -- WHILE: condición de continuación @@FETCH_STATUS = 0
  -- IF/ELSE: criterio de clasificación
  -- TRY/CATCH: manejo de errores y ROLLBACK
  -- DEALLOCATE: liberación de recursos
```

### Prompt 5 — Trigger de ingreso:
```
Lee CLAUDE.md y GUIA_SESION_VENTA.md.

Crea el archivo Sistema.Docs/DataBase/06_TRG_Ingreso.sql con:
- TRG01: trg_Ingreso_ActualizarStock (AFTER INSERT en detalle_ingreso)

Sigue el patrón EXACTO de trg_Venta_DescontarStock en 04_Triggers.sql
(que ya existe y funciona), pero para incrementar stock al comprar.
Incluye comentarios académicos explicando el uso de la tabla 'inserted'.
```

---

## NOTAS IMPORTANTES

**Sobre la conexión a BD:**
- Conexion.vb usa `Me.Seguridad = True` por defecto → Integrated Security (Windows)
- El servidor es `DESKTOP-CPRU6SB\SQLEXPRESS`, base de datos `dbsistema`
- DVenta ya hereda de Conexion → usa `MyBase.conn` directamente

**Sobre TRG02 y TRG03:**
- TRG02 (DescontarStock) se activa automáticamente cuando DVenta llama a sp_DetalleVentaInsertar
- TRG03 (RestaurarStock) se activa cuando DVenta llama a sp_VentaAnular
- FrmVenta NO necesita manejar el stock manualmente — los triggers lo hacen
- Si hay error de stock, TRG02 hace ROLLBACK y VB captura la excepción

**Sobre validaciones en NVenta:**
- Validar IdCliente > 0 antes de insertar
- Validar DtDetalle.Rows.Count > 0
- La validación de stock puede quedarse en NVenta O dejar que TRG02 la rechace
- Usar la validación existente en NVenta.Insertar como referencia

**Sobre el menú principal:**
- FrmVenta debe abrirse desde `VentasToolStripMenuItem1_Click` en FrmPrincipal.vb
- FrmPrincipal ya tiene el item de menú — solo falta el handler

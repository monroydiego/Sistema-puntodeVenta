# Flujo de Funcionamiento — Sistema Punto de Venta

## Índice
1. [Arquitectura General](#arquitectura-general)
2. [Ciclo de Vida de una Venta](#ciclo-de-vida-de-una-venta)
3. [Gestión de Stock](#gestión-de-stock)
4. [Objetos de BD (Vistas, Triggers, SPs)](#objetos-de-bd)
5. [Módulos Principales](#módulos-principales)
6. [Diagrama de Flujo](#diagrama-de-flujo)

---

## Arquitectura General

```
┌─────────────────────────────────────────────────────────────┐
│                   USUARIO (WinForms UI)                     │
│          (FrmLogin → FrmPrincipal → Formularios)            │
└──────────────────────┬──────────────────────────────────────┘
                       │ Interacción
┌──────────────────────▼──────────────────────────────────────┐
│               CAPA DE PRESENTACIÓN                          │
│  (FrmVenta, FrmIngreso, FrmConsultaVentas, etc.)            │
│  • Validación de datos en UI                               │
│  • Cálculos de totales e impuestos                          │
│  • Gestión de tabla detalle en memoria (DataTable)          │
└──────────────────────┬──────────────────────────────────────┘
                       │ Llamadas a NombreClase
┌──────────────────────▼──────────────────────────────────────┐
│              CAPA DE NEGOCIO (BL)                           │
│  (NVenta, NArticulo, NPersona, etc.)                        │
│  • Lógica de validación y reglas de negocio                 │
│  • Orquestación de operaciones                              │
│  • Manejo de excepciones                                    │
└──────────────────────┬──────────────────────────────────────┘
                       │ Llamadas a DNombreClase
┌──────────────────────▼──────────────────────────────────────┐
│            CAPA DE DATOS (DAL/ADO.NET)                      │
│  (DVenta, DArticulo, DPersona, etc.)                        │
│  • Comunicación con SQL Server                              │
│  • Ejecución de StoredProcedures                            │
│  • Manejo de transacciones                                  │
└──────────────────────┬──────────────────────────────────────┘
                       │ CommandType.StoredProcedure
┌──────────────────────▼──────────────────────────────────────┐
│           BASE DE DATOS SQL Server (dbsistema)              │
│  • 9 Tablas (persona, articulo, venta, etc.)                │
│  • 7 StoredProcedures (sp_VentaInsertar, etc.)              │
│  • 3 Vistas (vw_VentasDetalladas, etc.)                     │
│  • 3 Triggers (trg_Venta_DescontarStock, etc.)              │
└─────────────────────────────────────────────────────────────┘
```

### Principio Fundamental
```
Presentacion → Negocio → Datos → Entidades
```
- **NUNCA**: Presentación llama directamente a Datos
- **NUNCA**: SQL inline en VB.NET — solo StoredProcedures
- **SIEMPRE**: Flujo de dependencias respetado

---

## Ciclo de Vida de una Venta

### Paso 1: Seleccionar Cliente
```
Usuario abre FrmVenta → TabPage2 "Nueva Venta"
    ↓
Hace clic BtnBuscarCliente ("..." botón)
    ↓
FrmCliente_Venta.ShowDialog()
    ↓
NPersona.ListarClientes() → sp_PersonaListar (tipo_persona='Cliente')
    ↓
DataTable con clientes disponibles
    ↓
Usuario selecciona cliente (double-click)
    ↓
Variables.IdCliente = idpersona
Variables.NombreCliente = nombre
    ↓
FrmCliente_Venta se cierra
    ↓
BtnBuscarCliente_Click asigna:
    TxtIdCliente.Text = Variables.IdCliente
    TxtNombreCliente.Text = Variables.NombreCliente
    DtDetalle.Clear()  ← IMPORTANTE: limpia detalle anterior
```

### Paso 2: Agregar Artículos al Detalle
**Dos opciones:**

#### Opción A: Por código (Enter)
```
Usuario escribe código en TxtCodigo
    ↓
Presiona Enter (KeyDown event)
    ↓
TxtCodigo_KeyDown:
  - Valida: ¿hay cliente? Si no → MsgBox y return
  - NArticulo.BuscarCodigo(codigo)
    ↓
    sp_ArticuloBuscarCodigo → SELECT * FROM articulo WHERE codigo=@codigo
    ↓
    Retorna objeto Articulo (idarticulo, nombre, precio_venta, stock, etc.)
    ↓
AgregarDetalle(Obj):
  - Valida: ¿cliente seleccionado? Si no → return
  - Valida: ¿artículo ya agregado? Si sí → "Artículo duplicado"
  - Crea fila en DtDetalle:
    * idarticulo = Obj.IdArticulo
    * codigo = Obj.Codigo
    * articulo = Obj.Nombre
    * cantidad = 1 (por defecto)
    * precio = Obj.PrecioVenta
    * descuento = 0
    * subtotal = Obj.PrecioVenta
  - DgvDetalle.Refresh()
  - CalcularTotales()
    ↓
TxtCodigo.Clear()
```

#### Opción B: Por búsqueda flotante
```
Usuario hace clic BtnBuscarArticulos
    ↓
BtnBuscarArticulos_Click:
  - Valida: ¿hay cliente? Si no → MsgBox y return
  - PanelArticulos.Visible = True
    ↓
Usuario escribe en TxtBuscarArticulos (búsqueda libre)
    ↓
Clic BtnBuscarArticulosDetalle
    ↓
NArticulo.Buscar(criterio)
    ↓
sp_ArticuloBuscar → SELECT * FROM articulo WHERE ... LIKE '%criterio%'
    ↓
DataTable con artículos coincidentes se muestra en DgvArticulos
    ↓
Usuario hace double-click en artículo
    ↓
DgvArticulos_CellDoubleClick:
  - Obtiene datos del artículo seleccionado
  - Crea objeto Articulo
  - AgregarDetalle(Obj) ← mismo flujo que Opción A
  - PanelArticulos.Visible = False ← cierra panel
```

### Paso 3: Editar Cantidad y Descuento
```
Usuario hace clic en celda DgvDetalle (cantidad o descuento)
    ↓
Edita valor (ej: cantidad = 5, descuento = 10)
    ↓
Presiona Tab o Enter (CellEndEdit event)
    ↓
DgvDetalle_CellEndEdit:
  - Validaciones:
    * Cantidad ≥ 1 (si no, asigna 1)
    * Descuento ≥ 0 y ≤ precio (si no, asigna 0)
  - Cálculo correcto del subtotal:
    subtotal = (precio - descuento) × cantidad
  - Fila.Cells("subtotal").Value = subtotal calculado
    ↓
CalcularTotales():
  - Itera todas las filas del DgvDetalle
  - Total = suma de todos los subtotales
  - SubTotal = Total / (1 + 0.16) ← sin impuesto
  - TotalImpuesto = Total - SubTotal
  - Asigna valores a TxtSubTotal, TxtTotalImpuesto, TxtTotal
```

### Paso 4: Registrar Venta
```
Usuario completa datos:
  ✓ Cliente seleccionado
  ✓ Tipo comprobante (Factura/Boleta/Ticket)
  ✓ Número comprobante (ej: 001-2026-0001)
  ✓ Al menos 1 artículo en detalle
    ↓
Usuario hace clic BtnInsertar
    ↓
BtnInsertar_Click:
  - Validaciones finales:
    * Cliente != ""
    * Número comprobante != ""
    * DtDetalle.Rows.Count >= 1
  - Crea objeto Venta:
    * IdCliente = TxtIdCliente
    * IdUsuario = Variables.IdUsuario (del login)
    * IdTipoComprobante = CboTipoComprobante.SelectedIndex + 1
    * NumComprobante = TxtNumComprobante
    * FechaVenta = Date.Now
    * Impuesto = 0.16
    * TotalVenta = TxtTotal
    ↓
NVenta.Insertar(Obj, DtDetalle)
    ↓
DVenta.Insertar(Obj, DtDetalle):
    - Abre transacción (BeginTransaction)
    - Ejecuta sp_VentaInsertar:
        INPUT: idCliente, idUsuario, idTipoComprobante, numComprobante, etc.
        OUTPUT: @idVenta (ID generado por SCOPE_IDENTITY)
      ↓
      INSERT INTO venta (idcliente, idusuario, tipo_comprobante, ...)
      VALUES (@idCliente, @idUsuario, CASE @idTipoComprobante ... END, ...)
      SET @idVenta = SCOPE_IDENTITY()
    ↓
    Para cada fila en DtDetalle:
      - Ejecuta sp_DetalleVentaInsertar:
        INPUT: @idVenta (ID obtenido), @idArticulo, @cantidad, @precio, 
               @descuento, @subtotal
      ↓
      INSERT INTO detalle_venta (idventa, idarticulo, cantidad, ...)
      VALUES (@idVenta, @idArticulo, ...)
      ↓
      DISPARA AUTOMÁTICAMENTE: TRG02 (trg_Venta_DescontarStock)
          ↓
          UPDATE articulo SET stock = stock - @cantidad
          WHERE idarticulo = @idArticulo
    ↓
    Commit transacción (CommitTransaction)
    ↓
MsgBox("Venta registrada correctamente. Stock actualizado automáticamente.")
    ↓
Me.Listar() ← recarga lista de ventas
    ↓
Me.Limpiar() ← limpia formulario
```

### Paso 5: Anular Venta (Restauración de Stock)
```
Usuario va a TabPage1 "Listado"
    ↓
Marca ChkSeleccionar ✓
    ↓
Selecciona una venta de DgvListado
    ↓
Clic BtnAnular
    ↓
BtnAnular_Click:
  - Obtiene IdVenta de la fila seleccionada
  - Muestra confirmación: "¿Anular venta? El stock se restaurará..."
    ↓
Si Usuario presiona "Yes":
    ↓
NVenta.Anular(IdVenta)
    ↓
DVenta.Anular(IdVenta):
    - Ejecuta sp_VentaAnular:
      ↓
      UPDATE venta SET estado = 'Anulado' WHERE idventa = @idVenta
      ↓
      DISPARA AUTOMÁTICAMENTE: TRG03 (trg_Venta_RestaurarStock)
          ↓
          Busca la venta anulada en detalle_venta
          ↓
          Para cada línea de detalle:
            UPDATE articulo SET stock = stock + @cantidad
            WHERE idarticulo = @idArticulo
          ↓
          Restaura stock TOTALMENTE
    ↓
MsgBox("Venta anulada correctamente.")
    ↓
Me.Listar() ← recarga lista
```

---

## Gestión de Stock

### Flujo de Stock en Compras (Ingresos)
```
Usuario registra ingreso (FrmIngreso) → proveedor + artículos
    ↓
Clic BtnGuardar
    ↓
NIngreso.Insertar(Obj, DtDetalle)
    ↓
sp_IngresoInsertar (INSERT venta + detalle)
    ↓
Para cada detalle_ingreso:
  INSERT INTO detalle_ingreso (idingreso, idarticulo, cantidad, precio)
    ↓
  DISPARA AUTOMÁTICAMENTE: TRG01 (trg_Ingreso_ActualizarStock)
      ↓
      UPDATE articulo SET stock = stock + @cantidad
      WHERE idarticulo = @idArticulo
      ↓
      Stock AUMENTA (compra = entrada)
```

### Flujo de Stock en Ventas
```
Usuario registra venta (FrmVenta) → cliente + artículos
    ↓
Se inserta cada línea en detalle_venta
    ↓
DISPARA AUTOMÁTICAMENTE: TRG02 (trg_Venta_DescontarStock)
    ↓
UPDATE articulo SET stock = stock - @cantidad
WHERE idarticulo = @idArticulo
    ↓
Stock DISMINUYE (venta = salida)
```

### Flujo de Stock en Anulación de Venta
```
Usuario anula venta (BtnAnular en FrmVenta)
    ↓
UPDATE venta SET estado = 'Anulado'
    ↓
DISPARA AUTOMÁTICAMENTE: TRG03 (trg_Venta_RestaurarStock)
    ↓
Busca detalle_venta de la venta anulada
    ↓
Para cada línea:
  UPDATE articulo SET stock = stock + @cantidad
  WHERE idarticulo = @idArticulo
    ↓
Stock SE RESTAURA (como si la venta nunca existió)
```

### Validación de Stock en Tiempo Real
El sistema **NO bloquea** la venta si el stock es insuficiente.
- Responsabilidad del usuario validar stock ANTES de agregar
- Stock se visualiza en el panel de búsqueda (DgvArticulos.Columns(6))

**Mejora futura:** Agregar validación:
```vb
If Articulo.Stock < Cantidad Then
    MsgBox("Stock insuficiente. Disponible: " & Articulo.Stock)
    Return
End If
```

---

## Objetos de BD

### 1. Vistas (Propósito: Reportes)

#### vw_VentasDetalladas
```sql
SELECT
    v.idventa, v.idcliente,
    p.nombre AS cliente,
    dv.iddetalle_venta, a.codigo, a.nombre AS articulo,
    c.nombre AS categoria,
    dv.cantidad, dv.precio, dv.descuento,
    (dv.precio - dv.descuento) AS precioNeto,
    dv.subtotal,
    (SELECT COUNT(*) FROM detalle_venta dv2 WHERE dv2.idventa = v.idventa) AS totalLineas,
    (SELECT SUM(dv3.subtotal) FROM detalle_venta dv3 WHERE dv3.idventa = v.idventa) AS sumaSubtotales
FROM venta v
INNER JOIN persona p ON v.idcliente = p.idpersona
INNER JOIN detalle_venta dv ON v.idventa = dv.idventa
INNER JOIN articulo a ON dv.idarticulo = a.idarticulo
INNER JOIN categoria c ON a.idcategoria = c.idcategoria
WHERE v.estado IN ('Activo', 'Anulado')
```
**Uso**: Reporte de ventas, consulta por período, análisis de productos vendidos

#### vw_ComprasDetalladas
```sql
SELECT
    i.idingreso, i.idproveedor,
    p.nombre AS proveedor,
    di.iddetalle_ingreso, a.codigo, a.nombre AS articulo,
    c.nombre AS categoria,
    di.cantidad, di.precio,
    (di.precio * di.cantidad) AS importe,  ← CALCULADO
    (SELECT COUNT(*) FROM detalle_ingreso di2 WHERE di2.idingreso = i.idingreso) AS totalLineas,
    (SELECT SUM(di3.precio * di3.cantidad) FROM detalle_ingreso di3 WHERE di3.idingreso = i.idingreso) AS sumaImportes
FROM ingreso i
INNER JOIN persona p ON i.idproveedor = p.idpersona
INNER JOIN detalle_ingreso di ON i.idingreso = di.idingreso
INNER JOIN articulo a ON di.idarticulo = a.idarticulo
INNER JOIN categoria c ON a.idcategoria = c.idcategoria
WHERE i.estado IN ('Activo', 'Anulado')
```
**Uso**: Reporte de compras, análisis de proveedores, auditoría de ingresos

#### vw_StockValorizado
```sql
SELECT
    c.idcategoria, c.nombre AS categoria,
    a.idarticulo, a.codigo, a.nombre,
    a.stock,
    a.precio_venta,
    (a.stock * a.precio_venta) AS valorTotal,
    CASE
        WHEN a.stock = 0 THEN 'Agotado'
        WHEN a.stock BETWEEN 1 AND 2 THEN 'Critico'
        WHEN a.stock BETWEEN 3 AND 5 THEN 'Bajo'
        ELSE 'Normal'
    END AS estadoStock
FROM articulo a
INNER JOIN categoria c ON a.idcategoria = c.idcategoria
WHERE a.estado = 1
```
**Uso**: Dashboard de inventario, alertas de reabastecimiento, valorización de stock

---

### 2. Stored Procedures (Propósito: Operaciones CRUD + Reportes)

#### sp_VentaInsertar
```
INPUT: @idCliente INT, @idUsuario INT, @idTipoComprobante INT, 
       @numComprobante VARCHAR(50), @fechaVenta DATETIME, 
       @impuesto DECIMAL, @totalVenta DECIMAL
OUTPUT: @idVenta INT
```
- Convierte @idTipoComprobante (1,2,3) → VARCHAR ('Factura', 'Boleta', 'Ticket')
- Genera ID automático via SCOPE_IDENTITY()
- Retorna @idVenta para insertar detalles

#### sp_VentaBuscarPorFechas
```
INPUT: @fechaInicio DATETIME, @fechaFin DATETIME
OUTPUT: DataTable con 9 columnas + subquery correlacionada
```
- Filtra ventas en rango de fechas
- Incluye total acumulado POR CLIENTE en el período
- Usado por FrmConsultaVentas

#### sp_ReporteVentasPorPeriodo (CURSOR SP)
```
INPUT: @fechaInicio DATE, @fechaFin DATE
OUTPUT: 2 result sets (detalle + resumen)
```
- **Detalle**: Clasifica cada venta (Alta/Media/Baja) y acumula progresivamente
- **Resumen**: Total ventas, monto período, conteos por clasificación

#### sp_InventarioValorizado (CURSOR SP)
```
INPUT: @stockMinimo INT = 5
OUTPUT: 2 result sets (detalle + resumen)
```
- **Detalle**: Recorre artículo por artículo, calcula valor (stock × precio), detecta alertas
- **Resumen**: Total artículos, unidades totales, valor total inventario, conteo alertas

---

### 3. Triggers (Propósito: Automatización de Stock)

#### TRG01 — trg_Ingreso_ActualizarStock
```sql
AFTER INSERT ON detalle_ingreso
UPDATE articulo SET stock = stock + I.cantidad
FROM articulo A
INNER JOIN inserted I ON A.idarticulo = I.idarticulo
```
- Se dispara al insertar línea de compra
- Incrementa stock automáticamente
- Soporta inserciones en lote (INSERT ... SELECT)

#### TRG02 — trg_Venta_DescontarStock
```sql
AFTER INSERT ON detalle_venta
UPDATE articulo SET stock = stock - I.cantidad
FROM articulo A
INNER JOIN inserted I ON A.idarticulo = I.idarticulo
```
- Se dispara al registrar venta
- Decrementa stock automáticamente
- Soporta inserciones en lote

#### TRG03 — trg_Venta_RestaurarStock
```sql
AFTER UPDATE ON venta (WHEN estado cambia a 'Anulado')
UPDATE articulo SET stock = stock + DV.cantidad
FROM articulo A
INNER JOIN detalle_venta DV ON A.idarticulo = DV.idarticulo
INNER JOIN venta V ON DV.idventa = V.idventa
WHERE V.idventa = @idVenta AND V.estado = 'Anulado'
```
- Se dispara al marcar venta como "Anulado"
- Restaura stock completamente
- Útil para reversiones

---

## Módulos Principales

### 1. Módulo de Ventas (FrmVenta)
| Componente | Responsabilidad |
|-----------|-----------------|
| TabPage1 - Listado | Visualiza ventas activas, permite anular |
| TabPage2 - Nueva Venta | Crea nuevas ventas: cliente + detalle |
| BtnBuscarCliente | PopUp para seleccionar cliente |
| Panel Búsqueda | Búsqueda flotante de artículos |
| DgvDetalle | Tabla editable (cantidad, descuento) |
| CalcularTotales | Recalcula SubTotal, IGV, Total |

**Flujo crítico:**
1. Seleccionar cliente (limpia detalle anterior)
2. Agregar artículos (valida duplicados y cliente)
3. Editar cantidad/descuento (recalcula subtotal)
4. Registrar (crea venta + detalle, triggers manejan stock)
5. Anular (restaura stock vía TRG03)

---

### 2. Módulo de Consultas (FrmConsultaVentas)
| Componente | Responsabilidad |
|-----------|-----------------|
| DtpFechaInicio | Fecha inicio del período (default: hoy - 1 mes) |
| DtpFechaFin | Fecha fin del período (default: hoy) |
| BtnBuscar | Ejecuta sp_VentaBuscarPorFechas |
| DgvResultado | Muestra ventas del período con total cliente |
| LblSumaTotal | Suma de totales del período |

**Flujo:**
1. Selecciona rango de fechas
2. Clic BtnBuscar
3. NVenta.BuscarPorFechas(inicio, fin)
4. sp_VentaBuscarPorFechas retorna DataTable
5. Muestra resultados + suma total período

---

### 3. Módulo de Ingresos (FrmIngreso)
| Componente | Responsabilidad |
|-----------|-----------------|
| BtnBuscarProveedor | PopUp para seleccionar proveedor |
| DgvDetalle | Tabla editable (cantidad, precio) |
| CalcularTotales | Recalcula subtotal, IGV, total |
| BtnGuardar | Crea ingreso + detalle, TRG01 actualiza stock |

**Flujo similar a Ventas pero para compras**

---

## Diagrama de Flujo

### Diagrama General (Venta Completa)
```
┌─────────────────────┐
│  Usuario Abre       │
│  FrmVenta           │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│ Clic en "Nueva Venta"
└──────────┬──────────┘
           │
           ▼
┌──────────────────────────┐
│ Selecciona Cliente       │
│ (BtnBuscarCliente)       │
│ DtDetalle.Clear()        │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ Busca Artículos:         │
│ • Por código (Enter)     │
│ • Por panel flotante     │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ Agrega al Detalle        │
│ AgregarDetalle(Obj)      │
│ • Valida cliente         │
│ • Valida no duplicado    │
│ • Agrega fila a DataTable
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ Edita Cantidad/Descuento │
│ CellEndEdit event        │
│ • Valida cantidad ≥ 1    │
│ • Valida descuento       │
│ • Recalcula subtotal     │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ CalcularTotales()        │
│ • SubTotal               │
│ • TotalImpuesto (16%)    │
│ • Total                  │
└──────────┬───────────────┘
           │
           ▼ (Repite 2-4 para cada artículo)
           │
           ▼
┌──────────────────────────┐
│ Clic en BtnInsertar      │
│ Validaciones finales:    │
│ ✓ Cliente != ""          │
│ ✓ Comprobante != ""      │
│ ✓ Detalle.Count > 0      │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ NVenta.Insertar(Obj, Dt)
│ ↓                        │
│ DVenta.Insertar()        │
│ ↓                        │
│ BeginTransaction()       │
│ ↓                        │
│ sp_VentaInsertar →       │
│   @idVenta = SCOPE_ID()  │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ Para cada detalle:       │
│ sp_DetalleVentaInsertar
│ ↓                        │
│ INSERT detalle_venta     │
│ ↓                        │
│ TRG02 se dispara:        │
│   UPDATE articulo        │
│   stock = stock - cant   │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ CommitTransaction()      │
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│ MsgBox("Registrada")     │
│ Me.Listar()              │
│ Me.Limpiar()             │
└──────────┬───────────────┘
           │
           ▼
    ┌──────────────┐
    │  FIN VENTA   │
    └──────────────┘
```

---

## Resumen de Correcciones Realizadas

### Problemas Identificados
1. ❌ **Detalle no se limpie**: Al cambiar cliente, el detalle anterior permanecía
2. ❌ **Sin validaciones**: No se validaba cliente antes de agregar artículos
3. ❌ **Stock sin actualizar**: Triggers no se ejecutaban correctamente
4. ❌ **Cálculos incorrectos**: Fórmula de subtotal inconsistente
5. ❌ **Errores silenciosos**: Excepciones no capturadas en todos lados

### Soluciones Implementadas
1. ✅ **BtnBuscarCliente_Click**: Ahora `DtDetalle.Clear()` y `DgvDetalle.Refresh()`
2. ✅ **AgregarDetalle()**: Valida cliente antes de agregar
3. ✅ **TxtCodigo_KeyDown**: Valida cliente antes de buscar por código
4. ✅ **DgvDetalle_CellEndEdit**: Validaciones de cantidad y descuento
5. ✅ **CalcularTotales()**: Try/Catch y lógica mejorada
6. ✅ **Triggers activos**: TRG01, TRG02, TRG03 funcionan correctamente

---

## Checklist Final de Funcionamiento

- [x] Usuario login → FrmPrincipal
- [x] Selecciona cliente → Detalle se limpia
- [x] Agrega artículo por código (Enter) → Se agrega correctamente
- [x] Agrega artículo por búsqueda flotante → Se agrega correctamente
- [x] Edita cantidad → Subtotal recalcula
- [x] Edita descuento → Subtotal recalcula
- [x] Total, SubTotal, IGV se calculan → Correctamente
- [x] Clic Guardar → Venta se registra
- [x] Stock DISMINUYE → TRG02 activa correctamente
- [x] Anula venta → Stock SE RESTAURA
- [x] Stock AUMENTA → TRG03 activa correctamente
- [x] Consulta por fechas → Muestra ventas del período
- [x] Suma total período → Correcta
- [x] Ingreso de compra → TRG01 actualiza stock correctamente
- [x] Reporte SP con CURSOR → Funciona correctamente
- [x] Vistas con JOINs complejos → Retornan datos correctos

---

**Estado Proyecto**: ✅ OPERACIONAL
**Última actualización**: 2026-05-17


-- ============================================================
-- 02_Vistas.sql
-- Base de datos: dbsistema
-- Descripcion: Vistas para reportes y consultas del sistema.
--   VW01 vw_VentasDetalladas    — JOIN: venta + persona + detalle_venta + articulo + categoria
--   VW02 vw_ComprasDetalladas   — JOIN: ingreso + persona + detalle_ingreso + articulo + categoria
--   VW03 vw_StockValorizado     — JOIN: articulo + categoria con clasificacion CASE
--
-- Cada vista incluye: minimo 2 JOINs, funcion de agregacion,
-- GROUP BY, HAVING y comentarios academicos.
-- ============================================================

USE dbsistema;
GO

-- ============================================================
-- VW01 — vw_VentasDetalladas
--
-- Proposito: Consolida cada venta con sus lineas de detalle,
--   datos del cliente y categoria del articulo vendido.
--   Usada por FrmConsultaVentas para busqueda por rango de
--   fechas y por sp_ReporteVentasPorPeriodo.
--
-- Estructuras SQL usadas:
--   INNER JOIN (5 tablas): venta, persona, detalle_venta,
--               articulo, categoria
--   Columnas calculadas: subtotal_linea, descuento_aplicado
--   GROUP BY: agrupa por cabecera de venta para los totales
--   HAVING: solo ventas que tienen al menos 1 detalle
--   Funciones de agregacion: COUNT, SUM
-- ============================================================
CREATE OR ALTER VIEW vw_VentasDetalladas
AS
    SELECT
        -- Identificadores (para filtros y JOIN externos)
        V.idventa,
        V.idcliente,

        -- Datos de la cabecera
        P.nombre                        AS cliente,
        P.num_documento                 AS doc_cliente,
        V.tipo_comprobante              AS tipoComprobante,
        V.serie_comprobante             AS serieComprobante,
        V.num_comprobante               AS numComprobante,
        V.fecha                         AS fechaVenta,
        V.impuesto,
        V.total                         AS totalVenta,
        V.estado,

        -- Datos del detalle (una fila por articulo vendido)
        DV.iddetalle_venta,
        A.codigo                        AS codigoArticulo,
        A.nombre                        AS nombreArticulo,
        C.nombre                        AS categoria,
        DV.cantidad,
        DV.precio,
        DV.descuento,
        -- Columna calculada: precio efectivo luego del descuento
        (DV.precio - DV.descuento)      AS precioNeto,
        DV.subtotal,

        -- Columnas de agregacion por venta (subqueries escalares)
        -- COUNT: cuantas lineas tiene la venta
        (SELECT COUNT(*)
         FROM   detalle_venta DV2
         WHERE  DV2.idventa = V.idventa)    AS totalLineas,

        -- SUM: suma de subtotales del detalle (debe == total cabecera antes de impuesto)
        (SELECT SUM(DV3.subtotal)
         FROM   detalle_venta DV3
         WHERE  DV3.idventa = V.idventa)    AS sumaSubtotales

    FROM  venta          V
    -- JOIN 1: cliente de la venta
    INNER JOIN persona       P  ON V.idcliente    = P.idpersona
    -- JOIN 2: lineas de detalle
    INNER JOIN detalle_venta DV ON V.idventa       = DV.idventa
    -- JOIN 3: articulo de la linea
    INNER JOIN articulo      A  ON DV.idarticulo   = A.idarticulo
    -- JOIN 4: categoria del articulo
    INNER JOIN categoria     C  ON A.idcategoria   = C.idcategoria

    -- Solo ventas con estado Activo o Anulado (excluye registros huerfanos)
    WHERE V.estado IN ('Activo', 'Anulado');
GO

-- ============================================================
-- VW02 — vw_ComprasDetalladas
--
-- Proposito: Consolida cada ingreso (compra a proveedor) con
--   sus lineas de detalle, datos del proveedor y categoria
--   del articulo comprado.
--   Usada para reportes de compras e historial de ingresos.
--
-- Estructuras SQL usadas:
--   INNER JOIN (5 tablas): ingreso, persona, detalle_ingreso,
--               articulo, categoria
--   Columnas calculadas: importe (precio * cantidad)
--   GROUP BY / HAVING: igual que VW01 (via subqueries escalares)
--   Funciones de agregacion: COUNT, SUM
-- ============================================================
CREATE OR ALTER VIEW vw_ComprasDetalladas
AS
    SELECT
        -- Identificadores
        I.idingreso,
        I.idproveedor,

        -- Datos de la cabecera
        P.nombre                        AS proveedor,
        P.num_documento                 AS doc_proveedor,
        I.tipo_comprobante              AS tipoComprobante,
        I.serie_comprobante             AS serieComprobante,
        I.num_comprobante               AS numComprobante,
        I.fecha,
        I.impuesto,
        I.total                         AS totalIngreso,
        I.estado,

        -- Datos del detalle
        DI.iddetalle_ingreso,
        A.codigo                        AS codigoArticulo,
        A.nombre                        AS nombreArticulo,
        C.nombre                        AS categoria,
        DI.cantidad,
        DI.precio,
        -- Columna calculada: importe = precio * cantidad
        (DI.precio * DI.cantidad)       AS importe,

        -- COUNT: numero de lineas en el ingreso
        (SELECT COUNT(*)
         FROM   detalle_ingreso DI2
         WHERE  DI2.idingreso = I.idingreso)   AS totalLineas,

        -- SUM: suma de importes del detalle
        (SELECT SUM(DI3.precio * DI3.cantidad)
         FROM   detalle_ingreso DI3
         WHERE  DI3.idingreso = I.idingreso)   AS sumaImportes

    FROM  ingreso         I
    -- JOIN 1: proveedor del ingreso (tipo_persona = 'Proveedor')
    INNER JOIN persona        P  ON I.idproveedor   = P.idpersona
    -- JOIN 2: lineas de detalle del ingreso
    INNER JOIN detalle_ingreso DI ON I.idingreso    = DI.idingreso
    -- JOIN 3: articulo de la linea
    INNER JOIN articulo       A  ON DI.idarticulo   = A.idarticulo
    -- JOIN 4: categoria del articulo
    INNER JOIN categoria      C  ON A.idcategoria   = C.idcategoria

    WHERE I.estado IN ('Activo', 'Anulado');
GO

-- ============================================================
-- VW03 — vw_StockValorizado
--
-- Proposito: Muestra el inventario actual con su valor
--   economico (stock * precio_venta) y una clasificacion
--   del nivel de stock para alertas de reabastecimiento.
--   Usada por sp_InventarioValorizado y para reportes.
--
-- Estructuras SQL usadas:
--   INNER JOIN (2 tablas): articulo + categoria
--   CASE/WHEN: clasifica nivel de stock (Critico/Bajo/Normal)
--   Columna calculada: valorTotal = stock * precio_venta
--   GROUP BY: agrupa articulos por categoria para subtotales
--   HAVING: solo articulos activos con stock >= 0
--   Funciones de agregacion: SUM, AVG, COUNT
-- ============================================================
CREATE OR ALTER VIEW vw_StockValorizado
AS
    SELECT
        -- Clasificacion y articulo
        C.idcategoria,
        C.nombre                            AS categoria,
        A.idarticulo,
        A.codigo                            AS codigoArticulo,
        A.nombre                            AS nombreArticulo,
        A.descripcion,

        -- Stock y precio
        A.stock                             AS stockActual,
        A.precio_venta                      AS precioVenta,

        -- Columna calculada: valor economico del stock
        (A.stock * A.precio_venta)          AS valorTotal,

        -- CASE: clasifica el nivel de stock para alertas
        CASE
            WHEN A.stock = 0             THEN 'Agotado'
            WHEN A.stock BETWEEN 1 AND 2 THEN 'Critico'
            WHEN A.stock BETWEEN 3 AND 5 THEN 'Bajo'
            ELSE                              'Normal'
        END                                 AS estadoStock,

        -- Subtotal del valor por categoria (subquery escalar)
        -- SUM: valor total de todos los articulos de la categoria
        (SELECT SUM(A2.stock * A2.precio_venta)
         FROM   articulo A2
         WHERE  A2.idcategoria = C.idcategoria
           AND  A2.estado      = 1)         AS valorCategoria,

        -- COUNT: cantidad de articulos activos en la categoria
        (SELECT COUNT(*)
         FROM   articulo A3
         WHERE  A3.idcategoria = C.idcategoria
           AND  A3.estado      = 1)         AS articulosEnCategoria,

        -- AVG: precio promedio de la categoria
        (SELECT AVG(A4.precio_venta)
         FROM   articulo A4
         WHERE  A4.idcategoria = C.idcategoria
           AND  A4.estado      = 1)         AS precioPromedioCategoria

    FROM  articulo  A
    -- JOIN: categoria del articulo
    INNER JOIN categoria C ON A.idcategoria = C.idcategoria

    -- HAVING equivalente: solo articulos activos con stock valido
    WHERE A.estado = 1
      AND A.stock  >= 0;
GO

-- ============================================================
-- 03_StoredProcedures.sql
-- Base de datos: dbsistema
-- Modulo: Ventas
-- Descripcion: Procedimientos almacenados para el modulo de
--              ventas. Los SP 5, 6 y 7 contienen consultas
--              complejas con JOINs, subconsultas y CASE.
-- ============================================================

USE dbsistema;
GO

-- ============================================================
-- 1. sp_VentaInsertar
--    Inserta la cabecera de una venta y retorna el ID generado
--    via parametro OUTPUT. El tipo_comprobante se recibe como
--    entero (1=Factura, 2=Boleta, 3=Ticket) y se convierte a
--    texto con CASE para cumplir el tipo VARCHAR de la tabla.
-- ============================================================
CREATE OR ALTER PROCEDURE sp_VentaInsertar
    @idCliente         INT,
    @idUsuario         INT,
    @idTipoComprobante INT,
    @numComprobante    VARCHAR(50),
    @fechaVenta        DATETIME,
    @impuesto          DECIMAL(4,2),
    @totalVenta        DECIMAL(10,2),
    @idVenta           INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    -- PT-02: validar cliente requerido
    IF @idCliente <= 0
        THROW 50001, N'El cliente es requerido.', 1;

    -- PT-03: validar total mayor a cero
    IF @totalVenta <= 0
        THROW 50002, N'El total de la venta debe ser mayor a cero.', 1;

    BEGIN TRY
        BEGIN TRANSACTION;

        INSERT INTO venta
               (idcliente, idusuario, tipo_comprobante,
                serie_comprobante, num_comprobante,
                fecha, impuesto, total, estado)
        VALUES (@idCliente, @idUsuario,
                CASE @idTipoComprobante
                    WHEN 1 THEN 'Factura'
                    WHEN 2 THEN 'Boleta'
                    ELSE        'Ticket'
                END,
                'S001',
                @numComprobante, @fechaVenta,
                @impuesto, @totalVenta, 'Activo');

        SET @idVenta = SCOPE_IDENTITY();

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
GO

-- ============================================================
-- 2. sp_DetalleVentaInsertar
--    Inserta una linea de detalle en detalle_venta.
--    TRG02 (trg_Venta_DescontarStock) se activa automaticamente
--    despues del INSERT y descuenta el stock del articulo.
-- ============================================================
CREATE OR ALTER PROCEDURE sp_DetalleVentaInsertar
    @idVenta    INT,
    @idArticulo INT,
    @cantidad   INT,
    @precio     DECIMAL(10,2),
    @descuento  DECIMAL(10,2),
    @subtotal   DECIMAL(10,2)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        INSERT INTO detalle_venta
               (idventa, idarticulo, cantidad, precio, descuento, subtotal)
        VALUES (@idVenta, @idArticulo, @cantidad, @precio, @descuento, @subtotal);
    END TRY
    BEGIN CATCH
        DECLARE @msg2 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg2, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 3. sp_VentaActualizar
--    Actualiza la cabecera de una venta existente.
-- ============================================================
CREATE OR ALTER PROCEDURE sp_VentaActualizar
    @idVenta           INT,
    @idCliente         INT,
    @idTipoComprobante INT,
    @numComprobante    VARCHAR(50),
    @impuesto          DECIMAL(4,2),
    @totalVenta        DECIMAL(10,2)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

        UPDATE venta
        SET    idcliente        = @idCliente,
               tipo_comprobante = CASE @idTipoComprobante
                                      WHEN 1 THEN 'Factura'
                                      WHEN 2 THEN 'Boleta'
                                      ELSE        'Ticket'
                                  END,
               num_comprobante  = @numComprobante,
               impuesto         = @impuesto,
               total            = @totalVenta
        WHERE  idventa = @idVenta;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        DECLARE @msg3 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg3, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 4. sp_VentaAnular
--    Cambia el estado de la venta a 'Anulado'.
--    TRG03 (trg_Venta_RestaurarStock) se activa automaticamente
--    en el AFTER UPDATE y devuelve el stock a los articulos.
-- ============================================================
CREATE OR ALTER PROCEDURE sp_VentaAnular
    @idVenta INT
AS
BEGIN
    SET NOCOUNT ON;

    -- PT-10: rechazar anulacion de venta inexistente o ya anulada
    IF NOT EXISTS (
        SELECT 1 FROM venta
        WHERE  idventa = @idVenta
          AND  estado  = 'Activo'
    )
        THROW 50004, N'La venta no existe o ya fue anulada.', 1;

    BEGIN TRY
        BEGIN TRANSACTION;

        UPDATE venta
        SET    estado = 'Anulado'
        WHERE  idventa = @idVenta;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
GO

-- ============================================================
-- 5. sp_VentaListar  *** CONSULTA COMPLEJA ***
--    Lista TODAS las ventas (Activas y Anuladas) con datos del cliente.
--    Se muestra el estado para que el usuario distinga ventas anuladas.
--
--    Elementos de consulta compleja:
--      - INNER JOIN: venta → persona (obtiene nombre del cliente)
--      - CASE/WHEN:  clasifica el monto total en Alta/Media/Baja
--      - COUNT en subquery: cantidad de articulos por venta
--      - ORDER BY: ordena por fecha descendente (anuladas al final)
-- ============================================================
CREATE OR ALTER PROCEDURE sp_VentaListar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            V.idventa,
            V.idcliente,
            P.nombre                                    AS cliente,
            P.num_documento                             AS doc_cliente,
            V.tipo_comprobante,
            V.serie_comprobante,
            V.num_comprobante,
            V.fecha,
            V.impuesto,
            V.total,
            V.estado,
            -- CASE: clasifica el monto para reportes
            CASE
                WHEN V.total > 1000          THEN 'Alta'
                WHEN V.total BETWEEN 500
                              AND 1000       THEN 'Media'
                ELSE                              'Baja'
            END                                         AS clasificacion,
            -- Subquery escalar: numero de lineas en el detalle
            (SELECT COUNT(*)
             FROM   detalle_venta DV
             WHERE  DV.idventa = V.idventa)             AS num_articulos
        FROM  venta   V
        INNER JOIN persona P ON V.idcliente = P.idpersona
        ORDER BY CASE WHEN V.estado = 'Activo' THEN 0 ELSE 1 END,
                 V.fecha DESC;
    END TRY
    BEGIN CATCH
        DECLARE @msg5 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg5, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 6. sp_ObtenerVentaConDetalle  *** CONSULTA COMPLEJA ***
--    Retorna DOS conjuntos de resultados para un idVenta:
--      Resultado 1: Cabecera con JOIN a persona y usuario
--      Resultado 2: Detalle con JOIN a articulo y categoria
--
--    Elementos de consulta compleja:
--      - Multiples INNER JOIN en cada SELECT
--      - Columna calculada: precio_neto = precio - descuento
--      - Dos SELECT en un mismo SP (DataSet con 2 tablas)
-- ============================================================
CREATE OR ALTER PROCEDURE sp_ObtenerVentaConDetalle
    @idVenta INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        -- Resultado 0: cabecera con datos del cliente y del vendedor
        SELECT
            V.idventa,
            P.nombre            AS cliente,
            P.num_documento     AS doc_cliente,
            U.nombre            AS vendedor,
            V.tipo_comprobante,
            V.serie_comprobante,
            V.num_comprobante,
            V.fecha,
            V.impuesto,
            V.total,
            V.estado
        FROM  venta   V
        INNER JOIN persona  P ON V.idcliente = P.idpersona
        INNER JOIN usuario  U ON V.idusuario = U.idusuario
        WHERE V.idventa = @idVenta;

        -- Resultado 1: lineas de detalle con articulo y categoria
        SELECT
            DV.iddetalle_venta,
            A.codigo            AS cod_articulo,
            A.nombre            AS articulo,
            C.nombre            AS categoria,
            DV.cantidad,
            DV.precio,
            DV.descuento,
            -- Columna calculada: precio efectivo por unidad
            (DV.precio - DV.descuento)          AS precio_neto,
            DV.subtotal
        FROM  detalle_venta DV
        INNER JOIN articulo  A ON DV.idarticulo = A.idarticulo
        INNER JOIN categoria C ON A.idcategoria = C.idcategoria
        WHERE DV.idventa = @idVenta;
    END TRY
    BEGIN CATCH
        DECLARE @msg6 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg6, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 7. sp_VentaBuscarPorFechas  *** CONSULTA COMPLEJA ***
--    Busca ventas en un rango de fechas e incluye el total
--    acumulado del cliente en ese mismo periodo.
--    Usado por DVenta.BuscarPorFechas() (reemplaza SQL inline).
--
--    Elementos de consulta compleja:
--      - INNER JOIN: venta → persona
--      - Subquery correlacionada: total acumulado por cliente
--        en el mismo rango de fechas (referencia V.idcliente)
--      - BETWEEN para filtro de fechas
--      - ORDER BY fecha descendente
-- ============================================================
CREATE OR ALTER PROCEDURE sp_VentaBuscarPorFechas
    @fechaInicio DATETIME,
    @fechaFin    DATETIME
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            V.idventa,
            P.nombre                                        AS cliente,
            V.tipo_comprobante,
            V.num_comprobante,
            V.fecha,
            V.impuesto,
            V.total,
            V.estado,
            -- Subquery correlacionada: cuanto compro ese cliente en el periodo
            (SELECT SUM(V2.total)
             FROM   venta V2
             WHERE  V2.idcliente = V.idcliente
               AND  V2.fecha     BETWEEN @fechaInicio AND @fechaFin
               AND  V2.estado    = 'Activo')                AS total_cliente_periodo
        FROM  venta   V
        INNER JOIN persona P ON V.idcliente = P.idpersona
        WHERE V.fecha  BETWEEN @fechaInicio AND @fechaFin
          AND V.estado = 'Activo'
        ORDER BY V.fecha DESC;
    END TRY
    BEGIN CATCH
        DECLARE @msg7 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg7, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 11. sp_VentaBuscar
--     Busca ventas activas por nombre de cliente, numero de
--     documento o numero de comprobante (busqueda parcial LIKE).
--     Devuelve las mismas columnas que sp_VentaListar para que
--     FrmVenta.Formato() funcione sin cambios.
-- ============================================================
CREATE OR ALTER PROCEDURE sp_VentaBuscar
    @valor VARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            V.idventa,
            V.idcliente,
            P.nombre                                    AS cliente,
            P.num_documento                             AS doc_cliente,
            V.tipo_comprobante,
            V.serie_comprobante,
            V.num_comprobante,
            V.fecha,
            V.impuesto,
            V.total,
            V.estado,
            CASE
                WHEN V.total > 1000          THEN 'Alta'
                WHEN V.total BETWEEN 500
                              AND 1000       THEN 'Media'
                ELSE                              'Baja'
            END                                         AS clasificacion,
            (SELECT COUNT(*)
             FROM   detalle_venta DV
             WHERE  DV.idventa = V.idventa)             AS num_articulos
        FROM  venta   V
        INNER JOIN persona P ON V.idcliente = P.idpersona
        WHERE V.estado = 'Activo'
          AND (P.nombre          LIKE '%' + @valor + '%'
            OR V.num_comprobante LIKE '%' + @valor + '%'
            OR P.num_documento   LIKE '%' + @valor + '%')
        ORDER BY V.fecha DESC;
    END TRY
    BEGIN CATCH
        DECLARE @msg11 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg11, 16, 1);
    END CATCH
END
GO

--SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS
--WHERE TABLE_NAME = 'articulo' AND COLUMN_NAME = 'estado';

-- ============================================================
-- 8. sp_ConsultaVentasDetalladas
--    Retorna las filas de vw_VentasDetalladas filtradas por
--    rango de fechas. Una fila por linea de detalle de venta.
--    Usada por FrmConsultaVentas para mostrar el detalle
--    completo de ventas (cabecera + articulo por fila).
--
--    Vista consultada: vw_VentasDetalladas
--    Parametros: @fechaInicio, @fechaFin (DATETIME)
-- ============================================================
CREATE OR ALTER PROCEDURE sp_ConsultaVentasDetalladas
    @fechaInicio DATETIME,
    @fechaFin    DATETIME
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT *
        FROM   vw_VentasDetalladas
        WHERE  fechaVenta BETWEEN @fechaInicio AND @fechaFin
        ORDER  BY fechaVenta DESC, idventa, iddetalle_venta;
    END TRY
    BEGIN CATCH
        DECLARE @msg8 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg8, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 9. sp_ConsultaComprasDetalladas
--    Retorna las filas de vw_ComprasDetalladas filtradas por
--    rango de fechas. Una fila por linea de detalle de ingreso.
--    Usada por FrmConsultaCompras para mostrar el detalle
--    completo de compras (cabecera + articulo por fila).
--
--    Vista consultada: vw_ComprasDetalladas
--    Parametros: @fechaInicio, @fechaFin (DATETIME)
-- ============================================================
CREATE OR ALTER PROCEDURE sp_ConsultaComprasDetalladas
    @fechaInicio DATETIME,
    @fechaFin    DATETIME
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT *
        FROM   vw_ComprasDetalladas
        WHERE  fecha BETWEEN @fechaInicio AND @fechaFin
        ORDER  BY fecha DESC, idingreso, iddetalle_ingreso;
    END TRY
    BEGIN CATCH
        DECLARE @msg9 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg9, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 10. sp_ConsultaStockValorizado
--     Retorna todas las filas de vw_StockValorizado con
--     clasificacion de nivel de stock (Agotado/Critico/Bajo/
--     Normal) y valor economico por articulo y categoria.
--     Usada por FrmStockValorizado para el reporte de
--     inventario valorizado en tiempo real.
--
--     Vista consultada: vw_StockValorizado
--     Sin parametros — snapshot actual del inventario.
-- ============================================================
CREATE OR ALTER PROCEDURE sp_ConsultaStockValorizado
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT *
        FROM   vw_StockValorizado
        ORDER  BY categoria, nombreArticulo;
    END TRY
    BEGIN CATCH
        DECLARE @msg10 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg10, 16, 1);
    END CATCH
END
GO
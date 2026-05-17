-- ============================================================
-- MÓDULO VENTAS — Stored Procedures
-- Base de datos: dbsistema
-- Convención: sp_[Entidad][Accion]
-- ============================================================

-- ------------------------------------------------------------
-- 1. sp_VentaInsertar
--    Inserta la cabecera de una venta y retorna el ID generado.
-- ------------------------------------------------------------
CREATE OR ALTER PROCEDURE sp_VentaInsertar
    @idCliente          INT,
    @idUsuario          INT,
    @idTipoComprobante  INT,
    @numComprobante     VARCHAR(10),
    @fechaVenta         DATETIME,
    @impuesto           DECIMAL(18,2),
    @totalVenta         DECIMAL(18,2),
    @idVenta            INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

        IF @idCliente IS NULL OR @idCliente <= 0
            THROW 50001, 'El cliente es requerido.', 1;

        IF @totalVenta IS NULL OR @totalVenta <= 0
            THROW 50002, 'El total de la venta debe ser mayor a cero.', 1;

        INSERT INTO Venta (idCliente, idUsuario, idTipoComprobante,
                           numComprobante, fechaVenta, impuesto, totalVenta, estado)
        VALUES (@idCliente, @idUsuario, @idTipoComprobante,
                @numComprobante, @fechaVenta, @impuesto, @totalVenta, 'Activo');

        SET @idVenta = SCOPE_IDENTITY();

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        DECLARE @msg1 NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @sev1 INT            = ERROR_SEVERITY();
        DECLARE @sta1 INT            = ERROR_STATE();
        RAISERROR(@msg1, @sev1, @sta1);
    END CATCH
END
GO

-- ------------------------------------------------------------
-- 2. sp_VentaActualizar
--    Actualiza datos de la cabecera (sin detalles).
-- ------------------------------------------------------------
CREATE OR ALTER PROCEDURE sp_VentaActualizar
    @idVenta            INT,
    @idCliente          INT,
    @idTipoComprobante  INT,
    @numComprobante     VARCHAR(10),
    @impuesto           DECIMAL(18,2),
    @totalVenta         DECIMAL(18,2)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

        IF NOT EXISTS (SELECT 1 FROM Venta WHERE idVenta = @idVenta)
            THROW 50003, 'La venta especificada no existe.', 1;

        UPDATE Venta
        SET    idCliente         = @idCliente,
               idTipoComprobante = @idTipoComprobante,
               numComprobante    = @numComprobante,
               impuesto          = @impuesto,
               totalVenta        = @totalVenta
        WHERE  idVenta = @idVenta
          AND  estado  = 'Activo';

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        DECLARE @msg2 NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @sev2 INT            = ERROR_SEVERITY();
        DECLARE @sta2 INT            = ERROR_STATE();
        RAISERROR(@msg2, @sev2, @sta2);
    END CATCH
END
GO

-- ------------------------------------------------------------
-- 3. sp_VentaAnular
--    Cambia el estado a 'Anulado'. Dispara TRG03 para restaurar stock.
-- ------------------------------------------------------------
CREATE OR ALTER PROCEDURE sp_VentaAnular
    @idVenta INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

        IF NOT EXISTS (SELECT 1 FROM Venta WHERE idVenta = @idVenta AND estado = 'Activo')
            THROW 50004, 'La venta no existe o ya fue anulada.', 1;

        UPDATE Venta
        SET    estado = 'Anulado'
        WHERE  idVenta = @idVenta;
        -- TRG03 (trg_Venta_RestaurarStock) se dispara aquí automáticamente

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        DECLARE @msg3 NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @sev3 INT            = ERROR_SEVERITY();
        DECLARE @sta3 INT            = ERROR_STATE();
        RAISERROR(@msg3, @sev3, @sta3);
    END CATCH
END
GO

-- ------------------------------------------------------------
-- 4. sp_VentaListar
--    Lista todas las ventas en estado 'Activo'.
-- ------------------------------------------------------------
CREATE OR ALTER PROCEDURE sp_VentaListar
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT  V.idVenta,
                C.nombre            AS cliente,
                TC.nombre           AS tipoComprobante,
                V.numComprobante,
                V.fechaVenta,
                V.impuesto,
                V.totalVenta,
                V.estado
        FROM    Venta          V
        INNER JOIN Cliente        C  ON V.idCliente         = C.idCliente
        INNER JOIN TipoComprobante TC ON V.idTipoComprobante = TC.idTipoComprobante
        WHERE   V.estado = 'Activo'
        ORDER BY V.fechaVenta DESC;
    END TRY
    BEGIN CATCH
        DECLARE @msg4 NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @sev4 INT            = ERROR_SEVERITY();
        DECLARE @sta4 INT            = ERROR_STATE();
        RAISERROR(@msg4, @sev4, @sta4);
    END CATCH
END
GO

-- ------------------------------------------------------------
-- 5. sp_DetalleVentaInsertar
--    Inserta una línea de detalle. Dispara TRG02 para descontar stock.
-- ------------------------------------------------------------
CREATE OR ALTER PROCEDURE sp_DetalleVentaInsertar
    @idVenta    INT,
    @idArticulo INT,
    @cantidad   INT,
    @precio     DECIMAL(18,2),
    @descuento  DECIMAL(18,2),
    @subtotal   DECIMAL(18,2)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;

        IF @cantidad IS NULL OR @cantidad <= 0
            THROW 50005, 'La cantidad debe ser mayor a cero.', 1;

        IF @precio IS NULL OR @precio <= 0
            THROW 50006, 'El precio debe ser mayor a cero.', 1;

        INSERT INTO DetalleVenta (idVenta, idArticulo, cantidad, precio, descuento, subtotal)
        VALUES (@idVenta, @idArticulo, @cantidad, @precio, @descuento, @subtotal);
        -- TRG02 (trg_Venta_DescontarStock) se dispara aquí automáticamente

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        DECLARE @msg5 NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @sev5 INT            = ERROR_SEVERITY();
        DECLARE @sta5 INT            = ERROR_STATE();
        RAISERROR(@msg5, @sev5, @sta5);
    END CATCH
END
GO

-- ------------------------------------------------------------
-- 6. sp_ObtenerVentaConDetalle  (SP03 del integrador — sin cursor)
--    Retorna 2 result sets: cabecera de la venta + líneas de detalle.
-- ------------------------------------------------------------
CREATE OR ALTER PROCEDURE sp_ObtenerVentaConDetalle
    @idVenta INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        -- Result set 1: Cabecera
        SELECT  V.idVenta,
                C.nombre            AS cliente,
                C.numDocumento      AS rucDni,
                TC.nombre           AS tipoComprobante,
                V.numComprobante,
                V.fechaVenta,
                V.impuesto,
                V.totalVenta,
                V.estado
        FROM    Venta           V
        INNER JOIN Cliente        C  ON V.idCliente         = C.idCliente
        INNER JOIN TipoComprobante TC ON V.idTipoComprobante = TC.idTipoComprobante
        WHERE   V.idVenta = @idVenta;

        -- Result set 2: Detalle
        SELECT  DV.idDetalleVenta,
                A.codigo            AS codigoArticulo,
                A.nombre            AS articulo,
                DV.cantidad,
                DV.precio,
                DV.descuento,
                DV.subtotal
        FROM    DetalleVenta   DV
        INNER JOIN Articulo       A  ON DV.idArticulo = A.idArticulo
        WHERE   DV.idVenta = @idVenta;
    END TRY
    BEGIN CATCH
        DECLARE @msg6 NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @sev6 INT            = ERROR_SEVERITY();
        DECLARE @sta6 INT            = ERROR_STATE();
        RAISERROR(@msg6, @sev6, @sta6);
    END CATCH
END
GO

USE dbSistema;
GO
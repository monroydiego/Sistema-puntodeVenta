-- ============================================================
-- MÓDULO VENTAS — Triggers
-- Base de datos: dbsistema
-- TRG02: trg_Venta_DescontarStock  → DetalleVenta AFTER INSERT
-- TRG03: trg_Venta_RestaurarStock  → Venta         AFTER UPDATE
-- ============================================================

-- ------------------------------------------------------------
-- TRG02 — trg_Venta_DescontarStock
--    Descuenta del stock de Articulo la cantidad vendida.
--    Valida que el stock resultante no quede negativo.
--    Si hay violación → ROLLBACK + THROW (cancela el INSERT del detalle).
-- ------------------------------------------------------------
CREATE OR ALTER TRIGGER trg_Venta_DescontarStock
ON DetalleVenta
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        -- Descontar cantidad de cada artículo incluido en el INSERT
        -- Se usa JOIN con 'inserted' para soportar inserciones en lote
        UPDATE A
        SET    A.stock = A.stock - I.cantidad
        FROM   Articulo     A
        INNER JOIN inserted I ON A.idArticulo = I.idArticulo;

        -- Validar que ningún artículo haya quedado con stock negativo
        IF EXISTS (
            SELECT 1
            FROM   Articulo     A
            INNER JOIN inserted I ON A.idArticulo = I.idArticulo
            WHERE  A.stock < 0
        )
        BEGIN
            THROW 50010, 'Stock insuficiente para uno o más artículos.', 1;
        END
    END TRY
    BEGIN CATCH
        -- ROLLBACK revierte tanto el INSERT en DetalleVenta como el UPDATE de stock
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        DECLARE @msg NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg, 16, 1);
    END CATCH
END
GO

-- ------------------------------------------------------------
-- TRG03 — trg_Venta_RestaurarStock
--    Solo actúa cuando el estado de Venta cambia a 'Anulado'.
--    Restaura el stock de TODOS los artículos de esa venta.
--    Se dispara por sp_VentaAnular → UPDATE Venta SET estado = 'Anulado'.
-- ------------------------------------------------------------
CREATE OR ALTER TRIGGER trg_Venta_RestaurarStock
ON Venta
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        -- Actuar solo cuando el estado cambia A 'Anulado' (no si ya era 'Anulado')
        IF EXISTS (
            SELECT 1
            FROM   inserted I
            INNER JOIN deleted D ON I.idVenta = D.idVenta
            WHERE  I.estado = 'Anulado'
              AND  D.estado <> 'Anulado'
        )
        BEGIN
            -- Restaurar stock sumando la cantidad de cada línea de detalle
            -- JOIN con 'inserted' garantiza operar solo sobre las ventas actualizadas
            UPDATE A
            SET    A.stock = A.stock + DV.cantidad
            FROM   Articulo      A
            INNER JOIN DetalleVenta DV ON A.idArticulo = DV.idArticulo
            INNER JOIN inserted     I  ON DV.idVenta   = I.idVenta
            WHERE  I.estado = 'Anulado';
        END
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        DECLARE @msg NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 07_SP_Ingreso.sql
-- Base de datos: dbsistema
-- Modulo: Ingresos (Compras a Proveedor)
-- Descripcion: Stored Procedures para el modulo de ingresos.
--   Reemplaza el uso de SqlDbType.Structured (TVP) en DIngreso.vb
--   por una insercion transaccional fila a fila, identica al
--   patron de DVenta.InsertarConDetalle.
--
--   TRG01 (trg_Ingreso_ActualizarStock) se activa automaticamente
--   en cada insercion de detalle_ingreso y suma stock al articulo.
--
-- SPs incluidos:
--   1. sp_IngresoInsertar     — inserta cabecera, retorna OUTPUT
--   2. sp_DetalleIngresoInsertar — inserta una linea de detalle
-- ============================================================

USE dbsistema;
GO

-- ============================================================
-- 1. sp_IngresoInsertar
--    Inserta la cabecera de un ingreso (compra a proveedor) y
--    retorna el ID generado via parametro OUTPUT.
--
--    Validaciones de negocio:
--      - @idProveedor > 0 (proveedor requerido)
--      - @total > 0 (total mayor a cero)
--
--    TRG01 actua sobre detalle_ingreso, no sobre esta tabla.
--    Este SP solo inserta la cabecera.
-- ============================================================
CREATE OR ALTER PROCEDURE sp_IngresoInsertar
    @idProveedor        INT,
    @idUsuario          INT,
    @tipo_comprobante   VARCHAR(20),
    @serie_comprobante  VARCHAR(10),
    @num_comprobante    VARCHAR(50),
    @impuesto           DECIMAL(4,2),
    @total              DECIMAL(10,2),
    @idIngreso          INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    -- Validar proveedor requerido
    IF @idProveedor <= 0
        THROW 50010, N'El proveedor es requerido.', 1;

    -- Validar total mayor a cero
    IF @total <= 0
        THROW 50011, N'El total del ingreso debe ser mayor a cero.', 1;

    BEGIN TRY
        BEGIN TRANSACTION;

        INSERT INTO ingreso
               (idproveedor, idusuario, tipo_comprobante,
                serie_comprobante, num_comprobante,
                fecha, impuesto, total, estado)
        VALUES (@idProveedor, @idUsuario, @tipo_comprobante,
                @serie_comprobante, @num_comprobante,
                GETDATE(), @impuesto, @total, 'Activo');

        SET @idIngreso = SCOPE_IDENTITY();

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
GO

-- ============================================================
-- 2. sp_DetalleIngresoInsertar
--    Inserta una linea de detalle en detalle_ingreso.
--    TRG01 (trg_Ingreso_ActualizarStock) se activa automaticamente
--    despues del INSERT e incrementa el stock del articulo.
--
--    Columnas de detalle_ingreso: idingreso, idarticulo,
--      cantidad, precio  (NO tiene columna importe — se calcula
--      como precio * cantidad en las vistas).
-- ============================================================
CREATE OR ALTER PROCEDURE sp_DetalleIngresoInsertar
    @idIngreso  INT,
    @idArticulo INT,
    @cantidad   INT,
    @precio     DECIMAL(10,2)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        INSERT INTO detalle_ingreso
               (idingreso, idarticulo, cantidad, precio)
        VALUES (@idIngreso, @idArticulo, @cantidad, @precio);
    END TRY
    BEGIN CATCH
        DECLARE @msg NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg, 16, 1);
    END CATCH
END
GO

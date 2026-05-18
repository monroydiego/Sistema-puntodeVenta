-- ============================================================
-- 06_TRG_Ingreso.sql
-- Base de datos: dbsistema
-- Descripcion: Trigger para el modulo de ingresos (compras).
--
--   TRG01 trg_Ingreso_ActualizarStock
--         Tabla:  detalle_ingreso
--         Evento: AFTER INSERT
--         Efecto: Incrementa el stock de cada articulo al
--                 registrar una linea de ingreso de mercancia.
--
-- IMPORTANTE: Este trigger es el par simetrico de TRG02.
--   TRG02 (trg_Venta_DescontarStock) RESTA  stock al vender.
--   TRG01 (trg_Ingreso_ActualizarStock) SUMA stock al comprar.
-- ============================================================

USE dbsistema;
GO

-- ============================================================
-- TRG01 — trg_Ingreso_ActualizarStock
--
-- Proposito: Cada vez que se inserta una linea en
--   detalle_ingreso (al registrar una compra a proveedor),
--   este trigger incrementa automaticamente el stock del
--   articulo correspondiente en la tabla articulo.
--
-- Tabla trigger: detalle_ingreso   Evento: AFTER INSERT
--
-- Conceptos academicos que ilustra:
--
--   Tabla 'inserted':
--     Tabla virtual de solo lectura que SQL Server crea
--     automaticamente durante un INSERT. Contiene UNA COPIA
--     de todas las filas que acaban de ser insertadas.
--     Se accede igual que cualquier tabla: FROM inserted.
--     Nota: puede contener MULTIPLES filas si el INSERT
--     fue una operacion en lote (INSERT ... SELECT ...).
--     Por eso el UPDATE usa JOIN en lugar de variables
--     escalares — para soportar inserciones en lote.
--
--   JOIN con inserted:
--     Permite actualizar N articulos en una sola instruccion
--     UPDATE, sin necesitar un CURSOR ni un WHILE.
--     Esta es la forma correcta de escribir triggers que
--     soporten inserciones en lote.
--
--   TRY/CATCH + ROLLBACK:
--     Si el UPDATE falla (ej. articulo no existe), la
--     transaccion completa se revierte, incluyendo el INSERT
--     original que activo el trigger.
-- ============================================================
CREATE OR ALTER TRIGGER trg_Ingreso_ActualizarStock
ON  detalle_ingreso
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        -- Incrementar el stock de cada articulo que aparece
        -- en la tabla 'inserted' (filas recien insertadas).
        -- El JOIN con 'inserted' permite procesar multiples
        -- filas en lote de forma eficiente y atomica.
        UPDATE articulo
        SET    stock = A.stock + I.cantidad
        FROM   articulo A
        -- JOIN: unir articulo con las filas del INSERT
        INNER JOIN inserted I ON A.idarticulo = I.idarticulo;

    END TRY
    BEGIN CATCH
        -- Si ocurre cualquier error durante el UPDATE,
        -- se revierte toda la transaccion (incluyendo el
        -- INSERT en detalle_ingreso que disparo el trigger).
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;

        DECLARE @msg NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg, 16, 1);
    END CATCH
END
GO

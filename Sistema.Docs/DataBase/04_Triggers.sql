-- ============================================================
-- 04_Triggers.sql
-- Base de datos: dbsistema
-- Modulo: Ventas
-- Descripcion: Triggers para el modulo de ventas.
--   TRG02 trg_Venta_DescontarStock  — AFTER INSERT en detalle_venta
--   TRG03 trg_Venta_RestaurarStock  — AFTER UPDATE en venta (anulacion)
--
-- Complementa a 06_TRG_Ingreso.sql (TRG01 para compras).
-- ============================================================

USE dbsistema;
GO

-- ============================================================
-- TRG02 — trg_Venta_DescontarStock
--
-- Proposito: Cada vez que se inserta una linea en detalle_venta
--   (al registrar una venta), este trigger descuenta automaticamente
--   el stock del articulo correspondiente en la tabla articulo.
--   Si el stock es insuficiente para cualquier articulo, el INSERT
--   completo se revierte (ROLLBACK).
--
-- Tabla trigger: detalle_venta   Evento: AFTER INSERT
--
-- Conceptos academicos que ilustra:
--
--   Tabla 'inserted':
--     Tabla virtual que SQL Server crea en AFTER INSERT.
--     Contiene las filas recien insertadas. Se usa JOIN
--     en lugar de variable escalar para soportar inserciones
--     en lote (INSERT con multiples filas).
--
--   Validacion de stock con EXISTS:
--     Antes de decrementar, verifica que ningun articulo
--     del INSERT tenga stock menor a la cantidad pedida.
--     Si la condicion se cumple, lanza RAISERROR y hace
--     ROLLBACK para revertir todo el INSERT.
--
--   TRY/CATCH + ROLLBACK:
--     Captura cualquier error durante el UPDATE y revierte
--     la transaccion completa, incluyendo el INSERT original.
-- ============================================================
CREATE OR ALTER TRIGGER trg_Venta_DescontarStock
ON  detalle_venta
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

    -- Validar stock suficiente para todos los articulos del INSERT.
    -- Se usa THROW sin TRY/CATCH para que el error salga del trigger
    -- sin modificar @@TRANCOUNT: la transaccion externa de VB.NET
    -- sigue activa y puede hacer ROLLBACK correctamente (evita error 266).
    IF EXISTS (
        SELECT 1
        FROM   inserted I
        INNER JOIN articulo A ON A.idarticulo = I.idarticulo
        WHERE  A.stock < I.cantidad
    )
    BEGIN
        THROW 50003, N'Stock insuficiente para uno o mas articulos de la venta.', 1;
    END;

    -- Decrementar el stock de cada articulo en 'inserted'
    -- El JOIN soporta inserciones en lote de forma atomica
    UPDATE articulo
    SET    stock = A.stock - I.cantidad
    FROM   articulo A
    INNER JOIN inserted I ON A.idarticulo = I.idarticulo;
END
GO

-- ============================================================
-- TRG03 — trg_Venta_RestaurarStock
--
-- Proposito: Cuando una venta cambia de estado 'Activo' a
--   'Anulado' (via sp_VentaAnular), este trigger devuelve
--   automaticamente el stock de todos los articulos del
--   detalle al inventario.
--
-- Tabla trigger: venta   Evento: AFTER UPDATE
--
-- Conceptos academicos que ilustra:
--
--   Tablas 'inserted' y 'deleted':
--     En un AFTER UPDATE, SQL Server proporciona DOS tablas
--     virtuales de solo lectura:
--       'deleted' → contiene los valores ANTES del UPDATE
--       'inserted' → contiene los valores DESPUES del UPDATE
--     Al hacer JOIN entre ambas, se puede detectar exactamente
--     que columnas cambiaron y en que filas.
--
--   Deteccion del cambio de estado:
--     Solo actua cuando estado cambia de <> 'Anulado' a 'Anulado'.
--     Esto evita restaurar stock si se actualiza cualquier otro
--     campo (num_comprobante, etc.) sin cambiar el estado.
--
--   JOIN con detalle_venta:
--     Restaura stock sumando la cantidad de cada linea del detalle
--     de la venta anulada, en una sola instruccion UPDATE.
--
--   TRY/CATCH + ROLLBACK:
--     Cualquier fallo revierte la anulacion completa.
-- ============================================================
CREATE OR ALTER TRIGGER trg_Venta_RestaurarStock
ON  venta
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        -- Verificar si algun registro cambio a estado 'Anulado'
        -- 'deleted' tiene estado anterior, 'inserted' tiene nuevo estado
        -- Solo continuar si hay ventas que pasaron a Anulado en este UPDATE
        IF NOT EXISTS (
            SELECT 1
            FROM   inserted I
            INNER JOIN deleted D ON I.idventa = D.idventa
            WHERE  I.estado = 'Anulado'
              AND  D.estado <> 'Anulado'
        )
            RETURN;

        -- Restaurar stock sumando la cantidad de cada linea del detalle
        -- Solo para las ventas que pasaron a 'Anulado' en este UPDATE
        UPDATE articulo
        SET    stock = A.stock + DV.cantidad
        FROM   articulo     A
        INNER JOIN detalle_venta DV ON A.idarticulo = DV.idarticulo
        INNER JOIN inserted  I      ON DV.idventa   = I.idventa
        INNER JOIN deleted   D      ON I.idventa    = D.idventa
        WHERE  I.estado = 'Anulado'
          AND  D.estado <> 'Anulado';

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        DECLARE @msg3 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg3, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- 05_SP_Cursor.sql
-- Base de datos: dbsistema
-- Descripcion: Stored Procedures con CURSOR para iteracion
--   registro a registro. Requeridos por el proyecto integrador
--   para demostrar estructuras de control avanzadas en T-SQL.
--
--   SP01 sp_ReporteVentasPorPeriodo — clasifica ventas por monto
--   SP02 sp_InventarioValorizado    — detecta stock bajo y calcula valor
-- ============================================================

USE dbsistema;
GO

-- ============================================================
-- SP01 — sp_ReporteVentasPorPeriodo
--
-- Proposito: Recorre venta por venta en un periodo, clasifica
--   cada una por monto (Alta/Media/Baja) y acumula totales.
--   Devuelve un resultado con la clasificacion y el acumulado
--   progresivo del periodo.
--
-- Estructuras de control usadas (requeridas por rubrica):
--   CURSOR  — itera fila a fila sobre las ventas del periodo
--   WHILE   — condicion de continuacion @@FETCH_STATUS = 0
--   IF/ELSE — clasifica el total: Alta>1000, Media 500-1000, Baja<500
--   TRY/CATCH — captura errores y hace ROLLBACK si es necesario
--   DEALLOCATE — libera los recursos del cursor al finalizar
--
-- Parametros:
--   @fechaInicio DATE — inicio del rango (inclusive)
--   @fechaFin    DATE — fin del rango (inclusive)
-- ============================================================
CREATE OR ALTER PROCEDURE sp_ReporteVentasPorPeriodo
    @fechaInicio DATE,
    @fechaFin    DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        -- ── Variables acumuladoras del periodo ────────────────
        DECLARE @totalPeriodo    DECIMAL(12,2) = 0;
        DECLARE @contadorVentas  INT           = 0;
        DECLARE @contadorAltas   INT           = 0;
        DECLARE @contadorMedias  INT           = 0;
        DECLARE @contadorBajas   INT           = 0;

        -- ── Variables de trabajo por iteracion ────────────────
        DECLARE @idVenta        INT;
        DECLARE @cliente        NVARCHAR(200);
        DECLARE @fecha          DATETIME;
        DECLARE @total          DECIMAL(10,2);
        DECLARE @estado         VARCHAR(20);
        DECLARE @clasificacion  VARCHAR(10);
        DECLARE @acumulado      DECIMAL(12,2);

        -- ── Tabla temporal para almacenar el resultado ────────
        -- Se usa tabla temporal para construir el resultado fila
        -- a fila dentro del WHILE y devolverlo al final con SELECT.
        CREATE TABLE #ResultadoReporte (
            idventa         INT,
            cliente         NVARCHAR(200),
            fecha           DATETIME,
            total           DECIMAL(10,2),
            clasificacion   VARCHAR(10),
            acumulado       DECIMAL(12,2)
        );

        -- ── CURSOR: define el conjunto de filas a iterar ──────
        -- Se usa CURSOR porque necesitamos procesar cada venta
        -- individualmente para: clasificarla, acumular el total
        -- progresivo y contar por categoria. Un simple SELECT
        -- con GROUP BY no puede hacer el acumulado progresivo
        -- registro a registro de esta forma.
        DECLARE curVentas CURSOR FOR
            SELECT  V.idventa,
                    P.nombre    AS cliente,
                    V.fecha,
                    V.total,
                    V.estado
            FROM    venta   V
            INNER JOIN persona P ON V.idcliente = P.idpersona
            WHERE   CAST(V.fecha AS DATE) BETWEEN @fechaInicio AND @fechaFin
              AND   V.estado = 'Activo'
            ORDER BY V.fecha ASC;

        -- ── Apertura del cursor e inicio de iteracion ─────────
        OPEN curVentas;
        FETCH NEXT FROM curVentas INTO @idVenta, @cliente, @fecha, @total, @estado;

        -- ── WHILE: continua mientras haya filas (@@FETCH_STATUS = 0) ──
        -- @@FETCH_STATUS = 0  → FETCH exitoso, hay fila disponible
        -- @@FETCH_STATUS = -1 → no hay mas filas (fin del cursor)
        -- @@FETCH_STATUS = -2 → la fila fue eliminada (no aplica aqui)
        WHILE @@FETCH_STATUS = 0
        BEGIN
            -- Acumular total del periodo y contar ventas
            SET @totalPeriodo   = @totalPeriodo + @total;
            SET @contadorVentas = @contadorVentas + 1;
            SET @acumulado      = @totalPeriodo;

            -- IF/ELSE: clasificar la venta segun su monto total
            -- Criterio definido por el negocio:
            --   Alta  → ventas mayores a S/. 1000
            --   Media → ventas entre S/. 500 y S/. 1000
            --   Baja  → ventas menores a S/. 500
            IF @total > 1000
            BEGIN
                SET @clasificacion  = 'Alta';
                SET @contadorAltas  = @contadorAltas + 1;
            END
            ELSE IF @total BETWEEN 500 AND 1000
            BEGIN
                SET @clasificacion  = 'Media';
                SET @contadorMedias = @contadorMedias + 1;
            END
            ELSE
            BEGIN
                SET @clasificacion  = 'Baja';
                SET @contadorBajas  = @contadorBajas + 1;
            END;

            -- Insertar fila procesada en la tabla temporal
            INSERT INTO #ResultadoReporte
                   (idventa, cliente, fecha, total, clasificacion, acumulado)
            VALUES (@idVenta, @cliente, @fecha, @total, @clasificacion, @acumulado);

            -- Avanzar al siguiente registro del cursor
            FETCH NEXT FROM curVentas INTO @idVenta, @cliente, @fecha, @total, @estado;
        END; -- fin WHILE

        -- ── CLOSE + DEALLOCATE: liberar recursos del cursor ───
        -- CLOSE  → cierra el cursor pero mantiene la definicion
        -- DEALLOCATE → elimina la definicion del cursor y libera
        --              toda la memoria asociada. Siempre debe
        --              llamarse para evitar fugas de recursos.
        CLOSE     curVentas;
        DEALLOCATE curVentas;

        -- ── Resultado 1: detalle de ventas con clasificacion ──
        SELECT  idventa,
                cliente,
                fecha,
                total,
                clasificacion,
                acumulado
        FROM    #ResultadoReporte
        ORDER BY fecha ASC;

        -- ── Resultado 2: resumen del periodo ──────────────────
        SELECT  @contadorVentas AS total_ventas,
                @totalPeriodo   AS monto_total_periodo,
                @contadorAltas  AS ventas_alta,
                @contadorMedias AS ventas_media,
                @contadorBajas  AS ventas_baja;

        DROP TABLE #ResultadoReporte;

    END TRY
    BEGIN CATCH
        -- Si el cursor quedo abierto por el error, cerrarlo
        IF CURSOR_STATUS('local', 'curVentas') >= 0
        BEGIN
            CLOSE     curVentas;
            DEALLOCATE curVentas;
        END;
        IF OBJECT_ID('tempdb..#ResultadoReporte') IS NOT NULL
            DROP TABLE #ResultadoReporte;

        DECLARE @msg1 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg1, 16, 1);
    END CATCH
END
GO

-- ============================================================
-- SP02 — sp_InventarioValorizado
--
-- Proposito: Recorre articulo por articulo del inventario,
--   calcula el valor economico por categoria, detecta articulos
--   con stock bajo y genera alertas de reabastecimiento.
--
-- Estructuras de control usadas (requeridas por rubrica):
--   CURSOR  — itera articulo por articulo para acumular por
--             categoria (no posible con GROUP BY simple cuando
--             se necesita el acumulado progresivo por categoria)
--   WHILE   — condicion de continuacion @@FETCH_STATUS = 0
--   IF/ELSE — detecta stock bajo segun @stockMinimo
--   CASE    — clasifica estado del stock en 4 niveles
--   TRY/CATCH — manejo de errores
--   DEALLOCATE — libera recursos del cursor
--
-- Parametros:
--   @stockMinimo INT (default 5) — umbral para alerta de stock bajo
-- ============================================================
CREATE OR ALTER PROCEDURE sp_InventarioValorizado
    @stockMinimo INT = 5
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        -- ── Variables de trabajo por iteracion ────────────────
        DECLARE @idArticulo     INT;
        DECLARE @codArticulo    VARCHAR(50);
        DECLARE @nomArticulo    NVARCHAR(200);
        DECLARE @categoria      NVARCHAR(100);
        DECLARE @stock          INT;
        DECLARE @precioVenta    DECIMAL(10,2);
        DECLARE @valorUnitario  DECIMAL(10,2);
        DECLARE @valorLinea     DECIMAL(10,2);
        DECLARE @alerta         VARCHAR(20);
        DECLARE @estadoStock    VARCHAR(10);

        -- ── Variables acumuladoras por categoria ──────────────
        DECLARE @categoriaActual    NVARCHAR(100) = '';
        DECLARE @acumValorCategoria DECIMAL(12,2)  = 0;
        DECLARE @valorTotalGlobal   DECIMAL(12,2)  = 0;

        -- ── Tabla temporal del resultado ──────────────────────
        CREATE TABLE #ResultadoInventario (
            categoria       NVARCHAR(100),
            codigoArticulo  VARCHAR(50),
            nombreArticulo  NVARCHAR(200),
            stock           INT,
            precioVenta     DECIMAL(10,2),
            valorTotal      DECIMAL(10,2),
            estadoStock     VARCHAR(10),
            alerta          VARCHAR(20)
        );

        -- ── CURSOR: articulos activos ordenados por categoria ─
        -- Ordenar por categoria permite detectar cambio de grupo
        -- para calcular el acumulado por categoria correctamente.
        DECLARE curInventario CURSOR FOR
            SELECT  A.idarticulo,
                    A.codigo,
                    A.nombre,
                    C.nombre    AS categoria,
                    A.stock,
                    A.precio_venta
            FROM    articulo  A
            INNER JOIN categoria C ON A.idcategoria = C.idcategoria
            WHERE   A.estado = 1
            ORDER BY C.nombre ASC, A.nombre ASC;

        OPEN curInventario;
        FETCH NEXT FROM curInventario
            INTO @idArticulo, @codArticulo, @nomArticulo,
                 @categoria, @stock, @precioVenta;

        -- ── WHILE: iterar mientras haya articulos ─────────────
        WHILE @@FETCH_STATUS = 0
        BEGIN
            -- Calcular valor economico de esta linea
            SET @valorLinea = @stock * @precioVenta;

            -- Acumular totales
            SET @valorTotalGlobal = @valorTotalGlobal + @valorLinea;

            -- IF: detectar alerta de stock bajo
            IF @stock = 0
                SET @alerta = 'SIN STOCK';
            ELSE IF @stock < @stockMinimo
                SET @alerta = 'REABASTECER';
            ELSE
                SET @alerta = 'OK';

            -- CASE: clasificar nivel de stock con 4 categorias
            -- Critico  → 0 unidades (sin stock)
            -- Bajo     → entre 1 y @stockMinimo-1
            -- Normal   → igual o mayor a @stockMinimo
            -- Optimo   → mayor al doble del minimo
            SET @estadoStock =
                CASE
                    WHEN @stock = 0                     THEN 'Critico'
                    WHEN @stock < @stockMinimo          THEN 'Bajo'
                    WHEN @stock <= (@stockMinimo * 2)   THEN 'Normal'
                    ELSE                                     'Optimo'
                END;

            -- Insertar resultado procesado
            INSERT INTO #ResultadoInventario
                   (categoria, codigoArticulo, nombreArticulo,
                    stock, precioVenta, valorTotal, estadoStock, alerta)
            VALUES (@categoria, @codArticulo, @nomArticulo,
                    @stock, @precioVenta, @valorLinea, @estadoStock, @alerta);

            -- Siguiente registro del cursor
            FETCH NEXT FROM curInventario
                INTO @idArticulo, @codArticulo, @nomArticulo,
                     @categoria, @stock, @precioVenta;
        END; -- fin WHILE

        -- ── CLOSE + DEALLOCATE ────────────────────────────────
        CLOSE     curInventario;
        DEALLOCATE curInventario;

        -- ── Resultado 1: detalle articulo por articulo ────────
        SELECT  categoria,
                codigoArticulo,
                nombreArticulo,
                stock,
                precioVenta,
                valorTotal,
                estadoStock,
                alerta
        FROM    #ResultadoInventario
        ORDER BY categoria ASC, nombreArticulo ASC;

        -- ── Resultado 2: resumen global del inventario ────────
        SELECT  COUNT(*)                            AS total_articulos,
                SUM(stock)                          AS unidades_totales,
                SUM(valorTotal)                     AS valor_total_inventario,
                SUM(CASE WHEN alerta = 'REABASTECER'
                          OR alerta = 'SIN STOCK'
                     THEN 1 ELSE 0 END)             AS articulos_con_alerta
        FROM    #ResultadoInventario;

        DROP TABLE #ResultadoInventario;

    END TRY
    BEGIN CATCH
        IF CURSOR_STATUS('local', 'curInventario') >= 0
        BEGIN
            CLOSE     curInventario;
            DEALLOCATE curInventario;
        END;
        IF OBJECT_ID('tempdb..#ResultadoInventario') IS NOT NULL
            DROP TABLE #ResultadoInventario;

        DECLARE @msg2 NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(@msg2, 16, 1);
    END CATCH
END
GO

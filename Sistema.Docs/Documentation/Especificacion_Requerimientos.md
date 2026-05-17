# Especificación de Requerimientos de Software (ERS)

> **Nota:** De acuerdo a Roger S. Pressman, una Especificación de Requerimientos de Software (ERS) es un documento que contiene una descripción detallada de todas las funciones del software que se va a desarrollar o analizar, estableciendo las bases del comportamiento del sistema y la interacción con sus usuarios.

## 1. Descripción General del Proyecto

### 1.1 Antecedentes
En el ámbito comercial y de administración de negocios, el control manual de inventarios, compras y acceso de personal suele presentar retos importantes, tales como la pérdida de información, errores humanos en el conteo de artículos, y falta de restricciones de seguridad. Los sistemas de información empresariales surgen como respuesta a la necesidad de controlar de forma automatizada y precisa los flujos de un almacén o punto de venta.

### 1.2 Motivación
La principal motivación para este proyecto es proporcionar a la empresa o negocio una herramienta robusta, centralizada y fácil de usar que reduzca los tiempos en la administración de productos (artículos), el registro de ingresos y la gestión del personal y clientes.

### 1.3 Justificación
El desarrollo e implementación de este sistema se justifica por la disminución de pérdidas económicas derivadas de inventarios mal calculados o no actualizados en tiempo real. Un sistema automatizado proporciona una fuente de verdad única de la cual la gerencia puede extraer reportes, analizar ventas, supervisar ingresos de mercancía por proveedor y controlar los niveles de acceso de sus empleados (Roles), permitiendo una toma de decisiones más rápida y un servicio al cliente más ágil.

### 1.4 Objetivos
* **Objetivo General:** Administrar de manera eficiente e integral los módulos de inventario, compras (ingresos) y usuarios de un negocio a través de una aplicación de escritorio conectada a un origen de datos central.
* **Objetivos Específicos:**
  * Implementar control de inventarios mediante altas, bajas y modificaciones de Artículos y Categorías.
  * Gestionar los accesos de los empleados mediante un sistema de Roles de seguridad y autenticación (Login).
  * Registrar correctamente los movimientos y compras, vinculando a los proveedores con los productos que ingresan al almacén.

### 1.5 Hipótesis
La implementación de un sistema de escritorio bajo una arquitectura de capas centralizará la operativa administrativa del negocio, reduciendo en un porcentaje significativo el tiempo invertido en auditorías de inventario manuales y minimizando los errores de captura por parte del personal de ventas y almacén. 

### 1.6 Alcances
El sistema contempla las siguientes áreas o módulos funcionales operativos:
* **Acceso y Seguridad:** Autenticación de usuarios y asignación de roles.
* **Módulo de Inventario:** Catálogo de artículos (con imagen, precios, stock y códigos de barras/sku) y categorías.
* **Módulo de Compras:** Catálogo de proveedores y registro de ingresos de mercancía que impactan directamente el stock.
* **Módulo de Terceros:** Catálogo de clientes.

### 1.7 Beneficios
* **Control y Trazabilidad:** El administrador puede conocer el stock real en cualquier momento.
* **Seguridad:** Cada usuario tiene credenciales, permitiendo restringir o auditar las acciones en el sistema.
* **Escalabilidad:** Al estar diseñado en un modelo N-Capas, se facilita el mantenimiento o la posterior ampliación modular.

---

## 2. Tecnologías y Entorno de Desarrollo

Este proyecto está construido bajo una arquitectura robusta de **Capas (N-Tier)**, dividiendo la lógica en Presentación, Negocio, Datos y Entidades. Las tecnologías consideradas para su construcción y ejecución son las siguientes:

* **Lenguaje de Programación:** Visual Basic (.vb).
* **Framework:** .NET Framework 4.7.2 (lo que lo orienta a aplicaciones de entorno Windows clásico).
* **Tipo de Aplicación:** Windows Forms App (WinForms).
* **Diseño y UI:** Controles visuales estándar de .NET más la implementación de iconos mediante la librería `FontAwesome.Sharp`.
* **Base de Datos:** Debido a la arquitectura empresarial evidenciada en las capas `Sistema.Datos` y `Sistema.Entidades`, el sistema requiere conectar con un servidor de Bases de Datos Relacional, típicamente **Microsoft SQL Server**, a través de la tecnología ADO.NET (o similar).
* **Entorno de Desarrollo (IDE):** Microsoft Visual Studio (Community/Professional/Enterprise).

---

## 3. Requerimientos Funcionales y Verificación de Operaciones (CRUD)

### 3.1 Verificación de Funciones Principales (Justificación)
Tras analizar el código fuente del repositorio (específicamente la capa de presentación y sus eventos), he comprobado que el sistema **SÍ CUMPLE** rigurosamente con al menos las 4 funciones principales requeridas en las entidades principales (Alta, Consulta, Modificación y Eliminación), formando un CRUD completo. 

A continuación se justifica tomando como evidencia la funcionalidad de `FrmArticulo.vb`:

1. **Consulta (Read):**  
   Se encuentra justificado en los métodos `Listar()` y `Buscar()`. El sistema recupera la información de los artículos desde la capa de negocio (`Neg.Listar()`) y los plasma en el componente `DgvListado` (DataGridView), permitiendo al usuario visualizar el catálogo actual y aplicar filtros a la información mostrada.

2. **Alta / Registro (Create):**  
   Se justifica en el evento `BtnInsertar_Click`. El sistema valida que los campos requeridos (código, precio, nombre, etc.) no estén vacíos, instancia la entidad `Entidades.Articulo`, la empaqueta y luego invoca al método `Neg.Insertar(Obj)`. Adicionalmente, incluye lógica para transferir y guardar permanentemente la imagen asociada al producto de forma física.

3. **Actualización / Modificación (Update):**  
   Implementado y justificado en el evento `BtnActualizar_Click`. Cuando un usuario hace *doble clic* en la tabla, el formulario se rellena. Permite alterar precios, stock, imágenes y presionar "Actualizar", enviando el ID inmutable (`Obj.IdArticulo = TxtId.Text`) al servicio `Neg.Actualizar(Obj)` para reflejar un `UPDATE` controlado en la BD.

4. **Eliminación y Desactivación (Delete):**  
   Implementado y justificado en el evento `BtnEliminar_Click`. El sistema incluye checkboxes para realizar operaciones por lotes. Llama a la instrucción `Neg.Eliminar(OneKey)` para eliminar directamente en la base de datos y también emplea `File.Delete()` para eliminar la imagen física del disco local. Además de la eliminación, ofrece funciones de *Soft Delete* (Baja lógica) mediante los botones de **Activar** y **Desactivar**.

### 3.2 Tabla de Requerimientos Funcionales

| Módulo | Id | Componente | Requerimiento | Requerimiento Descripción |
| :--- | :---: | :--- | :--- | :--- |
| **Acceso** | 1.0 | Login | Ingresar al sistema | Ingresar al sistema con una cuenta o correo como usuario y su respectiva contraseña. |
| | 1.1 | | Restricción de usuarios | No permitir el ingreso de usuarios no autorizados, inactivos o con credenciales erróneas. |
| **Configuración** | 2.0 | Inicio y Menú | Limitar la información | Se visualiza la información necesaria y módulos disponibles para el usuario dependiendo del rol asignado. |
| **Configuración** | 3.0 | Módulo Usuarios | Mostrar Usuarios | Mostrar datos relevantes de usuarios y el tipo/nombre de rol asociado en un listado (DataGridView). |
| | 3.1 | | Registrar usuarios | Permite el alta (ingreso) de nuevos usuarios al sistema y asignarles contraseñas y permisos. |
| | 3.2 | | Editar y Estado | Se permite la modificación de sus datos, así como la posibilidad de "Desactivar" al usuario para impedir su acceso sin tener que eliminar su registro histórico. |
| **Inventario** | 4.0 | Módulo Artículos | Gestión de catálogo | Permitir el registro de nombre, descripción, stock base, código de barras, precio y cargar una fotografía representativa del artículo. |
| | 4.1 | | Búsqueda dinámica | Contar con un componente de búsqueda rápida para filtrar artículos por coincidencia de nombre o código. |
| | 4.2 | | Gestión de Categorías | Poder crear, leer, actualizar y eliminar las distintas categorías o departamentos para asociar a los artículos. |
| **Operativa** | 5.0 | Módulo Ingresos | Compras a Proveedores | Registrar ingresos de mercancía por un comprobante, relacionando un proveedor, incrementando automáticamente el Stock de los artículos ingresados. |
| | 5.1 | | Módulo Clientes | Disponer de un directorio mantenible (Alta/Baja/Cambio) con la lista de clientes para ser usados en procesos de facturación o ventas posteriores. |

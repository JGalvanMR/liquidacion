# 📦 Nombre del Módulo: **liquidacion** (repositorio `JGalvanMR/liquidacion`, rama `master`)

## 🧭 Propósito

`liquidacion` es una aplicación de escritorio Windows Forms (.NET Framework 4.0, "SISEMP") que forma parte de la suite SisGabWeb. Su propósito es calcular y generar la **liquidación de pago a proveedores de producto agrícola**, a partir de dos posibles orígenes de mercancía —recepción de producto terminado comprado (PTC) o producción propia (PRO)— separando siempre las ventas en **Nacional** y **Exportación**, aplicando conceptos de flete, préstamos, merma, comisiones y descuentos, y dejando registro contable/operativo de cada liquidación generada, modificada o cancelada.

## ⚙️ Responsabilidades

- Validar que exista una sesión de usuario activa (iniciada por la suite SisGabWeb) antes de permitir el uso de la aplicación.
- Calcular, para un proveedor y rango de fechas dado, las unidades vendidas Nacional vs Exportación por producto, a partir de facturación (`tb_det_facturas`, `tb_mstr_facturas_nal`).
- Calcular las unidades recibidas pendientes de liquidar por producto, a partir de recepciones (`tb_hist_recepcion`, `tb_mstr_recepcion_pt`).
- Presentar una grilla de liquidación con indicadores visuales (verde/rojo) de coincidencia entre lo recibido y lo liquidado.
- Permitir capturar el número de cajas a liquidar mediante un diálogo modal (`cajas`).
- Abrir la pantalla de liquidación preliminar (`preliminar`) donde se capturan los conceptos de la liquidación (flete, préstamo, merma, comisión, descuento, etc.), se valida el tipo de cambio del día y se calculan los totales.
- Persistir la liquidación (nueva, modificación o cancelación) en las tablas maestro/detalle correspondientes y actualizar el estado de las recepciones y órdenes de compra relacionadas.
- Consultar y conciliar notas de crédito/cargo (nacionales y de exportación, incluyendo merma) mediante procedimientos almacenados (`Notas_Credito_Cargo`, `ConsultaController`/`ConsultaModels`).
- Mostrar anticipos/préstamos (`anticipos`) y liquidaciones anticipadas/parciales (`articipadas`) asociadas a un proveedor.
- Mostrar desgloses de detalle de precio por factura (`DetallePrecio`) y por concepto/caja (`detalle`).
- Registrar accesos y errores en bitácora local y enviar alertas por correo ante fallas de sistema.

## 🔄 Flujo de Funcionamiento

1. **Arranque (`Program.cs`)**: se valida la IP de la máquina (`Utilerias.Class1.validar_ip()`) y se ejecuta `validar_login()`, que busca en `tb_cat_historial_dia` una sesión abierta (`fin_sesion IS NULL`) para el nombre de máquina actual. Si no existe, se muestra un aviso, se lanza `SISEMP.exe` (aplicación central de la suite) y la aplicación actual se cierra. Si existe, se abre `Form1`.
2. **Pantalla principal (`Form1`)**: el usuario captura clave de proveedor, rango de fechas, tipo de liquidación ("Recepción PT" o "Producción") y, opcionalmente, rango de líneas/empaques. Al presionar generar (`btnGenera_Click`):
   - Se determina si el cálculo de fletes del periodo está actualizado (`ultima_fecha_calculada`).
   - Según el tipo seleccionado se ejecuta `producto_terminado()` (índice 0) o `produccion()` (índice 1), que consultan ventas nacionales/exportación y unidades recibidas, calculan porcentaje de venta Nacional/Exportación por producto y llenan la grilla `dtgLiquidacion`.
3. **Selección de producto en la grilla**: al hacer doble clic en la columna Nacional (índice 4) o Exportación (índice 5), se valida la **regla del 65%** (ver Reglas de Negocio) y se pregunta si se desea capturar el número de cajas (diálogo `cajas`). Con la cantidad definida, se abre el formulario `preliminar`, pasando todos los datos del producto/proveedor/periodo.
4. **Liquidación preliminar (`preliminar`)**: valida que exista tipo de cambio del día (`tb_cat_tipocambio`); permite capturar conceptos (flete, préstamo, merma, comisión, descuento, etc.) indicando unidades y precio; calcula el importe de cada concepto (suma o resta según el código de concepto); calcula totales; y, al guardar, ejecuta el flujo correspondiente (`guardanuevonal`, `guardanuevoexp`, `guardafaltantenal`, `guardafaltanteexp`, `modificarnacional`, `modificarexportacion`) tras validar cajas/palet, flejes/palet, totales e importes.
5. **Persistencia**: se inserta/actualiza en `tb_mstr_liquidacion`, `tb_det_liquidacion`, `tb_det_liquidacion_rec`, `tb_det_prestamo`, `tb_det_liq_planta`, `tb_det_anticipada_pt`, y se actualizan `tb_hist_recepcion`, `tb_mstr_ordencompra`, `tb_det_ordenescompra`, `tb_det_notascyc` según corresponda.
6. **Consulta de notas de crédito/cargo**: desde `Form1.btnImpuesto_Click` se abre `Notas_Credito_Cargo`, que usa `ConsultaController`/`ConsultaModels` (basados en procedimientos almacenados parametrizados) para traer notas nacionales, de exportación y de merma no liquidadas, filtradas por línea de producto.
7. **Cancelación**: desde `preliminar.btnCancela_Click`, se valida en `tb_mstr_liquidacion` que la liquidación no tenga orden de compra asociada, no esté ya cancelada y no haya afectado costos, antes de proceder a cancelarla.

## 📐 Reglas de Negocio

### 🔒 Restricciones
- **Umbral del 65%**: un producto solo puede liquidarse en su columna Nacional o Exportación si el porcentaje de venta correspondiente (`por`) es **mayor o igual a 65%**; de lo contrario se bloquea la liquidación y se muestra un aviso con el detalle de cajas nacionales/exportación.
- **Excepciones al umbral del 65%**: la validación se omite si la clave de proveedor es `01`, `03` o `1328`, o si la línea del producto (`ln`) es `12` o `19`.
- **Bandera de bloqueo activa (`lblBloqueo == "1"`, proveniente de `tb_liquidacion_bloqueo`)**: la validación del 65% solo se ejecuta cuando esta bandera está activa; no es determinable con la información disponible qué condición de negocio activa/desactiva dicha bandera.
- **Cancelación de liquidación restringida**: no se puede cancelar una liquidación si ya tiene número de orden de compra asociado (`liq_numoc1`), si su estatus ya es `"C"` (cancelada), o si `liq_afecto == "1"` (ya afectó costos).
- **Acceso condicionado a sesión externa**: la aplicación no permite operar si no existe una sesión de login activa registrada por la suite SisGabWeb para la máquina actual.

### ✅ Validaciones
- Debe seleccionarse el tipo de liquidación ("Recepción PT" / "Producción") antes de generar.
- Debe seleccionarse un rango de fechas antes de generar.
- El formato de fecha capturado debe ser válido.
- Al capturar un concepto en `preliminar`: debe seleccionarse un concepto, capturarse unidades, capturarse un precio, el precio debe ser numérico y mayor a cero.
- Antes de guardar una liquidación nueva/modificada: el valor de cajas por palet debe ser mayor a 0, el valor de flejes por palet debe ser mayor a 0, el total/monto a liquidar/costo unitario deben ser mayores a 0, y debe existir al menos un concepto de liquidación capturado.
- Debe existir un tipo de cambio registrado para el día actual (`tb_cat_tipocambio`) antes de continuar en `preliminar`; si no existe, se advierte al usuario verificar con el encargado.
- Debe existir el número de tarimas y el número de cajas/flejes registrado para el producto antes de continuar ciertos cálculos.

### 🔁 Agrupaciones
- Las ventas facturadas se agrupan por `lin_clave` + `prod_clave`, sumando unidades, separando **Nacional** (moneda `PESOS`) de **Exportación** (moneda `USD`).
- Las recepciones (`tb_hist_recepcion`) se agrupan por producto, línea, tipo de recepción, número de liquidación y proveedor, para determinar unidades recibidas pendientes de liquidar.
- Los códigos de concepto se agrupan en dos comportamientos de cálculo: los códigos `93, 95, 100, 102, 103, 104, 105, 106` se **suman** al total (créditos/adiciones); todos los demás códigos de concepto se **restan** del total (deducciones).

### ⚙️ Reglas Operativas
- El tipo de liquidación determina el flujo de cálculo: "Recepción PT" ejecuta `producto_terminado()`; "Producción" ejecuta `produccion()`.
- La procedencia (`NACIONAL` / `EXPORTACION`) determina la columna de la grilla utilizada y el método de guardado invocado (`guardanuevonal`/`guardanuevoexp`/`modificarnacional`/`modificarexportacion`).
- Si la fecha del último cálculo de fletes (`tb_registro_movimientos`, `tipo_mov = 'F'`) es anterior a la fecha final solicitada, se advierte al usuario que el cálculo de fletes no está actualizado para el periodo, pero el flujo puede continuar (bandera `nanana`), repitiendo la advertencia al interactuar con la grilla.
- La conversión por tipo de cambio se aplica únicamente a los conceptos que **no** están en la lista de códigos aditivos (93/95/100/102/103/104/105/106).
- Si el folio de liquidación cambia por concurrencia de red al momento de guardar, el sistema lo detecta, informa al usuario del nuevo folio asignado y reimprime automáticamente con dicho folio.
- Ante error de guardado (`SqlException`), se registra el error mediante `Utilerias.Class1.registro_errores` y se envía una alerta por correo a `aescamilla@mrlucky.com.mx`.

## 🔗 Dependencias

- **Utilerias.dll** (referenciada por ruta absoluta `C:\SisGabWeb\Utilerias.dll`): expone `ConnectionString`, estado de sesión (`Login`, `Usu_login`, `Grupo`), `validar_ip()`, `registro_errores()` y `SendMail()`.
- **System.Data.SqlClient** (ADO.NET) para acceso directo a SQL Server.
- **7 procedimientos almacenados** consumidos vía `ConsultaModels`/`ConsultaController`: `spSISEMPLiquidacionesNotasCreditoCargoNal`, `...Exp`, `...MermaNal`, `...MermaExp`, `spSISEMPLiquidacionesProductos`, `spSISEMPLiquidacionesNCRNCGExp`, `spSISEMPLiquidacionesNCRNCGNal`.
- **Tablas SQL Server principales**: `tb_cat_historial_dia`, `tb_cat_linea`, `tb_cat_producto`, `tb_liquidacion_bloqueo`, `tb_det_facturas`, `tb_mstr_facturas_nal`, `tb_hist_recepcion`, `tb_mstr_recepcion_pt`, `tb_cat_tipocambio`, `tb_registro_movimientos`, `tb_mstr_liquidacion`, `tb_det_liquidacion`, `tb_det_liquidacion_rec`, `tb_det_prestamo`, `tb_det_liq_planta`, `tb_det_anticipada_pt`, `tb_det_notascyc`, `tb_mstr_ordencompra`, `tb_det_ordenescompra`, `Tb_Prestamos_Prov`.
- **Sistema de archivos local**: `C:\SisGabWeb\fondo_formularios.jpg` (imagen de fondo de formularios), `C:\SisEmpWeb\eventlog.txt` (bitácora de acceso), `C:\SisGabWeb\SISEMP.exe` (aplicación externa de la suite requerida para autenticación).
- **SMTP** vía `Utilerias.Class1.SendMail`, con destinatario fijo `aescamilla@mrlucky.com.mx`.
- Formularios modales internos con paso de datos vía estado estático: `cajas.SharedData`, `calendario.SharedData`.

## ⚠️ Riesgos Técnicos

- **Inyección SQL generalizada**: la gran mayoría de las consultas en `Form1.cs`, `preliminar.cs` y `DetallePrecio.cs` se construyen concatenando directamente valores de controles de UI (fechas, claves de proveedor/línea, texto) sin parametrización. Solo `ConsultaModels.cs` usa procedimientos almacenados con parámetros (`SqlParameter`).
- **Acoplamiento fuerte a rutas absolutas de un solo servidor** (`C:\SisGabWeb\...`, `C:\SisEmpWeb\...`), lo que impide portabilidad o despliegue en otro entorno sin modificar código.
- **Dependencia crítica de una aplicación externa (SISEMP.exe/suite SisGabWeb) para autenticación**: `liquidacion.exe` no gestiona su propio login; delega en una sesión ya abierta registrada en `tb_cat_historial_dia` y en `Environment.MachineName`, lo cual es frágil si varios usuarios comparten máquina o quedan sesiones sin `fin_sesion` registrado.
- **Reglas de negocio críticas codificadas como literales dispersos** ("magic numbers/strings"): el umbral del 65%, las claves de proveedor exentas (`01`, `03`, `1328`), las líneas exentas (`12`, `19`) y los códigos de concepto aditivos (`93/95/100/102/103/104/105/106`) están hardcodeados dentro de la lógica de UI en vez de un catálogo configurable, dificultando su mantenimiento y auditoría.
- **Credenciales/destinatario de alerta por correo repetidos como literales** (`aescamilla@mrlucky.com.mx`, `atrejo`) en decenas de puntos del código de `preliminar.cs`.
- **Archivo/formulario `preliminar.cs` de ~11,600 líneas** que concentra validación, cálculo, persistencia y presentación en una sola clase, lo que dificulta pruebas unitarias, mantenimiento y aislamiento de errores (alto acoplamiento / "God Object").
- **Manejo de errores basado en captura genérica de `SqlException`** con el mensaje uniforme "Error de sistema", sin garantizar una transacción explícita entre las múltiples operaciones de inserción/actualización relacionadas (`tb_mstr_liquidacion`, `tb_det_liquidacion`, `tb_hist_recepcion`, `tb_det_prestamo`, etc.), lo que puede dejar datos parcialmente escritos si falla una operación intermedia.
- **Estado global mutable vía patrón estático `SharedData`** (`cajas.SharedData.Polino`, `calendario.SharedData.Polino`) para pasar datos entre formularios modales, con riesgo de condiciones de carrera o datos obsoletos si se reabren diálogos.
- **Referencia a `Utilerias.dll` mediante `HintPath` absoluto** en lugar de gestión de paquetes/versión, lo que puede provocar builds no reproducibles fuera de esa máquina específica.

## 🧪 Casos Edge

- Cambio de folio de liquidación por concurrencia de red al guardar: se detecta y se informa al usuario, pero no es determinable con la información disponible qué ocurre si el conflicto se presenta durante la escritura real en base de datos, más allá del mensaje mostrado.
- Producto con 0% de venta en uno de los dos canales (Nacional=0 con Exportación>0, o viceversa): se maneja en el coloreado de la grilla (verde/rojo), pero no es determinable con la información disponible si en ese caso mixto aplica o no la validación del 65%.
- Proveedor exento (`01`, `03`, `1328`) o línea exenta (`12`, `19`): se omite completamente la validación del umbral del 65%, permitiendo liquidar independientemente del porcentaje de venta real.
- `lblBloqueo` distinto de `"1"`: se omite por completo la validación del 65%; no es determinable con la información disponible bajo qué condiciones de negocio se activa o desactiva dicho bloqueo.
- Guardado cuando el tipo de cambio o la respuesta de tipo de cambio están vacíos: se omite la conversión de moneda extranjera y el concepto se guarda con el precio en su moneda original tal como fue capturado.

## 🧱 Suposiciones Detectadas

- Se asume que la sesión de usuario ya fue iniciada por otra aplicación de la suite (SISEMP/SisGabWeb) antes de abrir `liquidacion.exe`.
- Se asume una única instancia de SQL Server accesible vía `Utilerias.Class1.ConnectionString`, compartida por toda la suite de aplicaciones.
- Se asume que las rutas de imágenes y bitácoras (`C:\SisGabWeb\...`, `C:\SisEmpWeb\...`) existen en la máquina cliente donde se ejecuta la aplicación.
- Se asume que el umbral del 65% y las claves de proveedor/línea exentas (`01`, `03`, `1328`, `12`, `19`) permanecerán estables como reglas de negocio; no hay evidencia en el código de que sean configurables desde la aplicación misma.
- Se asume que el usuario que captura importes y porcentajes en `preliminar` los introduce correctamente, ya que la única validación aplicada es "es numérico" y "mayor a cero", sin validar rangos máximos razonables.

## 📈 Recomendaciones Técnicas

- Migrar todas las consultas construidas por concatenación de cadenas a comandos parametrizados (`SqlParameter`) o procedimientos almacenados, replicando el patrón ya existente en `ConsultaModels.cs`.
- Extraer las reglas de negocio codificadas como literales (umbral 65%, claves exentas `01/03/1328`, líneas exentas `12/19`, códigos de concepto aditivos `93/95/100/102/103/104/105/106`) a una tabla de configuración/catálogo consultable y editable sin recompilar la aplicación.
- Refactorizar `preliminar.cs` separando responsabilidades en capas (UI / lógica de cálculo / acceso a datos), reduciendo el tamaño del archivo y habilitando pruebas unitarias sobre el cálculo de la liquidación.
- Envolver las operaciones de guardado que afectan múltiples tablas (`tb_mstr_liquidacion`, `tb_det_liquidacion`, `tb_hist_recepcion`, `tb_det_prestamo`, etc.) en una transacción SQL explícita (`SqlTransaction`) para garantizar atomicidad.
- Externalizar el destinatario y credenciales de alerta por correo a configuración segura, en lugar de literales repetidos en el código fuente.
- Sustituir el patrón estático `SharedData` por el paso de resultados vía propiedades públicas del diálogo modal (combinado con `DialogResult`), evitando estado global compartido.
- Documentar y, de ser posible, desacoplar la lógica de sesión/login (`validar_login` en `Program.cs`) de la aplicación externa `SISEMP.exe`, reduciendo el acoplamiento entre aplicaciones de la suite.

## 🧾 Resumen Ejecutivo

Este sistema es la herramienta con la que el área administrativa calcula cuánto debe pagarse a cada proveedor de producto agrícola por lo que entregó en un periodo determinado. Antes de autorizar el pago, el sistema exige que al menos el 65% de lo vendido de ese producto se haya vendido por el canal correspondiente (nacional o exportación), salvo un pequeño grupo de proveedores y líneas de producto que están exentos de esa regla. Una vez validado, el usuario puede agregar conceptos como flete, préstamos, merma o comisiones, y el sistema calcula el monto final a liquidar, considerando el tipo de cambio del día cuando aplica. El resultado se guarda formalmente y puede modificarse o cancelarse bajo ciertas condiciones (por ejemplo, no se puede cancelar una liquidación que ya generó una orden de compra o que ya impactó los costos de la empresa). El sistema también permite conciliar notas de crédito y cargo pendientes antes de cerrar el proceso. Desde el punto de vista de riesgo operativo, la aplicación depende de que otra aplicación de la suite ya haya iniciado la sesión del usuario, y varias de sus reglas de negocio (como el umbral del 65% o las excepciones por proveedor) están fijas en el código, por lo que cualquier cambio en esas políticas de negocio requeriría modificar y volver a desplegar la aplicación.

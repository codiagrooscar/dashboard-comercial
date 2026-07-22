# Controller Comercial

Aplicación local para generar el informe diario de ventas a partir de cuatro archivos Excel:

- `ofertas.xlsx`
- `pedidos.xlsx`
- `albaranes.xlsx`
- `facturas.xlsx`

## Cómo abrirla

Desde esta carpeta, ejecutar:

```powershell
& 'C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' app.py
```

Después abrir:

```text
http://127.0.0.1:8765
```

La app permite subir los cuatro archivos o usar directamente los archivos actuales de la carpeta `Estado de pedidos`.

## Reglas aplicadas

- El día analizado se calcula como el día anterior al día real de ejecución. Por ejemplo, si hoy es 20/05, analiza 19/05; si mañana es 21/05, analiza 20/05.
- Nacional/Extranjero se clasifica por la serie del pedido: si contiene `EX`, se trata como Extranjero.
- Las ofertas se consideran convertidas cuando existe pedido coincidente por cliente y artículo cuando esas columnas están disponibles.
- `pedidos.xlsx` se considera la cartera pendiente de albarán. El estado de pedidos usa `UnidadesServidas` y `UnidadesPendientes`; de momento no usa la columna `Estado`.
- No se muestran estados `Enviados/Entregados` ni `Completados` en pedidos porque no aplican al listado de cartera pendiente.
- Para la previsión de cierre, si un pedido no tiene `FechaNecesaria`, se supone que puede fabricarse/cargarse 7 días después de `FechaPedido`. También se incluye backlog reciente si sigue pendiente y su fecha real o estimada de disponibilidad cae antes del fin del mes actual, excluyendo pedidos con más de un mes de antigüedad.
- `albaranes.xlsx` se considera la lista de albaranes pendientes de facturar; no se descuentan por coincidencias indirectas con facturas del mismo cliente.
- Las ofertas aprobadas se identifican por respaldo en pedidos y se muestran con fecha teórica de entrega `FechaOferta + 15 días`, sin sumarse a la previsión para evitar duplicidad.
- La previsión compara la facturación real y esperada contra un presupuesto mensual de `1.636.909 EUR`.
- La auditoría revisa pedidos sin `Fecha Necesaria`, albaranes con más de 7 días sin factura localizada y ofertas de alto valor con más de 10 días sin conversión.

Nota: los archivos de albaranes y facturas revisados no incluyen artículo ni referencia directa de pedido, por lo que esos cruces se hacen por cliente y fecha.

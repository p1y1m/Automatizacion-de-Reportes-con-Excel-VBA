# Automatización de Reportes con Excel y VBA

Este proyecto presenta una solución desarrollada en VBA para Excel que permite automatizar la generación y envío de reportes operativos a más de 100 sucursales bancarias. El objetivo principal fue reducir drásticamente el tiempo requerido en comparación con el proceso manual.

## Principales Logros
- Reducción del tiempo de 32 horas hombre (8 personas x 4 horas) a solo 2 horas de ejecución automática.
- Envío automatizado de 108 correos personalizados a sucursales.
- Consolidación de información proveniente de más de 25 hojas de Excel.
- Inclusión automática de destinatarios (jefes, supervisores y zonales) según la sucursal.

## Descripción del Funcionamiento
1. Se recorren los códigos de todas las sucursales desde un archivo central.
2. Se obtienen automáticamente los correos electrónicos asociados desde la hoja 'Datos Sucursales'.
3. Se redacta el cuerpo del correo, incluyendo mensajes estandarizados.
4. Se recorren más de 25 hojas de cálculo en los libros 'Excelencia Operativa' y 'Eficiencia en Costos', extrayendo la información relevante para cada sucursal.
5. Se construye un correo con los errores e indicadores encontrados.
6. Se envía automáticamente el correo a los destinatarios correspondientes.

## Tecnologías Utilizadas
- Microsoft Excel (macros en VBA)
- Microsoft Outlook (automatización de envíos)
- Conversión de rangos a HTML para incrustar tablas de Excel en correos


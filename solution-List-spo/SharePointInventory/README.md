# Inventario de columnas de SharePoint 2013

## ¿Qué es?

Es una aplicación de consola en C# que lee un informe CSV de migración,
identifica las filas de tipo `Lista` y consulta la configuración de columnas
de cada lista en su sitio de origen de SharePoint 2013. Al finalizar crea
`inventario.xlsx`, con una fila por columna y, si se producen errores, una hoja
adicional con el detalle de las listas que no se pudieron consultar.

El programa inventaría los metadatos y la configuración de las listas; no
descarga ni modifica los elementos, documentos o contenidos de estas.

## Requisitos

- Windows y SDK de .NET 8.
- Conectividad de red a los sitios de SharePoint indicados en el CSV.
- Una cuenta con permiso de lectura en los sitios y listas.
- Que SharePoint acepte autenticación Windows/NTLM para las credenciales
  configuradas. La autenticación basada en formularios o proveedores de
  identidad requiere adaptar el mecanismo de autenticación.

## Configuración

La configuración se lee de las variables de entorno del proceso. También se
puede usar un archivo `.env` en la carpeta desde la que se inicia el programa.
Define estas variables:

| Variable | Obligatoria | Descripción | Valor predeterminado |
| --- | --- | --- | --- |
| `SPO_USERNAME` | Sí | Usuario de Windows/SharePoint. | — |
| `SPO_PASSWORD` | Sí | Contraseña del usuario. | — |
| `SPO_DOMAIN` | No | Dominio Windows, si corresponde. | Vacío |
| `INPUT_CSV` | No | Ruta del informe CSV de entrada. | `StructureFailureReport_R1.csv` |
| `OUTPUT_EXCEL` | No | Ruta del libro Excel que se generará. | `inventario.xlsx` |

Ejemplo de archivo `.env`:

```dotenv
SPO_USERNAME=DOMINIO\usuario
SPO_PASSWORD=pon_aqui_la_contrasena
SPO_DOMAIN=DOMINIO
INPUT_CSV=StructureFailureReport_R1.csv
OUTPUT_EXCEL=inventario.xlsx
```

Guarda `.env` en la carpeta desde la que ejecutarás el comando. No compartas ni
subas ese archivo: contiene credenciales. Si el CSV está en otra ubicación,
indica una ruta absoluta o una ruta relativa a la carpeta actual.

El CSV debe incluir los encabezados `Structure type`, `Structure title` y
`Source structure URL`. Solo se procesan filas cuyo `Structure type` sea
`Lista`; el título identifica la lista y la URL identifica el sitio de origen.
Las filas duplicadas para el mismo sitio y título se procesan una sola vez.

## Ejecución

Abre PowerShell en la carpeta raíz del proyecto y ejecuta:

```powershell
dotnet run --project .\SharePointInventory\SharePointInventory.csproj
```

También puedes especificar las rutas por variables de entorno, en lugar de
definirlas en `.env`:

```powershell
$env:SPO_USERNAME = "DOMINIO\usuario"
$env:SPO_PASSWORD = "pon_aqui_la_contrasena"
$env:SPO_DOMAIN = "DOMINIO"
$env:INPUT_CSV = ".\StructureFailureReport_R1.csv"
$env:OUTPUT_EXCEL = ".\inventario.xlsx"
dotnet run --project .\SharePointInventory\SharePointInventory.csproj
```

Las rutas relativas se resuelven desde la carpeta actual de PowerShell. El
programa informa del avance y de los errores en la consola. Guarda el Excel
incluso si algunas listas fallan; en ese caso termina con un código distinto
de cero para señalar que el resultado es parcial.

## Cómo está armado el Excel

El libro contiene la hoja `Inventario`, con una fila por cada columna visible
consultada en una lista. Sus columnas son:

| Columna | Qué muestra |
| --- | --- |
| `Sitio de origen` | URL del sitio donde se encuentra la lista. |
| `Lista` | Título de la lista consultada. |
| `Nombre visible` | Etiqueta que SharePoint muestra para la columna. |
| `Nombre interno` | Nombre interno de la columna, usado por SharePoint. |
| `Tipo de dato` | Tipo definido por SharePoint, por ejemplo texto, fecha, búsqueda o elección. |
| `Obligatorio` | Si la columna requiere un valor. |
| `Valor predeterminado` | Valor que SharePoint propone por defecto, si está configurado. |
| `Unicidad` | Si SharePoint exige valores únicos en la columna. |
| `Opciones de elección` | Opciones configuradas para columnas de tipo elección. |
| `Fórmula` | Fórmula configurada en el esquema de la columna, si existe. |
| `Búsqueda (lista/campo)` | Identificador de la lista y campo mostrado por una columna de búsqueda, cuando están disponibles. |
| `Tipo de contenido` | Tipos de contenido de la lista que incluyen la columna. |

La primera fila contiene los encabezados y queda fija al desplazarse. La hoja
incluye filtros y anchos ajustados al contenido. Las columnas ocultas de
SharePoint se omiten. Si un dato no está configurado o no está expuesto en el
esquema consultado, su celda queda vacía.

Si una o más listas no se pudieron consultar, aparece también la hoja
`Errores`, con estas columnas:

| Columna | Qué muestra |
| --- | --- |
| `Sitio de origen` | URL del sitio que se intentó consultar. |
| `Lista` | Título de la lista que produjo el error. |
| `Error` | Mensaje devuelto al intentar consultar la lista. |

## Cómo usar el Excel

1. Abre `inventario.xlsx` con Microsoft Excel o una aplicación compatible.
2. En `Inventario`, usa los filtros de los encabezados para seleccionar un
   sitio, una lista, un tipo de dato u otra característica.
3. Compara `Nombre visible` con `Nombre interno` para identificar columnas
   cuyo nombre técnico difiere de la etiqueta mostrada.
4. Revisa `Obligatorio`, `Valor predeterminado` y `Unicidad` para analizar las
   restricciones de cada columna.
5. Consulta `Opciones de elección`, `Fórmula` y `Búsqueda (lista/campo)` cuando
   necesites revisar la configuración específica de esos tipos de columna.
6. Usa `Tipo de contenido` para ver en qué tipos de contenido de la lista
   aparece cada columna.
7. Si existe la hoja `Errores`, corrige o verifica el acceso a esas listas y
   vuelve a ejecutar el programa para completar el inventario.

Cada ejecución guarda el archivo en la ruta de `OUTPUT_EXCEL`; si ya existe un
libro en esa ubicación, se reemplaza al guardar el nuevo inventario.

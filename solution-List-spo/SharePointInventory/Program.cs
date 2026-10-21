using System.Globalization;
using System.Net;
using System.Xml.Linq;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.SharePoint.Client;
using IOFile = System.IO.File;

internal static class Program
{
    // Encabezados de la hoja principal del Excel, en el mismo orden que se exportan los datos.
    private static readonly string[] InventoryHeaders =
    [
        "Sitio de origen",
        "Lista",
        "Nombre visible",
        "Nombre interno",
        "Tipo de dato",
        "Obligatorio",
        "Valor predeterminado",
        "Unicidad",
        "Opciones de elección",
        "Fórmula",
        "Búsqueda (lista/campo)",
        "Tipo de contenido"
    ];

    private static int Main()
    {
        try
        {
            // Carga la configuración local y obtiene credenciales y rutas de entrada/salida.
            LoadDotEnv(Path.Combine(Directory.GetCurrentDirectory(), ".env"));

            var username = GetRequiredSetting("SPO_USERNAME");
            var password = GetRequiredSetting("SPO_PASSWORD");
            var domain = Environment.GetEnvironmentVariable("SPO_DOMAIN");
            var inputPath = ResolvePath(Environment.GetEnvironmentVariable("INPUT_CSV")
                ?? "StructureFailureReport_R1.csv");
            var outputPath = ResolvePath(Environment.GetEnvironmentVariable("OUTPUT_EXCEL")
                ?? "inventario.xlsx");

            if (!IOFile.Exists(inputPath))
            {
                throw new FileNotFoundException("No se encontró el CSV de entrada.", inputPath);
            }

            // Identifica las listas del CSV y prepara los datos y errores de la ejecución.
            var lists = ReadSourceLists(inputPath);
            if (lists.Count == 0)
            {
                throw new InvalidOperationException(
                    "El CSV no contiene filas con Structure type = 'Lista'.");
            }

            var rows = new List<InventoryRow>();
            var errors = new List<InventoryError>();

            foreach (var sourceList in lists)
            {
                try
                {
                    // Consulta las columnas de cada lista de forma independiente.
                    rows.AddRange(ReadListFields(sourceList, username, password, domain));
                    Console.WriteLine($"Leída: {sourceList.Title} ({sourceList.WebUrl})");
                }
                catch (Exception exception)
                {
                    // Conserva el error de esta lista y continúa con las demás.
                    errors.Add(new InventoryError(
                        sourceList.WebUrl,
                        sourceList.Title,
                        exception.Message));
                    Console.Error.WriteLine(
                        $"Error en '{sourceList.Title}' ({sourceList.WebUrl}): {exception.Message}");
                }
            }

            // Guarda los resultados incluso si algunas listas no pudieron consultarse.
            WriteWorkbook(outputPath, rows, errors);
            Console.WriteLine(
                $"Inventario guardado en: {outputPath}. " +
                $"Listas: {lists.Count}; columnas: {rows.Count}; errores: {errors.Count}.");

            return errors.Count == 0 ? 0 : 1;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"No se pudo generar el inventario: {exception.Message}");
            return 1;
        }
    }

    private static List<SourceList> ReadSourceLists(string inputPath)
    {
        // CsvHelper interpreta correctamente los encabezados y campos entrecomillados del informe.
        using var reader = new StreamReader(inputPath);
        using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HasHeaderRecord = true,
            TrimOptions = TrimOptions.Trim
        });

        // La clave compuesta evita procesar varias veces una lista repetida en el CSV.
        var lists = new Dictionary<string, SourceList>(StringComparer.OrdinalIgnoreCase);
        while (csv.Read())
        {
            // Solo interesan las filas que representan listas, no campos, vistas ni otros objetos.
            var structureType = csv.GetField("Structure type");
            if (!string.Equals(structureType, "Lista", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var title = csv.GetField("Structure title")?.Trim();
            var webUrl = csv.GetField("Source structure URL")?.Trim();
            if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(webUrl))
            {
                throw new InvalidDataException(
                    $"La fila {csv.Parser.Row} tiene una lista sin título o sin URL de origen.");
            }

            if (!Uri.TryCreate(webUrl, UriKind.Absolute, out var parsedUrl)
                || (parsedUrl.Scheme != Uri.UriSchemeHttps && parsedUrl.Scheme != Uri.UriSchemeHttp))
            {
                throw new InvalidDataException(
                    $"La URL de origen de la fila {csv.Parser.Row} no es válida: '{webUrl}'.");
            }

            // Quita la barra final para que diferencias de formato no creen entradas duplicadas.
            var normalizedUrl = parsedUrl.AbsoluteUri.TrimEnd('/');
            var key = $"{normalizedUrl}|{title}";
            lists.TryAdd(key, new SourceList(title, normalizedUrl));
        }

        return lists.Values.ToList();
    }

    private static IEnumerable<InventoryRow> ReadListFields(
        SourceList sourceList,
        string username,
        string password,
        string? domain)
    {
        // Crea una conexión CSOM al sitio de origen usando las credenciales proporcionadas.
        using var context = new ClientContext(sourceList.WebUrl)
        {
            Credentials = string.IsNullOrWhiteSpace(domain)
                ? new NetworkCredential(username, password)
                : new NetworkCredential(username, password, domain)
        };

        // Solicita los metadatos de las columnas y los tipos de contenido asociados a la lista.
        var list = context.Web.Lists.GetByTitle(sourceList.Title);
        context.Load(list, item => item.Title);
        context.Load(list.Fields, fields => fields.Include(
            field => field.Id,
            field => field.SchemaXml,
            field => field.Title,
            field => field.InternalName,
            field => field.TypeAsString,
            field => field.Required,
            field => field.DefaultValue,
            field => field.EnforceUniqueValues));
        context.Load(list.ContentTypes);
        context.ExecuteQuery();

        foreach (var contentType in list.ContentTypes)
        {
            context.Load(contentType.FieldLinks);
        }

        // Carga los vínculos entre tipos de contenido y columnas para completar el inventario.
        context.ExecuteQuery();

        // Indexa los tipos de contenido por identificador de columna para buscarlos rápidamente.
        var contentTypesByFieldId = new Dictionary<Guid, List<string>>();
        foreach (var contentType in list.ContentTypes)
        {
            foreach (var fieldLink in contentType.FieldLinks)
            {
                if (!contentTypesByFieldId.TryGetValue(fieldLink.Id, out var names))
                {
                    names = [];
                    contentTypesByFieldId.Add(fieldLink.Id, names);
                }

                names.Add(contentType.Name);
            }
        }

        foreach (var field in list.Fields)
        {
            // El XML del campo incluye detalles como opciones, fórmulas y configuración de búsqueda.
            var schema = XDocument.Parse(field.SchemaXml);
            var element = schema.Root
                ?? throw new InvalidDataException(
                    $"La definición XML de la columna '{field.SchemaXml}' está vacía.");
            if (bool.TryParse((string?)element.Attribute("Hidden"), out var isHidden) && isHidden)
            {
                // Las columnas ocultas no se muestran en la configuración normal de la lista.
                continue;
            }

            // Extrae configuraciones específicas del esquema XML que no se exponen como propiedades CSOM.
            contentTypesByFieldId.TryGetValue(field.Id, out var contentTypes);
            var choices = element
                .Descendants()
                .Where(item => item.Name.LocalName == "CHOICE")
                .Select(item => item.Value)
                .ToList();
            var lookupList = (string?)element.Attribute("List");
            var lookupField = (string?)element.Attribute("ShowField");
            var lookup = string.Join(" / ",
                new[] { lookupList, lookupField }.Where(value => !string.IsNullOrWhiteSpace(value)));

            // Devuelve una fila con los datos de esta columna para agregarla al Excel.
            yield return new InventoryRow(
                sourceList.WebUrl,
                sourceList.Title,
                field.Title,
                field.InternalName,
                field.TypeAsString,
                FormatBoolean(field.Required.ToString(CultureInfo.InvariantCulture)),
                field.DefaultValue ?? "",
                FormatBoolean(field.EnforceUniqueValues.ToString(CultureInfo.InvariantCulture)),
                string.Join("; ", choices),
                GetFormula(element),
                lookup,
                contentTypes is null ? "" : string.Join("; ", contentTypes.Distinct(StringComparer.OrdinalIgnoreCase)));
        }
    }

    private static string GetFormula(XElement fieldElement)
    {
        // Busca primero la fórmula como atributo y luego como elemento XML anidado.
        var formulaAttribute = (string?)fieldElement.Attribute("Formula");
        if (!string.IsNullOrWhiteSpace(formulaAttribute))
        {
            return formulaAttribute;
        }

        return fieldElement
            .Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "Formula")
            ?.Value ?? "";
    }

    private static string FormatBoolean(string? value)
    {
        // Convierte valores booleanos a etiquetas legibles en español; deja intactos otros valores.
        if (bool.TryParse(value, out var parsed))
        {
            return parsed ? "Sí" : "No";
        }

        return value ?? "";
    }

    private static void WriteWorkbook(
        string outputPath,
        IReadOnlyCollection<InventoryRow> rows,
        IReadOnlyCollection<InventoryError> errors)
    {
        // Asegura que exista la carpeta de destino antes de crear el archivo.
        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using var workbook = new XLWorkbook();
        // Escribe una fila del Excel por cada columna recopilada en SharePoint.
        var inventorySheet = workbook.Worksheets.Add("Inventario");
        for (var column = 0; column < InventoryHeaders.Length; column++)
        {
            inventorySheet.Cell(1, column + 1).Value = InventoryHeaders[column];
        }

        var rowNumber = 2;
        foreach (var item in rows)
        {
            var values = item.ToValues();
            for (var column = 0; column < values.Length; column++)
            {
                inventorySheet.Cell(rowNumber, column + 1).Value = values[column];
            }

            rowNumber++;
        }

        FormatSheet(inventorySheet, InventoryHeaders.Length, rowNumber - 1);

        if (errors.Count > 0)
        {
            // Si hubo fallos, agrega una hoja aparte para mostrar qué lista no se pudo leer.
            var errorSheet = workbook.Worksheets.Add("Errores");
            var headers = new[] { "Sitio de origen", "Lista", "Error" };
            for (var column = 0; column < headers.Length; column++)
            {
                errorSheet.Cell(1, column + 1).Value = headers[column];
            }

            rowNumber = 2;
            foreach (var error in errors)
            {
                errorSheet.Cell(rowNumber, 1).Value = error.WebUrl;
                errorSheet.Cell(rowNumber, 2).Value = error.Title;
                errorSheet.Cell(rowNumber, 3).Value = error.Message;
                rowNumber++;
            }

            FormatSheet(errorSheet, headers.Length, rowNumber - 1);
        }

        workbook.SaveAs(outputPath);
    }

    private static void FormatSheet(IXLWorksheet sheet, int columnCount, int lastRow)
    {
        // Facilita la lectura del libro: encabezado resaltado, filtros y primera fila fija.
        var header = sheet.Range(1, 1, 1, columnCount);
        header.Style.Font.Bold = true;
        header.Style.Fill.BackgroundColor = XLColor.LightBlue;
        sheet.SheetView.FreezeRows(1);
        sheet.Range(1, 1, Math.Max(1, lastRow), columnCount).SetAutoFilter();
        sheet.Columns(1, columnCount).AdjustToContents();
    }

    private static void LoadDotEnv(string path)
    {
        // Si no existe .env, las variables aún pueden llegar desde el entorno del proceso.
        if (!IOFile.Exists(path))
        {
            return;
        }

        foreach (var rawLine in IOFile.ReadLines(path))
        {
            var line = rawLine.Trim();
            // Ignora líneas vacías y comentarios del archivo de configuración.
            if (line.Length == 0 || line.StartsWith('#'))
            {
                continue;
            }

            var separator = line.IndexOf('=');
            if (separator <= 0)
            {
                throw new InvalidDataException($"Línea inválida en el archivo .env: '{line}'.");
            }

            var key = line[..separator].Trim();
            var value = line[(separator + 1)..].Trim();
            // Retira comillas envolventes para aceptar valores como "mi contraseña".
            if (value.Length >= 2
                && ((value[0] == '"' && value[^1] == '"')
                    || (value[0] == '\'' && value[^1] == '\'')))
            {
                value = value[1..^1];
            }

            if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(key)))
            {
                // No sobrescribe una variable que ya esté definida en el entorno.
                Environment.SetEnvironmentVariable(key, value);
            }
        }
    }

    private static string GetRequiredSetting(string name)
    {
        // Falla con un mensaje explícito si falta una credencial obligatoria.
        var value = Environment.GetEnvironmentVariable(name);
        return string.IsNullOrWhiteSpace(value)
            ? throw new InvalidOperationException(
                $"Falta la variable {name}. Configúrala en .env o en las variables de entorno.")
            : value;
    }

    private static string ResolvePath(string path) =>
        // Interpreta rutas relativas desde la carpeta desde la que se ejecutó el programa.
        Path.GetFullPath(path, Directory.GetCurrentDirectory());

    // Representa una lista de origen identificada por su título y la URL del sitio.
    private sealed record SourceList(string Title, string WebUrl);

    // Agrupa la ubicación de una lista y el mensaje si su consulta no tuvo éxito.
    private sealed record InventoryError(string WebUrl, string Title, string Message);

    // Modelo de una fila de la hoja Inventario, con una propiedad por columna exportada.
    private sealed record InventoryRow(
        string WebUrl,
        string ListTitle,
        string DisplayName,
        string InternalName,
        string DataType,
        string Required,
        string DefaultValue,
        string Unique,
        string Choices,
        string Formula,
        string Lookup,
        string ContentTypes)
    {
        // Mantiene los valores alineados con el orden definido en InventoryHeaders.
        public string[] ToValues() =>
        [
            WebUrl,
            ListTitle,
            DisplayName,
            InternalName,
            DataType,
            Required,
            DefaultValue,
            Unique,
            Choices,
            Formula,
            Lookup,
            ContentTypes
        ];
    }
}

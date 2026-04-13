using AddDodoProducts.Models;
using ClosedXML.Excel;
using DodoPayments.Client.Models.Products;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AddDodoProducts.Services;

/// <summary>
/// Reads dodoproducts/input.xlsx (sheet: create_product) and writes
/// dodoproducts/products.json.  Also reads the JSON back for API calls.
/// </summary>
public static class ExcelConverter
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    // ── Excel column headers are driven by JsonPropertyName on ProductEntry ──────
    // Any rename on the model automatically updates what column name is expected.

    public static List<ProductEntry> ConvertExcelToJson(string folderPath)
    {
        var xlsxPath = Path.Combine(folderPath, "input.xlsx");
        var jsonPath = Path.Combine(folderPath, "products.json");

        if (!File.Exists(xlsxPath))
            throw new FileNotFoundException($"Excel file not found: {xlsxPath}");

        using var wb = new XLWorkbook(xlsxPath);
        var ws = wb.Worksheet("create_product")
                 ?? throw new InvalidOperationException(
                     "Sheet 'create_product' not found in input.xlsx");

        var rows = ws.RangeUsed()?.RowsUsed().ToList()
                        ?? throw new InvalidOperationException("Sheet is empty.");
        var headerRow = rows.First();

        // Build column-index map from header row
        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in headerRow.CellsUsed())
            colMap[cell.GetString().Trim()] = cell.Address.ColumnNumber;

        string GetVal(IXLRow row, string col) =>
            colMap.TryGetValue(col, out var idx) ? row.Cell(idx).GetString().Trim() : "";

        // Change the LINQ .Where and .Select to use .WorksheetRow() to convert IXLRangeRow to IXLRow
        var entries = rows.Skip(1)
            .Where(r => !string.IsNullOrWhiteSpace(GetVal(r.WorksheetRow(), Col<ProductEntry>(x => x.Name))))
            .Select(r =>
            {
                var row = r.WorksheetRow();
                return new ProductEntry
                {
                    DodoProductId = GetVal(row, Col<ProductEntry>(x => x.DodoProductId)).NullIfEmpty(),
                    Name = GetVal(row, Col<ProductEntry>(x => x.Name)),
                    Description = GetVal(row, Col<ProductEntry>(x => x.Description)),
                    TaxCategory = GetVal(row, Col<ProductEntry>(x => x.TaxCategory)).IfEmpty("digital_products"),
                    PriceCents = int.TryParse(GetVal(row, Col<ProductEntry>(x => x.PriceCents)), out var p) ? p : 0,
                    Currency = GetVal(row, Col<ProductEntry>(x => x.Currency)).IfEmpty("USD"),
                    Discount = long.TryParse(GetVal(row, Col<ProductEntry>(x => x.Discount)), out var d) ? d : 0,
                    PurchasingPowerParity = GetVal(row, Col<ProductEntry>(x => x.PurchasingPowerParity)).IsTruthy(),
                    PayWhatYouWant = GetVal(row, Col<ProductEntry>(x => x.PayWhatYouWant)).IsTruthy(),
                    SuggestedPriceCents = int.TryParse(GetVal(row, Col<ProductEntry>(x => x.SuggestedPriceCents)), out var sp) ? sp : null,
                    Sku = GetVal(row, Col<ProductEntry>(x => x.Sku)).NullIfEmpty(),
                    Plan = GetVal(row, Col<ProductEntry>(x => x.Plan)).NullIfEmpty(),
                    LicenseKeyEnabled = GetVal(row, Col<ProductEntry>(x => x.LicenseKeyEnabled)).NullableTruthy(),
                    LicenseKeyActivationMessage = GetVal(row, Col<ProductEntry>(x => x.LicenseKeyActivationMessage)).NullIfEmpty(),
                    LicenseKeyActivationsLimit = int.TryParse(GetVal(row, Col<ProductEntry>(x => x.LicenseKeyActivationsLimit)), out var lkal) ? lkal : null,
                    LicenseKeyDurationDays = int.TryParse(GetVal(row, Col<ProductEntry>(x => x.LicenseKeyDurationDays)), out var lkdd) ? lkdd : null,
                };
            })
            .ToList();

        var json = JsonSerializer.Serialize(entries, JsonOpts);
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"  Converted {entries.Count} rows → {jsonPath}");
        return entries;
    }

    public static List<ProductEntry> LoadJson(string folderPath)
    {
        var jsonPath = Path.Combine(folderPath, "products.json");
        if (!File.Exists(jsonPath))
            throw new FileNotFoundException($"products.json not found: {jsonPath}");

        var json = File.ReadAllText(jsonPath);
        return JsonSerializer.Deserialize<List<ProductEntry>>(json, JsonOpts)
               ?? [];
    }

    public static void SaveJson(string folderPath, List<ProductEntry> entries)
    {
        var jsonPath = Path.Combine(folderPath, "products.json");
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(entries, JsonOpts));
    }

    /// <summary>
    /// Returns the <see cref="JsonPropertyNameAttribute"/> value for a property, which is
    /// exactly what the Excel column header must say — the single source of truth.
    /// </summary>
    private static string Col<T>(Expression<Func<T, object?>> expr)
    {
        var memberExpr = expr.Body is UnaryExpression unary
            ? (MemberExpression)unary.Operand
            : (MemberExpression)expr.Body;

        var attr = memberExpr.Member.GetCustomAttribute<JsonPropertyNameAttribute>();
        return attr?.Name ?? memberExpr.Member.Name;
    }

    /// <summary>
    /// Opens <paramref name="xlsxPath"/>, adds/finds a <c>dodo_product_id</c> column in the
    /// sheet and writes the id from each <see cref="ProductEntry"/> into the matching row
    /// (matched by name, case-insensitive).  Saves the workbook in-place.
    /// </summary>
    public static void WriteProductIdsToExcel(string xlsxPath, List<ProductEntry> entries)
    {
        if (!File.Exists(xlsxPath)) return;

        using var wb = new XLWorkbook(xlsxPath);
        var ws = wb.Worksheet("create_product");
        if (ws == null) return;

        var headerRow = ws.RangeUsed()?.RowsUsed().FirstOrDefault();
        if (headerRow == null) return;

        // Build column-index map from the header row
        var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int lastCol = 0;
        foreach (var cell in headerRow.WorksheetRow().CellsUsed())
        {
            var key = cell.GetString().Trim();
            if (!string.IsNullOrWhiteSpace(key))
                colMap[key] = cell.Address.ColumnNumber;
            if (cell.Address.ColumnNumber > lastCol)
                lastCol = cell.Address.ColumnNumber;
        }

        // Ensure the dodo_product_id column exists
        var idColName = Col<ProductEntry>(x => x.DodoProductId);
        if (!colMap.TryGetValue(idColName, out var idColIdx))
        {
            idColIdx = lastCol + 1;
            headerRow.WorksheetRow().Cell(idColIdx).Value = idColName;
            colMap[idColName] = idColIdx;
        }

        // Build a name → id lookup (skip entries without an id)
        var idLookup = entries
            .Where(e => !string.IsNullOrWhiteSpace(e.DodoProductId))
            .ToDictionary(e => e.Name.Trim(), e => e.DodoProductId!, StringComparer.OrdinalIgnoreCase);

        if (idLookup.Count == 0) return;

        // Find the name column index
        if (!colMap.TryGetValue(Col<ProductEntry>(x => x.Name), out var nameColIdx)) return;

        // Walk every data row and fill in the id
        var allRows = ws.RangeUsed()?.RowsUsed().Skip(1).ToList() ?? [];
        foreach (var rangeRow in allRows)
        {
            var row = rangeRow.WorksheetRow();
            var name = row.Cell(nameColIdx).GetString().Trim();
            if (idLookup.TryGetValue(name, out var id))
                row.Cell(idColIdx).Value = id;
        }

        wb.Save();
        Console.WriteLine($"  Written dodo_product_id column → {Path.GetFileName(xlsxPath)}");
    }

    /// <summary>
    /// Writes <paramref name="products"/> to a timestamped Excel file inside
    /// <c>dodoproducts/get/</c> and returns the full path of the written file.
    /// </summary>
    public static string WriteProductsToExcel(string folderPath, IEnumerable<ProductListResponse> products)
    {
        var getFolder = Path.Combine(folderPath, "get");
        Directory.CreateDirectory(getFolder);

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var filePath = Path.Combine(getFolder, $"products_{timestamp}.xlsx");

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("products");

        // Header row
        string[] headers = ["product_id", "name", "description", "tax_category", "currency", "price", "is_recurring", "created_at", "updated_at"];
        for (int c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }

        // Data rows
        int row = 2;
        foreach (var p in products)
        {
            ws.Cell(row, 1).Value = p.ProductID;
            ws.Cell(row, 2).Value = p.Name ?? "";
            ws.Cell(row, 3).Value = p.Description ?? "";
            ws.Cell(row, 4).Value = p.TaxCategory?.ToString() ?? "";
            ws.Cell(row, 5).Value = p.Currency?.ToString() ?? "";
            ws.Cell(row, 6).Value = p.Price;
            ws.Cell(row, 7).Value = p.IsRecurring;
            ws.Cell(row, 8).Value = p.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss");
            ws.Cell(row, 9).Value = p.UpdatedAt.ToString("yyyy-MM-dd HH:mm:ss");
            row++;
        }

        ws.Columns().AdjustToContents();
        wb.SaveAs(filePath);
        return filePath;
    }

    /// <summary>
    /// Writes a single <paramref name="product"/> (retrieved by ID) to a timestamped Excel file
    /// inside <c>dodoproducts/get/</c> and returns the full path of the written file.
    /// </summary>
    public static string WriteProductToExcel(string folderPath, Product product)
    {
        var getFolder = Path.Combine(folderPath, "get");
        Directory.CreateDirectory(getFolder);

        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var filePath = Path.Combine(getFolder, $"product_{product.ProductID}_{timestamp}.xlsx");

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("product");

        string[] headers = ["product_id", "name", "description", "tax_category", "license_key_enabled", "license_key_activation_message", "license_key_activations_limit", "created_at", "updated_at"];
        for (int c = 0; c < headers.Length; c++)
        {
            ws.Cell(1, c + 1).Value = headers[c];
            ws.Cell(1, c + 1).Style.Font.Bold = true;
        }

        ws.Cell(2, 1).Value = product.ProductID;
        ws.Cell(2, 2).Value = product.Name ?? "";
        ws.Cell(2, 3).Value = product.Description ?? "";
        ws.Cell(2, 4).Value = product.TaxCategory?.ToString() ?? "";
        ws.Cell(2, 5).Value = product.LicenseKeyEnabled;
        ws.Cell(2, 6).Value = product.LicenseKeyActivationMessage ?? "";
        ws.Cell(2, 7).Value = product.LicenseKeyActivationsLimit?.ToString() ?? "";
        ws.Cell(2, 8).Value = product.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss");
        ws.Cell(2, 9).Value = product.UpdatedAt.ToString("yyyy-MM-dd HH:mm:ss");

        ws.Columns().AdjustToContents();
        wb.SaveAs(filePath);
        return filePath;
    }
}

// ── Small string helpers ───────────────────────────────────────────────────────
file static class StringEx
{
    public static string IfEmpty(this string s, string fallback) =>
        string.IsNullOrWhiteSpace(s) ? fallback : s;

    public static string? NullIfEmpty(this string s) =>
        string.IsNullOrWhiteSpace(s) ? null : s;

    public static bool IsTruthy(this string s) =>
        s.Equals("true", StringComparison.OrdinalIgnoreCase)
     || s.Equals("yes", StringComparison.OrdinalIgnoreCase)
     || s == "1";

    public static bool? NullableTruthy(this string s) =>
        string.IsNullOrWhiteSpace(s) ? null : s.IsTruthy();
}



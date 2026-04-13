using AddDodoProducts.Models;
using AddDodoProducts.Services;
using DodoPayments.Client;
using DodoPayments.Client.Core;
using Microsoft.Extensions.Configuration;

// ── Config ────────────────────────────────────────────────────────────────────

var config = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: false)
    .AddEnvironmentVariables()          // env vars override appsettings (e.g. CI/CD)
    .Build();

var token = config["Dodo:BearerToken"]
               ?? throw new InvalidOperationException("Dodo:BearerToken is not configured.");
var testMode = config["Dodo:TestMode"]
               ?? throw new InvalidOperationException("Dodo:TestMode is not configured.");

var FolderPath = Path.Combine(FindProjectRoot(AppContext.BaseDirectory), "dodoproducts");

Directory.CreateDirectory(FolderPath);

var client = new DodoPaymentsClient
{
    BearerToken = token,
    BaseUrl = testMode.Equals("true", StringComparison.OrdinalIgnoreCase) ? EnvironmentUrl.TestMode : EnvironmentUrl.LiveMode
};

var productService = new ProductService(client);

// ── Main menu ─────────────────────────────────────────────────────────────────

while (true)
{
    Console.WriteLine();
    Console.WriteLine("╔══════════════════════════════════════╗");
    Console.WriteLine("║      Dodo Products Manager           ║");
    Console.WriteLine("╠══════════════════════════════════════╣");
    Console.WriteLine("║  1 │ Convert input.xlsx → JSON       ║");
    Console.WriteLine("║  2 │ Create all products from JSON   ║");
    Console.WriteLine("║  3 │ Convert Excel then Create all   ║");
    Console.WriteLine("║  4 │ Update all products from JSON   ║");
    Console.WriteLine("║  5 │ Convert Excel then Update all   ║");
    Console.WriteLine("║  6 │ Get all products → Excel        ║");
    Console.WriteLine("║  7 │ Get product by ID → Excel       ║");
    Console.WriteLine("║  0 │ Exit                            ║");
    Console.WriteLine("╚══════════════════════════════════════╝");
    Console.Write("  Choose: ");

    var choice = Console.ReadLine()?.Trim();
    Console.WriteLine();

    switch (choice)
    {
        case "1":
            ConvertOnly();
            break;
        case "2":
            await CreateAll();
            break;
        case "3":
            await ConvertThenCreate();
            break;
        case "4":
            await UpdateAll();
            break;
        case "5":
            await ConvertThenUpdate();
            break;
        case "6":
            await GetAllProducts();
            break;
        case "7":
            await GetProductById();
            break;
        case "0":
            Console.WriteLine("Bye!");
            return;
        default:
            Console.WriteLine("  Invalid choice, try again.");
            break;
    }
}

// ── Actions ───────────────────────────────────────────────────────────────────

void ConvertOnly()
{
    try
    {
        Console.WriteLine("  Converting input.xlsx → products.json ...");
        var entries = ExcelConverter.ConvertExcelToJson(FolderPath);
        Console.WriteLine($"  Done. {entries.Count} product(s) ready.");
    }
    catch (Exception ex) { PrintError(ex); }
}

async Task CreateAll()
{
    try
    {
        var entries = ExcelConverter.LoadJson(FolderPath);
        if (entries.Count == 0) { Console.WriteLine("  No entries in products.json."); return; }

        Console.WriteLine($"Creating {entries.Count} product(s) ...");
        int created = 0, skipped = 0;

        for (int i = 0; i < entries.Count; i++)
        {
            var e = entries[i];
            if (!string.IsNullOrWhiteSpace(e.DodoProductId))
            {
                Console.WriteLine($"  [{i + 1}] SKIP  {e.Name}  (already has id: {e.DodoProductId})");
                skipped++;
                continue;
            }

            Console.Write($"  [{i + 1}] Creating  {e.Name} ...");
            var id = await productService.CreateAsync(e);
            e.DodoProductId = id;
            ExcelConverter.SaveJson(FolderPath, entries);   // save after each so progress is kept
            Console.WriteLine($"  OK  →  {id}");
            created++;
        }

        Console.WriteLine($"\nDone. Created: {created}  Skipped: {skipped}");

        if (created > 0)
            ArchiveInputFile(FolderPath, "created", entries);
    }
    catch (Exception ex) { PrintError(ex); }
}

async Task ConvertThenCreate()
{
    ConvertOnly();
    await CreateAll();
}

async Task ConvertThenUpdate()
{
    ConvertOnly();
    await UpdateAll();
}

async Task UpdateAll()
{
    try
    {
        var entries = ExcelConverter.LoadJson(FolderPath);
        if (entries.Count == 0) { Console.WriteLine("  No entries in products.json."); return; }

        var toUpdate = entries.Where(e => !string.IsNullOrWhiteSpace(e.DodoProductId)).ToList();
        var noId = entries.Count - toUpdate.Count;

        if (toUpdate.Count == 0)
        {
            Console.WriteLine("  No products have a dodo_product_id yet. Run Create All first.");
            return;
        }

        if (noId > 0)
            Console.WriteLine($"  Note: {noId} product(s) skipped — no dodo_product_id.");

        Console.WriteLine($"  Updating {toUpdate.Count} product(s) ...");
        int updated = 0;

        foreach (var e in toUpdate)
        {
            Console.Write($"  Updating  {e.Name}  ({e.DodoProductId}) ...");
            await productService.UpdateAsync(e.DodoProductId!, e);
            Console.WriteLine("  OK");
            updated++;
        }

        Console.WriteLine($"\nDone. Updated: {updated}");

        if (updated > 0)
            ArchiveInputFile(FolderPath, "updated", entries);
    }
    catch (Exception ex) { PrintError(ex); }
}

void ArchiveInputFile(string folderPath, string operation, List<ProductEntry> entries)
{
    var srcPath = Path.Combine(folderPath, "input.xlsx");
    if (!File.Exists(srcPath))
    {
        Console.WriteLine($"  (input.xlsx not found — skipping archive step)");
        return;
    }

    ExcelConverter.WriteProductIdsToExcel(srcPath, entries);

    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
    var destFolder = Path.Combine(folderPath, operation);
    Directory.CreateDirectory(destFolder);

    var destFileName = $"input_{operation}_{timestamp}.xlsx";
    var destPath = Path.Combine(destFolder, destFileName);

    File.Move(srcPath, destPath);
    Console.WriteLine($"  Archived: input.xlsx → {Path.GetRelativePath(AppContext.BaseDirectory, destPath)}");
}

async Task GetAllProducts()
{
    try
    {
        Console.WriteLine("  Fetching all products from Dodo Payments ...");
        var products = await productService.GetAllAsync();

        if (products.Count == 0)
        {
            Console.WriteLine("  No products found.");
            return;
        }

        Console.WriteLine($"  Retrieved {products.Count} product(s). Writing to Excel ...");
        var filePath = ExcelConverter.WriteProductsToExcel(FolderPath, products);
        Console.WriteLine($"  Saved → {Path.GetRelativePath(AppContext.BaseDirectory, filePath)}");
    }
    catch (Exception ex) { PrintError(ex); }
}

async Task GetProductById()
{
    try
    {
        Console.Write("  Enter product ID: ");
        var productId = Console.ReadLine()?.Trim();

        if (string.IsNullOrWhiteSpace(productId))
        {
            Console.WriteLine("  No product ID entered.");
            return;
        }

        Console.WriteLine($"  Fetching product {productId} ...");
        var product = await productService.GetByIdAsync(productId);
        var filePath = ExcelConverter.WriteProductToExcel(FolderPath, product);
        Console.WriteLine($"  Saved → {Path.GetRelativePath(AppContext.BaseDirectory, filePath)}");
    }
    catch (Exception ex) { PrintError(ex); }
}

void PrintError(Exception ex) =>
    Console.WriteLine($"  ERROR: {ex.Message}");

// Walks up from the binary output folder until it finds a directory containing a .csproj file.
// This keeps dodoproducts/ next to the project source rather than inside bin\.
static string FindProjectRoot(string startDir)
{
    var dir = new DirectoryInfo(startDir);
    while (dir != null)
    {
        if (dir.GetFiles("*.csproj").Length > 0)
            return dir.FullName;
        dir = dir.Parent;
    }
    return startDir; // fallback: use the binary directory as-is
}

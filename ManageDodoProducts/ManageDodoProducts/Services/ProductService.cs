using AddDodoProducts.Models;
using DodoPayments.Client;
using DodoPayments.Client.Models.Misc;
using DodoPayments.Client.Models.Products;

namespace AddDodoProducts.Services;

/// <summary>
/// Wraps Dodo Payments product Create / Update / Get calls.
/// </summary>
public sealed class ProductService(DodoPaymentsClient client)
{
    //  Create 

    public async Task<string> CreateAsync(ProductEntry entry)
    {
        var metadata = new Dictionary<string, string>();
        if (!string.IsNullOrWhiteSpace(entry.Sku)) metadata["sku"] = entry.Sku!;
        if (!string.IsNullOrWhiteSpace(entry.Plan)) metadata["plan"] = entry.Plan!;

        var result = await client.Products.Create(new ProductCreateParams
        {
            Name = entry.Name,
            Description = entry.Description,
            TaxCategory = entry.TaxCategory,
            Metadata = metadata.Count > 0 ? metadata : null,
            Price = new OneTimePrice
            {
                Type = DodoPayments.Client.Models.Products.Type.OneTimePrice,
                Currency = ParseCurrency(entry.Currency),
                PriceValue = entry.PriceCents,
                Discount = entry.Discount,
                PurchasingPowerParity = entry.PurchasingPowerParity,
                PayWhatYouWant = entry.PayWhatYouWant ? true : null,
                SuggestedPrice = entry.SuggestedPriceCents,
            }
        });

        return result.ProductID;
    }

    //  Update 

    public async Task UpdateAsync(string dodoProductId, ProductEntry entry)
    {
        var metadata = new Dictionary<string, string>();
        if (!string.IsNullOrWhiteSpace(entry.Sku)) metadata["sku"] = entry.Sku!;
        if (!string.IsNullOrWhiteSpace(entry.Plan)) metadata["plan"] = entry.Plan!;

        await client.Products.Update(dodoProductId, new ProductUpdateParams
        {
            Name = entry.Name,
            Description = entry.Description,
            Metadata = metadata.Count > 0 ? metadata : null,
            Price = new OneTimePrice
            {
                Type = DodoPayments.Client.Models.Products.Type.OneTimePrice,
                Currency = ParseCurrency(entry.Currency),
                PriceValue = entry.PriceCents,
                Discount = entry.Discount,
                PurchasingPowerParity = entry.PurchasingPowerParity,
                PayWhatYouWant = entry.PayWhatYouWant ? true : null,
                SuggestedPrice = entry.SuggestedPriceCents,
            }
        });
    }

    //  Helpers ?

    private static Currency ParseCurrency(string code) =>
        code.ToUpperInvariant() switch
        {
            "USD" => Currency.Usd,
            "EUR" => Currency.Eur,
            "GBP" => Currency.Gbp,
            "INR" => Currency.Inr,
            _ => Currency.Usd,
        };

    //  Get (list all / by id) 

    public async Task<List<ProductListResponse>> GetAllAsync()
    {
        var results = new List<ProductListResponse>();
        var page = await client.Products.List(new ProductListParams { PageNumber = 0, PageSize = 50 });

        while (true)
        {
            results.AddRange(page.Items);
            Console.WriteLine($"  Fetched {results.Count} product(s) so far ...");

            if (!page.HasNext())
                break;

            page = await page.Next();
        }

        return results;
    }

    public async Task<Product> GetByIdAsync(string productId) =>
        await client.Products.Retrieve(productId);
}

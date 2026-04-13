using System.Text.Json.Serialization;
using System.Text.Json.Serialization;

namespace AddDodoProducts.Models;

/// <summary>
/// Represents one row from the create_product sheet in input.xlsx.
/// Also used as the JSON schema saved to dodoproducts/products.json.
/// </summary>
public sealed class ProductEntry
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("description")]
    public string Description { get; set; } = "";

    [JsonPropertyName("tax_category")]
    public string TaxCategory { get; set; } = "digital_products";

    /// <summary>Price in cents (e.g. 999 = $9.99).</summary>
    [JsonPropertyName("price_cents")]
    public int PriceCents { get; set; }

    [JsonPropertyName("currency")]
    public string Currency { get; set; } = "USD";

    [JsonPropertyName("discount")]
    public long Discount { get; set; } = 0;

    [JsonPropertyName("purchasing_power_parity")]
    public bool PurchasingPowerParity { get; set; } = true;

    [JsonPropertyName("pay_what_you_want")]
    public bool PayWhatYouWant { get; set; } = true;

    /// <summary>Optional suggested price in cents when PayWhatYouWant = true.</summary>
    [JsonPropertyName("suggested_price_cents")]
    public int? SuggestedPriceCents { get; set; }

    /// <summary>Optional SKU metadata tag.</summary>
    [JsonPropertyName("sku")]
    public string? Sku { get; set; }

    /// <summary>Optional plan label metadata tag.</summary>
    [JsonPropertyName("plan")]
    public string? Plan { get; set; }

    // ── License key fields ────────────────────────────────────────────────

    /// <summary>Whether the product requires a license key.</summary>
    [JsonPropertyName("license_key_enabled")]
    public bool? LicenseKeyEnabled { get; set; }

    /// <summary>Message sent to the customer upon license key activation.</summary>
    [JsonPropertyName("license_key_activation_message")]
    public string? LicenseKeyActivationMessage { get; set; }

    /// <summary>Maximum number of times the license key can be activated.</summary>
    [JsonPropertyName("license_key_activations_limit")]
    public int? LicenseKeyActivationsLimit { get; set; }

    /// <summary>Duration in days for which the license key is valid (null = lifetime).</summary>
    [JsonPropertyName("license_key_duration_days")]
    public int? LicenseKeyDurationDays { get; set; }

    // ── Populated after creation ──────────────────────────────────────────
    [JsonPropertyName("dodo_product_id")]
    public string? DodoProductId { get; set; }
}

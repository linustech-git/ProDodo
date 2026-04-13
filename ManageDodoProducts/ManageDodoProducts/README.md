# AddDodoProducts
# AddDodoProducts

A .NET 9 console tool for bulk-creating and updating products in **Dodo Payments** from an Excel spreadsheet.

---

## Table of Contents

1. [Overview](#overview)
2. [Project Structure](#project-structure)
3. [Prerequisites](#prerequisites)
4. [Configuration](#configuration)
5. [Excel Input Format](#excel-input-format)
6. [The products.json File](#the-productsjson-file)
7. [Running the Tool](#running-the-tool)
8. [Menu Options](#menu-options)
9. [Workflow Examples](#workflow-examples)
10. [Archive Behaviour](#archive-behaviour)
11. [Supported Currencies](#supported-currencies)
12. [Error Handling](#error-handling)

---

## Overview

This tool reads product definitions from `dodoproducts/input.xlsx`, converts them to `dodoproducts/products.json`, and sends create or update requests to the Dodo Payments API.

Key features:
- **Excel-first workflow** — define all products in a spreadsheet, no code changes needed
- **Column names enforced by the model** — Excel column headers must match the `[JsonPropertyName]` values on `ProductEntry`. Renaming a property automatically updates what the tool expects; no magic strings
- **Idempotent creates** — products that already have a `dodo_product_id` are skipped automatically
- **Progress is persisted** — after each successful create, the `dodo_product_id` is saved to `products.json` immediately, so a crash mid-run never loses progress
- **IDs written back to Excel** — before archiving, the tool writes the `dodo_product_id` for each product into a `dodo_product_id` column in `input.xlsx`
- **Automatic archiving** — after a successful create or update run, `input.xlsx` is renamed with the operation and a timestamp and moved to a `created/` or `updated/` subfolder
- **Test / Live switching** — a single config flag switches between Dodo test mode and live mode

---

## Project Structure

```
AddDodoProducts/
+-- dodoproducts/               <- created at runtime, not committed
|   +-- input.xlsx              <- you fill this in
|   +-- products.json           <- auto-generated, stores dodo_product_id per product
|   +-- created/                <- archived Excel files after create runs
|   +-- updated/                <- archived Excel files after update runs
+-- Models/
|   +-- ProductEntry.cs         <- C# class mirroring one Excel row / JSON object
+-- Services/
|   +-- ExcelConverter.cs       <- reads input.xlsx, writes products.json, writes IDs back to Excel
|   +-- ProductService.cs       <- calls Dodo Payments Create / Update API
+-- appsettings.json            <- bearer token + test/live flag (not committed to git)
+-- Program.cs                  <- console menu entry point
```

> `dodoproducts/` is created automatically at startup next to the `.csproj` file, not inside `bin/`, so your Excel files stay in the same place regardless of build configuration.

---

## Prerequisites

- [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0)
- A Dodo Payments account with an API key ([dashboard.dodopayments.com](https://dashboard.dodopayments.com))
- Microsoft Excel or any tool that can produce `.xlsx` files (LibreOffice, Google Sheets export, etc.)

---

## Configuration

Edit `AddDodoProducts/appsettings.json`:

```json
{
  "Dodo": {
    "BearerToken": "your-api-key-here",
    "TestMode": "true"
  }
}
```

| Key | Description |
|---|---|
| `Dodo:BearerToken` | Your Dodo Payments API key. Found in the Dodo dashboard under **Developer -> API Keys** |
| `Dodo:TestMode` | `"true"` uses the Dodo test environment. `"false"` uses the live environment |

> **Security:** `appsettings.json` is listed in `.gitignore` and will never be committed. Do not share this file.

### Environment Variable Override

Any key can be overridden with an environment variable using `__` as the separator:

```
Dodo__BearerToken=your-key
Dodo__TestMode=false
```

---

## Excel Input Format

Place your Excel file at:

```
AddDodoProducts/dodoproducts/input.xlsx
```

The file must contain a sheet named exactly **`create_product`**.

The **first row** is the header row. Column order does not matter — columns are matched by header name (case-insensitive). All column names are derived directly from the `[JsonPropertyName]` attributes on `ProductEntry` — the single source of truth.

### Columns

| Column | Required | Type | Description |
|---|---|---|---|
| `name` | Yes | string | Product name shown in the Dodo dashboard and checkout |
| `description` | No | string | Short product description (max 1000 characters) |
| `tax_category` | No | string | Defaults to `digital_products`. Other values: `saas`, `ebook`, etc. |
| `price_cents` | Yes | integer | Price in the smallest currency unit. E.g. `999` = $9.99 |
| `currency` | No | string | 3-letter ISO code. Defaults to `USD`. See [Supported Currencies](#supported-currencies) |
| `discount` | No | integer | Discount percentage 0-100. Defaults to `0` |
| `purchasing_power_parity` | No | boolean | Enables local currency price adjustments |
| `pay_what_you_want` | No | boolean | Lets customers pay any amount >= `price_cents` |
| `suggested_price_cents` | No | integer | Suggested amount in cents when `pay_what_you_want = true` |
| `sku` | No | string | Internal SKU stored in Dodo product metadata |
| `plan` | No | string | Plan label stored in Dodo product metadata. E.g. `monthly`, `yearly` |
| `license_key_enabled` | No | boolean | Whether the product requires a license key |
| `license_key_activation_message` | No | string | Message sent to the customer on license key activation |
| `license_key_activations_limit` | No | integer | Maximum number of times the license key can be activated |
| `license_key_duration_days` | No | integer | License key validity in days. Leave blank for lifetime |

> Rows with an empty `name` cell are silently skipped.
> Boolean columns accept: `true`, `false`, `yes`, `no`, `1`, `0` (case-insensitive).

### Example Sheet (`create_product`)

| name | description | price_cents | currency | sku | plan | license_key_enabled |
|---|---|---|---|---|---|---|
| Numerinus.Algebra - Monthly | Monthly license | 999 | USD | num-alg-monthly | monthly | false |
| Numerinus.Algebra - Yearly | Yearly license | 7999 | USD | num-alg-yearly | yearly | false |

---

## The products.json File

After running **Convert** (option 1), a `products.json` is written to `dodoproducts/`:

```json
[
  {
    "name": "Numerinus.Algebra - Monthly",
    "description": "Monthly license",
    "tax_category": "digital_products",
    "price_cents": 999,
    "currency": "USD",
    "discount": 0,
    "purchasing_power_parity": false,
    "pay_what_you_want": false,
    "suggested_price_cents": null,
    "sku": "num-alg-monthly",
    "plan": "monthly",
    "license_key_enabled": null,
    "license_key_activation_message": null,
    "license_key_activations_limit": null,
    "license_key_duration_days": null,
    "dodo_product_id": null
  }

```

After running **Create All** (option 2), `dodo_product_id` is filled in for each successfully created product:

```json
"dodo_product_id": "prod_01abc123xyz"
```

This file acts as the source of truth between runs. Keep it — it prevents duplicate creates.

---

## Running the Tool

```bash
cd AddDodoProducts
dotnet run
```

---

## Menu Options

```
+======================================+
|      Dodo Products Manager           |
+======================================+
|  1 | Convert input.xlsx -> JSON      |
|  2 | Create all products from JSON   |
|  3 | Convert Excel then Create all   |
|  4 | Update all products from JSON   |
|  5 | Convert Excel then Update all   |
|  0 | Exit                            |
+======================================+
```

| Option | What it does |
|---|---|
| **1 - Convert** | Reads `input.xlsx` (sheet: `create_product`) and writes `products.json`. Does **not** call the API. Use this to preview what will be sent. |
| **2 - Create All** | Reads `products.json` and creates each product that does **not** yet have a `dodo_product_id`. Saves the returned ID to `products.json` after each call, writes IDs back into `input.xlsx`, then archives it to `dodoproducts/created/`. |
| **3 - Convert then Create** | Runs option 1 then option 2 in sequence. Use this for the initial bulk load. |
| **4 - Update All** | Reads `products.json` and updates every product that **has** a `dodo_product_id`. Products without an ID are skipped with a note. Archives `input.xlsx` to `dodoproducts/updated/` when done. |
| **5 - Convert then Update** | Runs option 1 then option 4 in sequence. Use this when you have changed product details in Excel and want to push those changes to Dodo. |
| **0 - Exit** | Exits the application. |

---

## Workflow Examples

### Initial bulk product load

1. Fill in `dodoproducts/input.xlsx` (sheet: `create_product`)
2. Run the tool -> choose **3** (Convert then Create)
3. All products are created in Dodo; IDs are saved to `products.json` and written back into the Excel file
4. The Excel file is moved to `dodoproducts/created/input_created_YYYYMMDD_HHmmss.xlsx`

### Adding new products

1. Place a new `input.xlsx` in `dodoproducts/` with only the new rows
2. Run the tool -> choose **3** (Convert then Create)
3. Only rows without a `dodo_product_id` are created — existing ones are skipped

### Updating product details

1. Place an `input.xlsx` in `dodoproducts/` with the rows to update (must include a `dodo_product_id` column, or run from an existing `products.json`)
2. Run the tool -> choose **5** (Convert then Update)
3. All products with a `dodo_product_id` are updated in Dodo
4. The Excel file is moved to `dodoproducts/updated/input_updated_YYYYMMDD_HHmmss.xlsx`

---

## Archive Behaviour

After a successful **Create** or **Update** run where at least one product was processed:

1. The `dodo_product_id` value is written into a `dodo_product_id` column in `input.xlsx` for every matched product (the column is added automatically if it does not exist)
2. The file is renamed to `input_{operation}_{timestamp}.xlsx` — e.g. `input_created_20250601_143022.xlsx`
3. The renamed file is moved to `dodoproducts/created/` or `dodoproducts/updated/`

This gives a timestamped audit trail of every bulk operation with the Dodo IDs embedded directly in the spreadsheet.

---

## Supported Currencies

| Code | Currency |
|---|---|
| `USD` | US Dollar (default) |
| `EUR` | Euro |
| `GBP` | British Pound |
| `INR` | Indian Rupee |

Any unlisted currency code defaults to `USD`. To add more, edit `ParseCurrency()` in `Services/ProductService.cs`.

---

## Error Handling

- If an API call fails for a product, the error is printed and the loop continues with the next product — a single failure does not abort the entire run
- `products.json` is saved after **each** successful create, so if the tool crashes mid-run, already-created products retain their IDs and are skipped on the next run
- If `appsettings.json` is missing or `Dodo:BearerToken` is empty, the tool throws immediately on startup before any API calls are made
- If `input.xlsx` is not found when archiving, the archive step is skipped with a warning and the run result is still reported

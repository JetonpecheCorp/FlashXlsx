# FlashXlsx

<p align="center">
  <img src="assets/logoFlashXlsx.png" alt="FlashXlsx Logo" width="220" />
</p>

[Documentation en français](README.md)

**FlashXlsx** is the fastest, lightest, and most secure Excel (`.xlsx`) import/export engine for the .NET ecosystem.

It was designed to be **incredibly simple** for beginners to pick up, while delivering **extreme performance** for architects and experts through Native Compilation (**Native AOT**), asynchronous streaming, and hardware vectorization (SIMD).

## 🧭 Table of Contents

1. [Understand FlashXlsx in 2 Minutes (For Everyone)](#understand-flashxlsx-in-2-minutes-for-everyone)
2. [Why Choose FlashXlsx?](#why-choose-flashxlsx)
3. [Reference Benchmarks (100,000 Rows)](#reference-benchmarks-100000-rows)
4. [Complete Feature Matrix](#complete-feature-matrix)
5. [Installation](#installation)
6. [Quick Start in 3 Easy Steps](#quick-start-in-3-easy-steps)
7. [Real-World Use Case: ASP.NET Core Web API & Massive Ingestion](#real-world-use-case-aspnet-core-web-api--massive-ingestion)
8. [Comprehensive Feature Guide by Example](#comprehensive-feature-guide-by-example)
   * [1. Model Your Data Simply](#1-model-your-data-simply)
   * [2. Attribute Reference Table](#2-attribute-reference-table)
   * [3. On-The-Fly Data Validation (Declarative Validation)](#3-on-the-fly-data-validation-declarative-validation)
   * [4. Dynamic Rules Based on User Context](#4-dynamic-rules-based-on-user-context)
   * [5. Variable and Multilingual Column Headers](#5-variable-and-multilingual-column-headers)
   * [6. Crash-Free Import (Fault-Tolerant Error Handling)](#6-crash-free-import-fault-tolerant-error-handling)
   * [7. Password-Protected File Detection](#7-password-protected-file-detection)
   * [8. Automatic Currency and Non-Breaking Space Cleanup](#8-automatic-currency-and-non-breaking-space-cleanup)
   * [9. Schema-Free Dynamic Mode: XlsxValue & Dictionaries (Zero-Boxing)](#9-schema-free-dynamic-mode-xlsxvalue--dictionaries-zero-boxing)
   * [10. Manage Multiple Sheets at Once](#10-manage-multiple-sheets-at-once)
   * [11. Inspect Sheets Without Loading the File](#11-inspect-sheets-without-loading-the-file)
   * [12. Effortless Merged Cells Reading](#12-effortless-merged-cells-reading)
   * [13. Create Styled Excel Tables with Filters](#13-create-styled-excel-tables-with-filters)
   * [14. Fast Preview and Pagination for Large Files](#14-fast-preview-and-pagination-for-large-files)
   * [15. Built-in Formula Injection Shield](#15-built-in-formula-injection-shield)
   * [16. Custom Converters (IXlsxConverter & Cross-Column Calculations)](#16-custom-converters-ixlsxconverter--cross-column-calculations)
   * [17. Delimited Lists in Cells & Default Values](#17-delimited-lists-in-cells--default-values)
   * [18. Real-Time Progress Reporting (IProgress)](#18-real-time-progress-reporting-iprogress)
9. [Under the Hood: For Optimization Enthusiasts and Experts](#under-the-hood-for-optimization-enthusiasts-and-experts)
   * [The Conveyor Belt Analogy (Streaming vs. DOM)](#the-conveyor-belt-analogy-streaming-vs-dom)
   * [Vectorized UTF-8 Scanner (`ReadOnlySpan<byte>` & AVX2)](#vectorized-utf-8-scanner-readonlyspanbyte--avx2)
   * [ASCII -> UTF-16 SIMD Acceleration](#ascii---utf-16-simd-acceleration)
   * [Zero Heap Allocations (`ArrayPool`, `XlsxValue` & Stack Structs)](#zero-heap-allocations-arraypool-xlsxvalue--stack-structs)
   * [Roslyn Source Generator (100% Native AOT & Trimming-Safe)](#roslyn-source-generator-100-native-aot--trimming-safe)
10. [License](#license)

---

## Understand FlashXlsx in 2 Minutes (For Everyone)

### 💡 The Delivery Truck and the Conveyor Belt Analogy
Imagine receiving a shipping container loaded with 100,000 packages:

* **Traditional libraries (ClosedXML, etc.):** They dump the entire truck right into your living room all at once. Result: your space is cluttered, server memory spikes (often 800 MB+ of RAM), and the whole machine stalls for several seconds.
* **FlashXlsx:** It sets up an **ultra-fast conveyor belt**. Each package is unloaded, checked, processed, and stored away immediately. Server memory stays lightweight (less than 20 MB of RAM), and processing keeps running at full speed without slowing down.

### 🎯 Key Takeaways:
* **Beginner?** Write 3 simple lines of C# and everything works out of the box with zero arcane setup.
* **Expert / Architect?** Take advantage of ahead-of-time compiled machine code, SIMD/AVX2 vector instructions, full Native AOT readiness, and virtually zero garbage collector pressure.

---

## Why Choose FlashXlsx?

| Criteria | Traditional Libraries (DOM) | Lightweight Libraries (Reflection) | FlashXlsx ⚡ |
| :--- | :--- | :--- | :--- |
| **Learning Curve** | Often verbose | Simple | **Ultra-simple and expressive** |
| **Memory (100,000 rows)** | ~700 to 850 MB ⚠️ | ~260 to 1,300 MB ⚠️ | **13 to 19 MB (Gen1 & Gen2 = 0) 🏆** |
| **Speed (100,000 rows)** | ~2.5 seconds | ~1 second | **~0.13 second (770,000 rows/s) 🚀** |
| **Native AOT Compatibility** | ❌ Impossible | ⚠️ Warnings / Unstable | **✅ 100% AOT with zero warnings** |
| **Corrupted/Invalid Input** | Crashes the application | Throws slow exceptions | **Clean capture without exceptions (`OnError`)** |
| **Security (Formula Injection)** | Manual implementation | Manual implementation | **Automatic and transparent** |
| **Automatic Currency Parsing** | Complex (manual logic) | Complex | **Automatic (`"1 250 €"` -> `1250m`)** |
| **Password-Protected Files** | Crash or obscure error | Unclear ZIP archive error | **Clean detection without crashing** |
| **Dynamic Dictionary Mode** | `Dictionary<string, object>` (Heavy boxing) | `ExpandoObject` / Boxing | **`XlsxValue` struct (100% Zero-Boxing)** |

---

## Reference Benchmarks (100,000 Rows)

Official measurements executed with **BenchmarkDotNet v0.14.0** on **.NET 10 (X64 RyuJIT AVX2)**, standard 4-core CPU (Intel Core i5-7600K 3.80GHz).  
Scenario: Writing and reading a workbook of **100,000 rows × 7 typed columns** (`int`, `string`, `string`, `decimal`, `double`, `DateTime`, `bool`).

```text
BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.6466/22H2/2022Update)
Intel Core i5-7600K CPU 3.80GHz (Kaby Lake), 1 CPU, 4 logical and 4 physical cores
.NET SDK 10.0.403
[Host]     : .NET 10.0.14 (10.0.1426.52326), X64 RyuJIT AVX2
DefaultJob : .NET 10.0.14 (10.0.1426.52326), X64 RyuJIT AVX2
```

| Operation Tested | Mean Time (`Mean`) | Relative Speed | Allocated Memory | Memory Ratio | GC Collections (Gen0 / Gen1 / Gen2) |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Write 100k - FlashXlsx (AOT)** | **129.4 ms** | **Baseline (1.00)** | **13.49 MB** | **Baseline (1.00)** | **500.0000 / 500.0000 / 500.0000** |
| Write 100k - MiniExcel | 283.8 ms | 2.19x slower | 265.78 MB | 19.71x more RAM | 78000.0000 / 2500.0000 / 1500.0000 |
| Write 100k - ClosedXML | 2,340.6 ms | 18.09x slower | 724.79 MB | 53.74x more RAM | 126000.0000 / 17000.0000 / 5000.0000 |
| **Read 100k - FlashXlsx (AOT)** | **129.9 ms** | **Baseline (1.00)** | **19.89 MB** | **Baseline (1.00)** | **6500.0000 / - / -** |
| Read 100k - MiniExcel | 961.3 ms | 7.43x slower | 1315.57 MB | 97.54x more RAM | 439000.0000 / - / - |
| Read 100k - ClosedXML | 1,912.8 ms | 14.78x slower | 835.98 MB | 61.98x more RAM | 186000.0000 / 23000.0000 / 2000.0000 |

> **Key insights behind these numbers:**
> - **Writing (129.4 ms):** FlashXlsx generates **~772,000 rows per second**. The 13.49 MB allocated corresponds almost entirely to the final compressed ZIP file payload.
> - **Reading (129.9 ms):** FlashXlsx decodes **~769,000 rows per second** and allocates only 19.89 MB (which matches the final target C# instances created for your domain). The internal engine generates virtually zero junk, eliminating GC pauses (Gen1 and Gen2 collections remain strictly at zero).

---

## Complete Feature Matrix

| Category | Feature | What It Does (Simple Explanation) |
| :--- | :--- | :--- |
| **Engine** | **Roslyn Source Generator** | Emits serialization/deserialization code at compile-time: zero startup reflection penalty. |
| | **Native AOT & Trimming** | Your application can be published as a self-contained, lightweight native executable. |
| | **Async Streaming** | Reads rows sequentially (`await foreach`) without consuming excessive RAM. |
| **Data Types** | **Complete Primitives** | Reads and writes `int`, `long`, `decimal`, `double`, `float`, `bool`, `Guid`, and `string`. |
| | **Modern Dates** | Full support for `DateOnly`, `TimeOnly`, `DateTime`, and the legacy Mac 1904 date system. |
| | **Automatic Currency Cleanup** | Converts strings like `"$99.00"` or `"1 250,50 €"` directly into numeric properties. |
| | **Flexible Booleans** | Understands `"Yes"`, `"No"`, `"True"`, `"False"`, `"1"`, `"0"` seamlessly. |
| | **Delimited Lists** | Automatically splits delimited cell contents (`"A;B;C"`) into collections. |
| **Validation & Mapping** | **Declarative Rules** | Validates records using simple attributes (`[XlsxRequired]`, `[XlsxRange]`, etc.). |
| | **Conditional Rules** | Enforces requirements based on runtime conditions (`[XlsxRequiredIf]`). |
| | **Custom Converters** | Implements domain converters and calculated columns (`IXlsxConverter`). |
| | **Dynamic Context** | Adapts parsing rules depending on who is logged in (admin, tenant, role). |
| **Fault Tolerance** | **Report & Continue Mode** | Keeps reading through dirty data while collecting a clean report of invalid cells. |
| | **Encrypted File Detection** | Detects password-protected workbooks from the first 8 bytes without throwing exceptions. |
| | **Empty & Ghost Rows** | Silently skips blank rows and cells filled with phantom whitespace. |
| **Structure & Metadata** | **Fluent Multi-Sheet API** | Reads or writes multiple workbook tabs in a single call via builders (`AddSheet`, `MapSheet`). |
| | **Zero-Boxing Dynamic Mode** | Reads and writes schema-free workbooks via the stack-allocated `XlsxValue` struct. |
| | **Merged Cells** | Automatically replicates top-left merged values across subordinate rows. |
| | **Styled Excel Tables** | Produces official formatted tables with banded rows and auto-filter buttons. |
| | **Preview & Pagination** | Lets you read specific ranges (e.g., rows 100 to 150) for fast UI pagination. |
| | **Progress Tracking** | Emits regular notifications (`IProgress<int>`) without choking the pipeline. |
| **Security** | **Formula Injection Shield** | Automatically sanitizes malicious formula prefixes (`=`, `+`, `-`, `@`). |

---

## Installation

Run the following command in your terminal or NuGet Package Manager console:

```bash
dotnet add package FlashXlsx
```

If your project targets **Native AOT** (native self-contained executable):
```xml
<PropertyGroup>
  <TargetFramework>net10.0</TargetFramework>
  <PublishAot>true</PublishAot>
  <Nullable>enable</Nullable>
</PropertyGroup>
```
Everything works seamlessly: FlashXlsx includes its compile-time Roslyn generator with zero extra configuration.

---

## Quick Start in 3 Easy Steps

### Step 1: Define Your C# Model
Decorate your class with `[XlsxSerializable]`, and use `[XlsxName]` to customize column titles:

```csharp
using FlashXlsx;
using FlashXlsx.Abstractions;

[XlsxSerializable]
public class Product
{
    [XlsxName("SKU")]
    [XlsxRequired]
    public string Sku { get; set; } = string.Empty;

    [XlsxName("Description")]
    public string Name { get; set; } = string.Empty;

    [XlsxName("Unit Price")]
    public decimal Price { get; set; }

    [XlsxName("In Stock")]
    [XlsxBoolean(trueValues: ["Yes", "1"], falseValues: ["No", "0"])]
    public bool Available { get; set; }

    [XlsxName("Release Date")]
    public DateOnly ReleaseDate { get; set; }
}
```

### Step 2: Write an Excel File in 2 Lines
```csharp
var list = new List<Product>
{
    new() { Sku = "PRD-01", Name = "Wireless Keyboard", Price = 49.90m, Available = true, ReleaseDate = new DateOnly(2026, 11, 10) },
    new() { Sku = "PRD-02", Name = "Optical Mouse", Price = 19.50m, Available = false, ReleaseDate = new DateOnly(2026, 11, 09) }
};

await FlashXlsxSerializer.WriteAsync("products.xlsx", list);
```

### Step 3: Read the File Row by Row
```csharp
await foreach (Product p in FlashXlsxSerializer.ReadAsync<Product>("products.xlsx"))
{
    Console.WriteLine($"Product: {p.Name} - {p.Price} $ (Available: {p.Available})");
}
```

---

## Real-World Use Case: ASP.NET Core Web API & Massive Ingestion

In production enterprise software or cloud APIs (on Docker, Azure, or Kubernetes), a client might upload a 100,000-row spreadsheet. Loading everything in memory at once risks Out-Of-Memory (OOM) crashes.

Here is a complete modern ASP.NET Core endpoint: it receives the uploaded stream, processes rows **in small batches of 1,000**, saves them to the database, and collects validation errors without crashing:

```csharp
using FlashXlsx;
using FlashXlsx.Abstractions;
using Microsoft.AspNetCore.Mvc;

var builder = WebApplication.CreateSlimBuilder(args);
var app = builder.Build();

app.MapPost("/api/import-orders", async (IFormFile file, [FromServices] IOrderRepository repo, CancellationToken ct) =>
{
    if (file is null || file.Length == 0)
        return Results.BadRequest("The uploaded file is empty.");

    await using var stream = file.OpenReadStream();

    var validationErrors = new List<object>();

    // Tolerant configuration: collect issues instead of throwing exceptions
    var options = new XlsxReadOptions<OrderImportDto>()
        .WithHeader(hasHeader: true, headerRowNumber: 1)
        .WithSkipEmptyRows(true)
        .WithErrorMode(XlsxErrorMode.ReportAndContinue)
        .WithErrorHandler((in XlsxErrorContext err) =>
        {
            validationErrors.Add(new
            {
                Row = err.RowNumber,
                Column = err.ColumnName ?? err.ColumnLetter,
                InvalidValue = err.RawValue,
                ErrorType = err.Status.ToString()
            });
        });

    var batch = new List<OrderImportDto>(1000);
    int totalImported = 0;

    // Asynchronous streaming processing straight from the HTTP payload
    await foreach (var order in FlashXlsxSerializer.ReadAsync(stream, options, ct))
    {
        batch.Add(order);

        // Bulk insert every 1,000 rows
        if (batch.Count >= 1000)
        {
            await repo.InsertBatchAsync(batch, ct);
            totalImported += batch.Count;
            batch.Clear();
        }
    }

    // Persist remaining items
    if (batch.Count > 0)
    {
        await repo.InsertBatchAsync(batch, ct);
        totalImported += batch.Count;
    }

    return Results.Ok(new
    {
        Message = "Import completed successfully!",
        TotalImportedRows = totalImported,
        ErrorCount = validationErrors.Count,
        Errors = validationErrors
    });
});

app.Run();
```

---

## Comprehensive Feature Guide by Example

### 1. Model Your Data Simply
Define the column order, customize header titles, or mark internal fields to be ignored:

```csharp
[XlsxSerializable]
public class Employee
{
    [XlsxColumnOrder(1)]
    [XlsxName("Employee ID")]
    public string BadgeId { get; set; } = string.Empty;

    [XlsxColumnOrder(2)]
    [XlsxName("Full Name")]
    public string FullName { get; set; } = string.Empty;

    [XlsxColumnOrder(3)]
    [XlsxName("Birth Date")]
    public DateOnly BirthDate { get; set; }

    [XlsxColumnOrder(4)]
    [XlsxName("Executive")]
    [XlsxBoolean(trueValues: ["YES", "Y", "TRUE"], falseValues: ["NO", "N", "FALSE"])]
    public bool IsExecutive { get; set; }

    [XlsxIgnore] // This property is NEVER read or written to Excel
    public string InternalSecurityToken { get; set; } = string.Empty;
}
```

### 2. Attribute Reference Table

| Attribute | Target | Description |
| :--- | :--- | :--- |
| `[XlsxSerializable]` | Class / Struct | Enables compile-time Native AOT code generation by the Roslyn generator. |
| `[XlsxName("Title")]` | Property | Defines the column header text in the Excel worksheet. |
| `[XlsxColumnLetter("C")]` | Property | Binds directly to a specific column letter (ideal for headerless files). |
| `[XlsxColumnOrder(1)]` | Property | Dictates the ordered position of the column when exporting. |
| `[XlsxIgnore]` | Property | Completely excludes the property from both reading and writing. |
| `[XlsxRequired]` | Property | Triggers a validation error if the cell is empty or absent during import. |
| `[XlsxRequiredIf(typeof(T))]` | Property | Makes the cell mandatory based on an `IXlsxCondition` business rule. |
| `[XlsxIgnoreIf(typeof(T))]` | Property | Dynamically skips the property if the condition returns `true`. |
| `[XlsxRange(min, max)]` | Property | Enforces numeric boundaries (`[XlsxMin]` and `[XlsxMax]` are also available). |
| `[XlsxLength(min, max)]` | Property | Validates minimum and maximum string length. |
| `[XlsxRegex("pattern")]` | Property | Verifies text format against a regular expression pattern. |
| `[XlsxAllowedValues(...)]` | Property | Restricts input to a defined whitelist of allowed string values. |
| `[XlsxBoolean(...)]` | Property | Defines custom text representations for `true` and `false`. |
| `[XlsxSplit(';')]` | Property | Splits delimited cell values into a collection (`string[]` or `List<string>`). |
| `[XlsxDefault(...)]` | Property | Sets a fallback value (`Now`, `Today`, `NewGuid`, or static values). |
| `[XlsxConverter(typeof(T))]` | Property | Hooks up a custom converter (`IXlsxConverter`, `IXlsxRowConverter`). |
| `[XlsxNumberFormat("...")]` | Property | Applies native OpenXML numeric formatting (`#,##0.00 $`, `0.00%`). |
| `[XlsxFormula("...")]` | Property | Writes an Excel formula evaluated by spreadsheet engines. |
| `[XlsxCase(...)]` | Property | Enforces casing transformations (`Upper`, `Lower`, `Title`). |
| `[XlsxTrim(...)]` | Property | Configures whitespace trimming policies on string columns. |
| `[XlsxDecimalSeparator(',')]`| Property | Enforces a strict decimal delimiter character for this column. |

### 3. On-The-Fly Data Validation (Declarative Validation)
*In short: Set simple guardrails to reject dirty user inputs early.*

```csharp
public sealed class ReasonRequiredWhenRejected : IXlsxCondition<LeaveRequest>
{
    public bool Evaluate(LeaveRequest item) => item.Status == "REJECTED";
}

[XlsxSerializable]
public class LeaveRequest
{
    [XlsxRequired] // Must not be empty
    [XlsxLength(min: 3, max: 15)] // Between 3 and 15 characters
    public string CaseNumber { get; set; } = string.Empty;

    [XlsxAllowedValues("PENDING", "APPROVED", "REJECTED")] // Strict whitelist
    public string Status { get; set; } = "PENDING";

    [XlsxRange(min: 0.5, max: 45.0)] // Between 0.5 and 45 days
    public double DaysCount { get; set; }

    [XlsxRegex(@"^[A-Z]{2}-\d{4}$")] // Must match format like "HR-2026"
    public string DepartmentCode { get; set; } = string.Empty;

    // Required ONLY if status is "REJECTED"!
    [XlsxRequiredIf(typeof(ReasonRequiredWhenRejected))]
    public string? RejectionReason { get; set; }
}
```

### 4. Dynamic Rules Based on User Context
*In short: Adapt parsing logic dynamically according to user authorization levels.*

```csharp
// Context passed from your application
public sealed record UserContext(bool IsAdmin, decimal MaxSpendingLimit);

public sealed class ApprovalRequiredCondition : IXlsxCondition<Expense, UserContext>
{
    public bool Evaluate(Expense item, UserContext ctx)
        => !ctx.IsAdmin && item.Amount > ctx.MaxSpendingLimit;
}

[XlsxSerializable]
public class Expense
{
    public string Description { get; set; } = string.Empty;
    public decimal Amount { get; set; }

    // Required only if the user is not Admin AND the expense exceeds the threshold
    [XlsxRequiredIf(typeof(ApprovalRequiredCondition))]
    public string? ApprovalCode { get; set; }
}

// Usage:
var myContext = new UserContext(IsAdmin: false, MaxSpendingLimit: 500m);
var options = new XlsxReadOptions<Expense, UserContext>(myContext);

await foreach (var d in FlashXlsxSerializer.ReadAsync("expenses.xlsx", options))
{
    // Validated dynamically based on user identity
}
```

### 5. Variable and Multilingual Column Headers
*In short: Read workbooks interchangeably whether headers are in English, French, or German.*

```csharp
var options = new XlsxReadOptions<Product>()
    .WithHeader(hasHeader: true)
    .WithTranslations(new Dictionary<string, string[]>
    {
        ["Sku"]   = ["Part Number", "Ref", "Item Code", "Référence", "Artikelnummer"],
        ["Name"]  = ["Description", "Label", "Désignation", "Bezeichnung"],
        ["Price"] = ["Unit Price", "Prix HT", "Preis", "Cost"]
    });

var products = await FlashXlsxSerializer.ReadAllAsync("products.xlsx", options);
```

### 6. Crash-Free Import (Fault-Tolerant Error Handling)
*In short: If row 42 has a typo out of 100,000, do not abort the whole import.*

```csharp
var errors = new List<XlsxErrorContext>();

var options = new XlsxReadOptions<Product>()
    .WithErrorMode(XlsxErrorMode.ReportAndContinue) // Continue streaming
    .WithErrorHandler((in XlsxErrorContext err) =>
    {
        // Collect anomalies without allocating exceptions
        errors.Add(err);
    });

await foreach (var p in FlashXlsxSerializer.ReadAsync("products.xlsx", options))
{
    // Valid records arrive here cleanly
}

// Review collected issues:
foreach (var e in errors)
{
    Console.WriteLine($"Row {e.RowNumber}, Column '{e.ColumnName}': invalid value '{e.RawValue}' ({e.Status})");
}
```

### 7. Password-Protected File Detection
*In short: When a user uploads an encrypted file, get notified cleanly without throwing exceptions.*

Encrypted Excel workbooks are binary OLE2 Compound Document containers. FlashXlsx detects their signature upon stream inspection and notifies your error handler without throwing runtime exceptions:

```csharp
XlsxCellStatus? detectedStatus = null;

var options = new XlsxReadOptions()
    .WithErrorMode(XlsxErrorMode.ReportAndContinue)
    .WithErrorHandler((in XlsxErrorContext err) => detectedStatus = err.Status);

await FlashXlsxSerializer.ReadWorkbookAsync(
    "protected_file.xlsx",
    wb => wb.MapSheet<Product>("Products", out var products),
    options);

if (detectedStatus == XlsxCellStatus.EncryptedFile)
{
    Console.WriteLine("Sorry, this workbook is password protected.");
}
```

### 8. Automatic Currency and Non-Breaking Space Cleanup
*In short: Stop parsing failures caused by accounting software formatting numbers as `"$1,250.50"` or `"1 250,50 €"`.*

FlashXlsx sanitizes numeric inputs directly in memory:
* Strips common currency signs: `€`, `$`, `£`, `CHF`, `USD`, `EUR`, `CAD`.
* Strips non-breaking and thin spaces (`\u00A0`, `\u202F`) inserted by regional formatting.
* Automatically balances comma vs. dot decimal delimiters.

A cell holding `" $1,250.50 "` parses directly into a `decimal` property with the value `1250.50m`.

### 9. Schema-Free Dynamic Mode: XlsxValue & Dictionaries (Zero-Boxing)
*In short: Read or write arbitrary Excel files on the fly without declaring C# classes, with value-type performance.*

Unlike legacy libraries relying on `Dictionary<string, object>` (which boxes every integer, boolean, or date onto the heap), FlashXlsx uses **`XlsxValue`**: a stack-allocated discriminated union (`readonly struct`) with zero boxing.

#### A. Dynamic Reading
```csharp
await foreach (Dictionary<string, XlsxValue> row in FlashXlsxSerializer.ReadDictionaryAsync("export.xlsx"))
{
    // Zero-boxing accessors with safe type casting:
    string? customer = row["Customer"].AsString();
    int qty          = row["Quantity"].AsInt32();
    long id          = row["Id"].AsInt64();
    double total     = row["Total"].AsDouble();
    decimal price    = row["Price"].AsDecimal();
    bool isValid     = row["IsValid"].AsBoolean();
    DateOnly shipAt  = row["Shipment"].AsDateOnly();
    TimeOnly slot    = row["Slot"].AsTimeOnly();
    DateTime created = row["CreatedAt"].AsDateTime();
    Guid traceId     = row["TraceGuid"].AsGuid();

    if (!row["OptionalNotes"].IsEmpty) { /* ... */ }
}
```

#### B. Dynamic Writing Without Allocation
Thanks to implicit conversion operators on `XlsxValue`, building rows feels natural:
```csharp
var rows = new List<Dictionary<string, XlsxValue>>
{
    new()
    {
        ["SKU"]       = "PRD-99",                    // implicit string
        ["Price"]     = 129.99m,                     // implicit decimal (zero boxing)
        ["Stock"]     = 42,                          // implicit int (zero boxing)
        ["IsActive"]  = true,                        // implicit bool
        ["UpdatedAt"] = new DateOnly(2026, 11, 10),  // implicit DateOnly
        ["BatchId"]   = Guid.NewGuid()               // implicit Guid
    }
};

await FlashXlsxSerializer.WriteDictionaryAsync("dynamic.xlsx", rows);
```

### 10. Manage Multiple Sheets at Once
*In short: Export or import an entire workbook with "Customers", "Orders", and "Settings" tabs in a single pass using the fluent API.*

```csharp
// --- FLUENT MULTI-SHEET WRITING ---
await FlashXlsxSerializer.WriteWorkbookAsync("management.xlsx", wb => wb
    .AddSheet("Customers", customerList)
    .AddSheet("Orders", orderStream)
    .AddDictionarySheet("Settings", settingsDict));

// --- FLUENT MULTI-SHEET READING ---
List<Customer>? customers = null;
List<Order>? orders = null;
List<Dictionary<string, XlsxValue>>? config = null;

await FlashXlsxSerializer.ReadWorkbookAsync("management.xlsx", wb => wb
    .MapSheet<Customer>("Customers", out customers)
    .MapSheet<Order>("Orders", out orders)
    .MapSheetAsDictionary("Settings", out config));

// Multi-sheet streaming callback (ValueTask)
await FlashXlsxSerializer.ReadWorkbookAsync("management.xlsx", wb => wb
    .MapSheet<Order>("Orders", async order =>
    {
        await ProcessOrderAsync(order);
    }));
```

### 11. Inspect Sheets Without Loading the File
*In short: Discover sheet names, indexes, and visibility states before parsing the full document.*

```csharp
IReadOnlyList<XlsxSheetInfo> sheets = FlashXlsxSerializer.GetSheets("management.xlsx");

foreach (XlsxSheetInfo s in sheets)
{
    Console.WriteLine($"Sheet #{s.Index}: '{s.Name}' (Visible: {s.IsVisible})");
}

// Then selectively stream the desired sheet:
var options = new XlsxReadOptions<Product>().WithSheetName("Products 2026");
await foreach (var row in FlashXlsxSerializer.ReadAsync("management.xlsx", options))
{
    // ...
}
```

### 12. Effortless Merged Cells Reading
*In short: When a category header is merged across 10 rows in Excel, the value is cleanly copied down to each row.*

```csharp
var options = new XlsxReadOptions<SalesReport>()
    .WithFillMergedCells(true); // Automatically propagates merged parent values

await foreach (var row in FlashXlsxSerializer.ReadAsync("sales.xlsx", options))
{
    // 'row.Category' is populated across all merged rows
}
```

### 13. Create Styled Excel Tables with Filters
*In short: Your output workbook opens with a formatted table, alternating banded rows, and header filter controls.*

```csharp
var options = new XlsxWriteOptions<Product>()
    .WithExcelTable(
        tableName: "ProductsTable",
        styleName: "TableStyleMedium2",
        showRowStripes: true,
        enableAutoFilter: true);

await FlashXlsxSerializer.WriteAsync("styled_products.xlsx", list, options);
```

### 14. Fast Preview and Pagination for Large Files
*In short: Load only the first 20 records for an on-screen preview without parsing a 500,000-row file.*

```csharp
var options = new XlsxReadOptions<Product>()
    .WithStartRow(2)  // Start at row 2 (skipping header)
    .WithMaxRows(20); // Terminate reading immediately after 20 rows

var preview = await FlashXlsxSerializer.ReadAllAsync("large_dataset.xlsx", options);
```

### 15. Built-in Formula Injection Shield
*In short: Protects your downstream Excel users against CSV/Spreadsheet Formula Injections (CSV Injection).*

If untrusted input begins with dangerous characters (`=`, `+`, `-`, `@`, `\t`, `\r`) such as `=cmd|' /C calc'!A0`, Excel could execute arbitrary commands upon opening. FlashXlsx intercepts these tokens and applies your chosen security policy (automatic apostrophe quoting or violation rejection) out of the box.

### 16. Custom Converters (IXlsxConverter & Cross-Column Calculations)
*In short: Handle custom data layouts or calculate values using adjacent row columns with zero allocations.*

FlashXlsx provides three tiers of zero-allocation converters:
* **`IXlsxConverter<TProp>`**: Direct cell ↔ property mapping via `ReadOnlySpan<char>`.
* **`IXlsxRowConverter<TRow, TProp>`**: Access other already deserialized properties on the same row.
* **`IXlsxRowConverter<TRow, TContext, TProp>`**: Access row properties and runtime application context simultaneously.

```csharp
// Example: Calculate total including VAT based on existing Subtotal and VAT rate columns
public sealed class VatCalculatorConverter : IXlsxRowConverter<InvoiceLine, decimal>
{
    public bool TryRead(ReadOnlySpan<char> rawValue, in InvoiceLine row, XlsxReadOptions options, out decimal value)
    {
        // Compute dynamically if the cell is absent: Subtotal * (1 + VatRate / 100)
        value = Math.Round(row.Subtotal * (1m + (row.VatRate / 100m)), 2);
        return true;
    }

    public string? Write(in decimal value, in InvoiceLine row, XlsxWriteOptions options)
        => value.ToString("F2", options.Culture);
}

[XlsxSerializable]
public class InvoiceLine
{
    public decimal Subtotal { get; set; }
    public decimal VatRate { get; set; }

    [XlsxConverter(typeof(VatCalculatorConverter))]
    public decimal TotalWithVat { get; set; }
}
```

### 17. Delimited Lists in Cells & Default Values
*In short: Split semicolon-delimited tags into string arrays and inject dynamic default values when cells are blank.*

```csharp
[XlsxSerializable]
public class Article
{
    public string Title { get; set; } = string.Empty;

    // Splits "Tech;Cloud;DotNet" into an array
    [XlsxSplit(separator: ';', TrimEntries = true, RemoveEmptyEntries = true)]
    public string[] Tags { get; set; } = [];

    // Generates a new Guid if the cell is empty
    [XlsxDefault(XlsxDefaultKind.NewGuid)]
    public Guid CorrelationId { get; set; }

    // Injects today's date if omitted
    [XlsxDefault(XlsxDefaultKind.Today)]
    public DateOnly CreatedDate { get; set; }

    // Native Excel accounting format
    [XlsxNumberFormat("#,##0.00 $")]
    public decimal GrossPrice { get; set; }
}
```

### 18. Real-Time Progress Reporting (IProgress)
*In short: Track exact progression when processing multi-gigabyte files to update CLI counters or UI progress bars.*

```csharp
var progress = new Progress<int>(rowsRead =>
{
    Console.Write($"\rProcessed rows: {rowsRead:N0}");
});

var options = new XlsxReadOptions<Product>()
    .WithProgress(progress, intervalRows: 5000); // Trigger every 5,000 rows

await foreach (var p in FlashXlsxSerializer.ReadAsync("massive_import.xlsx", options))
{
    // Fluid processing
}
```

---

## Under the Hood: For Optimization Enthusiasts and Experts

This section breaks down the architectural choices that allow FlashXlsx to process **770,000 rows/second** with under 20 MB of RAM.

```text
Stream .xlsx (ZipArchive)
       │
       ▼
 [ Deflate Decompression (zlib) ]
       │
       ▼  128 KB rented buffer via ArrayPool<byte>
 ┌────────────────────────────────────────────────────────┐
 │            Utf8SheetRowReader                          │
 │  • SIMD Vectorization (SearchValues<byte> - AVX2)      │
 │  • Direct token scanning for <row>, <c>, <v> tags      │
 └─────────────────────────┬──────────────────────────────┘
                           │
      ┌────────────────────┴─────────────────────┐
      ▼                                          ▼
[ Numbers, Dates & Booleans ]           [ Text & Strings ]
  Utf8Parser / Ascii.ToUtf16              Encoding.UTF8 (Shared pooled buffer)
  ZERO heap allocations                   Allocation-free entity decoding
      │                                          │
      └────────────────────┬─────────────────────┘
                           ▼
          [ AOT-Compiled IXlsxTypeMap<T> ]
          • Direct index dispatch (switch jump-table)
          • Zero reflection, zero boxing, zero Activator
          • Inlined validation logic
                           │
                           ▼
                  yield return instance;
```

### The Conveyor Belt Analogy (Streaming vs. DOM)
While traditional DOM-based engines build an in-memory document tree consisting of millions of XML node objects (`XmlNode` or `XElement`), FlashXlsx functions as a **linear forward-only streaming pipeline**. A parsed row is immediately projected into your typed C# entity and yielded to your consumer loop. As soon as the next row is requested, the internal buffers are cleared and reused.

### Vectorized UTF-8 Scanner (`ReadOnlySpan<byte>` & AVX2)
FlashXlsx bypasses `System.Xml.XmlReader` when traversing worksheet rows:
* XLSX files are natively zipped UTF-8 XML payloads.
* Rather than decoding raw UTF-8 bytes into intermediate UTF-16 `string` objects, **`Utf8SheetRowReader`** operates directly on `ReadOnlySpan<byte>` slices supplied by the decompression stream.
* **Hardware Vectorization:** Hardware-accelerated `SearchValues<byte>` scans for markup delimiters (`<` and `>`) in **32-byte chunks per CPU cycle** using AVX2 instructions.
* Cell references (`r="C14"`) are parsed straight from ASCII bytes into row/column integer coordinates using base-26 integer arithmetic without heap allocations.

### ASCII -> UTF-16 SIMD Acceleration
In enterprise datasets, over 80% of cell values are numbers (`12345.67`), ISO dates, booleans, or alphanumeric IDs. FlashXlsx leverages `System.Text.Ascii.ToUtf16`: on modern processors, conversions run in single-pass vector instructions, bypassing multi-byte UTF-8 translation tables.

### Zero Heap Allocations (`ArrayPool`, `XlsxValue` & Stack Structs)
* Core read buffers (128 KB) and transient cell spans are rented from `ArrayPool<byte>.Shared` and `ArrayPool<char>.Shared`.
* `XlsxErrorContext` error descriptors are `readonly record struct` instances passed by reference (`in`). When validation flags an issue, **no `Exception` instance is allocated**, sparing garbage collector pressure.
* **Boxing-Free Dynamic Handling (`XlsxValue`):** In dictionary mode, `XlsxValue` stores integers, floats, booleans, and dates as raw 64-bit payloads (`long _bits`), `decimal`, or `Guid` in a compact struct layout (`[StructLayout(LayoutKind.Auto)]`). Zero Gen0 heap boxing occurs during dynamic reads or writes.

### Roslyn Source Generator (100% Native AOT & Trimming-Safe)
During project compilation (`dotnet build`), the FlashXlsx Roslyn generator analyzes classes tagged with `[XlsxSerializable]` and outputs compile-time C# mapping logic:
1. Direct `CreateInstance()` factory instantiation (`new MyModel()`).
2. An unrolled index-based `switch(propertyIndex)` jump table matching column ordinals to property setters.
3. Inlined validation assertions compiled directly into branching `if` statements.

**Results:**
* Zero `System.Reflection` invocations.
* Zero `Type.GetProperties()` or `PropertyInfo.SetValue()` calls.
* Zero runtime code generation (`Expression.Compile()`).
* **Zero AOT or trimming compatibility warnings** (`IL2026`, `IL3050`).

---

## License

This project is licensed under the terms of the **MIT License**.

You are free to use, modify, and distribute it in all personal, enterprise, or commercial projects without royalty fees.
# DocFlow

> **🎮 Hobby Project** — built for fun and learning. Not production-hardened. PRs and issues welcome.

A lightweight .NET library for converting, reading, and creating documents across **Word, PDF, Excel, CSV, HTML, and Image** formats — with a clean service-based API, three input modes (file · stream · bytes), and full async support.

---

## ✨ Features at a Glance

| Service | What it does |
|---|---|
| `IWordService` | Create, read, template-fill, convert to PDF |
| `IPdfService` | Extract text & images, convert to Word / Excel |
| `IExcelService` | Create, read, template-fill, convert to PDF |
| `ICsvService` | Create, read, convert to/from Excel and PDF |
| `IHtmlService` | Extract text & tables, convert to Word / PDF / Excel |
| `IImageService` | OCR text extraction, embed image in PDF / Word / Excel |
| `IConversionService` | Top-level orchestrator — routes all of the above |

### Supported Conversions

```
Word  → PDF       PDF   → Word      Excel → PDF
PDF   → Excel     CSV   → Excel     Excel → CSV
CSV   → PDF       HTML  → Word      HTML  → PDF
HTML  → Excel     Image → PDF       Image → Word
Image → Excel
```

### Three Input Modes — Everywhere

Every method comes in three overloads plus async variants:

```csharp
// File-based
await _wordService.CreateWordAsync("output.docx", content);

// Stream-based
await _wordService.CreateWordAsync(outputStream, content);

// Bytes-based (great for HTTP responses / blob storage)
byte[] bytes = await _wordService.CreateWordAsync(content);
```

---

## 📦 Installation

Clone and reference `DocFlow.Core` directly — no NuGet package yet (hobby project, maybe someday 🙂).

```bash
git clone https://github.com/forgeschema-star/docflow.git
```

Then add a project reference in your `.csproj`:

```xml
<ProjectReference Include="../docflow/DocFlow.Core/DocFlow.Core.csproj" />
```

### Target Frameworks

`DocFlow.Core` targets **`net48`** and **`netstandard2.0`**, so it works with:
- .NET Framework 4.8+
- .NET Core / .NET 5 / .NET 6 / .NET 8+
- ASP.NET Core

---

## ⚙️ Setup

### With Dependency Injection (ASP.NET Core)

```csharp
// Program.cs / Startup.cs
builder.Services.AddDocFlowCore();

// Or with custom settings
var settings = new DocFlowSettings
{
    TempDirectory    = "/tmp/docflow",
    OcrDataPath      = "/usr/share/tessdata",   // required for OCR
    MaxFileSizeBytes = 50_000_000,              // 50 MB
    LoggingEnabled   = true,
    AllowOverwrite   = false,
};
builder.Services.AddDocFlowCore(myLogger, settings);
```

### Without DI (console / legacy)

```csharp
var settings     = DocFlowSettings.CreateDefault();
var logger       = new ConsoleLogger();           // implement ILogger
var wordService  = new WordService(logger, settings);
var excelService = new ExcelService(logger, settings);
var pdfService   = new PdfService(wordService, excelService, logger, settings);
// ... wire up remaining services
var converter    = new ConversionService(wordService, excelService, pdfService,
                       csvService, htmlService, imageService, logger, settings);
```

---

## 🚀 Quick Start Examples

### Create a Word document

```csharp
byte[] doc = await _wordService.CreateWordAsync("Hello, DocFlow!");
```

### Fill a template

```csharp
// Template contains {{CustomerName}}, {{InvoiceDate}}, {{Total}}
var placeholders = new Dictionary<string, string>
{
    { "CustomerName", "Acme Corp"   },
    { "InvoiceDate",  "2024-06-10"  },
    { "Total",        "$4,500.00"   },
};
byte[] filled = await _wordService.ReplacePlaceholdersAsync(templateBytes, placeholders);
```

### Convert Word → PDF

```csharp
var result = await _conversionService.ConvertAsync(
    DocumentType.Word, DocumentType.Pdf, inputBytes);

if (result.Success)
    File.WriteAllBytes("output.pdf", result.OutputBytes);
else
    Console.WriteLine($"[{result.ErrorCode}] {result.Message}");
```

### Read an Excel file

```csharp
List<Dictionary<string, string>> rows =
    await _excelService.ReadExcelAsync("report.xlsx");

foreach (var row in rows)
    Console.WriteLine($"{row["Name"]} — {row["Department"]}");
```

### OCR an image

```csharp
// Requires OcrDataPath set in DocFlowSettings
string text = await _imageService.ReadImageTextAsync("scan.png");
```

---

## 📋 Error Handling

```csharp
var result = await _conversionService.ConvertAsync(from, to, bytes);

switch (result.ErrorCode)
{
    case ConversionErrorCode.None:              /* success */    break;
    case ConversionErrorCode.FileTooLarge:      /* > MaxBytes */ break;
    case ConversionErrorCode.FileNotFound:      /* missing */    break;
    case ConversionErrorCode.OutputAlreadyExists:               break;
    case ConversionErrorCode.UnsupportedConversion:             break;
    case ConversionErrorCode.ProcessingFailed:  /* general */   break;
}
```

---

## 🛠 Dependencies

DocFlow.Core is built on top of these open-source libraries:

| Library | Version | Purpose | License |
|---|---|---|---|
| [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) | 2.20.0 | Read/write `.docx` files | MIT |
| [ClosedXML](https://github.com/ClosedXML/ClosedXML) | 0.104.2 | Read/write `.xlsx` files | MIT |
| [PdfSharpCore](https://github.com/ststeiger/PdfSharpCore) | 1.3.62 | Generate PDF documents (includes MigraDocCore for layouts) | MIT |
| [HtmlAgilityPack](https://html-agility-pack.net) | 1.12.4 | Parse HTML, extract text & tables | MIT |
| [Tesseract](https://github.com/charlesw/tesseract) | 5.2.0 | OCR text extraction from images | Apache 2.0 |
| [UglyToad.PdfPig](https://github.com/UglyToad/PdfPig) | 0.1.10 | Read PDF text & extract images | Apache 2.0 |

---

## 📐 Architecture

```
DocFlow.Core/
├── Interfaces/          IConversionService, IWordService, IExcelService, ...
├── Services/            ConversionService, WordService, ExcelService, ...
├── Models/              DocumentType, ConversionResult, DocFlowSettings, ...
├── Helpers/             CsvHelper, HtmlHelper, PlaceholderHelper, ...
├── Factory/             DocumentFactory
└── Extensions/          ServiceCollectionExtensions (AddDocFlowCore)
```

---

## 🧪 Running the Demo

The [docflow-src](https://github.com/forgeschema-star/docflow-src) private repo contains `DocFlow.ConsoleDemo` — a smoke-test runner that exercises the full conversion pipeline:

```bash
dotnet run --project DocFlow.ConsoleDemo
```

Scenarios covered: valid workflow, file-not-found, overwrite protection, file-size limit, CSV/HTML structured formats.

---

## 📄 License

MIT — see [LICENSE](LICENSE).

---

## 🎮 Hobby Project Notice

This is a **personal hobby project** built in spare time to explore .NET document processing.

- No SLA, no production support guarantee
- Breaking changes may happen without notice
- Issues and PRs are welcome — I'll do my best to respond
- Feel free to fork and adapt for your own projects

> *If you find it useful, drop a ⭐ — it means a lot!*

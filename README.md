# Nedev.FileConverters.DocxToDoc

A high-performance .NET library for converting OpenXML `.docx` documents to legacy binary `.doc` format. The package targets **.NET 8.0** and **.NET Standard 2.1** and exposes both a reusable class library and a simple command‑line interface (CLI).

> **Prerequisite**: the library depends on [`Nedev.FileConverters.Core`](https://www.nuget.org/packages/Nedev.FileConverters.Core) for shared utilities. The CLI project references the library, so installing the main package transitively brings in the core dependency.

---

## Features

- Fast, zero-external‑dependency conversion from `.docx` to `.doc`.
- Supports large documents with minimal memory overhead.
- Multi‑targeted for broad platform support.
- Includes an MIT‑licensed CLI for batch processing or scripting.
- Packaged with rich metadata (README & LICENSE) for NuGet.

### Feature completeness

The converter implements the following capabilities:

| Area | Status |
|------|--------|
| Paragraphs and runs | ✅ Complete |
| Tables | ✅ Complete |
| Styles and fonts | ✅ Complete |
| Sections and headers/footers | ✅ Complete |
| Numbering and lists | ✅ Complete |
| Images (embedded) | ⚠️ Partial (extracted but not re‑embedded in DOC) |
| Complex fields (TOC, hyperlinks) | ⚠️ Partial |

This project aims to faithfully represent most common Word constructs; however, very advanced features such as smart art, equations or macros are not supported. Contributions to expand coverage are welcome.

---

## Installation

Use NuGet to install the library in your project:

```powershell
dotnet add package Nedev.FileConverters.DocxToDoc --version 0.1.0
```

or add the package reference directly to your `.csproj`:

```xml
<PackageReference Include="Nedev.FileConverters.DocxToDoc" Version="0.1.0" />
```

The `Nedev.FileConverters.Core` dependency will be resolved automatically.

---

## Library Usage

```csharp
using Nedev.FileConverters.DocxToDoc;

// convert files on disk
var converter = new DocxToDocConverter();
converter.Convert("input.docx", "output.doc");

// or work with streams
using var inStream = File.OpenRead("input.docx");
using var outStream = File.Create("output.doc");
converter.Convert(inStream, outStream);
```

The converter throws `ArgumentNullException` for invalid paths and wraps I/O exceptions for other failures.

---

## CLI Usage

The CLI is a thin wrapper around the library. Build it with:

```powershell
cd src\Nedev.FileConverters.DocxToDoc.Cli
dotnet build -c Release
```

Run directly from the build output:

```powershell
dotnet bin\Release\net8.0\Nedev.FileConverters.DocxToDoc.Cli.dll <input.docx> <output.doc>
```

It returns `0` on success, non‑zero codes for missing arguments or errors. Use `-h` or `--help` to display usage.

---

## Development

Clone the repo and open it in Visual Studio or VS Code.

```powershell
git clone <repository-url>
cd Nedev.FileConverters.DocxToDoc
dotnet restore
```

Projects are located under `src/`:

- `Nedev.FileConverters.DocxToDoc` – core library
- `Nedev.FileConverters.DocxToDoc.Cli` – executable frontend
- `Nedev.FileConverters.DocxToDoc.Tests` – unit tests (xUnit)

Run tests with `dotnet test`.

---

## Packaging

A `.nuspec` file is not required; `dotnet pack` already includes README and LICENSE. Example:

```powershell
dotnet pack src\Nedev.FileConverters.DocxToDoc\Nedev.FileConverters.DocxToDoc.csproj -c Release -o nupkg
```

The resulting package is versioned `0.1.0` and can be pushed to NuGet.org.

---

## License

This project is licensed under the **MIT License** – see [LICENSE](LICENSE) for details.

---

## Contributing

Contributions are welcome! Please fork the repo, make changes on a feature branch, and open a pull request. Adhere to the existing coding style and update tests as appropriate.

---

## Author

Developed by Nedev – feel free to reach out on GitHub with issues or suggestions.

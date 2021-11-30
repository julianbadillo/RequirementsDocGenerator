# Requirements Document Generator

Generates a MS Word document from an MS Excel spreadsheet with a list of functional requirements.

## Dependencies

- .Net Framework 2.0+
- OpenXML.DocumentFormat <https://www.nuget.org/packages/DocumentFormat.OpenXml/>

## Download and Build

This was developed with VS Code. To download the code:

```powershell
git clone https://github.com/julianbadillo/RequirementsDocGenerator.git
```

If you open it with VS Code, building happens automatically, or you can build by:

```powershell
cd RequirementsDocGenerator
dotnet build
```

## Usage

### Sample Data

There's a sample spreadsheet on `data/requirements_sample.xlsx`, with the expected columns.

### From Terminal

```powershell
cd RequirementsDocGenerator
dotnet run <INPUT FILE> <OUTPUT FILE>
```

### From C\#

Use the `DocGenerator` class:

```C#
using RequirementsDocGenerator;
...

var gen = new DocGenerator();
gen.Generate(inputFile, outputFile);

```

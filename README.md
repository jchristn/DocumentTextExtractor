![documenttextextractor](https://raw.githubusercontent.com/jchristn/DocumentTextExtractor/main/assets/logo.ico) 

# DocumentTextExtractor

## Simple C# library for extracting text and metadata from .docx, .pptx, and .xlsx files

[![NuGet Version](https://img.shields.io/nuget/v/DocumentTextExtractor.svg?style=flat)](https://www.nuget.org/packages/DocumentTextExtractor/) [![NuGet](https://img.shields.io/nuget/dt/DocumentTextExtractor.svg)](https://www.nuget.org/packages/DocumentTextExtractor)    

DocumentTextExtractor provides simple methods for extracting text and metadata from .docx, .pptx, and .xlsx files.

## New in v1.0.x

- Initial release
- Support for `docx`, `pptx`, `xlsx`, and `pdf`
- Contextual extraction for `pptx`, `xlsx`

## Disclaimer

This library has been tested on a limited set of documents.  It is highly likely that documents exist this from which the library, in its current state, cannot extract text.

The PDF implementation relies upon [PDFSharp](https://github.com/empira/PDFsharp) and [PDFPlumber](https://github.com/jsvine/pdfplumber).  The latter is written in Python and requires that you have installed Python and used `pip` to install the `pdfplumber` package.

## Simple Examples

Refer to the `Test` project for a full example.

```csharp
using DocumentTextExtractor;

void Main(string[] args)
{
  using (DocxTextExtractor docx = new DocxTextExtractor("./temp/", "mydocument.docx"))
  {
    string docxText = docx.ExtractText();
    Dictionary<string, string> docxMetadata = docx.ExtractMetadata();
  }

  using (PptxTextExtractor pptx = new PptxTextExtractor("./temp/", "mypresentation.pptx"))
  {
    string pptxText = pptx.ExtractText();
    Dictionary<string, string> pptxMetadata = pptx.ExtractMetadata();
  }

  using (XlsxTextExtractor xlsx = new XlsxTextExtractor("./temp/", "myspreadsheet.xlsx"))
  {
    string xlsxText = xlsx.ExtractText();
    Dictionary<string, string> xlsxMetadata = xlsx.ExtractMetadata();
  }

  using (PdfTextExtractor pdf = new PdfTextExtractor("myfile.pdf"))
  {
    string pdfText = pdf.ExtractText();
    Dictionary<string, string> pdfMetadata = pdf.ExtractMetadata();
  }
}
```

## Contextual Extraction

For certain document types (e.g. `pptx` and `xlsx`) text can be extracted with an identifier for the slide or sheet number associated with the document.

```csharp
using (PptxTextExtractor pptx = new PptxTextExtractor("./temp/", "mypresentation.pptx"))
{
  IEnumerable<KeyValuePair<int, string>> slideContent = pptx.ExtractTextBySlide();
  foreach (KeyvaluePair<int, string> kvp in slideContent)
  {
    Console.WriteLine("Slide " + kvp.Key + ": " + kvp.Value);
  }
}

using (XlsxTextExtractor xlsx = new XlsxTextExtractor("./temp/", "mypresentation.pptx"))
{
  IEnumerable<KeyValuePair<int, string>> sheetContent = xlsx.ExtractTextBySheet();
  foreach (KeyvaluePair<int, string> kvp in sheetContent)
  {
    Console.WriteLine("Sheet " + kvp.Key + ": " + kvp.Value);
  }
}
```

## Version History

Please refer to CHANGELOG.md.

![documenttextextractor](https://raw.githubusercontent.com/jchristn/DocumentTextExtractor/main/assets/logo.ico) 

# DocumentTextExtractor

## Simple C# library for extracting text and metadata from .docx, .pptx, and .xlsx files

[![NuGet Version](https://img.shields.io/nuget/v/DocumentTextExtractor.svg?style=flat)](https://www.nuget.org/packages/DocumentTextExtractor/) [![NuGet](https://img.shields.io/nuget/dt/DocumentTextExtractor.svg)](https://www.nuget.org/packages/DocumentTextExtractor)    

DocumentTextExtractor provides simple methods for extracting text and metadata from .docx, .pptx, and .xlsx files.

## New in v1.0.x

- Initial release
- Support for ```docx```, ```pptx```, and ```xlsx```

## Disclaimer

This library has been tested on a limited set of documents.  It is highly likely that documents exist this from which the library, in its current state, cannot extract text.

## Simple Examples

Refer to the ```Test``` project for a full example.

```csharp
using DocumentTextExtractor;

void Main(string[] args)
{
  using (DocxTextExtractor docx = new DocxTextExtractor("./temp/", "mydocument.docx"))
  {
    string docxText = docx.ExtractText();
    Dictionary<string, string> docxMetadata = docx.ExtractMetadata();
  }

  using (PptxTextExtractor pptx = new DocxTextExtractor("./temp/", "mypresentation.pptx"))
  {
    string pptxText = pptx.ExtractText();
    Dictionary<string, string> pptxMetadata = pptx.ExtractMetadata();
  }

  using (XlsxTextExtractor xlsx = new XlsxTextExtractor("./temp/", "mypresentation.pptx"))
  {
    string xlsxText = xlsx.ExtractText();
    Dictionary<string, string> xlsxMetadata = xlsx.ExtractMetadata();
  }
}
```

## Version History

Please refer to CHANGELOG.md.

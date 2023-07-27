![documentparser](https://raw.githubusercontent.com/jchristn/DocumentParser/main/assets/logo.ico)

# DocumentParser

## Simple C# library for extracting text and metadata from .docx, .pptx, and .xlsx files

[![NuGet Version](https://img.shields.io/nuget/v/DocumentParser.svg?style=flat)](https://www.nuget.org/packages/DocumentParser/) [![NuGet](https://img.shields.io/nuget/dt/DocumentParser.svg)](https://www.nuget.org/packages/DocumentParser)    

DocumentParser provides simple methods for extracting text and metadata from .docx, .pptx, and .xlsx files.

## New in v1.0.x

- Initial release
- Support for ```docx```, ```pptx```, and ```xlsx```

## Disclaimer

This library has been tested on a limited set of documents.  It is highly likely that documents exist this from which the library, in its current state, cannot extract text.

## Simple Examples

Refer to the ```Test.DocumentParser``` project for a full example.

```csharp
using DocumentParser;

void Main(string[] args)
{
  using (DocxParser docx = new DocxParser("./temp/", "mydocument.docx"))
  {
    string docxText = docx.ExtractText();
    Dictionary<string, string> docxMetadata = docx.ExtractMetadata();
  }

  using (PptxParser pptx = new DocxParser("./temp/", "mypresentation.pptx"))
  {
    string pptxText = pptx.ExtractText();
    Dictionary<string, string> pptxMetadata = pptx.ExtractMetadata();
  }

  using (XlsxParser xlsx = new XlsxParser("./temp/", "mypresentation.pptx"))
  {
    string xlsxText = xlsx.ExtractText();
    Dictionary<string, string> xlsxMetadata = xlsx.ExtractMetadata();
  }
}
```

## Version History

Please refer to CHANGELOG.md.

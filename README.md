![documentparser](https://raw.githubusercontent.com/jchristn/DocumentParser/main/assets/logo.ico)

# DocumentParser

## Simple C# library for extracting text and metadata from .docx and .pptx files

[![NuGet Version](https://img.shields.io/nuget/v/DocumentParser.svg?style=flat)](https://www.nuget.org/packages/DocumentParser/) [![NuGet](https://img.shields.io/nuget/dt/DocumentParser.svg)](https://www.nuget.org/packages/DocumentParser)    

DocumentParser provides simple methods for extracting text and metadata from .docx files.

## New in v1.0.x

- Initial release
- Support for ```docx``` and ```pptx```

## Help or Feedback

Need help or have feedback? Please file an issue here!

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
}
```

## Version History

Please refer to CHANGELOG.md.

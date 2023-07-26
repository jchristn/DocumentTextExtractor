![documentparser](https://github.com/jchristn/documentparser/blob/master/assets/logo.ico)

# DocumentParser

## Simple C# library for extracting text and metadata from .docx files

[![NuGet Version](https://img.shields.io/nuget/v/DocumentParser.svg?style=flat)](https://www.nuget.org/packages/DocumentParser/) [![NuGet](https://img.shields.io/nuget/dt/DocumentParser.svg)](https://www.nuget.org/packages/DocumentParser)    

DocumentParser provides simple methods for extracting text and metadata from .docx files.

## New in v1.0.x

- Initial release

## Help or Feedback

Need help or have feedback? Please file an issue here!

## Simple Examples

Refer to the ```Test.DocumentParser``` project for a full example.

```csharp
using DocumentParser;

void Main(string[] args)
{
  using (DocxParser parser = new DocxParser("./temp/", "mydocument.docx"))
  {
    string text = parser.ExtractText();
    Dictionary<string, string> metadata = parser.ExtractMetadata();
  }
}
```

## Version History

Please refer to CHANGELOG.md.

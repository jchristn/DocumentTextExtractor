namespace DocumentParser
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Pptx text extractor.
    /// </summary>
    public class PptxTextExtractor : IDocumentTextExtractor, IDisposable
    {
        #region Public-Members

        #endregion

        #region Private-Members

        private const string _WXmlNamespace = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string _CpXmlNamespace = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string _DcXmlNamespace = @"http://purl.org/dc/elements/1.1/";
        private const string _AXmlNamespace = @"http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string _RXmlNamespace = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string _PXmlNamespace = @"http://schemas.openxmlformats.org/presentationml/2006/main";

        private const string _SlidesSubdirectory = "ppt/slides/";
        private const string _DocumentBodyXPath = "/p:sld/p:cSld/p:spTree";

        private const string _MetadataFile = "docProps/core.xml";
        private const string _MetadataXPath = "/cp:coreProperties";

        #endregion

        #region Constructors-and-Factories

        /// <summary>
        /// Instantiate.
        /// </summary>
        /// <param name="tempDirectory">Base temp directory.</param>
        /// <param name="filename">Filename.</param>
        public PptxTextExtractor(string tempDirectory, string filename)
        {
            if (String.IsNullOrEmpty(tempDirectory)) throw new ArgumentNullException(nameof(tempDirectory));
            if (String.IsNullOrEmpty(filename)) throw new ArgumentNullException(nameof(filename));

            TempDirectory = tempDirectory;
            Filename = filename;

            using (ZipArchive archive = ZipFile.OpenRead(Filename))
            {
                archive.ExtractToDirectory(TempDirectory);
            }
        }

        #endregion

        #region Public-Methods

        /// <summary>
        /// Dispose of resources.
        /// </summary>
        public void Dispose()
        {
            RecursiveDelete(DirInfo, true);
            Directory.Delete(TempDirectory, true);
        }

        /// <summary>
        /// Extract metadata from document.
        /// </summary>
        /// <returns>Dictionary containing metadata.</returns>
        public override Dictionary<string, string> ExtractMetadata()
        {
            Dictionary<string, string> ret = new Dictionary<string, string>();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.PreserveWhitespace = true;
            xmlDoc.Load(TempDirectory + _MetadataFile);

            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", _WXmlNamespace);
            nsmgr.AddNamespace("cp", _CpXmlNamespace);
            nsmgr.AddNamespace("dc", _DcXmlNamespace);
            nsmgr.AddNamespace("a", _AXmlNamespace);
            nsmgr.AddNamespace("r", _RXmlNamespace);
            nsmgr.AddNamespace("p", _PXmlNamespace);

            foreach (XmlNode node in xmlDoc.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element) continue;

                if (node.HasChildNodes)
                {
                    foreach (XmlNode child in node.ChildNodes)
                    {
                        ret.Add(child.LocalName, child.InnerText);
                    }
                }
            }

            return ret;
        }

        /// <summary>
        /// Extract text from document.
        /// </summary>
        /// <returns>Text contents.</returns>
        public override string ExtractText()
        {
            // see https://www.codeproject.com/Articles/20529/Using-DocxToText-to-Extract-Text-from-DOCX-Files

            string root = TempDirectory + _SlidesSubdirectory;
            StringBuilder sb = new StringBuilder();

            FileInfo[] files = new DirectoryInfo(TempDirectory + _SlidesSubdirectory).GetFiles();
            SortedDictionary<int, FileInfo> filesOrdered = new SortedDictionary<int, FileInfo>();

            foreach (FileInfo fi in files)
            {
                if (fi.Name.StartsWith("slide") && fi.Name.EndsWith(".xml"))
                {
                    string temp = fi.Name.Replace("slide", "").Replace(".xml", "");
                    int num = Convert.ToInt32(temp);
                    filesOrdered.Add(num, fi);
                }
            }

            foreach (KeyValuePair<int, FileInfo> kvp in filesOrdered)
            {
                FileInfo fi = kvp.Value;

                if (fi.Name.StartsWith("slide") && fi.Name.EndsWith(".xml"))
                {
                    // Console.WriteLine("File " + fi.Name);

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.PreserveWhitespace = true;
                    xmlDoc.Load(fi.FullName);

                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                    nsmgr.AddNamespace("w", _WXmlNamespace);
                    nsmgr.AddNamespace("cp", _CpXmlNamespace);
                    nsmgr.AddNamespace("dc", _DcXmlNamespace);
                    nsmgr.AddNamespace("a", _AXmlNamespace);
                    nsmgr.AddNamespace("r", _RXmlNamespace);
                    nsmgr.AddNamespace("p", _PXmlNamespace);

                    XmlNode node = xmlDoc.DocumentElement.SelectSingleNode(_DocumentBodyXPath, nsmgr);
                    sb.Append(ReadNode(node));
                    // sb.Append(Environment.NewLine);
                }
            }

            string ret = sb.ToString();

            while (ret.Contains("  ")) ret = ret.Replace("  ", " ");
            while (ret.Contains(Environment.NewLine + Environment.NewLine)) 
                ret = ret.Replace(
                    Environment.NewLine + Environment.NewLine, 
                    Environment.NewLine);
            
            return ret;
        }

        /// <summary>
        /// Extract text from document, delivered as a dictionary where the key is the slide number.
        /// </summary>
        /// <returns>Enumerable of key-value pairs, where the key is the slide number, and the value is the text content.</returns>
        public IEnumerable<KeyValuePair<int, string>> ExtractTextBySlide()
        {
            // see https://www.codeproject.com/Articles/20529/Using-DocxToText-to-Extract-Text-from-DOCX-Files

            string root = TempDirectory + _SlidesSubdirectory;
            StringBuilder sb = new StringBuilder();

            FileInfo[] files = new DirectoryInfo(TempDirectory + _SlidesSubdirectory).GetFiles();
            SortedDictionary<int, FileInfo> filesOrdered = new SortedDictionary<int, FileInfo>();

            foreach (FileInfo fi in files)
            {
                if (fi.Name.StartsWith("slide") && fi.Name.EndsWith(".xml"))
                {
                    string temp = fi.Name.Replace("slide", "").Replace(".xml", "");
                    int num = Convert.ToInt32(temp);
                    filesOrdered.Add(num, fi);
                }
            }

            foreach (KeyValuePair<int, FileInfo> kvp in filesOrdered)
            {
                FileInfo fi = kvp.Value;

                if (fi.Name.StartsWith("slide") && fi.Name.EndsWith(".xml"))
                {
                    // Console.WriteLine("File " + fi.Name);

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.PreserveWhitespace = true;
                    xmlDoc.Load(fi.FullName);

                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                    nsmgr.AddNamespace("w", _WXmlNamespace);
                    nsmgr.AddNamespace("cp", _CpXmlNamespace);
                    nsmgr.AddNamespace("dc", _DcXmlNamespace);
                    nsmgr.AddNamespace("a", _AXmlNamespace);
                    nsmgr.AddNamespace("r", _RXmlNamespace);
                    nsmgr.AddNamespace("p", _PXmlNamespace);

                    XmlNode node = xmlDoc.DocumentElement.SelectSingleNode(_DocumentBodyXPath, nsmgr);
                    
                    string content = ReadNode(node);
                    
                    while (content.Contains("  ")) content = content.Replace("  ", " ");
                    while (content.Contains(Environment.NewLine + Environment.NewLine))
                        content = content.Replace(
                            Environment.NewLine + Environment.NewLine, 
                            Environment.NewLine);

                    yield return new KeyValuePair<int, string>(kvp.Key, content + Environment.NewLine);
                }
            }
        }

        #endregion

        #region Private-Methods

        private string ReadNode(XmlNode node)
        {
            if (node == null || node.NodeType != XmlNodeType.Element)
                return string.Empty;

            StringBuilder sb = new StringBuilder();

            foreach (XmlNode child in node.ChildNodes)
            {
                if (child.NodeType != XmlNodeType.Element) continue;
                if (String.IsNullOrEmpty(child.InnerText)) continue;

                // Console.WriteLine(child.NodeType.ToString() + " " + child.Name + " " + child.LocalName + " " + child.InnerText);

                switch (child.LocalName)
                {   
                    // case "p":   // Paragraph
                    // case "r":   // Run
                    case "t":   // Text
                        sb.Append(child.InnerText + " ");
                        sb.Append(ReadNode(child));
                        break;

                    case "cr":                          // Carriage return
                    case "br":                          // Page break
                        sb.Append(Environment.NewLine);
                        break;

                    default:
                        sb.Append(ReadNode(child));
                        break;
                }
            }

            // sb.Append(Environment.NewLine);
            return sb.ToString();
        }

        #endregion
    }
}

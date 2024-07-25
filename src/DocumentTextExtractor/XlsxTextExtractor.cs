namespace DocumentParser
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Xlsx text extractor.
    /// </summary>
    public class XlsxTextExtractor : IDocumentTextExtractor, IDisposable
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

        private const string _SheetsSubdirectory = "xl/worksheets/";
        private const string _DocumentBodyXPath = "/worksheet";

        private const string _MetadataFile = "docProps/core.xml";
        private const string _MetadataXPath = "/cp:coreProperties";

        private const string _SharedStringsFile = "xl/sharedStrings.xml";
        private const string _SharedStringsXPath = "/sst";

        private string[] _SharedStrings = null;

        #endregion

        #region Constructors-and-Factories

        /// <summary>
        /// Instantiate.
        /// </summary>
        /// <param name="tempDirectory">Base temp directory.</param>
        /// <param name="filename">Filename.</param>
        public XlsxTextExtractor(string tempDirectory, string filename)
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
            _SharedStrings = ExtractSharedStrings();

            // see https://www.codeproject.com/Articles/20529/Using-DocxToText-to-Extract-Text-from-DOCX-Files

            string root = TempDirectory + _SheetsSubdirectory;
            StringBuilder sb = new StringBuilder();

            FileInfo[] files = new DirectoryInfo(TempDirectory + _SheetsSubdirectory).GetFiles();
            SortedDictionary<int, FileInfo> filesOrdered = new SortedDictionary<int, FileInfo>();

            foreach (FileInfo fi in files)
            {
                if (fi.Name.StartsWith("sheet") && fi.Name.EndsWith(".xml"))
                {
                    string temp = fi.Name.Replace("sheet", "").Replace(".xml", "");
                    int num = Convert.ToInt32(temp);
                    filesOrdered.Add(num, fi);
                }
            }

            foreach (KeyValuePair<int, FileInfo> kvp in filesOrdered)
            {
                FileInfo fi = kvp.Value;

                if (fi.Name.StartsWith("sheet") && fi.Name.EndsWith(".xml"))
                {
                    // Console.WriteLine("File " + fi.Name);

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.PreserveWhitespace = true;
                    xmlDoc.Load(fi.FullName);

                    foreach (XmlNode node in xmlDoc.ChildNodes)
                    {
                        // Console.WriteLine(node.NodeType.ToString() + " " + node.Name + " " + node.LocalName + " " + node.InnerText);
                        sb.Append(ReadNode(node));
                        sb.Append(Environment.NewLine);
                    }
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
        /// Extract text from document, delivered as a dictionary where the key is the sheet number.
        /// </summary>
        /// <returns>Enumerable of key-value pairs, where the key is the sheet number, and the value is the text content.</returns>
        public IEnumerable<KeyValuePair<int, string>> ExtractTextBySheet()
        {
            _SharedStrings = ExtractSharedStrings();

            // see https://www.codeproject.com/Articles/20529/Using-DocxToText-to-Extract-Text-from-DOCX-Files

            string root = TempDirectory + _SheetsSubdirectory;

            FileInfo[] files = new DirectoryInfo(TempDirectory + _SheetsSubdirectory).GetFiles();
            SortedDictionary<int, FileInfo> filesOrdered = new SortedDictionary<int, FileInfo>();

            foreach (FileInfo fi in files)
            {
                if (fi.Name.StartsWith("sheet") && fi.Name.EndsWith(".xml"))
                {
                    string temp = fi.Name.Replace("sheet", "").Replace(".xml", "");
                    int num = Convert.ToInt32(temp);
                    filesOrdered.Add(num, fi);
                }
            }

            foreach (KeyValuePair<int, FileInfo> kvp in filesOrdered)
            {
                FileInfo fi = kvp.Value;

                if (fi.Name.StartsWith("sheet") && fi.Name.EndsWith(".xml"))
                {
                    // Console.WriteLine("File " + fi.Name);

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.PreserveWhitespace = true;
                    xmlDoc.Load(fi.FullName);

                    StringBuilder sb = new StringBuilder();

                    foreach (XmlNode node in xmlDoc.ChildNodes)
                    {
                        // Console.WriteLine(node.NodeType.ToString() + " " + node.Name + " " + node.LocalName + " " + node.InnerText);
                        sb.Append(ReadNode(node));
                        sb.Append(Environment.NewLine);
                    }

                    string content = sb.ToString();

                    while (content.Contains("  ")) content = content.Replace("  ", " ");
                    while (content.Contains(Environment.NewLine + Environment.NewLine))
                        content = content.Replace(
                            Environment.NewLine + Environment.NewLine,
                            Environment.NewLine);

                    yield return new KeyValuePair<int, string>(kvp.Key, content);
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
                // Console.WriteLine("  " + child.NodeType.ToString() + " " + child.Name + " " + child.LocalName + " " + child.InnerXml);
                if (child.NodeType != XmlNodeType.Element) continue;
                if (String.IsNullOrEmpty(child.InnerText)) continue;

                if (child.LocalName == "c") // content
                {
                    bool isFormula = false;
                    string formula = null;

                    XmlNodeList contents = child.ChildNodes;
                    foreach (XmlNode content in contents)
                    {
                        if (content.LocalName == "f")
                        {
                            isFormula = true;
                            formula = content.InnerText;
                        }
                        else if (content.LocalName == "v")
                        {
                            if (isFormula)
                            {
                                sb.Append(content.InnerText + " ");
                            }
                            else
                            {
                                if (_SharedStrings != null)
                                {
                                    if (Int32.TryParse(content.InnerText, out int val))
                                    {
                                        if (_SharedStrings.Length > val)
                                        {
                                            sb.Append(_SharedStrings[val] + " ");
                                            sb.Append(ReadNode(child));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    sb.Append(ReadNode(child));
                }
            }

            return sb.ToString();
        }

        private string[] ExtractSharedStrings()
        {
            List<string> ret = new List<string>();

            if (File.Exists(TempDirectory + _SharedStringsFile))
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.PreserveWhitespace = true;
                xmlDoc.Load(TempDirectory + _SharedStringsFile);

                foreach (XmlNode node in xmlDoc.ChildNodes)
                {
                    if (node.NodeType != XmlNodeType.Element) continue;

                    if (node.HasChildNodes)
                    {
                        foreach (XmlNode child in node.ChildNodes)
                        {
                            ret.Add(child.InnerText);
                        }
                    }
                }
            }

            string[] retArray = ret.ToArray();
            // for (int i = 0; i < retArray.Length; i++) Console.WriteLine(i + ": " + retArray[i]);

            return retArray;
        }

        #endregion
    }
}

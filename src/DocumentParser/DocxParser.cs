using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml;
using System.Xml.Linq;
using DocumentParser;
using XmlToPox;

namespace DocumentParser
{
    /// <summary>
    /// Docx parser.
    /// </summary>
    public class DocxParser : IDocumentParser, IDisposable
    {
        #region Public-Members

        /// <summary>
        /// Serialization helper.
        /// </summary>
        public SerializationHelper Serializer
        {
            get
            {
                return _Serializer;
            }
            set
            {
                if (value == null) throw new ArgumentNullException(nameof(Serializer));
                _Serializer = value;
            }
        }

        /// <summary>
        /// Temporary directory.
        /// </summary>
        public string TempDirectory
        {
            get
            {
                return _TempDirectory;
            }
        }

        /// <summary>
        /// Directory info.
        /// </summary>
        public DirectoryInfo DirInfo
        {
            get
            {
                return _DirInfo;
            }
        }

        /// <summary>
        /// Filename.
        /// </summary>
        public string Filename
        {
            get
            {
                return _Filename;
            }
        }

        #endregion

        #region Private-Members

        private Guid _Guid = Guid.NewGuid();
        private SerializationHelper _Serializer = new SerializationHelper();
        private string _TempDirectory = null;
        private DirectoryInfo _DirInfo = null;
        private string _Filename = null;

        private const string _ContentTypeNamespace =
            @"http://schemas.openxmlformats.org/package/2006/content-types";

        private const string _WordXmlNamespace =
            @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private const string _WordMetadataXmlNamespace =
            @"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

        private const string _WordDcXmlNamespace =
            @"http://purl.org/dc/elements/1.1/";

        private const string _DocumentXmlPath =
            "/t:Types/t:Override[@ContentType=" +
            "application/vnd.openxmlformats-officedocument." +
            "wordprocessingml.document.main+xml\"]";

        private const string _DocumentBodyFile = "word/document.xml";
        private const string _DocumentBodyXPath = "/w:document/w:body";

        private const string _MetadataFile = "docProps/core.xml";
        private const string _MetadataXPath = "/cp:coreProperties";

        #endregion

        #region Constructors-and-Factories

        /// <summary>
        /// Instantiate.
        /// </summary>
        /// <param name="tempDirectory">Base temp directory.</param>
        /// <param name="filename">Filename.</param>
        public DocxParser(string tempDirectory, string filename)
        {
            if (String.IsNullOrEmpty(tempDirectory)) throw new ArgumentNullException(nameof(tempDirectory));
            if (String.IsNullOrEmpty(filename)) throw new ArgumentNullException(nameof(filename));

            tempDirectory = tempDirectory.Replace("\\", "/");
            if (!tempDirectory.EndsWith("/")) tempDirectory += "/";
            _TempDirectory = tempDirectory + _Guid.ToString() + "/";
            if (!Directory.Exists(_TempDirectory)) Directory.CreateDirectory(_TempDirectory);
            _DirInfo = new DirectoryInfo(_TempDirectory);

            _Filename = filename;

            using (ZipArchive archive = ZipFile.OpenRead(_Filename))
            {
                archive.ExtractToDirectory(_TempDirectory);
            }
        }

        #endregion

        #region Public-Methods

        /// <summary>
        /// Dispose of resources.
        /// </summary>
        public void Dispose()
        {
            RecursiveDelete(_DirInfo, true);
            Directory.Delete(_TempDirectory, true);
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
            xmlDoc.Load(_TempDirectory + _MetadataFile);

            XmlNamespaceManager nsmgr =
                new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", _WordXmlNamespace);
            nsmgr.AddNamespace("cp", _WordXmlNamespace);
            nsmgr.AddNamespace("dc", _WordDcXmlNamespace);

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

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.PreserveWhitespace = true;
            xmlDoc.Load(_TempDirectory + "word/document.xml");

            XmlNamespaceManager nsmgr =
                new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", _WordXmlNamespace);

            XmlNode node =
                xmlDoc.DocumentElement.SelectSingleNode(_DocumentBodyXPath, nsmgr);

            if (node == null)
                return string.Empty;

            StringBuilder sb = new StringBuilder();
            sb.Append(ReadNode(node));
            string ret = sb.ToString();

            while (ret.Contains("  ")) ret = ret.Replace("  ", " ");
            while (ret.Contains(Environment.NewLine + Environment.NewLine + Environment.NewLine)) ret = ret.Replace(Environment.NewLine + Environment.NewLine + Environment.NewLine, Environment.NewLine + Environment.NewLine);
            return ret;
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

                switch (child.LocalName)
                {
                    case "t":                           // Text
                        sb.Append(child.InnerText.TrimEnd());

                        string space =
                            ((XmlElement)child).GetAttribute("xml:space");
                        if (!string.IsNullOrEmpty(space) &&
                            space == "preserve")
                            sb.Append(' ');

                        break;

                    case "cr":                          // Carriage return
                    case "br":                          // Page break
                        sb.Append(Environment.NewLine);
                        break;

                    case "tab":                         // Tab
                        sb.Append("\t");
                        break;

                    case "p":                           // Paragraph
                        sb.Append(ReadNode(child));
                        sb.Append(Environment.NewLine);
                        sb.Append(Environment.NewLine);
                        break;

                    default:
                        sb.Append(ReadNode(child));
                        break;
                }
            }
            return sb.ToString();
        }

        private void RecursiveDelete(DirectoryInfo baseDir, bool isRootDir)
        {
            if (!baseDir.Exists) return;
            foreach (DirectoryInfo dir in baseDir.EnumerateDirectories()) RecursiveDelete(dir, false);
            foreach (FileInfo file in baseDir.GetFiles())
            {
                file.IsReadOnly = false;
                file.Delete();
            }
            if (!isRootDir)
            {
                baseDir.Delete();
            }
        }

        #endregion
    }
}

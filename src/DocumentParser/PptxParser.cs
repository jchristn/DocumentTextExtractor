﻿using System;
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
    /// Pptx parser.
    /// </summary>
    public class PptxParser : IDocumentParser, IDisposable
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
        public PptxParser(string tempDirectory, string filename)
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

            string root = _TempDirectory + _SlidesSubdirectory;
            StringBuilder sb = new StringBuilder();

            FileInfo[] files = new DirectoryInfo(_TempDirectory + _SlidesSubdirectory).GetFiles();
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
                    sb.Append(Environment.NewLine);
                }
            }

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
                if (String.IsNullOrEmpty(child.InnerText)) continue;

                // Console.WriteLine(child.NodeType.ToString() + " " + child.Name + " " + child.LocalName + " " + child.InnerText);

                switch (child.LocalName)
                {   
                    // case "p":   // Paragraph
                    // case "r":   // Run
                    case "t":   // Text
                        sb.Append(child.InnerText + " ");
                        sb.Append(ReadNode(child));
                        // sb.Append(Environment.NewLine);
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

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

namespace DocumentParser
{
    /// <summary>
    /// Pdf parser.
    /// </summary>
    public class PdfParser : IDocumentParser, IDisposable
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

        #endregion

        #region Constructors-and-Factories

        /// <summary>
        /// Instantiate.
        /// </summary>
        /// <param name="tempDirectory">Base temp directory.</param>
        /// <param name="filename">Filename.</param>
        public PdfParser(string tempDirectory, string filename)
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
            return ret;
        }

        /// <summary>
        /// Extract text from document.
        /// </summary>
        /// <returns>Text contents.</returns>
        public override string ExtractText()
        {
            return string.Empty;
        }

        #endregion

        #region Private-Methods

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

namespace DocumentParser
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Threading;

    /// <summary>
    /// Document text extractor abstract class.
    /// </summary>
    public abstract class IDocumentTextExtractor
    {
        #region Public-Members

        /// <summary>
        /// GUID.
        /// </summary>
        public Guid Guid
        {
            get
            {
                return _Guid;
            }
        }

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
            set
            {
                if (String.IsNullOrEmpty(value)) throw new ArgumentNullException(nameof(TempDirectory));

                value = value.Replace("\\", "/");
                if (!value.EndsWith("/")) value += "/";
                _TempDirectory = value + Guid.ToString() + "/";

                if (!Directory.Exists(TempDirectory)) Directory.CreateDirectory(TempDirectory);
                _DirInfo = new DirectoryInfo(TempDirectory);
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
            set
            {
                if (String.IsNullOrEmpty(value)) throw new ArgumentNullException(nameof(Filename));
                _Filename = value;
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

        #endregion

        #region Public-Methods

        /// <summary>
        /// Extract metadata from document.
        /// </summary>
        /// <returns>Dictionary containing metadata.</returns>
        public abstract Dictionary<string, string> ExtractMetadata();

        /// <summary>
        /// Extract text from document.
        /// </summary>
        /// <returns>Text contents.</returns>
        public abstract string ExtractText();

        /// <summary>
        /// Recursively delete a directory.
        /// </summary>
        /// <param name="baseDir">Base directory.</param>
        /// <param name="isRootDir">True to indicate the supplied directory is the root directory.</param>
        public void RecursiveDelete(DirectoryInfo baseDir, bool isRootDir)
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

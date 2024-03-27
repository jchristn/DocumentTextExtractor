using System;
using System.Collections.Generic;
using System.Text;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using HeyShelli;
using System.Threading.Tasks;
using DocumentParser;
using System.Reflection.PortableExecutable;

namespace DocumentTextExtractor
{
    /// <summary>
    /// PDF text extractor.
    /// </summary>
    public class PdfTextExtractor : IDocumentTextExtractor, IDisposable
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
                _Serializer = value ?? throw new ArgumentNullException(nameof(Serializer));
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

        /// <summary>
        /// Method to invoke to send log messages.
        /// </summary>
        public Action<string> Logger { get; set; } = null;

        #endregion

        #region Private-Members

        private string _Header = "[PdfParser] ";
        private SerializationHelper _Serializer = new();
        private readonly string _Filename = null;

        #endregion

        #region Constructors-and-Factories

        /// <summary>
        /// Instantiate.
        /// </summary>
        /// <param name="filename">Filename.</param>
        public PdfTextExtractor(string filename)
        {
            if (String.IsNullOrEmpty(filename)) throw new ArgumentNullException(nameof(filename));

            _Filename = filename;
        }

        #endregion

        #region Public-Methods

        /// <summary>
        /// Dispose of resources.
        /// </summary>
        public void Dispose()
        {
        }

        /// <summary>
        /// Extract metadata from document.
        /// </summary>
        /// <returns>Dictionary containing metadata.</returns>
        public override Dictionary<string, string> ExtractMetadata()
        {
            Dictionary<string, string> ret = new();

            PdfDocument doc = PdfReader.Open(_Filename);
            var metadata = doc.Info.Elements;
            foreach (var element in metadata)
            {
                ret.Add(element.Key, element.Value.ToString());
            }

            return ret;
        }

        /// <summary>
        /// Extract text from document.
        /// </summary>
        /// <returns>Text contents.</returns>
        public override string ExtractText()
        {
            StringBuilder dataSb = new StringBuilder();
            StringBuilder errorSb = new StringBuilder();

            DateTime lastDataReceived = DateTime.UtcNow;
            DateTime lastErrorReceived = DateTime.UtcNow;

            string command = "";

            if (OperatingSystem.IsWindows())
            {
                // https://stackoverflow.com/questions/14284269/why-doesnt-python-recognize-my-utf-8-encoded-source-file/14284404#14284404
                command += "chcp 65001 && SET PYTHONIOENCODING=utf-8 && ";
            }

            Shelli.OutputDataReceived = (s) =>
            {
                lastDataReceived = DateTime.UtcNow;
                dataSb.Append(s + Environment.NewLine);
            };

            Shelli.ErrorDataReceived = (s) =>
            {
                lastErrorReceived = DateTime.UtcNow;
                errorSb.Append(s + Environment.NewLine);
            };

            command += "py pdf.py " + _Filename;

            int returnCode = Shelli.Go(command);

            while
                (
                    DateTime.UtcNow < lastDataReceived.AddMilliseconds(250)
                    || DateTime.UtcNow < lastErrorReceived.AddMilliseconds(250)
                ) // may be more data
            {

            }

            if (returnCode == 0)
            {
                return dataSb.ToString();
            }
            else
            {
                Logger?.Invoke(_Header + "non-zero return code received from pdf.py: " + returnCode);
                return "Error: " + errorSb.ToString();
            }
        }

        #endregion

        #region Private-Methods

        #endregion
    }
}

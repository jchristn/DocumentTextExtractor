using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml;
using System.Xml.Serialization;

namespace DocumentParser
{
    /// <summary>
    /// Default serialization helper.
    /// </summary>
    public class SerializationHelper
    {
        #region Public-Members

        #endregion

        #region Private-Members

        private ExceptionConverter<Exception> _ExceptionConverter = new ExceptionConverter<Exception>();
        private NameValueCollectionConverter _NameValueCollectionConverter = new NameValueCollectionConverter();
        private JsonStringEnumConverter _StringEnumConverter = new JsonStringEnumConverter();

        #endregion

        #region Constructors-and-Factories

        /// <summary>
        /// Instantiate.
        /// </summary>
        public SerializationHelper()
        {
            InstantiateConverter();
        }

        #endregion

        #region Public-Methods

        /// <summary>
        /// Instantiation method to support fixups for various environments, e.g. Unity.
        /// </summary>
        public void InstantiateConverter()
        {
            try
            {
                Activator.CreateInstance<JsonStringEnumConverter>();
            }
            catch (Exception)
            {

            }
        }

        /// <summary>
        /// Deserialize JSON to an instance.
        /// </summary>
        /// <typeparam name="T">Type.</typeparam>
        /// <param name="json">JSON string.</param>
        /// <returns>Instance.</returns>
        public T DeserializeJson<T>(string json)
        {
            JsonSerializerOptions options = new JsonSerializerOptions();
            options.AllowTrailingCommas = true;
            options.ReadCommentHandling = JsonCommentHandling.Skip;

            options.Converters.Add(_ExceptionConverter);
            options.Converters.Add(_NameValueCollectionConverter);
            options.Converters.Add(_StringEnumConverter);
            return JsonSerializer.Deserialize<T>(json, options);
        }

        /// <summary>
        /// Serialize object to JSON.
        /// </summary>
        /// <param name="obj">Object.</param>
        /// <param name="pretty">Pretty print.</param>
        /// <returns>JSON.</returns>
        public string SerializeJson(object obj, bool pretty = true)
        {
            if (obj == null) return null;

            JsonSerializerOptions options = new JsonSerializerOptions();
            options.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull;

            // see https://github.com/dotnet/runtime/issues/43026
            options.Converters.Add(_ExceptionConverter);
            options.Converters.Add(_NameValueCollectionConverter);
            options.Converters.Add(_StringEnumConverter);

            if (!pretty)
            {
                options.WriteIndented = false;
                return JsonSerializer.Serialize(obj, options);
            }
            else
            {
                options.WriteIndented = true;
                return JsonSerializer.Serialize(obj, options);
            }
        }

        /// <summary>
        /// Deserialize XML.
        /// </summary>
        /// <typeparam name="T">Type.</typeparam>
        /// <param name="bytes">XML data.</param>
        /// <returns>Instance.</returns>
        public T DeserializeXml<T>(byte[] bytes) where T : class
        {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            return DeserializeXml<T>(Encoding.UTF8.GetString(bytes));
        }

        /// <summary>
        /// Deserialize XML.
        /// </summary>
        /// <typeparam name="T">Type.</typeparam>
        /// <param name="xml">XML string.</param>
        /// <returns>Instance.</returns>
        public T DeserializeXml<T>(string xml) where T : class
        {
            if (String.IsNullOrEmpty(xml)) throw new ArgumentNullException(nameof(xml));

            // remove preamble if exists
            string byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
            while (xml.StartsWith(byteOrderMarkUtf8, StringComparison.Ordinal))
            {
                xml = xml.Remove(0, byteOrderMarkUtf8.Length);
            }

            /*
             * 
             * This code respects the supplied namespace and validates it vs the model in code
             * 
             * 
            XmlSerializer xmls = new XmlSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                return (T)xmls.Deserialize(ms);
            }
            */

            // The code that follows ignores namespaces

            T obj = null;

            using (TextReader textReader = new StringReader(xml))
            {
                using (XmlTextReader reader = new XmlTextReader(textReader))
                {
                    reader.Namespaces = false;
                    XmlSerializer serializer = new XmlSerializer(typeof(T));
                    obj = (T)serializer.Deserialize(reader);
                }
            }

            return obj;
        }

        /// <summary>
        /// Serialize XML.
        /// </summary>
        /// <param name="obj">Object.</param>
        /// <param name="pretty">Pretty print.</param>
        /// <returns>XML string.</returns>
        public string SerializeXml(object obj, bool pretty = false)
        {
            if (obj == null) throw new ArgumentNullException(nameof(obj));

            XmlSerializer xml = new XmlSerializer(obj.GetType());

            using (MemoryStream stream = new MemoryStream())
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Encoding = Encoding.GetEncoding("ISO-8859-1");
                settings.NewLineChars = Environment.NewLine;
                settings.ConformanceLevel = ConformanceLevel.Document;
                if (pretty) settings.Indent = true;

                using (XmlWriter writer = XmlWriter.Create(stream, settings))
                {
                    xml.Serialize(new XmlWriterExtended(writer), obj);
                    byte[] bytes = stream.ToArray();
                    string ret = Encoding.UTF8.GetString(bytes, 0, bytes.Length);

                    // remove preamble if exists
                    string byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
                    while (ret.StartsWith(byteOrderMarkUtf8, StringComparison.Ordinal))
                    {
                        ret = ret.Remove(0, byteOrderMarkUtf8.Length);
                    }

                    return ret;
                }
            }
        }

        #endregion

        #region Private-Methods

        #endregion

        #region Private-Classes

        private class ExceptionConverter<TExceptionType> : JsonConverter<TExceptionType>
        {
            public override bool CanConvert(Type typeToConvert)
            {
                return typeof(Exception).IsAssignableFrom(typeToConvert);
            }

            public override TExceptionType Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            {
                throw new NotSupportedException("Deserializing exceptions is not allowed");
            }

            public override void Write(Utf8JsonWriter writer, TExceptionType value, JsonSerializerOptions options)
            {
                var serializableProperties = value.GetType()
                    .GetProperties()
                    .Select(uu => new { uu.Name, Value = uu.GetValue(value) })
                    .Where(uu => uu.Name != nameof(Exception.TargetSite));

                if (options.DefaultIgnoreCondition == JsonIgnoreCondition.WhenWritingNull)
                {
                    serializableProperties = serializableProperties.Where(uu => uu.Value != null);
                }

                var propList = serializableProperties.ToList();

                if (propList.Count == 0)
                {
                    // Nothing to write
                    return;
                }

                writer.WriteStartObject();

                foreach (var prop in propList)
                {
                    writer.WritePropertyName(prop.Name);
                    JsonSerializer.Serialize(writer, prop.Value, options);
                }

                writer.WriteEndObject();
            }
        }

        private class NameValueCollectionConverter : JsonConverter<NameValueCollection>
        {
            public override NameValueCollection Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) => throw new NotImplementedException();

            public override void Write(Utf8JsonWriter writer, NameValueCollection value, JsonSerializerOptions options)
            {
                var val = value.Keys.Cast<string>()
                    .ToDictionary(k => k, k => string.Join(", ", value.GetValues(k)));
                System.Text.Json.JsonSerializer.Serialize(writer, val);
            }
        }

        private class XmlWriterExtended : XmlWriter
        {
            private XmlWriter baseWriter;

            public XmlWriterExtended(XmlWriter w)
            {
                baseWriter = w;
            }

            // Force WriteEndElement to use WriteFullEndElement
            public override void WriteEndElement() { baseWriter.WriteFullEndElement(); }

            public override void WriteFullEndElement()
            {
                baseWriter.WriteFullEndElement();
            }

            public override void Close()
            {
                baseWriter.Close();
            }

            public override void Flush()
            {
                baseWriter.Flush();
            }

            public override string LookupPrefix(string ns)
            {
                return (baseWriter.LookupPrefix(ns));
            }

            public override void WriteBase64(byte[] buffer, int index, int count)
            {
                baseWriter.WriteBase64(buffer, index, count);
            }

            public override void WriteCData(string text)
            {
                baseWriter.WriteCData(text);
            }

            public override void WriteCharEntity(char ch)
            {
                baseWriter.WriteCharEntity(ch);
            }

            public override void WriteChars(char[] buffer, int index, int count)
            {
                baseWriter.WriteChars(buffer, index, count);
            }

            public override void WriteComment(string text)
            {
                baseWriter.WriteComment(text);
            }

            public override void WriteDocType(string name, string pubid, string sysid, string subset)
            {
                baseWriter.WriteDocType(name, pubid, sysid, subset);
            }

            public override void WriteEndAttribute()
            {
                baseWriter.WriteEndAttribute();
            }

            public override void WriteEndDocument()
            {
                baseWriter.WriteEndDocument();
            }

            public override void WriteEntityRef(string name)
            {
                baseWriter.WriteEntityRef(name);
            }

            public override void WriteProcessingInstruction(string name, string text)
            {
                baseWriter.WriteProcessingInstruction(name, text);
            }

            public override void WriteRaw(string data)
            {
                baseWriter.WriteRaw(data);
            }

            public override void WriteRaw(char[] buffer, int index, int count)
            {
                baseWriter.WriteRaw(buffer, index, count);
            }

            public override void WriteStartAttribute(string prefix, string localName, string ns)
            {
                baseWriter.WriteStartAttribute(prefix, localName, ns);
            }

            public override void WriteStartDocument(bool standalone)
            {
                baseWriter.WriteStartDocument(standalone);
            }

            public override void WriteStartDocument()
            {
                baseWriter.WriteStartDocument();
            }

            public override void WriteStartElement(string prefix, string localName, string ns)
            {
                baseWriter.WriteStartElement(prefix, localName, ns);
            }

            public override WriteState WriteState
            {
                get { return baseWriter.WriteState; }
            }

            public override void WriteString(string text)
            {
                baseWriter.WriteString(text);
            }

            public override void WriteSurrogateCharEntity(char lowChar, char highChar)
            {
                baseWriter.WriteSurrogateCharEntity(lowChar, highChar);
            }

            public override void WriteWhitespace(string ws)
            {
                baseWriter.WriteWhitespace(ws);
            }
        }
    }

    #endregion
}
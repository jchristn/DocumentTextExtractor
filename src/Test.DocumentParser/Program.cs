using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using DocumentParser;
using GetSomeInput;
using static System.Net.Mime.MediaTypeNames;

namespace Test.DocumentParser
{
    public static class Program
    {
        private static bool _RunForever = true;
        private static string _TempDirectory = "./temp/";
        private static DocxParser _Word = null;
        private static PptxParser _PowerPoint = null;

        public static void Main(string[] args)
        {
            while (_RunForever)
            {
                string userInput = Inputty.GetString("Command [?/help]:", null, false);

                string[] parts = null;

                if (userInput.StartsWith("tokenize ") && userInput.Length > "tokenize ".Length)
                {
                    parts = userInput.Split(" ", 2);
                    if (parts.Length > 1) ExtractTokensFromFile(parts[1]);
                }
                else if (userInput.StartsWith("metadata ") && userInput.Length > "metadata ".Length)
                {
                    parts = userInput.Split(" ", 2);
                    if (parts.Length > 1) ExtractMetadataFromFile(parts[1]);
                }
                else if (userInput.StartsWith("both ") && userInput.Length > "both ".Length)
                {
                    parts = userInput.Split(" ", 2);
                    ExtractBoth(parts[1]);
                }
                else
                {
                    switch (userInput)
                    {
                        case "?":
                            Menu();
                            break;
                        case "c":
                        case "cls":
                            Console.Clear();
                            break;
                        case "q":
                            _RunForever = false;
                            break;
                    }
                }
            }
        }

        public static void Menu()
        {
            Console.WriteLine("");
            Console.WriteLine("Available commands:");
            Console.WriteLine("  ?                 help, this menu");
            Console.WriteLine("  q                 quit this program");
            Console.WriteLine("  cls               clear the screen");
            Console.WriteLine("  tokenize [file]   tokenize file [filename]");
            Console.WriteLine("  metadata [file]   extract metadata from file [filename]");
            Console.WriteLine("  both [file]       tokenize and extract metadata from file [filename]");
            Console.WriteLine("");
        }

        private static void ExtractTokensFromFile(string filename)
        {
            string text = null;

            if (File.Exists(filename))
            {
                if (filename.ToLower().EndsWith(".docx"))
                {
                    using (_Word = new DocxParser(_TempDirectory, filename))
                    {
                        text = _Word.ExtractText();
                    }
                }
                else if (filename.ToLower().EndsWith(".pptx"))
                {
                    using (_PowerPoint = new PptxParser(_TempDirectory, filename))
                    {
                        text = _PowerPoint.ExtractText();
                    }
                }

                if (!String.IsNullOrEmpty(text))
                {
                    Console.WriteLine("");
                    Console.WriteLine(text);
                    Console.WriteLine("");
                }
                else
                {
                    Console.WriteLine("");
                    Console.WriteLine("No text.");
                    Console.WriteLine("");
                }
            }
            else
            {
                Console.WriteLine("");
                Console.WriteLine("*** File '" + filename + "' does not exist.");
                Console.WriteLine("");
            }
        }

        private static void ExtractMetadataFromFile(string filename)
        {
            Dictionary<string, string> metadata = null;

            if (File.Exists(filename))
            {
                if (filename.ToLower().EndsWith(".docx"))
                {
                    using (_Word = new DocxParser(_TempDirectory, filename))
                    {
                        metadata = _Word.ExtractMetadata();
                    }
                }
                else if (filename.ToLower().EndsWith(".pptx"))
                {
                    using (_PowerPoint = new PptxParser(_TempDirectory, filename))
                    {
                        metadata = _PowerPoint.ExtractMetadata();
                    }
                }

                if (metadata != null)
                {
                    Console.WriteLine("");
                    Console.WriteLine("Metadata:");
                    Console.WriteLine("");

                    foreach (KeyValuePair<string, string> kvp in metadata)
                    {
                        Console.WriteLine(kvp.Key + ": " + kvp.Value);
                    }

                    Console.WriteLine("");
                }
                else
                {
                    Console.WriteLine("");
                    Console.WriteLine("No metadata.");
                    Console.WriteLine("");
                }
            }
            else
            {
                Console.WriteLine("");
                Console.WriteLine("*** File '" + filename + "' does not exist.");
                Console.WriteLine("");
            }
        }

        private static void ExtractBoth(string filename)
        {
            string text = null;
            Dictionary<string, string> metadata = null;

            if (File.Exists(filename))
            {
                if (filename.ToLower().EndsWith(".docx"))
                {
                    using (_Word = new DocxParser(_TempDirectory, filename))
                    {
                        text = _Word.ExtractText();
                        metadata = _Word.ExtractMetadata();
                    }
                }
                else if (filename.ToLower().EndsWith(".pptx"))
                {
                    using (_PowerPoint = new PptxParser(_TempDirectory, filename))
                    {
                        text = _PowerPoint.ExtractText();
                        metadata = _PowerPoint.ExtractMetadata();
                    }
                }
            }

            if (!String.IsNullOrEmpty(text))
            {
                Console.WriteLine("");
                Console.WriteLine("Contents:");
                Console.WriteLine("");
                Console.WriteLine(text);
                Console.WriteLine("");
            }
            else
            {
                Console.WriteLine("");
                Console.WriteLine("No text.");
                Console.WriteLine("");
            }

            if (metadata != null)
            {
                Console.WriteLine("");
                Console.WriteLine("Metadata:");
                Console.WriteLine("");

                foreach (KeyValuePair<string, string> kvp in metadata)
                {
                    Console.WriteLine(kvp.Key + ": " + kvp.Value);
                }

                Console.WriteLine("");
            }
            else
            {
                Console.WriteLine("");
                Console.WriteLine("No metadata.");
                Console.WriteLine("");
            }
        }
    }
}
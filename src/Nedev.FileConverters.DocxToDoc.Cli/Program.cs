using System;
using System.IO;
using Nedev.FileConverters.DocxToDoc;

namespace Nedev.FileConverters.DocxToDoc.Cli
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length < 2 || args[0] == "-h" || args[0] == "--help")
            {
                ShowHelp();
                return 1;
            }

            string input = args[0];
            string output = args[1];

            try
            {
                if (!File.Exists(input))
                {
                    Console.Error.WriteLine($"Error: input file '{input}' does not exist.");
                    return 2;
                }

                var converter = new DocxToDocConverter();
                converter.Convert(input, output);
                Console.WriteLine($"Converted '{input}' -> '{output}'");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Conversion failed: " + ex.Message);
                return 3;
            }
        }

        private static void ShowHelp()
        {
            Console.WriteLine("Usage: dotnet Nedev.FileConverters.DocxToDoc.Cli.dll <input.docx> <output.doc>");
            Console.WriteLine();
            Console.WriteLine("Simple command‑line front‑end for DocxToDocConverter.");
            Console.WriteLine("Provide a path to a .docx file followed by the destination .doc file.");
        }
    }
}
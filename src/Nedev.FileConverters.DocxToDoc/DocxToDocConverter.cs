using System;
using System.IO;
using System.Text;

namespace Nedev.FileConverters.DocxToDoc
{
    /// <summary>
    /// Provides high-performance conversion from OpenXML (.docx) into MS-DOC legacy binary (.doc) format.
    /// Does not rely on any third-party libraries.
    /// </summary>
    public class DocxToDocConverter
    {
        static DocxToDocConverter()
        {
            // attempt to register code pages encoding provider if available (some targets may not include the
            // type by default, e.g. netstandard2.1). Use reflection so compilation succeeds across all TFM.
            var providerType = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages");
            if (providerType != null)
            {
                var instanceProp = providerType.GetProperty("Instance", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var instance = instanceProp?.GetValue(null);
                if (instance is EncodingProvider ep)
                {
                    Encoding.RegisterProvider(ep);
                }
            }
        }

        /// <summary>
        /// Converts a .docx file to a .doc file.
        /// </summary>
        /// <param name="docxPath">The path to the source .docx file.</param>
        /// <param name="docPath">The path to the destination .doc file.</param>
        public void Convert(string docxPath, string docPath)
        {
            if (string.IsNullOrWhiteSpace(docxPath))
                throw new ArgumentNullException(nameof(docxPath));
            if (string.IsNullOrWhiteSpace(docPath))
                throw new ArgumentNullException(nameof(docPath));

            using var inputStream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var outputStream = new FileStream(docPath, FileMode.Create, FileAccess.Write, FileShare.None);
            
            Convert(inputStream, outputStream);
        }

        /// <summary>
        /// Converts a .docx stream to a .doc stream.
        /// </summary>
        /// <param name="docxStream">A stream containing the OpenXML document. Must support Read.</param>
        /// <param name="docStream">A stream where the .doc binary will be written. Must support Write.</param>
        public void Convert(Stream docxStream, Stream docStream)
        {
            if (docxStream == null) throw new ArgumentNullException(nameof(docxStream));
            if (docStream == null) throw new ArgumentNullException(nameof(docStream));

            // Create a DocxReader to parse the document contents
            using var reader = new Format.DocxReader(docxStream);
            
            // Extract necessary layout/styles/content out of OpenXML and map to MS-DOC
            var documentModel = reader.ReadDocument();
            
            // Provide data blocks for the MS-DOC writer
            var writer = new Format.DocWriter();
            writer.WriteDocBlocks(documentModel, docStream);
        }
    }
}

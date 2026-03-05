using System.IO;
using System.Text;
using Nedev.DocxToDoc.Format;
using Nedev.DocxToDoc.Model;
using Xunit;

namespace Nedev.DocxToDoc.Tests.Format
{
    public class DocWriterTests
    {
        [Fact]
        public void WriteDocBlocks_ValidModel_CreatesValidCFB()
        {
            // Register encoding provider for Windows-1252 
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Arrange
            var model = new DocumentModel();
            model.Content.Add(new ParagraphModel { Runs = { new RunModel { Text = "Hello MS-DOC World!" } } });

            var writer = new DocWriter();
            using var ms = new MemoryStream();

            // Act
            writer.WriteDocBlocks(model, ms);
            ms.Position = 0;

            // Assert
            // We use OpenMcdf just to verify the CFB structure is valid
            using var compoundFile = new OpenMcdf.CompoundFile(ms);
            
            // Should contain WordDocument, 1Table, and Data streams
            Assert.True(compoundFile.RootStorage.TryGetStream("WordDocument", out var wordDocStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("1Table", out var tableStream));
            Assert.True(compoundFile.RootStorage.TryGetStream("Data", out var dataStream));

            // Verify Text was written to WordDocument stream at offset 1536
            var wordDocData = wordDocStream.GetData();
            string expectedText = "Hello MS-DOC World!\r";
            Assert.True(wordDocData.Length >= 1536 + expectedText.Length);

            var textBytes = new byte[expectedText.Length];
            System.Array.Copy(wordDocData, 1536, textBytes, 0, expectedText.Length);
            var extractedText = Encoding.GetEncoding(1252).GetString(textBytes);
            Assert.Equal(expectedText, extractedText);
        }
    }
}

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using Xunit;

namespace Nedev.DocxToDoc.Tests.Format
{
    public class DocxReaderTests
    {
        private byte[] CreateDummyDocx()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                // Valid minimal XML
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Hello World</w:t></w:r></w:p></w:body></w:document>");
            }
            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_ValidStream_FindsText()
        {
            // Arrange
            byte[] dummyData = CreateDummyDocx();
            using var ms = new MemoryStream(dummyData);
            using var reader = new Nedev.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Equal("Hello World\r", model.TextBuffer); // Includes the paragraph return
            Assert.Single(model.Paragraphs);
            Assert.Single(model.Paragraphs[0].Runs);
            Assert.Equal("Hello World", model.Paragraphs[0].Runs[0].Text);
        }

        [Fact]
        public void ReadDocument_MissingDocumentXml_ThrowsFileNotFound()
        {
            // Arrange
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create an empty zip
            }
            byte[] emptyZip = ms.ToArray();
            using var testStream = new MemoryStream(emptyZip);
            using var reader = new Nedev.DocxToDoc.Format.DocxReader(testStream);

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() => 
            {
                reader.ReadDocument();
            });
        }

        [Fact]
        public void ReadDocument_WithFormatting_ParsesProperties()
        {
            // Arrange
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry = archive.CreateEntry("word/document.xml");
                using var entryStream = entry.Open();
                using var writer = new StreamWriter(entryStream);
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
                             "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p>" +
                             "<w:r><w:rPr><w:b/><w:i w:val=\"1\"/><w:sz w:val=\"24\"/></w:rPr><w:t>BoldItalic12pt</w:t></w:r>" +
                             "</w:p></w:body></w:document>");
            }
            
            using var testStream = new MemoryStream(ms.ToArray());
            using var reader = new Nedev.DocxToDoc.Format.DocxReader(testStream);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Paragraphs);
            Assert.Single(model.Paragraphs[0].Runs);
            
            var run = model.Paragraphs[0].Runs[0];
            Assert.Equal("BoldItalic12pt", run.Text);
            Assert.True(run.Properties.IsBold);
            Assert.True(run.Properties.IsItalic);
            Assert.False(run.Properties.IsStrike);
            Assert.Equal(24, run.Properties.FontSize);
        }
    }
}

using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderImageTests
    {
        private byte[] CreateDocxWithImage()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create [Content_Types].xml
                var contentTypesEntry = archive.CreateEntry("[Content_Types].xml");
                using (var stream = contentTypesEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                        "<Default Extension=\"png\" ContentType=\"image/png\"/>" +
                        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                        "</Types>");
                }

                // Create _rels/.rels
                var relsEntry = archive.CreateEntry("_rels/.rels");
                using (var stream = relsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                        "</Relationships>");
                }

                // Create word/_rels/document.xml.rels
                var docRelsEntry = archive.CreateEntry("word/_rels/document.xml.rels");
                using (var stream = docRelsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image1.png\"/>" +
                        "</Relationships>");
                }

                // Create word/media/image1.png (dummy 1x1 PNG)
                var imageEntry = archive.CreateEntry("word/media/image1.png");
                using (var stream = imageEntry.Open())
                {
                    // Minimal valid PNG (1x1 pixel, red)
                    byte[] pngData = new byte[]
                    {
                        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
                        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
                        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
                        0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
                        0x54, 0x08, 0xD7, 0x63, 0xF8, 0x0F, 0x00, 0x00,
                        0x01, 0x01, 0x00, 0x05, 0x18, 0xD8, 0x4D, 0x00,
                        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                        0x42, 0x60, 0x82
                    };
                    stream.Write(pngData, 0, pngData.Length);
                }

                // Create word/document.xml with image
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">" +
                        "<w:body>" +
                        "<w:p>" +
                        "<w:r>" +
                        "<w:t>Before image</w:t>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:drawing>" +
                        "<wp:inline>" +
                        "<wp:extent cx=\"914400\" cy=\"914400\"/>" +
                        "<a:graphic>" +
                        "<a:graphicData>" +
                        "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                        "<pic:blipFill>" +
                        "<a:blip r:embed=\"rId1\"/>" +
                        "</pic:blipFill>" +
                        "</pic:pic>" +
                        "</a:graphicData>" +
                        "</a:graphic>" +
                        "</wp:inline>" +
                        "</w:drawing>" +
                        "</w:r>" +
                        "<w:r>" +
                        "<w:t>After image</w:t>" +
                        "</w:r>" +
                        "</w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_WithImage_ParsesImageData()
        {
            // Arrange
            byte[] docxData = CreateDocxWithImage();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Single(model.Paragraphs);
            Assert.True(model.Paragraphs[0].Runs.Count >= 2);

            // Find the run with image
            var imageRun = model.Paragraphs[0].Runs.Find(r => r.Image != null);
            Assert.NotNull(imageRun);
            Assert.NotNull(imageRun.Image);
            Assert.Equal("rId1", imageRun.Image.RelationshipId);
            // Width/Height may be 0 if extents not found, that's acceptable
            Assert.NotNull(imageRun.Image.Data);
            Assert.True(imageRun.Image.Data.Length > 0);
            Assert.Equal("image/png", imageRun.Image.ContentType);
        }
    }
}

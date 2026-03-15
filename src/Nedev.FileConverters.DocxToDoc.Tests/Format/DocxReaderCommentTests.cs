using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace Nedev.FileConverters.DocxToDoc.Tests.Format
{
    public class DocxReaderCommentTests
    {
        private byte[] CreateDocxWithComments()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/comments.xml
                var commentsEntry = archive.CreateEntry("word/comments.xml");
                using (var stream = commentsEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                        "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">" +
                        "<w:comment w:id=\"0\" w:author=\"John Doe\" w:initials=\"JD\" w:date=\"2024-01-15T10:30:00Z\">" +
                        "<w:p><w:r><w:t>This is a test comment</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "<w:comment w:id=\"1\" w:author=\"Jane Smith\" w:initials=\"JS\" w:date=\"2024-01-16T14:20:00Z\" w:done=\"1\">" +
                        "<w:p><w:r><w:t>Another comment that is resolved</w:t></w:r></w:p>" +
                        "</w:comment>" +
                        "</w:comments>");
                }

                // Create word/document.xml
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Document with comments</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        private byte[] CreateDocxWithoutComments()
        {
            using var ms = new MemoryStream();
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                // Create word/document.xml without comments
                var docEntry = archive.CreateEntry("word/document.xml");
                using (var stream = docEntry.Open())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                        "<w:body>" +
                        "<w:p><w:r><w:t>Document without comments</w:t></w:r></w:p>" +
                        "</w:body>" +
                        "</w:document>");
                }
            }
            return ms.ToArray();
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesAllComments()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Equal(2, model.Comments.Count);
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentId()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.Id == "0");
            Assert.Contains(model.Comments, c => c.Id == "1");
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentAuthor()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.Author == "John Doe");
            Assert.Contains(model.Comments, c => c.Author == "Jane Smith");
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentInitials()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.Initials == "JD");
            Assert.Contains(model.Comments, c => c.Initials == "JS");
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentText()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.Text == "This is a test comment");
            Assert.Contains(model.Comments, c => c.Text == "Another comment that is resolved");
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesCommentDate()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.All(model.Comments, c => Assert.NotNull(c.Date));
        }

        [Fact]
        public void ReadDocument_WithComments_ParsesDoneStatus()
        {
            // Arrange
            byte[] docxData = CreateDocxWithComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Contains(model.Comments, c => c.IsDone == true);
            Assert.Contains(model.Comments, c => c.IsDone == false);
        }

        [Fact]
        public void ReadDocument_WithoutComments_HasEmptyCommentsList()
        {
            // Arrange
            byte[] docxData = CreateDocxWithoutComments();
            using var ms = new MemoryStream(docxData);
            using var reader = new Nedev.FileConverters.DocxToDoc.Format.DocxReader(ms);

            // Act
            var model = reader.ReadDocument();

            // Assert
            Assert.Empty(model.Comments);
        }
    }
}
